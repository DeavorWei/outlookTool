package server

import (
	"context"
	"embed"
	"encoding/json"
	"fmt"
	"io/fs"
	"net"
	"net/http"
	"os"
	"os/exec"
	"sync"
	"time"

	"go.uber.org/zap"
	"outlook-archiver/internal/config"
	"outlook-archiver/internal/logger"
	"outlook-archiver/internal/registry"
	"outlook-archiver/internal/scheduler"
)

//go:embed web/*
var webFS embed.FS

type WebServer struct {
	logger     *zap.Logger
	sched      *scheduler.Scheduler
	cfgPath    string
	onShutdown func() // Web 服务停止时回调（仅取消菜单勾选）
	onAppQuit  func() // 用户请求退出整个程序时回调（触发托盘退出）
	ctx        context.Context

	mu            sync.Mutex
	srv           *http.Server
	isRunning     bool
	port          int
	lastHeartbeat time.Time
	shutdownCh    chan struct{}
}

func NewWebServer(ctx context.Context, logger *zap.Logger, sched *scheduler.Scheduler, cfgPath string, onShutdown func(), onAppQuit func()) *WebServer {
	return &WebServer{
		logger:     logger,
		sched:      sched,
		cfgPath:    cfgPath,
		onShutdown: onShutdown,
		onAppQuit:  onAppQuit,
		ctx:        ctx,
	}
}

func (ws *WebServer) Start() (int, error) {
	ws.mu.Lock()
	defer ws.mu.Unlock()

	if ws.isRunning {
		return ws.port, nil
	}

	ws.lastHeartbeat = time.Now()
	ws.shutdownCh = make(chan struct{})

	mux := http.NewServeMux()

	// API routes
	mux.HandleFunc("/api/config", ws.handleConfig)
	mux.HandleFunc("/api/heartbeat", ws.handleHeartbeat)
	mux.HandleFunc("/api/mounted-psts", ws.handleMountedPSTs)
	mux.HandleFunc("/api/actions/restore", ws.handleRestore)

	mux.HandleFunc("/api/logs/stream", ws.handleLogStream)
	mux.HandleFunc("/api/actions/trigger", ws.handleActionTrigger)
	mux.HandleFunc("/api/actions/reorganize", ws.handleActionReorganize)
	mux.HandleFunc("/api/actions/reload", ws.handleActionReload)
	mux.HandleFunc("/api/actions/open-dir", ws.handleActionOpenDir)
	mux.HandleFunc("/api/actions/autostart", ws.handleActionAutostart)
	mux.HandleFunc("/api/actions/quit", ws.handleActionQuit)

	// Static files
	webSubFS, err := fs.Sub(webFS, "web")
	if err != nil {
		return 0, err
	}
	mux.Handle("/", http.FileServer(http.FS(webSubFS)))

	// Find available port starting from 8443
	port := 8443
	var listener net.Listener
	for {
		addr := fmt.Sprintf("127.0.0.1:%d", port)
		l, err := net.Listen("tcp", addr)
		if err == nil {
			listener = l
			break
		}
		port++
		if port > 8500 {
			return 0, fmt.Errorf("no available ports found between 8443 and 8500")
		}
	}

	ws.port = port
	ws.srv = &http.Server{
		Handler:     mux,
		BaseContext: func(_ net.Listener) context.Context { return ws.ctx },
	}
	ws.isRunning = true

	go func() {
		ws.logger.Info("Web server started", zap.Int("port", port))
		if err := ws.srv.Serve(listener); err != nil && err != http.ErrServerClosed {
			ws.logger.Error("Web server error", zap.Error(err))
		}
	}()

	go ws.monitorHeartbeat()

	return port, nil
}

func (ws *WebServer) Stop() {
	ws.mu.Lock()
	if !ws.isRunning {
		ws.mu.Unlock()
		return
	}
	ws.isRunning = false
	close(ws.shutdownCh) // stops monitor
	srv := ws.srv
	ws.mu.Unlock()

	ctx, cancel := context.WithTimeout(context.Background(), 10*time.Second)
	defer cancel()
	if srv != nil {
		srv.Shutdown(ctx)
	}

	if ws.onShutdown != nil {
		ws.onShutdown()
	}
	ws.logger.Info("Web server stopped due to manual stop or heartbeat timeout")
}

func (ws *WebServer) monitorHeartbeat() {
	ticker := time.NewTicker(30 * time.Second)
	defer ticker.Stop()

	for {
		select {
		case <-ws.shutdownCh:
			return
		case <-ticker.C:
			ws.mu.Lock()
			elapsed := time.Since(ws.lastHeartbeat)
			ws.mu.Unlock()

			if elapsed > 5*time.Minute {
				ws.logger.Info("Web server heartbeat timeout, shutting down")
				ws.Stop()
				return
			}
		}
	}
}

func (ws *WebServer) handleHeartbeat(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
		return
	}
	ws.mu.Lock()
	ws.lastHeartbeat = time.Now()
	ws.mu.Unlock()
	w.WriteHeader(http.StatusOK)
}

func (ws *WebServer) handleConfig(w http.ResponseWriter, r *http.Request) {
	w.Header().Set("Access-Control-Allow-Origin", "*")
	w.Header().Set("Content-Type", "application/json")

	if r.Method == http.MethodGet {
		cfgCopy := ws.sched.GetConfigCopy()

		// Read registry state for UI
		cfgCopy.AutoStart = registry.IsAutoStartEnabled()

		if err := json.NewEncoder(w).Encode(cfgCopy); err != nil {
			http.Error(w, err.Error(), http.StatusInternalServerError)
		}
		return
	}

	if r.Method == http.MethodPost {
		var newCfg config.Config
		if err := json.NewDecoder(r.Body).Decode(&newCfg); err != nil {
			http.Error(w, err.Error(), http.StatusBadRequest)
			return
		}

		// Validate in memory first
		if err := config.ValidateConfig(&newCfg); err != nil {
			http.Error(w, err.Error(), http.StatusBadRequest)
			return
		}

		// Update registry for autostart
		if err := registry.SetAutoStart(newCfg.AutoStart); err != nil {
			ws.logger.Error("Failed to update autostart setting", zap.Error(err))
		}

		// Use atomic write to file (simple approach: write temp, then rename)
		tmpPath := ws.cfgPath + ".tmp"
		if err := config.SaveConfig(tmpPath, &newCfg); err != nil {
			http.Error(w, err.Error(), http.StatusInternalServerError)
			return
		}
		if err := os.Rename(tmpPath, ws.cfgPath); err != nil {
			os.Remove(tmpPath)
			http.Error(w, "Failed to update config file atomically: "+err.Error(), http.StatusInternalServerError)
			return
		}

		// Reload configuration into scheduler
		if err := ws.sched.ReloadConfig(ws.cfgPath); err != nil {
			http.Error(w, "已保存到文件，但热加载失败 (程序可能正在归档中，请稍后手动重载或重启): "+err.Error(), http.StatusInternalServerError)
			return
		}

		w.WriteHeader(http.StatusOK)
		w.Write([]byte(`{"status":"success"}`))
		return
	}

	http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
}

func (ws *WebServer) handleMountedPSTs(w http.ResponseWriter, r *http.Request) {
	w.Header().Set("Access-Control-Allow-Origin", "*")
	w.Header().Set("Content-Type", "application/json")

	if r.Method != http.MethodGet {
		http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
		return
	}

	psts, err := ws.sched.GetMountedPSTs()
	if err != nil {
		ws.logger.Error("获取挂载的PST失败", zap.Error(err))
		http.Error(w, err.Error(), http.StatusInternalServerError)
		return
	}

	if err := json.NewEncoder(w).Encode(psts); err != nil {
		http.Error(w, err.Error(), http.StatusInternalServerError)
	}
}

func (ws *WebServer) handleRestore(w http.ResponseWriter, r *http.Request) {
	w.Header().Set("Access-Control-Allow-Origin", "*")
	w.Header().Set("Content-Type", "application/json")

	if r.Method != http.MethodPost {
		http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
		return
	}

	var req scheduler.RestoreRequest
	if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
		http.Error(w, "Invalid request body", http.StatusBadRequest)
		return
	}

	// Trigger restore asynchronously to avoid blocking the HTTP response
	go func() {
		_ = ws.sched.TriggerRestore(ws.ctx, req)
	}()

	w.WriteHeader(http.StatusOK)
	w.Write([]byte(`{"status":"success"}`))
}

func (ws *WebServer) IsRunning() bool {
	ws.mu.Lock()
	defer ws.mu.Unlock()
	return ws.isRunning
}

// Actions

func (ws *WebServer) handleActionTrigger(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
		return
	}
	go ws.sched.TriggerOnce(ws.ctx)
	w.WriteHeader(http.StatusOK)
	w.Write([]byte(`{"status":"success"}`))
}

func (ws *WebServer) handleActionReorganize(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
		return
	}
	go ws.sched.TriggerReorganize(ws.ctx, func(info scheduler.ProgressInfo) {
		ws.logger.Info("全量整理进度", zap.Int("phase", info.Phase), zap.Int("processed", info.Processed))
	})
	w.WriteHeader(http.StatusOK)
	w.Write([]byte(`{"status":"success"}`))
}

func (ws *WebServer) handleActionReload(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
		return
	}
	err := ws.sched.ReloadConfig(ws.cfgPath)
	if err != nil {
		ws.logger.Error("重新加载配置失败", zap.Error(err))
		http.Error(w, err.Error(), http.StatusInternalServerError)
		return
	}
	ws.logger.Info("重新加载配置成功")
	w.WriteHeader(http.StatusOK)
	w.Write([]byte(`{"status":"success"}`))
}

func (ws *WebServer) handleActionOpenDir(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
		return
	}
	target := r.URL.Query().Get("target")
	if target == "log" {
		if logger.CurrentLogDir != "" {
			exec.Command("explorer", logger.CurrentLogDir).Start()
		}
	} else if target == "config" {
		exec.Command("cmd", "/c", "start", "", ws.cfgPath).Start()
	} else {
		http.Error(w, "Invalid target", http.StatusBadRequest)
		return
	}
	w.WriteHeader(http.StatusOK)
	w.Write([]byte(`{"status":"success"}`))
}

func (ws *WebServer) handleActionAutostart(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
		return
	}
	var req struct {
		Enabled bool `json:"enabled"`
	}
	if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
		http.Error(w, err.Error(), http.StatusBadRequest)
		return
	}
	if err := registry.SetAutoStart(req.Enabled); err != nil {
		ws.logger.Error("设置开机自启失败", zap.Error(err))
		http.Error(w, err.Error(), http.StatusInternalServerError)
		return
	}
	if req.Enabled {
		ws.logger.Info("已开启开机自启")
	} else {
		ws.logger.Info("已关闭开机自启")
	}
	w.WriteHeader(http.StatusOK)
	w.Write([]byte(`{"status":"success"}`))
}

func (ws *WebServer) handleActionQuit(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
		return
	}
	if ws.sched.GetState() == scheduler.StateReorganizing {
		http.Error(w, "全量整理中，禁止退出", http.StatusConflict)
		return
	}
	ws.logger.Info("用户通过控制台触发退出程序")
	w.WriteHeader(http.StatusOK)
	w.Write([]byte(`{"status":"success"}`))

	go func() {
		if ws.onAppQuit != nil {
			ws.onAppQuit()
		}
	}()
}
