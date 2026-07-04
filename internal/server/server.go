package server

import (
	"context"
	"embed"
	"encoding/json"
	"fmt"
	"io/fs"
	"net"
	"net/http"
	"sync"
	"time"

	"go.uber.org/zap"
	"outlook-archiver/internal/config"
	"outlook-archiver/internal/registry"
	"outlook-archiver/internal/scheduler"
)

//go:embed web/*
var webFS embed.FS

type WebServer struct {
	logger        *zap.Logger
	sched         *scheduler.Scheduler
	cfgPath       string
	onShutdown    func()

	mu            sync.Mutex
	srv           *http.Server
	isRunning     bool
	port          int
	lastHeartbeat time.Time
	shutdownCh    chan struct{}
}

func NewWebServer(logger *zap.Logger, sched *scheduler.Scheduler, cfgPath string, onShutdown func()) *WebServer {
	return &WebServer{
		logger:     logger,
		sched:      sched,
		cfgPath:    cfgPath,
		onShutdown: onShutdown,
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
	mux.HandleFunc("/api/restore", ws.handleRestore)
	
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
	ws.srv = &http.Server{Handler: mux}
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

	ctx, cancel := context.WithTimeout(context.Background(), 2*time.Second)
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

		if err := config.ValidateConfig(&newCfg); err != nil {
			http.Error(w, err.Error(), http.StatusBadRequest)
			return
		}

		// Update registry for autostart
		if err := registry.SetAutoStart(newCfg.AutoStart); err != nil {
			ws.logger.Error("Failed to update autostart setting", zap.Error(err))
		}

		// Save to disk
		if err := config.SaveConfig(ws.cfgPath, &newCfg); err != nil {
			http.Error(w, err.Error(), http.StatusInternalServerError)
			return
		}

		// Reload configuration into scheduler
		if err := ws.sched.ReloadConfig(ws.cfgPath); err != nil {
			http.Error(w, "已保存到文件，但热加载失败 (程序正在归档中，请稍后手动重载或重启): "+err.Error(), http.StatusInternalServerError)
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
		_ = ws.sched.TriggerRestore(context.Background(), req)
	}()

	w.WriteHeader(http.StatusOK)
	w.Write([]byte(`{"status":"success"}`))
}

func (ws *WebServer) IsRunning() bool {
	ws.mu.Lock()
	defer ws.mu.Unlock()
	return ws.isRunning
}
