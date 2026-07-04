package server

import (
	"fmt"
	"net/http"
	"sync"

	"outlook-archiver/internal/logger"
)

type LogHub struct {
	clients map[chan string]bool
	mu      sync.Mutex
}

var globalLogHub *LogHub
var hubOnce sync.Once

func GetLogHub() *LogHub {
	hubOnce.Do(func() {
		globalLogHub = &LogHub{
			clients: make(map[chan string]bool),
		}
		go globalLogHub.run()
	})
	return globalLogHub
}

func (h *LogHub) run() {
	for logMsg := range logger.LogBroadcast {
		h.mu.Lock()
		for client := range h.clients {
			select {
			case client <- logMsg:
			default:
				// If client channel is full, skip to avoid blocking others
			}
		}
		h.mu.Unlock()
	}
}

func (ws *WebServer) handleLogStream(w http.ResponseWriter, r *http.Request) {
	w.Header().Set("Content-Type", "text/event-stream")
	w.Header().Set("Cache-Control", "no-cache")
	w.Header().Set("Connection", "keep-alive")
	w.Header().Set("Access-Control-Allow-Origin", "*")

	hub := GetLogHub()
	
	clientChan := make(chan string, 100)
	hub.mu.Lock()
	hub.clients[clientChan] = true
	hub.mu.Unlock()

	defer func() {
		hub.mu.Lock()
		delete(hub.clients, clientChan)
		hub.mu.Unlock()
		close(clientChan)
	}()

	flusher, ok := w.(http.Flusher)
	if !ok {
		http.Error(w, "Streaming unsupported!", http.StatusInternalServerError)
		return
	}

	// Send an initial connected message (JSON formatted to match Zap output)
	fmt.Fprintf(w, "data: {\"level\":\"info\",\"msg\":\"Log stream connected...\"}\n\n")
	flusher.Flush()

	ctx := r.Context()
	for {
		select {
		case <-ctx.Done():
			return
		case logMsg := <-clientChan:
			fmt.Fprintf(w, "data: %s\n\n", logMsg)
			flusher.Flush()
		}
	}
}
