package main

import (
	"fmt"
	"os/exec"
	"strings"
	"time"
)

type ScanResult struct {
    IP            string
    Ports         []string
    Services      []string
    Versions      []string
    OS            string
    OSGuess       string
}

func performNmapScan(ip string, nmapCmd string) (*ScanResult, error) {
    args := strings.Split(nmapCmd, " ")
    args = append(args, ip)
    
    start := time.Now()
    cmd := exec.Command("nmap", args...)
    output, err := cmd.CombinedOutput()
    if err != nil {
        return nil, err
    }
    
    duration := time.Since(start)
    
    result := &ScanResult{
        IP: fmt.Sprintf("%s (扫描用时: %v)", ip, duration),
    }
    
    // 解析nmap输出
    lines := strings.Split(string(output), "\n")
    for _, line := range lines {
        if strings.Contains(line, "open") {
            parts := strings.Fields(line)
            if len(parts) >= 3 {
                result.Ports = append(result.Ports, parts[0])
                result.Services = append(result.Services, parts[2])
                if len(parts) > 3 {
                    result.Versions = append(result.Versions, strings.Join(parts[3:], " "))
                }
            }
        }
        if strings.Contains(line, "OS details:") {
            result.OS = strings.TrimPrefix(line, "OS details: ")
        }
    }
    
    return result, nil
} 