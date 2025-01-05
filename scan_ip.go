package main

import (
	"bufio"
	"flag"
	"fmt"
	"os"
	"os/exec"
	"regexp"
	"strings"

	"github.com/xuri/excelize/v2"
)

type ScanResult struct {
	OS    []string
	Ports []PortInfo
}

type PortInfo struct {
	Port     string
	Protocol string
	Service  string
	Version  string
	State    string
}

// 添加新的结构体用于存储Excel中的信息
type ExcelInfo struct {
	Number   string
	Name     string
	Domain   string
	IP       string
}

func parseNmapOutput(output string) ScanResult {
	result := ScanResult{
		OS:    make([]string, 0),
		Ports: make([]PortInfo, 0),
	}

	// 解析操作系统信息
	osRegex := regexp.MustCompile(`OS details: (.+)`)
	if matches := osRegex.FindStringSubmatch(output); len(matches) > 1 {
		result.OS = append(result.OS, matches[1])
	}

	// 修改解析端口信息部分
	portRegex := regexp.MustCompile(`(\d+)/(tcp|udp)\s+(\w+)\s+(.*)`)
	scanner := bufio.NewScanner(strings.NewReader(output))
	for scanner.Scan() {
		line := scanner.Text()
		if matches := portRegex.FindStringSubmatch(line); len(matches) > 1 {
			service := matches[4]
			
			// 处理服务信息
			if len(service) >= 23 {
				// 如果长度大于等于23，检查第23位是否为空
				if service[22] != ' ' {
					// 如果第23位非空，保留第23位及后面的内容
					service = service[22:]
				}
			} else {
				// 长度小于23，直接去除末尾的问号
				service = strings.TrimSuffix(service, "?")
			}
			
			portInfo := PortInfo{
				Port:     matches[1],
				Protocol: matches[2],
				State:    matches[3],
				Service:  strings.TrimSpace(service), // 确保去除首尾空格
			}
			result.Ports = append(result.Ports, portInfo)
		}
	}

	return result
}

func scanIP(ip string, nmapArgs string) (ScanResult, error) {
	args := append(strings.Split(nmapArgs, " "), ip)
	cmd := exec.Command("nmap", args...)
	
	output, err := cmd.CombinedOutput()
	if err != nil {
		return ScanResult{}, fmt.Errorf("扫描错误: %v", err)
	}
	
	// 解析扫描结果
	result := parseNmapOutput(string(output))
	
	// 输出格式化结果
	fmt.Printf("\n%s\n", strings.Repeat("=", 50))
	fmt.Printf("IP地址: %s\n\n", ip)
	
	// 输出操作系统信息
	fmt.Println("操作系统:")
	if len(result.OS) > 0 {
		for _, os := range result.OS {
			fmt.Printf("- %s\n", os)
		}
	} else {
		fmt.Println("- 未检测到操作系统")
	}
	
	// 输出端口信息
	fmt.Println("\n端口信息:")
	for _, port := range result.Ports {
		fmt.Printf("- %s/%s: %s %s\n",
			port.Port,
			port.Protocol,
			port.Service,
			port.State)
	}
	
	return result, nil
}

// 添加读取Excel的函数
func readExcel(filename string) ([]ExcelInfo, error) {
	f, err := excelize.OpenFile(filename)
	if err != nil {
		return nil, fmt.Errorf("打开Excel文件失败: %v", err)
	}
	defer f.Close()

	rows, err := f.GetRows("Sheet1")
	if err != nil {
		return nil, fmt.Errorf("读取工作表失败: %v", err)
	}

	var infos []ExcelInfo
	// 跳过表头
	for i := 1; i < len(rows); i++ {
		row := rows[i]
		if len(row) >= 4 {
			info := ExcelInfo{
				Number: row[0],
				Name:   row[1],
				Domain: row[2],
				IP:     row[3],
			}
			infos = append(infos, info)
		}
	}
	return infos, nil
}

// 修改exportToExcel函数
func exportToExcel(results map[string]ScanResult, sourceInfos []ExcelInfo, filename string) error {
	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Printf("关闭Excel文件时出错: %v\n", err)
		}
	}()
	
	// 修改表头顺序，添加新列
	headers := []string{"序号", "名称", "域名", "IP地址", "端口", "服务", "操作系统", "备注", "协议", "版本", "状态"}
	for i, header := range headers {
		cell, _ := excelize.CoordinatesToCellName(i+1, 1)
		f.SetCellValue("Sheet1", cell, header)
	}
	
	currentRow := 2
	// 使用map存储IP对应的信息，方便查找
	infoMap := make(map[string]ExcelInfo)
	for _, info := range sourceInfos {
		infoMap[info.IP] = info
	}

	for ip, result := range results {
		startRow := currentRow
		info := infoMap[ip]
		osInfo := " "
		if len(result.OS) > 0 {
			osInfo = strings.Join(result.OS, "\n")
		}
		
		if len(result.Ports) == 0 {
			// 如果没有端口数据，仍然记录基本信息
			f.SetCellValue("Sheet1", fmt.Sprintf("A%d", currentRow), info.Number)
			f.SetCellValue("Sheet1", fmt.Sprintf("B%d", currentRow), info.Name)
			f.SetCellValue("Sheet1", fmt.Sprintf("C%d", currentRow), info.Domain)
			f.SetCellValue("Sheet1", fmt.Sprintf("D%d", currentRow), ip)
			f.SetCellValue("Sheet1", fmt.Sprintf("G%d", currentRow), osInfo)
			currentRow++
			continue
		}
		
		for _, port := range result.Ports {
			f.SetCellValue("Sheet1", fmt.Sprintf("A%d", currentRow), info.Number)
			f.SetCellValue("Sheet1", fmt.Sprintf("B%d", currentRow), info.Name)
			f.SetCellValue("Sheet1", fmt.Sprintf("C%d", currentRow), info.Domain)
			f.SetCellValue("Sheet1", fmt.Sprintf("D%d", currentRow), ip)
			f.SetCellValue("Sheet1", fmt.Sprintf("E%d", currentRow), port.Port)
			f.SetCellValue("Sheet1", fmt.Sprintf("F%d", currentRow), port.Service)
			f.SetCellValue("Sheet1", fmt.Sprintf("G%d", currentRow), osInfo)
			f.SetCellValue("Sheet1", fmt.Sprintf("H%d", currentRow), "")  // 备注列
			f.SetCellValue("Sheet1", fmt.Sprintf("I%d", currentRow), port.Protocol)
			f.SetCellValue("Sheet1", fmt.Sprintf("J%d", currentRow), port.Version)
			f.SetCellValue("Sheet1", fmt.Sprintf("K%d", currentRow), port.State)
			currentRow++
		}
		
		// 合并单元格
		if currentRow > startRow+1 {
			// 合并序号、名称、域名、IP地址、操作系统和备注列
			cols := []string{"A", "B", "C", "D", "G", "H"}
			for _, col := range cols {
				f.MergeCell("Sheet1", fmt.Sprintf("%s%d", col, startRow), 
							fmt.Sprintf("%s%d", col, currentRow-1))
			}
			
			// 设置单元格样式
			style, _ := f.NewStyle(&excelize.Style{
				Alignment: &excelize.Alignment{
					Vertical:   "center",
					WrapText:  true,
				},
			})
			for _, col := range cols {
				f.SetCellStyle("Sheet1", fmt.Sprintf("%s%d", col, startRow), 
								fmt.Sprintf("%s%d", col, currentRow-1), style)
			}
		}
	}
	
	// 设置列宽
	columnWidths := map[int]float64{
		1: 10,  // 序号
		2: 15,  // 名称
		3: 20,  // 域名
		4: 15,  // IP地址
		5: 10,  // 端口
		6: 15,  // 服务
		7: 25,  // 操作系统
		8: 20,  // 备注
		9: 10,  // 协议
		10: 15, // 版本
		11: 10, // 状态
	}
	
	for col, width := range columnWidths {
		colName, _ := excelize.ColumnNumberToName(col)
		f.SetColWidth("Sheet1", colName, colName, width)
	}
	
	return f.SaveAs(filename)
}

func main() {
	// 添加新的命令行参数
	sourceExcel := flag.String("s", "", "源Excel文件路径")
	filePath := flag.String("f", "", "包含IP列表的文件路径")
	ipList := flag.String("i", "", "IP地址列表，用逗号分隔")
	nmapArgs := flag.String("a", "-sV -O", "nmap扫描参数")
	excelOutput := flag.String("e", "", "输出结果到Excel文件")
	flag.Parse()

	var ips []string
	var sourceInfos []ExcelInfo
	var err error

	// 从Excel文件读取信息
	if *sourceExcel != "" {
		sourceInfos, err = readExcel(*sourceExcel)
		if err != nil {
			fmt.Printf("读取Excel文件失败: %v\n", err)
			return
		}
		// 提取IP列表
		for _, info := range sourceInfos {
			ips = append(ips, info.IP)
		}
	} else {
		// 原有的IP列表读取逻辑保持不变
		// 从文件读取IP
		if *filePath != "" {
			file, err := os.Open(*filePath)
			if err != nil {
				fmt.Printf("无法打开文件: %v\n", err)
				return
			}
			defer file.Close()
			
			scanner := bufio.NewScanner(file)
			for scanner.Scan() {
				if ip := strings.TrimSpace(scanner.Text()); ip != "" {
					ips = append(ips, ip)
				}
			}
		} else if *ipList != "" {
			// 从命令行参数获取IP
			for _, ip := range strings.Split(*ipList, ",") {
				if ip := strings.TrimSpace(ip); ip != "" {
					ips = append(ips, ip)
				}
			}
		} else {
			fmt.Println("请提供扫描内容")
			return
		}
	}

	// 收集所有扫描结果
	results := make(map[string]ScanResult)
	for _, ip := range ips {
		fmt.Printf("正在扫描 %s...\n", ip)
		result, err := scanIP(ip, *nmapArgs)
		if err != nil {
			fmt.Printf("扫描 %s 时出错: %v\n", ip, err)
			continue
		}
		results[ip] = result
	}

	// 输出到Excel
	if *excelOutput != "" {
		if err := exportToExcel(results, sourceInfos, *excelOutput); err != nil {
			fmt.Printf("保存Excel文件时出错: %v\n", err)
		} else {
			fmt.Printf("\n结果已保存到Excel文件: %s\n", *excelOutput)
		}
	}
} 