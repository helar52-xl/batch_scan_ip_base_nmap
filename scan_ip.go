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
	OS        []string
	OSGuesses []string
	Ports     []PortInfo
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
		OS:        make([]string, 0),
		OSGuesses: make([]string, 0),
		Ports:     make([]PortInfo, 0),
	}

	// 解析操作系统信息
	osRegex := regexp.MustCompile(`OS details: (.+)`)
	if matches := osRegex.FindStringSubmatch(output); len(matches) > 1 {
		result.OS = append(result.OS, matches[1])
	}

	// 添加解析操作系统猜测信息
	osGuessRegex := regexp.MustCompile(`Aggressive OS guesses: (.+)`)
	if matches := osGuessRegex.FindStringSubmatch(output); len(matches) > 1 {
		result.OSGuesses = append(result.OSGuesses, matches[1])
	}

	// 修改解析端口信息部分
	portRegex := regexp.MustCompile(`(\d+)/(tcp|udp)\s+(\w+)\s+(.*)`)
	scanner := bufio.NewScanner(strings.NewReader(output))
	for scanner.Scan() {
		line := scanner.Text()
		if matches := portRegex.FindStringSubmatch(line); len(matches) > 1 {
			serviceInfo := matches[4]
			
			// 分离服务名称和版本信息
			service := serviceInfo
			version := ""
			
			// 如果包含空格，第一个空格前的是服务名，后面的是版本信息
			if idx := strings.Index(serviceInfo, " "); idx != -1 {
				service = strings.TrimSpace(serviceInfo[:idx])
				version = strings.TrimSpace(serviceInfo[idx+1:])
				
				// 如果有版本信息，则交换服务和版本
				if version != "" {
					service, version = version, service
				}
			}
			
			portInfo := PortInfo{
				Port:     matches[1],
				Protocol: matches[2],
				State:    matches[3],
				Service:  service,
				Version:  version,
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

// 修改readExcel函数
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
		// 修改判断逻辑：只要前三列有值就处理该行
		if len(row) >= 3 && (row[0] != "" || row[1] != "" || row[2] != "") {
			info := ExcelInfo{
				Number: row[0],
				Name:   row[1],
				Domain: row[2],
				IP:     "", // IP默认为空字符串
			}
			// 如果存在第四列（IP列）且不为空，则设置IP值
			if len(row) >= 4 {
				info.IP = row[3]
			}
			infos = append(infos, info)
		}
	}
	return infos, nil
}

// 修改exportToExcel函数，添加append参数支持追加模式
func exportToExcel(results map[string]ScanResult, sourceInfos []ExcelInfo, filename string, append bool) error {
	var f *excelize.File
	var currentRow int

	if append {
		// 如果文件存在则打开，不存在则创建新文件
		if _, statErr := os.Stat(filename); statErr == nil {
			var openErr error
			f, openErr = excelize.OpenFile(filename)
			if openErr != nil {
				return fmt.Errorf("打开Excel文件失败: %v", openErr)
			}
			// 获取最后一行的行号
			rows, _ := f.GetRows("Sheet1")
			currentRow = len(rows) + 1
		} else {
			f = excelize.NewFile()
			currentRow = 2 // 新文件从第二行开始写入数据
			// 写入表头
			headers := []string{"所属单位", "网站名称", "网站地址", "IP", "端口", "应用", "操作系统", "操作系统猜测", "备注", "协议", "版本", "状态"}
			for i, header := range headers {
				cell, _ := excelize.CoordinatesToCellName(i+1, 1)
				f.SetCellValue("Sheet1", cell, header)
			}
		}
	} else {
		f = excelize.NewFile()
		currentRow = 2
		// 写入表头
		headers := []string{"所属单位", "网站名称", "网站地址", "IP", "端口", "应用", "操作系统", "操作系统猜测", "备注", "协议", "版本", "状态"}
		for i, header := range headers {
			cell, _ := excelize.CoordinatesToCellName(i+1, 1)
			f.SetCellValue("Sheet1", cell, header)
		}
	}

	defer func() {
		if err := f.Close(); err != nil {
			fmt.Printf("关闭Excel文件时出错: %v\n", err)
		}
	}()

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
		
		osGuessInfo := " "
		if len(result.OSGuesses) > 0 {
			osGuessInfo = strings.Join(result.OSGuesses, "\n")
		}
		
		if len(result.Ports) == 0 || ip == "" {
			// 写入基本信息，其他字段留空
			f.SetCellValue("Sheet1", fmt.Sprintf("A%d", currentRow), info.Number)
			f.SetCellValue("Sheet1", fmt.Sprintf("B%d", currentRow), info.Name)
			f.SetCellValue("Sheet1", fmt.Sprintf("C%d", currentRow), info.Domain)
			f.SetCellValue("Sheet1", fmt.Sprintf("D%d", currentRow), ip)
			f.SetCellValue("Sheet1", fmt.Sprintf("E%d", currentRow), "")  // 端口为空
			f.SetCellValue("Sheet1", fmt.Sprintf("F%d", currentRow), "")  // 应用为空
			f.SetCellValue("Sheet1", fmt.Sprintf("G%d", currentRow), "")  // 操作系统为空
			f.SetCellValue("Sheet1", fmt.Sprintf("H%d", currentRow), "")  // 操作系统猜测为空
			f.SetCellValue("Sheet1", fmt.Sprintf("I%d", currentRow), "")  // 备注为空
			f.SetCellValue("Sheet1", fmt.Sprintf("J%d", currentRow), "")  // 协议为空
			f.SetCellValue("Sheet1", fmt.Sprintf("K%d", currentRow), "")  // 版本为空
			f.SetCellValue("Sheet1", fmt.Sprintf("L%d", currentRow), "")  // 状态为空
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
			f.SetCellValue("Sheet1", fmt.Sprintf("H%d", currentRow), osGuessInfo)  // 操作系统猜测
			f.SetCellValue("Sheet1", fmt.Sprintf("I%d", currentRow), "")  // 备注列
			f.SetCellValue("Sheet1", fmt.Sprintf("J%d", currentRow), port.Protocol)
			f.SetCellValue("Sheet1", fmt.Sprintf("K%d", currentRow), port.Version)
			f.SetCellValue("Sheet1", fmt.Sprintf("L%d", currentRow), port.State)
			currentRow++
		}
		
		// 合并单元格时需要包含新的操作系统猜测列
		if currentRow > startRow+1 {
			cols := []string{"A", "B", "C", "D", "G", "H", "I"}  // 添加 H 列到合并列表
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
	
	// 更新列宽设置
	columnWidths := map[int]float64{
		1: 10,  // 序号
		2: 15,  // 名称
		3: 20,  // 域名
		4: 15,  // IP地址
		5: 10,  // 端口
		6: 15,  // 服务
		7: 25,  // 操作系统
		8: 25,  // 操作系统猜测
		9: 20,  // 备注
		10: 10, // 协议
		11: 15, // 版本
		12: 10, // 状态
	}
	
	for col, width := range columnWidths {
		colName, _ := excelize.ColumnNumberToName(col)
		f.SetColWidth("Sheet1", colName, colName, width)
	}
	
	return f.SaveAs(filename)
}

// 添加新的函数用于写入单个IP的扫描结果
func appendScanResult(ip string, result ScanResult, info ExcelInfo, filename string) error {
	singleResult := make(map[string]ScanResult)
	singleResult[ip] = result
	
	singleInfo := []ExcelInfo{info}
	return exportToExcel(singleResult, singleInfo, filename, true)
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
		// 只提取有IP的行到IP列表，但保留所有sourceInfos
		for _, info := range sourceInfos {
			if info.IP != "" {
				ips = append(ips, info.IP)
			}
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

	// 创建一个空的Excel文件
	if *excelOutput != "" {
		emptyResults := make(map[string]ScanResult)
		if err := exportToExcel(emptyResults, nil, *excelOutput, false); err != nil {
			fmt.Printf("创建Excel文件时出错: %v\n", err)
			return
		}
	}

	// 按照源Excel的顺序处理所有记录
	if *sourceExcel != "" {
		for _, info := range sourceInfos {
			if info.IP == "" {
				// 对于没有IP的记录，直接写入空结果
				if *excelOutput != "" {
					if err := appendScanResult("", ScanResult{}, info, *excelOutput); err != nil {
						fmt.Printf("写入无IP记录时出错: %v\n", err)
					}
				}
			} else {
				// 对有IP的记录进行扫描
				fmt.Printf("正在扫描 %s...\n", info.IP)
				result, err := scanIP(info.IP, *nmapArgs)
				if err != nil {
					fmt.Printf("扫描 %s 时出错: %v\n", info.IP, err)
					if *excelOutput != "" {
						failedResult := ScanResult{
							OS:    []string{"扫描失败: " + err.Error()},
							Ports: []PortInfo{},
						}
						if err := appendScanResult(info.IP, failedResult, info, *excelOutput); err != nil {
							fmt.Printf("写入 %s 的失败结果时出错: %v\n", info.IP, err)
						}
					}
					continue
				}

				if *excelOutput != "" {
					if err := appendScanResult(info.IP, result, info, *excelOutput); err != nil {
						fmt.Printf("写入 %s 的扫描结果时出错: %v\n", info.IP, err)
					} else {
						fmt.Printf("%s 的扫描结果已写入文件\n", info.IP)
					}
				}
			}
		}
	} else {
		// 处理从文件或命令行参数读取的IP列表
		// ... 原有的IP列表处理代码 ...
	}

	fmt.Printf("\n所有扫描结果已保存到Excel文件: %s\n", *excelOutput)
} 