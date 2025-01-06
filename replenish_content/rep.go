package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"os"

	"github.com/xuri/excelize/v2"
)

// 从JSON文件读取键值对
func readJSON(filename string) (map[string]string, error) {
	data, err := os.ReadFile(filename)
	if err != nil {
		return nil, fmt.Errorf("读取JSON文件失败: %v", err)
	}

	var result map[string]string
	if err := json.Unmarshal(data, &result); err != nil {
		return nil, fmt.Errorf("解析JSON失败: %v", err)
	}

	return result, nil
}

// 处理Excel文件
func processExcel(inputFile string, outputFile string, jsonData map[string]string) error {
	// 打开输入Excel文件
	f, err := excelize.OpenFile(inputFile)
	if err != nil {
		return fmt.Errorf("打开输入Excel文件失败: %v", err)
	}
	defer f.Close()

	// 创建新的Excel文件
	newFile := excelize.NewFile()
	defer func() {
		if err := newFile.Close(); err != nil {
			fmt.Printf("关闭Excel文件时出错: %v\n", err)
		}
	}()

	// 获取所有行
	rows, err := f.GetRows("Sheet1")
	if err != nil {
		return fmt.Errorf("读取工作表失败: %v", err)
	}

	// 处理每一行
	for rowIndex, row := range rows {
		// 确保至少有8列
		newRow := make([]string, 8)
		copy(newRow, row)

		// 如果行数据少于8列，补充空字符串
		for i := len(row); i < 8; i++ {
			newRow[i] = ""
		}

		// 处理第二列为空且第一列有值的情况
		if rowIndex > 0 && newRow[1] == "" && newRow[0] != "" {
			if value, exists := jsonData[newRow[0]]; exists {
				newRow[1] = value
			}
		}

		// 写入新文件
		for colIndex, cellValue := range newRow {
			cell, _ := excelize.CoordinatesToCellName(colIndex+1, rowIndex+1)
			if err := newFile.SetCellValue("Sheet1", cell, cellValue); err != nil {
				return fmt.Errorf("写入单元格失败: %v", err)
			}
		}
	}

	// 设置列宽
	columnWidths := map[int]float64{
		1: 30, // 第一列
		2: 30, // 第二列
		3: 30, // 第三列
		4: 15, // 第四列
		5: 15, // 第五列
		6: 20, // 第六列
		7: 20, // 第七列
		8: 20, // 第八列
	}

	// 应用列宽设置
	for col, width := range columnWidths {
		colName, _ := excelize.ColumnNumberToName(col)
		if err := newFile.SetColWidth("Sheet1", colName, colName, width); err != nil {
			return fmt.Errorf("设置列宽失败: %v", err)
		}
	}

	// 保存新文件
	if err := newFile.SaveAs(outputFile); err != nil {
		return fmt.Errorf("保存Excel文件失败: %v", err)
	}

	return nil
}

func main() {
	// 定义命令行参数
	jsonFile := flag.String("j", "", "输入JSON文件路径")
	sourceExcel := flag.String("s", "", "输入Excel文件路径")
	outputExcel := flag.String("e", "", "输出Excel文件路径")
	flag.Parse()

	// 检查必要参数
	if *jsonFile == "" || *sourceExcel == "" || *outputExcel == "" {
		fmt.Println("请提供所有必要参数:")
		fmt.Println("-j JSON文件路径")
		fmt.Println("-s 输入Excel文件路径")
		fmt.Println("-e 输出Excel文件路径")
		return
	}

	// 读取JSON文件
	jsonData, err := readJSON(*jsonFile)
	if err != nil {
		fmt.Printf("读取JSON文件失败: %v\n", err)
		return
	}

	// 处理Excel文件
	if err := processExcel(*sourceExcel, *outputExcel, jsonData); err != nil {
		fmt.Printf("处理Excel文件失败: %v\n", err)
		return
	}

	fmt.Println("处理完成！")
}
