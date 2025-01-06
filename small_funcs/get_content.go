package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"os"
	"path/filepath"

	"github.com/xuri/excelize/v2"
)

// 读取Excel文件并返回第一列和第二列的键值对map
func readExcelColumns(filename string) (map[string]string, error) {
	// 打开Excel文件
	f, err := excelize.OpenFile(filename)
	if err != nil {
		return nil, fmt.Errorf("打开Excel文件失败: %v", err)
	}
	defer f.Close()

	// 获取第一个工作表中的所有行
	rows, err := f.GetRows("Sheet1")
	if err != nil {
		return nil, fmt.Errorf("读取工作表失败: %v", err)
	}

	// 创建map用于存储键值对
	result := make(map[string]string)

	// 从第二行开始读取（跳过表头）
	for i := 1; i < len(rows); i++ {
		row := rows[i]
		// 确保行至少有两列
		if len(row) >= 2 {
			key := row[0]
			value := row[1]
			// 只有当key和value都不为空时才添加到map中
			if key != "" && value != "" {
				result[key] = value
			}
		}
	}

	return result, nil
}

// 将map保存为JSON文件
func saveToJSON(data map[string]string, filename string) error {
	// 将map转换为JSON格式
	jsonData, err := json.MarshalIndent(data, "", "    ")
	if err != nil {
		return fmt.Errorf("转换JSON失败: %v", err)
	}

	// 写入文件
	err = os.WriteFile(filename, jsonData, 0644)
	if err != nil {
		return fmt.Errorf("写入JSON文件失败: %v", err)
	}

	return nil
}

func main() {
	// 定义命令行参数
	sourceExcel := flag.String("s", "", "源Excel文件路径")
	flag.Parse()

	// 检查是否提供了Excel文件路径
	if *sourceExcel == "" {
		fmt.Println("请使用 -s 参数指定Excel文件路径")
		return
	}

	// 读取Excel文件
	data, err := readExcelColumns(*sourceExcel)
	if err != nil {
		fmt.Printf("读取Excel文件失败: %v\n", err)
		return
	}

	// 生成输出JSON文件名（使用输入文件名，但改为.json扩展名）
	inputFile := *sourceExcel
	outputFile := filepath.Join(
		filepath.Dir(inputFile),
		filepath.Base(inputFile[:len(inputFile)-len(filepath.Ext(inputFile))]+".json"),
	)

	// 保存为JSON文件
	err = saveToJSON(data, outputFile)
	if err != nil {
		fmt.Printf("保存JSON文件失败: %v\n", err)
		return
	}

	fmt.Printf("处理完成！数据已保存到: %s\n", outputFile)
	fmt.Printf("共处理 %d 条记录\n", len(data))
}
