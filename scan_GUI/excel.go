package main

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

type Record struct {
    Organization string
    WebsiteName  string
    WebsiteAddr  string
    IP           string
    Port         string
    Protocol     string
    Service      string
    Version      string
    Status       string
    OS           string
    Notes        string
    OSGuess      string
}

func readExcelFile(filePath string) ([]Record, error) {
    f, err := excelize.OpenFile(filePath)
    if err != nil {
        return nil, err
    }
    defer f.Close()
    
    rows, err := f.GetRows("Sheet1")
    if err != nil {
        return nil, err
    }
    
    var records []Record
    for i, row := range rows {
        if i == 0 {
            continue // 跳过标题行
        }
        if len(row) < 12 {
            continue
        }
        record := Record{
            Organization: row[0],
            WebsiteName:  row[1],
            WebsiteAddr:  row[2],
            IP:          row[3],
            Port:        row[4],
            Protocol:    row[5],
            Service:     row[6],
            Version:     row[7],
            Status:      row[8],
            OS:          row[9],
            Notes:       row[10],
            OSGuess:     row[11],
        }
        records = append(records, record)
    }
    
    return records, nil
}

func appendToExcel(filePath string, records []Record) error {
    f := excelize.NewFile()
    
    // 写入标题行
    headers := []string{"所属单位", "网站名称", "网站地址", "IP", "端口", "协议", 
                       "服务", "版本", "状态", "操作系统", "备注", "操作系统猜测"}
    for i, header := range headers {
        cell := fmt.Sprintf("%c1", 'A'+i)
        f.SetCellValue("Sheet1", cell, header)
    }
    
    // 写入数据
    for i, record := range records {
        row := i + 2
        f.SetCellValue("Sheet1", fmt.Sprintf("A%d", row), record.Organization)
        f.SetCellValue("Sheet1", fmt.Sprintf("B%d", row), record.WebsiteName)
        f.SetCellValue("Sheet1", fmt.Sprintf("C%d", row), record.WebsiteAddr)
        f.SetCellValue("Sheet1", fmt.Sprintf("D%d", row), record.IP)
        f.SetCellValue("Sheet1", fmt.Sprintf("E%d", row), record.Port)
        f.SetCellValue("Sheet1", fmt.Sprintf("F%d", row), record.Protocol)
        f.SetCellValue("Sheet1", fmt.Sprintf("G%d", row), record.Service)
        f.SetCellValue("Sheet1", fmt.Sprintf("H%d", row), record.Version)
        f.SetCellValue("Sheet1", fmt.Sprintf("I%d", row), record.Status)
        f.SetCellValue("Sheet1", fmt.Sprintf("J%d", row), record.OS)
        f.SetCellValue("Sheet1", fmt.Sprintf("K%d", row), record.Notes)
        f.SetCellValue("Sheet1", fmt.Sprintf("L%d", row), record.OSGuess)
    }
    
    return f.SaveAs(filePath)
} 