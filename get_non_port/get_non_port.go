package getnonport

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

type ExcelInfo struct {
	line1  string
	line2  string
	line3  string
	IP     string
	PORT   string
	REMARK string
}

func ReadExcel(filename string) ([]ExcelInfo, error) {
	f, err := excelize.OpenFile(filename)
	if err != nil {
		return nil, fmt.Errorf("打开Excel文件失败:%v", err)
	}
	defer f.Close()

	rows, err := f.GetRows("Sheet1")

	if err != nil {
		return nil, fmt.Errorf("读取工作表失败:%v", err)
	}

	var infos []ExcelInfo

	for i := 1; i < len(rows); i++ {
		row := rows[i]
		if len(row) >= 3 && (row[0] != "" || row[1] != "" || row[2] != "") {
			info := ExcelInfo{
				line1:  row[0],
				line2:  row[1],
				line3:  row[2],
				IP:     "", //ip
				PORT:   "", //port
				REMARK: "", //remark
			}

			if len(row) >= 4 {
				info.IP = row[3]
			}
			if len(row) >= 5 {
				info.PORT = row[4]
			}
			if len(row) >= 8 {
				info.REMARK = row[7]
			}
			infos = append(infos, info)

		}
	}
	return infos, nil
}

func saveIt(infos []ExcelInfo, filename string) error {

}

func get_non_port(infos []ExcelInfo, filename string) error {
	var infos_non []ExcelInfo

	for _, info := range infos {
		if info.PORT == "" {
			infos_non = append(infos_non, info)
		}
	}

	saveIt(infos_non)

}
