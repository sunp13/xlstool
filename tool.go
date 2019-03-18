package xlstool

import (
	"fmt"

	"github.com/tealeg/xlsx"
)

// OpenFile 打开xls文件,返回行数组
func OpenFile(fileName string, sheetIndex ...int) (data []map[string]string, err error) {

	f, err := xlsx.OpenFile(fileName)
	if err != nil {
		return nil, err
	}
	sheetIdx := 0
	if len(sheetIndex) > 0 {
		sheetIdx = sheetIndex[0]
	}

	sheet := f.Sheets[sheetIdx]
	// sheetName := sheet.Name

	sheetTitle := make([]string, 0)
	sheetRows := make([]map[string]string, 0)

	for i, row := range sheet.Rows {
		data := make(map[string]string, 0)
		// 行号
		data["_index"] = fmt.Sprintf("%d", i+1)
		for j, cell := range row.Cells {
			if i == 0 {
				sheetTitle = append(sheetTitle, cell.String())
			} else {
				// 判断一下cell.type
				data[sheetTitle[j]] = fmt.Sprintf("%v", cell.Value)
			}
		}
		sheetRows = append(sheetRows, data)
	}
	return sheetRows, nil
}
