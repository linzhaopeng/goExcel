package goExcel

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"strconv"
)

const A = 65

func Export(dataList []map[string]string, headList map[string]int, fileName string) {
	f, newHeadList := excelize.NewFile(), make(map[string]string)

	// 设置表头
	for k, v := range headList{
		row := string(A + v)
		if v > 25 {
			if v < 52 {
				row = "A" + string(A + v - 26)
			} else {
				row = "B" + string(A + v - 52)
			}
		}

		f.SetCellValue("Sheet1", row + "1", k)
		newHeadList[k] = row
	}

	for dataIndex, data := range dataList{
		row := strconv.Itoa(dataIndex + 2)
		for k, v := range data{
			f.SetCellValue("Sheet1", newHeadList[k] + row, v);
		}
	}

	if err := f.SaveAs(fileName + ".xlsx"); err != nil {
		fmt.Println("save excelFile err: ", err.Error())
	}
}