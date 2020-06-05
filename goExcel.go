package goExcel

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"strconv"
)

const A = 65

func Export(dataList []map[string]string, headList []string, fileName string) {
	f, newHeadList := excelize.NewFile(), make(map[string]string)

	// set header columns
	// use uppercase from A to Z
	// if run out of alphabet,rearrange from A to Z and join another alphabet also from A to Z after it
	for k, v := range headList {
		row := string(A + k)
		if k > 25 {
			// I think 27*26 columns is enough
			times, offset := k / 26, k % 26
			row = string(A+times-1) + string(A+offset-1)
		}

		// the header column is set on row 1st
		f.SetCellValue("Sheet1", row+"1", k)
		newHeadList[v] = row
	}

	for dataIndex, data := range dataList {
		// because the header column is set on row 1st,the data's row is from 2nd
		row := strconv.Itoa(dataIndex + 2)
		for k, v := range data {
			f.SetCellValue("Sheet1", newHeadList[k]+row, v)
		}
	}

	if err := f.SaveAs(fileName + ".xlsx"); err != nil {
		fmt.Println("save excelFile err: ", err.Error())
	}
}
