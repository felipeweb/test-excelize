package main

import (
	"flag"
	"fmt"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/nuveo/log"
)

func main() {
	fileName := flag.String("name", "", "file path")
	flag.Parse()
	if *fileName == "" {
		log.Fatal("file name is required")
	}
	body, err := toCSV(*fileName)
	if err != nil {
		log.Errorln(err)
		return
	}
	log.Printf("Body: %s", body)
}

func toCSV(fileName string) (body string, err error) {
	xlsFile, err := excelize.OpenFile(fileName)
	if err != nil {
		err = fmt.Errorf("error convert with 'excelize' on xlsx parsing: %v", err)
		return
	}
	for sheetIndex := range xlsFile.GetSheetMap() {
		fmt.Println(">>>>>", xlsFile.GetSheetMap())
		rows := xlsFile.GetRows("sheet" + strconv.Itoa(sheetIndex))
		if len(rows) == 0 {
			err = fmt.Errorf("error on convert with excelize")
			return
		}
		fmt.Println("rows >>>>>>", len(rows))
		for rowIndex, row := range rows {
			for cellIndex, colCell := range row {
				text := colCell
				body += text
				if cellIndex < len(row)-1 {
					body += ","
				}
			}
			if rowIndex < len(rows)-1 {
				body += "\n"
			}
		}
	}
	return
}
