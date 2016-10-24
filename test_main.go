package main

import (
	"fmt"
	"os"
	"path/filepath"

	"github.com/andlabs/ui"
	"github.com/tealeg/xlsx"
)

func main() {

	err := ui.Main(func() {
		button := ui.NewButton("选择需要转化的文件")

		box := ui.NewVerticalBox()
		box.Append(button, false)
		window := ui.NewWindow("文字表转换工具 XLS --> JSON对象", 400, 200, false)
		window.SetChild(box)
		button.OnClicked(func(*ui.Button) {
			openFile := ui.OpenFile(window)
			_err := processXLSX(openFile)
			if _err != nil {
				appendLog("进程出错:")
				appendLog(_err.Error())
			}
		})
		window.OnClosing(func(*ui.Window) bool {
			ui.Quit()
			return true
		})
		window.Show()
	})
	if err != nil {
		panic(err)
	}
}

func appendLog(a ...interface{}) {
	fmt.Println(a)
}

func processXLSX(_filePath string) error {
	appendLog("选择了文件：" + _filePath)
	excelFileName := _filePath // "/Users/mapzchen/Documents/golangEx/src/github.com/Mapz/guiTest/test_xls_text.xlsx"
	dir, _ := filepath.Split(_filePath)
	appendLog("路径：" + dir)

	xlFile, err := xlsx.OpenFile(excelFileName)
	if err != nil {
		return err
	}
	sheet := xlFile.Sheets[0] //读取第个sheet
	//读取表头，顺便创建语言目录
	headRow := sheet.Rows[0]
	languageMap := make(map[int]string)

	for colIndex, cell := range headRow.Cells[1:len(headRow.Cells)] {
		_str, _err := cell.String()
		if _err != nil {
			return _err
		}
		appendLog("语言版本：" + _str)
		languageMap[colIndex] = _str
	}

	appendLog(languageMap)

	//创建导出文件
	colFileMap := make(map[int]*os.File)
	for col, ver := range languageMap {
		userFile := ver + ".js"
		_path := dir + userFile
		fout, err := os.Create(_path)
		appendLog("创建文件:" + _path)
		defer fout.Close()
		defer appendLog("完成导出:" + _path)
		if err != nil {
			appendLog(userFile, err)
			return err
		}
		fout.WriteString("module.exports = {\r\n")
		colFileMap[col] = fout

	}

	//开始读取数据
	for _, row := range sheet.Rows[1:len(sheet.Rows)] {
		_id, _err := row.Cells[0].String()
		if _err != nil {
			return _err
		}
		for col, cell := range row.Cells[1:len(row.Cells)] {
			_str, _err := cell.String()
			if _err != nil {
				return _err
			}
			colFileMap[col].WriteString("\t\"" + _id + `":"` + _str + "\",\r\n")
			fmt.Printf("%s    %s\n", _err, _str)
		}
	}

	for _, file := range colFileMap {
		file.WriteString("}")
	}

	return nil
}
