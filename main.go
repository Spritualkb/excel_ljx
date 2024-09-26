package main

import (
	"fmt"
	"log"
	"os"
	"strings"

	"github.com/xuri/excelize/v2"
)

// 打开Excel文件
func openExcel(filename string) (*excelize.File, error) {
	f, err := excelize.OpenFile(filename)
	if err != nil {
		return nil, err
	}
	return f, nil
}

// 填充第一行空单元格
func fillEmptyCellsInFirstRow(f *excelize.File, sheetName string) {
	for col := 1; col <= 45; col++ {
		cellAddr, _ := excelize.CoordinatesToCellName(col, 1) // 第1行的单元格
		cellValue, err := f.GetCellValue(sheetName, cellAddr)
		if err != nil {
			log.Fatalf("读取单元格失败: %v", err)
		}
		if strings.TrimSpace(cellValue) == "" {
			nextCellAddr, _ := excelize.CoordinatesToCellName(col, 2)
			nextCellValue, _ := f.GetCellValue(sheetName, nextCellAddr)
			if nextCellValue != "" {
				err := f.SetCellValue(sheetName, cellAddr, nextCellValue)
				if err != nil {
					return
				}
				fmt.Printf("第%d列，第一行为空格，已赋值为: %s\n", col, nextCellValue)
			}
		}
	}
}

// 查找指定列并插入新列标题
func insertHeaders(f *excelize.File, sheetName string, targetColName string, newHeaders []string) {
	rows, err := f.GetRows(sheetName)
	if err != nil {
		log.Fatalf("读取工作表失败: %v", err)
	}

	// 查找目标列
	colIndex := -1
	for i, col := range rows[0] {
		if col == targetColName {
			colIndex = i + 1 // Excel列从1开始计数
			break
		}
	}

	if colIndex == -1 {
		fmt.Printf("未找到列 '%s'\n", targetColName)
		return
	}

	// 插入新的列标题
	for i, header := range newHeaders {
		cellAddr, _ := excelize.CoordinatesToCellName(colIndex+i, 1)
		err := f.SetCellValue(sheetName, cellAddr, header)
		if err != nil {
			return
		}
	}
	fmt.Println("列拼接和填充完成！")
}

// 根据给定列计算和填充公式
func fillFormulasOrCopy(f *excelize.File, sheetName string, singleColumns map[string]string, sumColumns map[string][]string) {
	rows, err := f.GetRows(sheetName)
	if err != nil {
		log.Fatalf("读取工作表失败: %v", err)
	}

	// 处理需要求和的列
	for targetCol, inputCols := range sumColumns {
		var inputColIndices []int
		for _, colName := range inputCols {
			for i, col := range rows[0] {
				if col == colName {
					inputColIndices = append(inputColIndices, i+1)
				}
			}
		}

		// 查找目标列的索引
		targetColIndex := -1
		for i, col := range rows[0] {
			if col == targetCol {
				targetColIndex = i + 1
				break
			}
		}

		if targetColIndex == -1 {
			fmt.Printf("未找到目标列 '%s'\n", targetCol)
			continue
		}

		// 填充求和公式
		lastRow := len(rows)
		for row := 2; row <= lastRow; row++ {
			cellAddr, _ := excelize.CoordinatesToCellName(targetColIndex, row)
			var formulaParts []string
			for _, colIndex := range inputColIndices {
				colAddr, _ := excelize.CoordinatesToCellName(colIndex, row)
				formulaParts = append(formulaParts, colAddr)
			}
			formula := fmt.Sprintf("SUM(%s)", strings.Join(formulaParts, ","))
			err := f.SetCellFormula(sheetName, cellAddr, formula)
			if err != nil {
				return
			}
		}
	}

	// 处理直接复制的列
	for targetCol, inputCol := range singleColumns {
		var inputColIndex int
		for i, col := range rows[0] {
			if col == inputCol {
				inputColIndex = i + 1
				break
			}
		}

		targetColIndex := -1
		for i, col := range rows[0] {
			if col == targetCol {
				targetColIndex = i + 1
				break
			}
		}

		if targetColIndex == -1 {
			fmt.Printf("未找到目标列 '%s'\n", targetCol)
			continue
		}

		// 填充直接复制公式
		lastRow := len(rows)
		for row := 2; row <= lastRow; row++ {
			cellAddr, _ := excelize.CoordinatesToCellName(targetColIndex, row)
			inputAddr, _ := excelize.CoordinatesToCellName(inputColIndex, row)
			formula := fmt.Sprintf("=%s", inputAddr)
			err := f.SetCellFormula(sheetName, cellAddr, formula)
			if err != nil {
				return
			}
		}
	}
}

// 创建透视表工作表
func createTable(f *excelize.File, sheetName string) {

	// 创建新工作表
	index, err := f.NewSheet(sheetName)

	if err != nil {
		fmt.Println(err)
		return
	}

	// 设置新工作表为当前显示的工作表
	f.SetActiveSheet(index)

	fmt.Println("新建" + sheetName + "表成功!")
}

// 获取数据范围
func getSheetDataRange(f *excelize.File, sheetName string) (string, error) {
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return "", err
	}

	// 判断数据是否为空
	if len(rows) == 0 {
		return "", fmt.Errorf("工作表 %s 没有数据", sheetName)
	}

	// 假设数据从 A1 开始
	startRow, startCol := 1, 1
	endRow, endCol := 0, 0

	// 遍历所有行
	for rowIndex, row := range rows {
		if len(row) > 0 {
			endRow = rowIndex + 1 // 行号，从 1 开始

			for colIndex, cell := range row {
				if cell != "" { // 非空单元格
					if colIndex+1 > endCol {
						endCol = colIndex + 1
					}
				}
			}
		}
	}

	// 将列号转换为 Excel 列字母
	startCell, _ := excelize.CoordinatesToCellName(startCol, startRow)
	endCell, _ := excelize.CoordinatesToCellName(endCol, endRow)

	// 返回数据范围字符串
	dataRange := fmt.Sprintf("%s!%s:%s", sheetName, startCell, endCell)
	return dataRange, nil
}

func main() {
	// 获取拖入的 Excel 文件路径
	if len(os.Args) < 2 {
		fmt.Println("请将 Excel 文件拖到此程序上运行。")
		return
	}

	filename := os.Args[1]
	f, err := openExcel(filename)
	if err != nil {
		log.Fatalf("无法打开文件: %v", err)
	}
	defer f.SaveAs(filename)

	// 获取工作表
	sheetName := "Sheet1"

	// 填充第一行没有数据的单元格
	fillEmptyCellsInFirstRow(f, sheetName)

	// 插入列标题
	newHeaders := []string{"大于等于1826天金额", "1年内", "1到2", "2到3", "3到4", "4到5", "5以上"}
	insertHeaders(f, sheetName, "大于等于1826天金额", newHeaders)

	// 填充列数据：分为求和的列和直接复制的列
	singleColumns := map[string]string{
		"1到2": "366-730天金额",
		"2到3": "731-1095天金额",
		"3到4": "1096-1460天金额",
		"4到5": "1461-1825天金额",
		"5以上": "大于等于1826天金额",
	}

	sumColumns := map[string][]string{
		"1年内": {"180天内金额", "181-365天金额"},
	}

	// 处理列公式
	fillFormulasOrCopy(f, sheetName, singleColumns, sumColumns)

	// 创建新工作表 '透视表1'
	createTable(f, "透视表1")

	// 获取 Sheet1 的数据范围
	dataRange, err := getSheetDataRange(f, sheetName)
	if err != nil {
		log.Fatalf("获取数据范围失败: %v", err)
	}

	// 输出数据范围
	fmt.Printf("数据范围为: %s\n", dataRange)

	if err := f.AddPivotTable(&excelize.PivotTableOptions{
		DataRange:       dataRange,
		PivotTableRange: "透视表1!A1:AX5538",
		Rows: []excelize.PivotTableField{
			{Data: "评估分类", DefaultSubtotal: true},
			{Data: "ABC 类", DefaultSubtotal: true}},
		//Filter: []excelize.PivotTableField{
		//	{Data: "Region"}},
		//Columns: []excelize.PivotTableField{
		//
		//},
		Data: []excelize.PivotTableField{
			{Data: "1年内", Name: "求和项：1年内", Subtotal: "Sum"},
			{Data: "1到2", Name: "求和项：1到2", Subtotal: "Sum"},
			{Data: "2到3", Name: "求和项：2到3", Subtotal: "Sum"},
			{Data: "3到4", Name: "求和项：3到4", Subtotal: "Sum"},
			{Data: "4到5", Name: "求和项：4到5", Subtotal: "Sum"},
			{Data: "5以上", Name: "求和项：5以上", Subtotal: "Sum"},
		},
		RowGrandTotals: true,
		ColGrandTotals: true,
		ShowDrill:      true,
		ShowRowHeaders: true,
		ShowColHeaders: true,
		ShowLastColumn: true,
	}); err != nil {
		fmt.Println(err)
		return
	}

	fmt.Println("Excel 处理完成")
}
