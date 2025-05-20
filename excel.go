package exporter

import (
	"fmt"
	"io"

	"github.com/xuri/excelize/v2"
)

type ExcelExporter struct {
	file        *excelize.File
	sheetName   string
	currentRow  int
	columnsMeta *columnMeta
	headers     []Header
}

// 新建导出器
func NewExcelExporter() *ExcelExporter {
	return &ExcelExporter{
		file:       excelize.NewFile(),
		sheetName:  "Sheet1",
		currentRow: 1,
	}
}

func (e *ExcelExporter) CreateSheet(sheetName string) error {
	e.sheetName = sheetName
	e.file.NewSheet(sheetName)
	e.currentRow = 1
	return nil
}

// 保留简单表头
func (e *ExcelExporter) SetSimpleHeaders(headers []string) error {
	for col, header := range headers {
		cell, _ := excelize.CoordinatesToCellName(col+1, e.currentRow)
		e.file.SetCellValue(e.sheetName, cell, header)
	}
	e.columnsMeta = nil // 取消复杂表头映射，避免误用
	e.currentRow++
	return nil
}

// 保留简单追加数据行，按顺序写入
func (e *ExcelExporter) AppendData(data []interface{}) error {
	if e.currentRow <= 1 {
		return fmt.Errorf("请先设置表头")
	}
	// 如果是复杂表头模式，建议用 AppendRow
	if e.columnsMeta != nil {
		return fmt.Errorf("当前为复杂表头模式，请使用 AppendRow 方法")
	}
	for col, value := range data {
		cell, _ := excelize.CoordinatesToCellName(col+1, e.currentRow)
		e.file.SetCellValue(e.sheetName, cell, value)
	}
	e.currentRow++
	return nil
}

// 复杂多层表头
func (e *ExcelExporter) SetComplexHeaders(headers []Header) error {
	e.columnsMeta = &columnMeta{
		keyToColumn: make(map[string]string),
	}
	e.headers = headers

	e.generateHeaders(headers, e.currentRow, 1)

	e.currentRow = e.calculateMaxDepth(headers, 1) + 1
	return nil
}

// 计算表头最大深度
func (e *ExcelExporter) calculateMaxDepth(headers []Header, currentDepth int) int {
	maxDepth := currentDepth
	for _, h := range headers {
		if len(h.Children) > 0 {
			d := e.calculateMaxDepth(h.Children, currentDepth+1)
			if d > maxDepth {
				maxDepth = d
			}
		}
	}
	return maxDepth
}

// 计算叶子节点数量
func (e *ExcelExporter) countLeafColumns(headers []Header) int {
	count := 0
	for _, h := range headers {
		if len(h.Children) > 0 {
			count += e.countLeafColumns(h.Children)
		} else {
			count++
		}
	}
	return count
}

// 递归生成复杂表头
func (e *ExcelExporter) generateHeaders(headers []Header, row, colStart int) int {
	col := colStart
	for _, h := range headers {
		if len(h.Children) > 0 {
			colSpan := e.countLeafColumns(h.Children)

			startCell, _ := excelize.CoordinatesToCellName(col, row)
			endCell, _ := excelize.CoordinatesToCellName(col+colSpan-1, row)
			e.file.SetCellValue(e.sheetName, startCell, h.Title)
			e.file.MergeCell(e.sheetName, startCell, endCell)

			styleID, _ := e.file.NewStyle(&excelize.Style{
				Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
				Border: []excelize.Border{
					{Type: "left", Color: "000000", Style: 1},
					{Type: "top", Color: "000000", Style: 1},
					{Type: "right", Color: "000000", Style: 1},
					{Type: "bottom", Color: "000000", Style: 1},
				},
			})
			e.file.SetCellStyle(e.sheetName, startCell, endCell, styleID)

			e.generateHeaders(h.Children, row+1, col)
			col += colSpan
		} else {
			maxDepth := e.calculateMaxDepth(e.headers, 1)
			startCell, _ := excelize.CoordinatesToCellName(col, row)
			endCell, _ := excelize.CoordinatesToCellName(col, maxDepth)
			e.file.SetCellValue(e.sheetName, startCell, h.Title)
			e.file.MergeCell(e.sheetName, startCell, endCell)

			styleID, _ := e.file.NewStyle(&excelize.Style{
				Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
				Border: []excelize.Border{
					{Type: "left", Color: "000000", Style: 1},
					{Type: "top", Color: "000000", Style: 1},
					{Type: "right", Color: "000000", Style: 1},
					{Type: "bottom", Color: "000000", Style: 1},
				},
			})
			e.file.SetCellStyle(e.sheetName, startCell, endCell, styleID)

			e.columnsMeta.keyToColumn[h.Key] = startCell
			col++
		}
	}
	return col - colStart
}

// 追加复杂表头数据行，按key映射写入
func (e *ExcelExporter) AppendRow(data map[string]interface{}) error {
	if e.currentRow <= 1 {
		return fmt.Errorf("请先设置表头")
	}
	if e.columnsMeta == nil {
		return fmt.Errorf("当前不是复杂表头模式，不能使用 AppendRow")
	}
	for key, value := range data {
		if cell, ok := e.columnsMeta.keyToColumn[key]; ok {
			col, _, err := excelize.CellNameToCoordinates(cell)
			if err != nil {
				return err
			}
			newCell, _ := excelize.CoordinatesToCellName(col, e.currentRow)
			e.file.SetCellValue(e.sheetName, newCell, value)
		}
	}
	e.currentRow++
	return nil
}

// 你可以继续保留 SaveToFile、Save、GetFile 等方法
func (e *ExcelExporter) SaveToFile(path string) error {
	return e.file.SaveAs(path)
}

func (e *ExcelExporter) Save(writer io.Writer) error {
	return e.file.Write(writer)
}

func (e *ExcelExporter) GetFile() *excelize.File {
	return e.file
}
