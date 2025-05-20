package exporter

import "io"

// 导出器核心接口
type Exporter interface {
	CreateSheet(sheetName string) error
	AppendData(data []interface{}) error
	SetStyle(styleCfg *StyleConfig) error
	Save(writer io.Writer) error
	SaveToFile(filePath string) error
	SetSimpleHeaders(headers []string) error     // 简单表头
	SetComplexHeaders(headers []Header) error    // 复杂表头
	AppendRow(data map[string]interface{}) error // 按字段名填充

}

// 样式配置
type StyleConfig struct {
	HeaderFontColor string
	CellBackground  string
	AutoColumnWidth bool
}

// 列位置元数据

type columnMeta struct {
	keyToColumn map[string]string
}
