package exporter

import (
	"bytes"
	"fmt"
	"testing"
)

func TestSimpleHeaders(t *testing.T) {
	// 初始化导出器
	ex := NewExcelExporter()
	// 设置表头
	headers := []string{"姓名", "年龄", "城市"}
	ex.SetSimpleHeaders(headers)
	// 追加数据
	data := []interface{}{"张三", 25, "北京"}
	ex.AppendData(data)
	data = []interface{}{"李四", 30, "上海"}
	ex.AppendData(data)
	var buf bytes.Buffer
	ex.Save(&buf)
	ex.SaveToFile("./test.xlsx")
}
func TestSetComplexHeaders(t *testing.T) {

	headers := []Header{
		{
			Title:    "是否有孩子",
			Key:      "marriage1",
			Children: []Header{},
		},
		{
			Title: "基本信息",
			Key:   "basic",
			Children: []Header{
				{Title: "姓名", Key: "name"},
				{Title: "年龄", Key: "age"},
				{Title: "性别", Key: "gender"},
			},
		},
		{
			Title: "联系方式",
			Children: []Header{
				{Title: "手机", Key: "phone"},
				{Title: "邮箱", Key: "email", Children: []Header{
					{Title: "QQ", Key: "qq"},
					{Title: "微信", Key: "wechat"},
				}},
			},
		},
		{
			Title:    "是否已婚",
			Key:      "marriage",
			Children: []Header{},
		},
	}
	// 创建Excel导出器
	exporter := NewExcelExporter()

	// 设置复杂表头
	if err := exporter.SetComplexHeaders(headers); err != nil {
		fmt.Println("设置表头失败:", err)
		return
	}
	// // 追加数据
	data := map[string]interface{}{
		"marriage1": "否",
		"name":      "张三",
		"age":       25,
		"gender":    "男",
		"phone":     "123456789",
		"email":     "zhangsan@example.com",
		"marriage":  "否",
		"qq":        "123456",
		"wechat":    "abcdefg",
	}
	if err := exporter.AppendRow(data); err != nil {
		fmt.Println("追加数据失败:", err)
		return
	}
	data = map[string]interface{}{
		"marriage1": "2",
		"name":      "张三1",
		"age":       251,
		"gender":    "男1",
		"phone":     "1234516789",
		"email":     "zhangsan1@example.com",
		"marriage":  "0否",
	}
	if err := exporter.AppendRow(data); err != nil {
		fmt.Println("追加数据失败:", err)
		return
	}
	// 保存Excel文件
	if err := exporter.SaveToFile("output.xlsx"); err != nil {
		fmt.Println("保存文件失败:", err)
		return
	}

	fmt.Println("Excel文件生成成功！")

}
