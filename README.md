# Excel Exporter

一个基于 [excelize](https://github.com/qax-os/excelize) 的 Go 语言 Excel 导出工具，支持：

- 简单表头导出
- 多层级表头（任意深度嵌套）
- 单元格合并、自动列宽设置、居中样式
- 文件保存到本地或通过流导出
- 简洁易用的 API 接口设计

---

## ✨ 功能亮点

- ✅ `SetSimpleHeaders([]string)` 简单表头设置
- ✅ `SetComplexHeaders([]Header)` 支持无限嵌套表头结构
- ✅ `AppendData([]interface{})` 追加简单行数据
- ✅ `AppendRow(map[string]interface{})` 对应复杂表头的数据行追加
- ✅ 自动合并单元格，设置列宽及样式

---

## 🧱 安装依赖

```bash
go get github.com/xuri/excelize/v2
```
## 📦 使用方式
### 简单导出
```go
exporter := NewExcelExporter()
exporter.SetSimpleHeaders([]string{"姓名", "年龄", "性别"})
exporter.AppendData([]interface{}{"张三", 28, "男"})
exporter.SaveToFile("simple.xlsx")
```
### 多层级表头导出
```go
headers := []Header{
    {
        Title: "基本信息",
        Key:   "base",
        Children: []Header{
            {
                Title: "个人",
                Key:   "personal",
                Children: []Header{
                    {Title: "姓名", Key: "name"},
                    {Title: "性别", Key: "gender"},
                },
            },
            {Title: "年龄", Key: "age"},
        },
    },
    {Title: "是否有孩子", Key: "hasChild"},
}

exporter := NewExcelExporter()
exporter.SetComplexHeaders(headers)

data := map[string]interface{}{
    "name":     "张三",
    "gender":   "男",
    "age":      30,
    "hasChild": "是",
}
exporter.AppendRow(data)
exporter.SaveToFile("complex.xlsx")
```
