package exporter

type Header struct {
	Title    string   `json:"title"`
	Children []Header `json:"children"`
	Key      string   `json:"key"`
}
