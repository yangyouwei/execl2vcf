package main

import (
	"bytes"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"io/ioutil"
	"os"
	"text/template"
)

var user string = `BEGIN:VCARD
N;CHARSET=UTF-8:{{.Name}}
FN;CHARSET=UTF-8:{{.FullNane}}
ORG;CHARSET=UTF-8:{{.Org}}
TEL;TYPE=CELL:{{.Tel}}
END:VCARD
`

type userstruct struct {
	Name     string
	FullNane string
	Org      string
	Tel      string
}

func main() {
	fulldata := readexecl()
	//fmt.Println(fulldata)
	Genuserfile(fulldata)
}

//读取execl
func readexecl() []userstruct {
	f, err := excelize.OpenFile("./客户.xlsx")
	if err != nil {
		fmt.Println(err)
		panic("读取execl失败")
	}
	a := f.GetSheetName(1)
	var s []userstruct
	fulluser := f.GetRows(a)
	for _, oneuser := range fulluser {
		OneUSER := userstruct{Name: oneuser[1], FullNane: oneuser[1], Org: oneuser[0], Tel: oneuser[2]}
		s = append(s, OneUSER)
	}
	return s
}

//读取模板文件
func ReadTpl(p string) string {
	b, err := ioutil.ReadFile(p)
	if err != nil {
		fmt.Println(err)
		panic("读取模板文件失败")
	}
	return string(b)
}

//根据模板生成文件
func Genuserfile(s []userstruct) {
	//创建文件
	fd, err := os.OpenFile("./user.vcf", os.O_WRONLY|os.O_CREATE|os.O_APPEND, 0644)
	if err != nil {
		fmt.Println(err)
	}

	//使用模板解析
	for _, u := range s {
		templateText := ReadTpl("./vcf-tpl")
		buffer := new(bytes.Buffer)
		t := template.Must(template.New("user").Parse(templateText))
		err := t.Execute(buffer, u)
		if err != nil {
			fmt.Println(err)
			panic("解析模板失败")
		}
		fd.WriteString(buffer.String())
	}
}
