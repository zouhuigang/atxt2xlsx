package main

import (
	//"encoding/base64"
	//"bytes"
	//"code.google.com/p/mahonia"
	"fmt"
	"github.com/henrylee2cn/mahonia"
	"github.com/tealeg/xlsx"
	"io/ioutil"
	"log"
	"os"
	"path"
	"path/filepath"
	"strings"
	"time"
	"unicode"
)

var (
	newfile  *xlsx.File
	newsheet *xlsx.Sheet
	newrow   *xlsx.Row
	err      error
	txtFile  []byte
)

//得到文件根路径
func getCurrentDirectory() string {
	dir, err := filepath.Abs(filepath.Dir(os.Args[0]))
	if err != nil {
		log.Fatal(err)
	}
	return strings.Replace(dir, "\\", "/", -1)
}

//搜索文件夹下的所有文件目录
func getFilelist(paths string) {
	err := filepath.Walk(paths, func(paths string, f os.FileInfo, err error) error {
		if f == nil {
			return err
		}
		if f.IsDir() {
			return nil
		}
		paths = filepath.ToSlash(paths)
		filenameWithSuffix := path.Base(paths)
		//fmt.Println("filenameWithSuffix =", filenameWithSuffix)
		var fileSuffix string
		fileSuffix = path.Ext(filenameWithSuffix)

		if fileSuffix == ".TXT" {
			fmt.Println("去读文件", filenameWithSuffix)
			dofile(filenameWithSuffix)
		}

		//fmt.Println(path)
		//delFile(path)
		return nil
	})
	if err != nil {
		fmt.Printf("filepath.Walk() returned %v\n", err)
	}

}

func main() {
	//创建新表格
	newfile = xlsx.NewFile()
	newsheet, err = newfile.AddSheet("Sheet1")
	if err != nil {
		fmt.Printf(err.Error())
	}
	fmt.Println("=====================      程序开始运行,版本1.0.0      ===================\n")
	fmt.Println("======================作者:邹慧刚 邮箱:952750120@qq.com===================\n")
	fmt.Println("======================        制作日期:20170413        ===================\n")
	//得到数据目录

	testFile := getCurrentDirectory()
	getFilelist(testFile)

	newsheet.SetColWidth(1, 2, 38) //需要表格中有数据才行,值*10个像素，即380px宽,0-1为第一列,1-2为第二列
	err = newfile.Save("adata.xlsx")
	if err != nil {
		fmt.Printf(err.Error())
	}

	fmt.Println("全部任务已完成,10秒后程序将自动关闭...")
	time.Sleep(10 * time.Second)
}

func dofile(pathfile string) {

	txtFile, err = ioutil.ReadFile(pathfile)
	if err != nil {
		panic(err)
	}

	domain(txtFile)
}

func domain(txtFile []byte) {

	a_txt := strings.Replace(string(txtFile), "\r\n", "\n", -1)
	//a_txt = strings.Trim(a_txt, "\n") //去除末尾
	arr_txt := strings.Split(a_txt, "\n")

	//存数组，k-v键值对
	var RealNameArr []string = []string{}
	var InfoArr []string = []string{}
	for _, value := range arr_txt {
		if len(value) == 0 { //空行
			continue
		}
		srcCoder := mahonia.NewDecoder("GB2312") //asci转utf-8
		valueUTF8 := srcCoder.ConvertString(value)
		if IsChineseChar(valueUTF8) {
			RealNameArr = append(RealNameArr, valueUTF8)
			//addRow(valueUTF8) //中文名字
		} else {
			InfoArr = append(InfoArr, valueUTF8)
			//addRow1(valueUTF8) //其他字符
		}

	}

	if len(RealNameArr) == len(InfoArr) {
		for i := 0; i < len(InfoArr); i++ {
			//fmt.Printf("姓名%d,信息%d\n", RealNameArr[i], InfoArr[i])
			addRow(RealNameArr[i], InfoArr[i])
		}
	} else {
		fmt.Printf("姓名数组长度%d,信息数组长度%d\n", len(RealNameArr), len(InfoArr))
	}

	// output: 1 <nil>
}

/*
my program is below:
file = xlsx.NewFile()
sheet, _ = file.AddSheet("Sheet1")
row = sheet.AddRow()
row.SetHeightCM(1.3)
cell = row.AddCell()
//cell.SetStyle(style)
cell.Value = "aaa"
cell.Merge(3, 0)
sheet.SetColWidth(0, 0, 25.0)
sheet.SetColWidth(1, 4, 15.0)
*/
func addRow(nickname string, info string) {
	newrow = newsheet.AddRow()
	newrow.AddCell().SetString(nickname)
	newrow.AddCell().SetString(info)
}

func ConvertToString(src string, srcCode string, tagCode string) string {
	srcCoder := mahonia.NewDecoder(srcCode)
	srcResult := srcCoder.ConvertString(src)
	tagCoder := mahonia.NewDecoder(tagCode)
	_, cdata, _ := tagCoder.Translate([]byte(srcResult), true)
	result := string(cdata)
	return result
}

/*
判断字符串是否包含中文字符
*/
func IsChineseChar(str string) bool {
	for _, r := range str {
		if unicode.Is(unicode.Scripts["Han"], r) {
			return true
		}
	}
	return false
}
