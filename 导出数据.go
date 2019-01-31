package main

import (
	"fmt"
	"github.com/Luxurioust/excelize"
	"io/ioutil"
	"strconv"
	"strings"
)

func main() {

	toexcel()
}

func toexcel() {
	var path string
	path = "./spider-BaiduIndex/data/"   //填入data数据的路径，最后需要以 / 结尾
	files, _ := ioutil.ReadDir(path)

	for _, jingdian := range files {
		count1 := 1 // 避免多次插入时间(省份维度）
		var sfrow, csrow = 2, 2
		var time []string
		var csdata []int
		path := path + jingdian.Name() + "/"
		fis, _ := ioutil.ReadDir(path)

		sfxlsx := excelize.NewFile()
		csxlsx := excelize.NewFile()
		for _, shengfen := range fis {
			count2 := 1 // 避免多次插入时间（城市维度）
			var sfdata []int

			path := path + shengfen.Name() + "/"
			fs, _ := ioutil.ReadDir(path)

			for _, diqu := range fs {

				file := path + diqu.Name()
				time, csdata = getcsdata(file)
				for len(sfdata) <= len(csdata) {
					sfdata = append(sfdata, 0)
				}

				getsfdata(sfdata, csdata)
				csxlsx.SetCellValue("Sheet1", "A1", jingdian.Name())                                            //设置表首景点名
				csxlsx.SetCellValue("sheet1", Axis(csrow)+"1", shengfen.Name()+FileNameWithoutTxt(diqu.Name())) //每列数据来源
				//设置第一列为时间
				if count2 == 1 {
					inserttime(csxlsx, time)
					count2--
				}
				for i, val := range csdata {
					axis := Axis(csrow)
					axis += strconv.Itoa(i + 2)
					csxlsx.SetCellValue("Sheet1", axis, val)
				}
				csrow++
			}
			sfxlsx.SetCellValue("Sheet1", "A1", jingdian.Name())            //设置表首景点名
			sfxlsx.SetCellValue("sheet1", Axis(sfrow)+"1", shengfen.Name()) //每列数据来源
			//设置第一列为时间
			if count1 == 1 {
				inserttime(sfxlsx, time)
				count1--
			}
			for i, val := range sfdata {
				axis := Axis(sfrow)
				axis += strconv.Itoa(i + 2)
				sfxlsx.SetCellValue("Sheet1", axis, val)
			}
			sfrow++
			csxlsx.SaveAs(jingdian.Name() + "城市.xlsx")
			sfxlsx.SaveAs(jingdian.Name() + "省份.xlsx")
		}
		fmt.Println("完成一个景点")
	}
}

func inserttime(xlsx *excelize.File, time []string) {
	for i, val := range time {
		xlsx.SetCellValue("Sheet1", "A"+strconv.Itoa(i+2), val)
	}
}

func getsfdata(sfdata, csdata []int) {
	for i, val := range csdata {
		sfdata[i] += val
	}
}

func getcsdata(file string) ([]string, []int) {
	Data, err := ioutil.ReadFile(file)
	var time []string
	var count []int
	if err != nil {
		panic(err)
	}
	for _, val := range strings.Split(string(Data), "\n") {
		if len(val) >= 10 {
			time = append(time, val[:10])
			if len(val) > 14 {
				data, err := strconv.Atoi(val[14 : len(val)-1])
				if err != nil {
					count = append(count, 0)
				} else {
					count = append(count, data)
				}
			}
		}
	}
	return time, count
}

func FileNameWithoutTxt(name string) string {
	return name[0 : len(name)-4]
}

func Axis(row int) string {
	var a = [26]string{
		"Z", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y",
	}
	var axis string
	var rowslice []int
	if row <= 0 {
		return "err"
	}
	for row > 0 {
		rowslice = append(rowslice, row%26)
		row = (row - 1) / 26
	}
	l := len(rowslice)
	for l > 0 {
		axis = axis + a[rowslice[l-1]]
		l--
	}
	return axis
}
