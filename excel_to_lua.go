package main

import (
	"fmt"
	"github.com/tealeg/xlsx"
	"strings"
	"strconv"
	"io/ioutil"
	"os"
)
func main() {
	arr := os.Args
	len := len(arr)
	if len<2{
		fmt.Println("args Len must be 2 , \n Usage: nb.exe c:/a/in/ c:/a/out/")
		return
	}
	inpath := arr[1]
	outpath := arr[2]
	files,_:=ioutil.ReadDir(inpath)
	for _,file := range files{
		if strings.HasPrefix(file.Name(),"~") || strings.Contains(file.Name(),"$"){
			continue
		}
		if !strings.HasSuffix(file.Name(),".xlsx"){
			continue
		}
		fmt.Println("process:"+inpath+file.Name())
		once(inpath,file.Name(),outpath)
	}
}

func getCol(sheet *xlsx.Sheet) int{
		r:=0
		for _, cell := range sheet.Rows[0].Cells {
			text := strings.TrimSpace(cell.String())
			if len(text)>0{
				r+=1
			}
		}
		return r
}

func once(path string,file string,outputPath string){
	xlFile, err := xlsx.OpenFile(path+file)
	if err != nil {
		fmt.Println("open file error")
	}
	sheet := xlFile.Sheets[0]
	fileNameArr := strings.Split(file,".")
	outputFileName := fileNameArr[0]+".lua"
	//celLen := len(sheet.Cols)
	celLen := getCol(sheet)
	var field= make([]string, celLen)
	var types= make([]string, celLen)

	var fieldClient= make([]interface{}, celLen)
	//cbody[0] = fieldClient

	for idxRow, row := range sheet.Rows {
		if idxRow == 0 || idxRow == 1 || idxRow == 2 {
			for cellIdx, cell := range row.Cells {
				if cellIdx>=celLen{
					continue
				}
				text := strings.TrimSpace(cell.String())
				if idxRow == 1 {
					field[cellIdx] = text
					fieldClient[cellIdx] = text
					continue
				}
				if (idxRow == 2) {
					types[cellIdx] = text
					continue
				}
				if (idxRow == 0) {
					continue
				}
			}
			continue
		}
	}
	 d := [][]string{}
	for i, row := range sheet.Rows {
		if i == 0 || i == 1 || i == 2 {
			continue
		}
		sl := make([]string,0,celLen)
		for j, cell := range row.Cells {
			if j>=celLen{
				continue
			}
			itemCell :=""
			if types[j] == "table" {
				itemCell = strings.TrimSpace(cell.String())
				itemCell = strings.Replace(itemCell,"[","",-1)
				itemCell = strings.Replace(itemCell,"]","",-1)
			} else {
				itemCell=strings.TrimSpace(cell.String())
			}
			sl = append(sl, itemCell)
		}
		d = append(d, sl)
	}

	c0 := make([]string,0,celLen)
	c1 := make([]string,0,celLen)
	r :=""
	for i := 0;i<celLen;i++{
		if len(r)>0{
			r+=","
		}
		itemStr := "\n\t[\""+field[i]+"\"]={"
		isStr,isTable := false,false

		if types[i]=="string"{
			isStr = true
		}
		if types[i]=="table"{
			isTable = true
		}
		dlen := len(d)
		for loop,data := range d{
			itemData:=data[i]
			if isStr{
				itemStr += "\""
				itemStr += itemData
				itemStr += "\""
			}else if isTable{
				tr := "{"
				arr := strings.Split(itemData,"|")
				if 1==len(arr){
					arrItem := arr[0]
					itemArr :=strings.Split(arrItem,",")
					for itemArrItemIdx,itemArrItem := range itemArr{
						temp,error := strconv.Atoi(itemArrItem)
						if error!=nil{//str
							tr+="\""
							tr+=itemArrItem
							tr+="\""
						}else{//int
							tr+=strconv.Itoa(temp)
						}
						if itemArrItemIdx!= len(itemArr)-1{
							tr+=","
						}
					}
				}else{
					for arrItemIdx,arrItem := range arr{
						td :="{"
						itemArr :=strings.Split(arrItem,",")
						for itemArrItemIdx,itemArrItem := range itemArr{
							temp,error := strconv.Atoi(itemArrItem)
							if error!=nil{//str
								td+="\""
								td+=itemArrItem
								td+="\""
							}else{//int
								td+=strconv.Itoa(temp)
							}
							if itemArrItemIdx!= len(itemArr)-1{
								td+=","
							}
						}
						td +="}"
						if arrItemIdx != len(arr)-1{
							td +=","
						}
						tr+=td
					}
				}
				tr += "}"
				itemStr += tr
			}else{
				itemStr += itemData
				if i==0{
					c0 = append(c0,itemData)
				}
				if i==1{
					c1 = append(c1,itemData)
				}
			}
			if loop != dlen-1{
				itemStr += ","
			}
		}
		itemStr+="}"
		r+=itemStr
	}
	head :="\tkeys={"
	headItemLen := len(c0)
	for i:=0;i<headItemLen;i++{
		head += "["+c1[i]+"]="+c0[i]
		if i != headItemLen-1{
			head+=","
		}
	}
	head += "},"

	final :=""
	final += "ATItem={\n"
	final+=head
	final+=r
	final+="\n}"
	ioutil.WriteFile(outputPath+outputFileName,[]byte(final),0666)
	fmt.Println("done")
}