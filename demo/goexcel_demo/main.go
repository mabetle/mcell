package main

import(
	"github.com/mabetle/mcell"
	"fmt"
)

func DemoReadCell(){
	f:="../data/demo.xls"
	fmt.Println(mcell.NewWorkbook(f).GetSheet(1).GetCell(1,1).String())
}

func PrintSheet(){
	f:="../data/demo.xls"
	sheet:=mcell.NewWorkbook(f).GetSheet(1)
	// walk sheet
	//for 
	fmt.Println(sheet)
}

func main(){
	DemoReadCell()
	PrintSheet()
}
