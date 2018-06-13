package main

import (
	"strconv"
    "os"
    "fmt"
    "github.com/tealeg/xlsx"
    "path/filepath"
)

func main() {
    fmt.Printf("植物医生表格转换程序 2018.06.13.001 by wxl\n\n请将待处理表格放至程序同级目录下并重名名为\"input.xlsx\".\n确认文件存在后按回车开始处理.\n")
    fmt.Scanln()
    dir, err := filepath.Abs(filepath.Dir(os.Args[0]))
    if err != nil {
        fmt.Println(err)
    }
    fmt.Printf("%s\n",dir)
    excelFileName := dir +"\\input.xlsx"
    saveFileName := dir +"\\output.xlsx"
    // excelFileName := "input.xlsx"
    // saveFileName := "output.xlsx"
    xlFile, err := xlsx.OpenFile(excelFileName)
    if err != nil {
        fmt.Printf("文件打开失败: %s\n", err)
        fmt.Scanln()

    }

    newfile := xlsx.NewFile()
    newsheet, newerr := newfile.AddSheet("Sheet1")
    if newerr != nil {
        fmt.Printf(newerr.Error())
    }
    
    sheet := xlFile.Sheets[0]
    index := 0
    for _, row := range sheet.Rows {
        if index == 0 {//字段名复制
            newrow := newsheet.AddRow()
            newrow.SetHeightCM(1)
            for b := 0; b < 23; b++ {//复制原来的 cell 值
                newcell := newrow.AddCell()
                newcell.Value = row.Cells[b].String()
            }
            //多加两个字段
            newcell1 := newrow.AddCell()
            newcell1.Value = "订单类型"
            newcell2 := newrow.AddCell()
            newcell2.Value = "订单金额"
        }else{
            fmt.Printf("正在处理第%d条数据......\n" ,index)
            count,err := strconv.Atoi(row.Cells[12].String()) 
            if err != nil {
                fmt.Println(err)
                fmt.Scanln()
            }
            for a := 0; a < count; a++ {//每个订单新建一个包裹
                houzhui := fmt.Sprintf("-%03d",a+1)
                orderno := row.Cells[3].String() + houzhui
                newrow := newsheet.AddRow()
                newrow.SetHeightCM(1)
                for b := 0; b < 23; b++ {//复制原来的 cell 值
                    newcell := newrow.AddCell()
                    newcell.Value = row.Cells[b].String()
                }
                newcell1 := newrow.AddCell()
                newcell1.Value = "配送"
                newcell2 := newrow.AddCell()
                newcell2.Value = "0"
                newrow.Cells[3].Value = orderno //更改为加后缀的单号
            }

        }
        index++
    }
    err = newfile.Save(saveFileName)
    if err != nil {
        fmt.Printf(err.Error())
        fmt.Scanln()
    }
    fmt.Printf("处理完成,共%d条数据.\n" ,index-1)
    fmt.Printf("处理后文件为\"output.xlsx\".\n")
    fmt.Printf("按回车键退出.\n")
    fmt.Scanln()
}