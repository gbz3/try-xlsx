package main

import (
  "flag"
  "fmt"
  "io/ioutil"
  "os"
  "strings"

  "github.com/tealeg/xlsx"
)

var fKW string
var fAll bool
var fRow int
var fCol int
var fRowNum int
var fColNum int

func init() {
  flag.StringVar( &fKW, "keyword", "", "search keyword" )
  flag.BoolVar( &fAll, "all", false, "all" )
  flag.IntVar( &fRow, "row", -1, "row index" )
  flag.IntVar( &fCol, "col", -1, "col index" )
  flag.IntVar( &fRowNum, "rownum", 1, "num of row" )
  flag.IntVar( &fColNum, "colnum", 1, "num of col" )
}

func inDump() bool { return len( fKW ) == 0 }
func inSearchOne() bool { return len( fKW ) > 0 && fAll == false }
func inPointed() bool { return fRow >= 0 && fCol >= 0 }
func inRect() bool { return inPointed() && ( fRowNum > 1 || fColNum > 1 ) }

func main() {
  flag.Parse()

  // Excel ファイルを読み込む
  b, err := ioutil.ReadAll( os.Stdin )
  if err != nil {
    panic( err )
  }

  book, err := xlsx.OpenBinary( b )
  if err != nil {
    panic( err )
  }
  fmt.Printf( "file opened.\n" )

  for i, sheet := range book.Sheets {
    fmt.Printf( "[%03d] sheet[%s]: \n", i, sheet.Name )
    if inRect() {
      for j := 0; j < fRowNum; j++ {
        for k := 0; k < fColNum; k++ {
          if k > 0 { fmt.Print( "\t" ) }
          fmt.Print( sheet.Cell( fRow+j, fCol+k ).String() )
        }
        fmt.Print( "\n" )
      }
      os.Exit( 0 )
    } else if inPointed() {
      cell := sheet.Cell( fRow, fCol )
      fmt.Printf( "\trow=%d cell=%d Text=[%s]\n", fRow, fCol, cell.String() )
      os.Exit( 0 )
    }
    for j, row := range sheet.Rows {
      if inDump() {
        fmt.Printf( "%03d: ", j )
      }
      evaluateCells( j, row.Cells )
      if inDump() {
        fmt.Print( "\n" )
      }
    }
  }

}

func evaluateCells( numOfRow int, cells []*xlsx.Cell ) {
  for i, cell := range cells {
    if len( fKW ) > 0 {
      if strings.Contains( cell.String(), fKW ) {
        fmt.Printf( "\trow=%d cell=%d Text=[%s]\n", numOfRow, i, cell.String() )
        if inSearchOne() {
          os.Exit( 0 )
        }
      }
    } else {
      if i > 0 {
        fmt.Print( "\t" )
      }
      fmt.Print( cell.String() )
    }
  }
}

