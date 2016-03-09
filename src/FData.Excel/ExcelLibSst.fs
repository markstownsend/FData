module FData.Excel.Sst

open System.Xml
open FSharp.Data
open System.Linq
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Spreadsheet

let getAllStrings (workbook:string) = 
    use stream = new System.IO.FileStream(workbook, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite)

    let document = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(stream, false)
    
    let wbPart = document.WorkbookPart
    let sstPart = wbPart.GetPartsOfType<SharedStringTablePart>().First()
    let sst = sstPart.SharedStringTable
    
    let results = 
        sst.Elements<SharedStringItem>().ToArray()
            |> Array.map(fun f -> workbook, f.Text.Text)

    document.Close()

    results
    