module FData.Excel.Formula

open System.Xml
open FSharp.Data
open System.Linq
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Spreadsheet
open FSharp.Collections.ParallelSeq

/// <summary>
/// Gets all the cells with formulas out of the worksheet
/// </summary>
/// <param name="sheet">the worksheet to look through</param>
let getFormulaCells (sheet:WorksheetPart) =  
    let cells = sheet.Worksheet.Descendants<Cell>().ToArray()
    cells |> Array.where(fun c -> c.InnerXml.Contains("<x:f"))
    
/// <summary>
/// Gets all the cells with formulas out of the worksheet.
/// Runs asynchronously so can run in parallel across all the sheets in the workbook
/// </summary>
/// <param name="sheet">the worksheet to look through</param>
//let getFormulaCellsAsync (sheet:WorksheetPart) = async {
//        let cells = 
//            |> PSeq.collect sheet.Worksheet.Descendants<Cell>())
//        return cells |> Array.where(fun c -> c.InnerXml.Contains("<x:f"))
//    }

/// <summary>
/// Gets all the formula cells out of the workbook.
/// Runs in parallel against the cells collection on each worksheet in order to speed up
/// </summary>
/// <param name="workbook">the full path to the workbook to examine</param>
let getAllFormulasP (workbook:string) = 
    use stream = new System.IO.FileStream(workbook, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite)
    let document = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(stream, false)
    let wbPart = document.WorkbookPart
    

    // this just seems to get a collection of worksheets and doesn't iterate the cells
    // collection
    let results = 
        wbPart.WorksheetParts.
            |> PSeq..map (fun sht -> sht.Worksheet.Descendants<Cell>().AsEnumerable())
//                // get formula cells
            |> PSeq.filter (fun cl -> cl.InnerXml.Contains("<x:f"))
//                // get master formula cells only
//            |> PSeq.filter (fun cl -> cl.CellFormula.FormulaType.Value <> CellFormulaValues.Shared)
//                // return something flatter than a cell
            |> PSeq.map (fun cl -> "hello world")
            |> PSeq.toArray
    document.Close()
    results
//
//let hasFormula (worksheetCell:cell) = 
//    try


let getAllFormulas (workbook:string) = 
    use stream = new System.IO.FileStream(workbook, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite)

    let document = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(stream, false)
    
    let wbPart = document.WorkbookPart

    let sheets = wbPart.WorksheetParts


    let sheet = sheets.First<WorksheetPart>();
    let results = 
        getFormulaCells sheet
            |> Array.map(fun f -> sheet.Worksheet.SheetProperties.CodeName.Value, f.CellFormula)

    document.Close()

    results

   