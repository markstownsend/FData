#r @"..\..\packages\DocumentFormat.OpenXml\lib\DocumentFormat.OpenXml.dll"
#r @"C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.5.2\WindowsBase.dll"
#r @"..\..\packages\FSharp.Collections.ParallelSeq\lib\net40\FSharp.Collections.ParallelSeq.dll"


#load "ExcelLibSst.fs"
//#load "ExcelLibFormula.fs"

open FData.Excel.Sst
//open FData.Excel.Formula
open System.IO
open FSharp.Data
open FSharp.Collections

let files = Directory.EnumerateFiles(@"C:\mark\excel\deutsche", @"B*.xls*", SearchOption.AllDirectories)

let results = 
    files 
        |> Seq.map (fun f -> getAllStrings f)

results
    |> Seq.iter (fun tp -> tp |> Array.iter (fun s -> printfn "%s, %s" (fst s) (snd s)))
            

