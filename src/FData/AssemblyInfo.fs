namespace System
open System.Reflection

[<assembly: AssemblyTitleAttribute("FData")>]
[<assembly: AssemblyProductAttribute("FData")>]
[<assembly: AssemblyDescriptionAttribute("Mucking around with FSLab")>]
[<assembly: AssemblyVersionAttribute("1.0")>]
[<assembly: AssemblyFileVersionAttribute("1.0")>]
do ()

module internal AssemblyVersionInformation =
    let [<Literal>] Version = "1.0"
