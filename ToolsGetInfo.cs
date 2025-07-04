using ExcelDna.Integration;

// https://github.com/Excel-DNA/Samples/blob/master/GetInfoFunctions/GetInfoAddIn.cs
// https://groups.google.com/g/exceldna/c/InANJKfh5_8/m/8UtGE3r2AwAJ

public class GetInfoFunctions
{
    [ExcelFunction(Description = "Returns the result of xlfGetCell.", IsMacroType = true)]
    public static object GetCell(int type_num, [ExcelArgument(AllowReference = true)] object reference)
    {
        return XlCall.Excel(XlCall.xlfGetCell, type_num, reference);
    }

    [ExcelFunction(Description = "Returns the result of xlfGetDocument.", IsMacroType = true)]
    public static object GetDocument(int type_num, string name_text)
    {
        return XlCall.Excel(XlCall.xlfGetDocument, type_num, name_text);
    }

    [ExcelFunction(Description = "Returns the result of xlfGetWorkbook.", IsMacroType = true)]
    public static object GetWorkbook(int type_num, string name_text)
    {
        return XlCall.Excel(XlCall.xlfGetWorkbook, type_num, name_text);
    }

    [ExcelFunction(Description = "Returns the result of xlfGetWorkbook.", IsMacroType = false)]
    public static object GetWorkbookActive(int type_num)
    {
        return XlCall.Excel(XlCall.xlfGetWorkbook, type_num);
    }

    [ExcelFunction(Description = "Returns the result of xlfGetWorkspace.", IsMacroType = true)]
    public static object GetWorkspace(int type_num)
    {
        return XlCall.Excel(XlCall.xlfGetWorkspace, type_num);
    }

    [ExcelFunction(Description = "Returns the current list separator.", IsMacroType = true)]
    public static string GetListSeparator(int type_num)
    {
        object[,] workspaceSettings = (object[,])XlCall.Excel(XlCall.xlfGetWorkspace, 37);
        string listSeparator = (string)workspaceSettings[0, 4];
        return listSeparator;
    }
}