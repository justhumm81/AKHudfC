using ExcelDna.Integration;
using System.Reflection;

// ==================== START of NAMESPACE ====================
// note that this namespace is split across multiple CS files
//
// https://excel-dna.net/docs/introduction
//
namespace AKHudfC
{
    public class ToolsRange : XlCall
    // ==================== START of Class ====================
    // "XlCall" allows direct calling of Excel's native functions (I think).
    // class of routines and methods to help facilitate referencing
    // cells & ranges in Excel
    // https://groups.google.com/g/exceldna/c/zqzEIos7ma0/m/7XfV544o3Y8J
    // ---------------------------------------------------------
    {
        [ExcelFunction(IsMacroType = true)]
        public static object TestReferenceToRange()
        // ==================== START of Function ====================
        // ...
        // ---------------------------------------------------------
        {
            ExcelReference caller = (ExcelReference)XlCall.Excel(XlCall.xlfCaller);
            object callerRange = ReferenceToRange(caller);
            object callerAddress = callerRange.GetType().InvokeMember("Address",
                    BindingFlags.Public | BindingFlags.GetProperty,
                    null, callerRange, null);
            return callerAddress;
        } // -------------------- END of Method --------------------

        private static object ReferenceToRange(ExcelReference xlref)
        // ==================== START of Method ====================
        // ...
        // ---------------------------------------------------------
        {
            object app = ExcelDnaUtil.Application;
            object refText = Excel(xlfReftext, xlref, true);
            object range = app.GetType().InvokeMember("Range",
                    BindingFlags.Public | BindingFlags.GetProperty,
                    null, app, new object[] { refText });
            return range;
        } // -------------------- END of Method --------------------

        // --------------------------------------------------------------
        // Description for IntelliSense Tool Tip
        // --------------------------------------------------------------
        [ExcelFunction(IsMacroType = true)]
        public static object C_Indirect(
            [ExcelArgument(
            AllowReference = false,  // Don't use "AllowReference = true" for object arguments (sometimes)!!!
            Name = "Range String", Description = "Cell containing formula.")]
            string strRange
            )
        // ==================== START of Function ====================
        // simple call to Excel's native INDIRECT function to get values from other cells, based on text
        // ===========================================================
        {
            object objRngVal = Excel(xlfIndirect, strRange, true);

            return objRngVal;
        } // -------------------- END of Method --------------------

    }  // ========== END Class ==========
} // ========== END Namespace ==========