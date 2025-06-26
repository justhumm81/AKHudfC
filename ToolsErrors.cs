using ExcelDna.Integration;

// ==================== START of NAMESPACE ====================
// note that this namespace is split across multiple CS files

namespace AKHudfC
{
    public class ToolsErrors : XlCall
    // ==================== START of Class ====================
    // class of routines and methods to help facilitate referencing
    // cells & ranges in Excel
    // https://groups.google.com/g/exceldna/c/zqzEIos7ma0/m/7XfV544o3Y8J
    // ---------------------------------------------------------
    {

        public static object GetErrorValue()
        // ==================== START of Method ====================
        // ...
        // =========================================================
        {
            return ExcelError.ExcelErrorValue;
        }
        // -------------------- END of Method ---------------------

        public static object GetErrorRef()
        // ==================== START of Method ====================
        // when this method is called, it will return the "#N/A" error to Excel
        // =========================================================
        {
            return ExcelError.ExcelErrorRef;
        }

        public static object GetErrorNA()
        {
            return ExcelError.ExcelErrorNA;
        }


    }  // ========== END Class ==========
} // ========== END Namespace ==========