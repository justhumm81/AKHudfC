using ExcelDna.Integration;
using System;
//using System.Collections.Generic;
//using System.Security.Policy;
//using System.Reflection;
//using ExcelDna.IntelliSense;
// ==================== START of NAMESPACE ====================
// note that this namespace is split across multiple CS files
//
//
namespace AKHudfC
{
    public class ToolsArgs : XlCall
    // ==================== START of Class ====================
    // https://excel-dna.net/docs/guides-basic/optional-parameters-and-default-values/
    // Here is the helper class - add to it or change as you require
    // This method is a work-around for creating optional arguments (assiging default values when user does not specify an argument value).
    // "Check" appears to be an OVERLOADED method/function. https://www.w3schools.com/cs/cs_method_overloading.php
    {

        // --------------------------------------------------------------------------------
        // check to see if a string is / can be a double
        // --------------------------------------------------------------------------------
        public static bool ChkStringDouble(string strTest)
        {
            if (double.TryParse(strTest, out double d) && !Double.IsNaN(d) && !Double.IsInfinity(d))
            {
                return true;
            }
            return false;
        }

        // --------------------------------------------------------------------------------
        // Get the value of a reference and/or object
        // used when requesting fuction argument is maked as "ALLOWREFERENCE = TRUE"
        // --------------------------------------------------------------------------------
        public static object objRefVal(object objArg)
        {
            try
            {
                ExcelReference refVal = (ExcelReference)objArg;
                return refVal.GetValue();
            }
            catch (Exception ex)
            {
                int a = 0;
            }
            return 0;
        }

        // --------------------------------------------------------------------------------
        // OVERLOADED Methods/Functions
        // Methods/Functions to CHECK DATA TYPE of an argument and convert to DOUBLE where possible
        // For ExcelDNA, when the function is called, each OBJECT argument value will then be one of the possible types:
        // double, string,bool, ExcelMissing, ExcelEmpty, ExcelError.ExcelErrorXXXX, 
        // object[,] with a combination of the above
        // --------------------------------------------------------------------------------

        internal static double GetObjVal(object dblArg)
        {
            Type T1 = dblArg.GetType();
            if (dblArg is double)
                return (double)dblArg;
            else
                throw new ArgumentException();  // Will return #VALUE to Excel
        }

        internal static double GetDbl(object arg)
        //internal static double GetDbl(object arg, double Optional = 0)
        {
            Type T1 = arg.GetType();
            if (arg is double)
                return (double)arg;
            //else if (arg is object[,])
            //{
            //    //tell compiler that it's already an array
            //    object[,] objArr = (object[,])arg;

            //    //convert array data type from object to double
            //    double[,] dblArr = ToolsArray.ArrayObjToDbl(objArr);
            //    //double dblArrSum = ToolsArray.ArraySum(dblArr);
            //    //return dblArrSum;
            //    return dblArr;
            //}
            else if (arg is string)
                return (double)0;   // ignore it
            else if (arg is ExcelMissing || arg is ExcelEmpty)
                return (double)0;   // ignore it
            else
                throw new ArgumentException();  // Will return #VALUE to Excel
        }

        internal static string GetStr(object arg)
        //internal static string GetStr(object arg, string Optional = "0")
        {
            Type T1 = arg.GetType();
            if (arg is string)
                return (string)arg;
            else if (arg is double)
                return (string)arg.ToString();   // ignore it
            else if (arg is ExcelMissing || arg is ExcelEmpty)
                return (string)arg;   // ignore it
            else
                throw new ArgumentException();  // Will return #VALUE to Excel
        }

        // --------------------------------------------------------------------------------
        // OVERLOADED Methods/Functions
        // Return the data type of an OBJECT argument
        //
        // FOR NOW...use a DYNAMIC data type to make it work for differnt argument data types
        // --------------------------------------------------------------------------------
        internal static dynamic ObjGet(object arg)
        {
            Type T1 = arg.GetType();

            if (arg is string)
                return (string)arg;
            else if (arg is double)
                return (double)arg;
            else if (arg is ExcelMissing)
                throw new ArgumentException();  // Will return #VALUE to Excel
            else
                throw new ArgumentException();  // Will return #VALUE to Excel
        }


        // --------------------------------------------------------------------------------
        // OVERLOADED Methods/Functions
        // Check to see if an OPTIONAL FUNCTION ARGUMENT, if present, is a certain data type.
        // --------------------------------------------------------------------------------
        internal static string CheckOpt(object arg, string defaultValue)
        {
            if (arg is string)
                return (string)arg;
            else if (arg is ExcelMissing)
                return defaultValue;
            else if (arg is ExcelReference)
                //// Calling xlfRefText here requires IsMacroType=true for this function.
                return "Reference: " + XlCall.Excel(XlCall.xlfReftext, arg, true);
            else
                return arg.ToString();  // Or whatever you want to do here....
        }
        ////internal static double CheckOpt(object arg, double defaultValue)
        ////{
        ////    if (arg is double)
        ////        return (double)arg;
        ////    else if (arg is ExcelMissing)
        ////        return defaultValue;
        ////    else
        ////        throw new ArgumentException();  // Will return #VALUE to Excel
        ////}
        ///
        internal static double CheckOpt(object arg, double defaultValue)
        {
            if (arg is ExcelMissing || arg is ExcelEmpty)
                return defaultValue;
            else if (arg is double)
                return (double)arg;
            else if (arg is string s && double.TryParse(s, out double result))
                return result;
            else
                throw new ArgumentException("Invalid numeric argument.");
        }

        internal static bool CheckOpt(object arg, bool defaultValue)
        {
            if (arg is bool)
                return (bool)arg;
            else if (arg is double d)
                return d != 0; // Excel may pass TRUE as 1.0 and FALSE as 0.0
            else if (arg is string s)
                return s.Trim().ToLower() switch
                {
                    "true" or "yes" or "1" => true,
                    "false" or "no" or "0" => false,
                    _ => defaultValue
                };
            else if (arg is ExcelMissing)
                return defaultValue;
            else
                return defaultValue;
        }


        // check to see if optional argument, if present, is a date/time
        // This one is more tricky - we have to do the double->Date conversions ourselves
        internal static DateTime CheckOpt(object arg, DateTime defaultValue)
        {
            if (arg is double)
                return DateTime.FromOADate((double)arg);    // Here is the conversion
            else if (arg is string)
                return DateTime.Parse((string)arg);
            else if (arg is ExcelMissing)
                return defaultValue;
            else
                throw new ArgumentException();  // Or defaultValue or whatever
        }
        internal static string[] CheckOpt(object[] arg, string defaultValue)
        // check to see if optional argument, if present, is a 1d array (object[,])
        // The object array returned here may contain a mixture of types, reflecting the different cell contents.
        {
            if (arg is Array)
            {
                //get size (row X col) of array
                int iRows = arg.GetLength(0);

                //Create an empty array of the same size
                string[] ArgArr = new string[iRows];

                //Cast and fill each item
                // iterate through each element of 2d array
                for (int i = 0; i < iRows - 1; i++)
                {
                    ArgArr[i] = (string)((object[])arg)[i];
                }
                return ArgArr;
            }
            else if (arg is ExcelMissing)
            {
                string[] DefaultArr = new string[] { }; //empty array
                //DefaultArr[0] = defaultValue;
                return DefaultArr;
            }
            else
                throw new ArgumentException();  // Will return #VALUE to Excel
        }
        internal static double[,] CheckOpt(object arg)
        // check to see if optional argument, if present, is a 2d array (object[,])
        // The object array returned here may contain a mixture of types, reflecting the different cell contents.
        {
            if (arg is double[,])
                throw new ArgumentException();
            // return string.Format("Array[{0},{1}]({0},{1})",
            //        ((object[,](,)(,))arg).GetLength(0),
            //        ((object[,](,)(,))arg).GetLength(1));
            // else if (arg is ExcelMissing)
            //     return {0,0};  // return defaultValue;
            else
                throw new ArgumentException();  // Will return #VALUE to Excel
        }


        // ======================================================================
        // Perhaps check for other types and do whatever you think is right ....
        //else if (arg is double)
        //    return "Double: " + (double)arg;
        //else if (arg is bool)
        //    return "Boolean: " + (bool)arg;
        //else if (arg is ExcelError)
        //    return "ExcelError: " + arg.ToString();
        //else if (arg is object[,](,))
        //    // The object array returned here may contain a mixture of types,
        //    // reflecting the different cell contents.
        //    return string.Format("Array[{0},{1}]({0},{1})",
        //      ((object[,](,)(,))arg).GetLength(0), ((object[,](,)(,))arg).GetLength(1));
        //else if (arg is ExcelEmpty)
        //    return "<<Empty>>"; // Would have been null
        //else if (arg is ExcelReference)
        //  // Calling xlfRefText here requires IsMacroType=true for this function.
        //                return "Reference: " +
        //                     XlCall.Excel(XlCall.xlfReftext, arg, true);
        //            else
        //                return "!? Unheard Of ?!";
        // ======================================================================
    }
    // ==================== END of Class ====================

} // ========== END Namespace ==========
