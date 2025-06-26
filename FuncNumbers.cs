using ExcelDna.Integration;
using System;
using System.Collections.Generic;

// ==================== START of NAMESPACE ====================
// note that this namespace is split across multiple CS files
//
// https://excel-dna.net/docs/introduction
//
namespace AKHudfC
{
    public class FuncNumbers : XlCall
    // ==================== START of Class ====================
    // Make sure to register this class with ": XlCall".
    {
        // ==================== START of Function ====================
        // Description for IntelliSense Tool Tipz
        [ExcelFunction(Description = "Find the Absolute Minimum (closest number to zero).")]
        public static double C_AbsMin
            ([ExcelArgument(
                AllowReference = false,  // Don't use "AllowReference = true" for object arguments!!!
                Name ="Range Arguments",Description ="cell(s) range of values to compare. (up to 9, so far)"
                )
             ]
            object Range1, object Range2, object Range3, object Range4, object Range5, object Range6, object Range7, object Range8, object Range9
            )
        {
            // Create List of function arguments and Loop through each
            List<object> ArgList = new List<object> { Range1, Range2, Range3, Range4, Range5, Range6, Range7, Range8, Range9 };

            // Go through arguments and break up ranges to individual elements
            ArgList = ToolsArray.objList(ArgList);

            // Make sure all arguments are doubles and put in list
            List<double> dblList = new List<double>();
            foreach (var i in ArgList)
            {
                // make sure all arguments are (or are turned into) doubles
                // see "ToolsArgs" helper class for more info
                dblList.Add(ToolsArgs.GetDbl(i));
            }
            // maximum (farthest from zero) possible double value for initial value
            double AbsMin = 1.7976931348623157E+308;

            foreach (double i in dblList)
            {
                if (i != 0 && Math.Abs(i) < Math.Abs(AbsMin))
                    AbsMin = i;
            }
            return AbsMin;
        } // -------------------- END of Function --------------------

        // ==================== START of Function ====================
        // Description for IntelliSense Tool Tipz
        [ExcelFunction(Description = "Find the Absolute Maximum (farthest number to zero).")]
        public static double C_AbsMax
            ([ExcelArgument(
                AllowReference = false,  // Don't use "AllowReference = true" for object arguments!!!
                Name ="Range Arguments",Description ="cell(s) range of values to compare. (up to 9, so far)"
                )
             ]
            object Range1, object Range2, object Range3, object Range4, object Range5, object Range6, object Range7, object Range8, object Range9
            )
        {
            // Create List of function arguments and Loop through each
            List<object> ArgList = new List<object> { Range1, Range2, Range3, Range4, Range5, Range6, Range7, Range8, Range9 };

            // Go through arguments and break up ranges to individual elements
            ArgList = ToolsArray.objList(ArgList);

            // Make sure all arguments are doubles and put in list
            List<double> dblList = new List<double>();
            foreach (var i in ArgList)
            {
                // make sure all arguments are (or are turned into) doubles
                // see "ToolsArgs" helper class for more info
                dblList.Add(ToolsArgs.GetDbl(i));
            }

            // smallest (closeest from zero) possible double value for initial value
            double AbsMax = 4.94065645841247E-324;

            foreach (double i in dblList)
            {
                if (i != 0 && Math.Abs(i) > Math.Abs(AbsMax))
                    AbsMax = i;
            }
            return AbsMax;
        } // -------------------- END of Function --------------------

        // ==================== START of Private Method ====================
        // Common routines for AbsMax and AbsMin functions
        // =================================================================
        //internal static 

    } // ========== END Class ==========
}   // ========== END Namespace ==========
