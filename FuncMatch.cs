using ExcelDna.Integration;
//using static ExcelDna.Integration.XlCall; // You don’t have to have your class inherit from XlCall. That can be better done these days by adding <that>.
//using System.Reflection; // required for .GetValue Method
using System; // this is needed for TYPE definitions
using System.Collections.Generic; // need this for LIST operations
using System.Text.RegularExpressions;

// ==================== START of NAMESPACE ====================
// note that this namespace is split across multiple CS files
//
// https://excel-dna.net/docs/introduction
//
namespace AKHudfC
{

    public class FuncMatch : XlCall
    // ==================== START of Class ====================
    // public static class FuncMatch 
    // --------------------------------------------------------------------------------
    {
        // ==================== START of Function ====================
        // A MATCH function for multiple criteria
        // This is indended to / could be returned from Excel's native MATCH function
        // make these arguments DOUBLE, not INT
        //============================================================
        [ExcelFunction(IsMacroType = false, IsVolatile = false,
        Description = "Returns row no. of first (exact) MATCH of multiple (3) COLUMN criteria.")]
        public static object C_MMatch
            ([ExcelArgument(AllowReference = true, Name="Range0", Description ="Range to look for criteria match. Length must equal other ranges.")]
            object[] rng0,
            [ExcelArgument(AllowReference = false, Name="Criteria0", Description ="Criteria to look for match in range.")]
            object objCrit0,
            [ExcelArgument(AllowReference = true, Name="Range1", Description ="Range to look for criteria match. Length must equal other ranges.")]
            object[] rng1,
            [ExcelArgument(AllowReference = false, Name="Criteria1", Description ="Criteria to look for match in range.")]
            object objCrit1,
            //
            // consider these to be an OPTIONAL ARGUMENTS
            //
            [ExcelArgument(AllowReference = true, Name="[Range2]", Description ="Optional Range to look for criteria match.")]
            object[] rng2,
            [ExcelArgument(AllowReference = false, Name="[Criteria2]", Description ="Optional criteria to look for match in range.")]
            object objCrit2,
            [ExcelArgument(AllowReference = true, Name="[Range 3]", Description ="Optional range to look for criteria match.")]
            object[] rng3,
            [ExcelArgument(AllowReference = false, Name="[Criteria3]", Description ="Optional criteria to look for match in range.")]
            object objCrit3
            )
        {
            // ------------------------------------------------------------
            // Preliminary work to filter out unused, optional arguments
            // ------------------------------------------------------------

            // Check for OPTIONAL ARGUMENTS, set to NULL if not present
            rng2 = ToolsArgs.CheckOpt(rng2, "NULL");
            rng3 = ToolsArgs.CheckOpt(rng3, "NULL");
            objCrit2 = ToolsArgs.CheckOpt(objCrit2, "NULL");
            objCrit3 = ToolsArgs.CheckOpt(objCrit3, "NULL");

            int i = 0; int j = 0; int k = 0;
            int q = rng0.GetLength(0);
            int r = rng1.GetLength(0);
            int s = rng2.GetLength(0);
            int t = rng3.GetLength(0);
            int? z = null;   // set NULLABLE value for matching row

            //check range lengths, to make sure same size
            if (q != r)
            {
                return ToolsErrors.GetErrorNA();
            }

            // Create a List of criteria  
            // Create a List of ranges/arrays
            List<object> objListCrit = new List<object> { objCrit0, objCrit1, objCrit2, objCrit3 };
            List<object[]> objListRng = new List<object[]> { rng0, rng1, rng2, rng3 };

            // Remove Omitted Arguments from Lists
            // Assumes 3 (index 2) is the first optional argument in the list
            // go backwards through loop
            for (i = objListCrit.Count - 1; i >= 2; i--)
            {
                if (ToolsArgs.GetStr(objListCrit[i]) == "NULL")
                {
                    objListCrit.RemoveAt(i);
                    objListRng.RemoveAt(i);
                }
            }

            //check range lengths, to make sure same size
            int iArrs = objListCrit.Count - 1;
            for (i = 0; i < iArrs; i++)
            {
                j = objListRng[i].GetLength(0) - 1;
                k = objListRng[i + 1].GetLength(0) - 1;

                if (j != k)
                {
                    return ToolsErrors.GetErrorNA();
                }

                else continue;
            }

            //combine 1d range arrays into a single 2d array
            //builds a WIDE array...may need to be inverted later
            int iRows = rng0.GetLength(0);
            int iCols = objListCrit.Count;
            object[,] arrRng = new object[iRows, iCols];
            for (i = 0; i < iCols; i++)
            {
                for (j = 0; j < iRows; j++)
                {
                    arrRng[j, i] = objListRng[i][j];
                }

            }

            i = 0; j = 0; k = 0;
            int iLimit = arrRng.GetLength(0) - 1; // length of each loop (number of rows/columns)
            int jLimit = arrRng.GetLength(1) - 1; // number of criteria to loop through

            string strRngItem = null;
            string strCriteria = null;
            double? dblRngItem = null; // nullable double declaration
            double? dblCriteria = null;
            bool bMatch = false;

            // ------------------------------------------------------------
            // Start Comparisons
            // For Each Range in the List...
            // ------------------------------------------------------------
            // i is number of data rows/columns to check
            // j is each criteria in list
            // OUTER LOOP
            for (i = 0; i < iLimit; i++)
            {
            RestartOutLoop:
                // INNER LOOP
                for (j = 0; j < jLimit; j++)
                {
                RestartInLoop:
                    object objRngItem = arrRng[i, j];
                    var varRngItem = ToolsArgs.ObjGet(objRngItem); // var varRngItem = ToolsArgs.ObjGet(objRngItem, 0);
                    object objCriteria = objListCrit[j];
                    var varCriteria = ToolsArgs.ObjGet(objCriteria); // var varCriteria = ToolsArgs.ObjGet(objCriteria, 0);

                    // turn this IF into a TRY statement?
                    if (objRngItem is string && objCriteria is string)
                    {
                        // Replace Excel Wildcards with Regex Wildcards
                        strRngItem = (string)objRngItem;
                        strCriteria = (string)objCriteria;
                        // strRngItem = strRngItem.Replace("*", ".");
                        strCriteria = strCriteria.Replace("*", ".");

                        Match mMatch = Regex.Match(strRngItem, strCriteria, RegexOptions.IgnoreCase);

                        if (mMatch.Success)
                        {
                            bMatch = true;
                        }
                    }
                    else if (objRngItem is double && objCriteria is double)
                    {
                        //round values to X decimal places to eliminate rounding error
                        dblRngItem = Math.Round(ToolsArgs.GetDbl(objRngItem), 8);
                        dblCriteria = Math.Round(ToolsArgs.GetDbl(objCriteria), 8);

                        // used to be varRngItem
                        if (dblRngItem == dblCriteria)
                        {
                            bMatch = true;
                        }

                    }
                    else
                    {
                        return ToolsErrors.GetErrorNA();
                    }

                    // if row IS a match for this criteria
                    // and has NOT checked all criteria yet
                    if (j != jLimit && bMatch == true)
                    {
                        z = i;
                        j = j + 1;
                        bMatch = false; // reset value for next loop
                        goto RestartInLoop;
                    }
                    // if row IS a match for this criteria
                    // and HAS checked all other criteria
                    else if (j == jLimit && bMatch == true)
                    {
                        z = i;
                        z = z + 1; // +1 for last loop, because arrays are base 0, excel rows are base 1
                        bMatch = false; // reset value for next loop
                        return z;
                    }
                    // if row is NOT a match
                    else // if (bMatch == false)
                    {
                        // if row is not a match, reset return value to NULL
                        z = null;
                        i = i + 1;
                        bMatch = false; // reset value for next loop
                        goto RestartOutLoop; // then skip to the next row
                    }
                }
            }

            // if match is NEVER found
            if (z == null)
            {
                return ToolsErrors.GetErrorNA();
            }

            return z; // can we delete this line?

        } // -------------------- End of Function --------------------

    } // ==================== END of Class ====================
} // ==================== END of Namespace ====================
