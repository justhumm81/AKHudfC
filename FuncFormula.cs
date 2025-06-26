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
    public class FuncFormula : XlCall
    // ==================== START of Class ====================
    // "XlCall" allows direct calling of Excel's native functions (I think).
    // public static class FuncString   
    // static classes are non-instantiable. Static classes cannot be inherited from another class.
    // ---------------------------------------------------------------
    {
        // --------------------------------------------------------------
        // Description for IntelliSense Tool Tip
        // --------------------------------------------------------------
        [ExcelFunction(IsMacroType = true, Description = "Retrieves the numeric formula of a referenced cell.")]
        public static string C_Formula([
            ExcelArgument(AllowReference = true, Name = "Cell", Description = "Cell containing formula.")]
            object objFormula)
        // ==================== START of Function ====================
        // Creates a string of the referenced cell's formula, including referenced values
        // ===========================================================
        {
            // variables that get used later
            string strFormula = "";
            string strFormText = "";
            string strVal = "";

            // use the native excel FORMULATEXT function
			// get textual formula that is in referenced cell/argument
            try
            {
                strFormText = (string)Excel(xlfFormulatext, objFormula);
            }
            // WHAT TO DO WHEN ARGUMENT IS NOT A FORMULA (it's a hardcoded value)
            catch (InvalidCastException)
            {
                try
                {
                    // try to extract the actual value
					object objVal = ToolsArgs.objRefVal(objFormula);
                    strVal = objVal.ToString();
                    // if it's empty, return 0
					return objVal is ExcelEmpty ? "0" : strVal;
                }
                catch
                {
                    return "Not a formula; value unavailable.";
                }
            }
            catch
            {
                return "Unexpected input error.";
            }

            // remove = and $ from the string formula
            strFormText = strFormText.Replace("$", "").Replace("=", "");

            // Create List of mathematical operators that may be encountered in formula
            List<string> strList = new List<string> { "(", ")", ",", "+", "-", "*", "/", ":", "&", "^", "<=", ">=", "<", ">" };

            // ------------------------------------------------------------
            // set initial counter values that will/may change
            // and proceed character by character through cell formula
            // ------------------------------------------------------------
            string strSheet = "";
            string strCell = "";
            string strI = "";
            char chI = 'i';

            // p1 will be first character in a substring
            // p2 will be last chacter in a substring
            int p1 = 0, p2 = 0;
            int i;

            // Main For Loop to try to assemble cell addresses in formula
            for (i = 0; i < strFormText.Length; i++)
            {
                chI = strFormText[i];
                strI = chI.ToString();

                int intExtSht = 0; // marker for later use

                // if character indicates reference to other worksheet
                if (strI == "!")
                {
                    int sheetEnd = i - 1;
                    strSheet = strFormText.Substring(p1, sheetEnd - p1 + 1);
                    if (strSheet.Contains(" "))
                        strSheet = "'" + strSheet + "'";
                    strSheet += "!";
                    p1 = i + 1;
                    continue;  // used to increment back to top of loop, instead of goto
                }

                // if math operator is encountered
                else if (strList.Contains(strI))
                {
                    // if this is the first character in the string
                    if (i == 0)
                    {
                        strFormula = strI;
                        p1 = i + 1;  // update pointer
                        continue; // skip & continue to top of loop
                    }

                    // else...if (i > 0)
                    p2 = i - 1;
                    strCell = strFormText.Substring(p1, p2 - p1 + 1);


                    if (strSheet.Contains("!"))
                    {
                        intExtSht = 1; //marker for external worksheet
                    }

                    strCell = strSheet + strCell;
                    strCell = ResolveCellValue(strCell, intExtSht);

                    intExtSht = 0; // reset externaal sheet marker to 0

                    // increment p1 for the next string (cell address)
                    // and insert the math operator into strFormula
                    p1 = i + 1;
                    strFormula += strCell + strI;
                    strSheet = "";
                    continue;
                }

                // if character is the last character in the formula
                if (i == strFormText.Length - 1)
                {
                    // case where refCell only contains hardcoded text/string
                    if (p2 == 0 && p1 == 0)
                    {
                        strCell = strFormText;
                    }
                    else
                    {
                        p2 = p2 + 2; // character just after last math operator was found
                        p1 = i;      // the (current) last character
                        strCell = strFormText.Substring(p2, p1 - p2 + 1);
                    }

                    if (strSheet.Contains("!"))
                    {
                        intExtSht = 1; //marker for external worksheet
                    }

                    strCell = strSheet + strCell;
                    strCell = ResolveCellValue(strCell, intExtSht);

                    intExtSht = 0; // reset externaal sheet marker to 0

                    strFormula += strCell;
                    strSheet = "";
                }
            } // end of for loop


            return strFormula;
        } // -------------------- End of Function -------------------- 

        // ==================== START OF HELPER METHOD ====================
        // Helper Method called by the main function
        // ===========================================================
        public static string ResolveCellValue(string strCellFull, int intExtSht)
        {
            object objCell = strCellFull;

            try
            {
                object refCellObj;

                if (intExtSht == 0)
                {
                    refCellObj = Excel(xlfIndirect, strCellFull);
                }
                else if (intExtSht == 1)
                {
                    refCellObj = Excel(xlfIndirect, $"\"\"{strCellFull}\"\"");
                }
                else
                {
                    refCellObj = null;
                }

                // DEBUGGING VARIABLES
                string strDebug = "";
                strDebug = $"\"{strCellFull}\"";
                strDebug = "";
                strDebug = $"[DEBUG] INDIRECT({strCellFull}) → {refCellObj?.GetType().Name}";


                // if refCellObj is just a string, not a valid cell reference, throw it to catch below
                if (!(refCellObj is ExcelReference refCell))
                {
                    throw new InvalidCastException();
                }
                // else if it is a valid cell reference, return the formatted contents
                //else if (refCellObj is refCell)
                //{
                    object val = refCell.GetValue();
                    string format = (string)Excel(xlfGetCell, 7, refCell);
                    return (string)Excel(xlfText, val, format);
                //}
                //else if (refCellObj is ExcelError)
                //{
                //    return "Invalid reference";
                //}
                //else
                //{
                //    return "Indirect did not process.";
                //}
            }
            catch (InvalidCastException)
            {
                if (objCell is ExcelEmpty) return "0";
                else if (ToolsArgs.ChkStringDouble(strCellFull)) return strCellFull;
                else return strCellFull;
                //else return "InvalidCast";
            }
            catch
            {
                return "Indirect exception";
            }
        }
        // -------------------- End of Helper Method -------------------- 

    } // ==================== END of Class ====================
} // ==================== END of Namespace ====================
