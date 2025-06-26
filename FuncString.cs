using ExcelDna.Integration;
using System.Globalization;

// ==================== START of NAMESPACE ====================
// note that this namespace is split across multiple CS files
//
// https://excel-dna.net/docs/introduction
//
namespace AKHudfC
{

    public class FuncString : XlCall
    // ==================== START of Class ====================
    // "XlCall" allows direct calling of Excel's native functions (I think).
    // public static class FuncString   
    // static classes are non-instantiable. Static classes cannot be inherited from another class.
    // --------------------------------------------------------------------------------
    {
        // ==================== START of Function ====================
        // Description for IntelliSense Tool Tip
        [ExcelFunction(Description = "Reverses the order of letters in a single text string. CAT --> TAC")]
        public static string C_RevStr
            ([ExcelArgument(
                AllowReference = false,  // Don't use "AllowReference = true" for object arguments!!!
                Name="String",Description ="Text string that is to be reversed.")]
             string str
            )
        {
            if (str.Length > 0)
                return str[str.Length - 1] + C_RevStr(str.Substring(0, str.Length - 1));
            else
                return str;
        } // -------------------- End of Function --------------------

        // ==================== START of Function ====================
        // Description for IntelliSense Tool Tip
        [ExcelFunction(Description = "Changes the case of the letters in a single text string.")]
        public static string C_ChangeCase
            ([ExcelArgument(Name="String",Description ="Text string that is to have the case of letters changed.")]
              string ChangeStr,
             [ExcelArgument(Name ="Case",Description ="[optional] 0 = no case change, 1 = change to UPPERCASE, 2 = change to lowercase, 3 = Change To Title Case")]
              object CaseArg // optional argument
            )
        {
            // A call to the "Check" operation, defined in the "ToolsArgs" Class.
            // assign default value to optional argument, see "ToolsArgs" helper class for more info
            // Note that Excel doesn't store numbers as integers, so use double data type instead
            double bCase = ToolsArgs.CheckOpt(CaseArg, 0); // "0" is default value

            // Creates a TextInfo based on the "en-US" culture.
            TextInfo strTI = new CultureInfo("en-US", false).TextInfo;
            if (bCase == 0)
            {
                return ChangeStr;
            }
            else if (bCase == 1)   // "==" is the equivalence operator
            {
                return ChangeStr.ToUpper();
            }
            else if (bCase == 2)
            {
                return ChangeStr.ToLower();
            }
            else if (bCase == 3)
            {
                return strTI.ToTitleCase(ChangeStr);
            }
            else
            {
                return ChangeStr;
            }
        } // -------------------- End of Function --------------------


    } // ==================== END of Class ====================
} // ==================== END of Namespace ====================
