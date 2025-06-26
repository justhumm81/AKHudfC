using ExcelDna.Integration;

// ==================== START of NAMESPACE ====================
// note that this namespace is split across multiple CS files
//
// https://excel-dna.net/docs/introduction
//
namespace AKHudfC
{
    public class FuncTesting : XlCall
    // ==================== START of Class ====================
    // Make sure to register this class with ": XlCall".
    {
        //// ==================== START of Function ===================
        //[ExcelFunction(Description = "A useful test function that adds two numbers, and returns the sum.")]
        //public static double C_AddThem
        //    ([ExcelArgument(Name ="Add1",Description ="is the first number, to which will be added")]
        //    double v1,
        //    [ExcelArgument(Name ="Add2", Description ="is the second number that will be added")]
        //    double v2,
        //    [ExcelArgument(Name ="Add3", Description ="is the third number that will be added")]
        //    object v3
        //    )
        //{
        //    double dblArg = ToolsArgs.CheckOpt(v3, 100); // "100" is default value
        //    return v1 + v2 + dblArg;
        //} // -------------------- END of Function --------------------

        // ==================== START of Function ====================
        // Description for IntelliSense Tool Tip
        [ExcelFunction(IsMacroType = false, IsVolatile = false,
                        Description = "A speed test function that counts from 0 to the BigNumber, using Step as increment.")]
        public static string C_BiggerNumber
            ([ExcelArgument(AllowReference = false, Name="BigNumber",Description ="The number that will be counted to.")]
              double NumbDbl,
             [ExcelArgument(AllowReference = false, Name ="Step",Description ="[optional] step")]
              object StepArg // optional int argument
            )

        {
            {   // assign default value to optional argument, see "ToolsArgs" helper class for more info
                double Step = ToolsArgs.CheckOpt(StepArg, 1); // "1" is default value

                double counter = 0;

                var watch = new System.Diagnostics.Stopwatch();

                watch.Start();

                while (counter <= NumbDbl)
                {
                    counter = counter + Step;
                }

                watch.Stop();

                long lngWatch = watch.ElapsedMilliseconds;

                return "This took " + lngWatch + " milliseconds.";
            }
        } // -------------------- END of Function --------------------

        //// ==================== START of Function ====================
        //// https://groups.google.com/g/exceldna/c/kf76nqAqDUo
        //// As of June 2023, this function works for both multiple absolute arguments =MySum(1,2,3)
        //// and it works for multiple individual range arguments =MySum(A1,A2,A3)
        //// but it does NOT work for a large range argument \=MySum(A1:A3)
        //// Description for IntelliSense Tool Tip
        //[ExcelFunction(Description = "Test function...first test for params (varying number of arguments in C#). Adds together all arguments.")]
        //public static double C_MySum(params double[] values)
        //{

        //    return values.Sum();
        //}
        //// -------------------- END of Function --------------------

        //// ==================== START of Function ====================
        //// Description for IntelliSense Tool Tip
        //[ExcelFunction(Description = "Test function...test to Add together all arguments in a single 2D range.")]
        //public static double C_RangeSum
        //([ExcelArgument(AllowReference = true)] double[,] arrValues)      // Don't use "AllowReference = true" for object arguments
        //{
        //    // call to function/subroutine in separate class
        //    return ToolsArray.ArraySum(arrValues);
        //}
        //// -------------------- END of Function --------------------

        //// ==================== START of Function ====================
        //// Description for IntelliSense Tool Tip
        //[ExcelFunction(Description = "Test function...test to Add together all arguments in defined ranges.")]
        //public static double C_SumRanges
        //([ExcelArgument(AllowReference = false)]   // Don't use "AllowReference = true" for object arguments
        // object Range1, object Range2, object Range3, object Range4, object Range5
        //) 
        //{
        //    double dblSum = 0;
        //    // Create List of function arguments and Loop through each
        //    List<object> ArgList = new List<object> {Range1,Range2,Range3,Range4,Range5};
        //    foreach (var i in ArgList)
        //    {
        //        // see "ToolsArgs" helper class for more info
        //        double dblArg = ToolsArgs.GetDbl(i);
        //        dblSum = dblSum + dblArg;
        //    }
        //    return dblSum;
        //} // -------------------- END of Function --------------------

    } // -------------------- END of CLASS --------------------
} // -------------------- END of NAMESPACE--------------------
