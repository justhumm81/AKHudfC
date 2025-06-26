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

    public static class ToolsArray
    // ==================== START of Class ====================
    // static classes are non-instantiable. Static classes cannot be inherited from another class.
    // The intent of this class is to provide secondary procedures/subroutines that are repetative
    // and could get used in primary function definitions
    // --------------------------------------------------------------------------------
    {
        // ==================== START of Method ====================
        // Simple routine that will sum all the elements of a 2D array
        // ===========================================================
        internal static double ArraySum(double[,] arrValues)
        {
            int iRows = arrValues.GetLength(0);
            int iCols = arrValues.GetLength(1);

            double dblSum = 0;

            // iterate through each element of 2d array
            for (int i = 0; i < iRows; i++)
            {
                for (int j = 0; j < iCols; j++)
                {
                    dblSum = dblSum + arrValues[i, j];
                }
            }
            return dblSum;
        } // -------------------- END of Method --------------------

        // ==================== START of Method ====================
        // Create List of function arguments and Loop through each
        // Break up any range arguments and add the individual elements to the list
        // ===========================================================
        internal static List<object> objList(List<object> ArgList)
        {
            int i = 0; int j = 0; int k = 0;

            // --------------------------------------------------------------------------------
            // Declare a new, empty list, then populate it
            // --------------------------------------------------------------------------------
            List<object> objList = new List<object>();
            foreach (object arg in ArgList)
            //foreach (var i in ArgList)
            {
                Type T1 = arg.GetType();
                // --------------------------------------------------------------------------------
                // Insert arguments into List for processing, whether it's a value or an array/range
                // --------------------------------------------------------------------------------
                if (arg is object[,] && !(arg is ExcelMissing || arg is ExcelEmpty))
                {
                    object[,] objArr = (object[,])arg;
                    int iRows = objArr.GetLength(0);
                    int iCols = objArr.GetLength(1);
                    // --------------------------------------------------------------------------------
                    // iterate through each element of 2d array
                    // b/c .addrange doesn't seem to be working here
                    // --------------------------------------------------------------------------------
                    for (i = 0; i < iRows; i++)
                    {
                        for (j = 0; j < iCols; j++)
                        {
                            objList.Add(objArr[i, j]);
                        }
                    }
                    //objList.AddRange(objArr);
                }
                else if (arg is object && !(arg is ExcelMissing || arg is ExcelEmpty))
                    objList.Add(arg);
                else if (arg is ExcelMissing || arg is ExcelEmpty)
                    j = 666;
                else
                    objList.Add(800851);   // BOOBS! Error
            }
            return objList;
        } // -------------------- END of Method --------------------

        // ==================== START of Method ====================
        // Converts 2D object[,] to a double[,] array...I hope
        // ===========================================================
        internal static double[,] ArrayObjToDbl(object[,] objArr)
        {
            //get size (row X col) of array
            int iRows = objArr.GetLength(0);
            int iCols = objArr.GetLength(1);

            //Create an array of doubles, which is the same size
            double[,] dblArr = new double[iRows, iCols];

            //Cast and fill each item
            // iterate through each element of 2d array
            for (int i = 0; i < iRows; i++)
            {
                for (int j = 0; j < iCols; j++)
                {
                    dblArr[i, j] = (double)((object[,])objArr)[i, j];
                }
            }
            return dblArr;
        } // -------------------- END of Method --------------------

        // ==================== START of Private Subroutine ====================
        // Note: Array indexes start with 0: [0] is the first element. [1] is the second element, etc.
        //
        private static double BuildArray(double arrArgs, double rngArgs)
        {
            int k = 0;

            return arrArgs;
        }

        //////        Private Sub BuildArray(arrArg() As Variant, rngArgs() As Variant)

        //////    Dim arrSubArg() As Variant
        //////    m = UBound(rngArgs())
        //////    For i = 0 To UBound(rngArgs())

        //////        'POPULATE  SUB-ARRAY FROM DEFINED RANGE ARGUMENT
        //////        If rngArgs(i) Is Nothing Then
        //////            ReDim arrSubArg(0 To 0)
        //////            arrSubArg(0) = arrArg(0)
        //////        Else
        //////            k = rngArgs(i).Cells.Count
        //////            ReDim arrSubArg(0 To (k - 1))
        //////            For j = 0 To(k - 1)
        //////                arrSubArg(j) = rngArgs(i)(j + 1).Value
        //////            Next j
        //////        End If
        //////        'APPEND SUB-ARRAY TO ANY PREVIOUS RANGE ARGUMENTS
        //////        j = UBound(arrSubArg)
        //////        k = UBound(arrArg)
        //////        ReDim Preserve arrArg(0 To (k + UBound(arrSubArg) + 1))
        //////        For j = 0 To UBound(arrSubArg)
        //////            arrArg(k + j + 1) = arrSubArg(j)
        //////        Next j
        //////    Next i
        //////'''    'MAKE ALL VALUES IN ARRAY ARE NUMERIC(NOT TEXT)
        //////'''    For i = 1 To UBound(arrArg)
        //////'''        If Not IsNumeric(arrArg(i)) Then
        //////'''            arrArg(i) = 0
        //////'''        End If
        //////'''    Next i
        //////End Sub


    } // ========== END Class ==========
} // ========== END Namespace ==========
