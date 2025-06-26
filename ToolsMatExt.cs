
// ==================== START of NAMESPACE ====================
// note that this namespace is split across multiple CS files

//
namespace AKHudfC
{

    public static class ToolsMatExt
    // ==================== START of Class ====================
    // Programming tools that use Matrix Extensions
    // The MatrixExtensions type provides methods to transform a Matrix (Rotate, Scale, Translate, etc...).
    // These are a similar subset of methods originally provided in the System.Windows.Media.Matrix class.
    // https://learn.microsoft.com/en-us/windows/communitytoolkit/extensions/matrixextensions
    // https://learn.microsoft.com/en-us/dotnet/csharp/programming-guide/classes-and-structs/extension-methods
    //
    // https://stackoverflow.com/questions/16636019/how-to-get-1d-column-array-and-1d-row-array-from-2d-array-c-net
    // Usage example:
    //		double[,] myMatrix = ... // Initialize with desired size and values.
    //		double[] myRowVector = myMatrix.GetRow(2); // Gets the third row.
    //		double[] myColVector = myMatrix.GetCol(1); // Gets the second column.
    //		myMatrix.SetCol(2, myColVector); // Sets the third column to the second column.
    // --------------------------------------------------------------------------------
    {
        // ==================== START of Method ====================
        // Returns the row with number 'row' of this matrix as a 1D-Array.
        // ===========================================================
        public static T[] GetRow<T>(this T[,] matrix, int row)
        {
            var rowLength = matrix.GetLength(1);
            var rowVector = new T[rowLength];

            for (var i = 0; i < rowLength; i++)
                rowVector[i] = matrix[row, i];

            return rowVector;
        }
        // -------------------- END of Method --------------------

        // ==================== START of Method ====================
        // Sets the row with number 'row' of this 2D-matrix to the parameter 'rowVector'.
        // ===========================================================
        public static void SetRow<T>(this T[,] matrix, int row, T[] rowVector)
        {
            var rowLength = matrix.GetLength(1);

            for (var i = 0; i < rowLength; i++)
                matrix[row, i] = rowVector[i];
        }
        // -------------------- END of Method --------------------

        // ==================== START of Method ====================
        // Returns the column with number 'col' of this matrix as a 1D-Array.
        // ===========================================================
        public static T[] GetCol<T>(this T[,] matrix, int col)
        {
            var colLength = matrix.GetLength(0);
            var colVector = new T[colLength];

            for (var i = 0; i < colLength; i++)
                colVector[i] = matrix[i, col];

            return colVector;
        }
        // -------------------- END of Method --------------------

        // ==================== START of Method ====================
        // Sets the column with number 'col' of this 2D-matrix to the parameter 'colVector'.
        // ===========================================================
        public static void SetCol<T>(this T[,] matrix, int col, T[] colVector)
        {
            var colLength = matrix.GetLength(0);

            for (var i = 0; i < colLength; i++)
                matrix[i, col] = colVector[i];
        }
        // -------------------- END of Method --------------------  

    } // ========== END Class ==========
} // ========== END Namespace ==========
