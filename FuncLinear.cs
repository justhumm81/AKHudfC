using ExcelDna.Integration;

public static class InterpolationFunctions
{
    [ExcelFunction(Description = "Linear interpolation between (x1, y1) and (x2, y2) for a given x.")]
    public static object C_Linterp(object x, object x1, object x2, object y1, object y2)
    {
        // Try to parse inputs as doubles
        if (!TryGetDouble(x, out double dx) ||
            !TryGetDouble(x1, out double dx1) ||
            !TryGetDouble(x2, out double dx2) ||
            !TryGetDouble(y1, out double dy1) ||
            !TryGetDouble(y2, out double dy2))
        {
            return ExcelError.ExcelErrorValue;
        }

        // Handle boundary conditions
        if (dx == dx1) return dy1;
        if (dx == dx2) return dy2;

        // Perform interpolation
        return dy1 + (dx - dx1) * (dy2 - dy1) / (dx2 - dx1);
    }

    // Helper method to safely convert Excel inputs to double
    private static bool TryGetDouble(object input, out double result)
    {
        if (input is double d)
        {
            result = d;
            return true;
        }

        if (input is string s && double.TryParse(s, out double parsed))
        {
            result = parsed;
            return true;
        }

        result = double.NaN;
        return false;
    }
}
