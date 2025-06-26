using AKHudfC;
using ExcelDna.Integration;
using System;
using System.Security.Policy;

public static class UnitConversion
{
    // ==================== START of Function ====================
    // Description for IntelliSense Tool Tip
    // -----------------------------------------------------------
    [ExcelFunction(Description = "Converts decimal feet to a mixed format, feet-inches, string.", Name = "C_FootToMix")]
    public static string C_FootToMix
            ([ExcelArgument(Name = "(Foot) Decimal Number", Description = "Decimal number to be converted to mixed format")]
             double number,
             [ExcelArgument(Name = "Fraction Rounding", Description = "[optional, 1/2\" default] fractional rounding of inches.")]
             object fracRnd,
             [ExcelArgument(Name = "Text Boolean", Description = "[optional, 0 (defult) for #'-#\", 1 for #ft-#in")]
             object txtBool
            )
    {
        if (!TryParseFootToMixArgs(number, fracRnd, txtBool, out double num, out double rnd, out bool txt, out string err))
            return err;

        // return ToMix(number, fracRnd, txtBool);
        string strMix = ToMix(num, rnd, txt);

        //string footPart, inchPart;
        //SplitFootInch(strMix, out footPart, out inchPart);

        //string mixFrac;
        //inchPart = DecimalToFraction(inchPart, out mixFrac);


        return strMix;

    } // -------------------- End of Function --------------------

    // ==================== START of Function ====================
    // Description for IntelliSense Tool Tip
    // -----------------------------------------------------------
    [ExcelFunction(Description = "Converts decimal inches to feet-inches format.", Name = "C_InchToMix")]
    public static string C_InchToMix
            ([ExcelArgument(Name = "(Inches) Decimal Number", Description = "Decimal number to be converted to mixed format")]
             double number,
             [ExcelArgument(Name = "Fraction Rounding", Description = "[optional, 1/2\" default] fractional rounding of inches.")]
             object fracRnd,
             [ExcelArgument(Name = "Text Boolean", Description = "[optional, 0 (defult) for #'-#\", 1 for #ft-#in")]
             object txtBool
            )

    {

        // convert inch decimal number to foot
        number = number * (1.0 / 12.0);

        // apply default values for missing arguments

        if (!TryParseFootToMixArgs(
                number, fracRnd, txtBool,
                out double num, out double rnd, out bool txt, out string err))
        {
            return err;
        }

        return ToMix(num, rnd, txt);
    }

    // ==================== START of Function ====================
    // Description for IntelliSense Tool Tip
    // -----------------------------------------------------------
    [ExcelFunction(Description = "Converts a feet-inches string back to decimal inches.", Name = "C_MixToInch")]
    public static double C_MixToInch
            ([ExcelArgument(Name = "Mixed Ft-Inch string", Description = "Mixed Ft-Inch string to be converted to decimal inches.")]
             string input
            )
    {
        double feet = 0;
        double inches = 0;

        input = input.Trim().ToLowerInvariant();

        // Normalize curly quotes to straight
        input = input.Replace("’", "'").Replace("‘", "'").Replace("“", "\"").Replace("”", "\"");

        // Remove final inch symbol for simpler parsing
        input = input.Replace("\"", "").Replace("inches", "").Replace("inch", "").Replace("in", "").Trim();

        // Fall back to original logic (e.g., "5' 10.5\"")
        int footMark = input.IndexOf('\'');
        if (footMark < 0) footMark = input.IndexOf("ft");
        if (footMark < 0) footMark = input.IndexOf("foot");
        if (footMark < 0) footMark = input.IndexOf("feet");

        if (footMark >= 0)
        {
            string footPart = input.Substring(0, footMark).Trim();
            double.TryParse(footPart, out feet);
        }

        int inchStart = (footMark >= 0) ? footMark + 1 : 0;
        string inchPart = input.Substring(inchStart).Trim();

        // Remove any dashes before the first digit
        int firstDigitIdx = inchPart.IndexOfAny("0123456789".ToCharArray());
        if (firstDigitIdx > 0)
        {
            string before = inchPart.Substring(0, firstDigitIdx).Replace("-", "");
            string after = inchPart.Substring(firstDigitIdx);
            inchPart = before + after;
            inchPart = inchPart.Trim();
        }

        // Case: whole inches and fraction separated by dash or space
        if (inchPart.Contains("-") || inchPart.Contains(" "))
        {
            char separator = inchPart.Contains("-") ? '-' : ' ';
            var split = inchPart.Split(new[] { separator }, 2);

            if (split.Length == 2 &&
                double.TryParse(split[0], out double whole) &&
                TryParseFraction(split[1], out double frac))
            {
                inches = whole + frac;
            }
        }

        // Case: only a fraction
        else if (inchPart.Contains("/"))
        {
            TryParseFraction(inchPart, out inches);
        }
        // Case: plain decimal number
        else
        {
            double.TryParse(inchPart, out inches);
        }

        inches = ParseInchesFromString(inchPart);

        return feet * 12 + inches;
    }

    // ==================== START HELPER METHOD ====================
    // assumes number fed into method is a foot 
    private static string ToMix(double number, double fracRnd, bool txtBool)
    {
        int feet = (int)Math.Floor(number);
        double inches = (number - feet) * 12;
        inches = Math.Round(inches / fracRnd) * fracRnd;

        // If rounding pushes inches to 12, bump feet
        if (inches >= 12.0)
        {
            feet += 1;
            inches = 0.0;
        }

        string strInch = FormatInchesAsFraction(inches, fracRnd);

        string result = $"{feet}'-{strInch}\"";

        return txtBool
            ? result.Replace("'-", " ft ").Replace("\"", " in")
            : result;
    } // -------------------- End of Helper Method --------------------

    // ==================== START HELPER METHOD ====================
    private static bool TryParseFootToMixArgs(object number, object fracRnd, object txtBool,
                                          out double num, out double rnd, out bool txt, out string error)
    {
        // default values
        error = null;
        num = 0;
        rnd = 0.5;
        txt = false;

        try
        {
            num = ToolsArgs.CheckOpt(number, 0.0); // no default — required
            rnd = ToolsArgs.CheckOpt(fracRnd, 0.5);
            txt = ToolsArgs.CheckOpt(txtBool, false);

            if (rnd <= 0 || rnd > 1)
            {
                error = "Round factor must be > 0 and <= 1.";
                return false;
            }

            return true;
        }
        catch (Exception ex)
        {
            error = "Invalid input.";
            return false;
        }
    } // -------------------- End of Helper Method --------------------

    // ==================== START HELPER METHOD ====================
    private static void SplitFootInch(string input, out string footPart, out string inchPart)
    {
        footPart = "";
        inchPart = "";

        // Find first character of foot marker
        int footMark = input.IndexOf('\'');
        if (footMark < 0) footMark = input.IndexOf("ft");
        if (footMark < 0) footMark = input.IndexOf("foot");
        if (footMark < 0) footMark = input.IndexOf("feet");

        if (footMark >= 0)
        {
            footPart = input.Substring(0, footMark).Trim();
            inchPart = input.Substring(footMark + 1).Trim();
        }

        // Remove any dashes before the first digit
        int firstDigitIdx = inchPart.IndexOfAny("0123456789".ToCharArray());
        if (firstDigitIdx > 0)
        {
            string before = inchPart.Substring(0, firstDigitIdx).Replace("-", "");
            string after = inchPart.Substring(firstDigitIdx);
            inchPart = before + after;
            inchPart = inchPart.Trim();
        }

    } // -------------------- End of Helper Method --------------------


    // ==================== START HELPER METHOD ====================
    private static string FormatInchesAsFraction(double inches, double fracRnd)
    {
        int whole = (int)Math.Floor(inches);
        double frac = inches - whole;

        int denom = (int)Math.Round(1.0 / fracRnd);
        int numer = (int)Math.Round(frac * denom);

        // Adjust for rounding overflow
        if (numer >= denom)
        {
            whole += 1;
            numer = 0;
        }

        // Simplify the fraction
        int gcd = GCD(numer, denom);
        numer /= gcd;
        denom /= gcd;

        if (numer == 0)
            return whole.ToString();
        else if (whole == 0)
            return $"{numer}/{denom}";
        else
            return $"{whole}-{numer}/{denom}";
    } // -------------------- End of Helper Method --------------------

    // ==================== START HELPER METHOD ====================
    private static int GCD(int a, int b)
    {
        while (b != 0)
        {
            int temp = b;
            b = a % b;
            a = temp;
        }
        return a == 0 ? 1 : a;
    } // -------------------- End of Helper Method --------------------


    // ==================== START HELPER METHOD ====================
    private static bool TryParseFraction(string fracStr, out double value)
    {
        value = 0;
        var parts = fracStr.Split('/');
        if (parts.Length == 2 &&
            double.TryParse(parts[0], out double num) &&
            double.TryParse(parts[1], out double denom) &&
            denom != 0)
        {
            value = num / denom;
            return true;
        }
        return false;
    } // -------------------- End of Helper Method --------------------


    // ==================== START HELPER METHOD ====================
    private static double ParseInchesFromString(string strInch)
    {
        double inches = 0;

        strInch = strInch.Trim();

        if (strInch.Contains(" "))
        {
            var split = strInch.Split(' ');
            if (split.Length == 2 &&
                double.TryParse(split[0], out double whole) &&
                TryParseFraction(split[1], out double frac))
            {
                inches = whole + frac;
            }
        }
        else if (strInch.Contains("-"))
        {
            var split = strInch.Split('-');
            if (split.Length == 2 &&
                double.TryParse(split[0], out double whole) &&
                TryParseFraction(split[1], out double frac))
            {
                inches = whole + frac;
            }
        }
        else if (strInch.Contains("/"))
        {
            TryParseFraction(strInch, out inches);
        }
        else
        {
            double.TryParse(strInch, out inches);
        }

        return inches;
    } // -------------------- End of Helper Method --------------------


}
