MS Excel add-in of UDFs written in C# / ExcelDNA, primarily for experimentation and assistance in preparing civil engineering calculations.

UDFs currently included in the add-in:
- C_AbsMax (extracts the absolute maximum from a range or list of values)
- C_AbsMin (extracts the absolute minimum from a range or list of values)
- C_BiggerNumber (just a speed test...counts the miliseconds to reach a large number via specified increments)
- C_ChangeCase (changes the case of a text string)
- C_FootToMix (converts a decimal number - assumed FEET - to a Foot-Inch string)
- C_Formula (similar to the native FORMULATEXT, but converts the cell addresses to the cell contents)
- C_InchToMix (converts a decimal number - assumed INCHES - to a Foot-Inch string)
- C_Indirect (not working properly yet)
- C_Linterp (performs a linerar interpolation based on two know X-Y values)
- C_MixToInch (converts a Foot-Inch string to a decimal number of inches)
- C_MMatch (not working properly yet)
- C_RevStr (reverses the order of characters in a string, CAT --> TAC)

![image](https://github.com/user-attachments/assets/ce95a1ee-47fa-4869-adac-f265313e354f)


Initial inspiration for this endeavor was from Morefunc (by Laurent Longre), which was written in C++ and basically went defunct (2007) due to incompatability with newer versions of Excel. This repository is NOT meant to act as a replacement for it.
https://web.archive.org/web/20060601112922/http://xcell05.free.fr/
