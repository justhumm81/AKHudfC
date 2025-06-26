
// https://excel-dna.net/docs/introduction


// EXCELFUNCTIONATTRIBUTE
// https://excel-dna.net/docs/archive/wiki/ExcelFunction-and-other-attributes/


// Accepting RANGE PARAMETERS in UDFs
// https://excel-dna.net/docs/guides-basic/accepting-range-parameters-in-udfs/


// FUNCTION PARAMETERS and RETURN Data Types
//The allowed function parameter and return types are
//incoming function parameters of type Object will only arrive as one of the following
//return values of type Object are allowed to be
//https://excel-dna.net/reference-data-type-marshalling

//CLASSES
// ": XlCall" allows direct calling of Excel's native functions (I think).
// The class must be public for the static functions to be exposed to Excel.
// static classes are non-instantiable. Static classes cannot be inherited from another class

EXCELDNA ADDIN RELEASE NOTES

https://groups.google.com/g/exceldna/c/GfTCDsiwdc8/m/3uqoacrlAgAJ

What exactly you need to distribute depends a bit on which Excel-DNA package you're using, and how complicated your add-in is.
For a simple add-in it should work with the three files (your .dll and the .xll and .dna file for the right bitness of your Excel).
If you have other dependencies, you might see those .dll files in the output too, and you also need them to have a workign add-in.

To make distribution easier, as part of the build process, we make XXX-packed.xll and XXX-packed64.xll add-ins, which (if things are configured right) can be a single file distribution of your add-in.
(Well, two single-files, for 32-bit and 64-bit Excel.)
In the latest ExcelDna.AddIn package versions (1.7.0-rc4 is the latest), the packed add-ins go into a "publish" subdirectory of your build output.