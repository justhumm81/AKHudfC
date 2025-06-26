using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using ExcelDna.Registration;

// note that this namespace is split across multiple CS files
namespace AKHudfC
{
    // In order for the on-sheet Tool-Tips and descriptions to work, you need to
    // Register the IntelliSenseServer in your add-in's AutoOpen() implmentation
    // https://github.com/Excel-DNA/IntelliSense/wiki/Usage-Instructions
    //
    // And supposedly you need to AutoOpen the params registration
    // https://groups.google.com/g/exceldna/c/kf76nqAqDUo

    public class Auto : IExcelAddIn   // this is a Class inheritance, where "Auto" is the child
    {
        public void AutoOpen()
        {
            ExcelRegistration
                .GetExcelFunctions()
                    .ProcessParamsRegistrations()
                .RegisterFunctions();

            // the intllisenseServer has to be installed AFTER the Excel Function Registration
            IntelliSenseServer.Install();
        }

        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();
        }
    }
}