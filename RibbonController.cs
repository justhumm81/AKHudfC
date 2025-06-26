using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;


// CUI Ribbon exampe for ExcelDNA
// https://github.com/Excel-DNA/Samples/tree/master/Ribbon

namespace Ribbon
{
    // RIBBON XLM information contained in *Addin.dna file.
    //
    // ==================== START of Class ====================
    [ComVisible(true)]
    public class RibbonController : ExcelRibbon
    {
        dynamic Application = ExcelDnaUtil.Application;
        string strFunc = "";

        // ------------------------------------------------------------
        // OPEN EXCEL FUNCTION DIALOG / WIZARD
        // ------------------------------------------------------------
        public void OpenDialog(IRibbonControl control)
        {
            // open the function wizard
            dynamic cellActive = Application.ActiveCell;
            cellActive.FunctionWizard();
        }

        // ------------------------------------------------------------
        // FUNCTION LIBRARY CONTROLS
        // MENU 1 - Text
        // ------------------------------------------------------------
        public void Func211(IRibbonControl control)
        {
            strFunc = "=C_ChangeCase()";
            FuncAll(strFunc);
        }
        public void Func212(IRibbonControl control)
        {
            strFunc = "=C_Formula()";
            FuncAll(strFunc);
        }
        public void Func213(IRibbonControl control)
        {
            //strFunc = "=C_RevStr(\"TAC\")"; // Excel doesn't recognized single quotation marks
            strFunc = "=C_RevStr()";
            FuncAll(strFunc);
        }

        // ------------------------------------------------------------
        // FUNCTION LIBRARY CONTROLS
        // MENU 2 - Lookup & Ref
        // ------------------------------------------------------------
        public void Func221(IRibbonControl control)
        {
            strFunc = "=C_MMatch()";
            FuncAll(strFunc);
        }

        // ------------------------------------------------------------
        // FUNCTION LIBRARY CONTROLS
        // MENU 3 - Math & Trig
        // ------------------------------------------------------------
        public void Func231(IRibbonControl control)
        {
            strFunc = "=C_AbsMax()";
            FuncAll(strFunc);
        }
        public void Func232(IRibbonControl control)
        {
            strFunc = "=C_AbsMin()";
            FuncAll(strFunc);
        }
        public void Func233(IRibbonControl control)
        {
            strFunc = "=C_Linterp()";
            FuncAll(strFunc);
        }
        public void FuncAll(string strFunc)
        {
            dynamic cellActive = Application.ActiveCell;
            cellActive.formula = strFunc;  // just puts function string in cell
            cellActive.FunctionWizard(); // opens up the arguments dialog box for UDF 
        }

        // ------------------------------------------------------------
        // VERSION BUTTON CONTROLS
        // ------------------------------------------------------------
        public void PressButtonVersion(IRibbonControl control)
        {
            // compiler run date (hopefully)
            // https://stackoverflow.com/questions/1276437/compile-date-and-time
            //
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            DateTime dateBuild = new DateTime(2000, 1, 1)
                .AddDays(version.Build)
                .AddSeconds(version.Revision * 2);

            // const string
            string strCaption = "AKHudfC";
            string strMessage1 = "Excel Add-in of User Defined Functions (UDF).";
            string strDate = dateBuild.ToString("dd MMMM yyyy HH:mm");
            MessageBox.Show(
                "Beta Version" + Environment.NewLine + strMessage1 + Environment.NewLine + "Written in C#, Compile Date: " + strDate,
                strCaption,
                MessageBoxButtons.OK,
                MessageBoxIcon.Information); ;
        }
    }
    // ========== END Class ==========
}