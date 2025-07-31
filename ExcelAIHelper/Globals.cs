using System;

namespace ExcelAIHelper
{
    /// <summary>
    /// Helper class for working with Excel Application
    /// </summary>
    public static class ExcelAppHelper
    {
        /// <summary>
        /// Gets the Excel Application instance
        /// </summary>
        public static Microsoft.Office.Interop.Excel.Application GetExcelApp()
        {
            return Globals.ThisAddIn.Application;
        }
    }
}