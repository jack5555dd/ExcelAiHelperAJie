using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace ExcelAIHelper
{
    [ComVisible(true)]
    public class AiRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public string GetCustomUI(string ribbonID)
        {
            System.Diagnostics.Debug.WriteLine($"GetCustomUI called with {ribbonID}");
            return new StreamReader(Assembly.GetExecutingAssembly()
                .GetManifestResourceStream("ExcelAIHelper.AiRibbon.xml")).ReadToEnd();
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            System.Diagnostics.Debug.WriteLine("Ribbon_Load called");
            this.ribbon = ribbonUI;
        }

        public void OnChatPaneClick(Office.IRibbonControl control)
        {
            System.Diagnostics.Debug.WriteLine("OnChatPaneClick called");
            var pane = ThisAddIn.ChatPane;
            pane.Visible = !pane.Visible;   // 切换显示
        }

        public void OnApiClick(Office.IRibbonControl control)
        {
            System.Diagnostics.Debug.WriteLine("OnApiClick called");
            using (var dlg = new ApiSettingsForm())
            {
                dlg.ShowDialog();
            }
        }
    }
}