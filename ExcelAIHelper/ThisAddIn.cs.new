using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools;

namespace ExcelAIHelper
{
    public partial class ThisAddIn
    {
        internal static CustomTaskPane ChatPane;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            var chatControl = new ChatPaneControl();
            ChatPane = this.CustomTaskPanes.Add(chatControl, "AI Chat");
            ChatPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight; // 右侧
            ChatPane.Visible = false;   // 初始隐藏
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new AiRibbon();   // 告诉 VSTO 用这个类
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}