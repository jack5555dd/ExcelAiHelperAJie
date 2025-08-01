using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools;

namespace ExcelAIHelper
{
    public partial class ThisAddIn
    {
        internal static CustomTaskPane ChatPane;
        
        // Application property is already provided by the base class

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                // Configure SSL/TLS settings globally for the application
                // This is required for HTTPS connections to modern APIs
                System.Net.ServicePointManager.SecurityProtocol = 
                    System.Net.SecurityProtocolType.Tls12 | 
                    System.Net.SecurityProtocolType.Tls11 | 
                    System.Net.SecurityProtocolType.Tls;
                
                // Set connection limits for better performance
                System.Net.ServicePointManager.DefaultConnectionLimit = 10;
                
                System.Diagnostics.Debug.WriteLine("SSL/TLS configuration applied: " + 
                    System.Net.ServicePointManager.SecurityProtocol.ToString());
                
                // Create chat pane
                var chatControl = new ChatPaneControl();
                ChatPane = this.CustomTaskPanes.Add(chatControl, "AI Chat");
                ChatPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight; // 右侧
                ChatPane.Visible = false;   // 初始隐藏
                
                System.Diagnostics.Debug.WriteLine("Excel AI Helper started successfully");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"启动失败: {ex.Message}", "Excel AI Helper", MessageBoxButtons.OK, MessageBoxIcon.Error);
                System.Diagnostics.Debug.WriteLine($"Failed to start Excel AI Helper: {ex.Message}");
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                // 清理聚光灯资源
                SpotlightManager.Cleanup();
                
                // 清理快速录入资源
                QuickInputManager.Stop();
                
                System.Diagnostics.Debug.WriteLine("Excel AI Helper shutting down");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Shutdown cleanup error: {ex.Message}");
            }
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new AiRibbon();
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