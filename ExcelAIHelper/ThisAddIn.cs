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
                System.Diagnostics.Debug.WriteLine("ThisAddIn_Startup called");
                
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
                
                // 延迟初始化UI组件，避免在Excel完全启动前创建复杂控件
                System.Windows.Forms.Timer startupTimer = new System.Windows.Forms.Timer();
                startupTimer.Interval = 1000; // 1秒延迟
                startupTimer.Tick += (s, args) =>
                {
                    startupTimer.Stop();
                    startupTimer.Dispose();
                    DelayedInitialization();
                };
                startupTimer.Start();
                
                System.Diagnostics.Debug.WriteLine("Excel AI Helper startup initiated");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ThisAddIn_Startup failed: {ex}");
                // 不显示MessageBox，避免阻塞Excel启动
            }
        }
        
        private void DelayedInitialization()
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("DelayedInitialization called");
                
                // Create chat pane with error handling
                var chatControl = new ChatPaneControl();
                ChatPane = this.CustomTaskPanes.Add(chatControl, "AI Chat");
                ChatPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                ChatPane.Visible = false;
                
                System.Diagnostics.Debug.WriteLine("Excel AI Helper started successfully");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"DelayedInitialization failed: {ex}");
                // 如果延迟初始化也失败，记录错误但不阻塞Excel
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