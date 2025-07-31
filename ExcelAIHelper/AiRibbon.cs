using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;
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

        #region 面板功能
        public void OnChatPaneClick(Office.IRibbonControl control)
        {
            System.Diagnostics.Debug.WriteLine("OnChatPaneClick called");
            var pane = ThisAddIn.ChatPane;
            pane.Visible = !pane.Visible;   // 切换显示
        }
        #endregion

        #region 工具箱功能
        // 数据工具
        public void OnDataAnalysisClick(Office.IRibbonControl control)
        {
            MessageBox.Show("数据分析功能正在开发中...", "AI 助手", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void OnDataCleanClick(Office.IRibbonControl control)
        {
            MessageBox.Show("数据清洗功能正在开发中...", "AI 助手", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void OnDataImportClick(Office.IRibbonControl control)
        {
            MessageBox.Show("数据导入功能正在开发中...", "AI 助手", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // 格式工具
        public void OnAutoFormatClick(Office.IRibbonControl control)
        {
            MessageBox.Show("智能格式化功能正在开发中...", "AI 助手", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void OnStyleApplyClick(Office.IRibbonControl control)
        {
            MessageBox.Show("样式应用功能正在开发中...", "AI 助手", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void OnTableFormatClick(Office.IRibbonControl control)
        {
            MessageBox.Show("表格美化功能正在开发中...", "AI 助手", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // 公式工具
        public void OnFormulaHelperClick(Office.IRibbonControl control)
        {
            MessageBox.Show("公式助手功能正在开发中...", "AI 助手", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void OnFormulaCheckClick(Office.IRibbonControl control)
        {
            MessageBox.Show("公式检查功能正在开发中...", "AI 助手", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void OnFormulaOptimizeClick(Office.IRibbonControl control)
        {
            MessageBox.Show("公式优化功能正在开发中...", "AI 助手", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        #endregion

        #region 聚光灯功能
        public void OnSpotlightClick(Office.IRibbonControl control)
        {
            SpotlightManager.Toggle();
        }


        #endregion

        #region 设置和帮助功能
        public void OnSettingsClick(Office.IRibbonControl control)
        {
            try
            {
                var settingsForm = new SettingsForm();
                settingsForm.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开设置窗口时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnExampleClick(Office.IRibbonControl control)
        {
            try
            {
                var exampleForm = new ExampleForm();
                exampleForm.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开示例窗口时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnAboutClick(Office.IRibbonControl control)
        {
            try
            {
                var aboutForm = new AboutForm();
                aboutForm.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开关于窗口时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion
    }
}