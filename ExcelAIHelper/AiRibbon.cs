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
        /// <summary>
        /// 聚光灯主按钮点击 - 切换聚光灯状态
        /// </summary>
        public void OnSpotlightToggle(Office.IRibbonControl control)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("OnSpotlightToggle called");
                SpotlightManager.Toggle();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"OnSpotlightToggle error: {ex.Message}");
                MessageBox.Show($"聚光灯操作失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        /// <summary>
        /// 关闭聚光灯
        /// </summary>
        public void OnSpotlightClose(Office.IRibbonControl control)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("OnSpotlightClose called");
                if (SpotlightManager.IsActive)
                {
                    SpotlightManager.Stop();
                    MessageBox.Show("聚光灯已关闭", "聚光灯", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("聚光灯当前未开启", "聚光灯", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"OnSpotlightClose error: {ex.Message}");
                MessageBox.Show($"关闭聚光灯失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        /// <summary>
        /// 聚光灯设置
        /// </summary>
        public void OnSpotlightSettings(Office.IRibbonControl control)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("OnSpotlightSettings called");
                ShowSpotlightSettings();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"OnSpotlightSettings error: {ex.Message}");
                MessageBox.Show($"打开聚光灯设置失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        /// <summary>
        /// 显示聚光灯设置对话框
        /// </summary>
        private void ShowSpotlightSettings()
        {
            try
            {
                using (var settingsForm = new SpotlightSettingsForm())
                {
                    if (settingsForm.ShowDialog() == DialogResult.OK)
                    {
                        // 设置已在SpotlightSettingsForm中应用
                        MessageBox.Show("设置已保存", "聚光灯设置", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"显示设置对话框失败: {ex.Message}");
                MessageBox.Show($"显示设置对话框失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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