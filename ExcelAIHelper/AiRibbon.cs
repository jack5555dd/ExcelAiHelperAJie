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
        private bool spotlightEnabled = false;
        private Color spotlightColor = Color.Yellow;

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
            try
            {
                spotlightEnabled = !spotlightEnabled;
                if (spotlightEnabled)
                {
                    ApplySpotlight();
                    MessageBox.Show("聚光灯已开启，点击单元格或选择区域查看效果", "AI 助手", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    RemoveSpotlight();
                    MessageBox.Show("聚光灯已关闭", "AI 助手", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"聚光灯功能出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnSpotlightOffClick(Office.IRibbonControl control)
        {
            spotlightEnabled = false;
            RemoveSpotlight();
            MessageBox.Show("聚光灯已关闭", "AI 助手", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void OnSpotlightYellowClick(Office.IRibbonControl control)
        {
            spotlightColor = Color.Yellow;
            if (spotlightEnabled) ApplySpotlight();
        }

        public void OnSpotlightBlueClick(Office.IRibbonControl control)
        {
            spotlightColor = Color.LightBlue;
            if (spotlightEnabled) ApplySpotlight();
        }

        public void OnSpotlightGreenClick(Office.IRibbonControl control)
        {
            spotlightColor = Color.LightGreen;
            if (spotlightEnabled) ApplySpotlight();
        }

        public void OnSpotlightRedClick(Office.IRibbonControl control)
        {
            spotlightColor = Color.LightCoral;
            if (spotlightEnabled) ApplySpotlight();
        }

        public void OnSpotlightCustomClick(Office.IRibbonControl control)
        {
            using (ColorDialog colorDialog = new ColorDialog())
            {
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    spotlightColor = colorDialog.Color;
                    if (spotlightEnabled) ApplySpotlight();
                }
            }
        }

        private void ApplySpotlight()
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Excel.Range selection = app.Selection as Excel.Range;
                if (selection != null)
                {
                    selection.Interior.Color = ColorTranslator.ToOle(spotlightColor);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ApplySpotlight error: {ex.Message}");
            }
        }

        private void RemoveSpotlight()
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Excel.Range selection = app.Selection as Excel.Range;
                if (selection != null)
                {
                    selection.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"RemoveSpotlight error: {ex.Message}");
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