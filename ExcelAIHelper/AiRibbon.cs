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
        
        // 快速录入
        public void OnQuickInputSettingsClick(Office.IRibbonControl control)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("OnQuickInputSettingsClick called");
                var form = new QuickInputSettingsForm();
                form.Show(); // 使用非模态对话框
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"OnQuickInputSettingsClick error: {ex.Message}");
                MessageBox.Show($"打开快速录入设置失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        public void OnSequenceNumbersClick(Office.IRibbonControl control)
        {
            MessageBox.Show("序号功能正在开发中...", "快速录入", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnDateSeriesClick(Office.IRibbonControl control)
        {
            MessageBox.Show("日期序列功能正在开发中...", "快速录入", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnCustomSeriesClick(Office.IRibbonControl control)
        {
            MessageBox.Show("自定义序列功能正在开发中...", "快速录入", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        // 提取/过滤
        public void OnExtractNumbersClick(Office.IRibbonControl control)
        {
            MessageBox.Show("提取数字功能正在开发中...", "提取/过滤", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnExtractTextClick(Office.IRibbonControl control)
        {
            MessageBox.Show("提取文本功能正在开发中...", "提取/过滤", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnExtractDateClick(Office.IRibbonControl control)
        {
            MessageBox.Show("提取日期功能正在开发中...", "提取/过滤", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnRemoveDuplicatesClick(Office.IRibbonControl control)
        {
            MessageBox.Show("去重功能正在开发中...", "提取/过滤", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnAdvancedFilterClick(Office.IRibbonControl control)
        {
            MessageBox.Show("高级筛选功能正在开发中...", "提取/过滤", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        // 数值批量计算
        public void OnBatchAddClick(Office.IRibbonControl control)
        {
            MessageBox.Show("批量加法功能正在开发中...", "数值批量计算", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnBatchSubtractClick(Office.IRibbonControl control)
        {
            MessageBox.Show("批量减法功能正在开发中...", "数值批量计算", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnBatchMultiplyClick(Office.IRibbonControl control)
        {
            MessageBox.Show("批量乘法功能正在开发中...", "数值批量计算", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnBatchDivideClick(Office.IRibbonControl control)
        {
            MessageBox.Show("批量除法功能正在开发中...", "数值批量计算", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnBatchPercentClick(Office.IRibbonControl control)
        {
            MessageBox.Show("批量百分比功能正在开发中...", "数值批量计算", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnBatchPowerClick(Office.IRibbonControl control)
        {
            MessageBox.Show("批量乘方功能正在开发中...", "数值批量计算", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        // 四舍五入
        public void OnRoundToIntegerClick(Office.IRibbonControl control)
        {
            MessageBox.Show("取整功能正在开发中...", "四舍五入", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnRoundToDecimalClick(Office.IRibbonControl control)
        {
            MessageBox.Show("保留小数功能正在开发中...", "四舍五入", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnRoundUpClick(Office.IRibbonControl control)
        {
            MessageBox.Show("向上取整功能正在开发中...", "四舍五入", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnRoundDownClick(Office.IRibbonControl control)
        {
            MessageBox.Show("向下取整功能正在开发中...", "四舍五入", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnRoundToThousandClick(Office.IRibbonControl control)
        {
            MessageBox.Show("千位取整功能正在开发中...", "四舍五入", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnRoundToTenThousandClick(Office.IRibbonControl control)
        {
            MessageBox.Show("万位取整功能正在开发中...", "四舍五入", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        // 数字格式刷
        public void OnFormatGeneralClick(Office.IRibbonControl control)
        {
            MessageBox.Show("常规格式功能正在开发中...", "数字格式刷", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnFormatNumberClick(Office.IRibbonControl control)
        {
            MessageBox.Show("数值格式功能正在开发中...", "数字格式刷", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnFormatPercentClick(Office.IRibbonControl control)
        {
            MessageBox.Show("百分比格式功能正在开发中...", "数字格式刷", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnFormatDateClick(Office.IRibbonControl control)
        {
            MessageBox.Show("日期格式功能正在开发中...", "数字格式刷", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnFormatTimeClick(Office.IRibbonControl control)
        {
            MessageBox.Show("时间格式功能正在开发中...", "数字格式刷", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnFormatTextClick(Office.IRibbonControl control)
        {
            MessageBox.Show("文本格式功能正在开发中...", "数字格式刷", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        // 金额格式刷
        public void OnFormatCNYClick(Office.IRibbonControl control)
        {
            MessageBox.Show("人民币格式功能正在开发中...", "金额格式刷", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnFormatUSDClick(Office.IRibbonControl control)
        {
            MessageBox.Show("美元格式功能正在开发中...", "金额格式刷", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnFormatEURClick(Office.IRibbonControl control)
        {
            MessageBox.Show("欧元格式功能正在开发中...", "金额格式刷", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnFormatAccountingClick(Office.IRibbonControl control)
        {
            MessageBox.Show("会计格式功能正在开发中...", "金额格式刷", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnFormatFinancialClick(Office.IRibbonControl control)
        {
            MessageBox.Show("财务格式功能正在开发中...", "金额格式刷", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        // 姓名处理
        public void OnSplitNameClick(Office.IRibbonControl control)
        {
            MessageBox.Show("姓名拆分功能正在开发中...", "姓名处理", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnMergeNameClick(Office.IRibbonControl control)
        {
            MessageBox.Show("姓名合并功能正在开发中...", "姓名处理", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnNameValidationClick(Office.IRibbonControl control)
        {
            MessageBox.Show("姓名校验功能正在开发中...", "姓名处理", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnRemoveSpacesClick(Office.IRibbonControl control)
        {
            MessageBox.Show("去除空格功能正在开发中...", "姓名处理", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnStandardizeNameClick(Office.IRibbonControl control)
        {
            MessageBox.Show("姓名标准化功能正在开发中...", "姓名处理", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        // 身份证
        public void OnIDValidationClick(Office.IRibbonControl control)
        {
            MessageBox.Show("身份证校验功能正在开发中...", "身份证", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnExtractBirthdayClick(Office.IRibbonControl control)
        {
            MessageBox.Show("提取生日功能正在开发中...", "身份证", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnExtractAgeClick(Office.IRibbonControl control)
        {
            MessageBox.Show("提取年龄功能正在开发中...", "身份证", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnExtractGenderClick(Office.IRibbonControl control)
        {
            MessageBox.Show("提取性别功能正在开发中...", "身份证", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnIDMaskClick(Office.IRibbonControl control)
        {
            MessageBox.Show("身份证脱敏功能正在开发中...", "身份证", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnIDFormatClick(Office.IRibbonControl control)
        {
            MessageBox.Show("身份证格式化功能正在开发中...", "身份证", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        // 手机号
        public void OnPhoneValidationClick(Office.IRibbonControl control)
        {
            MessageBox.Show("手机号校验功能正在开发中...", "手机号", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnPhoneMaskClick(Office.IRibbonControl control)
        {
            MessageBox.Show("手机号脱敏功能正在开发中...", "手机号", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnPhoneFormatClick(Office.IRibbonControl control)
        {
            MessageBox.Show("手机号格式化功能正在开发中...", "手机号", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnExtractCarrierClick(Office.IRibbonControl control)
        {
            MessageBox.Show("提取运营商功能正在开发中...", "手机号", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public void OnExtractRegionClick(Office.IRibbonControl control)
        {
            MessageBox.Show("提取归属地功能正在开发中...", "手机号", MessageBoxButtons.OK, MessageBoxIcon.Information);
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