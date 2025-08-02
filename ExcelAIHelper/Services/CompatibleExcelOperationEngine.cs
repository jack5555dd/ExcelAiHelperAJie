using System;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;
using ExcelAIHelper.Exceptions;
using ExcelAIHelper.Models;

namespace ExcelAIHelper.Services
{
    /// <summary>
    /// 兼容版Excel操作引擎 - 支持Microsoft Excel和WPS表格
    /// </summary>
    public class CompatibleExcelOperationEngine
    {
        private readonly IExcelApplication _excelApp;
        
        /// <summary>
        /// 创建兼容版Excel操作引擎实例
        /// </summary>
        public CompatibleExcelOperationEngine()
        {
            _excelApp = ExcelApplicationFactory.GetCurrentApplication();
        }
        
        /// <summary>
        /// 获取当前选中的区域
        /// </summary>
        /// <returns>选中的区域</returns>
        public async Task<object> GetSelectedRangeAsync()
        {
            return await Task.Run(() => {
                try
                {
                    return _excelApp.Selection;
                }
                catch (Exception ex)
                {
                    throw new AiOperationException("Failed to get selected range", ex);
                }
            });
        }

        /// <summary>
        /// 根据指令获取目标区域，带智能回退
        /// </summary>
        /// <param name="targetRangeSpec">指令中的目标区域规范</param>
        /// <returns>解析的目标区域</returns>
        private async Task<object> GetTargetRangeAsync(string targetRangeSpec)
        {
            try
            {
                // 如果提供了特定区域，使用它
                if (!string.IsNullOrEmpty(targetRangeSpec))
                {
                    // 清理区域规范
                    string cleanRange = targetRangeSpec.Trim().ToUpper();
                    
                    // 处理常见变体
                    if (cleanRange.Contains("CURRENT") || cleanRange.Contains("SELECTED") || cleanRange.Contains("SELECTION"))
                    {
                        return await GetSelectedRangeAsync();
                    }
                    
                    // 尝试解析为Excel区域
                    try
                    {
                        return _excelApp.GetRange(cleanRange);
                    }
                    catch
                    {
                        // 如果解析失败，回退到选中区域
                        System.Diagnostics.Debug.WriteLine($"Failed to parse range '{targetRangeSpec}', using selected range");
                        return await GetSelectedRangeAsync();
                    }
                }
                else
                {
                    // 没有指定目标，使用选中区域
                    return await GetSelectedRangeAsync();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error resolving target range '{targetRangeSpec}': {ex.Message}");
                // 最终回退：尝试获取选中区域
                try
                {
                    return await GetSelectedRangeAsync();
                }
                catch
                {
                    // 终极回退：使用A1
                    return _excelApp.GetRange("A1");
                }
            }
        }
        
        /// <summary>
        /// 从区域获取数据
        /// </summary>
        /// <param name="range">要获取数据的区域</param>
        /// <returns>包含区域数据的DataTable</returns>
        public async Task<DataTable> GetRangeDataAsync(object range)
        {
            return await Task.Run(() => {
                try
                {
                    if (range == null) throw new ArgumentNullException(nameof(range));
                    
                    var data = new DataTable();
                    
                    // 获取区域的值
                    object rangeValue = _excelApp.GetCellValue(range);
                    if (rangeValue == null) return data;
                    
                    // 获取区域地址以确定大小
                    string address = _excelApp.GetRangeAddress(range);
                    var (rows, cols) = ParseRangeSize(address);
                    
                    // 创建列
                    for (int c = 1; c <= cols; c++)
                    {
                        data.Columns.Add($"Column{c}");
                    }
                    
                    // 处理单元格vs多元格区域
                    if (rows == 1 && cols == 1)
                    {
                        // 单元格
                        var row = data.NewRow();
                        row[0] = rangeValue ?? DBNull.Value;
                        data.Rows.Add(row);
                    }
                    else
                    {
                        // 多元格区域 - 这里需要根据具体实现调整
                        // 由于我们使用通用接口，这部分可能需要特殊处理
                        var row = data.NewRow();
                        row[0] = rangeValue ?? DBNull.Value;
                        data.Rows.Add(row);
                    }
                    
                    return data;
                }
                catch (Exception ex)
                {
                    throw new AiOperationException("Failed to get range data", ex);
                }
            });
        }
        
        /// <summary>
        /// 设置单元格值
        /// </summary>
        /// <param name="range">要设置的区域</param>
        /// <param name="value">要设置的值</param>
        public async Task SetCellValueAsync(object range, object value)
        {
            await Task.Run(() => {
                try
                {
                    if (range == null) throw new ArgumentNullException(nameof(range));
                    _excelApp.SetCellValue(range, value);
                }
                catch (Exception ex)
                {
                    throw new AiOperationException("Failed to set cell value", ex);
                }
            });
        }

        /// <summary>
        /// 清除区域内容
        /// </summary>
        /// <param name="range">要清除的区域</param>
        public async Task ClearRangeContentAsync(object range)
        {
            await Task.Run(() => {
                try
                {
                    if (range == null) throw new ArgumentNullException(nameof(range));
                    _excelApp.SetCellValue(range, "");
                }
                catch (Exception ex)
                {
                    throw new AiOperationException("Failed to clear range content", ex);
                }
            });
        }

        /// <summary>
        /// 设置区域的背景色
        /// </summary>
        /// <param name="range">要格式化的区域</param>
        /// <param name="color">颜色（十六进制格式如#FF0000或颜色名称）</param>
        public async Task SetBackgroundColorAsync(object range, string color)
        {
            await Task.Run(() => {
                try
                {
                    if (range == null) throw new ArgumentNullException(nameof(range));
                    if (string.IsNullOrEmpty(color)) throw new ArgumentNullException(nameof(color));
                    
                    // 这里需要根据具体的Excel应用程序类型来实现
                    // 由于接口限制，我们可能需要使用反射或特定的实现
                    System.Diagnostics.Debug.WriteLine($"Setting background color {color} for range {_excelApp.GetRangeAddress(range)}");
                    
                    // 暂时记录操作，具体实现可能需要在各自的应用程序类中完成
                    _excelApp.ShowMessage($"背景色设置功能正在开发中。颜色: {color}", "提示");
                }
                catch (Exception ex)
                {
                    throw new AiOperationException($"Failed to set background color '{color}'", ex);
                }
            });
        }

        /// <summary>
        /// 应用公式到区域
        /// </summary>
        /// <param name="range">要应用公式的区域</param>
        /// <param name="formula">要应用的公式</param>
        public async Task ApplyFormulaAsync(object range, string formula)
        {
            await Task.Run(() => {
                try
                {
                    if (range == null) throw new ArgumentNullException(nameof(range));
                    if (string.IsNullOrEmpty(formula)) throw new ArgumentNullException(nameof(formula));
                    
                    // 验证和改进公式
                    string validatedFormula = ValidateAndImproveFormula(formula, range);
                    
                    System.Diagnostics.Debug.WriteLine($"Applying formula: {validatedFormula} to range: {_excelApp.GetRangeAddress(range)}");
                    
                    _excelApp.SetCellValue(range, validatedFormula);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Formula application failed: {ex.Message}");
                    throw new AiOperationException($"Failed to apply formula '{formula}': {ex.Message}", ex);
                }
            });
        }

        /// <summary>
        /// 在应用公式之前验证和改进公式
        /// </summary>
        /// <param name="formula">原始公式</param>
        /// <param name="targetRange">目标区域</param>
        /// <returns>验证和改进后的公式</returns>
        private string ValidateAndImproveFormula(string formula, object targetRange)
        {
            if (string.IsNullOrEmpty(formula))
                return formula;

            // 确保公式以=开头
            if (!formula.StartsWith("="))
                formula = "=" + formula;

            // 处理常见的公式改进
            if (formula.Equals("=SUM()", StringComparison.OrdinalIgnoreCase))
            {
                // 尝试智能确定SUM范围
                string suggestedRange = GetSuggestedSumRange(targetRange);
                if (!string.IsNullOrEmpty(suggestedRange))
                {
                    formula = $"=SUM({suggestedRange})";
                    System.Diagnostics.Debug.WriteLine($"Improved SUM formula: {formula}");
                }
            }

            return formula;
        }

        /// <summary>
        /// 基于附近数据为SUM公式获取建议范围
        /// </summary>
        /// <param name="targetRange">目标区域</param>
        /// <returns>建议的范围字符串或null</returns>
        private string GetSuggestedSumRange(object targetRange)
        {
            try
            {
                // 这是一个简化的实现，实际可能需要更复杂的逻辑
                string address = _excelApp.GetRangeAddress(targetRange);
                var match = System.Text.RegularExpressions.Regex.Match(address, @"([A-Z]+)(\d+)");
                if (match.Success)
                {
                    string column = match.Groups[1].Value;
                    int row = int.Parse(match.Groups[2].Value);
                    if (row > 1)
                    {
                        return $"{column}1:{column}{row - 1}";
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error suggesting SUM range: {ex.Message}");
            }

            return null;
        }

        /// <summary>
        /// 解析区域大小
        /// </summary>
        /// <param name="address">区域地址</param>
        /// <returns>行数和列数的元组</returns>
        private (int rows, int cols) ParseRangeSize(string address)
        {
            try
            {
                if (string.IsNullOrEmpty(address))
                    return (1, 1);

                // 简单的单元格地址（如A1）
                if (!address.Contains(":"))
                    return (1, 1);

                // 区域地址（如A1:B2）
                var parts = address.Split(':');
                if (parts.Length == 2)
                {
                    var start = ParseCellAddress(parts[0]);
                    var end = ParseCellAddress(parts[1]);
                    return (end.row - start.row + 1, end.col - start.col + 1);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error parsing range size: {ex.Message}");
            }

            return (1, 1);
        }

        /// <summary>
        /// 解析单元格地址
        /// </summary>
        /// <param name="cellAddress">单元格地址（如A1）</param>
        /// <returns>行号和列号的元组</returns>
        private (int row, int col) ParseCellAddress(string cellAddress)
        {
            var match = System.Text.RegularExpressions.Regex.Match(cellAddress.Replace("$", ""), @"([A-Z]+)(\d+)");
            if (match.Success)
            {
                string columnLetters = match.Groups[1].Value;
                int row = int.Parse(match.Groups[2].Value);
                int col = ColumnLettersToNumber(columnLetters);
                return (row, col);
            }
            return (1, 1);
        }

        /// <summary>
        /// 将列字母转换为数字
        /// </summary>
        /// <param name="columnLetters">列字母（如A, B, AA）</param>
        /// <returns>列号</returns>
        private int ColumnLettersToNumber(string columnLetters)
        {
            int result = 0;
            for (int i = 0; i < columnLetters.Length; i++)
            {
                result = result * 26 + (columnLetters[i] - 'A' + 1);
            }
            return result;
        }

        /// <summary>
        /// 执行VBA代码
        /// </summary>
        /// <param name="vbaCode">VBA代码</param>
        public async Task ExecuteVbaCodeAsync(string vbaCode)
        {
            await Task.Run(() => {
                try
                {
                    _excelApp.ExecuteVbaCode(vbaCode);
                }
                catch (Exception ex)
                {
                    throw new AiOperationException("Failed to execute VBA code", ex);
                }
            });
        }

        /// <summary>
        /// 检查VBA访问权限
        /// </summary>
        /// <returns>是否有VBA访问权限</returns>
        public bool HasVbaAccess()
        {
            return _excelApp.HasVbaAccess();
        }

        /// <summary>
        /// 获取当前应用程序名称
        /// </summary>
        /// <returns>应用程序名称</returns>
        public string GetApplicationName()
        {
            return ExcelApplicationFactory.GetCurrentApplicationName();
        }
    }
}