using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelAIHelper
{
    internal static class SpotlightManager
    {
        private static bool _isActive = false;
        private static Excel.AppEvents_SheetSelectionChangeEventHandler _selectionChangeHandler;
        private static Color _currentHighlightColor = Color.FromArgb(255, 255, 0); // 默认黄色
        private static SpotlightMode _currentMode = SpotlightMode.Both;
        
        // 公共属性
        public static bool IsActive => _isActive;
        public static Color CurrentHighlightColor => _currentHighlightColor;
        public static SpotlightMode CurrentMode => _currentMode;
        
        /// <summary>
        /// 启动聚光灯功能
        /// </summary>
        public static void Start()
        {
            if (_isActive) return;
            
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app == null)
                {
                    System.Diagnostics.Debug.WriteLine("Excel应用程序未就绪");
                    return;
                }
                
                // 注册选择变化事件
                _selectionChangeHandler = new Excel.AppEvents_SheetSelectionChangeEventHandler(OnSelectionChange);
                app.SheetSelectionChange += _selectionChangeHandler;
                
                _isActive = true;
                
                // 立即高亮当前选择
                try
                {
                    var currentSelection = app.Selection as Excel.Range;
                    if (currentSelection != null)
                    {
                        HighlightCrossHair(currentSelection);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"初始高亮失败: {ex.Message}");
                }
                
                System.Diagnostics.Debug.WriteLine("聚光灯已启动");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"启动聚光灯失败: {ex.Message}");
                throw; // 重新抛出异常，让调用者处理
            }
        }
        
        /// <summary>
        /// 关闭聚光灯功能
        /// </summary>
        public static void Stop()
        {
            if (!_isActive) return;
            
            try
            {
                var app = Globals.ThisAddIn.Application;
                
                // 取消注册事件
                if (_selectionChangeHandler != null && app != null)
                {
                    app.SheetSelectionChange -= _selectionChangeHandler;
                    _selectionChangeHandler = null;
                }
                
                // 清除所有高亮
                ClearAllHighlights();
                
                _isActive = false;
                
                System.Diagnostics.Debug.WriteLine("聚光灯已关闭");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"关闭聚光灯失败: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 切换聚光灯状态
        /// </summary>
        public static void Toggle()
        {
            try
            {
                if (_isActive)
                {
                    Stop();
                    MessageBox.Show("聚光灯已关闭", "聚光灯", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    Start();
                    MessageBox.Show("聚光灯已开启\n移动到不同单元格查看十字光标效果", "聚光灯", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"切换聚光灯失败: {ex.Message}");
                MessageBox.Show($"切换聚光灯失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        /// <summary>
        /// 选择变化事件处理
        /// </summary>
        private static void OnSelectionChange(object sh, Excel.Range target)
        {
            try
            {
                if (!_isActive) return;
                
                // 清除之前的高亮
                ClearAllHighlights();
                
                // 高亮新的十字光标
                HighlightCrossHair(target);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"选择变化处理失败: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 高亮十字光标（行和列）
        /// </summary>
        private static void HighlightCrossHair(Excel.Range target)
        {
            try
            {
                if (target == null) return;
                
                var worksheet = target.Worksheet;
                int targetRow = target.Row;
                int targetColumn = target.Column;
                
                // 转换颜色为Excel格式
                int excelColor = ColorTranslator.ToOle(_currentHighlightColor);
                
                // 根据模式高亮不同区域
                switch (_currentMode)
                {
                    case SpotlightMode.Both:
                        HighlightRow(worksheet, targetRow, excelColor);
                        HighlightColumn(worksheet, targetColumn, excelColor);
                        break;
                    case SpotlightMode.Row:
                        HighlightRow(worksheet, targetRow, excelColor);
                        break;
                    case SpotlightMode.Column:
                        HighlightColumn(worksheet, targetColumn, excelColor);
                        break;
                }
                
                System.Diagnostics.Debug.WriteLine($"已高亮十字光标: 行{targetRow}, 列{targetColumn}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"高亮十字光标失败: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 高亮整行
        /// </summary>
        private static void HighlightRow(Excel.Worksheet worksheet, int row, int color)
        {
            try
            {
                // 获取可见区域的行范围（提高性能）
                var visibleCells = worksheet.Application.ActiveWindow.VisibleRange;
                int startCol = Math.Max(1, visibleCells.Column - 5);
                int endCol = Math.Min(visibleCells.Column + visibleCells.Columns.Count + 5, 50); // 限制在50列内
                
                var rowRange = worksheet.Range[worksheet.Cells[row, startCol], worksheet.Cells[row, endCol]];
                
                // 设置背景色
                rowRange.Interior.Color = color;
                
                // 设置透明度
                try
                {
                    rowRange.Interior.TintAndShade = 0.7; // 使颜色更浅
                }
                catch { }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"高亮行失败: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 高亮整列
        /// </summary>
        private static void HighlightColumn(Excel.Worksheet worksheet, int column, int color)
        {
            try
            {
                // 获取可见区域的列范围（提高性能）
                var visibleCells = worksheet.Application.ActiveWindow.VisibleRange;
                int startRow = Math.Max(1, visibleCells.Row - 5);
                int endRow = Math.Min(visibleCells.Row + visibleCells.Rows.Count + 5, 100); // 限制在100行内
                
                var columnRange = worksheet.Range[worksheet.Cells[startRow, column], worksheet.Cells[endRow, column]];
                
                // 设置背景色
                columnRange.Interior.Color = color;
                
                // 设置透明度
                try
                {
                    columnRange.Interior.TintAndShade = 0.7; // 使颜色更浅
                }
                catch { }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"高亮列失败: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 清除所有高亮
        /// </summary>
        private static void ClearAllHighlights()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var worksheet = app.ActiveSheet as Excel.Worksheet;
                
                if (worksheet == null) return;
                
                // 只清除可见区域的高亮（提高性能）
                var visibleRange = app.ActiveWindow.VisibleRange;
                if (visibleRange != null)
                {
                    // 扩展一点范围确保清理完全
                    int startRow = Math.Max(1, visibleRange.Row - 10);
                    int endRow = Math.Min(visibleRange.Row + visibleRange.Rows.Count + 10, 200);
                    int startCol = Math.Max(1, visibleRange.Column - 10);
                    int endCol = Math.Min(visibleRange.Column + visibleRange.Columns.Count + 10, 100);
                    
                    var clearRange = worksheet.Range[worksheet.Cells[startRow, startCol], worksheet.Cells[endRow, endCol]];
                    clearRange.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"清除高亮失败: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 设置高亮颜色
        /// </summary>
        public static void SetHighlightColor(Color color)
        {
            _currentHighlightColor = color;
            
            // 如果聚光灯正在运行，立即应用新颜色
            if (_isActive)
            {
                try
                {
                    var app = Globals.ThisAddIn.Application;
                    var currentSelection = app.Selection as Excel.Range;
                    if (currentSelection != null)
                    {
                        ClearAllHighlights();
                        HighlightCrossHair(currentSelection);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"应用新颜色失败: {ex.Message}");
                }
            }
            
            System.Diagnostics.Debug.WriteLine($"聚光灯颜色已设置为: {color.Name}");
        }
        
        /// <summary>
        /// 设置显示模式
        /// </summary>
        public static void SetMode(SpotlightMode mode)
        {
            _currentMode = mode;
            
            // 如果聚光灯正在运行，立即应用新模式
            if (_isActive)
            {
                try
                {
                    var app = Globals.ThisAddIn.Application;
                    var currentSelection = app.Selection as Excel.Range;
                    if (currentSelection != null)
                    {
                        ClearAllHighlights();
                        HighlightCrossHair(currentSelection);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"应用新模式失败: {ex.Message}");
                }
            }
            
            System.Diagnostics.Debug.WriteLine($"聚光灯模式已设置为: {mode}");
        }
        
        /// <summary>
        /// 清理资源（在加载项关闭时调用）
        /// </summary>
        public static void Cleanup()
        {
            try
            {
                Stop();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"清理聚光灯资源失败: {ex.Message}");
            }
        }
    }
    
    /// <summary>
    /// 聚光灯显示模式
    /// </summary>
    public enum SpotlightMode
    {
        Both = 0,    // 全部(行+列)
        Row = 1,     // 行
        Column = 2   // 列
    }
}