using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using ExcelAIHelper.Exceptions;
using ExcelAIHelper.Models;
using Newtonsoft.Json.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAIHelper.Services
{
    /// <summary>
    /// 命令执行引擎 - 处理JSON协议命令的执行
    /// </summary>
    public class CommandExecutionEngine
    {
        private readonly ExcelOperationEngine _operationEngine;
        private readonly Excel.Application _excelApp;
        
        /// <summary>
        /// 创建命令执行引擎实例
        /// </summary>
        /// <param name="excelApp">Excel应用程序实例</param>
        public CommandExecutionEngine(Excel.Application excelApp)
        {
            _excelApp = excelApp ?? throw new ArgumentNullException(nameof(excelApp));
            _operationEngine = new ExcelOperationEngine(excelApp);
        }
        
        /// <summary>
        /// 执行JSON命令集
        /// </summary>
        /// <param name="validatedJson">已验证的JSON命令</param>
        /// <param name="dryRun">是否为预览模式</param>
        /// <returns>执行结果</returns>
        public async Task<CommandExecutionResult> ExecuteCommandsAsync(JObject validatedJson, bool dryRun = false)
        {
            var result = new CommandExecutionResult
            {
                Success = true,
                ExecutedCommands = new List<CommandResult>(),
                Summary = validatedJson["summary"]?.ToString() ?? ""
            };
            
            try
            {
                var commands = validatedJson["commands"] as JArray;
                if (commands == null || commands.Count == 0)
                {
                    result.Success = false;
                    result.ErrorMessage = "没有找到可执行的命令";
                    return result;
                }
                
                foreach (var command in commands)
                {
                    var commandResult = await ExecuteSingleCommandAsync(command as JObject, dryRun);
                    result.ExecutedCommands.Add(commandResult);
                    
                    if (!commandResult.Success)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"命令执行失败: {commandResult.ErrorMessage}";
                        break;
                    }
                }
                
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"命令执行异常: {ex.Message}";
                return result;
            }
        }
        
        /// <summary>
        /// 执行单个命令
        /// </summary>
        /// <param name="command">命令JSON对象</param>
        /// <param name="dryRun">是否为预览模式</param>
        /// <returns>命令执行结果</returns>
        private async Task<CommandResult> ExecuteSingleCommandAsync(JObject command, bool dryRun)
        {
            var result = new CommandResult
            {
                Function = command["function"]?.ToString(),
                Description = command["description"]?.ToString() ?? "",
                Success = true
            };
            
            try
            {
                if (dryRun)
                {
                    result.Message = $"[预览] 将执行: {result.Description}";
                    return result;
                }
                
                var function = command["function"]?.ToString();
                var arguments = command["arguments"] as JObject;
                
                switch (function)
                {
                    case "setCellValue":
                        await ExecuteSetCellValueAsync(arguments, result);
                        break;
                        
                    case "applyCellFormula":
                        await ExecuteApplyCellFormulaAsync(arguments, result);
                        break;
                        
                    case "setCellStyle":
                        await ExecuteSetCellStyleAsync(arguments, result);
                        break;
                        
                    case "setCellFormat":
                        await ExecuteSetCellFormatAsync(arguments, result);
                        break;
                        
                    case "clearCellContent":
                        await ExecuteClearCellContentAsync(arguments, result);
                        break;
                        
                    case "insertRows":
                        await ExecuteInsertRowsAsync(arguments, result);
                        break;
                        
                    case "insertColumns":
                        await ExecuteInsertColumnsAsync(arguments, result);
                        break;
                        
                    case "deleteRows":
                        await ExecuteDeleteRowsAsync(arguments, result);
                        break;
                        
                    case "deleteColumns":
                        await ExecuteDeleteColumnsAsync(arguments, result);
                        break;
                        
                    case "copyRange":
                        await ExecuteCopyRangeAsync(arguments, result);
                        break;
                        
                    case "findReplace":
                        await ExecuteFindReplaceAsync(arguments, result);
                        break;
                        
                    case "sortRange":
                        await ExecuteSortRangeAsync(arguments, result);
                        break;
                        
                    case "filterRange":
                        await ExecuteFilterRangeAsync(arguments, result);
                        break;
                        
                    default:
                        result.Success = false;
                        result.ErrorMessage = $"不支持的函数: {function}";
                        break;
                }
                
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return result;
            }
        }
        
        /// <summary>
        /// 执行设置单元格值命令
        /// </summary>
        private async Task ExecuteSetCellValueAsync(JObject arguments, CommandResult result)
        {
            var range = await GetTargetRangeAsync(arguments["range"]?.ToString());
            var value = arguments["value"];
            var dataType = arguments["dataType"]?.ToString() ?? "auto";
            
            // 根据数据类型转换值
            object convertedValue = ConvertValueByDataType(value, dataType);
            
            await _operationEngine.SetCellValueAsync(range, convertedValue);
            result.Message = $"✓ 在 {range.Address} 设置值: {convertedValue}";
        }
        
        /// <summary>
        /// 执行应用公式命令
        /// </summary>
        private async Task ExecuteApplyCellFormulaAsync(JObject arguments, CommandResult result)
        {
            var range = await GetTargetRangeAsync(arguments["range"]?.ToString());
            var formula = arguments["formula"]?.ToString();
            
            await _operationEngine.ApplyFormulaAsync(range, formula);
            result.Message = $"✓ 在 {range.Address} 应用公式: {formula}";
        }
        
        /// <summary>
        /// 执行设置单元格样式命令
        /// </summary>
        private async Task ExecuteSetCellStyleAsync(JObject arguments, CommandResult result)
        {
            var range = await GetTargetRangeAsync(arguments["range"]?.ToString());
            var styleChanges = new List<string>();
            
            // 背景颜色
            if (arguments["backgroundColor"] != null)
            {
                var bgColor = arguments["backgroundColor"].ToString();
                await _operationEngine.SetBackgroundColorAsync(range, bgColor);
                styleChanges.Add($"背景色: {bgColor}");
            }
            
            // 字体颜色
            if (arguments["fontColor"] != null)
            {
                var fontColor = arguments["fontColor"].ToString();
                await _operationEngine.SetFontColorAsync(range, fontColor);
                styleChanges.Add($"字体色: {fontColor}");
            }
            
            // 字体样式
            bool? bold = arguments["bold"]?.ToObject<bool?>();
            bool? italic = arguments["italic"]?.ToObject<bool?>();
            bool? underline = arguments["underline"]?.ToObject<bool?>();
            int? fontSize = arguments["fontSize"]?.ToObject<int?>();
            string fontName = arguments["fontName"]?.ToString();
            
            if (bold.HasValue || italic.HasValue || underline.HasValue || fontSize.HasValue || !string.IsNullOrEmpty(fontName))
            {
                await _operationEngine.SetFontStyleAsync(range, bold, italic, underline, fontSize, fontName);
                
                if (bold.HasValue) styleChanges.Add($"粗体: {bold.Value}");
                if (italic.HasValue) styleChanges.Add($"斜体: {italic.Value}");
                if (underline.HasValue) styleChanges.Add($"下划线: {underline.Value}");
                if (fontSize.HasValue) styleChanges.Add($"字号: {fontSize.Value}");
                if (!string.IsNullOrEmpty(fontName)) styleChanges.Add($"字体: {fontName}");
            }
            
            result.Message = $"✓ 在 {range.Address} 设置样式: {string.Join(", ", styleChanges)}";
        }
        
        /// <summary>
        /// 执行设置单元格格式命令
        /// </summary>
        private async Task ExecuteSetCellFormatAsync(JObject arguments, CommandResult result)
        {
            var range = await GetTargetRangeAsync(arguments["range"]?.ToString());
            var format = arguments["format"]?.ToString();
            
            await _operationEngine.SetNumberFormatAsync(range, format);
            result.Message = $"✓ 在 {range.Address} 设置格式: {format}";
        }
        
        /// <summary>
        /// 执行清除单元格内容命令
        /// </summary>
        private async Task ExecuteClearCellContentAsync(JObject arguments, CommandResult result)
        {
            var range = await GetTargetRangeAsync(arguments["range"]?.ToString());
            var clearType = arguments["clearType"]?.ToString() ?? "content";
            
            switch (clearType.ToLower())
            {
                case "content":
                    await _operationEngine.ClearRangeContentAsync(range);
                    break;
                case "format":
                    range.ClearFormats();
                    break;
                case "all":
                    range.Clear();
                    break;
                default:
                    await _operationEngine.ClearRangeContentAsync(range);
                    break;
            }
            
            result.Message = $"✓ 清除 {range.Address} 的{clearType}";
        }
        
        /// <summary>
        /// 执行插入行命令
        /// </summary>
        private async Task ExecuteInsertRowsAsync(JObject arguments, CommandResult result)
        {
            var position = arguments["position"]?.ToObject<int>() ?? 1;
            var count = arguments["count"]?.ToObject<int>() ?? 1;
            
            var range = _excelApp.ActiveSheet.Rows[position];
            for (int i = 0; i < count; i++)
            {
                range.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
            }
            
            result.Message = $"✓ 在第{position}行插入{count}行";
        }
        
        /// <summary>
        /// 执行插入列命令
        /// </summary>
        private async Task ExecuteInsertColumnsAsync(JObject arguments, CommandResult result)
        {
            var position = arguments["position"]?.ToString();
            var count = arguments["count"]?.ToObject<int>() ?? 1;
            
            var range = _excelApp.ActiveSheet.Columns[position];
            for (int i = 0; i < count; i++)
            {
                range.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
            }
            
            result.Message = $"✓ 在{position}列插入{count}列";
        }
        
        /// <summary>
        /// 执行删除行命令
        /// </summary>
        private async Task ExecuteDeleteRowsAsync(JObject arguments, CommandResult result)
        {
            var rangeSpec = arguments["range"]?.ToString();
            var range = _excelApp.ActiveSheet.Range[rangeSpec].EntireRow;
            
            range.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
            result.Message = $"✓ 删除行: {rangeSpec}";
        }
        
        /// <summary>
        /// 执行删除列命令
        /// </summary>
        private async Task ExecuteDeleteColumnsAsync(JObject arguments, CommandResult result)
        {
            var rangeSpec = arguments["range"]?.ToString();
            var range = _excelApp.ActiveSheet.Range[rangeSpec].EntireColumn;
            
            range.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);
            result.Message = $"✓ 删除列: {rangeSpec}";
        }
        
        /// <summary>
        /// 执行复制范围命令
        /// </summary>
        private async Task ExecuteCopyRangeAsync(JObject arguments, CommandResult result)
        {
            var sourceRange = await GetTargetRangeAsync(arguments["sourceRange"]?.ToString());
            var targetRange = await GetTargetRangeAsync(arguments["targetRange"]?.ToString());
            var copyType = ParseCopyType(arguments["copyType"]?.ToString());
            
            await _operationEngine.CopyRangeAsync(sourceRange, targetRange, copyType);
            result.Message = $"✓ 从 {sourceRange.Address} 复制到 {targetRange.Address} (类型: {copyType})";
        }
        
        /// <summary>
        /// 执行查找替换命令
        /// </summary>
        private async Task ExecuteFindReplaceAsync(JObject arguments, CommandResult result)
        {
            var range = await GetTargetRangeAsync(arguments["range"]?.ToString());
            var findWhat = arguments["findWhat"]?.ToString();
            var replaceWith = arguments["replaceWith"]?.ToString();
            var matchCase = arguments["matchCase"]?.ToObject<bool>() ?? false;
            var matchEntireCell = arguments["matchEntireCell"]?.ToObject<bool>() ?? false;
            
            var replaceCount = await _operationEngine.FindAndReplaceAsync(range, findWhat, replaceWith, matchCase, matchEntireCell);
            result.Message = $"✓ 在 {range.Address} 替换了 {replaceCount} 个 '{findWhat}' 为 '{replaceWith}'";
        }
        
        /// <summary>
        /// 执行排序命令
        /// </summary>
        private async Task ExecuteSortRangeAsync(JObject arguments, CommandResult result)
        {
            var range = await GetTargetRangeAsync(arguments["range"]?.ToString());
            var sortColumn = arguments["sortColumn"]?.ToObject<int>() ?? 1;
            var ascending = arguments["ascending"]?.ToObject<bool>() ?? true;
            var hasHeaders = arguments["hasHeaders"]?.ToObject<bool>() ?? true;
            
            await _operationEngine.SortRangeAsync(range, sortColumn, ascending, hasHeaders);
            var direction = ascending ? "升序" : "降序";
            result.Message = $"✓ 对 {range.Address} 按第{sortColumn}列{direction}排序";
        }
        
        /// <summary>
        /// 执行筛选命令
        /// </summary>
        private async Task ExecuteFilterRangeAsync(JObject arguments, CommandResult result)
        {
            var range = await GetTargetRangeAsync(arguments["range"]?.ToString());
            var column = arguments["column"]?.ToObject<int>() ?? 1;
            var criteria = arguments["criteria"]?.ToString();
            
            await _operationEngine.ApplyAutoFilterAsync(range, column, criteria);
            result.Message = $"✓ 对 {range.Address} 第{column}列应用筛选条件: {criteria}";
        }
        
        /// <summary>
        /// 获取目标范围
        /// </summary>
        private async Task<Excel.Range> GetTargetRangeAsync(string rangeSpec)
        {
            if (string.IsNullOrEmpty(rangeSpec) || rangeSpec == "CURRENT_SELECTION")
            {
                return _excelApp.Selection as Excel.Range ?? _excelApp.ActiveCell;
            }
            
            try
            {
                return _excelApp.ActiveSheet.Range[rangeSpec];
            }
            catch
            {
                // 如果解析失败，返回当前选区
                return _excelApp.Selection as Excel.Range ?? _excelApp.ActiveCell;
            }
        }
        
        /// <summary>
        /// 根据数据类型转换值
        /// </summary>
        private object ConvertValueByDataType(JToken value, string dataType)
        {
            if (value == null) return null;
            
            switch (dataType.ToLower())
            {
                case "number":
                    return value.ToObject<double>();
                case "text":
                    return value.ToString();
                case "boolean":
                    return value.ToObject<bool>();
                case "date":
                    return DateTime.Parse(value.ToString());
                case "auto":
                default:
                    // 自动检测类型
                    if (value.Type == JTokenType.Integer || value.Type == JTokenType.Float)
                        return value.ToObject<double>();
                    if (value.Type == JTokenType.Boolean)
                        return value.ToObject<bool>();
                    if (DateTime.TryParse(value.ToString(), out DateTime dateValue))
                        return dateValue;
                    return value.ToString();
            }
        }
        
        /// <summary>
        /// 解析复制类型
        /// </summary>
        private CopyType ParseCopyType(string copyTypeStr)
        {
            if (string.IsNullOrEmpty(copyTypeStr))
                return CopyType.All;
                
            switch (copyTypeStr.ToLower())
            {
                case "values":
                    return CopyType.Values;
                case "formulas":
                    return CopyType.Formulas;
                case "formats":
                    return CopyType.Formats;
                case "all":
                default:
                    return CopyType.All;
            }
        }
    }
    
    /// <summary>
    /// 命令执行结果
    /// </summary>
    public class CommandExecutionResult
    {
        /// <summary>
        /// 是否执行成功
        /// </summary>
        public bool Success { get; set; }
        
        /// <summary>
        /// 错误消息
        /// </summary>
        public string ErrorMessage { get; set; }
        
        /// <summary>
        /// 操作摘要
        /// </summary>
        public string Summary { get; set; }
        
        /// <summary>
        /// 已执行的命令列表
        /// </summary>
        public List<CommandResult> ExecutedCommands { get; set; }
        
        /// <summary>
        /// 获取执行结果摘要
        /// </summary>
        public string GetResultSummary()
        {
            if (!Success)
                return $"❌ 执行失败: {ErrorMessage}";
                
            var successCount = ExecutedCommands?.Count ?? 0;
            var messages = ExecutedCommands?.ConvertAll(c => c.Message) ?? new List<string>();
            
            return $"✅ 成功执行 {successCount} 个命令:\n" + string.Join("\n", messages);
        }
    }
    
    /// <summary>
    /// 单个命令执行结果
    /// </summary>
    public class CommandResult
    {
        /// <summary>
        /// 函数名
        /// </summary>
        public string Function { get; set; }
        
        /// <summary>
        /// 描述
        /// </summary>
        public string Description { get; set; }
        
        /// <summary>
        /// 是否成功
        /// </summary>
        public bool Success { get; set; }
        
        /// <summary>
        /// 执行消息
        /// </summary>
        public string Message { get; set; }
        
        /// <summary>
        /// 错误消息
        /// </summary>
        public string ErrorMessage { get; set; }
    }
}