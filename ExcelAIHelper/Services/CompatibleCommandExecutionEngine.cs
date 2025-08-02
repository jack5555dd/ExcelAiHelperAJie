using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using ExcelAIHelper.Exceptions;
using ExcelAIHelper.Models;
using Newtonsoft.Json.Linq;

namespace ExcelAIHelper.Services
{
    /// <summary>
    /// 兼容版命令执行引擎 - 支持Microsoft Excel和WPS表格
    /// </summary>
    public class CompatibleCommandExecutionEngine
    {
        private readonly CompatibleExcelOperationEngine _operationEngine;
        private readonly IExcelApplication _excelApp;
        
        /// <summary>
        /// 创建兼容版命令执行引擎实例
        /// </summary>
        public CompatibleCommandExecutionEngine()
        {
            _excelApp = ExcelApplicationFactory.GetCurrentApplication();
            _operationEngine = new CompatibleExcelOperationEngine();
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

                System.Diagnostics.Debug.WriteLine($"开始执行 {commands.Count} 个命令 (DryRun: {dryRun})");
                
                foreach (JObject command in commands)
                {
                    var commandResult = await ExecuteSingleCommandAsync(command, dryRun);
                    result.ExecutedCommands.Add(commandResult);
                    
                    if (!commandResult.Success)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"命令执行失败: {commandResult.ErrorMessage}";
                        break;
                    }
                }
                
                System.Diagnostics.Debug.WriteLine($"命令执行完成，成功: {result.Success}");
                return result;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"命令执行引擎异常: {ex.Message}");
                result.Success = false;
                result.ErrorMessage = $"执行过程中发生异常: {ex.Message}";
                return result;
            }
        }
        
        /// <summary>
        /// 执行单个命令
        /// </summary>
        /// <param name="command">命令对象</param>
        /// <param name="dryRun">是否为预览模式</param>
        /// <returns>命令执行结果</returns>
        private async Task<CommandResult> ExecuteSingleCommandAsync(JObject command, bool dryRun)
        {
            var commandResult = new CommandResult
            {
                Function = command["action"]?.ToString() ?? "unknown",
                Success = true
            };
            
            try
            {
                string action = command["action"]?.ToString()?.ToLower();
                System.Diagnostics.Debug.WriteLine($"执行命令: {action} (DryRun: {dryRun})");
                
                if (dryRun)
                {
                    commandResult.Description = $"预览: {GetCommandDescription(command)}";
                    return commandResult;
                }
                
                switch (action)
                {
                    case "set_value":
                        await ExecuteSetValueCommand(command);
                        commandResult.Description = "设置单元格值";
                        break;
                        
                    case "apply_formula":
                        await ExecuteApplyFormulaCommand(command);
                        commandResult.Description = "应用公式";
                        break;
                        
                    case "format_cells":
                        await ExecuteFormatCellsCommand(command);
                        commandResult.Description = "格式化单元格";
                        break;
                        
                    case "clear_content":
                        await ExecuteClearContentCommand(command);
                        commandResult.Description = "清除内容";
                        break;
                        
                    case "execute_vba":
                        await ExecuteVbaCommand(command);
                        commandResult.Description = "执行VBA代码";
                        break;
                        
                    default:
                        throw new AiOperationException($"不支持的命令类型: {action}");
                }
                
                System.Diagnostics.Debug.WriteLine($"命令 {action} 执行成功");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"命令执行失败: {ex.Message}");
                commandResult.Success = false;
                commandResult.ErrorMessage = ex.Message;
            }
            
            return commandResult;
        }
        
        /// <summary>
        /// 获取命令描述（用于预览模式）
        /// </summary>
        /// <param name="command">命令对象</param>
        /// <returns>命令描述</returns>
        private string GetCommandDescription(JObject command)
        {
            string action = command["action"]?.ToString()?.ToLower();
            string target = command["target"]?.ToString() ?? "当前选择";
            
            switch (action)
            {
                case "set_value":
                    string value = command["value"]?.ToString() ?? "";
                    return $"在 {target} 设置值: {value}";
                    
                case "apply_formula":
                    string formula = command["formula"]?.ToString() ?? "";
                    return $"在 {target} 应用公式: {formula}";
                    
                case "format_cells":
                    return $"格式化 {target}";
                    
                case "clear_content":
                    return $"清除 {target} 的内容";
                    
                case "execute_vba":
                    return $"执行VBA代码";
                    
                default:
                    return $"执行 {action} 操作";
            }
        }
        
        /// <summary>
        /// 执行设置值命令
        /// </summary>
        /// <param name="command">命令对象</param>
        private async Task ExecuteSetValueCommand(JObject command)
        {
            string target = command["target"]?.ToString();
            object value = command["value"];
            
            var range = await GetTargetRange(target);
            await _operationEngine.SetCellValueAsync(range, value);
        }
        
        /// <summary>
        /// 执行应用公式命令
        /// </summary>
        /// <param name="command">命令对象</param>
        private async Task ExecuteApplyFormulaCommand(JObject command)
        {
            string target = command["target"]?.ToString();
            string formula = command["formula"]?.ToString();
            
            if (string.IsNullOrEmpty(formula))
                throw new AiOperationException("公式不能为空");
            
            var range = await GetTargetRange(target);
            await _operationEngine.ApplyFormulaAsync(range, formula);
        }
        
        /// <summary>
        /// 执行格式化单元格命令
        /// </summary>
        /// <param name="command">命令对象</param>
        private async Task ExecuteFormatCellsCommand(JObject command)
        {
            string target = command["target"]?.ToString();
            var formatting = command["formatting"] as JObject;
            
            if (formatting == null)
                throw new AiOperationException("格式化参数不能为空");
            
            var range = await GetTargetRange(target);
            
            // 处理背景色
            string backgroundColor = formatting["background_color"]?.ToString();
            if (!string.IsNullOrEmpty(backgroundColor))
            {
                await _operationEngine.SetBackgroundColorAsync(range, backgroundColor);
            }
            
            // 其他格式化选项可以在这里添加
        }
        
        /// <summary>
        /// 执行清除内容命令
        /// </summary>
        /// <param name="command">命令对象</param>
        private async Task ExecuteClearContentCommand(JObject command)
        {
            string target = command["target"]?.ToString();
            var range = await GetTargetRange(target);
            await _operationEngine.ClearRangeContentAsync(range);
        }
        
        /// <summary>
        /// 执行VBA命令
        /// </summary>
        /// <param name="command">命令对象</param>
        private async Task ExecuteVbaCommand(JObject command)
        {
            string vbaCode = command["code"]?.ToString();
            
            if (string.IsNullOrEmpty(vbaCode))
                throw new AiOperationException("VBA代码不能为空");
            
            if (!_operationEngine.HasVbaAccess())
                throw new AiOperationException("没有VBA访问权限");
            
            await _operationEngine.ExecuteVbaCodeAsync(vbaCode);
        }
        
        /// <summary>
        /// 获取目标区域
        /// </summary>
        /// <param name="target">目标规范</param>
        /// <returns>目标区域对象</returns>
        private async Task<object> GetTargetRange(string target)
        {
            if (string.IsNullOrEmpty(target) || target.ToLower() == "selection")
            {
                return await _operationEngine.GetSelectedRangeAsync();
            }
            
            return _excelApp.GetRange(target);
        }
        
        /// <summary>
        /// 获取当前应用程序信息
        /// </summary>
        /// <returns>应用程序信息</returns>
        public string GetApplicationInfo()
        {
            return $"当前运行在: {_operationEngine.GetApplicationName()}";
        }
    }
}