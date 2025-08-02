using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using ExcelAIHelper.Exceptions;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;

namespace ExcelAIHelper.Services
{
    /// <summary>
    /// VBA生成结果
    /// </summary>
    public class VbaGenerationResult
    {
        /// <summary>
        /// 是否成功
        /// </summary>
        public bool Success { get; set; }
        
        /// <summary>
        /// 宏名称
        /// </summary>
        public string MacroName { get; set; }
        
        /// <summary>
        /// VBA代码
        /// </summary>
        public string VbaCode { get; set; }
        
        /// <summary>
        /// 功能描述
        /// </summary>
        public string Description { get; set; }
        
        /// <summary>
        /// 风险级别
        /// </summary>
        public string RiskLevel { get; set; }
        
        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }
        
        /// <summary>
        /// 安全扫描结果
        /// </summary>
        public SecurityScanResult SecurityScanResult { get; set; }

        public static VbaGenerationResult CreateSuccess(string macroName, string vbaCode, string description, string riskLevel)
        {
            return new VbaGenerationResult
            {
                Success = true,
                MacroName = macroName,
                VbaCode = vbaCode,
                Description = description,
                RiskLevel = riskLevel
            };
        }

        public static VbaGenerationResult CreateFailure(string errorMessage)
        {
            return new VbaGenerationResult
            {
                Success = false,
                ErrorMessage = errorMessage
            };
        }
    }

    /// <summary>
    /// VBA执行结果
    /// </summary>
    public class VbaExecutionResult
    {
        /// <summary>
        /// 是否成功
        /// </summary>
        public bool Success { get; set; }
        
        /// <summary>
        /// 执行结果
        /// </summary>
        public object Result { get; set; }
        
        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }
        
        /// <summary>
        /// 执行时间（毫秒）
        /// </summary>
        public long ExecutionTimeMs { get; set; }
        
        /// <summary>
        /// 临时模块名称
        /// </summary>
        public string TempModuleName { get; set; }

        public static VbaExecutionResult CreateSuccess(object result, long executionTimeMs, string tempModuleName = "")
        {
            return new VbaExecutionResult
            {
                Success = true,
                Result = result,
                ExecutionTimeMs = executionTimeMs,
                TempModuleName = tempModuleName
            };
        }

        public static VbaExecutionResult CreateFailure(string errorMessage, long executionTimeMs = 0)
        {
            return new VbaExecutionResult
            {
                Success = false,
                ErrorMessage = errorMessage,
                ExecutionTimeMs = executionTimeMs
            };
        }
    }

    /// <summary>
    /// VBA执行日志
    /// </summary>
    public class VbaExecutionLog
    {
        /// <summary>
        /// 执行时间
        /// </summary>
        public DateTime ExecutionTime { get; set; }
        
        /// <summary>
        /// 宏名称
        /// </summary>
        public string MacroName { get; set; }
        
        /// <summary>
        /// VBA代码
        /// </summary>
        public string VbaCode { get; set; }
        
        /// <summary>
        /// 是否成功
        /// </summary>
        public bool Success { get; set; }
        
        /// <summary>
        /// 执行时间（毫秒）
        /// </summary>
        public long ExecutionTimeMs { get; set; }
        
        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }
        
        /// <summary>
        /// 用户请求
        /// </summary>
        public string UserRequest { get; set; }
    }

    /// <summary>
    /// VBA注入引擎
    /// 负责生成、注入和执行VBA代码
    /// </summary>
    public class VbaInjectionEngine
    {
        private readonly DeepSeekClient _aiClient;
        private readonly VbaSecurityScanner _securityScanner;
        private readonly List<VbaExecutionLog> _executionHistory;
        private readonly object _lockObject = new object();

        public VbaInjectionEngine(DeepSeekClient aiClient)
        {
            _aiClient = aiClient ?? throw new ArgumentNullException(nameof(aiClient));
            _securityScanner = new VbaSecurityScanner();
            _executionHistory = new List<VbaExecutionLog>();
        }

        /// <summary>
        /// 检查VBE访问权限
        /// </summary>
        /// <returns>是否有访问权限</returns>
        public bool CheckVbeAccess()
        {
            try
            {
                var excelApp = Globals.ThisAddIn?.Application;
                if (excelApp == null)
                {
                    Debug.WriteLine("[VbaInjectionEngine] Excel应用程序不可用");
                    return false;
                }

                Debug.WriteLine("[VbaInjectionEngine] 开始VBE访问权限检查...");
                
                // 尝试访问VBE对象
                var vbe = excelApp.VBE;
                Debug.WriteLine("[VbaInjectionEngine] VBE对象获取成功");
                
                var vbProject = vbe.ActiveVBProject;
                if (vbProject == null)
                {
                    Debug.WriteLine("[VbaInjectionEngine] ActiveVBProject为null");
                    return false;
                }
                
                Debug.WriteLine($"[VbaInjectionEngine] VBA项目访问成功: {vbProject.Name}");
                
                // 尝试访问VBComponents来进一步验证权限
                var components = vbProject.VBComponents;
                Debug.WriteLine($"[VbaInjectionEngine] VBComponents访问成功，组件数量: {components.Count}");
                
                Debug.WriteLine("[VbaInjectionEngine] VBE访问权限检查通过");
                return true;
            }
            catch (System.UnauthorizedAccessException ex)
            {
                Debug.WriteLine($"[VbaInjectionEngine] VBE访问被拒绝: {ex.Message}");
                Debug.WriteLine("[VbaInjectionEngine] 请确保已启用'信任对VBA项目对象模型的访问'");
                return false;
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                Debug.WriteLine($"[VbaInjectionEngine] COM异常: {ex.Message} (HRESULT: 0x{ex.HResult:X8})");
                if (ex.HResult == unchecked((int)0x800A03EC)) // VBA项目被保护
                {
                    Debug.WriteLine("[VbaInjectionEngine] VBA项目被密码保护或锁定");
                }
                else if (ex.HResult == unchecked((int)0x800A9C68)) // 无法访问VBA项目
                {
                    Debug.WriteLine("[VbaInjectionEngine] 无法访问VBA项目，可能需要启用宏或VBE访问权限");
                }
                return false;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[VbaInjectionEngine] VBE访问权限检查失败: {ex.Message}");
                Debug.WriteLine($"[VbaInjectionEngine] 异常类型: {ex.GetType().Name}");
                return false;
            }
        }

        /// <summary>
        /// 生成VBA代码
        /// </summary>
        /// <param name="userRequest">用户请求</param>
        /// <returns>VBA生成结果</returns>
        public async Task<VbaGenerationResult> GenerateVbaCodeAsync(string userRequest)
        {
            try
            {
                Debug.WriteLine($"[VbaInjectionEngine] 开始生成VBA代码，用户请求: {userRequest}");

                // 构建VBA专用提示
                string systemPrompt = BuildVbaSystemPrompt();
                string userPrompt = BuildVbaUserPrompt(userRequest);

                // 调用AI生成VBA代码
                string aiResponse = await _aiClient.AskAsync(userPrompt, systemPrompt);
                
                Debug.WriteLine($"[VbaInjectionEngine] AI响应: {aiResponse}");

                // 解析AI响应
                var generationResult = ParseVbaResponse(aiResponse);
                if (!generationResult.Success)
                {
                    return generationResult;
                }

                // 安全扫描
                var scanResult = _securityScanner.ScanCode(generationResult.VbaCode);
                generationResult.SecurityScanResult = scanResult;

                if (!scanResult.IsSafe)
                {
                    Debug.WriteLine($"[VbaInjectionEngine] 安全扫描未通过: {scanResult.Summary}");
                    return VbaGenerationResult.CreateFailure($"安全扫描未通过: {scanResult.Summary}");
                }

                Debug.WriteLine("[VbaInjectionEngine] VBA代码生成成功");
                return generationResult;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[VbaInjectionEngine] VBA代码生成失败: {ex.Message}");
                return VbaGenerationResult.CreateFailure($"VBA代码生成失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 注入并执行VBA代码
        /// </summary>
        /// <param name="macroName">宏名称</param>
        /// <param name="vbaCode">VBA代码</param>
        /// <param name="userRequest">用户原始请求（用于日志）</param>
        /// <returns>执行结果</returns>
        public async Task<VbaExecutionResult> InjectAndExecuteAsync(string macroName, string vbaCode, string userRequest = "")
        {
            var stopwatch = Stopwatch.StartNew();
            string tempModuleName = "";
            
            try
            {
                Debug.WriteLine($"[VbaInjectionEngine] 开始注入并执行VBA代码: {macroName}");

                // 检查VBE访问权限
                if (!CheckVbeAccess())
                {
                    return VbaExecutionResult.CreateFailure("VBE访问权限不足，请在Excel选项中启用'信任对VBA项目对象模型的访问'");
                }

                var excelApp = Globals.ThisAddIn.Application;
                var vbe = excelApp.VBE;
                var vbProject = vbe.ActiveVBProject;

                // 生成临时模块名称
                tempModuleName = $"TempModule_{DateTime.Now:yyyyMMddHHmmss}_{new Random().Next(1000, 9999)}";

                // 创建临时模块
                var tempModule = vbProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
                tempModule.Name = tempModuleName;

                Debug.WriteLine($"[VbaInjectionEngine] 创建临时模块: {tempModuleName}");

                // 注入VBA代码
                tempModule.CodeModule.AddFromString(vbaCode);
                
                Debug.WriteLine("[VbaInjectionEngine] VBA代码注入完成");

                // 执行宏
                object result = null;
                try
                {
                    result = excelApp.Run(macroName);
                    Debug.WriteLine("[VbaInjectionEngine] VBA宏执行成功");
                }
                catch (Exception execEx)
                {
                    Debug.WriteLine($"[VbaInjectionEngine] VBA宏执行失败: {execEx.Message}");
                    throw new AiOperationException($"VBA宏执行失败: {execEx.Message}", execEx);
                }
                finally
                {
                    // 清理临时模块
                    try
                    {
                        vbProject.VBComponents.Remove(tempModule);
                        Debug.WriteLine($"[VbaInjectionEngine] 临时模块已清理: {tempModuleName}");
                    }
                    catch (Exception cleanupEx)
                    {
                        Debug.WriteLine($"[VbaInjectionEngine] 清理临时模块失败: {cleanupEx.Message}");
                    }
                }

                stopwatch.Stop();
                
                // 记录执行日志
                LogExecution(macroName, vbaCode, userRequest, true, stopwatch.ElapsedMilliseconds, "");

                return VbaExecutionResult.CreateSuccess(result, stopwatch.ElapsedMilliseconds, tempModuleName);
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                
                Debug.WriteLine($"[VbaInjectionEngine] VBA注入执行失败: {ex.Message}");
                
                // 记录执行日志
                LogExecution(macroName, vbaCode, userRequest, false, stopwatch.ElapsedMilliseconds, ex.Message);
                
                return VbaExecutionResult.CreateFailure(ex.Message, stopwatch.ElapsedMilliseconds);
            }
        }

        /// <summary>
        /// 清理临时模块
        /// </summary>
        public void CleanupTempModules()
        {
            try
            {
                var excelApp = Globals.ThisAddIn?.Application;
                if (excelApp == null) return;

                var vbe = excelApp.VBE;
                var vbProject = vbe.ActiveVBProject;

                // 查找并删除临时模块
                var componentsToRemove = new List<VBComponent>();
                
                foreach (VBComponent component in vbProject.VBComponents)
                {
                    if (component.Name.StartsWith("TempModule_"))
                    {
                        componentsToRemove.Add(component);
                    }
                }

                foreach (var component in componentsToRemove)
                {
                    try
                    {
                        vbProject.VBComponents.Remove(component);
                        Debug.WriteLine($"[VbaInjectionEngine] 清理临时模块: {component.Name}");
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[VbaInjectionEngine] 清理模块失败 {component.Name}: {ex.Message}");
                    }
                }

                Debug.WriteLine($"[VbaInjectionEngine] 临时模块清理完成，共清理 {componentsToRemove.Count} 个模块");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[VbaInjectionEngine] 清理临时模块异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取执行历史
        /// </summary>
        /// <returns>执行历史列表</returns>
        public List<VbaExecutionLog> GetExecutionHistory()
        {
            lock (_lockObject)
            {
                return new List<VbaExecutionLog>(_executionHistory);
            }
        }

        /// <summary>
        /// 清除执行历史
        /// </summary>
        public void ClearExecutionHistory()
        {
            lock (_lockObject)
            {
                _executionHistory.Clear();
                Debug.WriteLine("[VbaInjectionEngine] 执行历史已清除");
            }
        }

        /// <summary>
        /// 构建VBA系统提示
        /// </summary>
        private string BuildVbaSystemPrompt()
        {
            return @"你是一名资深Excel VBA开发者。请根据用户需求生成VBA代码，严格按照以下JSON格式返回：

{
  ""macroName"": ""生成的宏名称"",
  ""vbaCode"": ""完整的VBA代码"",
  ""description"": ""功能描述"",
  ""riskLevel"": ""low""
}

安全要求：
1. 禁止使用：Shell, Kill, CreateObject(""WScript.Shell""), FileSystemObject, Dir, ChDir, MkDir, RmDir
2. 只能操作当前工作簿和工作表
3. 使用Application、Workbook、Worksheet、Range等Excel对象
4. 代码必须包含错误处理
5. 宏名称要简洁明了，使用英文

代码模板：
Sub MacroName()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' 你的操作代码
    
    Exit Sub
ErrorHandler:
    MsgBox ""操作失败: "" & Err.Description, vbCritical, ""错误""
End Sub

示例操作：
- 设置单元格值：Range(""A1"").Value = ""Hello""
- 应用公式：Range(""B1"").Formula = ""=SUM(A1:A10)""
- 设置格式：Range(""A1"").Font.Bold = True
- 循环操作：For i = 1 To 10 ... Next i

请确保生成的代码安全、高效、易读。";
        }

        /// <summary>
        /// 构建VBA用户提示
        /// </summary>
        private string BuildVbaUserPrompt(string userRequest)
        {
            return $@"用户需求：{userRequest}

请生成相应的VBA代码来实现这个需求。注意：
1. 代码要安全可靠
2. 包含适当的错误处理
3. 使用Excel内置对象和方法
4. 代码要简洁高效

请严格按照JSON格式返回结果。";
        }

        /// <summary>
        /// 解析VBA响应
        /// </summary>
        private VbaGenerationResult ParseVbaResponse(string aiResponse)
        {
            try
            {
                // 清理响应中的markdown包装
                string cleanedResponse = CleanJsonResponse(aiResponse);
                
                // 解析JSON
                var jsonObject = JObject.Parse(cleanedResponse);
                
                string macroName = jsonObject["macroName"]?.ToString();
                string vbaCode = jsonObject["vbaCode"]?.ToString();
                string description = jsonObject["description"]?.ToString() ?? "";
                string riskLevel = jsonObject["riskLevel"]?.ToString() ?? "medium";
                
                if (string.IsNullOrEmpty(macroName) || string.IsNullOrEmpty(vbaCode))
                {
                    return VbaGenerationResult.CreateFailure("AI响应格式不正确：缺少必要字段");
                }
                
                // 验证宏名称格式
                if (!IsValidMacroName(macroName))
                {
                    return VbaGenerationResult.CreateFailure($"宏名称格式不正确: {macroName}");
                }
                
                return VbaGenerationResult.CreateSuccess(macroName, vbaCode, description, riskLevel);
            }
            catch (JsonException ex)
            {
                Debug.WriteLine($"[VbaInjectionEngine] JSON解析失败: {ex.Message}");
                return VbaGenerationResult.CreateFailure($"AI响应解析失败: {ex.Message}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[VbaInjectionEngine] 响应解析异常: {ex.Message}");
                return VbaGenerationResult.CreateFailure($"响应解析异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 清理JSON响应
        /// </summary>
        private string CleanJsonResponse(string response)
        {
            if (string.IsNullOrWhiteSpace(response))
                return response;

            string cleaned = response.Trim();
            
            // 检测并移除markdown代码块
            if (cleaned.StartsWith("```json", StringComparison.OrdinalIgnoreCase))
            {
                int startIndex = cleaned.IndexOf('\n') + 1;
                int endIndex = cleaned.LastIndexOf("```");
                if (endIndex > startIndex)
                {
                    cleaned = cleaned.Substring(startIndex, endIndex - startIndex).Trim();
                }
            }
            else if (cleaned.StartsWith("```"))
            {
                int startIndex = cleaned.IndexOf('\n') + 1;
                int endIndex = cleaned.LastIndexOf("```");
                if (endIndex > startIndex)
                {
                    cleaned = cleaned.Substring(startIndex, endIndex - startIndex).Trim();
                }
            }
            
            return cleaned;
        }

        /// <summary>
        /// 验证宏名称格式
        /// </summary>
        private bool IsValidMacroName(string macroName)
        {
            if (string.IsNullOrWhiteSpace(macroName))
                return false;
            
            // VBA宏名称规则：以字母开头，只能包含字母、数字和下划线，长度不超过255
            return Regex.IsMatch(macroName, @"^[a-zA-Z][a-zA-Z0-9_]{0,254}$");
        }

        /// <summary>
        /// 记录执行日志
        /// </summary>
        private void LogExecution(string macroName, string vbaCode, string userRequest, bool success, long executionTimeMs, string errorMessage)
        {
            lock (_lockObject)
            {
                var log = new VbaExecutionLog
                {
                    ExecutionTime = DateTime.Now,
                    MacroName = macroName,
                    VbaCode = vbaCode,
                    UserRequest = userRequest,
                    Success = success,
                    ExecutionTimeMs = executionTimeMs,
                    ErrorMessage = errorMessage
                };
                
                _executionHistory.Add(log);
                
                // 保持最近100条记录
                if (_executionHistory.Count > 100)
                {
                    _executionHistory.RemoveAt(0);
                }
                
                Debug.WriteLine($"[VbaInjectionEngine] 执行日志已记录: {macroName} ({(success ? "成功" : "失败")})");
            }
        }
    }
}