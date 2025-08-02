using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Diagnostics;

namespace ExcelAIHelper.Services
{
    /// <summary>
    /// 安全级别枚举
    /// </summary>
    public enum SecurityLevel
    {
        /// <summary>
        /// 安全 - 无风险
        /// </summary>
        Safe,
        
        /// <summary>
        /// 低风险 - 可能有轻微风险但通常安全
        /// </summary>
        Low,
        
        /// <summary>
        /// 中等风险 - 需要用户确认
        /// </summary>
        Medium,
        
        /// <summary>
        /// 高风险 - 建议拒绝执行
        /// </summary>
        High,
        
        /// <summary>
        /// 危险 - 禁止执行
        /// </summary>
        Dangerous
    }

    /// <summary>
    /// 安全问题类型
    /// </summary>
    public enum SecurityIssueType
    {
        /// <summary>
        /// 系统调用
        /// </summary>
        SystemCall,
        
        /// <summary>
        /// 文件系统操作
        /// </summary>
        FileSystemAccess,
        
        /// <summary>
        /// 网络访问
        /// </summary>
        NetworkAccess,
        
        /// <summary>
        /// 注册表操作
        /// </summary>
        RegistryAccess,
        
        /// <summary>
        /// 外部程序执行
        /// </summary>
        ExternalExecution,
        
        /// <summary>
        /// 危险函数调用
        /// </summary>
        DangerousFunction,
        
        /// <summary>
        /// 可疑模式
        /// </summary>
        SuspiciousPattern
    }

    /// <summary>
    /// 安全问题详情
    /// </summary>
    public class SecurityIssue
    {
        /// <summary>
        /// 问题类型
        /// </summary>
        public SecurityIssueType Type { get; set; }
        
        /// <summary>
        /// 安全级别
        /// </summary>
        public SecurityLevel Level { get; set; }
        
        /// <summary>
        /// 问题描述
        /// </summary>
        public string Description { get; set; }
        
        /// <summary>
        /// 发现的代码片段
        /// </summary>
        public string CodeSnippet { get; set; }
        
        /// <summary>
        /// 行号
        /// </summary>
        public int LineNumber { get; set; }
        
        /// <summary>
        /// 建议
        /// </summary>
        public string Suggestion { get; set; }

        public SecurityIssue(SecurityIssueType type, SecurityLevel level, string description, string codeSnippet, int lineNumber, string suggestion = "")
        {
            Type = type;
            Level = level;
            Description = description;
            CodeSnippet = codeSnippet;
            LineNumber = lineNumber;
            Suggestion = suggestion;
        }
    }

    /// <summary>
    /// 安全扫描结果
    /// </summary>
    public class SecurityScanResult
    {
        /// <summary>
        /// 是否安全
        /// </summary>
        public bool IsSafe { get; set; }
        
        /// <summary>
        /// 安全问题列表
        /// </summary>
        public List<SecurityIssue> Issues { get; set; }
        
        /// <summary>
        /// 整体安全级别
        /// </summary>
        public SecurityLevel Level { get; set; }
        
        /// <summary>
        /// 扫描摘要
        /// </summary>
        public string Summary { get; set; }

        public SecurityScanResult()
        {
            Issues = new List<SecurityIssue>();
            IsSafe = true;
            Level = SecurityLevel.Safe;
        }
    }

    /// <summary>
    /// VBA安全扫描器
    /// 负责扫描VBA代码中的潜在安全风险
    /// </summary>
    public class VbaSecurityScanner
    {
        /// <summary>
        /// 危险关键字列表 - 禁止使用
        /// </summary>
        private static readonly Dictionary<string, SecurityIssueType> DangerousKeywords = new Dictionary<string, SecurityIssueType>(StringComparer.OrdinalIgnoreCase)
        {
            // 系统调用
            { "Shell", SecurityIssueType.SystemCall },
            { "Environ", SecurityIssueType.SystemCall },
            { "Command", SecurityIssueType.SystemCall },
            
            // 文件系统操作
            { "Kill", SecurityIssueType.FileSystemAccess },
            { "Dir", SecurityIssueType.FileSystemAccess },
            { "ChDir", SecurityIssueType.FileSystemAccess },
            { "MkDir", SecurityIssueType.FileSystemAccess },
            { "RmDir", SecurityIssueType.FileSystemAccess },
            { "FileCopy", SecurityIssueType.FileSystemAccess },
            { "FileSystemObject", SecurityIssueType.FileSystemAccess },
            
            // 外部对象创建
            { "CreateObject", SecurityIssueType.ExternalExecution },
            { "GetObject", SecurityIssueType.ExternalExecution },
            
            // 网络和脚本
            { "WScript", SecurityIssueType.ExternalExecution },
            { "WshShell", SecurityIssueType.ExternalExecution },
            { "InternetExplorer", SecurityIssueType.NetworkAccess },
            { "XMLHTTP", SecurityIssueType.NetworkAccess },
            { "WinHttp", SecurityIssueType.NetworkAccess }
        };

        /// <summary>
        /// 可疑模式列表 - 需要警告
        /// </summary>
        private static readonly Dictionary<string, SecurityIssueType> SuspiciousPatterns = new Dictionary<string, SecurityIssueType>(StringComparer.OrdinalIgnoreCase)
        {
            { "Registry", SecurityIssueType.RegistryAccess },
            { "SendKeys", SecurityIssueType.SuspiciousPattern },
            { "DoEvents", SecurityIssueType.SuspiciousPattern },
            { "Application.Run", SecurityIssueType.SuspiciousPattern },
            { "Eval", SecurityIssueType.DangerousFunction },
            { "Execute", SecurityIssueType.DangerousFunction }
        };

        /// <summary>
        /// 允许的安全函数白名单
        /// </summary>
        private static readonly HashSet<string> SafeFunctions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Application", "ActiveSheet", "ActiveWorkbook", "Workbooks", "Worksheets",
            "Range", "Cells", "Rows", "Columns", "Selection",
            "WorksheetFunction", "MsgBox", "InputBox",
            "Format", "CStr", "CInt", "CDbl", "CDate",
            "Left", "Right", "Mid", "Len", "InStr", "Replace",
            "UCase", "LCase", "Trim", "Split", "Join",
            "IsEmpty", "IsNull", "IsNumeric", "IsDate",
            "Now", "Date", "Time", "DateAdd", "DateDiff",
            "Abs", "Int", "Round", "Sqr", "Log", "Exp"
        };

        /// <summary>
        /// 扫描VBA代码
        /// </summary>
        /// <param name="vbaCode">要扫描的VBA代码</param>
        /// <returns>扫描结果</returns>
        public SecurityScanResult ScanCode(string vbaCode)
        {
            var result = new SecurityScanResult();
            
            try
            {
                if (string.IsNullOrWhiteSpace(vbaCode))
                {
                    result.Summary = "代码为空";
                    return result;
                }

                Debug.WriteLine("[VbaSecurityScanner] 开始安全扫描");

                // 按行扫描代码
                string[] lines = vbaCode.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                
                for (int i = 0; i < lines.Length; i++)
                {
                    string line = lines[i].Trim();
                    int lineNumber = i + 1;
                    
                    // 跳过注释行
                    if (line.StartsWith("'") || line.StartsWith("Rem ", StringComparison.OrdinalIgnoreCase))
                        continue;
                    
                    // 扫描危险关键字
                    ScanDangerousKeywords(line, lineNumber, result);
                    
                    // 扫描可疑模式
                    ScanSuspiciousPatterns(line, lineNumber, result);
                    
                    // 扫描特殊模式
                    ScanSpecialPatterns(line, lineNumber, result);
                }

                // 计算整体安全级别
                CalculateOverallSecurityLevel(result);
                
                // 生成扫描摘要
                GenerateScanSummary(result);
                
                Debug.WriteLine($"[VbaSecurityScanner] 扫描完成，发现 {result.Issues.Count} 个问题，安全级别: {result.Level}");
                
                return result;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[VbaSecurityScanner] 扫描异常: {ex.Message}");
                result.IsSafe = false;
                result.Level = SecurityLevel.High;
                result.Summary = $"扫描过程中发生错误: {ex.Message}";
                return result;
            }
        }

        /// <summary>
        /// 检查代码是否安全
        /// </summary>
        /// <param name="vbaCode">VBA代码</param>
        /// <returns>是否安全</returns>
        public bool IsCodeSafe(string vbaCode)
        {
            var result = ScanCode(vbaCode);
            return result.IsSafe && result.Level <= SecurityLevel.Low;
        }

        /// <summary>
        /// 获取安全问题列表
        /// </summary>
        /// <param name="vbaCode">VBA代码</param>
        /// <returns>安全问题列表</returns>
        public List<SecurityIssue> GetSecurityIssues(string vbaCode)
        {
            var result = ScanCode(vbaCode);
            return result.Issues;
        }

        /// <summary>
        /// 扫描危险关键字
        /// </summary>
        private void ScanDangerousKeywords(string line, int lineNumber, SecurityScanResult result)
        {
            foreach (var keyword in DangerousKeywords)
            {
                if (ContainsKeyword(line, keyword.Key))
                {
                    var issue = new SecurityIssue(
                        keyword.Value,
                        SecurityLevel.Dangerous,
                        $"检测到危险函数调用: {keyword.Key}",
                        line,
                        lineNumber,
                        "此函数可能存在安全风险，建议使用Excel内置对象替代"
                    );
                    
                    result.Issues.Add(issue);
                    result.IsSafe = false;
                    
                    Debug.WriteLine($"[VbaSecurityScanner] 发现危险关键字: {keyword.Key} (行 {lineNumber})");
                }
            }
        }

        /// <summary>
        /// 扫描可疑模式
        /// </summary>
        private void ScanSuspiciousPatterns(string line, int lineNumber, SecurityScanResult result)
        {
            foreach (var pattern in SuspiciousPatterns)
            {
                if (ContainsKeyword(line, pattern.Key))
                {
                    var issue = new SecurityIssue(
                        pattern.Value,
                        SecurityLevel.Medium,
                        $"检测到可疑函数调用: {pattern.Key}",
                        line,
                        lineNumber,
                        "此函数可能需要额外权限或存在潜在风险"
                    );
                    
                    result.Issues.Add(issue);
                    
                    Debug.WriteLine($"[VbaSecurityScanner] 发现可疑模式: {pattern.Key} (行 {lineNumber})");
                }
            }
        }

        /// <summary>
        /// 扫描特殊模式
        /// </summary>
        private void ScanSpecialPatterns(string line, int lineNumber, SecurityScanResult result)
        {
            // 检查是否包含可疑的字符串模式
            var suspiciousStringPatterns = new[]
            {
                @"\.exe""",
                @"\.bat""",
                @"\.cmd""",
                @"\.vbs""",
                @"\.ps1""",
                @"powershell",
                @"cmd\.exe"
            };

            foreach (var pattern in suspiciousStringPatterns)
            {
                if (Regex.IsMatch(line, pattern, RegexOptions.IgnoreCase))
                {
                    var issue = new SecurityIssue(
                        SecurityIssueType.ExternalExecution,
                        SecurityLevel.High,
                        "检测到可疑的可执行文件引用",
                        line,
                        lineNumber,
                        "避免在VBA中直接引用可执行文件"
                    );
                    
                    result.Issues.Add(issue);
                    
                    Debug.WriteLine($"[VbaSecurityScanner] 发现可疑字符串模式: {pattern} (行 {lineNumber})");
                }
            }

            // 检查网络URL模式
            if (Regex.IsMatch(line, @"https?://", RegexOptions.IgnoreCase))
            {
                var issue = new SecurityIssue(
                    SecurityIssueType.NetworkAccess,
                    SecurityLevel.Medium,
                    "检测到网络URL引用",
                    line,
                    lineNumber,
                    "网络访问可能存在安全风险"
                );
                
                result.Issues.Add(issue);
                
                Debug.WriteLine($"[VbaSecurityScanner] 发现网络URL (行 {lineNumber})");
            }
        }

        /// <summary>
        /// 检查行中是否包含关键字
        /// </summary>
        private bool ContainsKeyword(string line, string keyword)
        {
            // 使用正则表达式确保匹配完整的单词
            string pattern = @"\b" + Regex.Escape(keyword) + @"\b";
            return Regex.IsMatch(line, pattern, RegexOptions.IgnoreCase);
        }

        /// <summary>
        /// 计算整体安全级别
        /// </summary>
        private void CalculateOverallSecurityLevel(SecurityScanResult result)
        {
            if (result.Issues.Count == 0)
            {
                result.Level = SecurityLevel.Safe;
                result.IsSafe = true;
                return;
            }

            // 找到最高的安全级别
            SecurityLevel maxLevel = result.Issues.Max(issue => issue.Level);
            result.Level = maxLevel;
            
            // 如果有危险级别的问题，标记为不安全
            if (maxLevel >= SecurityLevel.High)
            {
                result.IsSafe = false;
            }
            else if (maxLevel == SecurityLevel.Medium && result.Issues.Count > 3)
            {
                // 如果中等风险问题太多，也标记为不安全
                result.IsSafe = false;
            }
            else
            {
                result.IsSafe = true;
            }
        }

        /// <summary>
        /// 生成扫描摘要
        /// </summary>
        private void GenerateScanSummary(SecurityScanResult result)
        {
            if (result.Issues.Count == 0)
            {
                result.Summary = "代码安全扫描通过，未发现安全问题";
                return;
            }

            var summary = $"发现 {result.Issues.Count} 个安全问题：";
            
            var groupedIssues = result.Issues.GroupBy(i => i.Level);
            foreach (var group in groupedIssues.OrderByDescending(g => g.Key))
            {
                summary += $"\n- {GetSecurityLevelDescription(group.Key)}: {group.Count()} 个";
            }

            if (!result.IsSafe)
            {
                summary += "\n\n建议：请仔细检查标记的问题，确认代码安全性后再执行";
            }

            result.Summary = summary;
        }

        /// <summary>
        /// 获取安全级别描述
        /// </summary>
        private string GetSecurityLevelDescription(SecurityLevel level)
        {
            switch (level)
            {
                case SecurityLevel.Safe:
                    return "安全";
                case SecurityLevel.Low:
                    return "低风险";
                case SecurityLevel.Medium:
                    return "中等风险";
                case SecurityLevel.High:
                    return "高风险";
                case SecurityLevel.Dangerous:
                    return "危险";
                default:
                    return "未知";
            }
        }
    }
}