using System;
using System.Threading.Tasks;
using System.Text;
using System.Diagnostics;

namespace ExcelAIHelper.Services
{
    /// <summary>
    /// VBA提示构建器
    /// 专门为VBA代码生成构建优化的提示
    /// </summary>
    public class VbaPromptBuilder
    {
        private readonly ContextManager _contextManager;

        public VbaPromptBuilder(ContextManager contextManager)
        {
            _contextManager = contextManager ?? throw new ArgumentNullException(nameof(contextManager));
        }

        /// <summary>
        /// 构建VBA系统提示
        /// </summary>
        /// <returns>系统提示</returns>
        public async Task<string> BuildVbaSystemPromptAsync()
        {
            try
            {
                var systemPrompt = new StringBuilder();
                
                // 基础角色定义
                systemPrompt.AppendLine("你是一名资深Excel VBA开发者和自动化专家。");
                systemPrompt.AppendLine("你的任务是根据用户的中文需求，生成安全、高效的VBA代码。");
                systemPrompt.AppendLine();
                
                // JSON格式要求
                systemPrompt.AppendLine("请严格按照以下JSON格式返回结果：");
                systemPrompt.AppendLine("{");
                systemPrompt.AppendLine("  \"macroName\": \"宏名称（英文，简洁明了）\",");
                systemPrompt.AppendLine("  \"vbaCode\": \"完整的VBA代码\",");
                systemPrompt.AppendLine("  \"description\": \"功能描述（中文）\",");
                systemPrompt.AppendLine("  \"riskLevel\": \"low|medium|high\"");
                systemPrompt.AppendLine("}");
                systemPrompt.AppendLine();
                
                // 安全要求
                systemPrompt.AppendLine("🔒 安全要求（必须严格遵守）：");
                systemPrompt.AppendLine("1. 禁止使用的危险函数：");
                systemPrompt.AppendLine("   - Shell, Kill, CreateObject(\"WScript.Shell\")");
                systemPrompt.AppendLine("   - FileSystemObject, Dir, ChDir, MkDir, RmDir");
                systemPrompt.AppendLine("   - Registry, Environ, Command");
                systemPrompt.AppendLine("   - 任何文件系统操作和外部程序调用");
                systemPrompt.AppendLine();
                systemPrompt.AppendLine("2. 只能使用Excel内置对象：");
                systemPrompt.AppendLine("   - Application, Workbook, Worksheet, Range, Cells");
                systemPrompt.AppendLine("   - WorksheetFunction, Selection, ActiveSheet, ActiveWorkbook");
                systemPrompt.AppendLine("   - 标准VBA函数：Format, CStr, CInt, Left, Right, Mid等");
                systemPrompt.AppendLine();
                
                // 代码质量要求
                systemPrompt.AppendLine("📝 代码质量要求：");
                systemPrompt.AppendLine("1. 必须包含错误处理（On Error GoTo ErrorHandler）");
                systemPrompt.AppendLine("2. 使用有意义的变量名");
                systemPrompt.AppendLine("3. 添加适当的注释");
                systemPrompt.AppendLine("4. 代码结构清晰，逻辑简洁");
                systemPrompt.AppendLine("5. 宏名称使用英文，遵循驼峰命名法");
                systemPrompt.AppendLine();
                
                // 标准代码模板
                systemPrompt.AppendLine("📋 标准代码模板：");
                systemPrompt.AppendLine("```vba");
                systemPrompt.AppendLine("Sub MacroName()");
                systemPrompt.AppendLine("    On Error GoTo ErrorHandler");
                systemPrompt.AppendLine("    ");
                systemPrompt.AppendLine("    ' 声明变量");
                systemPrompt.AppendLine("    Dim ws As Worksheet");
                systemPrompt.AppendLine("    Dim rng As Range");
                systemPrompt.AppendLine("    ");
                systemPrompt.AppendLine("    ' 设置对象引用");
                systemPrompt.AppendLine("    Set ws = ActiveSheet");
                systemPrompt.AppendLine("    ");
                systemPrompt.AppendLine("    ' 主要操作代码");
                systemPrompt.AppendLine("    ' ...");
                systemPrompt.AppendLine("    ");
                systemPrompt.AppendLine("    ' 清理对象引用");
                systemPrompt.AppendLine("    Set ws = Nothing");
                systemPrompt.AppendLine("    Set rng = Nothing");
                systemPrompt.AppendLine("    ");
                systemPrompt.AppendLine("    Exit Sub");
                systemPrompt.AppendLine("    ");
                systemPrompt.AppendLine("ErrorHandler:");
                systemPrompt.AppendLine("    MsgBox \"操作失败: \" & Err.Description, vbCritical, \"错误\"");
                systemPrompt.AppendLine("    ' 清理对象引用");
                systemPrompt.AppendLine("    Set ws = Nothing");
                systemPrompt.AppendLine("    Set rng = Nothing");
                systemPrompt.AppendLine("End Sub");
                systemPrompt.AppendLine("```");
                systemPrompt.AppendLine();
                
                // 常用操作示例
                systemPrompt.AppendLine("💡 常用操作示例：");
                systemPrompt.AppendLine("• 设置单元格值：Range(\"A1\").Value = \"Hello\"");
                systemPrompt.AppendLine("• 应用公式：Range(\"B1\").Formula = \"=SUM(A1:A10)\"");
                systemPrompt.AppendLine("• 设置字体：Range(\"A1\").Font.Bold = True");
                systemPrompt.AppendLine("• 设置背景色：Range(\"A1\").Interior.Color = RGB(255, 0, 0)");
                systemPrompt.AppendLine("• 循环处理：For i = 1 To 10: Cells(i, 1).Value = i: Next i");
                systemPrompt.AppendLine("• 查找数据：Set rng = ws.Range(\"A:A\").Find(\"查找内容\")");
                systemPrompt.AppendLine("• 排序数据：ws.Range(\"A1:C10\").Sort Key1:=ws.Range(\"A1\")");
                systemPrompt.AppendLine();
                
                // 获取当前Excel上下文
                var context = await _contextManager.GetCurrentContextAsync();
                if (context != null)
                {
                    systemPrompt.AppendLine("📊 当前Excel环境信息：");
                    if (context.CurrentWorksheet != null)
                    {
                        systemPrompt.AppendLine($"• 当前工作表：{context.CurrentWorksheet.Name}");
                    }
                    if (context.SelectedRange != null && !string.IsNullOrEmpty(context.SelectedRange.Address))
                    {
                        systemPrompt.AppendLine($"• 选中区域：{context.SelectedRange.Address}");
                        systemPrompt.AppendLine($"• 区域大小：{context.SelectedRange.RowCount}行 × {context.SelectedRange.ColumnCount}列");
                    }
                    systemPrompt.AppendLine();
                }
                
                // 风险级别说明
                systemPrompt.AppendLine("⚠️ 风险级别说明：");
                systemPrompt.AppendLine("• low: 基础操作，如设置值、格式等");
                systemPrompt.AppendLine("• medium: 复杂操作，如循环、查找替换等");
                systemPrompt.AppendLine("• high: 涉及大量数据或复杂逻辑的操作");
                systemPrompt.AppendLine();
                
                systemPrompt.AppendLine("请根据用户需求生成相应的VBA代码，确保代码安全、高效、易读。");
                
                Debug.WriteLine("[VbaPromptBuilder] VBA系统提示构建完成");
                return systemPrompt.ToString();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[VbaPromptBuilder] 构建系统提示失败: {ex.Message}");
                return GetFallbackSystemPrompt();
            }
        }

        /// <summary>
        /// 构建VBA用户提示
        /// </summary>
        /// <param name="userRequest">用户请求</param>
        /// <returns>用户提示</returns>
        public async Task<string> BuildVbaUserPromptAsync(string userRequest)
        {
            try
            {
                var userPrompt = new StringBuilder();
                
                userPrompt.AppendLine("🎯 用户需求：");
                userPrompt.AppendLine(userRequest);
                userPrompt.AppendLine();
                
                // 获取当前上下文信息
                var context = await _contextManager.GetCurrentContextAsync();
                if (context != null)
                {
                    userPrompt.AppendLine("📋 当前状态：");
                    if (context.CurrentWorksheet != null)
                    {
                        userPrompt.AppendLine($"• 工作表：{context.CurrentWorksheet.Name}");
                    }
                    
                    if (context.SelectedRange != null && !string.IsNullOrEmpty(context.SelectedRange.Address))
                    {
                        userPrompt.AppendLine($"• 选中区域：{context.SelectedRange.Address}");
                        
                        // 如果选区不大，可以提供一些上下文数据
                        if (context.SelectedRange.RowCount <= 10 && context.SelectedRange.ColumnCount <= 10)
                        {
                            try
                            {
                                var contextDescription = await _contextManager.GetContextDescriptionAsync();
                                if (!string.IsNullOrEmpty(contextDescription))
                                {
                                    userPrompt.AppendLine("• 选区内容概览：");
                                    userPrompt.AppendLine(contextDescription);
                                }
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"[VbaPromptBuilder] 获取上下文描述失败: {ex.Message}");
                            }
                        }
                    }
                    else
                    {
                        userPrompt.AppendLine("• 选中区域：当前活动单元格");
                    }
                    
                    userPrompt.AppendLine();
                }
                
                // 添加特定指导
                userPrompt.AppendLine("📝 请注意：");
                userPrompt.AppendLine("1. 如果用户提到\"当前选区\"或\"选中的区域\"，请使用Selection对象");
                userPrompt.AppendLine("2. 如果用户提到具体的单元格地址（如A1、B2），请直接使用Range对象");
                userPrompt.AppendLine("3. 如果需要循环处理数据，请考虑数据量大小，避免过长的执行时间");
                userPrompt.AppendLine("4. 生成的宏名称要能反映功能，如SetCellValue、FormatRange等");
                userPrompt.AppendLine();
                
                userPrompt.AppendLine("请严格按照JSON格式返回VBA代码。");
                
                Debug.WriteLine($"[VbaPromptBuilder] VBA用户提示构建完成，用户请求: {userRequest}");
                return userPrompt.ToString();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[VbaPromptBuilder] 构建用户提示失败: {ex.Message}");
                return $"用户需求：{userRequest}\n\n请生成相应的VBA代码来实现这个需求，严格按照JSON格式返回。";
            }
        }

        /// <summary>
        /// 获取VBA安全指导
        /// </summary>
        /// <returns>安全指导文本</returns>
        private string GetVbaSafetyInstructions()
        {
            return @"VBA安全编程指导：

🔒 禁止使用的函数和对象：
• Shell - 执行外部程序
• Kill - 删除文件
• CreateObject(""WScript.Shell"") - 创建脚本对象
• CreateObject(""Scripting.FileSystemObject"") - 文件系统对象
• Dir, ChDir, MkDir, RmDir - 文件夹操作
• Registry - 注册表操作
• Environ - 环境变量
• Command - 命令行参数

✅ 推荐使用的安全对象：
• Application - Excel应用程序对象
• Workbook, Workbooks - 工作簿对象
• Worksheet, Worksheets - 工作表对象
• Range, Cells - 单元格对象
• Selection - 选择对象
• WorksheetFunction - 工作表函数

💡 最佳实践：
• 始终使用错误处理
• 及时清理对象引用
• 避免无限循环
• 使用有意义的变量名
• 添加必要的注释";
        }

        /// <summary>
        /// 获取VBA模板示例
        /// </summary>
        /// <returns>模板示例</returns>
        private string GetVbaTemplateExamples()
        {
            return @"VBA代码模板示例：

1. 基础数据操作：
```vba
Sub SetCellData()
    On Error GoTo ErrorHandler
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ws.Range(""A1"").Value = ""Hello World""
    ws.Range(""B1"").Value = 100
    
    Set ws = Nothing
    Exit Sub
ErrorHandler:
    MsgBox ""操作失败: "" & Err.Description
    Set ws = Nothing
End Sub
```

2. 循环处理数据：
```vba
Sub ProcessData()
    On Error GoTo ErrorHandler
    Dim ws As Worksheet
    Dim i As Integer
    
    Set ws = ActiveSheet
    
    For i = 1 To 10
        ws.Cells(i, 1).Value = i
        ws.Cells(i, 2).Value = i * 2
    Next i
    
    Set ws = Nothing
    Exit Sub
ErrorHandler:
    MsgBox ""操作失败: "" & Err.Description
    Set ws = Nothing
End Sub
```

3. 格式设置：
```vba
Sub FormatCells()
    On Error GoTo ErrorHandler
    Dim rng As Range
    
    Set rng = Selection
    
    With rng
        .Font.Bold = True
        .Font.Color = RGB(255, 0, 0)
        .Interior.Color = RGB(255, 255, 0)
    End With
    
    Set rng = Nothing
    Exit Sub
ErrorHandler:
    MsgBox ""操作失败: "" & Err.Description
    Set rng = Nothing
End Sub
```";
        }

        /// <summary>
        /// 获取备用系统提示（简化版）
        /// </summary>
        /// <returns>备用系统提示</returns>
        private string GetFallbackSystemPrompt()
        {
            return @"你是Excel VBA专家。根据用户需求生成VBA代码，必须返回JSON格式：

{
  ""macroName"": ""宏名称"",
  ""vbaCode"": ""完整VBA代码"",
  ""description"": ""功能说明"",
  ""riskLevel"": ""low""
}

安全规则：
1. 只能使用Excel对象模型（Application, Workbook, Worksheet等）
2. 禁止文件系统操作、网络访问、系统调用
3. 必须包含错误处理
4. 代码要简洁高效

示例模板：
Sub GeneratedMacro()
    On Error GoTo ErrorHandler
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ' 你的操作代码
    Set ws = Nothing
    Exit Sub
ErrorHandler:
    MsgBox ""操作失败: "" & Err.Description
    Set ws = Nothing
End Sub";
        }
    }
}