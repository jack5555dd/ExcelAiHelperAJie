using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelAIHelper.Exceptions;
using ExcelAIHelper.Models;
using ExcelAIHelper.Services;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAIHelper
{
    public partial class ChatPaneControl : UserControl
    {
        private DeepSeekClient _aiClient;
        private ExcelOperationEngine _operationEngine;
        private ContextManager _contextManager;
        private PromptBuilder _promptBuilder;
        private InstructionParser _instructionParser;
        private OperationDispatcher _operationDispatcher;
        
        // VBA相关服务
        private VbaInjectionEngine _vbaInjectionEngine;
        private VbaPromptBuilder _vbaPromptBuilder;
        
        private string _currentUserRequest;
        private VbaGenerationResult _currentVbaResult;
        private bool _vbaCodeExpanded = false;
        
        public ChatPaneControl()
        {
            InitializeComponent();
            InitializeServices();
        }
        
        private void InitializeServices()
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("ChatPaneControl InitializeServices called");
                
                // 显示基本的欢迎信息
                if (rtbChatHistory != null)
                {
                    AppendToChatHistory("系统", "Excel AI助手正在启动...", Color.Blue);
                }
                
                // 安全地初始化UI状态，不依赖Excel对象
                SafeInitializeExecutionModeUI();
                
                // 使用Timer延迟初始化服务，避免在Excel完全启动之前初始化
                var initTimer = new System.Windows.Forms.Timer();
                initTimer.Interval = 2000; // 2秒延迟
                initTimer.Tick += (s, args) =>
                {
                    initTimer.Stop();
                    initTimer.Dispose();
                    DelayedInitializeServices();
                };
                initTimer.Start();
                
                System.Diagnostics.Debug.WriteLine("ChatPaneControl initialization timer started");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"InitializeServices failed: {ex}");
                // 确保至少显示错误信息
                try
                {
                    if (rtbChatHistory != null)
                    {
                        AppendToChatHistory("系统", $"❌ 初始化失败: {ex.Message}", Color.Red);
                    }
                }
                catch
                {
                    // 忽略二次异常
                }
            }
        }
        
        private void DelayedInitializeServices()
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("DelayedInitializeServices called");
                
                // 检查ThisAddIn是否可用
                if (Globals.ThisAddIn == null)
                {
                    System.Diagnostics.Debug.WriteLine("Globals.ThisAddIn is null, scheduling retry");
                    AppendToChatHistory("系统", "⚠️ Excel加载项尚未完全初始化，正在重试...", Color.Orange);
                    
                    // 再次延迟重试
                    var retryTimer = new System.Windows.Forms.Timer();
                    retryTimer.Interval = 3000; // 3秒后重试
                    retryTimer.Tick += (s, args) =>
                    {
                        retryTimer.Stop();
                        retryTimer.Dispose();
                        DelayedInitializeServices();
                    };
                    retryTimer.Start();
                    return;
                }

                // Get Excel application from ThisAddIn
                Excel.Application excelApp = null;
                try
                {
                    excelApp = Globals.ThisAddIn.Application;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Failed to get Excel Application: {ex}");
                    AppendToChatHistory("系统", "⚠️ Excel应用程序访问失败，将在基础模式下运行", Color.Orange);
                }
                
                if (excelApp == null)
                {
                    AppendToChatHistory("系统", "⚠️ Excel应用程序尚未可用，将在基础模式下运行", Color.Orange);
                    InitializeBasicMode();
                    return;
                }
                
                // Initialize services
                try
                {
                    _aiClient = new DeepSeekClient();
                    _operationEngine = new ExcelOperationEngine(excelApp);
                    _contextManager = new ContextManager(excelApp, _operationEngine);
                    _promptBuilder = new PromptBuilder(_contextManager);
                    _instructionParser = new InstructionParser();
                    _operationDispatcher = new OperationDispatcher(
                        _aiClient, 
                        _promptBuilder, 
                        _instructionParser, 
                        excelApp);
                    
                    // Initialize VBA services
                    _vbaInjectionEngine = new VbaInjectionEngine(_aiClient);
                    _vbaPromptBuilder = new VbaPromptBuilder(_contextManager);
                    
                    System.Diagnostics.Debug.WriteLine("All services initialized successfully");
                    
                    AppendToChatHistory("系统", "✅ Excel AI助手已启动，支持自然语言操作Excel表格。", Color.Green);
                    AppendToChatHistory("系统", "💡 提示：您可以说\"在A1输入100\"、\"给选中区域设置红色背景\"等。", Color.Gray);
                    
                    // 延迟检查VBA状态，避免阻塞主要功能
                    CheckVbaStatusAsync();
                    
                    // Test API connection if API key is set
                    if (!string.IsNullOrEmpty(Properties.Settings.Default.ApiKey))
                    {
                        _ = TestApiConnectionAsync(); // Fire and forget
                    }
                    else
                    {
                        AppendToChatHistory("系统", "⚠️ 请先点击\"API 设置\"配置DeepSeek API密钥。", Color.Orange);
                    }
                }
                catch (Exception serviceEx)
                {
                    System.Diagnostics.Debug.WriteLine($"Service initialization failed: {serviceEx}");
                    AppendToChatHistory("系统", $"❌ 服务初始化失败: {serviceEx.Message}", Color.Red);
                    InitializeBasicMode();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"DelayedInitializeServices failed: {ex}");
                AppendToChatHistory("系统", $"❌ 延迟初始化失败: {ex.Message}", Color.Red);
                InitializeBasicMode();
            }
        }
        
        /// <summary>
        /// 初始化基础模式（仅聊天功能）
        /// </summary>
        private void InitializeBasicMode()
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("Initializing basic mode");
                
                // 只初始化AI客户端
                _aiClient = new DeepSeekClient();
                
                AppendToChatHistory("系统", "⚠️ 以基础模式启动，仅支持聊天功能", Color.Orange);
                AppendToChatHistory("系统", "💡 您可以询问Excel相关问题，但无法直接操作表格", Color.Gray);
                
                // 强制切换到仅聊天模式
                if (rbChatOnlyMode != null)
                {
                    rbChatOnlyMode.Checked = true;
                    rbVstoMode.Enabled = false;
                    rbVbaMode.Enabled = false;
                }
                
                // Test API connection if API key is set
                if (!string.IsNullOrEmpty(Properties.Settings.Default.ApiKey))
                {
                    _ = TestApiConnectionAsync();
                }
                else
                {
                    AppendToChatHistory("系统", "⚠️ 请先点击\"API 设置\"配置DeepSeek API密钥。", Color.Orange);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"InitializeBasicMode failed: {ex}");
                AppendToChatHistory("系统", $"❌ 基础模式初始化失败: {ex.Message}", Color.Red);
            }
        }
        
        /// <summary>
        /// 异步检查VBA状态
        /// </summary>
        private async void CheckVbaStatusAsync()
        {
            try
            {
                await Task.Run(() =>
                {
                    try
                    {
                        System.Diagnostics.Debug.WriteLine("Checking VBA status asynchronously");
                        
                        // 安全地检查VBA状态
                        bool vbaAvailable = false;
                        try
                        {
                            // 不直接调用ExecutionModeManager.Initialize()，避免可能的异常
                            vbaAvailable = CheckVbaEnvironmentSafely();
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"VBA status check failed: {ex}");
                        }
                        
                        // 在UI线程中更新状态
                        this.BeginInvoke(new Action(() =>
                        {
                            try
                            {
                                if (rbVbaMode != null)
                                {
                                    if (vbaAvailable)
                                    {
                                        rbVbaMode.Enabled = true;
                                        rbVbaMode.Text = "VBA";
                                        AppendToChatHistory("系统", "🔧 VBA模式可用", Color.Blue);
                                    }
                                    else
                                    {
                                        rbVbaMode.Enabled = false;
                                        rbVbaMode.Text = "VBA (不可用)";
                                        AppendToChatHistory("系统", "⚠️ VBA模式不可用，请检查VBE访问权限", Color.Orange);
                                    }
                                }
                            }
                            catch (Exception uiEx)
                            {
                                System.Diagnostics.Debug.WriteLine($"UI update failed: {uiEx}");
                            }
                        }));
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Async VBA check failed: {ex}");
                    }
                });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"CheckVbaStatusAsync failed: {ex}");
            }
        }
        
        /// <summary>
        /// 安全地检查VBA环境
        /// </summary>
        private bool CheckVbaEnvironmentSafely()
        {
            try
            {
                // 检查Excel应用程序是否可用
                if (Globals.ThisAddIn?.Application == null)
                {
                    return false;
                }
                
                var excelApp = Globals.ThisAddIn.Application;
                
                // 简单的VBA环境检查，不进行复杂操作
                try
                {
                    var vbe = excelApp.VBE;
                    return vbe != null;
                }
                catch
                {
                    return false;
                }
            }
            catch
            {
                return false;
            }
        }
        
        private async void btnSend_Click(object sender, EventArgs e)
        {
            await SendUserMessageAsync();
        }
        
        private async void txtUserInput_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && e.Control)
            {
                e.SuppressKeyPress = true;
                await SendUserMessageAsync();
            }
        }
        
        private async Task SendUserMessageAsync()
        {
            string userMessage = txtUserInput.Text.Trim();
            if (string.IsNullOrEmpty(userMessage))
                return;
            
            // Disable input while processing
            SetInputEnabled(false);
            
            try
            {
                // Save the current request
                _currentUserRequest = userMessage;
                
                // Display user message
                AppendToChatHistory("用户", userMessage, Color.Blue);
                
                // Clear input
                txtUserInput.Clear();
                
                // Check if API key is set
                if (string.IsNullOrEmpty(Properties.Settings.Default.ApiKey))
                {
                    AppendToChatHistory("系统", "请先设置API密钥。点击'API 设置'按钮进行配置。", Color.Red);
                    SetInputEnabled(true);
                    return;
                }
                
                // 根据当前执行模式处理请求
                var currentMode = ExecutionModeManager.CurrentMode;
                
                switch (currentMode)
                {
                    case ExecutionMode.VSTO:
                        await HandleVstoModeRequestAsync(userMessage);
                        break;
                        
                    case ExecutionMode.VBA:
                        await HandleVbaModeRequestAsync(userMessage);
                        break;
                        
                    case ExecutionMode.ChatOnly:
                        await HandleChatOnlyModeRequestAsync(userMessage);
                        break;
                        
                    default:
                        AppendToChatHistory("系统", "未知的执行模式", Color.Red);
                        break;
                }
            }
            catch (Exception ex)
            {
                HandleError(ex);
            }
            finally
            {
                SetInputEnabled(true);
            }
        }

        /// <summary>
        /// 处理VSTO模式的用户请求
        /// </summary>
        private async Task HandleVstoModeRequestAsync(string userMessage)
        {
            try
            {
                // Show thinking indicator
                AppendToChatHistory("系统", "AI思考中...", Color.Gray);
                
                // Get preview of operations using new OperationResult
                var previewResult = await _operationDispatcher.ApplyAsync(userMessage, true);
                
                // Remove thinking indicator
                RemoveLastChatHistoryLine();
                
                if (previewResult.Success)
                {
                    // Show preview
                    ShowPreview(previewResult.GetUserMessage());
                }
                else
                {
                    // Handle error
                    AppendToChatHistory("系统", previewResult.GetUserMessage(), Color.Red);
                    
                    // If it's a format error and can retry, offer retry option
                    if (previewResult.CanRetry && previewResult.ErrorType == "协议格式错误")
                    {
                        AppendToChatHistory("系统", "正在尝试重新生成...", Color.Gray);
                        await RetryOperationAsync(userMessage);
                    }
                }
            }
            catch (Exception)
            {
                RemoveLastChatHistoryLine();
                throw;
            }
        }

        /// <summary>
        /// 处理仅聊天模式的用户请求
        /// </summary>
        private async Task HandleChatOnlyModeRequestAsync(string userMessage)
        {
            try
            {
                AppendToChatHistory("AI", "正在思考您的问题...", Color.Blue);
                
                // 构建仅聊天的提示
                string systemPrompt = "你是一个Excel专家助手。用户会问你关于Excel的问题，请提供详细的回答和建议，但不要生成任何可执行的代码或指令。只需要用自然语言回答问题。";
                string userPrompt = userMessage;
                
                // 调用AI获取回答
                string aiResponse = await _aiClient.AskAsync(userPrompt, systemPrompt);
                
                // 移除"正在思考"提示
                RemoveLastChatHistoryLine();
                
                // 显示AI回答
                AppendToChatHistory("AI", aiResponse, Color.Green);
            }
            catch (Exception)
            {
                RemoveLastChatHistoryLine();
                throw;
            }
        }
        
        private async void btnConfirm_Click(object sender, EventArgs e)
        {
            try
            {
                // Hide preview
                HidePreview();
                
                // Show executing indicator
                AppendToChatHistory("系统", "执行操作中...", Color.Gray);
                
                // Execute operations using new OperationResult
                var result = await _operationDispatcher.ApplyAsync(_currentUserRequest, false);
                
                // Remove executing indicator
                RemoveLastChatHistoryLine();
                
                if (result.Success)
                {
                    // Show success result
                    AppendToChatHistory("AI", result.GetUserMessage(), Color.Green);
                }
                else
                {
                    // Show error result
                    AppendToChatHistory("系统", result.GetUserMessage(), Color.Red);
                    
                    // If it's a format error and can retry, offer retry option
                    if (result.CanRetry && result.ErrorType == "协议格式错误")
                    {
                        AppendToChatHistory("系统", "正在尝试重新执行...", Color.Gray);
                        await RetryOperationAsync(_currentUserRequest, false);
                    }
                }
            }
            catch (Exception ex)
            {
                HandleError(ex);
            }
            finally
            {
                SetInputEnabled(true);
            }
        }
        
        private void btnCancel_Click(object sender, EventArgs e)
        {
            HidePreview();
            AppendToChatHistory("系统", "操作已取消", Color.Gray);
            SetInputEnabled(true);
        }
        
        private void AppendToChatHistory(string sender, string message, Color color)
        {
            if (rtbChatHistory.InvokeRequired)
            {
                rtbChatHistory.Invoke(new Action(() => AppendToChatHistory(sender, message, color)));
                return;
            }

            rtbChatHistory.SelectionStart = rtbChatHistory.TextLength;
            rtbChatHistory.SelectionLength = 0;
            
            // Add sender with bold formatting
            rtbChatHistory.SelectionFont = new Font(rtbChatHistory.Font, FontStyle.Bold);
            rtbChatHistory.SelectionColor = color;
            rtbChatHistory.AppendText($"{sender}: ");
            
            // Add message with normal formatting
            rtbChatHistory.SelectionFont = new Font(rtbChatHistory.Font, FontStyle.Regular);
            rtbChatHistory.SelectionColor = Color.Black;
            rtbChatHistory.AppendText($"{message}{Environment.NewLine}{Environment.NewLine}");
            
            // Scroll to end
            rtbChatHistory.ScrollToCaret();
        }
        
        private void RemoveLastChatHistoryLine()
        {
            if (rtbChatHistory.InvokeRequired)
            {
                rtbChatHistory.Invoke(new Action(RemoveLastChatHistoryLine));
                return;
            }

            // Find the last line break
            int lastLineBreakIndex = rtbChatHistory.Text.LastIndexOf(Environment.NewLine);
            if (lastLineBreakIndex > 0)
            {
                // Find the previous line break
                int previousLineBreakIndex = rtbChatHistory.Text.LastIndexOf(Environment.NewLine, lastLineBreakIndex - 1);
                if (previousLineBreakIndex >= 0)
                {
                    // Remove the last line
                    rtbChatHistory.Text = rtbChatHistory.Text.Substring(0, previousLineBreakIndex + Environment.NewLine.Length);
                    rtbChatHistory.SelectionStart = rtbChatHistory.TextLength;
                    rtbChatHistory.ScrollToCaret();
                }
            }
        }
        
        private void ShowPreview(string previewText)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => ShowPreview(previewText)));
                return;
            }

            rtbPreview.Text = previewText;
            pnlPreview.Visible = true;
            
            // Adjust layout
            rtbChatHistory.Height = pnlPreview.Location.Y;
        }
        
        private void HidePreview()
        {
            if (InvokeRequired)
            {
                Invoke(new Action(HidePreview));
                return;
            }

            pnlPreview.Visible = false;
            
            // Restore layout
            rtbChatHistory.Height = pnlInput.Location.Y;
        }
        
        private void SetInputEnabled(bool enabled)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => SetInputEnabled(enabled)));
                return;
            }

            txtUserInput.Enabled = enabled;
            btnSend.Enabled = enabled;
        }
        
        private void HandleError(Exception ex)
        {
            string errorMessage;
            
            if (ex is AiFormatException formatEx)
            {
                errorMessage = $"❌ AI响应格式错误: {formatEx.GetUserFriendlyMessage()}";
            }
            else if (ex is AiOperationException aiEx)
            {
                errorMessage = $"❌ 操作执行错误: {aiEx.Message}";
            }
            else
            {
                errorMessage = $"❌ 系统错误: {ex.Message}";
            }
            
            // Remove any pending indicators
            RemoveLastChatHistoryLine();
            
            AppendToChatHistory("系统", errorMessage, Color.Red);
            System.Diagnostics.Debug.WriteLine($"Error: {ex}");
        }
        
        /// <summary>
        /// 重试操作（用于协议违规后的自动重试）
        /// </summary>
        /// <param name="userRequest">用户请求</param>
        /// <param name="isDryRun">是否为预览模式</param>
        private async Task RetryOperationAsync(string userRequest, bool isDryRun = true)
        {
            try
            {
                var retryResult = await _operationDispatcher.RetryAsync(userRequest, "协议格式错误", isDryRun);
                
                // Remove retry indicator
                RemoveLastChatHistoryLine();
                
                if (retryResult.Success)
                {
                    if (isDryRun)
                    {
                        ShowPreview(retryResult.GetUserMessage());
                    }
                    else
                    {
                        AppendToChatHistory("AI", retryResult.GetUserMessage(), Color.Green);
                        SetInputEnabled(true);
                    }
                }
                else
                {
                    AppendToChatHistory("系统", $"❌ 重试失败: {retryResult.ErrorMessage}", Color.Red);
                    AppendToChatHistory("系统", "💡 请尝试重新描述您的需求，使用更清晰的表达。", Color.Gray);
                    SetInputEnabled(true);
                }
            }
            catch (Exception ex)
            {
                RemoveLastChatHistoryLine();
                AppendToChatHistory("系统", $"❌ 重试过程出错: {ex.Message}", Color.Red);
                SetInputEnabled(true);
            }
        }

        private async Task TestApiConnectionAsync()
        {
            try
            {
                AppendToChatHistory("系统", "正在测试API连接...", Color.Gray);
                
                var testResult = await _aiClient.TestConnectionAsync();
                
                if (testResult.Item1)
                {
                    AppendToChatHistory("系统", "✓ API连接正常", Color.Green);
                }
                else
                {
                    AppendToChatHistory("系统", $"✗ API连接失败: {testResult.Item2}", Color.Red);
                    AppendToChatHistory("系统", "请点击'API 设置'按钮检查配置", Color.Orange);
                }
            }
            catch (Exception ex)
            {
                AppendToChatHistory("系统", $"✗ API连接测试异常: {ex.Message}", Color.Red);
            }
        }

        #region VBA功能相关方法

        /// <summary>
        /// 安全地初始化执行模式UI
        /// </summary>
        private void SafeInitializeExecutionModeUI()
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("SafeInitializeExecutionModeUI called");
                
                // 检查控件是否已初始化
                if (rbVstoMode == null || rbVbaMode == null || rbChatOnlyMode == null)
                {
                    System.Diagnostics.Debug.WriteLine("UI controls not yet initialized, skipping");
                    return;
                }

                // 默认设置为VSTO模式
                rbVstoMode.Checked = true;
                rbVbaMode.Checked = false;
                rbChatOnlyMode.Checked = false;
                
                // 默认禁用VBA模式，直到检查完成
                rbVbaMode.Enabled = false;
                rbVbaMode.Text = "VBA (检查中...)";
                
                // 安全地初始化VBA代码面板状态
                SafeUpdateVbaCodePanelVisibility();
                
                System.Diagnostics.Debug.WriteLine("SafeInitializeExecutionModeUI completed");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SafeInitializeExecutionModeUI failed: {ex}");
                // 确保至少有一个模式被选中
                try
                {
                    if (rbVstoMode != null)
                        rbVstoMode.Checked = true;
                }
                catch
                {
                    // 忽略二次异常
                }
            }
        }
        
        /// <summary>
        /// 安全地更新VBA代码面板可见性
        /// </summary>
        private void SafeUpdateVbaCodePanelVisibility()
        {
            try
            {
                // 检查控件是否已初始化
                if (pnlVbaCode == null || rtbChatHistory == null || pnlInput == null)
                {
                    System.Diagnostics.Debug.WriteLine("VBA panel controls not initialized, skipping");
                    return;
                }

                // 默认不显示VBA面板
                pnlVbaCode.Visible = false;
                
                // 恢复聊天历史区域高度
                rtbChatHistory.Height = pnlInput.Location.Y - rtbChatHistory.Location.Y;
                
                System.Diagnostics.Debug.WriteLine("VBA panel visibility updated safely");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SafeUpdateVbaCodePanelVisibility failed: {ex}");
            }
        }

        /// <summary>
        /// 更新VBA代码面板可见性
        /// </summary>
        private void UpdateVbaCodePanelVisibility()
        {
            try
            {
                // 检查控件是否已初始化
                if (pnlVbaCode == null || rtbChatHistory == null || pnlInput == null)
                {
                    return;
                }

                // 默认不显示VBA面板，因为ExecutionModeManager可能未初始化
                bool showVbaPanel = false;
                
                try
                {
                    showVbaPanel = ExecutionModeManager.IsInitialized && (ExecutionModeManager.CurrentMode == ExecutionMode.VBA);
                }
                catch
                {
                    // 如果ExecutionModeManager访问失败，默认不显示VBA面板
                    showVbaPanel = false;
                }

                pnlVbaCode.Visible = showVbaPanel;
                
                if (showVbaPanel)
                {
                    // 调整聊天历史区域高度
                    rtbChatHistory.Height = pnlVbaCode.Location.Y - rtbChatHistory.Location.Y;
                }
                else
                {
                    // 恢复聊天历史区域高度
                    rtbChatHistory.Height = pnlInput.Location.Y - rtbChatHistory.Location.Y;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"UpdateVbaCodePanelVisibility failed: {ex}");
            }
        }

        /// <summary>
        /// 处理VBA模式的用户请求
        /// </summary>
        private async Task HandleVbaModeRequestAsync(string userRequest)
        {
            try
            {
                AppendToChatHistory("AI", "正在生成VBA代码...", Color.Blue);
                
                // 生成VBA代码
                var vbaResult = await _vbaInjectionEngine.GenerateVbaCodeAsync(userRequest);
                
                // 移除"正在生成"提示
                RemoveLastChatHistoryLine();
                
                if (!vbaResult.Success)
                {
                    AppendToChatHistory("AI", $"❌ VBA代码生成失败: {vbaResult.ErrorMessage}", Color.Red);
                    return;
                }
                
                // 保存当前VBA结果
                _currentVbaResult = vbaResult;
                
                // 显示VBA代码
                ShowVbaCode(vbaResult);
                
                // 在聊天记录中显示描述
                AppendToChatHistory("AI", $"✅ 已生成VBA代码: {vbaResult.Description}", Color.Green);
                AppendToChatHistory("AI", $"🔧 宏名称: {vbaResult.MacroName}", Color.Blue);
                AppendToChatHistory("AI", $"⚠️ 风险级别: {GetRiskLevelDescription(vbaResult.RiskLevel)}", Color.Orange);
                
                if (vbaResult.SecurityScanResult != null && !vbaResult.SecurityScanResult.IsSafe)
                {
                    AppendToChatHistory("AI", $"🚨 安全警告: {vbaResult.SecurityScanResult.Summary}", Color.Red);
                }
                else
                {
                    AppendToChatHistory("AI", "✅ 安全扫描通过，可以安全执行", Color.Green);
                }
            }
            catch (Exception ex)
            {
                RemoveLastChatHistoryLine();
                AppendToChatHistory("系统", $"❌ VBA处理失败: {ex.Message}", Color.Red);
            }
        }

        /// <summary>
        /// 显示VBA代码
        /// </summary>
        private void ShowVbaCode(VbaGenerationResult vbaResult)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => ShowVbaCode(vbaResult)));
                return;
            }
            
            rtbVbaCode.Text = vbaResult.VbaCode;
            pnlVbaCode.Visible = true;
            btnExecuteVba.Enabled = vbaResult.Success;
            
            // 如果代码较长，默认折叠
            if (vbaResult.VbaCode.Length > 200)
            {
                CollapseVbaCode();
            }
            else
            {
                ExpandVbaCode();
            }
        }

        /// <summary>
        /// 展开VBA代码显示
        /// </summary>
        private void ExpandVbaCode()
        {
            _vbaCodeExpanded = true;
            pnlVbaCode.Height = 120;
            rtbVbaCode.Height = 75;
            btnToggleVbaCode.Text = "折叠";
            
            // 调整其他控件位置
            pnlInput.Location = new Point(pnlInput.Location.X, pnlVbaCode.Location.Y + pnlVbaCode.Height);
        }

        /// <summary>
        /// 折叠VBA代码显示
        /// </summary>
        private void CollapseVbaCode()
        {
            _vbaCodeExpanded = false;
            pnlVbaCode.Height = 65;
            rtbVbaCode.Height = 20;
            btnToggleVbaCode.Text = "展开";
            
            // 调整其他控件位置
            pnlInput.Location = new Point(pnlInput.Location.X, pnlVbaCode.Location.Y + pnlVbaCode.Height);
        }

        /// <summary>
        /// 获取风险级别描述
        /// </summary>
        private string GetRiskLevelDescription(string riskLevel)
        {
            switch (riskLevel?.ToLower())
            {
                case "low":
                    return "低风险";
                case "medium":
                    return "中等风险";
                case "high":
                    return "高风险";
                default:
                    return "未知风险";
            }
        }

        /// <summary>
        /// 检查并更新VBA状态
        /// </summary>
        private void CheckAndUpdateVbaStatus()
        {
            try
            {
                // 检查Excel应用程序是否可用
                if (Globals.ThisAddIn?.Application == null)
                {
                    AppendToChatHistory("系统", "⚠️ Excel应用程序尚未完全初始化，稍后将自动检查VBA状态", Color.Orange);
                    return;
                }

                // 强制刷新VBA状态
                ExecutionModeManager.RefreshVbaStatus();
                
                // 更新UI状态
                SafeInitializeExecutionModeUI();
                
                // 显示状态信息
                if (ExecutionModeManager.IsVbaEnabled)
                {
                    AppendToChatHistory("系统", "🔧 VBA模式可用，您可以切换到VBA模式体验更强大的功能。", Color.Blue);
                }
                else
                {
                    AppendToChatHistory("系统", "⚠️ VBA模式不可用，请检查VBE访问权限设置。", Color.Orange);
                    
                    // 安全地运行快速诊断
                    try
                    {
                        var (isAvailable, reason) = VbaDiagnostics.QuickCheck();
                        if (!isAvailable)
                        {
                            AppendToChatHistory("系统", $"📋 诊断结果: {reason}", Color.Red);
                            AppendToChatHistory("系统", "💡 提示: 尝试切换到VBA模式时可以运行详细诊断", Color.Gray);
                        }
                    }
                    catch (Exception diagEx)
                    {
                        AppendToChatHistory("系统", $"📋 快速诊断失败: {diagEx.Message}", Color.Red);
                    }
                }
            }
            catch (Exception ex)
            {
                AppendToChatHistory("系统", $"❌ VBA状态检查失败: {ex.Message}", Color.Red);
                System.Diagnostics.Debug.WriteLine($"CheckAndUpdateVbaStatus failed: {ex}");
            }
        }

        #endregion

        #region VBA诊断方法

        /// <summary>
        /// 显示VBA诊断结果
        /// </summary>
        private void ShowVbaDiagnostics()
        {
            try
            {
                AppendToChatHistory("系统", "🔍 正在运行VBA环境诊断...", Color.Blue);
                
                // 运行诊断
                string diagnosticsReport = VbaDiagnostics.RunFullDiagnostics();
                
                // 创建诊断结果窗口
                var diagForm = new Form
                {
                    Text = "VBA环境诊断报告",
                    Size = new Size(800, 600),
                    StartPosition = FormStartPosition.CenterParent,
                    ShowIcon = false,
                    MaximizeBox = true,
                    MinimizeBox = false
                };

                var textBox = new TextBox
                {
                    Multiline = true,
                    ScrollBars = ScrollBars.Both,
                    ReadOnly = true,
                    Dock = DockStyle.Fill,
                    Font = new Font("Consolas", 9),
                    Text = diagnosticsReport
                };

                var buttonPanel = new Panel
                {
                    Height = 40,
                    Dock = DockStyle.Bottom
                };

                var copyButton = new Button
                {
                    Text = "复制报告",
                    Size = new Size(80, 30),
                    Location = new Point(10, 5)
                };
                copyButton.Click += (s, e) =>
                {
                    Clipboard.SetText(diagnosticsReport);
                    MessageBox.Show("诊断报告已复制到剪贴板", "复制成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                };

                var refreshButton = new Button
                {
                    Text = "重新检查",
                    Size = new Size(80, 30),
                    Location = new Point(100, 5)
                };
                refreshButton.Click += (s, e) =>
                {
                    ExecutionModeManager.RefreshVbaStatus();
                    textBox.Text = VbaDiagnostics.RunFullDiagnostics();
                    
                    // 检查是否已修复
                    if (ExecutionModeManager.IsVbaEnabled)
                    {
                        MessageBox.Show("✅ VBA环境已可用！现在可以使用VBA模式了。", "检查成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        diagForm.Close();
                        
                        // 自动切换到VBA模式
                        rbVbaMode.Checked = true;
                        rbVbaMode.Enabled = true;
                        rbVbaMode.Text = "VBA";
                        ExecutionModeManager.SwitchMode(ExecutionMode.VBA);
                        UpdateVbaCodePanelVisibility();
                        AppendToChatHistory("系统", "✅ VBA模式已启用！", Color.Green);
                    }
                    else
                    {
                        MessageBox.Show("VBA环境仍不可用，请参考诊断报告中的建议。", "检查结果", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                };

                var closeButton = new Button
                {
                    Text = "关闭",
                    Size = new Size(80, 30),
                    Location = new Point(190, 5)
                };
                closeButton.Click += (s, e) => diagForm.Close();

                buttonPanel.Controls.AddRange(new Control[] { copyButton, refreshButton, closeButton });
                diagForm.Controls.AddRange(new Control[] { textBox, buttonPanel });

                diagForm.ShowDialog(this);
                
                AppendToChatHistory("系统", "📋 VBA诊断完成，请查看详细报告", Color.Blue);
            }
            catch (Exception ex)
            {
                AppendToChatHistory("系统", $"❌ 诊断过程出错: {ex.Message}", Color.Red);
            }
        }

        #endregion

        #region 事件处理方法

        /// <summary>
        /// 刷新VBA状态按钮点击事件
        /// </summary>
        private void btnRefreshVbaStatus_Click(object sender, EventArgs e)
        {
            try
            {
                AppendToChatHistory("系统", "🔄 正在刷新VBA状态...", Color.Blue);
                CheckAndUpdateVbaStatus();
            }
            catch (Exception ex)
            {
                AppendToChatHistory("系统", $"❌ 刷新VBA状态失败: {ex.Message}", Color.Red);
            }
        }

        /// <summary>
        /// VSTO模式选择事件
        /// </summary>
        private void rbVstoMode_CheckedChanged(object sender, EventArgs e)
        {
            if (rbVstoMode.Checked)
            {
                ExecutionModeManager.SwitchMode(ExecutionMode.VSTO);
                UpdateVbaCodePanelVisibility();
                AppendToChatHistory("系统", "已切换到VSTO模式", Color.Blue);
            }
        }

        /// <summary>
        /// VBA模式选择事件
        /// </summary>
        private void rbVbaMode_CheckedChanged(object sender, EventArgs e)
        {
            if (rbVbaMode.Checked)
            {
                // 首次点击时进行延迟初始化和检查
                try
                {
                    AppendToChatHistory("系统", "🔄 正在检查VBA环境...", Color.Blue);
                    
                    // 延迟初始化ExecutionModeManager
                    if (!ExecutionModeManager.IsInitialized)
                    {
                        ExecutionModeManager.Initialize();
                    }
                    
                    // 检查VBA是否可用
                    if (ExecutionModeManager.IsVbaEnabled)
                    {
                        rbVbaMode.Enabled = true;
                        rbVbaMode.Text = "VBA";
                        ExecutionModeManager.SwitchMode(ExecutionMode.VBA);
                        UpdateVbaCodePanelVisibility();
                        AppendToChatHistory("系统", "✅ 已切换到VBA模式", Color.Green);
                        AppendToChatHistory("系统", "💡 在VBA模式下，AI将生成VBA代码供您执行", Color.Gray);
                        return;
                    }
                }
                catch (Exception ex)
                {
                    AppendToChatHistory("系统", $"❌ VBA环境检查失败: {ex.Message}", Color.Red);
                    System.Diagnostics.Debug.WriteLine($"VBA check failed: {ex}");
                }
                
                // VBA不可用，回退到VSTO模式
                rbVstoMode.Checked = true;
                
                // 显示VBA诊断和权限设置选项
                var diagResult = MessageBox.Show(
                    "⚠️ VBA模式不可用，请检查VBE访问权限设置。\n\n" +
                    "点击'是'进行详细诊断\n" +
                    "点击'否'查看设置指导\n" +
                    "点击'取消'返回",
                    "VBA模式不可用",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Warning);
                
                if (diagResult == DialogResult.Yes)
                {
                    // 运行VBA诊断
                    ShowVbaDiagnostics();
                }
                else if (diagResult == DialogResult.No)
                {
                    // 显示VBE权限设置对话框
                    var result = VbePermissionDialog.ShowVbePermissionDialog(this);
                    
                    if (result == DialogResult.OK)
                    {
                        // 用户已成功设置权限，重新尝试切换到VBA模式
                        try
                        {
                            ExecutionModeManager.RefreshVbaStatus();
                            if (ExecutionModeManager.IsVbaEnabled)
                            {
                                rbVbaMode.Checked = true;
                                rbVbaMode.Enabled = true;
                                rbVbaMode.Text = "VBA";
                                ExecutionModeManager.SwitchMode(ExecutionMode.VBA);
                                UpdateVbaCodePanelVisibility();
                                AppendToChatHistory("系统", "✅ VBA模式已启用！", Color.Green);
                                AppendToChatHistory("系统", "💡 在VBA模式下，AI将生成VBA代码供您执行", Color.Gray);
                            }
                            else
                            {
                                AppendToChatHistory("系统", "❌ VBA模式仍不可用，请尝试重启Excel", Color.Red);
                            }
                        }
                        catch (Exception ex)
                        {
                            AppendToChatHistory("系统", $"❌ VBA状态刷新失败: {ex.Message}", Color.Red);
                            System.Diagnostics.Debug.WriteLine($"VBA refresh failed: {ex}");
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 仅聊天模式选择事件
        /// </summary>
        private void rbChatOnlyMode_CheckedChanged(object sender, EventArgs e)
        {
            if (rbChatOnlyMode.Checked)
            {
                ExecutionModeManager.SwitchMode(ExecutionMode.ChatOnly);
                UpdateVbaCodePanelVisibility();
                AppendToChatHistory("系统", "已切换到仅聊天模式", Color.Blue);
                AppendToChatHistory("系统", "💬 在此模式下，AI只会回答问题，不会执行任何操作", Color.Gray);
            }
        }

        /// <summary>
        /// 执行VBA代码事件
        /// </summary>
        private async void btnExecuteVba_Click(object sender, EventArgs e)
        {
            if (_currentVbaResult == null || !_currentVbaResult.Success)
            {
                MessageBox.Show("没有可执行的VBA代码", "执行VBA", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                // 确认执行
                var confirmResult = MessageBox.Show(
                    $"确定要执行以下VBA代码吗？\n\n宏名称: {_currentVbaResult.MacroName}\n描述: {_currentVbaResult.Description}\n风险级别: {GetRiskLevelDescription(_currentVbaResult.RiskLevel)}",
                    "确认执行VBA代码",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (confirmResult != DialogResult.Yes)
                    return;

                btnExecuteVba.Enabled = false;
                AppendToChatHistory("系统", "正在执行VBA代码...", Color.Blue);

                // 执行VBA代码
                var executionResult = await _vbaInjectionEngine.InjectAndExecuteAsync(
                    _currentVbaResult.MacroName, 
                    _currentVbaResult.VbaCode, 
                    _currentUserRequest);

                // 移除"正在执行"提示
                RemoveLastChatHistoryLine();

                if (executionResult.Success)
                {
                    AppendToChatHistory("系统", $"✅ VBA代码执行成功 (耗时: {executionResult.ExecutionTimeMs}ms)", Color.Green);
                    if (executionResult.Result != null)
                    {
                        AppendToChatHistory("系统", $"执行结果: {executionResult.Result}", Color.Blue);
                    }
                }
                else
                {
                    AppendToChatHistory("系统", $"❌ VBA代码执行失败: {executionResult.ErrorMessage}", Color.Red);
                }
            }
            catch (Exception ex)
            {
                RemoveLastChatHistoryLine();
                AppendToChatHistory("系统", $"❌ VBA执行异常: {ex.Message}", Color.Red);
            }
            finally
            {
                btnExecuteVba.Enabled = true;
            }
        }

        /// <summary>
        /// 复制VBA代码事件
        /// </summary>
        private void btnCopyVbaCode_Click(object sender, EventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(rtbVbaCode.Text))
                {
                    Clipboard.SetText(rtbVbaCode.Text);
                    AppendToChatHistory("系统", "✅ VBA代码已复制到剪贴板", Color.Green);
                }
                else
                {
                    MessageBox.Show("没有可复制的VBA代码", "复制代码", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                AppendToChatHistory("系统", $"❌ 复制失败: {ex.Message}", Color.Red);
            }
        }

        /// <summary>
        /// 切换VBA代码显示事件
        /// </summary>
        private void btnToggleVbaCode_Click(object sender, EventArgs e)
        {
            if (_vbaCodeExpanded)
            {
                CollapseVbaCode();
            }
            else
            {
                ExpandVbaCode();
            }
        }

        #endregion
    }
}