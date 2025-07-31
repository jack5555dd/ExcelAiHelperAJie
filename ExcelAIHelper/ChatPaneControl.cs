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
        
        private string _currentUserRequest;
        
        public ChatPaneControl()
        {
            InitializeComponent();
            InitializeServices();
        }
        
        private void InitializeServices()
        {
            try
            {
                // Get Excel application from ThisAddIn
                Excel.Application excelApp = Globals.ThisAddIn.Application;
                
                // Initialize services
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
                
                AppendToChatHistory("系统", "Excel AI助手已启动，支持自然语言操作Excel表格。", Color.Green);
                AppendToChatHistory("系统", "💡 提示：您可以说\"在A1输入100\"、\"给选中区域设置红色背景\"等。", Color.Gray);
                
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
            catch (Exception ex)
            {
                AppendToChatHistory("系统", $"❌ 初始化失败: {ex.Message}", Color.Red);
                System.Diagnostics.Debug.WriteLine($"Service initialization failed: {ex.Message}");
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
                
                // Show thinking indicator
                AppendToChatHistory("系统", "AI思考中...", Color.Gray);
                
                // For debugging: check if API key is set
                if (string.IsNullOrEmpty(Properties.Settings.Default.ApiKey))
                {
                    RemoveLastChatHistoryLine();
                    AppendToChatHistory("系统", "请先设置API密钥。点击'API 设置'按钮进行配置。", Color.Red);
                    SetInputEnabled(true);
                    return;
                }
                
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
                    else
                    {
                        SetInputEnabled(true);
                    }
                }
            }
            catch (Exception ex)
            {
                HandleError(ex);
                SetInputEnabled(true);
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
    }
}