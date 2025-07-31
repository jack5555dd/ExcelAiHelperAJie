using System;
using System.Threading.Tasks;
using ExcelAIHelper.Exceptions;
using ExcelAIHelper.Models;
using ExcelAIHelper.Services;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAIHelper.Services
{
    /// <summary>
    /// Dispatches operations from user requests to Excel operations
    /// </summary>
    public class OperationDispatcher
    {
        private readonly DeepSeekClient _aiClient;
        private readonly PromptBuilder _promptBuilder;
        private readonly InstructionParser _instructionParser;
        private readonly ExcelOperationEngine _operationEngine;
        
        /// <summary>
        /// Creates a new OperationDispatcher
        /// </summary>
        public OperationDispatcher(
            DeepSeekClient aiClient,
            PromptBuilder promptBuilder,
            InstructionParser instructionParser,
            Excel.Application excelApp)
        {
            _aiClient = aiClient ?? throw new ArgumentNullException(nameof(aiClient));
            _promptBuilder = promptBuilder ?? throw new ArgumentNullException(nameof(promptBuilder));
            _instructionParser = instructionParser ?? throw new ArgumentNullException(nameof(instructionParser));
            _operationEngine = new ExcelOperationEngine(excelApp);
        }
        
        /// <summary>
        /// Applies user request to Excel
        /// </summary>
        /// <param name="userRequest">The user's request</param>
        /// <param name="dryRun">If true, only simulate the operation</param>
        /// <returns>Operation result</returns>
        public async Task<OperationResult> ApplyAsync(string userRequest, bool dryRun = false)
        {
            try
            {
                // Build prompts
                string systemPrompt = await _promptBuilder.BuildSystemPromptAsync();
                string userPrompt = await _promptBuilder.BuildUserPromptAsync(userRequest);
                
                // Get AI response
                string aiResponse = await _aiClient.AskAsync(systemPrompt, userPrompt);
                
                // Parse instructions
                InstructionSet instructionSet = await _instructionParser.ParseAsync(aiResponse);
                
                // Execute instructions
                string result = await _operationEngine.ExecuteInstructionsAsync(instructionSet, dryRun);
                
                return OperationResult.CreateSuccess(result);
            }
            catch (AiFormatException formatEx)
            {
                return OperationResult.CreateFailure(formatEx.GetUserFriendlyMessage(), "协议格式错误", formatEx.CanRetry);
            }
            catch (AiOperationException opEx)
            {
                return OperationResult.CreateFailure(opEx.Message, "操作执行错误", false);
            }
            catch (Exception ex)
            {
                return OperationResult.CreateFailure($"系统错误: {ex.Message}", "系统错误", false);
            }
        }
        
        /// <summary>
        /// Retries an operation after a failure
        /// </summary>
        /// <param name="userRequest">The original user request</param>
        /// <param name="errorType">The type of error that occurred</param>
        /// <param name="dryRun">If true, only simulate the operation</param>
        /// <returns>Retry operation result</returns>
        public async Task<OperationResult> RetryAsync(string userRequest, string errorType, bool dryRun = false)
        {
            try
            {
                // For format errors, we can try again with the same request
                if (errorType == "协议格式错误")
                {
                    return await ApplyAsync(userRequest, dryRun);
                }
                
                return OperationResult.CreateFailure("不支持的重试类型", errorType, false);
            }
            catch (Exception ex)
            {
                return OperationResult.CreateFailure($"重试失败: {ex.Message}", "重试错误", false);
            }
        }
    }
    
    /// <summary>
    /// Operation result class
    /// </summary>
    public class OperationResult
    {
        /// <summary>
        /// Whether the operation was successful
        /// </summary>
        public bool Success { get; private set; }
        
        /// <summary>
        /// The result message
        /// </summary>
        public string Message { get; private set; }
        
        /// <summary>
        /// Error message if operation failed
        /// </summary>
        public string ErrorMessage { get; private set; }
        
        /// <summary>
        /// Type of error that occurred
        /// </summary>
        public string ErrorType { get; private set; }
        
        /// <summary>
        /// Whether the operation can be retried
        /// </summary>
        public bool CanRetry { get; private set; }
        
        private OperationResult(bool success, string message, string errorMessage = null, string errorType = null, bool canRetry = false)
        {
            Success = success;
            Message = message;
            ErrorMessage = errorMessage;
            ErrorType = errorType;
            CanRetry = canRetry;
        }
        
        /// <summary>
        /// Creates a successful operation result
        /// </summary>
        public static OperationResult CreateSuccess(string message)
        {
            return new OperationResult(true, message);
        }
        
        /// <summary>
        /// Creates a failed operation result
        /// </summary>
        public static OperationResult CreateFailure(string errorMessage, string errorType = null, bool canRetry = false)
        {
            return new OperationResult(false, null, errorMessage, errorType, canRetry);
        }
        
        /// <summary>
        /// Gets user-friendly message
        /// </summary>
        public string GetUserMessage()
        {
            return Success ? Message : ErrorMessage;
        }
    }
}