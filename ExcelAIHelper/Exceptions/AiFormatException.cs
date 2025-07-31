using System;

namespace ExcelAIHelper.Exceptions
{
    /// <summary>
    /// AI响应格式异常，当AI返回的内容不符合协议要求时抛出
    /// </summary>
    public class AiFormatException : Exception
    {
        /// <summary>
        /// 原始AI响应内容
        /// </summary>
        public string OriginalResponse { get; }
        
        /// <summary>
        /// 协议违规详情
        /// </summary>
        public string ValidationDetails { get; }
        
        /// <summary>
        /// 是否可以重试
        /// </summary>
        public bool CanRetry { get; }
        
        /// <summary>
        /// 创建AI格式异常
        /// </summary>
        /// <param name="message">异常消息</param>
        public AiFormatException(string message) : base(message)
        {
            CanRetry = true;
        }
        
        /// <summary>
        /// 创建AI格式异常
        /// </summary>
        /// <param name="message">异常消息</param>
        /// <param name="originalResponse">原始AI响应</param>
        public AiFormatException(string message, string originalResponse) : base(message)
        {
            OriginalResponse = originalResponse;
            CanRetry = true;
        }
        
        /// <summary>
        /// 创建AI格式异常
        /// </summary>
        /// <param name="message">异常消息</param>
        /// <param name="originalResponse">原始AI响应</param>
        /// <param name="validationDetails">验证详情</param>
        public AiFormatException(string message, string originalResponse, string validationDetails) : base(message)
        {
            OriginalResponse = originalResponse;
            ValidationDetails = validationDetails;
            CanRetry = true;
        }
        
        /// <summary>
        /// 创建AI格式异常
        /// </summary>
        /// <param name="message">异常消息</param>
        /// <param name="originalResponse">原始AI响应</param>
        /// <param name="validationDetails">验证详情</param>
        /// <param name="canRetry">是否可以重试</param>
        public AiFormatException(string message, string originalResponse, string validationDetails, bool canRetry) : base(message)
        {
            OriginalResponse = originalResponse;
            ValidationDetails = validationDetails;
            CanRetry = canRetry;
        }
        
        /// <summary>
        /// 创建AI格式异常
        /// </summary>
        /// <param name="message">异常消息</param>
        /// <param name="innerException">内部异常</param>
        public AiFormatException(string message, Exception innerException) : base(message, innerException)
        {
            CanRetry = true;
        }
        
        /// <summary>
        /// 创建AI格式异常
        /// </summary>
        /// <param name="message">异常消息</param>
        /// <param name="originalResponse">原始AI响应</param>
        /// <param name="innerException">内部异常</param>
        public AiFormatException(string message, string originalResponse, Exception innerException) : base(message, innerException)
        {
            OriginalResponse = originalResponse;
            CanRetry = true;
        }
        
        /// <summary>
        /// 获取详细的错误信息
        /// </summary>
        /// <returns>包含所有错误详情的字符串</returns>
        public string GetDetailedErrorMessage()
        {
            var details = new System.Text.StringBuilder();
            details.AppendLine($"AI响应格式错误: {Message}");
            
            if (!string.IsNullOrEmpty(ValidationDetails))
            {
                details.AppendLine($"验证详情: {ValidationDetails}");
            }
            
            if (!string.IsNullOrEmpty(OriginalResponse))
            {
                details.AppendLine("原始响应:");
                details.AppendLine(OriginalResponse);
            }
            
            if (InnerException != null)
            {
                details.AppendLine($"内部异常: {InnerException.Message}");
            }
            
            details.AppendLine($"可重试: {(CanRetry ? "是" : "否")}");
            
            return details.ToString();
        }
        
        /// <summary>
        /// 获取用户友好的错误消息
        /// </summary>
        /// <returns>用户友好的错误描述</returns>
        public string GetUserFriendlyMessage()
        {
            if (!string.IsNullOrEmpty(ValidationDetails))
            {
                // 尝试提取具体的验证错误
                if (ValidationDetails.Contains("JSON格式错误"))
                {
                    return "AI返回的内容格式不正确，请重试。";
                }
                else if (ValidationDetails.Contains("协议违规"))
                {
                    return "AI返回的指令不符合规范，请重试或重新描述您的需求。";
                }
                else if (ValidationDetails.Contains("无效的单元格范围"))
                {
                    return "指定的单元格位置无效，请检查位置描述是否正确。";
                }
                else if (ValidationDetails.Contains("公式语法错误"))
                {
                    return "生成的Excel公式有语法错误，请重试。";
                }
            }
            
            return CanRetry ? 
                "AI响应格式有误，正在重试..." : 
                "AI响应处理失败，请重新发送您的请求。";
        }
        
        /// <summary>
        /// 创建协议违规异常
        /// </summary>
        /// <param name="originalResponse">原始响应</param>
        /// <param name="validationErrors">验证错误列表</param>
        /// <returns>格式异常实例</returns>
        public static AiFormatException CreateProtocolViolation(string originalResponse, string validationErrors)
        {
            return new AiFormatException(
                "AI响应违反了命令协议规范",
                originalResponse,
                validationErrors,
                true
            );
        }
        
        /// <summary>
        /// 创建JSON解析异常
        /// </summary>
        /// <param name="originalResponse">原始响应</param>
        /// <param name="parseException">解析异常</param>
        /// <returns>格式异常实例</returns>
        public static AiFormatException CreateJsonParseError(string originalResponse, Exception parseException)
        {
            return new AiFormatException(
                "AI返回的内容不是有效的JSON格式",
                originalResponse,
                parseException
            );
        }
        
        /// <summary>
        /// 创建业务规则违规异常
        /// </summary>
        /// <param name="originalResponse">原始响应</param>
        /// <param name="businessRuleError">业务规则错误</param>
        /// <returns>格式异常实例</returns>
        public static AiFormatException CreateBusinessRuleViolation(string originalResponse, string businessRuleError)
        {
            return new AiFormatException(
                "AI返回的指令违反了业务规则",
                originalResponse,
                businessRuleError,
                true
            );
        }
    }
}