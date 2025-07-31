using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Threading.Tasks;
using ExcelAIHelper.Exceptions;

namespace ExcelAIHelper.Services
{
    /// <summary>
    /// Client for communicating with DeepSeek API
    /// </summary>
    public class DeepSeekClient : IDisposable
    {
        private readonly HttpClient _httpClient;
        private readonly string _apiKey;
        private readonly string _model;
        private readonly double _temperature;
        private readonly int _maxTokens;

        /// <summary>
        /// Creates a new instance of DeepSeekClient
        /// </summary>
        public DeepSeekClient()
        {
            // Enable TLS 1.2 and higher versions for SSL/TLS connections
            // This is required for modern HTTPS APIs
            System.Net.ServicePointManager.SecurityProtocol = 
                System.Net.SecurityProtocolType.Tls12 | 
                System.Net.SecurityProtocolType.Tls11 | 
                System.Net.SecurityProtocolType.Tls;
            
            // Disable SSL certificate validation for debugging (remove in production)
            // System.Net.ServicePointManager.ServerCertificateValidationCallback = 
            //     (sender, certificate, chain, sslPolicyErrors) => true;
            
            _httpClient = new HttpClient();
            _httpClient.Timeout = TimeSpan.FromSeconds(30); // Set timeout
            _apiKey = Properties.Settings.Default.ApiKey;
            
            // Validate API key
            if (string.IsNullOrEmpty(_apiKey))
            {
                throw new AiOperationException("API密钥未设置，请先在API设置中配置密钥");
            }
            
            _httpClient.BaseAddress = new Uri("https://api.deepseek.com");  // Default endpoint
            _httpClient.DefaultRequestHeaders.Accept.Clear();
            _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _apiKey);
            
            // Add User-Agent header for better compatibility
            _httpClient.DefaultRequestHeaders.UserAgent.ParseAdd("ExcelAIHelper/1.0");
            
            _model = "deepseek-chat";  // Default model
            _temperature = 0.7;  // Default temperature
            _maxTokens = 2048;  // Default max tokens
        }

        /// <summary>
        /// Sends a message to the AI and gets a response
        /// </summary>
        /// <param name="prompt">The user's message</param>
        /// <param name="systemPrompt">Optional system prompt to guide the AI</param>
        /// <returns>The AI's response</returns>
        public async Task<string> AskAsync(string prompt, string systemPrompt = null)
        {
            try
            {
                if (string.IsNullOrEmpty(prompt))
                {
                    throw new ArgumentException("用户输入不能为空", nameof(prompt));
                }

                var messages = new List<object>();
                
                // Add system message if provided
                if (!string.IsNullOrEmpty(systemPrompt))
                {
                    messages.Add(new
                    {
                        role = "system",
                        content = systemPrompt
                    });
                }
                
                // Add user message
                messages.Add(new
                {
                    role = "user",
                    content = prompt
                });
                
                var requestData = new
                {
                    model = _model,
                    messages,
                    temperature = _temperature,
                    max_tokens = _maxTokens
                };
                
                var jsonContent = JsonConvert.SerializeObject(requestData);
                var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");
                
                System.Diagnostics.Debug.WriteLine($"Sending request to DeepSeek API: {jsonContent}");
                
                var response = await _httpClient.PostAsync("/v1/chat/completions", content);
                
                var jsonResponse = await response.Content.ReadAsStringAsync();
                System.Diagnostics.Debug.WriteLine($"Received response: {jsonResponse}");
                
                if (!response.IsSuccessStatusCode)
                {
                    throw new AiOperationException($"API调用失败: {response.StatusCode} - {jsonResponse}");
                }
                
                var responseObject = JObject.Parse(jsonResponse);
                
                // Check if response has the expected structure
                if (responseObject["choices"] == null || !responseObject["choices"].HasValues)
                {
                    throw new AiOperationException("API响应格式异常：缺少choices字段");
                }
                
                var content_text = responseObject["choices"][0]["message"]["content"]?.ToString();
                if (string.IsNullOrEmpty(content_text))
                {
                    throw new AiOperationException("API响应为空");
                }
                
                return content_text;
            }
            catch (HttpRequestException ex)
            {
                System.Diagnostics.Debug.WriteLine($"HTTP Error in DeepSeekClient: {ex.Message}");
                
                // Provide more specific error messages based on the exception
                string errorMessage = "网络连接失败";
                if (ex.Message.Contains("401"))
                {
                    errorMessage = "API密钥无效，请检查API密钥是否正确";
                }
                else if (ex.Message.Contains("403"))
                {
                    errorMessage = "API访问被拒绝，请检查API密钥权限";
                }
                else if (ex.Message.Contains("429"))
                {
                    errorMessage = "API调用频率超限，请稍后重试";
                }
                else if (ex.Message.Contains("500"))
                {
                    errorMessage = "API服务器内部错误，请稍后重试";
                }
                else if (ex.Message.Contains("timeout") || ex.Message.Contains("超时"))
                {
                    errorMessage = "网络连接超时，请检查网络连接";
                }
                else if (ex.Message.Contains("SSL") || ex.Message.Contains("certificate"))
                {
                    errorMessage = "SSL证书验证失败，请检查网络设置";
                }
                
                throw new AiOperationException($"{errorMessage}: {ex.Message}", ex);
            }
            catch (TaskCanceledException ex)
            {
                System.Diagnostics.Debug.WriteLine($"Timeout Error in DeepSeekClient: {ex.Message}");
                throw new AiOperationException("请求超时，请稍后重试", ex);
            }
            catch (JsonException ex)
            {
                System.Diagnostics.Debug.WriteLine($"JSON Error in DeepSeekClient: {ex.Message}");
                throw new AiOperationException("API响应解析失败", ex);
            }
            catch (AiOperationException)
            {
                // Re-throw our custom exceptions
                throw;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Unexpected Error in DeepSeekClient: {ex.Message}");
                throw new AiOperationException($"AI服务调用失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Sends a message to the AI and gets a structured JSON response
        /// </summary>
        /// <typeparam name="T">The type to deserialize the response to</typeparam>
        /// <param name="prompt">The user's message</param>
        /// <param name="systemPrompt">Optional system prompt to guide the AI</param>
        /// <returns>The deserialized response</returns>
        public async Task<T> AskForStructuredResponseAsync<T>(string prompt, string systemPrompt = null)
        {
            try
            {
                string jsonResponse = await AskAsync(prompt, systemPrompt);
                return JsonConvert.DeserializeObject<T>(jsonResponse);
            }
            catch (JsonException ex)
            {
                throw new AiOperationException("Failed to parse AI response as valid JSON", ex);
            }
        }

        /// <summary>
        /// Tests basic network connectivity to the API endpoint
        /// </summary>
        /// <returns>True if basic connectivity works</returns>
        public async Task<(bool Success, string Message)> TestBasicConnectivityAsync()
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("Testing basic connectivity to DeepSeek API...");
                
                // Enable TLS 1.2 for this test as well
                System.Net.ServicePointManager.SecurityProtocol = 
                    System.Net.SecurityProtocolType.Tls12 | 
                    System.Net.SecurityProtocolType.Tls11 | 
                    System.Net.SecurityProtocolType.Tls;
                
                // Test basic HTTP connectivity without authentication
                using (var testClient = new HttpClient())
                {
                    testClient.Timeout = TimeSpan.FromSeconds(10);
                    testClient.DefaultRequestHeaders.Add("User-Agent", "ExcelAIHelper/1.0");
                    
                    var response = await testClient.GetAsync("https://api.deepseek.com");
                    
                    System.Diagnostics.Debug.WriteLine($"Basic connectivity test status: {response.StatusCode}");
                    
                    if (response.StatusCode == System.Net.HttpStatusCode.NotFound || 
                        response.StatusCode == System.Net.HttpStatusCode.MethodNotAllowed ||
                        response.StatusCode == System.Net.HttpStatusCode.Unauthorized)
                    {
                        // These are expected responses that indicate connectivity is working
                        return (true, "网络连接正常，SSL/TLS握手成功");
                    }
                    else if (response.IsSuccessStatusCode)
                    {
                        return (true, "网络连接正常，SSL/TLS握手成功");
                    }
                    else
                    {
                        return (false, $"服务器响应异常: {response.StatusCode}");
                    }
                }
            }
            catch (HttpRequestException ex)
            {
                if (ex.Message.Contains("SSL") || ex.Message.Contains("TLS") || ex.Message.Contains("安全通道"))
                {
                    return (false, $"SSL/TLS连接失败: {ex.Message}\n建议: 检查系统时间、更新.NET Framework或联系IT部门");
                }
                return (false, $"网络连接失败: {ex.Message}");
            }
            catch (TaskCanceledException)
            {
                return (false, "网络连接超时");
            }
            catch (Exception ex)
            {
                return (false, $"连接测试异常: {ex.Message}");
            }
        }

        /// <summary>
        /// Tests the API connection and authentication
        /// </summary>
        /// <returns>True if connection is successful, false otherwise</returns>
        public async Task<(bool Success, string Message)> TestConnectionAsync()
        {
            try
            {
                // Simple test request with minimal content
                var testMessages = new List<object>
                {
                    new
                    {
                        role = "user",
                        content = "Hello"
                    }
                };

                var requestData = new
                {
                    model = _model,
                    messages = testMessages,
                    temperature = 0.1,
                    max_tokens = 10
                };

                var jsonContent = JsonConvert.SerializeObject(requestData);
                var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                System.Diagnostics.Debug.WriteLine($"Testing API connection...");

                var response = await _httpClient.PostAsync("/v1/chat/completions", content);
                var jsonResponse = await response.Content.ReadAsStringAsync();

                System.Diagnostics.Debug.WriteLine($"Test response status: {response.StatusCode}");
                System.Diagnostics.Debug.WriteLine($"Test response: {jsonResponse}");

                if (response.IsSuccessStatusCode)
                {
                    var responseObject = JObject.Parse(jsonResponse);
                    if (responseObject["choices"] != null && responseObject["choices"].HasValues)
                    {
                        return (true, "API连接测试成功");
                    }
                    else
                    {
                        return (false, "API响应格式异常");
                    }
                }
                else
                {
                    // Parse error response for more details
                    try
                    {
                        var errorObject = JObject.Parse(jsonResponse);
                        var errorMessage = errorObject["error"]?["message"]?.ToString() ?? "未知错误";
                        var errorType = errorObject["error"]?["type"]?.ToString() ?? "unknown";
                        
                        return (false, $"API调用失败 ({response.StatusCode}): {errorType} - {errorMessage}");
                    }
                    catch
                    {
                        return (false, $"API调用失败 ({response.StatusCode}): {jsonResponse}");
                    }
                }
            }
            catch (HttpRequestException ex)
            {
                System.Diagnostics.Debug.WriteLine($"HTTP Error in TestConnection: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Inner Exception: {ex.InnerException?.Message}");
                
                string errorMessage = "网络连接失败";
                string suggestion = "";
                
                if (ex.Message.Contains("401") || ex.Message.Contains("Unauthorized"))
                {
                    errorMessage = "API密钥无效";
                    suggestion = "请检查API密钥是否正确";
                }
                else if (ex.Message.Contains("403") || ex.Message.Contains("Forbidden"))
                {
                    errorMessage = "API访问被拒绝";
                    suggestion = "请检查API密钥权限或账户状态";
                }
                else if (ex.Message.Contains("404") || ex.Message.Contains("Not Found"))
                {
                    errorMessage = "API端点不存在";
                    suggestion = "请检查API端点URL是否正确";
                }
                else if (ex.Message.Contains("429") || ex.Message.Contains("Too Many Requests"))
                {
                    errorMessage = "API调用频率超限";
                    suggestion = "请稍后重试";
                }
                else if (ex.Message.Contains("500") || ex.Message.Contains("Internal Server Error"))
                {
                    errorMessage = "API服务器内部错误";
                    suggestion = "请稍后重试或联系API服务提供商";
                }
                else if (ex.Message.Contains("timeout") || ex.Message.Contains("超时"))
                {
                    errorMessage = "网络连接超时";
                    suggestion = "请检查网络连接或尝试使用代理";
                }
                else if (ex.Message.Contains("SSL") || ex.Message.Contains("certificate"))
                {
                    errorMessage = "SSL证书验证失败";
                    suggestion = "请检查网络代理设置或系统时间";
                }
                else if (ex.Message.Contains("Name or service not known") || ex.Message.Contains("No such host"))
                {
                    errorMessage = "无法解析域名";
                    suggestion = "请检查DNS设置或网络连接";
                }
                else if (ex.Message.Contains("Connection refused") || ex.Message.Contains("连接被拒绝"))
                {
                    errorMessage = "连接被拒绝";
                    suggestion = "请检查防火墙设置或网络代理";
                }
                else if (ex.InnerException != null)
                {
                    errorMessage = $"网络连接失败: {ex.InnerException.Message}";
                    suggestion = "请检查网络连接、防火墙或代理设置";
                }
                
                return (false, $"{errorMessage}\n建议: {suggestion}\n详细信息: {ex.Message}");
            }
            catch (TaskCanceledException ex)
            {
                return (false, $"请求超时\n建议: 请检查网络连接或稍后重试\n详细信息: {ex.Message}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Unexpected Error in TestConnection: {ex.Message}");
                return (false, $"连接测试失败\n详细信息: {ex.Message}");
            }
        }

        /// <summary>
        /// Disposes the HTTP client
        /// </summary>
        public void Dispose()
        {
            _httpClient?.Dispose();
        }
    }
}