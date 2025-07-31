using System;
using System.Net.NetworkInformation;
using System.Threading.Tasks;

namespace ExcelAIHelper.Services
{
    /// <summary>
    /// Network diagnostics utilities
    /// </summary>
    public static class NetworkDiagnostics
    {
        /// <summary>
        /// Tests basic internet connectivity
        /// </summary>
        /// <returns>True if internet is accessible</returns>
        public static async Task<(bool Success, string Message)> TestInternetConnectivityAsync()
        {
            try
            {
                using (var ping = new Ping())
                {
                    // Test connectivity to a reliable public DNS server
                    var reply = await ping.SendPingAsync("8.8.8.8", 5000);
                    
                    if (reply.Status == IPStatus.Success)
                    {
                        return (true, "互联网连接正常");
                    }
                    else
                    {
                        return (false, $"网络连接异常: {reply.Status}");
                    }
                }
            }
            catch (Exception ex)
            {
                return (false, $"网络测试失败: {ex.Message}");
            }
        }

        /// <summary>
        /// Tests DNS resolution for DeepSeek API
        /// </summary>
        /// <returns>True if DNS resolution works</returns>
        public static async Task<(bool Success, string Message)> TestDnsResolutionAsync()
        {
            try
            {
                using (var ping = new Ping())
                {
                    var reply = await ping.SendPingAsync("api.deepseek.com", 5000);
                    
                    if (reply.Status == IPStatus.Success)
                    {
                        return (true, "DNS解析正常");
                    }
                    else if (reply.Status == IPStatus.TimedOut)
                    {
                        return (false, "DNS解析超时，可能是网络或防火墙问题");
                    }
                    else
                    {
                        return (false, $"DNS解析失败: {reply.Status}");
                    }
                }
            }
            catch (Exception ex)
            {
                return (false, $"DNS测试失败: {ex.Message}");
            }
        }

        /// <summary>
        /// Tests SSL/TLS configuration
        /// </summary>
        /// <returns>SSL/TLS configuration status</returns>
        public static (bool Success, string Message) TestSslTlsConfiguration()
        {
            try
            {
                var currentProtocol = System.Net.ServicePointManager.SecurityProtocol;
                var supportedProtocols = new System.Text.StringBuilder();
                
                if ((currentProtocol & System.Net.SecurityProtocolType.Tls12) == System.Net.SecurityProtocolType.Tls12)
                    supportedProtocols.Append("TLS 1.2 ");
                if ((currentProtocol & System.Net.SecurityProtocolType.Tls11) == System.Net.SecurityProtocolType.Tls11)
                    supportedProtocols.Append("TLS 1.1 ");
                if ((currentProtocol & System.Net.SecurityProtocolType.Tls) == System.Net.SecurityProtocolType.Tls)
                    supportedProtocols.Append("TLS 1.0 ");
                
                if (supportedProtocols.Length == 0)
                {
                    return (false, "未启用任何TLS协议");
                }
                
                bool hasTls12 = (currentProtocol & System.Net.SecurityProtocolType.Tls12) == System.Net.SecurityProtocolType.Tls12;
                
                return (hasTls12, $"支持的协议: {supportedProtocols.ToString().Trim()}" + 
                                 (hasTls12 ? " (推荐)" : " (建议启用TLS 1.2)"));
            }
            catch (Exception ex)
            {
                return (false, $"SSL/TLS配置检查失败: {ex.Message}");
            }
        }

        /// <summary>
        /// Gets comprehensive network diagnostic information
        /// </summary>
        /// <returns>Diagnostic information</returns>
        public static async Task<string> GetNetworkDiagnosticsAsync()
        {
            var diagnostics = new System.Text.StringBuilder();
            
            diagnostics.AppendLine("=== 网络诊断信息 ===");
            diagnostics.AppendLine($"时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            diagnostics.AppendLine();
            
            // Test SSL/TLS configuration
            var (sslSuccess, sslMessage) = TestSslTlsConfiguration();
            diagnostics.AppendLine($"SSL/TLS配置: {(sslSuccess ? "✓" : "✗")} {sslMessage}");
            
            // Test internet connectivity
            var (internetSuccess, internetMessage) = await TestInternetConnectivityAsync();
            diagnostics.AppendLine($"互联网连接: {(internetSuccess ? "✓" : "✗")} {internetMessage}");
            
            // Test DNS resolution
            var (dnsSuccess, dnsMessage) = await TestDnsResolutionAsync();
            diagnostics.AppendLine($"DNS解析: {(dnsSuccess ? "✓" : "✗")} {dnsMessage}");
            
            // .NET Framework version
            try
            {
                var frameworkVersion = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription;
                diagnostics.AppendLine($".NET Framework: {frameworkVersion}");
            }
            catch
            {
                diagnostics.AppendLine($".NET Framework: {Environment.Version}");
            }
            
            // System time (important for SSL certificates)
            diagnostics.AppendLine($"系统时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss} (UTC: {DateTime.UtcNow:yyyy-MM-dd HH:mm:ss})");
            
            // Network interface information
            try
            {
                bool hasActiveConnection = false;
                foreach (NetworkInterface ni in NetworkInterface.GetAllNetworkInterfaces())
                {
                    if (ni.OperationalStatus == OperationalStatus.Up && 
                        ni.NetworkInterfaceType != NetworkInterfaceType.Loopback)
                    {
                        hasActiveConnection = true;
                        diagnostics.AppendLine($"网络接口: {ni.Name} ({ni.NetworkInterfaceType}) - 活动");
                        break;
                    }
                }
                
                if (!hasActiveConnection)
                {
                    diagnostics.AppendLine("网络接口: 未找到活动的网络连接");
                }
            }
            catch (Exception ex)
            {
                diagnostics.AppendLine($"网络接口检查失败: {ex.Message}");
            }
            
            return diagnostics.ToString();
        }
    }
}