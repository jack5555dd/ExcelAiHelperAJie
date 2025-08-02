using System;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace ExcelAIHelper.Services
{
    /// <summary>
    /// 应用程序兼容性管理器 - 检测并适配Microsoft Excel和WPS表格
    /// </summary>
    public static class ApplicationCompatibilityManager
    {
        /// <summary>
        /// 当前运行的Office应用程序类型
        /// </summary>
        public enum OfficeApplicationType
        {
            Unknown,
            MicrosoftExcel,
            WpsSpreadsheets
        }

        private static OfficeApplicationType? _currentApplicationType = null;

        /// <summary>
        /// 获取当前运行的Office应用程序类型
        /// </summary>
        public static OfficeApplicationType CurrentApplicationType
        {
            get
            {
                if (_currentApplicationType == null)
                {
                    _currentApplicationType = DetectApplicationType();
                }
                return _currentApplicationType.Value;
            }
        }

        /// <summary>
        /// 检测当前运行的Office应用程序类型
        /// </summary>
        /// <returns>应用程序类型</returns>
        private static OfficeApplicationType DetectApplicationType()
        {
            try
            {
                // 方法1: 通过进程名检测
                var currentProcess = Process.GetCurrentProcess();
                var processName = currentProcess.ProcessName.ToLower();
                
                Debug.WriteLine($"Current process name: {processName}");
                
                if (processName.Contains("excel"))
                {
                    Debug.WriteLine("Detected Microsoft Excel by process name");
                    return OfficeApplicationType.MicrosoftExcel;
                }
                else if (processName.Contains("et") || processName.Contains("wps"))
                {
                    Debug.WriteLine("Detected WPS Spreadsheets by process name");
                    return OfficeApplicationType.WpsSpreadsheets;
                }

                // 方法2: 通过COM对象检测
                try
                {
                    // 尝试获取Excel Application对象
                    var excelApp = Marshal.GetActiveObject("Excel.Application");
                    if (excelApp != null)
                    {
                        var appName = excelApp.GetType().InvokeMember("Name", 
                            System.Reflection.BindingFlags.GetProperty, null, excelApp, null) as string;
                        
                        Debug.WriteLine($"COM Application Name: {appName}");
                        
                        if (appName != null && appName.ToLower().Contains("excel"))
                        {
                            Debug.WriteLine("Detected Microsoft Excel by COM object");
                            return OfficeApplicationType.MicrosoftExcel;
                        }
                        else if (appName != null && (appName.ToLower().Contains("et") || appName.ToLower().Contains("wps")))
                        {
                            Debug.WriteLine("Detected WPS Spreadsheets by COM object");
                            return OfficeApplicationType.WpsSpreadsheets;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"COM detection failed: {ex.Message}");
                }

                // 方法3: 通过模块路径检测
                try
                {
                    var mainModule = currentProcess.MainModule;
                    if (mainModule != null)
                    {
                        var fileName = mainModule.FileName.ToLower();
                        Debug.WriteLine($"Main module file name: {fileName}");
                        
                        if (fileName.Contains("excel.exe"))
                        {
                            Debug.WriteLine("Detected Microsoft Excel by module path");
                            return OfficeApplicationType.MicrosoftExcel;
                        }
                        else if (fileName.Contains("et.exe") || fileName.Contains("wps"))
                        {
                            Debug.WriteLine("Detected WPS Spreadsheets by module path");
                            return OfficeApplicationType.WpsSpreadsheets;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Module path detection failed: {ex.Message}");
                }

                Debug.WriteLine("Could not detect application type, defaulting to Microsoft Excel");
                return OfficeApplicationType.MicrosoftExcel;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Application detection failed: {ex.Message}");
                return OfficeApplicationType.Unknown;
            }
        }

        /// <summary>
        /// 检查是否为Microsoft Excel
        /// </summary>
        public static bool IsMicrosoftExcel => CurrentApplicationType == OfficeApplicationType.MicrosoftExcel;

        /// <summary>
        /// 检查是否为WPS表格
        /// </summary>
        public static bool IsWpsSpreadsheets => CurrentApplicationType == OfficeApplicationType.WpsSpreadsheets;

        /// <summary>
        /// 获取应用程序显示名称
        /// </summary>
        public static string GetApplicationDisplayName()
        {
            switch (CurrentApplicationType)
            {
                case OfficeApplicationType.MicrosoftExcel:
                    return "Microsoft Excel";
                case OfficeApplicationType.WpsSpreadsheets:
                    return "WPS表格";
                default:
                    return "未知应用程序";
            }
        }

        /// <summary>
        /// 重置检测结果（用于测试或强制重新检测）
        /// </summary>
        public static void ResetDetection()
        {
            _currentApplicationType = null;
        }
    }
}