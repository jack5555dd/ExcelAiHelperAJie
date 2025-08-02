using System;
using System.Diagnostics;

namespace ExcelAIHelper.Services
{
    /// <summary>
    /// Excel应用程序工厂 - 根据当前环境创建合适的Excel应用程序实例
    /// </summary>
    public static class ExcelApplicationFactory
    {
        private static IExcelApplication _currentInstance = null;

        /// <summary>
        /// 获取当前Excel应用程序实例
        /// </summary>
        /// <returns>Excel应用程序实例</returns>
        public static IExcelApplication GetCurrentApplication()
        {
            if (_currentInstance == null)
            {
                _currentInstance = CreateApplication();
            }
            return _currentInstance;
        }

        /// <summary>
        /// 创建Excel应用程序实例
        /// </summary>
        /// <returns>Excel应用程序实例</returns>
        private static IExcelApplication CreateApplication()
        {
            try
            {
                var appType = ApplicationCompatibilityManager.CurrentApplicationType;
                Debug.WriteLine($"Creating Excel application for: {appType}");

                switch (appType)
                {
                    case ApplicationCompatibilityManager.OfficeApplicationType.MicrosoftExcel:
                        Debug.WriteLine("Creating MicrosoftExcelApplication instance");
                        return new MicrosoftExcelApplication();

                    case ApplicationCompatibilityManager.OfficeApplicationType.WpsSpreadsheets:
                        Debug.WriteLine("Creating WpsExcelApplication instance");
                        return new WpsExcelApplication();

                    default:
                        Debug.WriteLine("Unknown application type, defaulting to MicrosoftExcelApplication");
                        return new MicrosoftExcelApplication();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to create Excel application: {ex.Message}");
                // 如果创建失败，尝试创建Microsoft Excel实例作为后备
                try
                {
                    Debug.WriteLine("Attempting fallback to MicrosoftExcelApplication");
                    return new MicrosoftExcelApplication();
                }
                catch (Exception fallbackEx)
                {
                    Debug.WriteLine($"Fallback also failed: {fallbackEx.Message}");
                    throw new InvalidOperationException($"无法创建Excel应用程序实例: {ex.Message}", ex);
                }
            }
        }

        /// <summary>
        /// 重置当前实例（用于测试或强制重新创建）
        /// </summary>
        public static void ResetInstance()
        {
            _currentInstance = null;
        }

        /// <summary>
        /// 获取当前应用程序的显示名称
        /// </summary>
        /// <returns>应用程序显示名称</returns>
        public static string GetCurrentApplicationName()
        {
            return ApplicationCompatibilityManager.GetApplicationDisplayName();
        }

        /// <summary>
        /// 检查当前是否为Microsoft Excel
        /// </summary>
        /// <returns>是否为Microsoft Excel</returns>
        public static bool IsMicrosoftExcel()
        {
            return ApplicationCompatibilityManager.IsMicrosoftExcel;
        }

        /// <summary>
        /// 检查当前是否为WPS表格
        /// </summary>
        /// <returns>是否为WPS表格</returns>
        public static bool IsWpsSpreadsheets()
        {
            return ApplicationCompatibilityManager.IsWpsSpreadsheets;
        }
    }
}