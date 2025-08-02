using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ExcelAIHelper.Services
{
    /// <summary>
    /// WPS表格应用程序实现
    /// </summary>
    public class WpsExcelApplication : IExcelApplication
    {
        private dynamic _application;
        private bool _isInitialized = false;

        public WpsExcelApplication()
        {
            try
            {
                InitializeWpsApplication();
                Debug.WriteLine("WpsExcelApplication initialized successfully");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"WpsExcelApplication initialization failed: {ex.Message}");
                throw;
            }
        }

        private void InitializeWpsApplication()
        {
            try
            {
                // 方法1: 尝试通过Globals获取（如果在VSTO环境中）
                try
                {
                    _application = Globals.ThisAddIn.Application;
                    _isInitialized = true;
                    Debug.WriteLine("WPS Application obtained from Globals.ThisAddIn");
                    return;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Failed to get WPS application from Globals: {ex.Message}");
                }

                // 方法2: 尝试通过COM获取活动的WPS应用程序
                try
                {
                    _application = Marshal.GetActiveObject("ET.Application");
                    _isInitialized = true;
                    Debug.WriteLine("WPS Application obtained from ET.Application COM object");
                    return;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Failed to get active ET.Application: {ex.Message}");
                }

                // 方法3: 尝试通过Excel兼容接口获取
                try
                {
                    _application = Marshal.GetActiveObject("Excel.Application");
                    // 验证这确实是WPS而不是Microsoft Excel
                    var appName = _application.Name;
                    if (appName != null && (appName.ToString().Contains("ET") || appName.ToString().Contains("WPS")))
                    {
                        _isInitialized = true;
                        Debug.WriteLine("WPS Application obtained from Excel.Application COM object (WPS compatibility mode)");
                        return;
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Failed to get WPS through Excel compatibility: {ex.Message}");
                }

                throw new InvalidOperationException("无法初始化WPS表格应用程序对象");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"InitializeWpsApplication failed: {ex.Message}");
                throw;
            }
        }

        public object Application => _application;

        /// <summary>
        /// 获取WPS应用程序是否已成功初始化
        /// </summary>
        public bool IsInitialized => _isInitialized;

        public object ActiveWorkbook
        {
            get
            {
                try
                {
                    return _application?.ActiveWorkbook;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Failed to get ActiveWorkbook: {ex.Message}");
                    return null;
                }
            }
        }

        public object ActiveSheet
        {
            get
            {
                try
                {
                    return _application?.ActiveSheet;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Failed to get ActiveSheet: {ex.Message}");
                    return null;
                }
            }
        }

        public object Selection
        {
            get
            {
                try
                {
                    return _application?.Selection;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Failed to get Selection: {ex.Message}");
                    return null;
                }
            }
        }

        public object GetRange(string address)
        {
            try
            {
                var activeSheet = _application?.ActiveSheet;
                if (activeSheet != null)
                {
                    return activeSheet.Range[address];
                }
                return null;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"GetRange failed for address '{address}': {ex.Message}");
                return null;
            }
        }

        public object GetCellValue(object range)
        {
            try
            {
                if (range != null)
                {
                    // WPS使用与Excel兼容的Value2属性
                    return range.GetType().InvokeMember("Value2", 
                        System.Reflection.BindingFlags.GetProperty, null, range, null);
                }
                return null;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"GetCellValue failed: {ex.Message}");
                return null;
            }
        }

        public void SetCellValue(object range, object value)
        {
            try
            {
                if (range != null)
                {
                    // WPS使用与Excel兼容的Value2属性
                    range.GetType().InvokeMember("Value2", 
                        System.Reflection.BindingFlags.SetProperty, null, range, new[] { value });
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SetCellValue failed: {ex.Message}");
                throw;
            }
        }

        public string GetRangeAddress(object range)
        {
            try
            {
                if (range != null)
                {
                    var address = range.GetType().InvokeMember("Address", 
                        System.Reflection.BindingFlags.GetProperty, null, range, null);
                    return address?.ToString() ?? string.Empty;
                }
                return string.Empty;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"GetRangeAddress failed: {ex.Message}");
                return string.Empty;
            }
        }

        public bool HasVbaAccess()
        {
            try
            {
                // 尝试访问WPS的VBE对象
                var vbe = _application?.VBE;
                if (vbe != null)
                {
                    // 尝试访问VBE项目来测试权限
                    var projectCount = vbe.VBProjects.Count;
                    Debug.WriteLine($"WPS VBA access check successful, found {projectCount} projects");
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"WPS VBA access check failed: {ex.Message}");
                return false;
            }
        }

        public void ExecuteVbaCode(string vbaCode)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(vbaCode))
                {
                    throw new ArgumentException("VBA代码不能为空");
                }

                if (!HasVbaAccess())
                {
                    throw new InvalidOperationException("没有VBA访问权限");
                }

                var vbe = _application.VBE;
                var activeProject = vbe.ActiveVBProject;
                
                // 查找或创建临时模块
                dynamic tempModule = null;
                try
                {
                    tempModule = activeProject.VBComponents.Item("TempAIModule");
                }
                catch
                {
                    // WPS使用与Excel兼容的VBComponent类型
                    tempModule = activeProject.VBComponents.Add(1); // vbext_ct_StdModule = 1
                    tempModule.Name = "TempAIModule";
                }

                // 清空模块内容
                var codeModule = tempModule.CodeModule;
                if (codeModule.CountOfLines > 0)
                {
                    codeModule.DeleteLines(1, codeModule.CountOfLines);
                }

                // 添加新的VBA代码
                codeModule.AddFromString(vbaCode);

                // 查找并执行主过程
                string mainProcedure = ExtractMainProcedureName(vbaCode);
                if (!string.IsNullOrEmpty(mainProcedure))
                {
                    _application.Run($"TempAIModule.{mainProcedure}");
                    Debug.WriteLine($"WPS VBA code executed successfully: {mainProcedure}");
                }
                else
                {
                    throw new InvalidOperationException("无法找到可执行的VBA过程");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"WPS ExecuteVbaCode failed: {ex.Message}");
                throw;
            }
        }

        public void ShowMessage(string message, string title = "提示")
        {
            try
            {
                MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"ShowMessage failed: {ex.Message}");
            }
        }

        /// <summary>
        /// 从VBA代码中提取主过程名称
        /// </summary>
        /// <param name="vbaCode">VBA代码</param>
        /// <returns>主过程名称</returns>
        private string ExtractMainProcedureName(string vbaCode)
        {
            try
            {
                var lines = vbaCode.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var line in lines)
                {
                    var trimmedLine = line.Trim();
                    if (trimmedLine.StartsWith("Sub ", StringComparison.OrdinalIgnoreCase) ||
                        trimmedLine.StartsWith("Function ", StringComparison.OrdinalIgnoreCase))
                    {
                        var parts = trimmedLine.Split(' ');
                        if (parts.Length >= 2)
                        {
                            var procedureName = parts[1];
                            var parenIndex = procedureName.IndexOf('(');
                            if (parenIndex > 0)
                            {
                                procedureName = procedureName.Substring(0, parenIndex);
                            }
                            return procedureName;
                        }
                    }
                }
                return null;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"ExtractMainProcedureName failed: {ex.Message}");
                return null;
            }
        }
    }
}