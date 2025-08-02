using System;
using System.Diagnostics;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;

namespace ExcelAIHelper.Services
{
    /// <summary>
    /// Microsoft Excel应用程序实现
    /// </summary>
    public class MicrosoftExcelApplication : IExcelApplication
    {
        private Excel.Application _application;

        public MicrosoftExcelApplication()
        {
            try
            {
                _application = Globals.ThisAddIn.Application;
                Debug.WriteLine("MicrosoftExcelApplication initialized successfully");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"MicrosoftExcelApplication initialization failed: {ex.Message}");
                throw;
            }
        }

        public object Application => _application;

        public object ActiveWorkbook => _application?.ActiveWorkbook;

        public object ActiveSheet => _application?.ActiveSheet;

        public object Selection => _application?.Selection;

        public object GetRange(string address)
        {
            try
            {
                var activeSheet = _application?.ActiveSheet as Excel.Worksheet;
                return activeSheet?.Range[address];
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
                if (range is Excel.Range excelRange)
                {
                    return excelRange.Value2;
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
                if (range is Excel.Range excelRange)
                {
                    excelRange.Value2 = value;
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
                if (range is Excel.Range excelRange)
                {
                    return excelRange.Address;
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
                var vbe = _application?.VBE;
                if (vbe != null)
                {
                    // 尝试访问VBE项目来测试权限
                    var projectCount = vbe.VBProjects.Count;
                    Debug.WriteLine($"VBA access check successful, found {projectCount} projects");
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"VBA access check failed: {ex.Message}");
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
                VBComponent tempModule = null;
                try
                {
                    tempModule = activeProject.VBComponents.Item("TempAIModule");
                }
                catch
                {
                    tempModule = activeProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
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
                    Debug.WriteLine($"VBA code executed successfully: {mainProcedure}");
                }
                else
                {
                    throw new InvalidOperationException("无法找到可执行的VBA过程");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"ExecuteVbaCode failed: {ex.Message}");
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