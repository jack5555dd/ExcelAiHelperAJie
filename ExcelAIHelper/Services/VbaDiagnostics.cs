using System;
using System.Diagnostics;
using System.Text;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAIHelper.Services
{
    /// <summary>
    /// VBA权限诊断工具
    /// </summary>
    public static class VbaDiagnostics
    {
        /// <summary>
        /// 执行完整的VBA环境诊断
        /// </summary>
        /// <returns>诊断报告</returns>
        public static string RunFullDiagnostics()
        {
            var report = new StringBuilder();
            report.AppendLine("=== VBA环境诊断报告 ===");
            report.AppendLine($"诊断时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            report.AppendLine();

            // 1. Excel应用程序检查
            report.AppendLine("1. Excel应用程序检查:");
            var excelApp = Globals.ThisAddIn?.Application;
            if (excelApp != null)
            {
                report.AppendLine($"   ✅ Excel应用程序可用");
                report.AppendLine($"   版本: {excelApp.Version}");
                report.AppendLine($"   构建: {excelApp.Build}");
            }
            else
            {
                report.AppendLine("   ❌ Excel应用程序不可用");
                return report.ToString();
            }

            // 2. 注册表权限检查
            report.AppendLine();
            report.AppendLine("2. 注册表权限检查:");
            CheckRegistrySettings(report);

            // 3. VBE对象访问测试
            report.AppendLine();
            report.AppendLine("3. VBE对象访问测试:");
            TestVbeAccess(report, excelApp);

            // 4. 工作簿状态检查
            report.AppendLine();
            report.AppendLine("4. 工作簿状态检查:");
            CheckWorkbookStatus(report, excelApp);

            // 5. 建议和解决方案
            report.AppendLine();
            report.AppendLine("5. 建议和解决方案:");
            ProvideSolutions(report);

            return report.ToString();
        }

        /// <summary>
        /// 检查注册表设置
        /// </summary>
        private static void CheckRegistrySettings(StringBuilder report)
        {
            string[] officeVersions = { "16.0", "15.0", "14.0", "12.0" };
            bool foundSetting = false;

            foreach (string version in officeVersions)
            {
                string keyPath = $@"Software\Microsoft\Office\{version}\Excel\Security";
                
                try
                {
                    using (RegistryKey key = Registry.CurrentUser.OpenSubKey(keyPath))
                    {
                        if (key != null)
                        {
                            object value = key.GetValue("AccessVBOM");
                            if (value != null)
                            {
                                string status = value.ToString() == "1" ? "✅ 已启用" : "❌ 未启用";
                                report.AppendLine($"   Office {version}: {status} (值: {value})");
                                if (value.ToString() == "1")
                                {
                                    foundSetting = true;
                                }
                            }
                            else
                            {
                                report.AppendLine($"   Office {version}: ⚠️ 未找到AccessVBOM设置");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    report.AppendLine($"   Office {version}: ❌ 读取失败 ({ex.Message})");
                }
            }

            if (!foundSetting)
            {
                report.AppendLine("   ⚠️ 未找到任何启用的VBE访问权限设置");
            }
        }

        /// <summary>
        /// 测试VBE对象访问
        /// </summary>
        private static void TestVbeAccess(StringBuilder report, Excel.Application excelApp)
        {
            try
            {
                // 步骤1: 获取VBE对象
                report.AppendLine("   步骤1: 获取VBE对象...");
                var vbe = excelApp.VBE;
                report.AppendLine("   ✅ VBE对象获取成功");

                // 步骤2: 获取活动VB项目
                report.AppendLine("   步骤2: 获取活动VB项目...");
                var vbProject = vbe.ActiveVBProject;
                if (vbProject != null)
                {
                    report.AppendLine($"   ✅ VB项目获取成功: {vbProject.Name}");
                    
                    // 步骤3: 检查项目保护状态
                    report.AppendLine("   步骤3: 检查项目保护状态...");
                    try
                    {
                        var protection = vbProject.Protection;
                        report.AppendLine($"   ✅ 项目保护状态: {protection}");
                    }
                    catch (Exception ex)
                    {
                        report.AppendLine($"   ⚠️ 无法检查保护状态: {ex.Message}");
                    }

                    // 步骤4: 访问VBComponents
                    report.AppendLine("   步骤4: 访问VBComponents...");
                    var components = vbProject.VBComponents;
                    report.AppendLine($"   ✅ VBComponents访问成功，组件数量: {components.Count}");

                    // 步骤5: 尝试创建临时模块（测试写入权限）
                    report.AppendLine("   步骤5: 测试模块创建权限...");
                    TestModuleCreation(report, vbProject);
                }
                else
                {
                    report.AppendLine("   ❌ ActiveVBProject为null");
                }
            }
            catch (System.UnauthorizedAccessException ex)
            {
                report.AppendLine($"   ❌ 访问被拒绝: {ex.Message}");
                report.AppendLine("   原因: VBE访问权限未正确配置");
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                report.AppendLine($"   ❌ COM异常: {ex.Message}");
                report.AppendLine($"   HRESULT: 0x{ex.HResult:X8}");
                
                switch (ex.HResult)
                {
                    case unchecked((int)0x800A03EC):
                        report.AppendLine("   原因: VBA项目被密码保护或锁定");
                        break;
                    case unchecked((int)0x800A9C68):
                        report.AppendLine("   原因: 无法访问VBA项目，可能需要启用宏");
                        break;
                    default:
                        report.AppendLine("   原因: 未知的COM错误");
                        break;
                }
            }
            catch (Exception ex)
            {
                report.AppendLine($"   ❌ 其他异常: {ex.GetType().Name} - {ex.Message}");
            }
        }

        /// <summary>
        /// 测试模块创建权限
        /// </summary>
        private static void TestModuleCreation(StringBuilder report, dynamic vbProject)
        {
            string testModuleName = $"TestModule_{DateTime.Now:HHmmss}";
            dynamic testModule = null;
            
            try
            {
                // 尝试创建临时模块
                testModule = vbProject.VBComponents.Add(1); // vbext_ct_StdModule = 1
                testModule.Name = testModuleName;
                report.AppendLine($"   ✅ 临时模块创建成功: {testModuleName}");

                // 尝试添加代码
                testModule.CodeModule.AddFromString("' Test code");
                report.AppendLine("   ✅ 代码注入测试成功");

                // 清理临时模块
                vbProject.VBComponents.Remove(testModule);
                report.AppendLine("   ✅ 临时模块清理成功");
            }
            catch (Exception ex)
            {
                report.AppendLine($"   ❌ 模块操作失败: {ex.Message}");
                
                // 尝试清理
                if (testModule != null)
                {
                    try
                    {
                        vbProject.VBComponents.Remove(testModule);
                    }
                    catch
                    {
                        // 忽略清理错误
                    }
                }
            }
        }

        /// <summary>
        /// 检查工作簿状态
        /// </summary>
        private static void CheckWorkbookStatus(StringBuilder report, Excel.Application excelApp)
        {
            try
            {
                var activeWorkbook = excelApp.ActiveWorkbook;
                if (activeWorkbook != null)
                {
                    report.AppendLine($"   ✅ 活动工作簿: {activeWorkbook.Name}");
                    report.AppendLine($"   路径: {activeWorkbook.FullName}");
                    report.AppendLine($"   只读: {activeWorkbook.ReadOnly}");
                    report.AppendLine($"   已保存: {activeWorkbook.Saved}");
                    
                    // 检查宏设置
                    try
                    {
                        var hasVBProject = activeWorkbook.HasVBProject;
                        report.AppendLine($"   包含VBA项目: {hasVBProject}");
                    }
                    catch (Exception ex)
                    {
                        report.AppendLine($"   ⚠️ 无法检查VBA项目状态: {ex.Message}");
                    }
                }
                else
                {
                    report.AppendLine("   ❌ 没有活动工作簿");
                }
            }
            catch (Exception ex)
            {
                report.AppendLine($"   ❌ 工作簿检查失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 提供解决方案建议
        /// </summary>
        private static void ProvideSolutions(StringBuilder report)
        {
            report.AppendLine("   如果VBA功能不可用，请尝试以下解决方案：");
            report.AppendLine();
            report.AppendLine("   方案1: 检查信任中心设置");
            report.AppendLine("   - 文件 → 选项 → 信任中心 → 信任中心设置");
            report.AppendLine("   - 宏设置 → 勾选'信任对VBA项目对象模型的访问'");
            report.AppendLine();
            report.AppendLine("   方案2: 重启Excel");
            report.AppendLine("   - 完全关闭Excel应用程序");
            report.AppendLine("   - 重新打开Excel和工作簿");
            report.AppendLine();
            report.AppendLine("   方案3: 检查工作簿保护");
            report.AppendLine("   - 确保工作簿的VBA项目没有密码保护");
            report.AppendLine("   - 开发工具 → Visual Basic → 工具 → VBAProject属性");
            report.AppendLine();
            report.AppendLine("   方案4: 管理员权限");
            report.AppendLine("   - 以管理员身份运行Excel");
            report.AppendLine("   - 联系系统管理员检查组策略设置");
        }

        /// <summary>
        /// 快速检查VBA是否可用
        /// </summary>
        /// <returns>检查结果</returns>
        public static (bool IsAvailable, string Reason) QuickCheck()
        {
            try
            {
                // 检查ThisAddIn是否可用
                if (Globals.ThisAddIn == null)
                {
                    return (false, "ThisAddIn未初始化");
                }

                var excelApp = Globals.ThisAddIn.Application;
                if (excelApp == null)
                {
                    return (false, "Excel应用程序不可用");
                }

                // 安全地访问VBE
                dynamic vbe = null;
                try
                {
                    vbe = excelApp.VBE;
                    if (vbe == null)
                    {
                        return (false, "无法获取VBE对象");
                    }
                }
                catch (System.UnauthorizedAccessException)
                {
                    return (false, "VBE访问权限未启用");
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    return (false, $"VBE访问失败: 0x{ex.HResult:X8}");
                }

                // 安全地访问VBA项目
                dynamic vbProject = null;
                try
                {
                    vbProject = vbe.ActiveVBProject;
                    if (vbProject == null)
                    {
                        return (false, "无法访问VBA项目");
                    }
                }
                catch (System.UnauthorizedAccessException)
                {
                    return (false, "VBA项目访问权限不足");
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    if (ex.HResult == unchecked((int)0x800A03EC))
                    {
                        return (false, "VBA项目被密码保护");
                    }
                    return (false, $"VBA项目访问失败: 0x{ex.HResult:X8}");
                }

                // 尝试访问组件集合
                try
                {
                    var components = vbProject.VBComponents;
                    return (true, "VBA环境正常");
                }
                catch (System.UnauthorizedAccessException)
                {
                    return (false, "VBComponents访问权限不足");
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    return (false, $"VBComponents访问失败: 0x{ex.HResult:X8}");
                }
            }
            catch (System.UnauthorizedAccessException)
            {
                return (false, "VBE访问权限未启用");
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                if (ex.HResult == unchecked((int)0x800A03EC))
                {
                    return (false, "VBA项目被密码保护");
                }
                return (false, $"COM错误: 0x{ex.HResult:X8}");
            }
            catch (Exception ex)
            {
                return (false, $"未知错误: {ex.Message}");
            }
        }
    }
}