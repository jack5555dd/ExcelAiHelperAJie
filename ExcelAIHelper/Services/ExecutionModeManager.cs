using System;
using System.Configuration;
using System.Diagnostics;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAIHelper.Services
{
    /// <summary>
    /// 执行模式枚举
    /// </summary>
    public enum ExecutionMode
    {
        /// <summary>
        /// VSTO模式 - 使用现有的VSTO操作引擎
        /// </summary>
        VSTO,
        
        /// <summary>
        /// VBA模式 - 使用AI生成VBA代码并动态注入执行
        /// </summary>
        VBA,
        
        /// <summary>
        /// 仅聊天模式 - 只进行对话，不执行任何操作
        /// </summary>
        ChatOnly
    }

    /// <summary>
    /// 执行模式管理器
    /// 负责管理当前的执行模式，检查VBA环境可用性
    /// </summary>
    public static class ExecutionModeManager
    {
        private static ExecutionMode _currentMode = ExecutionMode.VSTO;
        private static bool? _vbaEnabled = null;
        private static bool _isInitialized = false;

        /// <summary>
        /// 当前执行模式
        /// </summary>
        public static ExecutionMode CurrentMode
        {
            get => _currentMode;
            set
            {
                if (_currentMode != value)
                {
                    _currentMode = value;
                    SaveCurrentMode();
                    OnModeChanged?.Invoke(value);
                }
            }
        }

        /// <summary>
        /// VBA功能是否启用
        /// </summary>
        public static bool IsVbaEnabled
        {
            get
            {
                if (_vbaEnabled == null)
                {
                    _vbaEnabled = CheckVbaEnvironment();
                }
                return _vbaEnabled.Value;
            }
        }

        /// <summary>
        /// 是否已初始化
        /// </summary>
        public static bool IsInitialized => _isInitialized;

        /// <summary>
        /// 模式变更事件
        /// </summary>
        public static event Action<ExecutionMode> OnModeChanged;

        /// <summary>
        /// 初始化执行模式管理器
        /// </summary>
        public static void Initialize()
        {
            try
            {
                // 从配置中加载上次保存的模式
                LoadCurrentMode();
                
                // 检查VBA环境
                _vbaEnabled = CheckVbaEnvironment();
                
                // 如果当前模式是VBA但VBA不可用，则切换到VSTO模式
                if (_currentMode == ExecutionMode.VBA && !_vbaEnabled.Value)
                {
                    Debug.WriteLine("[ExecutionModeManager] VBA模式不可用，切换到VSTO模式");
                    _currentMode = ExecutionMode.VSTO;
                    SaveCurrentMode();
                }
                
                Debug.WriteLine($"[ExecutionModeManager] 初始化完成，当前模式: {_currentMode}, VBA可用: {_vbaEnabled}");
                _isInitialized = true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ExecutionModeManager] 初始化失败: {ex.Message}");
                _currentMode = ExecutionMode.VSTO; // 默认使用VSTO模式
                _isInitialized = true; // 即使失败也标记为已初始化，避免重复尝试
            }
        }

        /// <summary>
        /// 检查VBA环境是否可用
        /// </summary>
        /// <returns>VBA环境是否可用</returns>
        public static bool CheckVbaEnvironment()
        {
            try
            {
                Debug.WriteLine("[ExecutionModeManager] 开始检查VBA环境");
                
                // 1. 检查Excel应用程序是否可用
                Excel.Application excelApp = null;
                try
                {
                    if (Globals.ThisAddIn == null)
                    {
                        Debug.WriteLine("[ExecutionModeManager] Globals.ThisAddIn为null");
                        return false;
                    }
                    
                    excelApp = Globals.ThisAddIn.Application;
                    if (excelApp == null)
                    {
                        Debug.WriteLine("[ExecutionModeManager] Excel应用程序不可用");
                        return false;
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[ExecutionModeManager] 获取Excel应用程序失败: {ex.Message}");
                    return false;
                }

                // 2. 检查VBE访问权限
                if (!CheckVbeAccess())
                {
                    Debug.WriteLine("[ExecutionModeManager] VBE访问权限未启用");
                    return false;
                }

                // 3. 检查是否有活动工作簿（可选检查，不强制要求）
                try
                {
                    if (excelApp.Workbooks.Count == 0)
                    {
                        Debug.WriteLine("[ExecutionModeManager] 没有活动工作簿，但VBE访问权限可用");
                        // 不返回false，因为VBE权限本身是可用的
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[ExecutionModeManager] 检查工作簿数量失败: {ex.Message}");
                    // 继续检查VBE对象
                }

                // 4. 尝试访问VBE对象
                try
                {
                    Debug.WriteLine("[ExecutionModeManager] 开始访问VBE对象...");
                    var vbe = excelApp.VBE;
                    if (vbe == null)
                    {
                        Debug.WriteLine("[ExecutionModeManager] VBE对象为null");
                        return false;
                    }
                    
                    Debug.WriteLine("[ExecutionModeManager] VBE对象获取成功");
                    
                    // 尝试获取ActiveVBProject，但不强制要求
                    try
                    {
                        var vbProject = vbe.ActiveVBProject;
                        if (vbProject != null)
                        {
                            Debug.WriteLine($"[ExecutionModeManager] VBA项目访问成功: {vbProject.Name}");
                            
                            // 尝试访问VBComponents来进一步验证权限
                            var components = vbProject.VBComponents;
                            Debug.WriteLine($"[ExecutionModeManager] VBComponents访问成功，组件数量: {components.Count}");
                        }
                        else
                        {
                            Debug.WriteLine("[ExecutionModeManager] ActiveVBProject为null，但VBE对象可用");
                        }
                    }
                    catch (Exception projectEx)
                    {
                        Debug.WriteLine($"[ExecutionModeManager] VBA项目访问失败: {projectEx.Message}");
                        // 即使项目访问失败，VBE对象本身可用也算成功
                    }
                }
                catch (System.UnauthorizedAccessException ex)
                {
                    Debug.WriteLine($"[ExecutionModeManager] VBE访问被拒绝: {ex.Message}");
                    Debug.WriteLine("[ExecutionModeManager] 请确保已启用'信任对VBA项目对象模型的访问'");
                    return false;
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    Debug.WriteLine($"[ExecutionModeManager] COM异常: {ex.Message} (HRESULT: 0x{ex.HResult:X8})");
                    if (ex.HResult == unchecked((int)0x800A03EC)) // VBA项目被保护
                    {
                        Debug.WriteLine("[ExecutionModeManager] VBA项目被密码保护或锁定");
                    }
                    return false;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[ExecutionModeManager] VBE对象访问失败: {ex.Message}");
                    Debug.WriteLine($"[ExecutionModeManager] 异常类型: {ex.GetType().Name}");
                    return false;
                }

                Debug.WriteLine("[ExecutionModeManager] VBA环境检查通过");
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ExecutionModeManager] VBA环境检查异常: {ex.Message}");
                Debug.WriteLine($"[ExecutionModeManager] 异常堆栈: {ex.StackTrace}");
                return false;
            }
        }

        /// <summary>
        /// 检查VBE访问权限
        /// </summary>
        /// <returns>是否有VBE访问权限</returns>
        public static bool CheckVbeAccess()
        {
            try
            {
                // 检查注册表中的AccessVBOM设置
                // HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Excel\Security\AccessVBOM
                // 或 HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Excel\Security\AccessVBOM
                
                string[] officeVersions = { "16.0", "15.0", "14.0", "12.0" };
                
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
                                if (value != null && value.ToString() == "1")
                                {
                                    Debug.WriteLine($"[ExecutionModeManager] 找到VBE访问权限设置 (Office {version})");
                                    return true;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[ExecutionModeManager] 检查注册表失败 (Office {version}): {ex.Message}");
                    }
                }
                
                Debug.WriteLine("[ExecutionModeManager] 未找到VBE访问权限设置");
                return false;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ExecutionModeManager] VBE访问权限检查异常: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 切换执行模式
        /// </summary>
        /// <param name="mode">目标模式</param>
        /// <returns>是否切换成功</returns>
        public static bool SwitchMode(ExecutionMode mode)
        {
            try
            {
                // 如果要切换到VBA模式，先检查VBA环境
                if (mode == ExecutionMode.VBA && !IsVbaEnabled)
                {
                    Debug.WriteLine("[ExecutionModeManager] VBA环境不可用，无法切换到VBA模式");
                    return false;
                }

                CurrentMode = mode;
                Debug.WriteLine($"[ExecutionModeManager] 成功切换到模式: {mode}");
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ExecutionModeManager] 模式切换失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 获取VBE访问设置指南
        /// </summary>
        /// <returns>设置指南文本</returns>
        public static string GetVbeAccessInstructions()
        {
            return @"要启用VBA功能，请按以下步骤设置：

1. 打开Excel，点击 文件 → 选项
2. 在左侧菜单中选择 信任中心
3. 点击 信任中心设置 按钮
4. 在左侧菜单中选择 宏设置
5. 勾选 信任对VBA项目对象模型的访问
6. 点击 确定 保存设置
7. 重启Excel使设置生效

注意：此设置需要管理员权限，如果无法修改请联系系统管理员。";
        }

        /// <summary>
        /// 强制刷新VBA环境状态
        /// </summary>
        public static void RefreshVbaStatus()
        {
            _vbaEnabled = null; // 清除缓存
            _vbaEnabled = CheckVbaEnvironment(); // 重新检查
            Debug.WriteLine($"[ExecutionModeManager] VBA状态已刷新: {_vbaEnabled}");
        }

        /// <summary>
        /// 从配置中加载当前模式
        /// </summary>
        private static void LoadCurrentMode()
        {
            try
            {
                string modeString = Properties.Settings.Default.ExecutionMode;
                if (!string.IsNullOrEmpty(modeString) && Enum.TryParse<ExecutionMode>(modeString, out ExecutionMode mode))
                {
                    _currentMode = mode;
                    Debug.WriteLine($"[ExecutionModeManager] 从配置加载模式: {mode}");
                }
                else
                {
                    _currentMode = ExecutionMode.VSTO; // 默认模式
                    Debug.WriteLine("[ExecutionModeManager] 使用默认模式: VSTO");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ExecutionModeManager] 加载配置失败: {ex.Message}");
                _currentMode = ExecutionMode.VSTO;
            }
        }

        /// <summary>
        /// 保存当前模式到配置
        /// </summary>
        private static void SaveCurrentMode()
        {
            try
            {
                Properties.Settings.Default.ExecutionMode = _currentMode.ToString();
                Properties.Settings.Default.Save();
                Debug.WriteLine($"[ExecutionModeManager] 模式已保存: {_currentMode}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ExecutionModeManager] 保存配置失败: {ex.Message}");
            }
        }
    }
}