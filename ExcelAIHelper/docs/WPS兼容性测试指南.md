# WPS兼容性测试指南

## 概述
本文档描述如何测试Excel AI Helper在WPS表格中的兼容性。

## 前置条件

### 1. 安装要求
- Windows 10/11
- WPS Office (推荐最新版本)
- .NET Framework 4.7.2 或更高版本
- Visual Studio 2019/2022 (用于开发和调试)

### 2. WPS Office 安装验证
确保WPS Office已正确安装：
```
检查路径：
- C:\Program Files\Kingsoft\WPS Office\office6\et.exe
- C:\Program Files (x86)\Kingsoft\WPS Office\office6\et.exe
```

## 部署步骤

### 1. 编译项目
```bash
# 在Visual Studio中编译项目
# 确保选择Release配置
```

### 2. 注册COM组件到WPS
使用PowerShell脚本（以管理员身份运行）：
```powershell
# 注册到WPS
.\RegisterWPS.ps1

# 如果需要卸载
.\RegisterWPS.ps1 -Unregister
```

或者手动使用注册表文件：
```bash
# 双击运行RegisterWPS.reg文件
# 注意：需要根据实际DLL路径修改注册表文件
```

### 3. 验证注册
检查注册表项是否创建：
```
HKEY_CURRENT_USER\Software\Kingsoft\Office\6.0\Common\AddIns\ExcelAIHelper.AiRibbon
HKEY_CURRENT_USER\Software\Kingsoft\Office\6.0\ET\AddIns\ExcelAIHelper.AiRibbon
```

## 功能测试

### 1. 基本加载测试
1. 启动WPS表格
2. 检查是否出现"AI助手"选项卡
3. 验证所有按钮是否正常显示

### 2. 应用程序检测测试
1. 打开AI聊天面板
2. 查看系统消息，确认显示"当前运行在 WPS表格"
3. 验证兼容模式提示信息

### 3. 基本操作测试
测试以下功能：
- [ ] 聊天面板打开/关闭
- [ ] 设置单元格值
- [ ] 应用简单公式
- [ ] 格式化操作
- [ ] 聚光灯功能
- [ ] 快速录入功能

### 4. AI功能测试
1. 配置DeepSeek API密钥
2. 测试自然语言指令：
   - "在A1输入100"
   - "给选中区域设置红色背景"
   - "计算A1到A10的总和"

### 5. VBA兼容性测试
1. 检查VBA访问权限
2. 测试简单VBA代码执行
3. 验证AI生成VBA代码功能

## 已知限制

### 1. WPS特定限制
- 某些高级Excel功能可能不完全兼容
- VBA对象模型可能存在细微差异
- 部分格式化选项可能需要特殊处理

### 2. 功能差异
- 图标显示可能与Excel版本略有不同
- 某些快捷键可能不同
- 错误消息格式可能有差异

## 故障排除

### 1. 加载项未显示
- 检查WPS版本兼容性
- 验证注册表项是否正确
- 检查DLL文件路径
- 确认.NET Framework版本

### 2. 功能异常
- 查看Windows事件日志
- 检查Debug输出
- 验证API配置
- 重新注册COM组件

### 3. 性能问题
- 检查WPS版本
- 验证系统资源
- 调整超时设置

## 调试技巧

### 1. 启用调试输出
```csharp
System.Diagnostics.Debug.WriteLine("Debug message");
```

### 2. 使用Visual Studio调试器
1. 附加到WPS进程 (et.exe)
2. 设置断点
3. 逐步调试

### 3. 日志记录
检查以下位置的日志：
- Windows事件查看器
- Visual Studio输出窗口
- 自定义日志文件

## 测试清单

### 基本功能
- [ ] WPS表格启动时加载项正常加载
- [ ] 菜单栏显示"AI助手"选项卡
- [ ] 所有按钮图标正常显示
- [ ] 聊天面板可以正常打开/关闭

### 兼容性检测
- [ ] 应用程序类型检测正确
- [ ] 显示WPS兼容模式提示
- [ ] 基本Excel操作正常工作

### AI功能
- [ ] API连接正常
- [ ] 自然语言指令解析正确
- [ ] 操作执行成功
- [ ] 错误处理正常

### 高级功能
- [ ] VBA权限检查正常
- [ ] VBA代码执行成功
- [ ] 聚光灯功能正常
- [ ] 快速录入功能正常

## 版本兼容性

### 支持的WPS版本
- WPS Office 2019
- WPS Office 2021
- WPS Office 365

### 测试环境
- Windows 10 (1909+)
- Windows 11
- .NET Framework 4.7.2+

## 反馈和报告

如果发现问题，请提供以下信息：
1. WPS Office版本号
2. Windows版本
3. 错误描述和重现步骤
4. 相关日志和截图
5. 系统配置信息

## 更新日志

### v1.0 (2025-08-02)
- 初始WPS兼容性支持
- 应用程序检测机制
- 基本功能适配
- 注册脚本和文档