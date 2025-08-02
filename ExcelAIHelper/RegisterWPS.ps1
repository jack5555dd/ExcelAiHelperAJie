# WPS表格 COM加载项注册脚本
# 此脚本将Excel AI Helper注册为WPS表格的COM加载项

param(
    [string]$DllPath = "",
    [switch]$Unregister = $false
)

# 检查管理员权限
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{
    Write-Host "此脚本需要管理员权限。请以管理员身份运行PowerShell。" -ForegroundColor Red
    exit 1
}

# 如果没有提供DLL路径，尝试自动检测
if ([string]::IsNullOrEmpty($DllPath)) {
    $currentDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    $DllPath = Join-Path $currentDir "bin\Release\ExcelAIHelper.dll"
    
    if (-not (Test-Path $DllPath)) {
        $DllPath = Join-Path $currentDir "bin\Debug\ExcelAIHelper.dll"
    }
    
    if (-not (Test-Path $DllPath)) {
        Write-Host "无法找到ExcelAIHelper.dll文件。请指定正确的路径。" -ForegroundColor Red
        exit 1
    }
}

Write-Host "使用DLL路径: $DllPath" -ForegroundColor Green

# COM组件信息
$ProgId = "ExcelAIHelper.AiRibbon"
$FriendlyName = "Excel AI Helper for WPS"
$Description = "AI-powered Excel assistant compatible with WPS Spreadsheets"

# WPS注册表路径
$WpsRegistryPaths = @(
    "HKCU:\Software\Kingsoft\Office\6.0\Common\AddIns\$ProgId",
    "HKLM:\Software\Kingsoft\Office\6.0\Common\AddIns\$ProgId",
    "HKCU:\Software\Kingsoft\Office\6.0\ET\AddIns\$ProgId"
)

if ($Unregister) {
    Write-Host "正在卸载WPS加载项..." -ForegroundColor Yellow
    
    # 从WPS注册表中移除
    foreach ($path in $WpsRegistryPaths) {
        if (Test-Path $path) {
            Remove-Item -Path $path -Recurse -Force
            Write-Host "已移除: $path" -ForegroundColor Green
        }
    }
    
    # 取消注册COM组件
    try {
        & regasm.exe /unregister "$DllPath" /codebase
        Write-Host "COM组件已取消注册" -ForegroundColor Green
    }
    catch {
        Write-Host "取消注册COM组件时出错: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "WPS加载项卸载完成" -ForegroundColor Green
}
else {
    Write-Host "正在注册WPS加载项..." -ForegroundColor Yellow
    
    # 首先注册COM组件
    try {
        & regasm.exe "$DllPath" /codebase
        Write-Host "COM组件已注册" -ForegroundColor Green
    }
    catch {
        Write-Host "注册COM组件时出错: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
    
    # 在WPS注册表中注册加载项
    foreach ($path in $WpsRegistryPaths) {
        try {
            # 创建注册表项
            if (-not (Test-Path $path)) {
                New-Item -Path $path -Force | Out-Null
            }
            
            # 设置加载项属性
            Set-ItemProperty -Path $path -Name "LoadBehavior" -Value 3 -Type DWord
            Set-ItemProperty -Path $path -Name "FriendlyName" -Value $FriendlyName -Type String
            Set-ItemProperty -Path $path -Name "Description" -Value $Description -Type String
            
            Write-Host "已注册到: $path" -ForegroundColor Green
        }
        catch {
            Write-Host "注册到 $path 时出错: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    
    Write-Host "WPS加载项注册完成" -ForegroundColor Green
    Write-Host "请重启WPS表格以加载插件" -ForegroundColor Cyan
}

# 检查WPS安装
$WpsPaths = @(
    "${env:ProgramFiles}\Kingsoft\WPS Office\office6\et.exe",
    "${env:ProgramFiles(x86)}\Kingsoft\WPS Office\office6\et.exe"
)

$WpsFound = $false
foreach ($wpsPath in $WpsPaths) {
    if (Test-Path $wpsPath) {
        Write-Host "找到WPS表格: $wpsPath" -ForegroundColor Green
        $WpsFound = $true
        break
    }
}

if (-not $WpsFound) {
    Write-Host "警告: 未找到WPS Office安装。请确保已安装WPS Office。" -ForegroundColor Yellow
}

Write-Host "脚本执行完成" -ForegroundColor Green