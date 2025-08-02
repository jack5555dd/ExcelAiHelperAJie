好的，陈梦洁。结合你**C# + VSTO COM 加载项**已在微软 Excel 跑通的现状，给出一份**确定的“最佳方案”**，目标是以**最小改动复用现有代码**，在 **WPS 表格（ET）** 中稳定加载与运行，并为后续扩展留出空间。

---

## 结论（最佳方案一览）

**用一个“纯 COM 外接程序（非 VSTO 依赖）”作为通用外壳（Shim）**：

* 在该外壳中实现 `IDTExtensibility2` + `IRibbonExtensibility`，用 **晚绑定（dynamic/反射）** 访问宿主 `Application` 对象（从而兼容 Excel 与 WPS ET）。
* 保留你现有业务逻辑为独立类库（Core），外壳只做 Ribbon 回调与宿主适配。
* 安装时**双注册**：按 Office 规范写入 `Microsoft\Office\Excel\Addins\<ProgID>`；**另外**写入 WPS 的 `Kingsoft\Office\ET\AddinsWL`（字符串值名=你的 ProgID），WPS 即可识别并加载。([微软学习][1], [add-in-express.com][2], [CSDN博客][3], [博客园][4], [CSDN文库][5])

> 为什么这是最优？
>
> 1. **改动最小**：无需重写成 WPS JavaScript 加载项；
> 2. **双端通用**：同一 DLL、同一套 RibbonXML 与逻辑，Office 与 WPS 同时兼容；
> 3. **稳定可控**：避开 2024 年起 WPS 对 JS 插件发布/调试策略的变动（oem.ini 方式受限），COM 路线更稳定。([bbs.wps.cn][6])

> 需要注意的现实差异
>
> * **自定义任务窗格（CTP）**：VSTO 专属的 `CustomTaskPaneCollection` 在 WPS 端没有等价托管接口；若你依赖“右侧任务窗格”体验，建议：
>
>   * 短期：用**无模式 WinForms** 窗体替代；
>   * 中期：在需要“内嵌任务窗格”的功能上，引入**WPS JavaScript 加载项**（仅做 UI/任务窗格，通过本地端口/管道调用你的 COM/后端）。官方 JS API 的 `CreateTaskPane` 已长期稳定可用。([qn.cache.wpscdn.cn][7], [qn.cache.wpscdn.cn][8], [qn.cache.wpscdn.cn][9])

---

## 落地实施步骤（按顺序执行）

### 1) 工程分层与外壳（Shim）创建

1. 保留你现有的业务库（**Core**）：AI 调用、指令解析、数据处理等不含宿主耦合代码。
2. 新建 **C# Class Library**（强命名，`ComVisible(true)`），作为 **Addin.Shim**：

   * 实现 `AddInDesignerObjects.IDTExtensibility2`（生命周期）与 `Office.IRibbonExtensibility`（Ribbon UI）。
   * 尽量 **不直接引用** `Microsoft.Office.Interop.Excel` 等具体模型，统一通过 **dynamic/反射** 访问 `Application/Workbook/Range/Selection`。这使得在 WPS ET 中同样可用。([博客园][10], [博客园][10])

> 参考：IDTExtensibility2/IRibbonExtensibility 是 Office/WPS 通用的 COM 接口，适合做“一套代码双宿主”。([微软学习][11], [l2m2.top][12])

### 2) Ribbon UI：使用 RibbonX（返回 XML 字符串）

* 在 `GetCustomUI(string ribbonID)` 返回你的 `customUI` XML（2009/07 命名空间），保持与你现有 VSTO 的按钮结构一致；
* 回调签名与 Office 一致（`void OnAction(Office.IRibbonControl control)` 等）。([博客园][10])

### 3) 宿主适配（关键）

* `OnConnection(object application, ...)` 接到的 `application` 即 Excel 或 WPS ET 的 `Application`。
* 用 `dynamic app = application;` 获取 `app.ActiveWorkbook`、`app.Selection` 等通用成员；避免 PIA 强绑定。
* 可通过进程名或应用名粗粒度识别宿主（Excel/WPS）：例如读取进程文件名包含 `EXCEL.EXE/et.exe`；Linux 资料指明 WPS 表格可执行为 `et`，Windows 也常用 `et.exe`。([博客园][13])

### 4) 平台与构建

* 产出 **x86** 与 **x64** 两个版本（WPS 既有 32 位也有 64 位版本；注册/加载需与进程位数一致）。官方下载页与规格页均覆盖 32/64 位系统。([WPS][14], [WPS Office][15])
* 打开 `Register for COM interop`，设置 `ComVisible` 与 `Guid`。
* Ribbon 的 `Office.dll` 建议 **Embed Interop Types** 或使用较低版本（如 Office 2007 12.0）以提高兼容性。([博客园][10])

### 5) 注册与让 WPS 识别（重点）

1. **COM 注册（二选一）**

   * 开发机临时：

     ```bat
     %WinDir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe YourAddin.dll /codebase
     %WinDir%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe YourAddin.dll /codebase
     ```
   * 安装包阶段用 MSI/安装器执行等效注册。
2. **Office 加载项键**（Excel 侧，常规做法）：

   ```
   HKCU\Software\Microsoft\Office\Excel\Addins\<Your.ProgID>\
     FriendlyName=REG_SZ
     Description=REG_SZ
     LoadBehavior=REG_DWORD 3
   ```

   规范详见 MS 文档与 Add-in Express 指南。([微软学习][1], [add-in-express.com][2])
3. **WPS 识别键**（WPS 表格侧，关键差异）：

   ```
   HKCU\Software\Kingsoft\Office\ET\AddinsWL
     (字符串值) 名称 = <Your.ProgID>，数据留空即可
   ```

   注意：注册表不区分大小写（常见 `Kingsoft`/`KingSoft` 都可），此处使用 **HKCU** 为主。该做法在多篇开发实践中被验证有效（Word 对应 `WPS\AddinsWL`，PPT 对应 `WPP\AddinsWL`；Excel 对应 `ET\AddinsWL`）。([CSDN博客][3], [博客园][4], [CSDN文库][5])

> **为什么不是 oem.ini/jsplugins.xml？**
> 这是 **JS 加载项** 的老部署路径，WPS 在 **2024-06** 起限制个人版用该方式，需改用新版发布流程（`wpsjs publish`）。你的场景是 **COM 插件**，不走这条线。了解此变更仅为避免误用。([bbs.wps.cn][6])

### 6) 打包与安装

* 用 **MSI/Advanced Installer/WiX** 生成两个包（x86/x64）或一个引导程序：

  * 执行 COM 注册（RegAsm 等效动作）；
  * 写入 Office `Addins\<ProgID>` 与 WPS `ET\AddinsWL` 键；
  * 可选：写入卸载脚本/日志路径。
* 若要“所有用户”，Office 侧可写 HKLM；WPS 侧以 HKCU 更稳（不同版本对 HKLM 支持不一致，建议以用户维度部署）。([add-in-express.com][2])

### 7) 测试清单（建议逐项打钩）

* [ ] Office Excel 正常加载（验证 `OnConnection`、Ribbon 按钮事件）。
* [ ] 安装 WPS（匹配位数），写入 `ET\AddinsWL` 后，WPS 表格启动触发 `OnConnection`。
* [ ] 在 WPS/Excel 分别验证：获取 `ActiveWorkbook/ActiveSheet/Selection`、读写单元格、范围遍历。
* [ ] Ribbon 图标（`imageMso` 若不兼容则改自定义图片）；右键菜单/快捷键如有。([CSDN博客][16])
* [ ] 无模式窗体（代替 CTP）在两端均正常，不阻塞编辑。
* [ ] 卸载清理（移除 COM 注册与注册表键）。

---

## 关键代码骨架（可直接移植）

**(1) Add-in 主类（纯 COM 外壳）**

```csharp
using System;
using System.Runtime.InteropServices;
using AddInDesignerObjects; // Extensibility.dll
using Office;               // Office.dll for IRibbonExtensibility

namespace ExcelAIHelper.Addin
{
    [ComVisible(true)]
    [Guid("PUT-YOUR-CLSID-HERE")]
    [ProgId("ExcelAIHelper.Addin")]
    public class Entry : IDTExtensibility2, IRibbonExtensibility
    {
        private dynamic _app; // Excel 或 WPS ET 的 Application；用 dynamic 以适配双端
        private static Core.Engine _engine; // 复用你的业务库

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            _app = application;
            _engine = new Core.Engine();
            // TODO: 初始化日志、宿主识别等
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom) { }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        // 返回 RibbonX
        public string GetCustomUI(string ribbonID) => @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon><tabs>
    <tab id='tabAI' label='AI助手'>
      <group id='grpMain' label='常用'>
        <button id='btnRun' label='运行' onAction='OnRun'/>
      </group>
    </tab>
  </tabs></ribbon>
</customUI>";

        // Ribbon 回调
        public void OnRun(IRibbonControl control)
        {
            // 例：读取当前选区文本并调用你的引擎
            var sel = _app?.Selection;                 // WPS/Excel 通用
            string text = TryGetRangeText(sel);
            var result = _engine.Process(text);        // 你的业务逻辑
            if (!string.IsNullOrEmpty(result))
                ShowModeless(result);                  // 用 WinForms 无模式窗体替代 CTP
        }

        private string TryGetRangeText(dynamic selection)
        {
            try { return selection?.Text?.ToString(); } catch { return ""; }
        }

        private void ShowModeless(string message)
        {
            var form = new System.Windows.Forms.Form { Text = "AI助手", TopMost = false };
            var tb = new System.Windows.Forms.TextBox { Multiline = true, Dock = System.Windows.Forms.DockStyle.Fill, Text = message };
            form.Controls.Add(tb);
            form.Show(); // 无模式
        }
    }
}
```

> 接口与做法参考了社区对 **IDTExtensibility2/IRibbonExtensibility** 的实现经验。([博客园][10], [博客园][10])

**(2) WPS 识别注册（.reg 片段，可随安装器写入）**

```reg
Windows Registry Editor Version 5.00

; Office Excel 常规加载项键
[HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\ExcelAIHelper.Addin]
"FriendlyName"="ExcelAI 助手"
"Description"="AI增强插件（Excel/WPS通用）"
"LoadBehavior"=dword:00000003

; WPS 表格识别键（关键）
[HKEY_CURRENT_USER\Software\Kingsoft\Office\ET\AddinsWL]
"ExcelAIHelper.Addin"=""
```

> Office 侧键位与 `LoadBehavior` 规范来自微软文档；WPS 侧 `ET\AddinsWL` 为识别列表（字符串名=ProgID）。([微软学习][1], [add-in-express.com][2], [CSDN博客][3], [博客园][4])

---

## 兼容与限制清单（提前规避）

1. **任务窗格（CTP）**：VSTO 的 CTP 不要指望在 WPS 端可用；用 **WinForms 无模式窗体** 或另起一套 **WPS JS 加载项** 的 `CreateTaskPane` 来实现右侧停靠页面。([qn.cache.wpscdn.cn][7], [qn.cache.wpscdn.cn][8])
2. **idMso 图标差异**：部分 `imageMso` 名称在 WPS 不匹配时，使用自定义位图更稳。([CSDN博客][16])
3. **对象模型细微差异**：极少数属性/事件在 WPS 不存在或行为不同。通过 **dynamic + Try/Catch** 兜底，避免强类型失败；对关键路径做“特性探测”。（社区长期总结表明两端大体一致，但需守住边界）([CSDN文库][17], [CSDN博客][18])
4. **发布与调试变更（JS 线）**：若将来叠加 JS 加载项，用 `wpsjs publish` 正式发布，别再依赖 `oem.ini`。([bbs.wps.cn][6])

---

## 里程碑与工期（保守估算）

* **D1-D2**：抽离 Core、创建 Addin.Shim、实现 `IDTExtensibility2`/`IRibbonExtensibility`。
* **D3**：Ribbon 回调映射 + 宿主适配（dynamic）+ 基础功能冒烟。
* **D4**：安装器脚本（RegAsm、Office 键、WPS `ET\AddinsWL`）、x86/x64 双构建。
* **D5**：测试矩阵（Excel 2016/2019/365、WPS 近期版本），修兼容性。
  若叠加“JS 任务窗格 UI 封装（可选）”，+2\~3 天：`ribbon.xml + main.js + CreateTaskPane`，本地 HTTP/命名管道与 COM 交互。([qn.cache.wpscdn.cn][7], [qn.cache.wpscdn.cn][9])

---

## 参考与依据（关键点对应文档/实战）

* **COM 外接程序通用接口**：`IDTExtensibility2` / `IRibbonExtensibility` 的实现与示例要点。([博客园][10], [博客园][10])
* **Office 侧注册规范**（Addins 键、LoadBehavior）：微软/第三方部署文档。([微软学习][1], [add-in-express.com][2])
* **WPS 侧识别键**：`HKCU\Software\Kingsoft\Office\ET\AddinsWL`（同理 WPS/WPP），社区多例实证。([CSDN博客][3], [博客园][4], [CSDN文库][5])
* **WPS JS 加载项任务窗格**（若需 UI 内嵌）：`CreateTaskPane` / 开发说明。([qn.cache.wpscdn.cn][7], [qn.cache.wpscdn.cn][9])
* **WPS 对 JS 加载项发布/调试策略更新（2024-06）**：避免走被限制的 oem.ini 旧路。([bbs.wps.cn][6])
* **宿主识别辅助**：WPS 表格可执行名 `et` 的参考。([博客园][13])

---

## 你现在可以直接做的三步

1. 在现有解决方案中**新建 Addin.Shim**，复制上面的 `Entry` 骨架，改 ProgID/GUID。
2. 生成 x86/x64，两条 `RegAsm` 注册后，在注册表写入：

   * `HKCU\Software\Microsoft\Office\Excel\Addins\<你的ProgID>`（`LoadBehavior=3`）；
   * `HKCU\Software\Kingsoft\Office\ET\AddinsWL` 下新建字符串值（名=你的 ProgID）。([微软学习][1], [add-in-express.com][2], [CSDN博客][3])
3. 启动 WPS 表格验证 `OnConnection` 与 Ribbon 按钮是否触发；若需右侧窗格体验，先用 WinForms 无模式窗体替代，后续再按需接入 JS 任务窗格。

---

如果你把**当前 ProgID、现有 Ribbon 按钮清单**发我，我可以给你：

* 一份**可直接双击导入的 .reg**（Office+WPS）和
* **Advanced Installer/WiX** 的安装项配置片段（含 x86/x64 分发逻辑），
  再附上**最小型宿主适配层**（dynamic + 兼容探测）代码，帮助你**一天内打通 WPS**。需要我直接把上述骨架打包成 VS 工程样板一起给你吗？

[1]: https://learn.microsoft.com/en-us/visualstudio/vsto/registry-entries-for-vsto-add-ins?view=vs-2022&utm_source=chatgpt.com "Registry entries for VSTO Add-ins - Visual Studio (Windows)"
[2]: https://www.add-in-express.com/docs/net-deploying-addins.php?utm_source=chatgpt.com "How to deploy and register Office addin: Outlook, Excel, Word"
[3]: https://blog.csdn.net/weixin_41116263/article/details/147572215?utm_source=chatgpt.com "C# 开发Office 和WPS COM 加载项转载"
[4]: https://www.cnblogs.com/SuperRight/p/18414061?utm_source=chatgpt.com "C#开发word(wps)插件（com加载项）作为一个web前端开发"
[5]: https://wenku.csdn.net/answer/30c0zayae4?utm_source=chatgpt.com "HKEY_CURRENT_USER\SOFTWARE\KingSoft\Office\ET ..."
[6]: https://bbs.wps.cn/topic/36774 "近期wps加载项不能调试和加载的问题说明"
[7]: https://qn.cache.wpscdn.cn/encs/doc/office_v11/topics/WPS%20%E5%8A%A0%E8%BD%BD%E9%A1%B9%E5%BC%80%E5%8F%91/%E5%8A%A0%E8%BD%BD%E9%A1%B9%20API%20%E5%8F%82%E8%80%83/%E4%BB%BB%E5%8A%A1%E7%AA%97%E6%A0%BC/%E4%BB%BB%E5%8A%A1%E7%AA%97%E6%A0%BC%E6%A6%82%E8%BF%B0.html?utm_source=chatgpt.com "任务窗格概述- WPS 加载项开发"
[8]: https://qn.cache.wpscdn.cn/encs/doc/office_v13/index.htm?page=WPS+%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3%2F%E5%8A%A0%E8%BD%BD%E9%A1%B9+API+%E5%8F%82%E8%80%83%2F%E4%BB%BB%E5%8A%A1%E7%AA%97%E6%A0%BC%2FTaskPane+%E5%AF%B9%E8%B1%A1.htm&utm_source=chatgpt.com "TaskPane 对象与office传统的taskpanes类似，不同之处理"
[9]: https://qn.cache.wpscdn.cn/encs/doc/office_v5/topics/WPS%20%E5%8A%A0%E8%BD%BD%E9%A1%B9%E5%BC%80%E5%8F%91/WPS%20%E5%8A%A0%E8%BD%BD%E9%A1%B9%E5%BC%80%E5%8F%91%E8%AF%B4%E6%98%8E.html?utm_source=chatgpt.com "WPS 加载项开发说明"
[10]: https://www.cnblogs.com/BluePointLilac/p/18802868?utm_source=chatgpt.com "C# 开发Office 和WPS COM 加载项- 蓝点lilac"
[11]: https://learn.microsoft.com/zh-cn/office/vba/outlook/how-to/office-fluent-ui-extensibility/implementing-the-iribbonextensibility-interface?utm_source=chatgpt.com "实现IRibbonExtensibility 接口"
[12]: https://l2m2.top/2021/10/22/2021-10-22-build-cplusplus-addin-for-ppt-1/?utm_source=chatgpt.com "C++编写PowerPoint插件（一）：第一个插件程序 - L2M2"
[13]: https://www.cnblogs.com/qingzhen/p/18223664?utm_source=chatgpt.com "office可执行文件位置- 麒麟正青春"
[14]: https://zh-hant.wps.com/office/windows/?utm_source=chatgpt.com "[官方] Windows 版WPS Office：下載免費的全功能Office 套件"
[15]: https://platform.wps.cn/?utm_source=chatgpt.com "WPS Office-支持多人在线编辑多种文档格式_WPS官方网站"
[16]: https://blog.csdn.net/weixin_45055317/article/details/133832128?utm_source=chatgpt.com "兼容Office和WPS中Word图标库原创"
[17]: https://wenku.csdn.net/answer/4ff07iknsy?utm_source=chatgpt.com "wps vba 和excel vba差异"
[18]: https://blog.csdn.net/HackMasterX/article/details/133425305?utm_source=chatgpt.com "WPS和Office：编程中的区别原创"
