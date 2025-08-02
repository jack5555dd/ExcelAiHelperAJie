# AI-VBA 动态注入可执行改造报告

> 目标：在 **Excel AI助手** 现有 VSTO 架构之上，新增「AI 动态生成 VBA 宏 → 注入并运行」能力，实现更高自由度的 Excel 自动化，同时保证安全性与兼容性。

---

## 1  当前基线

* 项目已采用 **分层架构**（UI → 业务逻辑 → 服务 → 操作引擎）并通过 `OperationDispatcher` 与 `InstructionParser` 串联 AI 指令流。
* 现有引擎直接调用 COM 对象模型完成操作；未涉及 VBA 注入。

---

## 2  引入 VBA 动态注入的必要性

| 场景               | VSTO 调用   | VBA 宏 | 说明           |
| ---------------- | --------- | ----- | ------------ |
| 高级公式、数组函数、迭代运算   | 复杂        | 简单    | VBA 原生语法直接支持 |
| 工作簿-级事件、UserForm | 需 C# 事件桥接 | 直接    | 开发效率更高       |
| 兼容老旧宏文件（\*.xlsm） | 受限        | 100%  | 无需迁移         |

> **注意：WPS Office 2025** 已停止原生 VBA 支持，只有安装 MS Office 的计算机可见开发者环境，因此 WPS 场景需降级为原有 VSTO 路径或提示用户安装 VBA 库([Reddit][1], [WPS][2])。

---

## 3  目标架构增量

```
┌───────── 业务逻辑层 ──────────┐
│  InstructionParser (扩展)    │
│  OperationDispatcher (扩展)  │
└────────────┬───────────────┘
             │
┌───────── 服务层 ─────────────┐
│  PromptBuilder  (新增模板)   │
│  DeepSeekClient (复用)       │
└────────────┬───────────────┘
             │
┌──────── 操作引擎层 ─────────┐
│  ExcelOperationEngine (现有)│
│  **VbaInjectionEngine** ★   │
└─────────────────────────────┘
```

### 3.1  核心新增组件 ― **VbaInjectionEngine**

| 方法                                 | 作用                                                                 |
| ---------------------------------- | ------------------------------------------------------------------ |
| `GeneratePrompt(...)`              | 拼接安全指令，要求 AI 仅返回 JSON `{ "macroName": "...", "vbaCode": "..." }`   |
| `StaticScan(vbaCode)`              | 黑/白名单扫描（禁止 `Shell`, `Kill`, `CreateObject("WScript.Shell")` 等高危调用） |
| `InjectAndRun(macroName, vbaCode)` | 调用 VBIDE 对象模型动态插入并执行                                               |

#### 注入流程

1. **解析**：`InstructionParser` 收到 `type:"vba"` 指令块后转交 `VbaInjectionEngine`。
2. **安全扫描**：高危关键字 → 拒绝或请求用户确认。
3. **注入**：

   ```csharp
   var vbProj = Globals.ThisAddIn.Application.VBE.ActiveVBProject;
   var module = vbProj.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
   module.CodeModule.AddFromString(vbaCode); // 直接插入宏:contentReference[oaicite:2]{index=2}
   Globals.ThisAddIn.Application.Run(macroName);
   vbProj.VBComponents.Remove(module); // 可选：执行完即清理
   ```
4. **回收与日志**：记录宏文本 & 执行结果，供审计与回滚。

---

## 4  安全与权限

| 风险            | 对策                                                          |
| ------------- | ----------------------------------------------------------- |
| **VBE 访问被禁止** | 引导用户在 Excel 选项 → 信任中心 → 勾选「信任对 VBA 项目对象模型的访问」([微软支持][3])    |
| 恶意 VBA        | ① 静态扫描；② 沙箱调用 `Application.Run` 捕获异常；③ 用户二次确认               |
| 自动启用宏警告       | 注入宏仅驻留内存，不写入磁盘，避免永久信任                                       |
| 企业策略锁定        | 提供注册表/组策略文档，对应 `AccessVBOM` 键值（仅管理员执行）([Stack Overflow][4]) |

---

## 5  兼容性矩阵

| 平台                       | 是否支持注入 | 兼容策略                                       |
| ------------------------ | ------ | ------------------------------------------ |
| **Excel 2013–365 (Win)** | ✔      | 正常启用                                       |
| **Excel for Mac**        | ✖      | COM + VBE 受限，保持原逻辑                         |
| **WPS Office ≥ 2023**    | ✖      | 显示提示并回退至 VSTO 指令；或指导下载 MSO VBA 库([WPS][2]) |

---

## 6  Prompt 模板（示例）

```text
系统提示：
你是一名资深 Excel VBA 开发者，只返回 JSON：
{
 "macroName": "MyTask",
 "vbaCode": "'#safe\nSub MyTask()\n  '...可使用 Application, WorksheetFunction 等\nEnd Sub"
}
禁止使用：Shell, Kill, CreateObject, FileSystemObject, WScript, Registry, Environ
```

---

## 7  任务拆解（可直接交给 AI-Dev Agent 执行）

| # | 模块   | 关键文件                            | 任务描述                                                       |
| - | ---- | ------------------------------- | ---------------------------------------------------------- |
| 1 | 项目配置 | `ExcelAIHelper.csproj`          | 引用 `Microsoft.Vbe.Interop`；设置 `Embed Interop Types = True` |
| 2 | 引擎   | `Engines/VbaInjectionEngine.cs` | 按 §3 实现类，含静态扫描+注入+执行                                       |
| 3 | 解析器  | `InstructionParser.cs`          | 支持 `{"type":"vba", ...}` 节点并路由                             |
| 4 | 构建器  | `PromptBuilder.cs`              | 新增 `BuildVbaPrompt()` 方法                                   |
| 5 | UI   | `ChatPaneControl.xaml.cs`       | 执行前显示 VBA 代码预览；提供「运行 / 取消」按钮                               |
| 6 | 日志   | `Logging/ActionLogger.cs`       | 记录 VBA 文本 & 执行耗时                                           |
| 7 | 安全   | `Security/Sandbox.cs`           | 实现黑/白名单扫描 + 关键词配置                                          |
| 8 | 测试   | `Tests/VbaInjectionTests.cs`    | 单元：安全扫描；集成：注入运行 ✔                                          |
| 9 | 文档   | `docs/security.md`              | 描述宏信任设置、风险与回滚指令                                            |

> **优先级**：1-4（核心链路）→ 5-6（体验+审计）→ 7-8-9（安全硬化 & 文档）

---

## 8  回归与验收

1. **功能**：自然语言 → AI → 宏注入 → 运行成功；复杂示例如自动插入透视表。
2. **安全**：注入含 `Shell` 关键字的代码应被阻止并报警。
3. **兼容**：Excel 2019、Excel 365 Win10/11 均通过；WPS 回退提示正常。
4. **性能**：宏注入+执行 < 500 ms；未显著增加内存占用。
5. **回滚**：执行后模块被移除；或用户可点击“撤销宏”。

---

## 9  风险与缓解

| 风险                 | 级别 | 缓解措施                                                                                          |
| ------------------ | -- | --------------------------------------------------------------------------------------------- |
| 用户未开启 `AccessVBOM` | 高  | 启动时检测并引导勾选；提供文档链接                                                                             |
| AI 生成无效 VBA        | 中  | 解析前编译 `Application.VBE.ActiveVBProject.VBComponents("...").Activate` → 捕获 `COMException`，回显错误 |
| 宏警报干扰体验            | 中  | 注入后立即删除模块避免持久化；提示“此宏仅在内存运行”                                                                   |

---

## 10  交付物

* **源代码**：全部 Pull Request（包含 90% 覆盖率测试）。
* **用户指南**：`docs/user_vba_injection.md` – 教程 + 常见问题。
* **安全白皮书**：`docs/security.md` – 威胁模型 & 控制。
* **版本标签**：`v0.9.0-vba-injection`。

---

### ✅ 执行此报告，即可在两周内完成 AI-VBA 功能落地，确保 Excel 原生宏能力与现有 VSTO 智能引擎无缝融合，为后续高级自动化奠定基础。

