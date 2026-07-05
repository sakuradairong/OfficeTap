# OfficeTap

Excel VSTO 外接程序示例：在 Excel 顶部（Ribbon 工具栏下方）显示一个自定义任务窗格，用类似浏览器标签页的按钮列出当前所有打开的工作簿，点击按钮即可切换对应工作簿。

## 功能

- 启动 Excel 后自动显示“工作簿标签”任务窗格。
- 实时监听工作簿的打开、新建、关闭、激活事件，动态刷新标签列表。
- 标签按钮显示工作簿文件名；鼠标悬停显示工作簿名，避免暴露完整路径。
- 当前活动工作簿高亮显示。
- 点击标签按钮调用 `Workbook.Activate()` 切换窗口。

## 项目结构

```
OfficeTap/
├── OfficeTap.csproj          # MSBuild 项目文件（传统 VSTO 格式）
├── ThisAddIn.cs              # 外接程序入口：创建 CustomTaskPane 并注册事件
├── WorkbookTaskPane.cs       # 任务窗格用户控件：标签按钮列表与切换逻辑
├── WorkbookTaskPane.Designer.cs
├── WorkbookTaskPane.resx
├── Properties/
│   └── AssemblyInfo.cs       # 程序集元数据
└── README.md
```

## 环境要求

- Windows
- Visual Studio 2019/2022（需安装“Office/SharePoint 开发”工作负载）
- Microsoft Excel 2010 或更高版本（项目引用的是 Excel 2013/2016/2019/365 的 PIA，版本 15.0）
- .NET Framework 4.8

> 说明：这是传统 .NET Framework 4.8 VSTO 项目，权威构建与运行验证必须在 Windows + Visual Studio Office/SharePoint 开发工作负载 + Excel 环境中完成。Linux 环境只能做静态检查或编辑器导航；即使安装跨平台 .NET SDK / C# LSP，也不能替代 VSTO 构建和 Excel 运行验证。

## 打开与构建

1. 双击 `OfficeTap.csproj` 或在 Visual Studio 中选择“打开项目/解决方案”。
2. 首次打开时，Visual Studio 可能会提示生成临时密钥文件 `OfficeTap_TemporaryKey.pfx`，点击“确定”即可。
3. 按 `F5` 启动调试。Visual Studio 会启动 Excel 并加载外接程序。

## 代码要点

### CustomTaskPane 的创建

```csharp
_taskPaneControl = new WorkbookTaskPane(Application);
_customTaskPane = CustomTaskPanes.Add(_taskPaneControl, "工作簿标签");
_customTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionTop;
_customTaskPane.Height = 42;
_customTaskPane.Visible = true;
```

### 工作簿事件监听

```csharp
Application.NewWorkbook += Application_NewWorkbook;
Application.WorkbookOpen += Application_WorkbookOpen;
Application.WorkbookBeforeClose += Application_WorkbookBeforeClose;
Application.WorkbookActivate += Application_WorkbookActivate;
```

关闭工作簿时，刷新会延迟到当前关闭事件返回之后执行，避免在 Excel 尚未移除工作簿时提前重建标签列表。

## 验证清单

Linux 可执行的检查仅限静态层面，例如确认事件订阅、延迟刷新、路径隐私和字体复用相关代码模式是否存在。完整验证需要在 Windows 上执行：

1. 使用 Visual Studio 构建 `OfficeTap.csproj`，或运行 `msbuild .\OfficeTap.csproj /p:Configuration=Debug /p:Platform=AnyCPU`。
2. 按 `F5` 启动 Excel 并加载外接程序。
3. 新建空白工作簿，确认标签立即出现。
4. 打开已有工作簿，确认标签出现。
5. 关闭工作簿，确认关闭完成后标签消失；取消关闭提示时标签列表仍正确。
6. 在多个工作簿间切换，确认高亮跟随真实活动工作簿。
7. 悬停标签，确认只显示工作簿名，不暴露完整文件路径。
8. 反复切换工作簿，观察 GDI 对象数量不应持续增长。

### 切换工作簿

```csharp
workbook.Activate();
```

## 常见问题

- **看不到任务窗格**：检查 `_customTaskPane.Visible` 是否为 `true`；也可通过“视图”→“任务窗格”手动切换。
- **按钮没反应**：确保目标工作簿窗口未被其他模态对话框阻塞；`workbook.Activate()` 在 Excel 忙时可能抛出异常。
- **多个 Excel 进程**：VSTO 外接程序对每个 Excel 进程单独加载，这是正常现象。

## 扩展建议

- 在标签按钮上添加关闭按钮（×），实现一键关闭工作簿。
- 按工作簿修改状态（是否有未保存更改）显示不同图标或颜色。
- 添加搜索框过滤标签列表。
- 将任务窗格位置改为 `Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionTop` 即可固定在 Ribbon 下方。

## 许可

MIT
