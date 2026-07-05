# PROJECT KNOWLEDGE BASE

**Generated:** 2026-07-05
**Commit:** 88cf81f
**Branch:** main

## OVERVIEW

OfficeTap is a small Excel VSTO add-in that shows open workbooks as browser-like tabs in a top-docked custom task pane. It is a classic .NET Framework 4.8 / WinForms / Excel interop project, not a cross-platform .NET project.

## STRUCTURE

```
office-tap/
├── OfficeTap.csproj              # Classic MSBuild/VSTO project manifest
├── ThisAddIn.cs                  # Hand-written add-in lifecycle and workbook events
├── ThisAddIn.Designer.cs         # VSTO-generated host/bootstrap plumbing
├── WorkbookTaskPane.cs           # Task pane UI and workbook tab behavior
├── WorkbookTaskPane.Designer.cs  # WinForms designer init/dispose half
├── WorkbookTaskPane.resx         # Designer resource companion
├── Globals.cs                    # VSTO singleton-style globals
├── Properties/AssemblyInfo.cs    # Assembly metadata and trust attributes
└── README.md                     # User-facing setup, behavior, and QA notes
```

## WHERE TO LOOK

| Task | Location | Notes |
|------|----------|-------|
| Add-in startup/shutdown | `ThisAddIn.cs` | Creates/removes `WorkbookTaskPane`, wires/unwires workbook events. |
| VSTO host bootstrap | `ThisAddIn.Designer.cs` | `Initialize`, `FinishInitialization`, `InternalStartup`, `OnShutdown`. Treat as generated plumbing. |
| Workbook tab rendering | `WorkbookTaskPane.cs` | `RefreshWorkbooks`, `CreateTabButton`, `HighlightActiveWorkbook`. |
| Workbook switching | `WorkbookTaskPane.cs` | `ActivateWorkbook` calls `Workbook.Activate()` and shows a warning on failure. |
| UI resource cleanup | `WorkbookTaskPane.Designer.cs` | Disposes `components`, `_toolTip`, and owned tab fonts. |
| Build composition | `OfficeTap.csproj` | Target framework, Office/VSTO references, compile/resource item list. |
| Human setup/QA | `README.md` | Windows/VSTO requirements and manual Excel validation checklist. |

## CODE MAP

LSP/codegraph unavailable in this workspace; centrality is from static source inspection.

| Symbol | Type | Location | Refs | Role |
|--------|------|----------|------|------|
| `ThisAddIn_Startup` | method | `ThisAddIn.cs` | VSTO startup event | Creates the custom task pane and starts event wiring. |
| `WireWorkbookEvents` | method | `ThisAddIn.cs` | startup | Subscribes to `NewWorkbook`, `WorkbookNewSheet`, `WorkbookOpen`, `WorkbookBeforeClose`, `WorkbookActivate`. |
| `Application_WorkbookBeforeClose` | method | `ThisAddIn.cs` | Excel event | Defers `RefreshTabs()` with `BeginInvoke` so Excel can finish closing first. |
| `RefreshTabs` | method | `ThisAddIn.cs` | event handlers | Delegates workbook-list refresh into the task pane. |
| `WorkbookTaskPane` | UserControl | `WorkbookTaskPane.cs` | task pane creation | Owns all tab UI state. |
| `RefreshWorkbooks` | method | `WorkbookTaskPane.cs` | constructor/events | Rebuilds tab buttons from `Application.Workbooks`. |
| `CreateTabButton` | method | `WorkbookTaskPane.cs` | refresh | Creates one button per workbook; tooltip uses workbook name only. |
| `SetActiveWorkbook` | method | `WorkbookTaskPane.cs` | `WorkbookActivate` | Updates highlight for the activated workbook. |
| `Globals` | class | `Globals.cs` | designer bootstrap | One-time VSTO global accessors for add-in and factory. |

## CONVENTIONS

- Keep hand-written behavior in `ThisAddIn.cs` and `WorkbookTaskPane.cs`; designer companions should stay limited to generated initialization and disposal glue.
- If adding Excel workbook lifecycle behavior, wire it in `WireWorkbookEvents()` and unwire it in `UnwireWorkbookEvents()` in the same edit.
- Closing workbooks is timing-sensitive: do not synchronously rebuild tabs inside `WorkbookBeforeClose`; defer UI refresh so Excel can update `Application.Workbooks` first.
- Workbook tab tooltip text intentionally avoids full paths; do not reintroduce `Workbook.FullName` in user-visible hover text without an explicit privacy decision.
- `WorkbookTaskPane` owns `_regularTabFont` and `_activeTabFont`; reuse them instead of allocating `Font` objects during every highlight pass.
- `Properties/AssemblyInfo.cs` is metadata/trust only. It does not define runtime behavior.

## ANTI-PATTERNS (THIS PROJECT)

- Do not assume Linux `dotnet build` validates this project. Classic VSTO requires Windows, Visual Studio Office tooling, Excel, and .NET Framework 4.8.
- Do not treat `.codegraph/`, `.omo/`, `.git/`, `bin/`, or `obj/` as product source.
- Do not create a subdirectory AGENTS.md for `Properties/`; one metadata file is covered by this root guide.
- Do not make broad architecture changes. The project is intentionally flat and small.

## COMMANDS

```bash
# Static inspection only in Linux
rg --files -g '!*bin*' -g '!*obj*'
rg -n "Application\.NewWorkbook|WorkbookBeforeClose|BeginInvoke|ReferenceEquals|SetToolTip|new Font|FullName" *.cs README.md

# Authoritative build on Windows with Visual Studio Office/SharePoint workload
msbuild .\OfficeTap.csproj /p:Configuration=Debug /p:Platform=AnyCPU
```

## NOTES

- There is no `.sln`, CI workflow, automated test project, `packages.config`, `global.json`, or build script in the repo.
- Expected debug flow is Visual Studio `F5`, which launches Excel and loads the add-in.
- README is Chinese and is the primary user-facing project guide.
- Existing commit messages are Chinese plain descriptions, not Conventional Commits.
