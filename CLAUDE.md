# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project overview

OfficeTap is a C# .NET Framework 4.8 Excel VSTO add-in. It shows a top-docked custom task pane named `工作簿标签` under the Ribbon and renders open Excel workbooks as browser-like tab buttons. Clicking a tab calls `Excel.Workbook.Activate()` to switch workbooks.

The repository tracks `OfficeTap.sln` for Visual Studio convenience and `OfficeTap.csproj` as the actual VSTO project. Visual Studio workspace files under `.vs/`, build outputs under `bin/`/`obj/`, and local signing keys (`*.pfx`) are ignored.

## Development commands

Run these from a Visual Studio Developer PowerShell/Command Prompt so MSBuild can find the VSTO targets and .NET Framework reference assemblies.

```powershell
# Restore is not needed: this project has no NuGet packages.

# Build debug
msbuild .\OfficeTap.csproj /p:Configuration=Debug /p:Platform=AnyCPU

# Build release
msbuild .\OfficeTap.csproj /p:Configuration=Release /p:Platform=AnyCPU

# Clean
msbuild .\OfficeTap.csproj /t:Clean /p:Configuration=Debug /p:Platform=AnyCPU
```

To run/debug the add-in, open `OfficeTap.csproj` in Visual Studio 2019/2022 with the “Office/SharePoint development” workload installed, then press `F5`. Visual Studio launches Excel and loads the add-in. On first open, Visual Studio may generate `OfficeTap_TemporaryKey.pfx` locally; keep it untracked.

There is currently no test project, test framework, lint target, or single-test command configured in this repository.

## Architecture notes

- `ThisAddIn.cs` is the VSTO entry point. `ThisAddIn_Startup` constructs `WorkbookTaskPane`, adds it to `CustomTaskPanes`, docks it at `msoCTPDockPositionTop`, sets a fixed height, makes it visible, and subscribes to Excel workbook events.
- Excel event flow is centralized in `ThisAddIn.cs`: `WorkbookNewSheet`, `WorkbookOpen`, and `WorkbookBeforeClose` refresh the tab list; `WorkbookActivate` updates the highlighted active workbook. `ThisAddIn_Shutdown` unwires these events, removes the task pane, and disposes the user control.
- `WorkbookTaskPane.cs` owns the WinForms UI and workbook switching logic. It creates a horizontal, non-wrapping `FlowLayoutPanel` at runtime, enumerates `_application.Workbooks`, creates one `Button` per workbook, stores workbook references in button `Tag`s, and keeps hover tooltips path-private by showing only workbook names.
- UI updates in `WorkbookTaskPane` guard with `InvokeRequired`, because Excel/VSTO event callbacks may need marshaling back to the WinForms UI thread.
- Active tab styling is recalculated by comparing workbook COM references with `ReferenceEquals`; hover state delegates back to `HighlightActiveWorkbook()` on mouse leave.
- The tab context menu can activate, save, close current, close others, close all, copy full path by explicit action, and set session-local color labels. Unsaved workbooks display a `*` marker in their tab text.
- OfficeTap simulates a single-window workflow by minimizing inactive workbook windows with `Excel.XlWindowState.xlMinimized`. Do not use `Excel.Window.Visible = false`, because hiding windows can hide the task pane host.
- `WorkbookTaskPane.Designer.cs` is designer-generated and currently only initializes the base user control and disposes `components`, `_toolTip`, `_tabContextMenu`, and owned tab fonts. Keep hand-written layout/switching changes in `WorkbookTaskPane.cs` unless intentionally using the WinForms designer.
- `OfficeTap.csproj` is a traditional VSTO project file importing `Microsoft.CSharp.targets` and `$(VSToolsPath)\OfficeTools\Microsoft.VisualStudio.Tools.Office.targets`. It references Excel PIA version 15.0 with embedded interop types and targets .NET Framework 4.8.
