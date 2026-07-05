using Microsoft.Office.Tools;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficeTap
{
    public partial class ThisAddIn
    {
        private WorkbookTaskPane _taskPaneControl;
        private CustomTaskPane _customTaskPane;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _taskPaneControl = new WorkbookTaskPane(Application);
            _customTaskPane = CustomTaskPanes.Add(_taskPaneControl, "工作簿标签");
            _customTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionTop;
            _customTaskPane.Height = 42;
            _customTaskPane.Visible = true;

            WireWorkbookEvents();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            UnwireWorkbookEvents();

            if (_customTaskPane != null)
            {
                CustomTaskPanes.Remove(_customTaskPane);
                _customTaskPane = null;
            }

            _taskPaneControl?.Dispose();
            _taskPaneControl = null;
        }

        private void WireWorkbookEvents()
        {
            Application.WorkbookNewSheet += Application_WorkbookNewSheet;
            Application.WorkbookOpen += Application_WorkbookOpen;
            Application.WorkbookBeforeClose += Application_WorkbookBeforeClose;
            Application.WorkbookActivate += Application_WorkbookActivate;
        }

        private void UnwireWorkbookEvents()
        {
            Application.WorkbookNewSheet -= Application_WorkbookNewSheet;
            Application.WorkbookOpen -= Application_WorkbookOpen;
            Application.WorkbookBeforeClose -= Application_WorkbookBeforeClose;
            Application.WorkbookActivate -= Application_WorkbookActivate;
        }

        private void Application_WorkbookNewSheet(Excel.Workbook workbook, object sheet)
        {
            RefreshTabs();
        }

        private void Application_WorkbookOpen(Excel.Workbook workbook)
        {
            RefreshTabs();
        }

        private void Application_WorkbookBeforeClose(Excel.Workbook workbook, ref bool cancel)
        {
            if (!cancel)
            {
                RefreshTabs();
            }
        }

        private void Application_WorkbookActivate(Excel.Workbook workbook)
        {
            _taskPaneControl?.SetActiveWorkbook(workbook);
        }

        private void RefreshTabs()
        {
            _taskPaneControl?.RefreshWorkbooks();
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要使用代码编辑器修改
        /// 此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        #endregion
    }
}
