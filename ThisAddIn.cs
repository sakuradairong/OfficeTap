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
            Application.NewWorkbook += Application_NewWorkbook;
            Application.WorkbookNewSheet += Application_WorkbookNewSheet;
            Application.WorkbookOpen += Application_WorkbookOpen;
            Application.WorkbookBeforeClose += Application_WorkbookBeforeClose;
            Application.WorkbookActivate += Application_WorkbookActivate;
        }

        private void UnwireWorkbookEvents()
        {
            Application.NewWorkbook -= Application_NewWorkbook;
            Application.WorkbookNewSheet -= Application_WorkbookNewSheet;
            Application.WorkbookOpen -= Application_WorkbookOpen;
            Application.WorkbookBeforeClose -= Application_WorkbookBeforeClose;
            Application.WorkbookActivate -= Application_WorkbookActivate;
        }

        private void Application_NewWorkbook(Excel.Workbook workbook)
        {
            RefreshTabs();
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
            if (!cancel && _taskPaneControl != null && !_taskPaneControl.IsDisposed && _taskPaneControl.IsHandleCreated)
            {
                _taskPaneControl.BeginInvoke(new System.Action(RefreshTabs));
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

        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }
    }
}
