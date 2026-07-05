using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficeTap
{
    public partial class WorkbookTaskPane : UserControl
    {
        private readonly Excel.Application _application;
        private readonly FlowLayoutPanel _tabContainer;
        private readonly Dictionary<Button, Excel.Workbook> _buttonToWorkbook = new Dictionary<Button, Excel.Workbook>();
        private readonly ToolTip _toolTip = new ToolTip();
        private readonly ContextMenuStrip _tabContextMenu = new ContextMenuStrip();
        private readonly Dictionary<Excel.Workbook, TabColor> _tabColors = new Dictionary<Excel.Workbook, TabColor>();
        private readonly Font _regularTabFont = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular);
        private readonly Font _activeTabFont = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold);
        private Excel.Workbook _contextMenuWorkbook;
        private Excel.Workbook _lastWindowLayoutWorkbook;
        private int _lastWindowLayoutWorkbookCount = -1;

        private enum TabColor
        {
            Default,
            Red,
            Green,
            Blue
        }

        public WorkbookTaskPane(Excel.Application application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            InitializeComponent();
            InitializeTabContextMenu();

            _tabContainer = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                WrapContents = false,
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(4, 2, 4, 2),
                BackColor = Color.FromArgb(243, 243, 243)
            };

            Controls.Add(_tabContainer);

            RefreshWorkbooks();
        }

        public void RefreshWorkbooks()
        {
            if (InvokeRequired)
            {
                Invoke(new Action(RefreshWorkbooks));
                return;
            }

            _tabContainer.SuspendLayout();
            try
            {
                var oldButtons = _tabContainer.Controls.OfType<Button>().ToList();
                _tabContainer.Controls.Clear();
                foreach (var button in oldButtons)
                {
                    _toolTip.SetToolTip(button, null);
                    button.Dispose();
                }

                _buttonToWorkbook.Clear();

                var books = GetOpenWorkbooks();
                RemoveClosedWorkbookColors(books);
                _lastWindowLayoutWorkbook = null;
                _lastWindowLayoutWorkbookCount = -1;
                foreach (var workbook in books)
                {
                    var button = CreateTabButton(workbook);
                    _tabContainer.Controls.Add(button);
                    _buttonToWorkbook[button] = workbook;
                }

                var active = GetActiveWorkbook();
                HighlightActiveWorkbook(active);
                ApplyWorkbookWindowLayout(active, books.Count);
            }
            finally
            {
                _tabContainer.ResumeLayout(true);
            }
        }

        public void SetActiveWorkbook(Excel.Workbook workbook)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<Excel.Workbook>(SetActiveWorkbook), workbook);
                return;
            }

            HighlightActiveWorkbook(workbook);
            ApplyWorkbookWindowLayout(workbook, GetOpenWorkbooks().Count);
        }

        public void RestoreAllWorkbookWindows()
        {
            _lastWindowLayoutWorkbook = null;
            _lastWindowLayoutWorkbookCount = -1;

            try
            {
                foreach (var workbook in GetOpenWorkbooks())
                {
                    RestoreWorkbookWindows(workbook);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine("OfficeTap window restoration skipped during shutdown: " + ex.Message);
            }
        }

        private static IReadOnlyList<Excel.Workbook> GetOpenWorkbooks(Excel.Workbooks workbooks)
        {
            var list = new List<Excel.Workbook>();
            foreach (Excel.Workbook book in workbooks)
            {
                list.Add(book);
            }
            return list;
        }

        private IReadOnlyList<Excel.Workbook> GetOpenWorkbooks()
        {
            return _application.Workbooks != null
                ? GetOpenWorkbooks(_application.Workbooks)
                : Array.Empty<Excel.Workbook>();
        }

        private Button CreateTabButton(Excel.Workbook workbook)
        {
            var name = GetWorkbookDisplayName(workbook);
            var tooltip = GetWorkbookName(workbook);

            var button = new Button
            {
                Text = name,
                TextAlign = ContentAlignment.MiddleCenter,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Height = Math.Max(_tabContainer.ClientSize.Height - 4, 28),
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(2, 0, 2, 0),
                Padding = new Padding(10, 0, 10, 0),
                Font = _regularTabFont,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(32, 32, 32),
                Cursor = Cursors.Hand,
                ContextMenuStrip = _tabContextMenu,
                Tag = workbook
            };

            button.FlatAppearance.BorderSize = 1;
            button.FlatAppearance.BorderColor = Color.FromArgb(210, 210, 210);
            button.MouseEnter += (s, e) => button.BackColor = Color.FromArgb(230, 242, 255);
            button.MouseLeave += (s, e) => HighlightActiveWorkbook();
            button.Click += (s, e) => ActivateWorkbook(workbook);
            button.KeyDown += TabButton_KeyDown;

            _toolTip.SetToolTip(button, tooltip);

            return button;
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (HandleShortcut(keyData))
            {
                return true;
            }

            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void TabButton_KeyDown(object sender, KeyEventArgs e)
        {
            if (HandleShortcut(e.KeyData))
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        private void InitializeTabContextMenu()
        {
            var activateItem = new ToolStripMenuItem("激活");
            activateItem.Click += (s, e) => ActivateWorkbook(_contextMenuWorkbook);

            var saveItem = new ToolStripMenuItem("保存");
            saveItem.Click += (s, e) => SaveWorkbook(_contextMenuWorkbook);

            var closeCurrentItem = new ToolStripMenuItem("关闭当前");
            closeCurrentItem.Click += (s, e) => CloseWorkbook(_contextMenuWorkbook);

            var closeOthersItem = new ToolStripMenuItem("关闭其他");
            closeOthersItem.Click += (s, e) => CloseOtherWorkbooks(_contextMenuWorkbook);

            var closeAllItem = new ToolStripMenuItem("关闭全部");
            closeAllItem.Click += (s, e) => CloseAllWorkbooks();

            var copyFullPathItem = new ToolStripMenuItem("复制完整路径");
            copyFullPathItem.Click += (s, e) => CopyWorkbookFullPath(_contextMenuWorkbook);

            var colorMenu = new ToolStripMenuItem("颜色标签");
            var defaultColorItem = CreateColorMenuItem("默认", TabColor.Default);
            var redColorItem = CreateColorMenuItem("红色", TabColor.Red);
            var greenColorItem = CreateColorMenuItem("绿色", TabColor.Green);
            var blueColorItem = CreateColorMenuItem("蓝色", TabColor.Blue);
            colorMenu.DropDownItems.AddRange(new ToolStripItem[]
            {
                defaultColorItem,
                redColorItem,
                greenColorItem,
                blueColorItem
            });

            _tabContextMenu.Items.AddRange(new ToolStripItem[]
            {
                activateItem,
                saveItem,
                new ToolStripSeparator(),
                closeCurrentItem,
                closeOthersItem,
                closeAllItem,
                new ToolStripSeparator(),
                copyFullPathItem,
                colorMenu
            });
            _tabContextMenu.Opening += (s, e) =>
            {
                _contextMenuWorkbook = null;

                var button = _tabContextMenu.SourceControl as Button;
                Excel.Workbook workbook;
                if (button != null && _buttonToWorkbook.TryGetValue(button, out workbook))
                {
                    _contextMenuWorkbook = workbook;
                }

                var hasWorkbook = _contextMenuWorkbook != null;
                activateItem.Enabled = hasWorkbook;
                saveItem.Enabled = hasWorkbook;
                closeCurrentItem.Enabled = hasWorkbook;
                closeOthersItem.Enabled = hasWorkbook;
                closeAllItem.Enabled = hasWorkbook;
                copyFullPathItem.Enabled = hasWorkbook;
                colorMenu.Enabled = hasWorkbook;
                defaultColorItem.Checked = hasWorkbook && GetTabColor(_contextMenuWorkbook) == TabColor.Default;
                redColorItem.Checked = hasWorkbook && GetTabColor(_contextMenuWorkbook) == TabColor.Red;
                greenColorItem.Checked = hasWorkbook && GetTabColor(_contextMenuWorkbook) == TabColor.Green;
                blueColorItem.Checked = hasWorkbook && GetTabColor(_contextMenuWorkbook) == TabColor.Blue;
                e.Cancel = !hasWorkbook;
            };
            _tabContextMenu.Closed += (s, e) => _contextMenuWorkbook = null;
        }

        private ToolStripMenuItem CreateColorMenuItem(string text, TabColor color)
        {
            var item = new ToolStripMenuItem(text)
            {
                CheckOnClick = false
            };
            item.Click += (s, e) => SetWorkbookColor(_contextMenuWorkbook, color);
            return item;
        }

        private void ActivateWorkbook(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return;
            }

            try
            {
                RestoreWorkbookWindows(workbook);
                workbook.Activate();
                HighlightActiveWorkbook(workbook);
                ApplyWorkbookWindowLayout(workbook, GetOpenWorkbooks().Count);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "无法切换工作簿：" + ex.Message,
                    "OfficeTap",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private void CloseWorkbook(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return;
            }

            try
            {
                CloseWorkbookWithPrompt(workbook);
                BeginRefreshWorkbooks();
            }
            catch (Exception ex)
            {
                ShowWorkbookWarning("无法关闭工作簿：", ex);
                BeginRefreshWorkbooks();
            }
        }

        private void CloseOtherWorkbooks(Excel.Workbook workbookToKeep)
        {
            if (workbookToKeep == null)
            {
                return;
            }

            foreach (var workbook in GetOpenWorkbooks().Where(w => !ReferenceEquals(w, workbookToKeep)).ToList())
            {
                if (!CloseWorkbookWithPrompt(workbook))
                {
                    break;
                }
            }

            BeginRefreshWorkbooks();
        }

        private void CloseAllWorkbooks()
        {
            foreach (var workbook in GetOpenWorkbooks().ToList())
            {
                if (!CloseWorkbookWithPrompt(workbook))
                {
                    break;
                }
            }

            BeginRefreshWorkbooks();
        }

        private bool CloseWorkbookWithPrompt(Excel.Workbook workbook)
        {
            try
            {
                workbook.Close();
                return !IsWorkbookOpen(workbook);
            }
            catch (Exception ex)
            {
                ShowWorkbookWarning("无法关闭工作簿：", ex);
                return false;
            }
        }

        private bool IsWorkbookOpen(Excel.Workbook workbook)
        {
            return GetOpenWorkbooks().Any(openWorkbook => ReferenceEquals(openWorkbook, workbook));
        }

        private void CopyWorkbookFullPath(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return;
            }

            try
            {
                Clipboard.SetText(workbook.FullName ?? GetWorkbookName(workbook));
            }
            catch (Exception ex)
            {
                ShowWorkbookWarning("无法复制完整路径：", ex);
            }
        }

        private void SaveWorkbook(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return;
            }

            try
            {
                workbook.Save();
                BeginRefreshWorkbooks();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "无法保存工作簿：" + ex.Message,
                    "OfficeTap",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private void SetWorkbookColor(Excel.Workbook workbook, TabColor color)
        {
            if (workbook == null)
            {
                return;
            }

            if (color == TabColor.Default)
            {
                _tabColors.Remove(workbook);
            }
            else
            {
                _tabColors[workbook] = color;
            }

            HighlightActiveWorkbook(GetActiveWorkbook());
        }

        private TabColor GetTabColor(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return TabColor.Default;
            }

            TabColor color;
            return _tabColors.TryGetValue(workbook, out color) ? color : TabColor.Default;
        }

        private void RemoveClosedWorkbookColors(IReadOnlyList<Excel.Workbook> openWorkbooks)
        {
            foreach (var workbook in _tabColors.Keys.ToList())
            {
                if (!openWorkbooks.Any(openWorkbook => ReferenceEquals(openWorkbook, workbook)))
                {
                    _tabColors.Remove(workbook);
                }
            }
        }

        private string GetWorkbookName(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return "未命名";
            }

            try
            {
                return workbook.Name ?? "未命名";
            }
            catch (Exception)
            {
                return "未命名";
            }
        }

        private string GetWorkbookDisplayName(Excel.Workbook workbook)
        {
            var name = GetWorkbookName(workbook);
            return IsWorkbookSaved(workbook) ? name : name + " *";
        }

        private bool IsWorkbookSaved(Excel.Workbook workbook)
        {
            try
            {
                return workbook == null || workbook.Saved;
            }
            catch (Exception)
            {
                return true;
            }
        }

        private bool HandleShortcut(Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.Tab))
            {
                ActivateRelativeWorkbook(1);
                return true;
            }

            if (keyData == (Keys.Control | Keys.Shift | Keys.Tab))
            {
                ActivateRelativeWorkbook(-1);
                return true;
            }

            if ((keyData & Keys.Alt) == Keys.Alt)
            {
                var keyCode = keyData & Keys.KeyCode;
                if (keyCode >= Keys.D1 && keyCode <= Keys.D9)
                {
                    ActivateWorkbookAtIndex((int)keyCode - (int)Keys.D1);
                    return true;
                }
            }

            return false;
        }

        private void ActivateRelativeWorkbook(int offset)
        {
            var workbooks = GetDisplayedWorkbooks();
            if (workbooks.Count == 0)
            {
                return;
            }

            var active = GetActiveWorkbook();
            var activeIndex = workbooks.FindIndex(workbook => ReferenceEquals(workbook, active));
            var startIndex = activeIndex >= 0 ? activeIndex : 0;
            var nextIndex = (startIndex + offset + workbooks.Count) % workbooks.Count;
            ActivateWorkbook(workbooks[nextIndex]);
        }

        private void ActivateWorkbookAtIndex(int index)
        {
            var workbooks = GetDisplayedWorkbooks();
            if (index >= 0 && index < workbooks.Count)
            {
                ActivateWorkbook(workbooks[index]);
            }
        }

        private List<Excel.Workbook> GetDisplayedWorkbooks()
        {
            return _tabContainer.Controls
                .OfType<Button>()
                .Select(button => button.Tag as Excel.Workbook)
                .Where(workbook => workbook != null)
                .ToList();
        }

        private void BeginRefreshWorkbooks()
        {
            if (IsDisposed)
            {
                return;
            }

            if (IsHandleCreated)
            {
                BeginInvoke(new Action(RefreshWorkbooks));
            }
        }

        private void ApplyWorkbookWindowLayout(Excel.Workbook activeWorkbook, int workbookCount)
        {
            if (activeWorkbook == null)
            {
                return;
            }

            if (ReferenceEquals(activeWorkbook, _lastWindowLayoutWorkbook) && workbookCount == _lastWindowLayoutWorkbookCount)
            {
                return;
            }

            try
            {
                foreach (var workbook in GetOpenWorkbooks())
                {
                    if (ReferenceEquals(workbook, activeWorkbook))
                    {
                        RestoreWorkbookWindows(workbook);
                    }
                    else
                    {
                        MinimizeWorkbookWindows(workbook);
                    }
                }

                _lastWindowLayoutWorkbook = activeWorkbook;
                _lastWindowLayoutWorkbookCount = workbookCount;
            }
            catch (Exception)
            {
                RestoreAllWorkbookWindows();
            }
        }

        private void RestoreWorkbookWindows(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return;
            }

            foreach (Excel.Window window in workbook.Windows)
            {
                if (window.WindowState == Excel.XlWindowState.xlMinimized)
                {
                    window.WindowState = Excel.XlWindowState.xlNormal;
                }
            }
        }

        private void MinimizeWorkbookWindows(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return;
            }

            foreach (Excel.Window window in workbook.Windows)
            {
                if (window.WindowState != Excel.XlWindowState.xlMinimized)
                {
                    window.WindowState = Excel.XlWindowState.xlMinimized;
                }
            }
        }

        private void ShowWorkbookWarning(string prefix, Exception ex)
        {
            MessageBox.Show(
                prefix + ex.Message,
                "OfficeTap",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }

        private Excel.Workbook GetActiveWorkbook()
        {
            try
            {
                return _application.ActiveWorkbook;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private void HighlightActiveWorkbook()
        {
            HighlightActiveWorkbook(GetActiveWorkbook());
        }

        private void HighlightActiveWorkbook(Excel.Workbook activeWorkbook)
        {
            foreach (Button button in _tabContainer.Controls.OfType<Button>())
            {
                var workbook = button.Tag as Excel.Workbook;
                bool isActive = workbook != null && activeWorkbook != null && ReferenceEquals(workbook, activeWorkbook);

                button.Text = GetWorkbookDisplayName(workbook);
                button.BackColor = isActive ? Color.FromArgb(0, 120, 215) : GetInactiveTabBackColor(workbook);
                button.ForeColor = isActive ? Color.White : Color.FromArgb(32, 32, 32);
                button.FlatAppearance.BorderColor = isActive ? Color.FromArgb(0, 120, 215) : Color.FromArgb(210, 210, 210);
                button.Font = isActive ? _activeTabFont : _regularTabFont;
            }
        }

        private Color GetInactiveTabBackColor(Excel.Workbook workbook)
        {
            switch (GetTabColor(workbook))
            {
                case TabColor.Red:
                    return Color.FromArgb(255, 235, 235);
                case TabColor.Green:
                    return Color.FromArgb(232, 246, 232);
                case TabColor.Blue:
                    return Color.FromArgb(230, 242, 255);
                default:
                    return Color.White;
            }
        }
    }
}
