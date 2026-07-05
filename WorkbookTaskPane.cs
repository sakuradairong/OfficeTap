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
        private readonly Font _regularTabFont = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular);
        private readonly Font _activeTabFont = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold);

        public WorkbookTaskPane(Excel.Application application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            InitializeComponent();

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
            _tabContainer.Controls.Clear();
            _buttonToWorkbook.Clear();

            var books = GetOpenWorkbooks();
            foreach (var workbook in books)
            {
                var button = CreateTabButton(workbook);
                _tabContainer.Controls.Add(button);
                _buttonToWorkbook[button] = workbook;
            }

            HighlightActiveWorkbook();
            _tabContainer.ResumeLayout(true);
        }

        public void SetActiveWorkbook(Excel.Workbook workbook)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<Excel.Workbook>(SetActiveWorkbook), workbook);
                return;
            }

            HighlightActiveWorkbook(workbook);
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
            var name = workbook.Name ?? "未命名";

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
                Tag = workbook
            };

            button.FlatAppearance.BorderSize = 1;
            button.FlatAppearance.BorderColor = Color.FromArgb(210, 210, 210);
            button.MouseEnter += (s, e) => button.BackColor = Color.FromArgb(230, 242, 255);
            button.MouseLeave += (s, e) => HighlightActiveWorkbook();
            button.Click += (s, e) => ActivateWorkbook(workbook);

            _toolTip.SetToolTip(button, name);

            return button;
        }

        private void ActivateWorkbook(Excel.Workbook workbook)
        {
            try
            {
                workbook.Activate();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"无法切换工作簿：{ex.Message}",
                    "OfficeTap",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private void HighlightActiveWorkbook()
        {
            Excel.Workbook active = null;
            try
            {
                active = _application.ActiveWorkbook;
            }
            catch
            {
                // 可能没有活动工作簿
            }

            HighlightActiveWorkbook(active);
        }

        private void HighlightActiveWorkbook(Excel.Workbook activeWorkbook)
        {
            foreach (Button button in _tabContainer.Controls.OfType<Button>())
            {
                var workbook = button.Tag as Excel.Workbook;
                bool isActive = workbook != null && activeWorkbook != null && ReferenceEquals(workbook, activeWorkbook);

                button.BackColor = isActive ? Color.FromArgb(0, 120, 215) : Color.White;
                button.ForeColor = isActive ? Color.White : Color.FromArgb(32, 32, 32);
                button.FlatAppearance.BorderColor = isActive ? Color.FromArgb(0, 120, 215) : Color.FromArgb(210, 210, 210);
                button.Font = isActive ? _activeTabFont : _regularTabFont;
            }
        }
    }
}
