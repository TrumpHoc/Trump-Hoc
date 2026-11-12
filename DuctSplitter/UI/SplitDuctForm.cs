using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using TD.DuctSplitter.Core;

namespace TD.DuctSplitter.UI
{
    /// <summary>Hộp thoại nhập thông số chia đốt ống gió.</summary>
    public class SplitDuctForm : Form
    {
        private readonly SplitSettings _s;

        private readonly NumericUpDown _rect = Num(50, 20000);
        private readonly NumericUpDown _round = Num(50, 20000);
        private readonly NumericUpDown _oval = Num(50, 20000);
        private readonly NumericUpDown _tail = Num(0, 5000);
        private readonly ComboBox _start = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 260 };
        private readonly RadioButton _scopeSel = new RadioButton { Text = "Ống gió đang chọn (chưa chọn thì pick)", AutoSize = true };
        private readonly RadioButton _scopeView = new RadioButton { Text = "Tất cả ống gió trong view hiện hành", AutoSize = true };
        private readonly CheckBox _dim = new CheckBox { Text = "Tự động dim các Union trong view hiện hành", AutoSize = true };
        private readonly CheckBox _dimEnds = new CheckBox { Text = "Dim cả 2 đầu tuyến ống", AutoSize = true };
        private readonly NumericUpDown _dimOffset = Num(0, 100);
        private readonly ComboBox _dimType = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 260 };

        public SplitDuctForm(SplitSettings settings, IList<string> dimTypeNames, bool viewAllowsDims, string viewNote)
        {
            _s = settings;
            Text = "Chia đốt ống gió";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Font = new Font("Segoe UI", 9f);
            Padding = new Padding(10);

            var root = new TableLayoutPanel { ColumnCount = 1, AutoSize = true, Dock = DockStyle.Fill };
            Controls.Add(root);

            // --- Chiều dài đốt
            var gLen = Group("Chiều dài đốt ống (mm) – đo theo tim ống");
            var tLen = Grid();
            AddRow(tLen, "Ống chữ nhật:", _rect);
            AddRow(tLen, "Ống tròn:", _round);
            AddRow(tLen, "Ống oval:", _oval);
            AddRow(tLen, "Đốt cuối tối thiểu:", _tail);
            AddRow(tLen, "Bắt đầu đo từ:", _start);
            gLen.Controls.Add(tLen);
            root.Controls.Add(gLen);

            _start.Items.AddRange(new object[]
            {
                "Tự động (đầu đang nối fitting/thiết bị)",
                "Điểm đầu (Start) của ống",
                "Điểm cuối (End) của ống"
            });

            // --- Phạm vi
            var gScope = Group("Phạm vi");
            var fScope = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, Dock = DockStyle.Fill };
            fScope.Controls.Add(_scopeSel);
            fScope.Controls.Add(_scopeView);
            gScope.Controls.Add(fScope);
            root.Controls.Add(gScope);

            // --- Dim
            var gDim = Group("Dim Union");
            var tDim = Grid();
            tDim.Controls.Add(_dim, 0, tDim.RowCount); tDim.SetColumnSpan(_dim, 2); tDim.RowCount++;
            tDim.Controls.Add(_dimEnds, 0, tDim.RowCount); tDim.SetColumnSpan(_dimEnds, 2); tDim.RowCount++;
            AddRow(tDim, "Khoảng cách dim (mm giấy):", _dimOffset);
            AddRow(tDim, "Kiểu dim:", _dimType);
            if (!viewAllowsDims)
            {
                var note = new Label { Text = "⚠ " + viewNote, AutoSize = true, ForeColor = Color.DarkRed, MaximumSize = new Size(420, 0) };
                tDim.Controls.Add(note, 0, tDim.RowCount); tDim.SetColumnSpan(note, 2); tDim.RowCount++;
            }
            gDim.Controls.Add(tDim);
            root.Controls.Add(gDim);

            _dimType.Items.Add("<Mặc định của dự án>");
            foreach (string n in dimTypeNames) _dimType.Items.Add(n);

            // --- Nút
            var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Dock = DockStyle.Fill };
            var ok = new Button { Text = "Chạy", DialogResult = DialogResult.OK, Width = 90 };
            var cancel = new Button { Text = "Huỷ", DialogResult = DialogResult.Cancel, Width = 90 };
            buttons.Controls.Add(cancel);
            buttons.Controls.Add(ok);
            root.Controls.Add(buttons);
            AcceptButton = ok;
            CancelButton = cancel;

            // --- Giá trị ban đầu
            _rect.Value = Clamp(_rect, settings.RectLengthMm);
            _round.Value = Clamp(_round, settings.RoundLengthMm);
            _oval.Value = Clamp(_oval, settings.OvalLengthMm);
            _tail.Value = Clamp(_tail, settings.MinTailMm);
            _start.SelectedIndex = (int)settings.StartSide;
            _scopeSel.Checked = settings.Scope == ScopeMode.Selection;
            _scopeView.Checked = settings.Scope == ScopeMode.ActiveView;
            _dim.Checked = settings.CreateDimensions && viewAllowsDims;
            _dim.Enabled = viewAllowsDims;
            _dimEnds.Checked = settings.DimIncludeEnds;
            _dimOffset.Value = Clamp(_dimOffset, settings.DimOffsetMm);
            int idx = _dimType.Items.IndexOf(settings.DimensionTypeName ?? "");
            _dimType.SelectedIndex = idx > 0 ? idx : 0;

            _dim.CheckedChanged += (o, e) => UpdateDimEnabled();
            UpdateDimEnabled();

            ok.Click += (o, e) => Apply(viewAllowsDims);
        }

        private void UpdateDimEnabled()
        {
            bool on = _dim.Checked && _dim.Enabled;
            _dimEnds.Enabled = _dimOffset.Enabled = _dimType.Enabled = on;
        }

        private void Apply(bool viewAllowsDims)
        {
            _s.RectLengthMm = (double)_rect.Value;
            _s.RoundLengthMm = (double)_round.Value;
            _s.OvalLengthMm = (double)_oval.Value;
            _s.MinTailMm = (double)_tail.Value;
            _s.StartSide = (StartSide)_start.SelectedIndex;
            _s.Scope = _scopeView.Checked ? ScopeMode.ActiveView : ScopeMode.Selection;
            if (viewAllowsDims) _s.CreateDimensions = _dim.Checked; // 3D view: giữ nguyên lựa chọn cũ cho lần sau
            _s.DimIncludeEnds = _dimEnds.Checked;
            _s.DimOffsetMm = (double)_dimOffset.Value;
            _s.DimensionTypeName = _dimType.SelectedIndex > 0 ? _dimType.SelectedItem.ToString() : "";
        }

        /// <summary>true nếu người dùng muốn tạo dim ở lần chạy này.</summary>
        public bool DimThisRun => _dim.Enabled && _dim.Checked;

        // ---------- helpers
        private static NumericUpDown Num(int min, int max) =>
            new NumericUpDown { Minimum = min, Maximum = max, DecimalPlaces = 0, Increment = 50, Width = 100, TextAlign = HorizontalAlignment.Right };

        private static decimal Clamp(NumericUpDown n, double v) =>
            Math.Max(n.Minimum, Math.Min(n.Maximum, (decimal)v));

        private static GroupBox Group(string text) =>
            new GroupBox { Text = text, AutoSize = true, Dock = DockStyle.Fill, Padding = new Padding(8) };

        private static TableLayoutPanel Grid()
        {
            var t = new TableLayoutPanel { ColumnCount = 2, AutoSize = true, Dock = DockStyle.Fill };
            t.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            return t;
        }

        private static void AddRow(TableLayoutPanel t, string label, Control c)
        {
            int r = t.RowCount;
            t.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(3, 7, 3, 3) }, 0, r);
            t.Controls.Add(c, 1, r);
            t.RowCount = r + 1;
        }
    }
}
