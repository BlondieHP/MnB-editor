using MnBEditor.Core;
using MnBEditor.Core.Models;

namespace MnBEditor.WinForms.Editors;

/// <summary>
/// Faction editor — mimics original frmFactions layout.
/// </summary>
public class FactionListEditor : UserControl
{
    private readonly DataRepository _repo;
    private DataGridView _grid = null!;
    private TextBox _txtStrId = null!, _txtName = null!, _txtFlags = null!, _txtColor = null!;
    private Label _lblStatus = null!;
    private int _currentIndex = -1;

    public FactionListEditor(DataRepository repo)
    {
        _repo = repo; Dock = DockStyle.Fill; InitializeUI(); RefreshList();
    }

    private void InitializeUI()
    {
        _grid = new DataGridView
        {
            Dock = DockStyle.Left, Width = 380,
            AllowUserToAddRows = false, AllowUserToDeleteRows = false,
            ReadOnly = true, SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            MultiSelect = false, RowHeadersVisible = false
        };
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "ID", Width = 36 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "StrID", Width = 160 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Name", Width = 150 });
        _grid.SelectionChanged += OnSel;

        var props = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(8) };
        var layout = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 2 };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        _txtStrId = AddF(layout, "StrID:"); _txtName = AddF(layout, "Name:");
        _txtFlags = AddF(layout, "Flags:"); _txtColor = AddF(layout, "Color:");

        var btn = new Button { Text = "&Apply", Width = 100 }; btn.Click += OnApply; layout.Controls.Add(btn);
        _lblStatus = new Label { Text = "", AutoSize = true }; layout.Controls.Add(_lblStatus);
        props.Controls.Add(layout);

        var split = new SplitContainer { Dock = DockStyle.Fill, SplitterDistance = 380 };
        split.Panel1.Controls.Add(_grid); split.Panel2.Controls.Add(props); Controls.Add(split);
    }

    private static TextBox AddF(TableLayoutPanel p, string l) { var tb = new TextBox { Dock = DockStyle.Fill }; p.Controls.Add(new Label { Text = l, TextAlign = ContentAlignment.MiddleRight, Dock = DockStyle.Fill }); p.Controls.Add(tb); return tb; }

    private void RefreshList() { _grid.Rows.Clear(); foreach (var f in _repo.Factions) _grid.Rows.Add(f.Id, f.StrId, f.Name); }

    private void OnSel(object? s, EventArgs e)
    {
        if (_grid.SelectedRows.Count == 0) return; _currentIndex = _grid.SelectedRows[0].Index;
        if (_currentIndex >= _repo.Factions.Count) return;
        var f = _repo.Factions[_currentIndex];
        _txtStrId.Text = f.StrId; _txtName.Text = f.Name; _txtFlags.Text = f.Flags.ToString();
        _txtColor.Text = f.Color; _lblStatus.Text = "";
    }

    private void OnApply(object? s, EventArgs e)
    {
        if (_currentIndex < 0) return;
        try
        {
            var f = _repo.Factions[_currentIndex];
            f.StrId = _txtStrId.Text; f.Name = _txtName.Text;
            f.Flags = long.TryParse(_txtFlags.Text, out var fl) ? fl : 0;
            f.Color = _txtColor.Text; f.IsEdited = true;
            _grid.Rows[_currentIndex].Cells[1].Value = f.StrId;
            _grid.Rows[_currentIndex].Cells[2].Value = f.Name;
            _lblStatus.Text = "Applied";
            _repo.NotifyFactionChanged(_currentIndex);
        }
        catch (Exception ex) { _lblStatus.Text = $"Error: {ex.Message}"; }
    }
}
