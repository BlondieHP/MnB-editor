using MnBEditor.Core;
using MnBEditor.Core.Models;

namespace MnBEditor.WinForms.Editors;

/// <summary>
/// Troop editor — mimics the original frmTroops layout.
/// </summary>
public class TroopListEditor : UserControl
{
    private readonly DataRepository _repo;
    private DataGridView _grid = null!;
    private Panel _props = null!;
    private TextBox _txtStrId = null!, _txtName = null!, _txtPlural = null!;
    private TextBox _txtLevel = null!, _txtStr = null!, _txtAgi = null!, _txtInt = null!, _txtCha = null!;
    private TextBox _txt1h = null!, _txt2h = null!, _txtPole = null!, _txtArch = null!, _txtXbow = null!, _txtThr = null!, _txtFire = null!;
    private ComboBox _cmbFaction = null!;
    private Label _lblStatus = null!;
    private int _currentIndex = -1;

    public TroopListEditor(DataRepository repo)
    {
        _repo = repo;
        Dock = DockStyle.Fill;
        InitializeUI();
        RefreshList();
    }

    private void InitializeUI()
    {
        // === Left: Grid ===
        _grid = new DataGridView
        {
            Dock = DockStyle.Left, Width = 380,
            AllowUserToAddRows = false, AllowUserToDeleteRows = false,
            ReadOnly = true, SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            MultiSelect = false, RowHeadersVisible = false
        };
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "ID", Width = 36 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "StrID", Width = 160 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Name", Width = 100 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Lv", Width = 36 });
        _grid.SelectionChanged += OnSelectionChanged;

        // === Right: Properties ===
        _props = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(8) };
        var layout = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 2 };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        _txtStrId = AddF(layout, "StrID:");
        _txtName = AddF(layout, "Name:");
        _txtPlural = AddF(layout, "Plural:");
        _txtLevel = AddF(layout, "Level:");
        _txtStr = AddF(layout, "Strength:");
        _txtAgi = AddF(layout, "Agility:");
        _txtInt = AddF(layout, "Intelligence:");
        _txtCha = AddF(layout, "Charisma:");
        _txt1h = AddF(layout, "1H Wpn:");
        _txt2h = AddF(layout, "2H Wpn:");
        _txtPole = AddF(layout, "Polearm:");
        _txtArch = AddF(layout, "Archery:");
        _txtXbow = AddF(layout, "Crossbow:");
        _txtThr = AddF(layout, "Throwing:");
        _txtFire = AddF(layout, "Firearm:");
        _cmbFaction = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
        layout.Controls.Add(new Label { Text = "Faction:", TextAlign = ContentAlignment.MiddleRight, Dock = DockStyle.Fill });
        layout.Controls.Add(_cmbFaction);

        var btnApply = new Button { Text = "&Apply", Width = 100 };
        btnApply.Click += OnApply;
        layout.Controls.Add(btnApply);
        _lblStatus = new Label { Text = "", AutoSize = true };
        layout.Controls.Add(_lblStatus);

        _props.Controls.Add(layout);

        // Layout
        var split = new SplitContainer { Dock = DockStyle.Fill, SplitterDistance = 380 };
        split.Panel1.Controls.Add(_grid);
        split.Panel2.Controls.Add(_props);
        Controls.Add(split);
    }

    private static TextBox AddF(TableLayoutPanel p, string l) { var tb = new TextBox { Dock = DockStyle.Fill }; p.Controls.Add(new Label { Text = l, TextAlign = ContentAlignment.MiddleRight, Dock = DockStyle.Fill }); p.Controls.Add(tb); return tb; }

    private void RefreshList()
    {
        _grid.Rows.Clear();
        _cmbFaction.Items.Clear();
        foreach (var f in _repo.Factions) _cmbFaction.Items.Add($"[{f.Id}] {f.Name}");
        foreach (var t in _repo.Troops) _grid.Rows.Add(t.Id, t.StrId, t.Name, t.Attributes.Level);
    }

    private void OnSelectionChanged(object? sender, EventArgs e)
    {
        if (_grid.SelectedRows.Count == 0) return;
        _currentIndex = _grid.SelectedRows[0].Index;
        if (_currentIndex >= _repo.Troops.Count) return;
        var t = _repo.Troops[_currentIndex];
        _txtStrId.Text = t.StrId; _txtName.Text = t.Name; _txtPlural.Text = t.PluralName;
        _txtLevel.Text = t.Attributes.Level.ToString();
        _txtStr.Text = t.Attributes.Str.ToString(); _txtAgi.Text = t.Attributes.Agi.ToString();
        _txtInt.Text = t.Attributes.Int.ToString(); _txtCha.Text = t.Attributes.Cha.ToString();
        _txt1h.Text = t.Proficiencies.OneHanded.ToString(); _txt2h.Text = t.Proficiencies.TwoHanded.ToString();
        _txtPole.Text = t.Proficiencies.Polearm.ToString(); _txtArch.Text = t.Proficiencies.Archery.ToString();
        _txtXbow.Text = t.Proficiencies.Crossbow.ToString(); _txtThr.Text = t.Proficiencies.Throwing.ToString();
        _txtFire.Text = t.Proficiencies.Firearm.ToString();
        if (t.Faction < _cmbFaction.Items.Count) _cmbFaction.SelectedIndex = (int)t.Faction;
        _lblStatus.Text = "";
    }

    private void OnApply(object? sender, EventArgs e)
    {
        if (_currentIndex < 0 || _currentIndex >= _repo.Troops.Count) return;
        try
        {
            var t = _repo.Troops[_currentIndex];
            t.StrId = _txtStrId.Text; t.Name = _txtName.Text; t.PluralName = _txtPlural.Text;
            t.Attributes = new Attributes
            {
                Level = int.TryParse(_txtLevel.Text, out var lv) ? lv : 0,
                Str = int.TryParse(_txtStr.Text, out var s) ? s : 0,
                Agi = int.TryParse(_txtAgi.Text, out var a) ? a : 0,
                Int = int.TryParse(_txtInt.Text, out var i) ? i : 0,
                Cha = int.TryParse(_txtCha.Text, out var c) ? c : 0
            };
            t.Proficiencies = new WeaponProficiencies
            {
                OneHanded = long.TryParse(_txt1h.Text, out var v1) ? v1 : 0,
                TwoHanded = long.TryParse(_txt2h.Text, out var v2) ? v2 : 0,
                Polearm = long.TryParse(_txtPole.Text, out var v3) ? v3 : 0,
                Archery = long.TryParse(_txtArch.Text, out var v4) ? v4 : 0,
                Crossbow = long.TryParse(_txtXbow.Text, out var v5) ? v5 : 0,
                Throwing = long.TryParse(_txtThr.Text, out var v6) ? v6 : 0,
                Firearm = long.TryParse(_txtFire.Text, out var v7) ? v7 : 0
            };
            if (_cmbFaction.SelectedIndex >= 0) t.Faction = _cmbFaction.SelectedIndex;
            t.IsEdited = true;
            _grid.Rows[_currentIndex].Cells[1].Value = t.StrId;
            _grid.Rows[_currentIndex].Cells[2].Value = t.Name;
            _lblStatus.Text = "Applied";
            _repo.NotifyTroopChanged(_currentIndex);
        }
        catch (Exception ex) { _lblStatus.Text = $"Error: {ex.Message}"; }
    }
}
