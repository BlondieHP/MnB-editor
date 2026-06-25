using MnBEditor.Core;
using MnBEditor.Core.Models;

namespace MnBEditor.WinForms.Editors;

/// <summary>
/// Item editor — mimics the original frmItems layout.
/// Left: DataGridView list | Right: Property panels | Bottom: Action buttons
/// </summary>
public class ItemListEditor : UserControl
{
    private readonly DataRepository _repo;
    private DataGridView _grid = null!;
    private Panel _props = null!;
    private TextBox _txtName = null!, _txtDisplayName = null!, _txtPrice = null!, _txtWeight = null!;
    private TextBox _txtAbundance = null!, _txtHeadArmor = null!, _txtBodyArmor = null!, _txtLegArmor = null!;
    private TextBox _txtDifficulty = null!, _txtHitPoints = null!, _txtSpeed = null!, _txtMissileSpeed = null!;
    private TextBox _txtWeaponLen = null!, _txtMaxAmmo = null!, _txtThrust = null!, _txtSwing = null!;
    private ComboBox _cmbType = null!;
    private Label _lblStatus = null!;
    private int _currentIndex = -1;

    private static readonly string[] ItemTypes =
        ["horse", "one_handed", "two_handed", "polearm", "arrows", "bolts", "shield", "bow",
         "crossbow", "thrown", "goods", "head_armor", "body_armor", "foot_armor", "hand_armor",
         "pistol", "musket", "bullets", "animal", "book"];

    public ItemListEditor(DataRepository repo)
    {
        _repo = repo;
        Dock = DockStyle.Fill;
        InitializeUI();
        RefreshList();
    }

    private void InitializeUI()
    {
        // === Left: DataGridView ===
        _grid = new DataGridView
        {
            Dock = DockStyle.Left, Width = 420,
            AllowUserToAddRows = false, AllowUserToDeleteRows = false,
            ReadOnly = true, SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            MultiSelect = false, RowHeadersVisible = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        };
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "ID", Width = 40, DataPropertyName = "Id" });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "DB Name", Width = 180, DataPropertyName = "DbName" });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Display Name", Width = 140, DataPropertyName = "DisplayName" });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Price", Width = 55, DataPropertyName = "Price" });
        _grid.SelectionChanged += OnGridSelectionChanged;

        // === Center: Property Panel ===
        _props = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(8) };
        var layout = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 2, Padding = new Padding(0) };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        _txtName = AddField(layout, "DB Name:");
        _txtDisplayName = AddField(layout, "Display Name:");
        _txtPrice = AddField(layout, "Price:");
        _txtWeight = AddField(layout, "Weight:");
        _cmbType = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
        foreach (var t in ItemTypes) _cmbType.Items.Add(t);
        AddControl(layout, "Type:", _cmbType);
        _txtAbundance = AddField(layout, "Abundance:");
        _txtHeadArmor = AddField(layout, "Head Armor:");
        _txtBodyArmor = AddField(layout, "Body Armor:");
        _txtLegArmor = AddField(layout, "Leg Armor:");
        _txtDifficulty = AddField(layout, "Difficulty:");
        _txtHitPoints = AddField(layout, "Hit Points:");
        _txtSpeed = AddField(layout, "Speed:");
        _txtMissileSpeed = AddField(layout, "Missile Speed:");
        _txtWeaponLen = AddField(layout, "Weapon Length:");
        _txtMaxAmmo = AddField(layout, "Max Ammo:");
        _txtThrust = AddField(layout, "Thrust Dmg:");
        _txtSwing = AddField(layout, "Swing Dmg:");

        var btnApply = new Button { Text = "&Apply", Dock = DockStyle.None, Width = 100 };
        btnApply.Click += OnApply;
        layout.Controls.Add(btnApply);
        _lblStatus = new Label { Text = "", ForeColor = Color.Green, AutoSize = true };
        layout.Controls.Add(_lblStatus);

        _props.Controls.Add(layout);

        // === Bottom: Search ===
        var bottom = new Panel { Dock = DockStyle.Bottom, Height = 30 };
        var txtSearch = new TextBox { Width = 200, Dock = DockStyle.Left };
        var btnSearch = new Button { Text = "Search", Dock = DockStyle.Left, Width = 70 };
        btnSearch.Click += (_, _) => { foreach (DataGridViewRow row in _grid.Rows) row.Visible = row.Cells[1].Value?.ToString()?.Contains(txtSearch.Text, StringComparison.OrdinalIgnoreCase) ?? true; };
        bottom.Controls.Add(btnSearch);
        bottom.Controls.Add(txtSearch);

        // === Layout ===
        var split = new SplitContainer { Dock = DockStyle.Fill, SplitterDistance = 420, Panel1MinSize = 200 };
        split.Panel1.Controls.Add(_grid);
        split.Panel2.Controls.Add(_props);
        Controls.Add(split);
        Controls.Add(bottom);
    }

    private static TextBox AddField(TableLayoutPanel parent, string label)
    {
        var tb = new TextBox { Dock = DockStyle.Fill };
        parent.Controls.Add(new Label { Text = label, TextAlign = ContentAlignment.MiddleRight, Dock = DockStyle.Fill });
        parent.Controls.Add(tb);
        return tb;
    }
    private static void AddControl(TableLayoutPanel parent, string label, Control ctrl)
    {
        parent.Controls.Add(new Label { Text = label, TextAlign = ContentAlignment.MiddleRight, Dock = DockStyle.Fill });
        parent.Controls.Add(ctrl);
    }

    private void RefreshList()
    {
        _grid.Rows.Clear();
        foreach (var it in _repo.Items)
            _grid.Rows.Add(it.Id, it.DbName, it.DisplayName, it.Price);
    }

    private void OnGridSelectionChanged(object? sender, EventArgs e)
    {
        if (_grid.SelectedRows.Count == 0) return;
        _currentIndex = _grid.SelectedRows[0].Index;
        if (_currentIndex >= _repo.Items.Count) return;
        var it = _repo.Items[_currentIndex];
        _txtName.Text = it.DbName;
        _txtDisplayName.Text = it.DisplayName;
        _txtPrice.Text = it.Price.ToString();
        _txtWeight.Text = it.Weight;
        _txtAbundance.Text = it.Abundance.ToString();
        _txtHeadArmor.Text = it.HeadArmor.ToString();
        _txtBodyArmor.Text = it.BodyArmor.ToString();
        _txtLegArmor.Text = it.LegArmor.ToString();
        _txtDifficulty.Text = it.Difficulty.ToString();
        _txtHitPoints.Text = it.HitPoints.ToString();
        _txtSpeed.Text = it.SpeedRating.ToString();
        _txtMissileSpeed.Text = it.MissileSpeed.ToString();
        _txtWeaponLen.Text = it.WeaponLength.ToString();
        _txtMaxAmmo.Text = it.MaxAmmo.ToString();
        _txtThrust.Text = it.ThrustDamage.ToString();
        _txtSwing.Text = it.SwingDamage.ToString();
        _lblStatus.Text = "";
    }

    private void OnApply(object? sender, EventArgs e)
    {
        if (_currentIndex < 0 || _currentIndex >= _repo.Items.Count) return;
        try
        {
            var it = _repo.Items[_currentIndex];
            it.DbName = _txtName.Text;
            it.DisplayName = _txtDisplayName.Text;
            it.Price = long.TryParse(_txtPrice.Text, out var p) ? p : 0;
            it.Weight = _txtWeight.Text;
            it.Abundance = long.TryParse(_txtAbundance.Text, out var a) ? a : 0;
            it.HeadArmor = long.TryParse(_txtHeadArmor.Text, out var ha) ? ha : 0;
            it.BodyArmor = long.TryParse(_txtBodyArmor.Text, out var ba) ? ba : 0;
            it.LegArmor = long.TryParse(_txtLegArmor.Text, out var la) ? la : 0;
            it.Difficulty = long.TryParse(_txtDifficulty.Text, out var d) ? d : 0;
            it.HitPoints = long.TryParse(_txtHitPoints.Text, out var hp) ? hp : 0;
            it.SpeedRating = long.TryParse(_txtSpeed.Text, out var sp) ? sp : 0;
            it.MissileSpeed = long.TryParse(_txtMissileSpeed.Text, out var ms) ? ms : 0;
            it.WeaponLength = long.TryParse(_txtWeaponLen.Text, out var wl) ? wl : 0;
            it.MaxAmmo = long.TryParse(_txtMaxAmmo.Text, out var ma) ? ma : 0;
            it.ThrustDamage = long.TryParse(_txtThrust.Text, out var td) ? td : 0;
            it.SwingDamage = long.TryParse(_txtSwing.Text, out var sd) ? sd : 0;
            it.IsEdited = true;
            _grid.Rows[_currentIndex].Cells[1].Value = it.DbName;
            _grid.Rows[_currentIndex].Cells[2].Value = it.DisplayName;
            _grid.Rows[_currentIndex].Cells[3].Value = it.Price;
            _lblStatus.Text = "Applied";
            _repo.NotifyItemChanged(_currentIndex);
        }
        catch (Exception ex) { _lblStatus.Text = $"Error: {ex.Message}"; _lblStatus.ForeColor = Color.Red; }
    }
}
