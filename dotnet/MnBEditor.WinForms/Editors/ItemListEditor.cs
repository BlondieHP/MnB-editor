using MnBEditor.Core;
using MnBEditor.Core.Models;

namespace MnBEditor.WinForms.Editors;

/// <summary>
/// Item editor — faithfully replicates the original frmItems VB6 form.
/// Left: ListView with search. Right: Two tab pages (Basic & Mod/Flags).
/// Bottom: Reset, Apply, Output buttons.
/// </summary>
public class ItemListEditor : UserControl
{
    private readonly DataRepository _repo;
    private ListView _lstItems = null!;
    private TextBox _txtSearch = null!;
    private TabControl _tabs = null!;
    private Label _lblStatus = null!;
    private List<Item> _filtered = new();
    private int _selIdx = -1;

    // === Basic tab controls ===
    private TextBox _tbDbName = null!, _tbDispName = null!, _tbTexName = null!;
    private TextBox _tbPrice = null!, _tbWeight = null!, _tbAbundance = null!;
    private TextBox _tbHeadArmor = null!, _tbBodyArmor = null!, _tbLegArmor = null!;
    private TextBox _tbDifficulty = null!, _tbHitPoints = null!, _tbSpeed = null!;
    private TextBox _tbMissileSpd = null!, _tbWeaponLen = null!, _tbMaxAmmo = null!;
    private TextBox _tbThrustDmg = null!, _tbSwingDmg = null!;
    private TextBox _tbTypeFlags = null!, _tbActionFlags = null!;
    private TextBox _tbPrefix = null!, _tbMeshCount = null!;
    private ListBox _lbMeshes = null!;
    private CheckedListBox _clbFlags = null!;

    // === Mod tab controls ===
    private TextBox _tbFactionCount = null!;
    private ListBox _lbFactions = null!;
    private TextBox _tbTriggerCount = null!;

    private static readonly (string, int)[] ItemFlags =
        [("Unique",12),("Always Loot",13),("No Parry",14),("Default Ammo",15),
         ("Merchandise",16),("Wooden Attack",17),("Wooden Parry",18),("Food",19),
         ("Cant Reload Horse",20),("Two Handed",21),("Primary",22),("Secondary",23),
         ("Covers Legs",24),("Consumable",25),("Bonus vs Shield",26),("Penalty Shield",27),
         ("Cant Use Horse",28),("Civilian",29),("Fit to Head",30),("Covers Head",31),
         ("Crush Through",32),("Knock Back",33),("Remove on Use",34),("Unbalanced",35),
         ("Covers Beard",36),("No Pickup",37),("Can Knock Down",38)];

    public ItemListEditor(DataRepository repo) { _repo = repo; Dock = DockStyle.Fill; BuildUI(); RefreshList(); }

    void BuildUI()
    {
        // === Left: ListView + Search ===
        var left = new Panel { Dock = DockStyle.Left, Width = 380 };
        var searchBar = new Panel { Dock = DockStyle.Top, Height = 30 };
        _txtSearch = new TextBox { Width = 280, Dock = DockStyle.Left };
        var btnSearch = new Button { Text = "Find", Width = 50, Dock = DockStyle.Left };
        btnSearch.Click += (_, _) => { ApplyFilter(); };
        _txtSearch.KeyDown += (_, e) => { if (e.KeyCode == Keys.Enter) ApplyFilter(); };
        searchBar.Controls.Add(btnSearch);
        searchBar.Controls.Add(_txtSearch);

        _lstItems = new ListView { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true,
            MultiSelect = false, HideSelection = false, HeaderStyle = ColumnHeaderStyle.Nonclickable };
        _lstItems.Columns.Add("ID", 40);
        _lstItems.Columns.Add("DB Name", 160);
        _lstItems.Columns.Add("Display Name", 130);
        _lstItems.Columns.Add("Price", 45);
        _lstItems.SelectedIndexChanged += OnItemSelected;
        left.Controls.Add(_lstItems);
        left.Controls.Add(searchBar);

        // === Right: TabControl ===
        _tabs = new TabControl { Dock = DockStyle.Fill };
        _tabs.TabPages.Add(BuildBasicTab());
        _tabs.TabPages.Add(BuildModTab());

        // === Bottom: Action buttons ===
        var bottom = new Panel { Dock = DockStyle.Bottom, Height = 40 };
        var btnReset = new Button { Text = "&Reset", Width = 90, Dock = DockStyle.Right };
        btnReset.Click += (_, _) => { if (_selIdx >= 0 && _selIdx < _filtered.Count) LoadItem(_filtered[_selIdx]); };
        var btnApply = new Button { Text = "&Apply", Width = 90, Dock = DockStyle.Right, BackColor = Color.LightCoral };
        btnApply.Click += (_, _) => ApplyChanges();
        var btnOutput = new Button { Text = "&Output", Width = 90, Dock = DockStyle.Right, BackColor = Color.LightGreen };
        btnOutput.Click += (_, _) => { if (_selIdx >= 0) MessageBox.Show(FileSaver.FormatItemPublic(_filtered[_selIdx]), "Output Item"); };
        _lblStatus = new Label { Text = "", AutoSize = true, Dock = DockStyle.Left, TextAlign = ContentAlignment.MiddleLeft };
        bottom.Controls.Add(btnOutput);
        bottom.Controls.Add(btnApply);
        bottom.Controls.Add(btnReset);
        bottom.Controls.Add(_lblStatus);

        // === Layout ===
        var split = new SplitContainer { Dock = DockStyle.Fill, SplitterDistance = 380, Panel1MinSize = 200 };
        split.Panel1.Controls.Add(left);
        split.Panel2.Controls.Add(_tabs);
        Controls.Add(split);
        Controls.Add(bottom);
    }

    // ====== Basic Tab ======
    TabPage BuildBasicTab()
    {
        var page = new TabPage("Basic Properties") { AutoScroll = true };
        var p = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 4, Padding = new Padding(6) };
        p.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 85));
        p.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));
        p.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 85));
        p.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));

        _tbDbName = AddRow(p, "DB Name:"); _tbDispName = AddRow(p, "Display:");
        _tbTexName = AddRow(p, "Texture:"); _tbPrice = AddRow(p, "Price:");
        _tbWeight = AddRow(p, "Weight:"); _tbAbundance = AddRow(p, "Abundance:");
        _tbPrefix = AddRow(p, "Prefix:"); _tbMeshCount = AddRow(p, "Mesh Count:");
        _tbHeadArmor = AddRow(p, "Head Armor:"); _tbBodyArmor = AddRow(p, "Body Armor:");
        _tbLegArmor = AddRow(p, "Leg Armor:"); _tbDifficulty = AddRow(p, "Difficulty:");
        _tbHitPoints = AddRow(p, "Hit Points:"); _tbSpeed = AddRow(p, "Speed:");
        _tbMissileSpd = AddRow(p, "Missile Spd:"); _tbWeaponLen = AddRow(p, "Weapon Len:");
        _tbMaxAmmo = AddRow(p, "Max Ammo:"); _tbThrustDmg = AddRow(p, "Thrust Dmg:");
        _tbSwingDmg = AddRow(p, "Swing Dmg:"); _tbTypeFlags = AddRow(p, "Type Flags:");
        _tbActionFlags = AddRow(p, "Action:");

        // Meshes list
        p.Controls.Add(new Label { Text = "Meshes:", TextAlign = ContentAlignment.TopRight, Dock = DockStyle.Fill });
        _lbMeshes = new ListBox { Height = 60, Dock = DockStyle.Fill };
        p.Controls.Add(_lbMeshes); p.SetColumnSpan(_lbMeshes, 3);

        // Item flags checklist
        p.Controls.Add(new Label { Text = "Flags:", TextAlign = ContentAlignment.TopRight, Dock = DockStyle.Fill });
        _clbFlags = new CheckedListBox { Height = 160, Dock = DockStyle.Fill, MultiColumn = true, ColumnWidth = 130 };
        foreach (var (name, _) in ItemFlags) _clbFlags.Items.Add(name);
        p.Controls.Add(_clbFlags); p.SetColumnSpan(_clbFlags, 3);

        page.Controls.Add(p);
        return page;
    }

    // ====== Mod Tab ======
    TabPage BuildModTab()
    {
        var page = new TabPage("Mod Data") { AutoScroll = true };
        var p = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 2, Padding = new Padding(6) };
        p.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));
        p.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        _tbFactionCount = AddRow(p, "Faction Count:");
        p.Controls.Add(new Label { Text = "Factions:", TextAlign = ContentAlignment.TopRight, Dock = DockStyle.Fill });
        _lbFactions = new ListBox { Height = 80, Dock = DockStyle.Fill }; p.Controls.Add(_lbFactions);

        _tbTriggerCount = AddRow(p, "Trigger Count:");
        p.Controls.Add(new Label { Text = "Triggers:", TextAlign = ContentAlignment.TopRight, Dock = DockStyle.Fill });
        p.Controls.Add(new Label { Text = "(triggers listed in Output)", Dock = DockStyle.Fill });

        page.Controls.Add(p);
        return page;
    }

    // ====== Helpers ======
    TextBox AddRow(TableLayoutPanel p, string label)
    {
        p.Controls.Add(new Label { Text = label, TextAlign = ContentAlignment.MiddleRight, Dock = DockStyle.Fill });
        var tb = new TextBox { Dock = DockStyle.Fill };
        p.Controls.Add(tb);
        return tb;
    }

    void RefreshList() { _lstItems.Items.Clear(); _filtered.Clear(); _filtered.AddRange(_repo.Items); PopulateList(); }
    void PopulateList() { _lstItems.BeginUpdate(); _lstItems.Items.Clear(); foreach (var it in _filtered) _lstItems.Items.Add(new ListViewItem([it.Id.ToString(), it.DbName, it.DisplayName, it.Price.ToString()])); _lstItems.EndUpdate(); }
    void ApplyFilter() { var q = _txtSearch.Text.Trim(); _filtered = q.Length == 0 ? new List<Item>(_repo.Items) : _repo.Items.Where(i => i.DbName.Contains(q, StringComparison.OrdinalIgnoreCase) || i.DisplayName.Contains(q, StringComparison.OrdinalIgnoreCase)).ToList(); Populate(); }
    void Populate() { PopulateList(); _lblStatus.Text = $"Showing {_filtered.Count}/{_repo.Items.Count} items"; }

    // ====== Selection ======
    void OnItemSelected(object? s, EventArgs e)
    {
        if (_lstItems.SelectedIndices.Count == 0) return;
        _selIdx = _lstItems.SelectedIndices[0];
        if (_selIdx < 0 || _selIdx >= _filtered.Count) return;
        LoadItem(_filtered[_selIdx]);
    }

    void LoadItem(Item it)
    {
        _tbDbName.Text = it.DbName; _tbDispName.Text = it.DisplayName; _tbTexName.Text = it.TextureName;
        _tbPrice.Text = it.Price.ToString(); _tbWeight.Text = it.Weight; _tbAbundance.Text = it.Abundance.ToString();
        _tbHeadArmor.Text = it.HeadArmor.ToString(); _tbBodyArmor.Text = it.BodyArmor.ToString(); _tbLegArmor.Text = it.LegArmor.ToString();
        _tbDifficulty.Text = it.Difficulty.ToString(); _tbHitPoints.Text = it.HitPoints.ToString(); _tbSpeed.Text = it.SpeedRating.ToString();
        _tbMissileSpd.Text = it.MissileSpeed.ToString(); _tbWeaponLen.Text = it.WeaponLength.ToString(); _tbMaxAmmo.Text = it.MaxAmmo.ToString();
        _tbThrustDmg.Text = it.ThrustDamage.ToString(); _tbSwingDmg.Text = it.SwingDamage.ToString();
        _tbTypeFlags.Text = it.ItemType; _tbActionFlags.Text = it.Action;
        _tbPrefix.Text = it.Prefix; _tbMeshCount.Text = it.MeshCount.ToString();
        _tbFactionCount.Text = it.FactionCount.ToString(); _tbTriggerCount.Text = it.TriggerCount.ToString();

        _lbMeshes.Items.Clear();
        for (int i = 0; i < it.MeshCount && i < it.MeshNames.Count; i++)
            _lbMeshes.Items.Add($"{it.MeshNames[i]} ({it.MeshParams[i]})");

        _lbFactions.Items.Clear();
        for (int i = 0; i < it.FactionCount && i < it.Factions.Count; i++)
            _lbFactions.Items.Add($"FacID: {it.Factions[i].Id}");

        // Update flag checkboxes
        var flags = TryParseHexOrDec(it.ItemType);
        for (int i = 0; i < ItemFlags.Length && i < _clbFlags.Items.Count; i++)
            _clbFlags.SetItemChecked(i, (flags & (1L << ItemFlags[i].Item2)) != 0);

        _lblStatus.Text = $"Loaded: {it.DbName}";
    }

    // ====== Apply ======
    void ApplyChanges()
    {
        if (_selIdx < 0 || _selIdx >= _filtered.Count) return;
        try
        {
            var it = _filtered[_selIdx];
            it.DbName = _tbDbName.Text; it.DisplayName = _tbDispName.Text; it.TextureName = _tbTexName.Text;
            it.Price = ParseL(_tbPrice.Text); it.Weight = _tbWeight.Text; it.Abundance = ParseL(_tbAbundance.Text);
            it.HeadArmor = ParseL(_tbHeadArmor.Text); it.BodyArmor = ParseL(_tbBodyArmor.Text); it.LegArmor = ParseL(_tbLegArmor.Text);
            it.Difficulty = ParseL(_tbDifficulty.Text); it.HitPoints = ParseL(_tbHitPoints.Text); it.SpeedRating = ParseL(_tbSpeed.Text);
            it.MissileSpeed = ParseL(_tbMissileSpd.Text); it.WeaponLength = ParseL(_tbWeaponLen.Text); it.MaxAmmo = ParseL(_tbMaxAmmo.Text);
            it.ThrustDamage = ParseL(_tbThrustDmg.Text); it.SwingDamage = ParseL(_tbSwingDmg.Text);
            it.ItemType = _tbTypeFlags.Text; it.Action = _tbActionFlags.Text;
            it.Prefix = _tbPrefix.Text; it.MeshCount = ParseL(_tbMeshCount.Text);
            it.FactionCount = ParseL(_tbFactionCount.Text); it.TriggerCount = ParseL(_tbTriggerCount.Text);
            it.IsEdited = true;

            if (_lstItems.SelectedIndices.Count > 0)
            {
                var lvi = _lstItems.Items[_lstItems.SelectedIndices[0]];
                lvi.SubItems[1].Text = it.DbName;
                lvi.SubItems[2].Text = it.DisplayName;
                lvi.SubItems[3].Text = it.Price.ToString();
            }
            _lblStatus.Text = "Applied ✓";
            _repo.NotifyItemChanged((int)it.Id);
        }
        catch (Exception ex) { _lblStatus.Text = $"Error: {ex.Message}"; }
    }

    static long ParseL(string s) => long.TryParse(s, out var v) ? v : 0;
    static long TryParseHexOrDec(string s) { if (string.IsNullOrEmpty(s)) return 0; if (s.StartsWith("0x") || s.Length > 10) return long.TryParse(s, System.Globalization.NumberStyles.HexNumber, null, out var h) ? h : 0; return long.TryParse(s, out var d) ? d : 0; }
}
