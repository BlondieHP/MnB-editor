using MnBEditor.Core;
using MnBEditor.Core.Models;

namespace MnBEditor.WinForms.Editors;

/// <summary>
/// Troop editor — replicates frmTroops with inventory grid, attributes,
/// proficiencies, skills, face codes, and upgrade tree.
/// </summary>
public class TroopListEditor : UserControl
{
    private readonly DataRepository _repo;
    private ListView _lst = null!;
    private TextBox _txtSearch = null!;
    private TabControl _tabs = null!;
    private Label _lblStatus = null!;
    private List<Troop> _filtered = new();
    private int _selIdx = -1;

    // Basic tab
    private TextBox _tbStrId = null!, _tbName = null!, _tbPlural = null!, _tbFlags = null!;
    private TextBox _tbScene = null!, _tbReserved = null!, _tbFaction = null!;
    private TextBox _tbUpgrade1 = null!, _tbUpgrade2 = null!;
    private Label _lblUp1Name = null!, _lblUp2Name = null!;

    // Attributes tab
    private TextBox _tbLevel = null!, _tbStr = null!, _tbAgi = null!, _tbInt = null!, _tbCha = null!;
    private TextBox _tb1H = null!, _tb2H = null!, _tbPole = null!, _tbArch = null!;
    private TextBox _tbXbow = null!, _tbThr = null!, _tbFire = null!;
    private TextBox _tbSk0 = null!, _tbSk1 = null!, _tbSk2 = null!, _tbSk3 = null!, _tbSk4 = null!, _tbSk5 = null!;

    // Inventory tab
    private ListBox _lbInventory = null!;

    // Face tab
    private TextBox _tbFace0 = null!, _tbFace1 = null!, _tbFace2 = null!, _tbFace3 = null!;
    private TextBox _tbFace4 = null!, _tbFace5 = null!, _tbFace6 = null!, _tbFace7 = null!;

    public TroopListEditor(DataRepository repo) { _repo = repo; Dock = DockStyle.Fill; BuildUI(); RefreshList(); }

    void BuildUI()
    {
        var left = new Panel { Dock = DockStyle.Left, Width = 350 };
        var searchBar = new Panel { Dock = DockStyle.Top, Height = 30 };
        _txtSearch = new TextBox { Width = 260, Dock = DockStyle.Left };
        var btnSearch = new Button { Text = "Find", Width = 50, Dock = DockStyle.Left };
        btnSearch.Click += (_, _) => ApplyFilter();
        _txtSearch.KeyDown += (_, e) => { if (e.KeyCode == Keys.Enter) ApplyFilter(); };
        searchBar.Controls.Add(btnSearch); searchBar.Controls.Add(_txtSearch);

        _lst = new ListView { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true,
            MultiSelect = false, HideSelection = false, HeaderStyle = ColumnHeaderStyle.Nonclickable };
        _lst.Columns.Add("ID", 36); _lst.Columns.Add("StrID", 150);
        _lst.Columns.Add("Name", 100); _lst.Columns.Add("Lv", 32);
        _lst.SelectedIndexChanged += OnSelect;
        left.Controls.Add(_lst); left.Controls.Add(searchBar);

        _tabs = new TabControl { Dock = DockStyle.Fill };
        _tabs.TabPages.Add(BuildBasicTab());
        _tabs.TabPages.Add(BuildAttrTab());
        _tabs.TabPages.Add(BuildInvTab());
        _tabs.TabPages.Add(BuildFaceTab());

        var bottom = new Panel { Dock = DockStyle.Bottom, Height = 40 };
        var btnReset = new Button { Text = "&Reset", Width = 80, Dock = DockStyle.Right };
        btnReset.Click += (_, _) => { if (_selIdx >= 0 && _selIdx < _filtered.Count) LoadTroop(_filtered[_selIdx]); };
        var btnApply = new Button { Text = "&Apply", Width = 80, Dock = DockStyle.Right, BackColor = Color.LightCoral };
        btnApply.Click += (_, _) => Apply();
        _lblStatus = new Label { Text = "", AutoSize = true, Dock = DockStyle.Left };
        bottom.Controls.Add(btnApply); bottom.Controls.Add(btnReset); bottom.Controls.Add(_lblStatus);

        var split = new SplitContainer { Dock = DockStyle.Fill, SplitterDistance = 350, Panel1MinSize = 200 };
        split.Panel1.Controls.Add(left); split.Panel2.Controls.Add(_tabs);
        Controls.Add(split); Controls.Add(bottom);
    }

    // ====== Tabs ======
    TabPage BuildBasicTab()
    {
        var p = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 4, Padding = new Padding(6) };
        p.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 85));
        p.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));
        p.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 85));
        p.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));

        _tbStrId = AddTF(p, "StrID:"); _tbName = AddTF(p, "Name:");
        _tbPlural = AddTF(p, "Plural:"); _tbFlags = AddTF(p, "Flags:");
        _tbScene = AddTF(p, "Scene:"); _tbReserved = AddTF(p, "Reserved:");
        _tbFaction = AddTF(p, "Faction:"); _tbUpgrade1 = AddTF(p, "Upgrade 1:");
        _lblUp1Name = new Label { Text = "", Dock = DockStyle.Fill, ForeColor = Color.Gray }; p.Controls.Add(_lblUp1Name);
        _tbUpgrade2 = AddTF(p, "Upgrade 2:");
        _lblUp2Name = new Label { Text = "", Dock = DockStyle.Fill, ForeColor = Color.Gray }; p.Controls.Add(_lblUp2Name);

        var page = new TabPage("Basic") { AutoScroll = true }; page.Controls.Add(p); return page;
    }

    TabPage BuildAttrTab()
    {
        var p = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 4, Padding = new Padding(6) };
        p.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 85));
        p.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));
        p.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 85));
        p.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));

        _tbLevel = AddTF(p, "Level:"); _tbStr = AddTF(p, "Strength:");
        _tbAgi = AddTF(p, "Agility:"); _tbInt = AddTF(p, "Intelligence:");
        _tbCha = AddTF(p, "Charisma:");
        AddSep(p, "── Proficiencies ──");
        _tb1H = AddTF(p, "One Handed:"); _tb2H = AddTF(p, "Two Handed:");
        _tbPole = AddTF(p, "Polearm:"); _tbArch = AddTF(p, "Archery:");
        _tbXbow = AddTF(p, "Crossbow:"); _tbThr = AddTF(p, "Throwing:");
        _tbFire = AddTF(p, "Firearm:");
        AddSep(p, "── Skills ──");
        _tbSk0 = AddTF(p, "Skill 0:"); _tbSk1 = AddTF(p, "Skill 1:");
        _tbSk2 = AddTF(p, "Skill 2:"); _tbSk3 = AddTF(p, "Skill 3:");
        _tbSk4 = AddTF(p, "Skill 4:"); _tbSk5 = AddTF(p, "Skill 5:");

        var page = new TabPage("Attrs/Profs") { AutoScroll = true }; page.Controls.Add(p); return page;
    }

    TabPage BuildInvTab()
    {
        var page = new TabPage("Inventory") { AutoScroll = true };
        _lbInventory = new ListBox { Dock = DockStyle.Fill, Font = new Font("Consolas", 9) };
        page.Controls.Add(_lbInventory);
        return page;
    }

    TabPage BuildFaceTab()
    {
        var p = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 4, Padding = new Padding(6) };
        p.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80));
        p.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));
        p.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80));
        p.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));

        _tbFace0 = AddTF(p, "Face 0:"); _tbFace1 = AddTF(p, "Face 1:");
        _tbFace2 = AddTF(p, "Face 2:"); _tbFace3 = AddTF(p, "Face 3:");
        _tbFace4 = AddTF(p, "Face 4:"); _tbFace5 = AddTF(p, "Face 5:");
        _tbFace6 = AddTF(p, "Face 6:"); _tbFace7 = AddTF(p, "Face 7:");

        var page = new TabPage("Face Codes") { AutoScroll = true }; page.Controls.Add(p); return page;
    }

    // ====== Helpers ======
    TextBox AddTF(TableLayoutPanel p, string label)
    {
        p.Controls.Add(new Label { Text = label, TextAlign = ContentAlignment.MiddleRight, Dock = DockStyle.Fill });
        var tb = new TextBox { Dock = DockStyle.Fill }; p.Controls.Add(tb); return tb;
    }
    void AddSep(TableLayoutPanel p, string text) { var l = new Label { Text = text, TextAlign = ContentAlignment.MiddleCenter, Font = new Font(Font.FontFamily, Font.Size, FontStyle.Bold) }; p.Controls.Add(l); p.SetColumnSpan(l, 4); }

    void RefreshList() { _lst.Items.Clear(); _filtered.Clear(); _filtered.AddRange(_repo.Troops); PopulateList(); }
    void PopulateList() { _lst.BeginUpdate(); _lst.Items.Clear(); foreach (var t in _filtered) _lst.Items.Add(new ListViewItem([t.Id.ToString(), t.StrId, t.Name, t.Attributes.Level.ToString()])); _lst.EndUpdate(); }
    void ApplyFilter() { var q = _txtSearch.Text.Trim(); _filtered = q.Length == 0 ? new List<Troop>(_repo.Troops) : _repo.Troops.Where(t => t.StrId.Contains(q, StringComparison.OrdinalIgnoreCase) || t.Name.Contains(q, StringComparison.OrdinalIgnoreCase)).ToList(); PopulateList(); _lblStatus.Text = $"Showing {_filtered.Count}/{_repo.Troops.Count}"; }

    void OnSelect(object? s, EventArgs e)
    {
        if (_lst.SelectedIndices.Count == 0) return;
        _selIdx = _lst.SelectedIndices[0];
        if (_selIdx >= 0 && _selIdx < _filtered.Count) LoadTroop(_filtered[_selIdx]);
    }

    void LoadTroop(Troop t)
    {
        _tbStrId.Text = t.StrId; _tbName.Text = t.Name; _tbPlural.Text = t.PluralName;
        _tbFlags.Text = t.Flags; _tbScene.Text = t.Scene.ToString(); _tbReserved.Text = t.Reserved.ToString();
        _tbFaction.Text = t.Faction.ToString(); _tbUpgrade1.Text = t.Upgrade1.ToString(); _tbUpgrade2.Text = t.Upgrade2.ToString();
        _lblUp1Name.Text = t.Upgrade1 < _repo.Troops.Count ? $"→ {_repo.Troops[(int)t.Upgrade1].Name}" : "";
        _lblUp2Name.Text = t.Upgrade2 < _repo.Troops.Count ? $"→ {_repo.Troops[(int)t.Upgrade2].Name}" : "";

        _tbLevel.Text = t.Attributes.Level.ToString(); _tbStr.Text = t.Attributes.Str.ToString();
        _tbAgi.Text = t.Attributes.Agi.ToString(); _tbInt.Text = t.Attributes.Int.ToString(); _tbCha.Text = t.Attributes.Cha.ToString();
        _tb1H.Text = t.Proficiencies.OneHanded.ToString(); _tb2H.Text = t.Proficiencies.TwoHanded.ToString();
        _tbPole.Text = t.Proficiencies.Polearm.ToString(); _tbArch.Text = t.Proficiencies.Archery.ToString();
        _tbXbow.Text = t.Proficiencies.Crossbow.ToString(); _tbThr.Text = t.Proficiencies.Throwing.ToString();
        _tbFire.Text = t.Proficiencies.Firearm.ToString();
        _tbSk0.Text = t.Skills[0].ToString(); _tbSk1.Text = t.Skills[1].ToString(); _tbSk2.Text = t.Skills[2].ToString();
        _tbSk3.Text = t.Skills[3].ToString(); _tbSk4.Text = t.Skills[4].ToString(); _tbSk5.Text = t.Skills[5].ToString();

        _lbInventory.Items.Clear();
        for (int i = 0; i < 64; i++)
        {
            if (t.Inventory[i].ItemId > 0 && t.Inventory[i].ItemId < _repo.Items.Count)
                _lbInventory.Items.Add($"[{i,2}] {_repo.Items[(int)t.Inventory[i].ItemId].DbName,-35} x{t.Inventory[i].Quantity}");
            else if (t.Inventory[i].ItemId > 0)
                _lbInventory.Items.Add($"[{i,2}] UnknownItem#{t.Inventory[i].ItemId} x{t.Inventory[i].Quantity}");
        }

        _tbFace0.Text = t.Face[0]; _tbFace1.Text = t.Face[1]; _tbFace2.Text = t.Face[2]; _tbFace3.Text = t.Face[3];
        _tbFace4.Text = t.Face[4]; _tbFace5.Text = t.Face[5]; _tbFace6.Text = t.Face[6]; _tbFace7.Text = t.Face[7];

        _lblStatus.Text = $"Loaded: {t.StrId} Lv{t.Attributes.Level}";
    }

    void Apply()
    {
        if (_selIdx < 0 || _selIdx >= _filtered.Count) return;
        try
        {
            var t = _filtered[_selIdx];
            t.StrId = _tbStrId.Text; t.Name = _tbName.Text; t.PluralName = _tbPlural.Text;
            t.Flags = _tbFlags.Text; t.Scene = P(_tbScene.Text); t.Reserved = P(_tbReserved.Text);
            t.Faction = P(_tbFaction.Text); t.Upgrade1 = P(_tbUpgrade1.Text); t.Upgrade2 = P(_tbUpgrade2.Text);
            t.Attributes = new Attributes { Level = I(_tbLevel.Text), Str = I(_tbStr.Text), Agi = I(_tbAgi.Text), Int = I(_tbInt.Text), Cha = I(_tbCha.Text) };
            t.Proficiencies = new WeaponProficiencies { OneHanded = P(_tb1H.Text), TwoHanded = P(_tb2H.Text), Polearm = P(_tbPole.Text), Archery = P(_tbArch.Text), Crossbow = P(_tbXbow.Text), Throwing = P(_tbThr.Text), Firearm = P(_tbFire.Text) };
            t.Skills[0] = P(_tbSk0.Text); t.Skills[1] = P(_tbSk1.Text); t.Skills[2] = P(_tbSk2.Text);
            t.Skills[3] = P(_tbSk3.Text); t.Skills[4] = P(_tbSk4.Text); t.Skills[5] = P(_tbSk5.Text);
            for (int f = 0; f < 8; f++) t.Face[f] = f switch { 0 => _tbFace0.Text, 1 => _tbFace1.Text, 2 => _tbFace2.Text, 3 => _tbFace3.Text, 4 => _tbFace4.Text, 5 => _tbFace5.Text, 6 => _tbFace6.Text, _ => _tbFace7.Text };
            t.IsEdited = true;

            if (_lst.SelectedIndices.Count > 0) { var lvi = _lst.Items[_lst.SelectedIndices[0]]; lvi.SubItems[1].Text = t.StrId; lvi.SubItems[2].Text = t.Name; lvi.SubItems[3].Text = t.Attributes.Level.ToString(); }
            _lblStatus.Text = "Applied ✓"; _repo.NotifyTroopChanged((int)t.Id);
        }
        catch (Exception ex) { _lblStatus.Text = $"Error: {ex.Message}"; }
    }

    static long P(string s) => long.TryParse(s, out var v) ? v : 0;
    static int I(string s) => int.TryParse(s, out var v) ? v : 0;
}
