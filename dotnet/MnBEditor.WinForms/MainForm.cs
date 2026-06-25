using MnBEditor.Core;
using MnBEditor.Core.Models;
using MnBEditor.WinForms.Editors;

namespace MnBEditor.WinForms;

public class MainForm : Form
{
    private readonly DataRepository _repo = new();
    private TreeView _tree = null!;
    private TabControl _tabs = null!;
    private ToolStrip _toolbar = null!;
    private StatusStrip _status = null!;
    private ToolStripStatusLabel _statusLabel = null!;
    private ItemListEditor? _itemEditor;
    private TroopListEditor? _troopEditor;
    private FactionListEditor? _factionEditor;

    public MainForm()
    {
        Text = "MnBEditor — M&B Warband MOD Editor";
        Size = new Size(1280, 800);
        StartPosition = FormStartPosition.CenterScreen;
        InitializeUI();
    }

    private void InitializeUI()
    {
        // === Toolbar ===
        _toolbar = new ToolStrip();
        _toolbar.Items.Add(new ToolStripButton("Open MOD", null, OnOpenMod) { DisplayStyle = ToolStripItemDisplayStyle.Text });
        _toolbar.Items.Add(new ToolStripSeparator());
        _toolbar.Items.Add(new ToolStripButton("Save All", null, OnSaveAll) { DisplayStyle = ToolStripItemDisplayStyle.Text });
        _toolbar.Items.Add(new ToolStripButton("Validate", null, OnValidate) { DisplayStyle = ToolStripItemDisplayStyle.Text });
        _toolbar.Items.Add(new ToolStripSeparator());
        _toolbar.Items.Add(new ToolStripButton("Stats", null, (_, _) => ShowStats()) { DisplayStyle = ToolStripItemDisplayStyle.Text });

        // === Left Panel: TreeView ===
        _tree = new TreeView { Dock = DockStyle.Left, Width = 220, BorderStyle = BorderStyle.FixedSingle };
        _tree.AfterSelect += OnTreeNodeSelected;

        // === Center: TabControl ===
        _tabs = new TabControl { Dock = DockStyle.Fill };

        // === Status bar ===
        _status = new StatusStrip();
        _statusLabel = new ToolStripStatusLabel("Ready — Open a MOD to begin");
        _status.Items.Add(_statusLabel);

        // === Layout ===
        var split = new SplitContainer { Dock = DockStyle.Fill, SplitterDistance = 220, Panel1MinSize = 150 };
        split.Panel1.Controls.Add(_tree);
        split.Panel2.Controls.Add(_tabs);

        Controls.Add(split);
        Controls.Add(_toolbar);
        Controls.Add(_status);
    }

    // ====== Menu ======

    private void OnOpenMod(object? sender, EventArgs e)
    {
        using var dlg = new FolderBrowserDialog { Description = "Select M&B Warband MOD directory" };
        if (dlg.ShowDialog() != DialogResult.OK) return;

        Cursor = Cursors.WaitCursor;
        _statusLabel.Text = "Loading MOD...";
        _status.Refresh();

        try
        {
            LoadMod(dlg.SelectedPath);
            BuildTree();
            _statusLabel.Text = $"Loaded: {_repo.Items.Count} items, {_repo.Troops.Count} troops, {_repo.Factions.Count} factions, {_repo.Parties.Count} parties";
            Text = $"MnBEditor — {_repo.ModName}";
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Error Loading MOD", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally { Cursor = Cursors.Default; }
    }

    private void LoadMod(string path)
    {
        _repo.ClearAll();
        _repo.ModPath = Path.GetFullPath(path);
        _repo.ModName = Path.GetFileName(Path.GetFullPath(path));

        LoadIfExists(Path.Combine(path, "item_kinds1.txt"), f => { var p = new FileParser(File.ReadAllBytes(f)); p.GetWord(); p.GetWord(); p.GetWord(); int c = (int)p.GetLong(); for (int i = 0; i < c; i++) { var it = ReadItem(p); it.Id = i; _repo.Items.Add(it); } });
        LoadIfExists(Path.Combine(path, "troops.txt"), f => { var p = new FileParser(File.ReadAllBytes(f)); p.GetWord(); p.GetWord(); p.GetWord(); int c = (int)p.GetLong(); for (int i = 0; i < c; i++) { var t = ReadTroop(p); t.Id = i; _repo.Troops.Add(t); } });
        LoadIfExists(Path.Combine(path, "factions.txt"), f => { var p = new FileParser(File.ReadAllBytes(f)); p.GetWord(); p.GetWord(); p.GetWord(); int c = (int)p.GetLong(); for (int i = 0; i < c; i++) { var fac = ReadFaction(p, c); fac.Id = i; _repo.Factions.Add(fac); } });
        LoadIfExists(Path.Combine(path, "party_templates.txt"), f => _repo.PartyTemplates.AddRange(EntityLoaders.LoadPartyTemplates(f)));
        LoadIfExists(Path.Combine(path, "parties.txt"), f => _repo.Parties.AddRange(EntityLoaders.LoadParties(f)));
        LoadIfExists(Path.Combine(path, "scenes.txt"), f => _repo.Scenes.AddRange(EntityLoaders.LoadScenes(f)));

        _repo.NotifyDataLoaded();
    }

    private static void LoadIfExists(string path, Action<string> load) { if (File.Exists(path)) load(path); }

    // ====== Tree ======

    private void BuildTree()
    {
        _tree.Nodes.Clear();
        AddTreeNode("Items", $"Items ({_repo.Items.Count})", 0);
        AddTreeNode("Troops", $"Troops ({_repo.Troops.Count})", 1);
        AddTreeNode("Factions", $"Factions ({_repo.Factions.Count})", 2);
        AddTreeNode("Parties", $"Parties ({_repo.Parties.Count})", 3);
        AddTreeNode("PartyTmpl", $"Party Templates ({_repo.PartyTemplates.Count})", 4);
        AddTreeNode("Scenes", $"Scenes ({_repo.Scenes.Count})", 5);
        _tree.ExpandAll();
    }

    private void AddTreeNode(string key, string text, int idx) => _tree.Nodes.Add(new TreeNode(text, idx, idx) { Tag = key });

    private void OnTreeNodeSelected(object? sender, TreeViewEventArgs e)
    {
        var tag = e.Node?.Tag?.ToString();
        if (tag == null) return;

        // Find or create tab
        foreach (TabPage tab in _tabs.TabPages)
            if ((string)tab.Tag == tag) { _tabs.SelectedTab = tab; return; }

        Control? editor = tag switch
        {
            "Items" => _itemEditor = new ItemListEditor(_repo),
            "Troops" => _troopEditor = new TroopListEditor(_repo),
            "Factions" => _factionEditor = new FactionListEditor(_repo),
            _ => new Label { Text = $"Editor for {tag} not yet implemented", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter }
        };

        var page = new TabPage(tag) { Tag = tag };
        page.Controls.Add(editor);
        editor.Dock = DockStyle.Fill;
        _tabs.TabPages.Add(page);
        _tabs.SelectedTab = page;
    }

    // ====== Actions ======

    private void OnSaveAll(object? sender, EventArgs e)
    {
        if (_repo.Items.Count == 0) { MessageBox.Show("No MOD loaded."); return; }
        try
        {
            FileSaver.SaveItems(_repo);
            FileSaver.SaveTroops(_repo);
            FileSaver.SaveFactions(_repo);
            _statusLabel.Text = $"Saved: {_repo.Items.Count} items, {_repo.Troops.Count} troops, {_repo.Factions.Count} factions";
            _repo.NotifyDataSaved();
            MessageBox.Show("MOD saved successfully.", "Save", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex) { MessageBox.Show($"Save failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
    }

    private void OnValidate(object? sender, EventArgs e) { var v = new DataValidator(); MessageBox.Show(v.Validate(_repo) ? "Validation PASSED" : $"{v.Errors.Count} issues found", "Validate"); }

    private void ShowStats()
    {
        var msg = $"MOD: {_repo.ModName}\n\n" +
                  $"Items: {_repo.Items.Count}\nTroops: {_repo.Troops.Count}\n" +
                  $"Factions: {_repo.Factions.Count}\nParties: {_repo.Parties.Count}\n" +
                  $"Party Templates: {_repo.PartyTemplates.Count}\nScenes: {_repo.Scenes.Count}";
        MessageBox.Show(msg, "MOD Statistics", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    // ====== Parsers (inline for speed) ======

    static Item ReadItem(FileParser p) { var it = new Item { DbName = p.GetWord(), DisplayName = p.GetWord(), TextureName = p.GetWord() }; it.MeshCount = p.GetLong(); for (int i = 0; i < it.MeshCount; i++) { it.MeshNames.Add(p.GetWord()); it.MeshParams.Add(p.GetWord()); } it.ItemType = p.GetWord(); it.Action = p.GetWord(); it.Price = p.GetLong(); it.Prefix = p.GetWord(); it.Weight = p.GetWord(); it.Abundance = p.GetLong(); it.HeadArmor = p.GetLong(); it.BodyArmor = p.GetLong(); it.LegArmor = p.GetLong(); it.Difficulty = p.GetLong(); it.HitPoints = p.GetLong(); it.SpeedRating = p.GetLong(); it.MissileSpeed = p.GetLong(); it.WeaponLength = p.GetLong(); it.MaxAmmo = p.GetLong(); it.ThrustDamage = p.GetLong(); it.SwingDamage = p.GetLong(); it.FactionCount = p.GetLong(); for (int i = 0; i < it.FactionCount; i++) it.Factions.Add(new FactionRef { Id = p.GetLong() }); it.TriggerCount = p.GetLong(); for (int i = 0; i < it.TriggerCount; i++) { p.GetDouble(); long ac = p.GetLong(); for (int j = 0; j < ac; j++) { p.GetWord(); p.SkipWords((int)p.GetLong()); } } return it; }
    static Troop ReadTroop(FileParser p) { var t = new Troop { StrId = p.GetWord(), Name = p.GetWord(), PluralName = p.GetWord(), UnknownWarband = p.GetWord(), Flags = p.GetWord(), Scene = p.GetLong(), Reserved = p.GetLong(), Faction = p.GetLong(), Upgrade1 = p.GetLong(), Upgrade2 = p.GetLong() }; for (int m = 0; m < 64; m++) t.Inventory[m] = new InventorySlot { ItemId = p.GetLong(), Quantity = p.GetLong() }; t.Attributes = new Attributes { Str = (int)p.GetLong(), Agi = (int)p.GetLong(), Int = (int)p.GetLong(), Cha = (int)p.GetLong(), Level = (int)p.GetLong() }; t.Proficiencies = new WeaponProficiencies { OneHanded = p.GetLong(), TwoHanded = p.GetLong(), Polearm = p.GetLong(), Archery = p.GetLong(), Crossbow = p.GetLong(), Throwing = p.GetLong(), Firearm = p.GetLong() }; for (int s = 0; s < 6; s++) t.Skills[s] = p.GetLong(); for (int f = 0; f < 8; f++) t.Face[f] = p.GetWord(); return t; }
    static Faction ReadFaction(FileParser p, int totalFactions) { var f = new Faction { StrId = p.GetWord(), Name = p.GetWord(), Flags = p.GetLong(), Color = p.GetWord() }; for (int i = 0; i < totalFactions; i++) f.Relationships.Add(new Relationship { FactionId = i, Value = p.GetDouble() }); p.GetLong(); return f; }
}
