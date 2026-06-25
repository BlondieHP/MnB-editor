using System.Text.Json;
using MnBEditor.Core;
using MnBEditor.Core.Models;

if (args.Length == 0 || args[0] == "--help" || args[0] == "-h") { PrintHelp(); return 0; }

var command = args[0].ToLowerInvariant();
var modPath = args.Length > 1 && !args[1].StartsWith('-') ? args[1] : ".";

try
{
    return command switch
    {
        "stats" => ShowStats(modPath),
        "validate" => ValidateMod(modPath),
        "items" => ListItems(modPath, Arg(args, 2)),
        "troops" => ListTroops(modPath, Arg(args, 2)),
        "info" => ShowInfo(modPath, Arg(args, 2), Arg(args, 3)),
        "search" => Search(modPath, Arg(args, 2), args.Skip(3).ToArray()),
        "export" => Export(modPath, Arg(args, 2), args.Contains("--json")),
        "compare" => Compare(Arg(args, 1), Arg(args, 2)),
        _ => UnknownCommand(command)
    };
}
catch (Exception ex)
{
    Console.Error.WriteLine($"Error: {ex.Message}");
    return 1;
}

static string? Arg(string[] a, int i) => i < a.Length && !a[i].StartsWith('-') ? a[i] : null;
static string[] Flags(string[] a, int start) => a.Skip(start).Where(x => x.StartsWith('-')).ToArray();
static string? OptVal(string[] a, string flag) { var i = Array.IndexOf(a, flag); return i >= 0 && i + 1 < a.Length ? a[i + 1] : null; }

static int UnknownCommand(string c) { Console.Error.WriteLine($"Unknown: {c}"); PrintHelp(); return 1; }

static void PrintHelp()
{
    Console.WriteLine("MnBEditor CLI - M&B Warband MOD Inspector");
    Console.WriteLine();
    Console.WriteLine("Usage: MnBEditor.Cli <command> [mod-path] [options]");
    Console.WriteLine();
    Console.WriteLine("Inspect:");
    Console.WriteLine("  stats <path>              Show entity counts + type distribution");
    Console.WriteLine("  validate <path>           Check data integrity");
    Console.WriteLine("  items <path> [keyword]    List items filtered by name");
    Console.WriteLine("  troops <path> [keyword]   List troops filtered by name");
    Console.WriteLine("  info <path> items|troops <id|name>  Show full entity details");
    Console.WriteLine();
    Console.WriteLine("Search:");
    Console.WriteLine("  search <path> troops [filters]");
    Console.WriteLine("    Filters: level>=N  level<=N  faction=N  str>=N  name=keyword");
    Console.WriteLine("    Example: search . troops level>=30 str>=20 name=knight");
    Console.WriteLine();
    Console.WriteLine("Export:");
    Console.WriteLine("  export <path> items|troops [--json]   CSV or JSON export");
    Console.WriteLine();
    Console.WriteLine("Compare:");
    Console.WriteLine("  compare <path1> <path2>  Diff two MODs, show added/removed/changed");
}

// ====== STATS ======

static int ShowStats(string modPath)
{
    var repo = LoadMod(modPath); if (repo == null) return 1;
    Console.WriteLine($"MOD: {repo.ModName}");
    Console.WriteLine($"Path: {repo.ModPath}\n");
    Console.WriteLine("=== Entity Counts ===");
    Console.WriteLine($"{"Items:",-20} {repo.Items.Count,6}");
    Console.WriteLine($"{"Troops:",-20} {repo.Troops.Count,6}");
    Console.WriteLine($"{"Factions:",-20} {repo.Factions.Count,6}");
    Console.WriteLine($"{"Parties:",-20} {repo.Parties.Count,6}");
    Console.WriteLine($"{"Party Templates:",-20} {repo.PartyTemplates.Count,6}");
    Console.WriteLine($"{"Scenes:",-20} {repo.Scenes.Count,6}");

    if (repo.Items.Count > 0)
    {
        Console.WriteLine("\n=== Item Types ===");
        var tc = new Dictionary<string, int>();
        foreach (var it in repo.Items) { var t = GetItemTypeName(it.ItemType); tc.TryGetValue(t, out var c); tc[t] = c + 1; }
        foreach (var kv in tc.OrderByDescending(kv => kv.Value)) Console.WriteLine($"  {kv.Key,-20} {kv.Value,6}");
    }

    if (repo.Troops.Count > 0)
    {
        Console.WriteLine("\n=== Troop Stats ===");
        Console.WriteLine($"  Avg Level: {repo.Troops.Average(t => (double)t.Attributes.Level):F1}");
        Console.WriteLine($"  Max Level: {repo.Troops.Max(t => t.Attributes.Level)}");
        var topLv = repo.Troops.OrderByDescending(t => t.Attributes.Level).Take(5);
        Console.WriteLine("  Top 5 by level:");
        foreach (var t in topLv) Console.WriteLine($"    [{t.Id}] {t.StrId,-35} Lv{t.Attributes.Level}");
    }

    if (repo.Factions.Count > 0)
    {
        Console.WriteLine("\n=== Factions ===");
        foreach (var f in repo.Factions.Take(20)) Console.WriteLine($"  [{f.Id,3}] {f.StrId,-30} {f.Name}");
        if (repo.Factions.Count > 20) Console.WriteLine($"  ... and {repo.Factions.Count - 20} more");
    }
    return 0;
}

// ====== VALIDATE ======

static int ValidateMod(string modPath)
{
    var repo = LoadMod(modPath); if (repo == null) return 1;
    var v = new DataValidator();
    if (v.Validate(repo)) { Console.WriteLine("PASSED - no issues."); return 0; }
    Console.Error.WriteLine($"FAILED - {v.Errors.Count} issue(s):");
    foreach (var e in v.Errors.Take(50)) Console.Error.WriteLine($"  - {e}");
    if (v.Errors.Count > 50) Console.Error.WriteLine($"  ... and {v.Errors.Count - 50} more");
    return 1;
}

// ====== LIST ======

static int ListItems(string modPath, string? filter)
{
    var repo = LoadMod(modPath); if (repo == null) return 1;
    int shown = 0;
    foreach (var it in repo.Items)
    {
        if (filter != null && !it.DbName.Contains(filter, StringComparison.OrdinalIgnoreCase) && !it.DisplayName.Contains(filter, StringComparison.OrdinalIgnoreCase)) continue;
        Console.WriteLine($"[{it.Id,4}] {it.DbName,-40} {it.DisplayName,-30} Price:{it.Price,6} {GetItemTypeName(it.ItemType)}");
        shown++;
    }
    Console.Error.WriteLine($"Shown: {shown}/{repo.Items.Count} items");
    return 0;
}

static int ListTroops(string modPath, string? filter)
{
    var repo = LoadMod(modPath); if (repo == null) return 1;
    int shown = 0;
    foreach (var t in repo.Troops)
    {
        if (filter != null && !t.StrId.Contains(filter, StringComparison.OrdinalIgnoreCase) && !t.Name.Contains(filter, StringComparison.OrdinalIgnoreCase)) continue;
        Console.WriteLine($"[{t.Id,4}] {t.StrId,-35} {t.Name,-25} Lv:{t.Attributes.Level,3} Str:{t.Attributes.Str,3} Agi:{t.Attributes.Agi,3}");
        shown++;
    }
    Console.Error.WriteLine($"Shown: {shown}/{repo.Troops.Count} troops");
    return 0;
}

// ====== INFO (NEW) ======

static int ShowInfo(string modPath, string? entityType, string? query)
{
    if (entityType == null || query == null) { Console.Error.WriteLine("Usage: info <path> items|troops <id|name>"); return 1; }
    var repo = LoadMod(modPath); if (repo == null) return 1;

    switch (entityType.ToLowerInvariant())
    {
        case "items":
        case "item":
            Item? item = null;
            if (int.TryParse(query, out var iid) && iid < repo.Items.Count) item = repo.Items[iid];
            else item = repo.Items.FirstOrDefault(x => x.DbName.Contains(query, StringComparison.OrdinalIgnoreCase));
            if (item == null) { Console.Error.WriteLine($"Item not found: {query}"); return 1; }
            PrintItemDetail(item);
            break;
        case "troops":
        case "troop":
            Troop? troop = null;
            if (int.TryParse(query, out var tid) && tid < repo.Troops.Count) troop = repo.Troops[tid];
            else troop = repo.Troops.FirstOrDefault(x => x.StrId.Contains(query, StringComparison.OrdinalIgnoreCase) || x.Name.Contains(query, StringComparison.OrdinalIgnoreCase));
            if (troop == null) { Console.Error.WriteLine($"Troop not found: {query}"); return 1; }
            PrintTroopDetail(troop, repo);
            break;
        default:
            Console.Error.WriteLine("Use 'items' or 'troops'"); return 1;
    }
    return 0;
}

static void PrintItemDetail(Item it)
{
    Console.WriteLine($"=== Item [{it.Id}] ===");
    Console.WriteLine($"  DB Name:      {it.DbName}");
    Console.WriteLine($"  Display Name: {it.DisplayName}");
    Console.WriteLine($"  Type:         {GetItemTypeName(it.ItemType)} (flags: {it.ItemType})");
    Console.WriteLine($"  Price:        {it.Price}");
    Console.WriteLine($"  Weight:       {it.Weight}");
    Console.WriteLine($"  Abundance:    {it.Abundance}");
    Console.WriteLine($"  Head Armor:   {it.HeadArmor}");
    Console.WriteLine($"  Body Armor:   {it.BodyArmor}");
    Console.WriteLine($"  Leg Armor:    {it.LegArmor}");
    Console.WriteLine($"  Difficulty:   {it.Difficulty}");
    Console.WriteLine($"  Hit Points:   {it.HitPoints}");
    Console.WriteLine($"  Speed:        {it.SpeedRating}");
    Console.WriteLine($"  Weapon Len:   {it.WeaponLength}");
    Console.WriteLine($"  Max Ammo:     {it.MaxAmmo}");
    Console.WriteLine($"  Thrust Dmg:   {it.ThrustDamage}");
    Console.WriteLine($"  Swing Dmg:    {it.SwingDamage}");
    Console.WriteLine($"  Meshes:       {it.MeshCount}");
    for (int i = 0; i < it.MeshCount; i++) Console.WriteLine($"    {it.MeshNames[i]} ({it.MeshParams[i]})");
    Console.WriteLine($"  Factions:     {it.FactionCount}");
    foreach (var f in it.Factions) Console.WriteLine($"    Fac[{f.Id}]");
    Console.WriteLine($"  Triggers:     {it.TriggerCount}");
    Console.WriteLine($"  Texture:      {it.TextureName}");
}

static void PrintTroopDetail(Troop t, DataRepository repo)
{
    Console.WriteLine($"=== Troop [{t.Id}] ===");
    Console.WriteLine($"  StrID:        {t.StrId}");
    Console.WriteLine($"  Name:         {t.Name} ({t.PluralName})");
    Console.WriteLine($"  Level:        {t.Attributes.Level}");
    Console.WriteLine($"  Attributes:   STR:{t.Attributes.Str} AGI:{t.Attributes.Agi} INT:{t.Attributes.Int} CHA:{t.Attributes.Cha}");
    Console.WriteLine($"  Proficiencies:");
    Console.WriteLine($"    1H:{t.Proficiencies.OneHanded} 2H:{t.Proficiencies.TwoHanded} Pole:{t.Proficiencies.Polearm}");
    Console.WriteLine($"    Arch:{t.Proficiencies.Archery} Xbow:{t.Proficiencies.Crossbow} Thr:{t.Proficiencies.Throwing} Fire:{t.Proficiencies.Firearm}");
    Console.WriteLine($"  Faction:      {(t.Faction < repo.Factions.Count ? repo.Factions[(int)t.Faction].Name : $"#{t.Faction}")}");
    Console.WriteLine($"  Upgrade1:     {(t.Upgrade1 < repo.Troops.Count ? repo.Troops[(int)t.Upgrade1].Name : "none")} (#{t.Upgrade1})");
    Console.WriteLine($"  Upgrade2:     {(t.Upgrade2 < repo.Troops.Count ? repo.Troops[(int)t.Upgrade2].Name : "none")} (#{t.Upgrade2})");
    Console.WriteLine($"  Inventory:    (first 8 non-empty slots)");
    int shown = 0;
    for (int i = 0; i < 64 && shown < 8; i++)
    {
        if (t.Inventory[i].ItemId > 0 && t.Inventory[i].ItemId < repo.Items.Count)
        {
            Console.WriteLine($"    [{i,2}] {repo.Items[(int)t.Inventory[i].ItemId].DbName,-35} x{t.Inventory[i].Quantity}");
            shown++;
        }
    }
    Console.WriteLine($"  Flags:        {t.Flags}");
    Console.WriteLine($"  Scene:        {t.Scene}");
}

// ====== SEARCH (NEW) ======

static int Search(string modPath, string? entityType, string[] filters)
{
    if (entityType != "troops") { Console.Error.WriteLine("Currently only 'search <path> troops' is supported"); return 1; }
    var repo = LoadMod(modPath); if (repo == null) return 1;

    // Parse filters: level>=30, level<=10, faction=5, str>=20, name=knight
    int? lvMin = null, lvMax = null, strMin = null, agiMin = null, faction = null;
    string? nameFilter = null;
    foreach (var raw in filters)
    {
        // Extract operator and values: handle >=, <=, >, <, =
        string op; string key; string val;
        if (raw.Contains(">=")) { var p = raw.Split(">="); key = p[0]; val = p[1]; op = ">="; }
        else if (raw.Contains("<=")) { var p = raw.Split("<="); key = p[0]; val = p[1]; op = "<="; }
        else if (raw.Contains(">")) { var p = raw.Split('>'); key = p[0]; val = p[1]; op = ">"; }
        else if (raw.Contains("<")) { var p = raw.Split('<'); key = p[0]; val = p[1]; op = "<"; }
        else if (raw.Contains('=')) { var p = raw.Split('='); key = p[0]; val = p[1]; op = "="; }
        else continue;

        key = key.ToLowerInvariant();

        if (key == "level" && op == ">=") lvMin = int.Parse(val);
        else if (key == "level" && op == "<=") lvMax = int.Parse(val);
        else if (key == "level" && op == "=") { lvMin = lvMax = int.Parse(val); }
        else if (key == "str" && op == ">=") strMin = int.Parse(val);
        else if (key == "agi" && op == ">=") agiMin = int.Parse(val);
        else if (key == "faction") faction = int.Parse(val);
        else if (key == "name") nameFilter = val;
    }

    var results = repo.Troops.AsEnumerable();
    if (lvMin.HasValue) { var m = lvMin.Value; results = results.Where(t => t.Attributes.Level >= m); }
    if (lvMax.HasValue) { var m = lvMax.Value; results = results.Where(t => t.Attributes.Level <= m); }
    if (strMin.HasValue) { var m = strMin.Value; results = results.Where(t => t.Attributes.Str >= m); }
    if (agiMin.HasValue) { var m = agiMin.Value; results = results.Where(t => t.Attributes.Agi >= m); }
    if (faction.HasValue) { var m = faction.Value; results = results.Where(t => t.Faction == m); }
    if (nameFilter != null) { var n = nameFilter; results = results.Where(t => t.Name.Contains(n, StringComparison.OrdinalIgnoreCase) || t.StrId.Contains(n, StringComparison.OrdinalIgnoreCase)); }

    var list = results.ToList();
    Console.WriteLine($"Search results: {list.Count}/{repo.Troops.Count} troops");
    Console.WriteLine();
    foreach (var t in list)
    {
        var facName = t.Faction < repo.Factions.Count ? repo.Factions[(int)t.Faction].Name : $"#{t.Faction}";
        Console.WriteLine($"[{t.Id,4}] {t.StrId,-35} {t.Name,-25} Lv:{t.Attributes.Level,3} STR:{t.Attributes.Str,3} AGI:{t.Attributes.Agi,3} {facName}");
    }
    return 0;
}

// ====== EXPORT (ENHANCED) ======

static int Export(string modPath, string? entityType, bool json)
{
    if (entityType == null) { Console.Error.WriteLine("Usage: export <path> items|troops [--json]"); return 1; }
    var repo = LoadMod(modPath); if (repo == null) return 1;

    if (json)
    {
        var opts = new JsonSerializerOptions { WriteIndented = true };
        switch (entityType.ToLowerInvariant())
        {
            case "items": Console.WriteLine(JsonSerializer.Serialize(repo.Items.Select(it => new { it.Id, it.DbName, it.DisplayName, Type = GetItemTypeName(it.ItemType), it.Price, it.Weight, it.Abundance, it.HeadArmor, it.BodyArmor, it.LegArmor, it.Difficulty, it.HitPoints, it.SpeedRating, it.WeaponLength, it.MaxAmmo, it.ThrustDamage, it.SwingDamage, it.MeshCount }), opts)); break;
            case "troops": Console.WriteLine(JsonSerializer.Serialize(repo.Troops.Select(t => new { t.Id, t.StrId, t.Name, t.Attributes.Level, t.Attributes.Str, t.Attributes.Agi, t.Attributes.Int, t.Attributes.Cha, t.Proficiencies.OneHanded, t.Proficiencies.TwoHanded, t.Proficiencies.Polearm, t.Proficiencies.Archery, t.Proficiencies.Crossbow, t.Proficiencies.Throwing, t.Proficiencies.Firearm, Faction = t.Faction < repo.Factions.Count ? repo.Factions[(int)t.Faction].Name : $"#{t.Faction}" }), opts)); break;
            default: Console.Error.WriteLine("Use 'items' or 'troops'"); return 1;
        }
    }
    else
    {
        switch (entityType.ToLowerInvariant())
        {
            case "items": Console.WriteLine("Id,DbName,DisplayName,Price,Weight,Type"); foreach (var it in repo.Items) Console.WriteLine($"{it.Id},{it.DbName},{it.DisplayName},{it.Price},{it.Weight},{GetItemTypeName(it.ItemType)}"); break;
            case "troops": Console.WriteLine("Id,StrId,Name,Level,Str,Agi,Int,Cha"); foreach (var t in repo.Troops) Console.WriteLine($"{t.Id},{t.StrId},{t.Name},{t.Attributes.Level},{t.Attributes.Str},{t.Attributes.Agi},{t.Attributes.Int},{t.Attributes.Cha}"); break;
            default: Console.Error.WriteLine("Use 'items' or 'troops'"); return 1;
        }
    }
    return 0;
}

// ====== COMPARE (NEW) ======

static int Compare(string? path1, string? path2)
{
    if (path1 == null || path2 == null) { Console.Error.WriteLine("Usage: compare <path1> <path2>"); return 1; }
    Console.Error.WriteLine($"Loading MOD A: {path1}...");
    var repo1 = LoadMod(path1); if (repo1 == null) return 1;
    Console.Error.WriteLine($"Loading MOD B: {path2}...");
    var repo2 = LoadMod(path2); if (repo2 == null) return 1;

    Console.WriteLine($"=== MOD Comparison ===");
    Console.WriteLine($"A: {repo1.ModName} ({repo1.Items.Count} items, {repo1.Troops.Count} troops)");
    Console.WriteLine($"B: {repo2.ModName} ({repo2.Items.Count} items, {repo2.Troops.Count} troops)");
    Console.WriteLine();

    // Compare items by DbName
    var items1 = repo1.Items.GroupBy(i => i.DbName).ToDictionary(g => g.Key, g => g.First());
    var items2 = repo2.Items.GroupBy(i => i.DbName).ToDictionary(g => g.Key, g => g.First());
    var addedItems = items2.Keys.Except(items1.Keys).ToList();
    var removedItems = items1.Keys.Except(items2.Keys).ToList();
    var commonItems = items1.Keys.Intersect(items2.Keys).ToList();

    // Find changed items (same DbName, different stats)
    var changedItems = commonItems.Where(k =>
    {
        var a = items1[k]; var b = items2[k];
        return a.Price != b.Price || a.Weight != b.Weight || a.ItemType != b.ItemType ||
               a.HeadArmor != b.HeadArmor || a.BodyArmor != b.BodyArmor || a.HitPoints != b.HitPoints;
    }).ToList();

    Console.WriteLine("--- Items ---");
    Console.WriteLine($"  Added:   {addedItems.Count}");
    foreach (var k in addedItems.Take(10)) Console.WriteLine($"    + {k}");
    if (addedItems.Count > 10) Console.WriteLine($"    ... {addedItems.Count - 10} more");

    Console.WriteLine($"  Removed: {removedItems.Count}");
    foreach (var k in removedItems.Take(10)) Console.WriteLine($"    - {k}");
    if (removedItems.Count > 10) Console.WriteLine($"    ... {removedItems.Count - 10} more");

    Console.WriteLine($"  Changed: {changedItems.Count}");
    foreach (var k in changedItems.Take(10))
    {
        var a = items1[k]; var b = items2[k];
        Console.WriteLine($"    ~ {k}: Price {a.Price}→{b.Price} Wt {a.Weight}→{b.Weight}");
    }
    if (changedItems.Count > 10) Console.WriteLine($"    ... {changedItems.Count - 10} more");
    Console.WriteLine($"  Unchanged: {commonItems.Count - changedItems.Count}");

    // Compare troops
    Console.WriteLine();
    var troops1 = repo1.Troops.GroupBy(t => t.StrId).ToDictionary(g => g.Key, g => g.First());
    var troops2 = repo2.Troops.GroupBy(t => t.StrId).ToDictionary(g => g.Key, g => g.First());
    var addedTroops = troops2.Keys.Except(troops1.Keys).ToList();
    var removedTroops = troops1.Keys.Except(troops2.Keys).ToList();
    var changedTroops = troops1.Keys.Intersect(troops2.Keys).Where(k =>
    {
        var a = troops1[k]; var b = troops2[k];
        return a.Attributes.Level != b.Attributes.Level || a.Attributes.Str != b.Attributes.Str;
    }).ToList();

    Console.WriteLine("--- Troops ---");
    Console.WriteLine($"  Added:   {addedTroops.Count}");
    foreach (var k in addedTroops.Take(5)) Console.WriteLine($"    + {k}");
    Console.WriteLine($"  Removed: {removedTroops.Count}");
    foreach (var k in removedTroops.Take(5)) Console.WriteLine($"    - {k}");
    Console.WriteLine($"  Changed: {changedTroops.Count}");
    foreach (var k in changedTroops.Take(5))
    {
        var a = troops1[k]; var b = troops2[k];
        Console.WriteLine($"    ~ {k}: Lv {a.Attributes.Level}→{b.Attributes.Level} STR {a.Attributes.Str}→{b.Attributes.Str}");
    }
    Console.WriteLine($"  Unchanged: {troops1.Keys.Intersect(troops2.Keys).Count() - changedTroops.Count}");

    return 0;
}

// ====== LOADER ======

static DataRepository? LoadMod(string modPath)
{
    if (!Directory.Exists(modPath)) { Console.Error.WriteLine($"Not found: {modPath}"); return null; }
    var repo = new DataRepository { ModPath = Path.GetFullPath(modPath), ModName = Path.GetFileName(Path.GetFullPath(modPath)) };

    var itemFile = Path.Combine(modPath, "item_kinds1.txt");
    if (File.Exists(itemFile)) { var p = new FileParser(File.ReadAllBytes(itemFile)); p.GetWord(); p.GetWord(); p.GetWord(); int c = (int)p.GetLong(); for (int i = 0; i < c; i++) { var it = ReadItem(p); it.Id = i; repo.Items.Add(it); } }

    var troopFile = Path.Combine(modPath, "troops.txt");
    if (File.Exists(troopFile)) { var p = new FileParser(File.ReadAllBytes(troopFile)); p.GetWord(); p.GetWord(); p.GetWord(); int c = (int)p.GetLong(); for (int i = 0; i < c; i++) { var t = ReadTroop(p); t.Id = i; repo.Troops.Add(t); } }

    var factionFile = Path.Combine(modPath, "factions.txt");
    if (File.Exists(factionFile)) { var p = new FileParser(File.ReadAllBytes(factionFile)); p.GetWord(); p.GetWord(); p.GetWord(); int c = (int)p.GetLong(); for (int i = 0; i < c; i++) { var f = ReadFaction(p, c); f.Id = i; repo.Factions.Add(f); } }

    var ptFile = Path.Combine(modPath, "party_templates.txt");
    if (File.Exists(ptFile)) { repo.PartyTemplates.AddRange(EntityLoaders.LoadPartyTemplates(ptFile)); }

    var partyFile = Path.Combine(modPath, "parties.txt");
    if (File.Exists(partyFile)) { repo.Parties.AddRange(EntityLoaders.LoadParties(partyFile)); }

    var sceneFile = Path.Combine(modPath, "scenes.txt");
    if (File.Exists(sceneFile)) { repo.Scenes.AddRange(EntityLoaders.LoadScenes(sceneFile)); }

    return repo;
}

static Item ReadItem(FileParser p) { var it = new Item { DbName = p.GetWord(), DisplayName = p.GetWord(), TextureName = p.GetWord() }; it.MeshCount = p.GetLong(); for (int i = 0; i < it.MeshCount; i++) { it.MeshNames.Add(p.GetWord()); it.MeshParams.Add(p.GetWord()); } it.ItemType = p.GetWord(); it.Action = p.GetWord(); it.Price = p.GetLong(); it.Prefix = p.GetWord(); it.Weight = p.GetWord(); it.Abundance = p.GetLong(); it.HeadArmor = p.GetLong(); it.BodyArmor = p.GetLong(); it.LegArmor = p.GetLong(); it.Difficulty = p.GetLong(); it.HitPoints = p.GetLong(); it.SpeedRating = p.GetLong(); it.MissileSpeed = p.GetLong(); it.WeaponLength = p.GetLong(); it.MaxAmmo = p.GetLong(); it.ThrustDamage = p.GetLong(); it.SwingDamage = p.GetLong(); it.FactionCount = p.GetLong(); for (int i = 0; i < it.FactionCount; i++) it.Factions.Add(new FactionRef { Id = p.GetLong() }); it.TriggerCount = p.GetLong(); for (int i = 0; i < it.TriggerCount; i++) { p.GetDouble(); long ac = p.GetLong(); for (int j = 0; j < ac; j++) { p.GetWord(); p.SkipWords((int)p.GetLong()); } } return it; }

static Troop ReadTroop(FileParser p) { var t = new Troop { StrId = p.GetWord(), Name = p.GetWord(), PluralName = p.GetWord(), UnknownWarband = p.GetWord(), Flags = p.GetWord(), Scene = p.GetLong(), Reserved = p.GetLong(), Faction = p.GetLong(), Upgrade1 = p.GetLong(), Upgrade2 = p.GetLong() }; for (int m = 0; m < 64; m++) t.Inventory[m] = new InventorySlot { ItemId = p.GetLong(), Quantity = p.GetLong() }; t.Attributes = new Attributes { Str = (int)p.GetLong(), Agi = (int)p.GetLong(), Int = (int)p.GetLong(), Cha = (int)p.GetLong(), Level = (int)p.GetLong() }; t.Proficiencies = new WeaponProficiencies { OneHanded = p.GetLong(), TwoHanded = p.GetLong(), Polearm = p.GetLong(), Archery = p.GetLong(), Crossbow = p.GetLong(), Throwing = p.GetLong(), Firearm = p.GetLong() }; for (int s = 0; s < 6; s++) t.Skills[s] = p.GetLong(); for (int f = 0; f < 8; f++) t.Face[f] = p.GetWord(); return t; }

static Faction ReadFaction(FileParser p, int totalFactions) { var f = new Faction { StrId = p.GetWord(), Name = p.GetWord(), Flags = p.GetLong(), Color = p.GetWord() }; for (int i = 0; i < totalFactions; i++) f.Relationships.Add(new Relationship { FactionId = i, Value = p.GetDouble() }); p.GetLong(); return f; }

static string GetItemTypeName(string tf) { if (string.IsNullOrEmpty(tf) || !long.TryParse(tf, out var f)) return "?"; var t = f & 0x1F; return t switch { 0x1 => "horse", 0x2 => "1h_wpn", 0x3 => "2h_wpn", 0x4 => "polearm", 0x5 => "arrows", 0x6 => "bolts", 0x7 => "shield", 0x8 => "bow", 0x9 => "crossbow", 0xA => "thrown", 0xB => "goods", 0xC => "head_arm", 0xD => "body_arm", 0xE => "foot_arm", 0xF => "hand_arm", 0x10 => "pistol", 0x11 => "musket", 0x12 => "bullets", 0x13 => "animal", 0x14 => "book", _ => $"0x{t:X}" }; }
