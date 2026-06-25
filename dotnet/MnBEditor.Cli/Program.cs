using MnBEditor.Core;
using MnBEditor.Core.Models;

if (args.Length == 0 || args[0] == "--help" || args[0] == "-h")
{
    PrintHelp();
    return 0;
}

var command = args[0].ToLowerInvariant();
var modPath = args.Length > 1 ? args[1] : ".";

try
{
    return command switch
    {
        "stats" => ShowStats(modPath),
        "validate" => ValidateMod(modPath),
        "items" => ListItems(modPath, args.Length > 2 ? args[2] : null),
        "troops" => ListTroops(modPath, args.Length > 2 ? args[2] : null),
        "export" => ExportCsv(modPath, args.Length > 2 ? args[2] : "items"),
        _ => UnknownCommand(command)
    };
}
catch (Exception ex)
{
    Console.Error.WriteLine($"Error: {ex.Message}");
    return 1;
}

static int UnknownCommand(string cmd)
{
    Console.Error.WriteLine($"Unknown command: {cmd}");
    PrintHelp();
    return 1;
}

static void PrintHelp()
{
    Console.WriteLine("MnBEditor CLI - M&B Warband MOD Inspector");
    Console.WriteLine();
    Console.WriteLine("Usage: MnBEditor.Cli <command> [mod-path] [options]");
    Console.WriteLine();
    Console.WriteLine("Commands:");
    Console.WriteLine("  stats     Show MOD statistics (item/troop/faction counts)");
    Console.WriteLine("  validate  Validate data integrity");
    Console.WriteLine("  items     List items, optionally filtered by keyword");
    Console.WriteLine("  troops    List troops, optionally filtered by keyword");
    Console.WriteLine("  export    Export items/troops to CSV");
    Console.WriteLine();
    Console.WriteLine("Examples:");
    Console.WriteLine("  MnBEditor.Cli stats ../Support");
    Console.WriteLine("  MnBEditor.Cli items ../Support horse");
    Console.WriteLine("  MnBEditor.Cli export ../Support items > items.csv");
}

static int ShowStats(string modPath)
{
    var repo = LoadMod(modPath);
    if (repo == null) return 1;

    Console.WriteLine($"MOD: {repo.ModName}");
    Console.WriteLine($"Path: {repo.ModPath}");
    Console.WriteLine();
    Console.WriteLine("=== Entity Counts ===");
    Console.WriteLine($"{"Items:",-20} {repo.Items.Count,6}");
    Console.WriteLine($"{"Troops:",-20} {repo.Troops.Count,6}");
    Console.WriteLine($"{"Factions:",-20} {repo.Factions.Count,6}");
    Console.WriteLine($"{"Parties:",-20} {repo.Parties.Count,6}");
    Console.WriteLine($"{"Party Templates:",-20} {repo.PartyTemplates.Count,6}");
    Console.WriteLine($"{"Scenes:",-20} {repo.Scenes.Count,6}");

    if (repo.Items.Count > 0)
    {
        Console.WriteLine();
        Console.WriteLine("=== Item Type Distribution ===");
        var typeCounts = new Dictionary<string, int>();
        foreach (var item in repo.Items)
        {
            var type = GetItemTypeName(item.ItemType);
            typeCounts.TryGetValue(type, out var c);
            typeCounts[type] = c + 1;
        }
        foreach (var kv in typeCounts.OrderByDescending(kv => kv.Value))
            Console.WriteLine($"  {kv.Key,-20} {kv.Value,6}");
    }

    if (repo.Troops.Count > 0)
    {
        Console.WriteLine();
        Console.WriteLine("=== Troop Stats ===");
        var avgLevel = repo.Troops.Average(t => (double)t.Attributes.Level);
        var maxLevel = repo.Troops.Max(t => t.Attributes.Level);
        Console.WriteLine($"  Average Level: {avgLevel:F1}");
        Console.WriteLine($"  Max Level:     {maxLevel}");
    }

    return 0;
}

static int ValidateMod(string modPath)
{
    var repo = LoadMod(modPath);
    if (repo == null) return 1;

    var validator = new DataValidator();
    var valid = validator.Validate(repo);

    if (valid)
    {
        Console.WriteLine("Validation PASSED - no issues found.");
        return 0;
    }
    else
    {
        Console.Error.WriteLine($"Validation FAILED - {validator.Errors.Count} issue(s):");
        foreach (var err in validator.Errors)
            Console.Error.WriteLine($"  - {err}");
        return 1;
    }
}

static int ListItems(string modPath, string? filter)
{
    var repo = LoadMod(modPath);
    if (repo == null) return 1;

    foreach (var item in repo.Items)
    {
        if (filter != null &&
            !item.DbName.Contains(filter, StringComparison.OrdinalIgnoreCase) &&
            !item.DisplayName.Contains(filter, StringComparison.OrdinalIgnoreCase))
            continue;

        Console.WriteLine($"[{item.Id,4}] {item.DbName,-40} {item.DisplayName,-30} " +
                         $"Price:{item.Price,6} Type:{GetItemTypeName(item.ItemType)}");
    }

    return 0;
}

static int ListTroops(string modPath, string? filter)
{
    var repo = LoadMod(modPath);
    if (repo == null) return 1;

    foreach (var troop in repo.Troops)
    {
        if (filter != null &&
            !troop.StrId.Contains(filter, StringComparison.OrdinalIgnoreCase) &&
            !troop.Name.Contains(filter, StringComparison.OrdinalIgnoreCase))
            continue;

        Console.WriteLine($"[{troop.Id,4}] {troop.StrId,-35} {troop.Name,-25} " +
                         $"Lv:{troop.Attributes.Level,3} Str:{troop.Attributes.Str,3}");
    }

    return 0;
}

static int ExportCsv(string modPath, string entityType)
{
    var repo = LoadMod(modPath);
    if (repo == null) return 1;

    switch (entityType.ToLowerInvariant())
    {
        case "items":
            Console.WriteLine("Id,DbName,DisplayName,Price,Weight,Type");
            foreach (var item in repo.Items)
                Console.WriteLine($"{item.Id},{item.DbName},{item.DisplayName}," +
                                 $"{item.Price},{item.Weight},{GetItemTypeName(item.ItemType)}");
            break;
        case "troops":
            Console.WriteLine("Id,StrId,Name,Level,Str,Agi,Int,Cha");
            foreach (var t in repo.Troops)
                Console.WriteLine($"{t.Id},{t.StrId},{t.Name},{t.Attributes.Level}," +
                                 $"{t.Attributes.Str},{t.Attributes.Agi},{t.Attributes.Int},{t.Attributes.Cha}");
            break;
        default:
            Console.Error.WriteLine($"Unknown entity type: {entityType}. Use 'items' or 'troops'.");
            return 1;
    }

    return 0;
}

static DataRepository? LoadMod(string modPath)
{
    if (!Directory.Exists(modPath))
    {
        Console.Error.WriteLine($"Mod path not found: {modPath}");
        return null;
    }

    var repo = new DataRepository
    {
        ModPath = Path.GetFullPath(modPath),
        ModName = Path.GetFileName(Path.GetFullPath(modPath))
    };

    // Load items
    var itemFile = Path.Combine(modPath, "item_kinds1.txt");
    if (File.Exists(itemFile))
    {
        var parser = new FileParser(File.ReadAllBytes(itemFile));
        parser.GetWord(); parser.GetWord(); parser.GetWord();
        int count = (int)parser.GetLong();
        for (int i = 0; i < count; i++)
        {
            var item = ReadItem(parser);
            item.Id = i;
            repo.Items.Add(item);
        }
    }

    // Load troops
    var troopFile = Path.Combine(modPath, "troops.txt");
    if (File.Exists(troopFile))
    {
        var parser = new FileParser(File.ReadAllBytes(troopFile));
        parser.GetWord(); parser.GetWord(); parser.GetWord();
        int count = (int)parser.GetLong();
        for (int i = 0; i < count; i++)
        {
            var troop = ReadTroop(parser);
            troop.Id = i;
            repo.Troops.Add(troop);
        }
    }

    // Load factions
    var factionFile = Path.Combine(modPath, "factions.txt");
    if (File.Exists(factionFile))
    {
        var parser = new FileParser(File.ReadAllBytes(factionFile));
        parser.GetWord(); parser.GetWord(); parser.GetWord();
        int count = (int)parser.GetLong();
        for (int i = 0; i < count; i++)
        {
            var f = ReadFaction(parser);
            f.Id = i;
            repo.Factions.Add(f);
        }
    }

    return repo;
}

static Item ReadItem(FileParser p)
{
    var item = new Item { DbName = p.GetWord(), DisplayName = p.GetWord(), TextureName = p.GetWord() };
    item.MeshCount = p.GetLong();
    for (int i = 0; i < item.MeshCount; i++) { item.MeshNames.Add(p.GetWord()); item.MeshParams.Add(p.GetWord()); }
    item.ItemType = p.GetWord(); item.Action = p.GetWord();
    item.Price = p.GetLong(); item.Prefix = p.GetWord(); item.Weight = p.GetWord();
    item.Abundance = p.GetLong(); item.HeadArmor = p.GetLong(); item.BodyArmor = p.GetLong(); item.LegArmor = p.GetLong();
    item.Difficulty = p.GetLong(); item.HitPoints = p.GetLong(); item.SpeedRating = p.GetLong();
    item.MissileSpeed = p.GetLong(); item.WeaponLength = p.GetLong(); item.MaxAmmo = p.GetLong();
    item.ThrustDamage = p.GetLong(); item.SwingDamage = p.GetLong();
    item.FactionCount = p.GetLong();
    for (int i = 0; i < item.FactionCount; i++) item.Factions.Add(new FactionRef { Id = p.GetLong() });
    item.TriggerCount = p.GetLong();
    for (int i = 0; i < item.TriggerCount; i++) { p.GetDouble(); long ac = p.GetLong(); for (int j = 0; j < ac; j++) { p.GetWord(); p.SkipWords((int)p.GetLong()); } }
    return item;
}

static Troop ReadTroop(FileParser p)
{
    var t = new Troop { StrId = p.GetWord(), Name = p.GetWord(), PluralName = p.GetWord(), UnknownWarband = p.GetWord(), Flags = p.GetWord(), Scene = p.GetLong(), Reserved = p.GetLong(), Faction = p.GetLong(), Upgrade1 = p.GetLong(), Upgrade2 = p.GetLong() };
    for (int m = 0; m < 64; m++) t.Inventory[m] = new InventorySlot { ItemId = p.GetLong(), Quantity = p.GetLong() };
    t.Attributes = new Attributes { Str = (int)p.GetLong(), Agi = (int)p.GetLong(), Int = (int)p.GetLong(), Cha = (int)p.GetLong(), Level = (int)p.GetLong() };
    t.Proficiencies = new WeaponProficiencies { OneHanded = p.GetLong(), TwoHanded = p.GetLong(), Polearm = p.GetLong(), Archery = p.GetLong(), Crossbow = p.GetLong(), Throwing = p.GetLong(), Firearm = p.GetLong() };
    for (int s = 0; s < 6; s++) t.Skills[s] = p.GetLong();
    for (int f = 0; f < 8; f++) t.Face[f] = p.GetWord();
    return t;
}

static Faction ReadFaction(FileParser p)
{
    var f = new Faction { StrId = p.GetWord(), Name = p.GetWord(), Flags = p.GetLong(), Color = p.GetWord() };
    // Read relationships: first word is count, then N double values
    var relCount = p.GetLong();
    for (int i = 0; i < relCount; i++) p.GetDouble();
    p.GetLong(); // reserved
    return f;
}

static string GetItemTypeName(string typeFlags)
{
    if (string.IsNullOrEmpty(typeFlags) || !long.TryParse(typeFlags, out var flags)) return "unknown";
    var t = flags & 0x1F;
    return t switch
    {
        0x1 => "horse", 0x2 => "one_handed", 0x3 => "two_handed", 0x4 => "polearm",
        0x5 => "arrows", 0x6 => "bolts", 0x7 => "shield", 0x8 => "bow",
        0x9 => "crossbow", 0xA => "thrown", 0xB => "goods", 0xC => "head_armor",
        0xD => "body_armor", 0xE => "foot_armor", 0xF => "hand_armor",
        0x10 => "pistol", 0x11 => "musket", 0x12 => "bullets",
        0x13 => "animal", 0x14 => "book", _ => $"type_{t:X}"
    };
}
