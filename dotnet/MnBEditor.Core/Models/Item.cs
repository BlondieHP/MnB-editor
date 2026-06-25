namespace MnBEditor.Core.Models;

/// <summary>
/// Item (weapon/armor/goods) definition.
/// Ported from Type_Item in Models.bas.
/// </summary>
public record Item
{
    public long Id { get; set; }
    public string DbName { get; set; } = "";      // Internal database name
    public string DisplayName { get; set; } = "";  // In-game display name
    public string TextureName { get; set; } = "";
    public long MeshCount { get; set; }
    public List<string> MeshNames { get; set; } = new();
    public List<string> MeshParams { get; set; } = new();
    public string ItemType { get; set; } = "";     // Hex flags string
    public string Action { get; set; } = "";       // Capability flags string
    public long Price { get; set; }
    public string Prefix { get; set; } = "";
    public string Weight { get; set; } = "";
    public long Abundance { get; set; }
    public long HeadArmor { get; set; }
    public long BodyArmor { get; set; }
    public long LegArmor { get; set; }
    public long Difficulty { get; set; }
    public long HitPoints { get; set; }
    public long SpeedRating { get; set; }
    public long MissileSpeed { get; set; }
    public long WeaponLength { get; set; }
    public long MaxAmmo { get; set; }
    public long ThrustDamage { get; set; }
    public long SwingDamage { get; set; }

    /// <summary>Factions this item belongs to.</summary>
    public long FactionCount { get; set; }
    public List<FactionRef> Factions { get; set; } = new();

    /// <summary>Item triggers (used for firearms/scripted items).</summary>
    public long TriggerCount { get; set; }
    public List<Trigger> Triggers { get; set; } = new();

    public bool IsEdited { get; set; }
    public string CsvName { get; set; } = "";
    public string CsvNamePl { get; set; } = "";
}

public record struct FactionRef
{
    public long Id { get; set; }
    public string StrId { get; set; }
}
