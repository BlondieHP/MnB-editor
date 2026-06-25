namespace MnBEditor.Core.Models;

/// <summary>
/// Troop (soldier/hero/NPC) definition.
/// Ported from Type_Troops in Models.bas.
/// </summary>
public record Troop
{
    public long Id { get; set; }
    public string StrId { get; set; } = "";          // e.g., "trp_player"
    public string Name { get; set; } = "";
    public string PluralName { get; set; } = "";
    public string UnknownWarband { get; set; } = "";  // Warband-specific field
    public string Flags { get; set; } = "";
    public long Scene { get; set; }
    public long SceneId { get; set; }
    public string SceneStrId { get; set; } = "";
    public long Entry { get; set; }
    public long Reserved { get; set; }
    public long Faction { get; set; }
    public string FactionStrId { get; set; } = "";
    public long Upgrade1 { get; set; }
    public string Upgrade1StrId { get; set; } = "";
    public long Upgrade2 { get; set; }
    public string Upgrade2StrId { get; set; } = "";

    /// <summary>64 inventory slots (item_id, quantity).</summary>
    public InventorySlot[] Inventory { get; set; } = new InventorySlot[64];

    public Attributes Attributes { get; set; }
    public WeaponProficiencies Proficiencies { get; set; }

    /// <summary>6 skill values (encoded).</summary>
    public long[] Skills { get; set; } = new long[6];

    /// <summary>8 face code values.</summary>
    public string[] Face { get; set; } = new string[8];

    public bool IsEdited { get; set; }
    public string CsvName { get; set; } = "";
    public string CsvNamePl { get; set; } = "";
}
