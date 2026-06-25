namespace MnBEditor.Core.Models;

/// <summary>
/// Simple XY coordinate pair.
/// </summary>
public record struct Point2D(double X, double Y);

/// <summary>
/// 3D coordinate.
/// </summary>
public record struct Point3D(double X, double Y, double Z);

/// <summary>
/// Inventory slot: item ID + quantity.
/// </summary>
public record InventorySlot
{
    public long ItemId { get; set; }
    public string ItemIdStr { get; set; } = "";
    public long Quantity { get; set; }
}

/// <summary>
/// Troop attributes (STR/AGI/INT/CHA + level).
/// </summary>
public record struct Attributes
{
    public int Str { get; set; }
    public int Agi { get; set; }
    public int Int { get; set; }
    public int Cha { get; set; }
    public int Level { get; set; }
}

/// <summary>
/// Weapon proficiency ratings.
/// </summary>
public record struct WeaponProficiencies
{
    public long OneHanded { get; set; }
    public long TwoHanded { get; set; }
    public long Polearm { get; set; }
    public long Archery { get; set; }
    public long Crossbow { get; set; }
    public long Throwing { get; set; }
    public long Firearm { get; set; }
}

/// <summary>
/// Stack entry in a party or party template.
/// </summary>
public record PartyStack
{
    public long TroopId { get; set; }
    public string TroopIdStr { get; set; } = "";
    public long Min { get; set; }
    public long Max { get; set; }
    public long Flags { get; set; } // 0=member, 1=prisoner
}

/// <summary>
/// Faction relationship entry.
/// </summary>
public record Relationship
{
    public long FactionId { get; set; }
    public string FactionIdStr { get; set; } = "";
    public double Value { get; set; }
}

/// <summary>
/// A single operation within a trigger or operation block.
/// </summary>
public record Operation
{
    public string OpCode { get; set; } = "";
    public long ParamCount { get; set; }
    public List<OperationParam> Parameters { get; set; } = new();
}

/// <summary>
/// Parameter within an operation.
/// </summary>
public record OperationParam
{
    public string Value { get; set; } = "";
    public string Type { get; set; } = "";
}

/// <summary>
/// A trigger with conditions and consequences.
/// </summary>
public record Trigger
{
    public double TriggerType { get; set; }
    public long ActionCount { get; set; }
    public List<Operation> Actions { get; set; } = new();
}
