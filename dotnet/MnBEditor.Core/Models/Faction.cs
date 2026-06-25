namespace MnBEditor.Core.Models;

/// <summary>
/// Faction (kingdom/culture) definition.
/// Ported from Type_Faction in Models.bas.
/// </summary>
public record Faction
{
    public long Id { get; set; }
    public string StrId { get; set; } = "";
    public string Name { get; set; } = "";
    public long Flags { get; set; }
    public string Color { get; set; } = "";         // Hex color string

    /// <summary>Relationships with all other factions.</summary>
    public List<Relationship> Relationships { get; set; } = new();

    public long Reserved { get; set; }
    public bool IsEdited { get; set; }
    public string CsvName { get; set; } = "";
}
