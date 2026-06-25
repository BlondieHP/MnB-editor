namespace MnBEditor.Core.Models;

/// <summary>
/// Party template definition.
/// Ported from Type_PT in Models.bas.
/// </summary>
public record PartyTemplate
{
    public long Id { get; set; }
    public string PtId { get; set; } = "";
    public string Name { get; set; } = "";
    public string Flags { get; set; } = "";
    public string Menu { get; set; } = "";
    public long Faction { get; set; }
    public string FactionStrId { get; set; } = "";
    public long Personality { get; set; }

    /// <summary>Up to 6 troop stacks in this template.</summary>
    public PartyStack[] Stacks { get; set; } = new PartyStack[6];

    public bool IsEdited { get; set; }
    public string CsvName { get; set; } = "";
}

/// <summary>
/// Party (map entity) definition.
/// Ported from Type_Party in Models.bas.
/// </summary>
public record Party
{
    public long UnknownTitle { get; set; }
    public long Id { get; set; }
    public long Id2 { get; set; }
    public string StrId { get; set; } = "";
    public string Name { get; set; } = "";
    public string Flags { get; set; } = "";
    public string MapIconStrId { get; set; } = "";
    public string Menu { get; set; } = "";
    public long Template { get; set; }
    public string TemplateStrId { get; set; } = "";
    public long Faction { get; set; }
    public string FactionStrId { get; set; } = "";
    public long[] Personality { get; set; } = new long[2];
    public string AiBehavior { get; set; } = "";
    public string AiTarget { get; set; } = "";
    public string AiTargetStrId { get; set; } = "";
    public long Reserved { get; set; }

    /// <summary>Initial positions [3].</summary>
    public Point2D[] InitPositions { get; set; } = new Point2D[3];

    public string UnknownStr { get; set; } = "";
    public long StacksCount { get; set; }
    public List<PartyStack> Stacks { get; set; } = new();
    public string Degree { get; set; } = "";

    public bool IsEdited { get; set; }
    public string CsvName { get; set; } = "";
}
