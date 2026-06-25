namespace MnBEditor.Core.Models;

/// <summary>
/// Scene definition.
/// Ported from Type_Scene in Models.bas.
/// </summary>
public record Scene
{
    public long Id { get; set; }
    public string StrId { get; set; } = "";
    public string Name { get; set; } = "";
    public string Flags { get; set; } = "";
    public string MeshName { get; set; } = "";
    public string BodyName { get; set; } = "";
    public Point2D Point1 { get; set; }
    public Point2D Point2 { get; set; }
    public double WaterLevel { get; set; }
    public string TerrainCode { get; set; } = "";
    public long AccessCount { get; set; }
    public List<string> Accesses { get; set; } = new();
    public long ChestCount { get; set; }
    public List<FactionRef> Chests { get; set; } = new();
    public string OuterTerrainType { get; set; } = "";
    public bool IsEdited { get; set; }
}
