namespace MnBEditor.Core.Models;

/// <summary>
/// Map icon definition.
/// </summary>
public record MapIcon
{
    public long Id { get; set; }
    public string StrId { get; set; } = "";
    public long Flags { get; set; }
    public string MeshName { get; set; } = "";
    public string Scale { get; set; } = "";
    public long Sound { get; set; }
    public string SoundName { get; set; } = "";
    public string[] Offset { get; set; } = new string[3];
    public long TriggerCount { get; set; }
    public List<Trigger> Triggers { get; set; } = new();
    public bool IsEdited { get; set; }
}

/// <summary>
/// Sound definition with resources.
/// </summary>
public record Sound
{
    public long Id { get; set; }
    public string Name { get; set; } = "";
    public string Flags { get; set; } = "";
    public long ResourceCount { get; set; }
    public List<SoundResource> Resources { get; set; } = new();
    public bool IsEdited { get; set; }
}

public record SoundResource
{
    public long Id { get; set; }
    public string StrId { get; set; }
    public long Unknown { get; set; }
}

/// <summary>
/// Particle system definition.
/// </summary>
public record ParticleSystem
{
    public long Id { get; set; }
    public string StrId { get; set; } = "";
    public string Flags { get; set; } = "";
    public string MeshName { get; set; } = "";
    public long ParticleCount { get; set; }
    public double Life { get; set; }
    public double Damping { get; set; }
    public double Gravity { get; set; }
    public double TurbulenceSize { get; set; }
    public double TurbulenceStrength { get; set; }
    public Point2D AlphaKey { get; set; }
    public Point2D RedKey { get; set; }
    public Point2D GreenKey { get; set; }
    public Point2D BlueKey { get; set; }
    public Point2D ScaleKey { get; set; }
    public double[] EmitBoxSize { get; set; } = new double[2];
    public double[] EmitVelocity { get; set; } = new double[2];
    public double EmitDirRandomness { get; set; }
    public double ParticleRotationSpeed { get; set; }
    public double ParticleRotationDamping { get; set; }
    public bool IsEdited { get; set; }
}

/// <summary>
/// Tableau material definition.
/// </summary>
public record TableauMaterial
{
    public long Id { get; set; }
    public string StrId { get; set; } = "";
    public string Flags { get; set; } = "";
    public string Sample { get; set; } = "";
    public long Width { get; set; }
    public long Height { get; set; }
    public Point2D Min { get; set; }
    public Point2D Max { get; set; }
    public long OpCount { get; set; }
    public List<Operation> OpBlock { get; set; } = new();
    public bool IsEdited { get; set; }
}

/// <summary>
/// Mesh definition with translation, rotation, scale.
/// </summary>
public record Mesh
{
    public long Id { get; set; }
    public string StrId { get; set; } = "";
    public string Flags { get; set; } = "";
    public string ResourceName { get; set; } = "";
    public Point3D Translation { get; set; }
    public Point3D Rotation { get; set; }
    public Point3D Scale { get; set; }
    public bool IsEdited { get; set; }
}

/// <summary>
/// Time-based trigger (simple trigger).
/// </summary>
public record TimeTrigger
{
    public long Id { get; set; }
    public double CheckInterval { get; set; }
    public double DelayInterval { get; set; }
    public double RearmInterval { get; set; }
    public long ConditionsCount { get; set; }
    public List<Operation> Conditions { get; set; } = new();
    public long ConsequencesCount { get; set; }
    public List<Operation> Consequences { get; set; } = new();
    public bool IsEdited { get; set; }
}

/// <summary>
/// Game string (localized text).
/// </summary>
public record GameString
{
    public long Id { get; set; }
    public string Name { get; set; } = "";
    public string Text { get; set; } = "";
    public string Csv { get; set; } = "";
    public bool IsEdited { get; set; }
}

/// <summary>
/// Global variable.
/// </summary>
public record GlobalVariable
{
    public long Id { get; set; }
    public string VarName { get; set; } = "";
}

/// <summary>
/// Item modifier (prefix) definition.
/// </summary>
public record ItemModifier
{
    public string Id { get; set; } = "";
    public string CsvName { get; set; } = "";
}
