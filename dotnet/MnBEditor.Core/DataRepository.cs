using MnBEditor.Core.Models;

namespace MnBEditor.Core;

/// <summary>
/// Centralized data store with change notification.
/// Ported from DataRepository.cls. In .NET, this will use
/// ObservableCollection&lt;T&gt; for UI binding in WinForms/WPF.
/// </summary>
public class DataRepository
{
    // --- Core data arrays ---
    public List<Troop> Troops { get; private set; } = new();
    public List<Item> Items { get; private set; } = new();
    public List<Faction> Factions { get; private set; } = new();
    public List<Party> Parties { get; private set; } = new();
    public List<PartyTemplate> PartyTemplates { get; private set; } = new();
    public List<Scene> Scenes { get; private set; } = new();
    public List<MapIcon> MapIcons { get; private set; } = new();
    public List<Sound> Sounds { get; private set; } = new();
    public List<ParticleSystem> ParticleSystems { get; private set; } = new();
    public List<TableauMaterial> TableauMaterials { get; private set; } = new();
    public List<Mesh> Meshes { get; private set; } = new();
    public List<TimeTrigger> TimeTriggers { get; private set; } = new();
    public List<GameString> Strings { get; private set; } = new();
    public List<GameString> QuickStrings { get; private set; } = new();
    public List<GlobalVariable> GlobalVariables { get; private set; } = new();
    public List<ItemModifier> ItemModifiers { get; private set; } = new();

    // --- Mod info ---
    public string ModPath { get; set; } = "";
    public string ModName { get; set; } = "";
    public string Language { get; set; } = "en";
    public bool IsFirstTimeEdit { get; set; }
    public long[] EditInfo { get; set; } = new long[14];

    // --- Change notification events ---
    public event Action<int>? TroopChanged;
    public event Action<int>? ItemChanged;
    public event Action<int>? FactionChanged;
    public event Action<int>? PartyChanged;
    public event Action<int>? SceneChanged;
    public event Action? DataLoaded;
    public event Action? DataSaved;

    public void ClearAll()
    {
        Troops.Clear();
        Items.Clear();
        Factions.Clear();
        Parties.Clear();
        PartyTemplates.Clear();
        Scenes.Clear();
        MapIcons.Clear();
        Sounds.Clear();
        ParticleSystems.Clear();
        TableauMaterials.Clear();
        Meshes.Clear();
        TimeTriggers.Clear();
        Strings.Clear();
        QuickStrings.Clear();
        GlobalVariables.Clear();
        ItemModifiers.Clear();
    }

    public void NotifyDataLoaded() => DataLoaded?.Invoke();
    public void NotifyDataSaved() => DataSaved?.Invoke();

    public void NotifyTroopChanged(int index) => TroopChanged?.Invoke(index);
    public void NotifyItemChanged(int index) => ItemChanged?.Invoke(index);
    public void NotifyFactionChanged(int index) => FactionChanged?.Invoke(index);
    public void NotifyPartyChanged(int index) => PartyChanged?.Invoke(index);
    public void NotifySceneChanged(int index) => SceneChanged?.Invoke(index);
}
