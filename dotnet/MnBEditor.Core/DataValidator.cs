using MnBEditor.Core.Models;

namespace MnBEditor.Core;

/// <summary>
/// Data integrity validation. Run before saving to catch data corruption early.
/// </summary>
public class DataValidator
{
    public List<string> Errors { get; } = new();

    public bool Validate(DataRepository repo)
    {
        Errors.Clear();
        ValidateTroops(repo);
        ValidateItems(repo);
        ValidateFactions(repo);
        ValidateParties(repo);
        ValidateScenes(repo);
        return Errors.Count == 0;
    }

    private void ValidateTroops(DataRepository repo)
    {
        for (int i = 0; i < repo.Troops.Count; i++)
        {
            var t = repo.Troops[i];
            if (string.IsNullOrWhiteSpace(t.StrId))
                Errors.Add($"Troop[{i}]: empty StrId");
            if (string.IsNullOrWhiteSpace(t.Name))
                Errors.Add($"Troop[{i}] '{t.StrId}': empty Name");
            if (t.Faction < 0 || t.Faction >= repo.Factions.Count)
                Errors.Add($"Troop[{i}] '{t.StrId}': Faction={t.Faction} out of range (0-{repo.Factions.Count - 1})");
            if (t.Attributes.Level < 0 || t.Attributes.Level > 100)
                Errors.Add($"Troop[{i}] '{t.StrId}': Level={t.Attributes.Level} out of range (0-100)");
            if (t.Upgrade1 >= repo.Troops.Count)
                Errors.Add($"Troop[{i}] '{t.StrId}': Upgrade1={t.Upgrade1} references non-existent troop");
            if (t.Upgrade2 >= repo.Troops.Count)
                Errors.Add($"Troop[{i}] '{t.StrId}': Upgrade2={t.Upgrade2} references non-existent troop");
        }
    }

    private void ValidateItems(DataRepository repo)
    {
        for (int i = 0; i < repo.Items.Count; i++)
        {
            var it = repo.Items[i];
            if (string.IsNullOrWhiteSpace(it.DbName))
                Errors.Add($"Item[{i}]: empty DbName");
            if (it.Price < 0)
                Errors.Add($"Item[{i}] '{it.DbName}': Price={it.Price} is negative");
            if (it.Abundance < 0 || it.Abundance > 255)
                Errors.Add($"Item[{i}] '{it.DbName}': Abundance={it.Abundance} out of range (0-255)");
            if (it.HeadArmor < 0 || it.BodyArmor < 0 || it.LegArmor < 0)
                Errors.Add($"Item[{i}] '{it.DbName}': armor value is negative");
            if (it.HitPoints < 0)
                Errors.Add($"Item[{i}] '{it.DbName}': HitPoints is negative");
            // Validate faction references
            foreach (var f in it.Factions)
            {
                if (f.Id >= repo.Factions.Count)
                    Errors.Add($"Item[{i}] '{it.DbName}': references non-existent faction {f.Id}");
            }
        }
    }

    private void ValidateFactions(DataRepository repo)
    {
        for (int i = 0; i < repo.Factions.Count; i++)
        {
            var f = repo.Factions[i];
            if (string.IsNullOrWhiteSpace(f.StrId))
                Errors.Add($"Faction[{i}]: empty StrId");
            if (string.IsNullOrWhiteSpace(f.Name))
                Errors.Add($"Faction[{i}] '{f.StrId}': empty Name");
            if (f.Relationships.Count != repo.Factions.Count)
                Errors.Add($"Faction[{i}] '{f.StrId}': {f.Relationships.Count} relationships, expected {repo.Factions.Count}");
        }
    }

    private void ValidateParties(DataRepository repo)
    {
        for (int i = 0; i < repo.Parties.Count; i++)
        {
            var p = repo.Parties[i];
            if (string.IsNullOrWhiteSpace(p.StrId))
                Errors.Add($"Party[{i}]: empty StrId");
            if (p.Faction >= repo.Factions.Count && p.Faction != 0)
                Errors.Add($"Party[{i}] '{p.StrId}': Faction={p.Faction} out of range");
            if (p.Template >= repo.PartyTemplates.Count && p.Template != 0)
                Errors.Add($"Party[{i}] '{p.StrId}': Template={p.Template} out of range");
            // Validate stacks reference valid troops
            for (int j = 0; j < p.Stacks.Count; j++)
            {
                if (p.Stacks[j].TroopId >= repo.Troops.Count)
                    Errors.Add($"Party[{i}] '{p.StrId}': Stack[{j}] references non-existent troop {p.Stacks[j].TroopId}");
            }
        }
    }

    private void ValidateScenes(DataRepository repo)
    {
        for (int i = 0; i < repo.Scenes.Count; i++)
        {
            var s = repo.Scenes[i];
            if (string.IsNullOrWhiteSpace(s.StrId))
                Errors.Add($"Scene[{i}]: empty StrId");
            if (string.IsNullOrWhiteSpace(s.Name))
                Errors.Add($"Scene[{i}] '{s.StrId}': empty Name");
        }
    }
}
