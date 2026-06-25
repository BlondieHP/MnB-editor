using System.Text;
using MnBEditor.Core.Models;

namespace MnBEditor.Core;

/// <summary>
/// CSV localization file loader — O(n) via Dictionary lookup.
/// Ported from CsvLoader.bas.
/// </summary>
public static class CsvLoader
{
    public static void LoadPartyCsv(DataRepository repo, string filePath)
    {
        if (!File.Exists(filePath) || repo.Language == "en") return;
        var dict = repo.Parties.ToDictionary(p => p.StrId.ToLowerInvariant());
        foreach (var (key, value) in ParseCsvFile(filePath))
            if (dict.TryGetValue(key.ToLowerInvariant(), out var party))
                party.CsvName = value;
    }

    public static void LoadPartyTemplateCsv(DataRepository repo, string filePath)
    {
        if (!File.Exists(filePath) || repo.Language == "en") return;
        var dict = repo.PartyTemplates.ToDictionary(p => p.PtId.ToLowerInvariant());
        foreach (var (key, value) in ParseCsvFile(filePath))
            if (dict.TryGetValue(key.ToLowerInvariant(), out var pt))
                pt.CsvName = value;
    }

    public static void LoadItemCsv(DataRepository repo, string filePath)
    {
        if (!File.Exists(filePath) || repo.Language == "en") return;
        var dict = repo.Items.ToDictionary(i => i.DbName.ToLowerInvariant());
        foreach (var (key, value) in ParseCsvFile(filePath))
        {
            if (key.EndsWith("_pl"))
            {
                var baseKey = key[..^3].ToLowerInvariant();
                if (dict.TryGetValue(baseKey, out var item))
                    item.CsvNamePl = value;
            }
            else if (dict.TryGetValue(key.ToLowerInvariant(), out var item))
            {
                item.CsvName = value;
            }
        }
    }

    public static void LoadTroopCsv(DataRepository repo, string filePath)
    {
        if (!File.Exists(filePath) || repo.Language == "en") return;
        var dict = repo.Troops.ToDictionary(t => t.StrId.ToLowerInvariant());
        foreach (var (key, value) in ParseCsvFile(filePath))
        {
            if (key.EndsWith("_pl"))
            {
                var baseKey = key[..^3].ToLowerInvariant();
                if (dict.TryGetValue(baseKey, out var troop))
                    troop.CsvNamePl = value;
            }
            else if (dict.TryGetValue(key.ToLowerInvariant(), out var troop))
            {
                troop.CsvName = value;
            }
        }
    }

    public static void LoadFactionCsv(DataRepository repo, string filePath)
    {
        if (!File.Exists(filePath) || repo.Language == "en") return;
        var dict = repo.Factions.ToDictionary(f => f.StrId.ToLowerInvariant());
        foreach (var (key, value) in ParseCsvFile(filePath))
            if (dict.TryGetValue(key.ToLowerInvariant(), out var faction))
                faction.CsvName = value;
    }

    /// <summary>
    /// Parses a CSV file into key-value pairs.
    /// Format: key|value per line (pipe-delimited).
    /// </summary>
    private static IEnumerable<(string Key, string Value)> ParseCsvFile(string filePath)
    {
        var text = File.ReadAllText(filePath, Encoding.UTF8);
        foreach (var line in text.Split('\n'))
        {
            var trimmed = line.Trim();
            if (trimmed.Length == 0) continue;

            var parts = trimmed.Split('|');
            if (parts.Length >= 2)
                yield return (parts[0], parts[1]);
        }
    }
}
