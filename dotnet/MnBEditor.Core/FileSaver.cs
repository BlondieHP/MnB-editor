using System.Text;
using MnBEditor.Core.Models;

namespace MnBEditor.Core;

/// <summary>
/// Saves entity data back to M&B TXT format.
/// Ported from FileSaver.bas.
/// </summary>
public static class FileSaver
{
    public static void SaveItems(DataRepository repo, string? path = null)
    {
        path ??= Path.Combine(repo.ModPath, "item_kinds1.txt");
        using var w = new StreamWriter(path, false, Encoding.ASCII);
        w.Write($"itemsfile version 3\r\n{repo.Items.Count}\r\n");
        foreach (var it in repo.Items) w.Write(FormatItem(it));
    }

    public static void SaveTroops(DataRepository repo, string? path = null)
    {
        path ??= Path.Combine(repo.ModPath, "troops.txt");
        using var w = new StreamWriter(path, false, Encoding.ASCII);
        w.Write($"troopsfile version 3\r\n{repo.Troops.Count} \r\n");
        foreach (var t in repo.Troops)
        {
            var sid = t.StrId.StartsWith("trp_") ? t.StrId : "trp_" + t.StrId;
            w.Write($" {sid} {t.Name} {t.PluralName} {t.UnknownWarband} {t.Flags}" +
                   $" {t.Scene} {t.Reserved} {t.Faction} {t.Upgrade1} {t.Upgrade2}\r\n");
            w.Write("  ");
            for (int m = 0; m < 64; m++) w.Write($"{t.Inventory[m].ItemId} {t.Inventory[m].Quantity} ");
            w.Write($"\r\n  {t.Attributes.Str} {t.Attributes.Agi} {t.Attributes.Int} {t.Attributes.Cha} {t.Attributes.Level}\r\n");
            w.Write($" {t.Proficiencies.OneHanded} {t.Proficiencies.TwoHanded} {t.Proficiencies.Polearm}" +
                   $" {t.Proficiencies.Archery} {t.Proficiencies.Crossbow} {t.Proficiencies.Throwing} {t.Proficiencies.Firearm}\r\n");
            w.Write($"{t.Skills[0]} {t.Skills[1]} {t.Skills[2]} {t.Skills[3]} {t.Skills[4]} {t.Skills[5]} \r\n");
            w.Write($"  {t.Face[0]} {t.Face[1]} {t.Face[2]} {t.Face[3]} {t.Face[4]} {t.Face[5]} {t.Face[6]} {t.Face[7]} \r\n\r\n");
        }
    }

    public static void SaveFactions(DataRepository repo, string? path = null)
    {
        path ??= Path.Combine(repo.ModPath, "factions.txt");
        using var w = new StreamWriter(path, false, Encoding.ASCII);
        w.Write($"factionsfile version 1\r\n{repo.Factions.Count}\r\n");
        foreach (var f in repo.Factions)
        {
            w.Write($" {f.StrId} {f.Name} {f.Flags} {f.Color}\r\n");
            foreach (var r in f.Relationships) w.Write($"{r.Value} ");
            w.Write($"\r\n{f.Reserved}\r\n");
        }
    }

    public static string FormatItemPublic(Item it) => FormatItem(it);
    private static string FormatItem(Item it)
    {
        var sb = new StringBuilder();
        sb.Append($" {it.DbName} {it.DisplayName} {it.TextureName} {it.MeshCount}");
        for (int i = 0; i < it.MeshCount; i++) sb.Append($" {it.MeshNames[i]} {it.MeshParams[i]}");
        sb.Append($" {it.ItemType} {it.Action} {it.Price} {it.Prefix} {it.Weight} {it.Abundance}" +
                 $" {it.HeadArmor} {it.BodyArmor} {it.LegArmor} {it.Difficulty} {it.HitPoints}" +
                 $" {it.SpeedRating} {it.MissileSpeed} {it.WeaponLength} {it.MaxAmmo}" +
                 $" {it.ThrustDamage} {it.SwingDamage} {it.FactionCount}");
        for (int i = 0; i < it.FactionCount; i++) sb.Append($" {it.Factions[i].Id}");
        sb.Append($" {it.TriggerCount}");
        for (int i = 0; i < it.TriggerCount; i++)
        {
            var t = it.Triggers[i];
            sb.Append($" {t.TriggerType} {t.ActionCount}");
            for (int j = 0; j < t.ActionCount; j++)
            {
                var op = t.Actions[j];
                sb.Append($" {op.OpCode} {op.ParamCount}");
                for (int k = 0; k < op.ParamCount; k++) sb.Append($" {op.Parameters[k].Value}");
            }
        }
        sb.Append("\r\n");
        return sb.ToString();
    }
}
