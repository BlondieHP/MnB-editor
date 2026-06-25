using MnBEditor.Core.Models;

namespace MnBEditor.Core;

/// <summary>
/// Item data file loader/saver — proof of concept for the migration.
/// Ported from LoadItemFile + ReadItemLine in FileLoader.bas/FileSaver.bas.
/// </summary>
public static class ItemLoader
{
    /// <summary>
    /// Loads item_kinds1.txt into the DataRepository.
    /// </summary>
    public static void LoadItems(DataRepository repo, string filePath)
    {
        var parser = new FileParser(filePath);

        // Read version info (3 words)
        var versionInfo = new string[3];
        for (int i = 0; i < 3; i++)
            versionInfo[i] = parser.GetWord();

        // Read item count
        int count = (int)parser.GetLong();
        repo.Items.Clear();
        repo.Items.Capacity = count;

        for (int i = 0; i < count; i++)
        {
            var item = ReadItem(parser);
            item.Id = i;
            item.IsEdited = i >= (repo.IsFirstTimeEdit ? repo.EditInfo[Constants.EditInfoItemsCount] : 0);
            repo.Items.Add(item);
        }
    }

    /// <summary>
    /// Reads a single item record from the parser.
    /// Ported from ReadItemLine in ModMain.bas.
    /// </summary>
    private static Item ReadItem(FileParser parser)
    {
        var item = new Item();

        item.DbName = parser.GetWord();
        item.DisplayName = parser.GetWord();
        item.TextureName = parser.GetWord();

        // Mesh count and mesh names
        item.MeshCount = parser.GetLong();
        for (int i = 0; i < item.MeshCount; i++)
        {
            item.MeshNames.Add(parser.GetWord());
            item.MeshParams.Add(parser.GetWord());
        }

        item.ItemType = parser.GetWord();
        item.Action = parser.GetWord();
        item.Price = parser.GetLong();
        item.Prefix = parser.GetWord();
        item.Weight = parser.GetWord();
        item.Abundance = parser.GetLong();
        item.HeadArmor = parser.GetLong();
        item.BodyArmor = parser.GetLong();
        item.LegArmor = parser.GetLong();
        item.Difficulty = parser.GetLong();
        item.HitPoints = parser.GetLong();
        item.SpeedRating = parser.GetLong();
        item.MissileSpeed = parser.GetLong();
        item.WeaponLength = parser.GetLong();
        item.MaxAmmo = parser.GetLong();
        item.ThrustDamage = parser.GetLong();
        item.SwingDamage = parser.GetLong();

        // Faction affiliations
        item.FactionCount = parser.GetLong();
        for (int i = 0; i < item.FactionCount; i++)
        {
            item.Factions.Add(new FactionRef
            {
                Id = parser.GetLong(),
                StrId = ""
            });
        }

        // Triggers
        item.TriggerCount = parser.GetLong();
        for (int i = 0; i < item.TriggerCount; i++)
        {
            var trigger = new Trigger
            {
                TriggerType = parser.GetDouble()
            };
            trigger.ActionCount = parser.GetLong();
            for (int j = 0; j < trigger.ActionCount; j++)
            {
                var op = new Operation
                {
                    OpCode = parser.GetWord()
                };
                op.ParamCount = parser.GetLong();
                for (int k = 0; k < op.ParamCount; k++)
                {
                    op.Parameters.Add(new OperationParam
                    {
                        Value = parser.GetWord(),
                        Type = ""
                    });
                }
                trigger.Actions.Add(op);
            }
            item.Triggers.Add(trigger);
        }

        return item;
    }

    /// <summary>
    /// Saves items to item_kinds1.txt format.
    /// Ported from SaveItemFile + OutputItemLine in ModMain.bas.
    /// </summary>
    public static void SaveItems(DataRepository repo, string filePath)
    {
        using var writer = new StreamWriter(filePath, false, System.Text.Encoding.GetEncoding(1252));

        // Write version header (3 words: "itemsfile version 3" or similar)
        writer.Write("itemsfile version 3 ");
        writer.Write(repo.Items.Count);
        writer.Write("\r\n");

        foreach (var item in repo.Items)
        {
            writer.Write(OutputItemLine(item));
            writer.Write("\r\n");
        }
    }

    /// <summary>
    /// Formats a single item as a space-delimited TXT line.
    /// </summary>
    private static string OutputItemLine(Item item)
    {
        var parts = new List<string>
        {
            $" {item.DbName}",
            $" {item.DisplayName.Replace(' ', '_')}",
            $" {item.TextureName}"
        };

        // Mesh count and meshes
        parts.Add($" {item.MeshCount}");
        for (int i = 0; i < item.MeshCount; i++)
        {
            parts.Add($" {item.MeshNames[i]}");
            parts.Add($" {item.MeshParams[i]}");
        }

        parts.Add($" {item.ItemType}");
        parts.Add($" {item.Action}");
        parts.Add($" {item.Price}");
        parts.Add($" {item.Prefix}");
        parts.Add($" {item.Weight}");
        parts.Add($" {item.Abundance}");
        parts.Add($" {item.HeadArmor}");
        parts.Add($" {item.BodyArmor}");
        parts.Add($" {item.LegArmor}");
        parts.Add($" {item.Difficulty}");
        parts.Add($" {item.HitPoints}");
        parts.Add($" {item.SpeedRating}");
        parts.Add($" {item.MissileSpeed}");
        parts.Add($" {item.WeaponLength}");
        parts.Add($" {item.MaxAmmo}");
        parts.Add($" {item.ThrustDamage}");
        parts.Add($" {item.SwingDamage}");

        // Faction affiliations
        parts.Add($" {item.FactionCount}");
        foreach (var f in item.Factions)
            parts.Add($" {f.Id}");

        // Triggers
        parts.Add($" {item.TriggerCount}");
        foreach (var t in item.Triggers)
        {
            parts.Add($" {t.TriggerType}");
            parts.Add($" {t.ActionCount}");
            foreach (var op in t.Actions)
            {
                parts.Add($" {op.OpCode} {op.ParamCount}");
                foreach (var p in op.Parameters)
                    parts.Add($" {p.Value}");
            }
        }

        return string.Concat(parts);
    }
}
