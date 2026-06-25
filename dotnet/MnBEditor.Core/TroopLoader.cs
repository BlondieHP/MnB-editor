using MnBEditor.Core.Models;

namespace MnBEditor.Core;

/// <summary>
/// Troop data file loader/saver.
/// Ported from LoadTroopFile + SaveTroopFile in FileLoader.bas/FileSaver.bas.
/// </summary>
public static class TroopLoader
{
    public static void LoadTroops(DataRepository repo, string filePath)
    {
        var parser = new FileParser(filePath);

        // Read version info (3 words)
        var versionInfo = new string[3];
        for (int i = 0; i < 3; i++)
            versionInfo[i] = parser.GetWord();

        // Read troop count
        int count = (int)parser.GetLong();
        repo.Troops.Clear();
        repo.Troops.Capacity = count;

        for (int i = 0; i < count; i++)
        {
            var troop = ReadTroop(parser);
            troop.Id = i;
            troop.IsEdited = i >= (repo.IsFirstTimeEdit ? repo.EditInfo[Constants.EditInfoTroopsCount] : 0);
            repo.Troops.Add(troop);
        }
    }

    private static Troop ReadTroop(FileParser parser)
    {
        var t = new Troop
        {
            StrId = parser.GetWord(),
            Name = parser.GetWord(),
            PluralName = parser.GetWord(),
            CsvName = "",  // Set later by CSV loader
            CsvNamePl = "",
            UnknownWarband = parser.GetWord(),
            Flags = parser.GetWord(),
            Scene = parser.GetLong(),
            Reserved = parser.GetLong(),
            Faction = parser.GetLong(),
            Upgrade1 = parser.GetLong(),
            Upgrade2 = parser.GetLong()
        };

        // 64 inventory slots (item_id, quantity pairs)
        for (int m = 0; m < 64; m++)
        {
            t.Inventory[m] = new InventorySlot
            {
                ItemId = parser.GetLong(),
                Quantity = parser.GetLong()
            };
        }

        // Attributes
        t.Attributes = new Attributes
        {
            Str = (int)parser.GetLong(),
            Agi = (int)parser.GetLong(),
            Int = (int)parser.GetLong(),
            Cha = (int)parser.GetLong(),
            Level = (int)parser.GetLong()
        };

        // Weapon proficiencies
        t.Proficiencies = new WeaponProficiencies
        {
            OneHanded = parser.GetLong(),
            TwoHanded = parser.GetLong(),
            Polearm = parser.GetLong(),
            Archery = parser.GetLong(),
            Crossbow = parser.GetLong(),
            Throwing = parser.GetLong(),
            Firearm = parser.GetLong()
        };

        // Skills (6 encoded values)
        for (int s = 0; s < 6; s++)
            t.Skills[s] = parser.GetLong();

        // Face codes (8 values)
        for (int f = 0; f < 8; f++)
            t.Face[f] = parser.GetWord();

        return t;
    }

    public static void SaveTroops(DataRepository repo, string filePath)
    {
        using var writer = new StreamWriter(filePath, false, System.Text.Encoding.ASCII);

        // Write version header
        writer.Write("troopsfile version 3 ");
        writer.Write(repo.Troops.Count);
        writer.Write("\r\n");

        foreach (var t in repo.Troops)
        {
            // Ensure trp_ prefix (preserved from VB6: SaveTroopFile mutates strID in place)
            var strId = t.StrId;
            if (!strId.StartsWith("trp_"))
                strId = "trp_" + strId;

            // Line 1: core fields
            writer.Write($" {strId} {t.Name} {t.PluralName} {t.UnknownWarband} {t.Flags}" +
                        $" {t.Scene} {t.Reserved} {t.Faction} {t.Upgrade1} {t.Upgrade2}\r\n");

            // Line 2: inventory
            writer.Write("  ");
            for (int m = 0; m < 64; m++)
                writer.Write($"{t.Inventory[m].ItemId} {t.Inventory[m].Quantity} ");
            writer.Write("\r\n");

            // Line 3: attributes
            writer.Write($"  {t.Attributes.Str} {t.Attributes.Agi} {t.Attributes.Int} {t.Attributes.Cha} {t.Attributes.Level}\r\n");

            // Line 4: proficiencies
            writer.Write($" {t.Proficiencies.OneHanded} {t.Proficiencies.TwoHanded} {t.Proficiencies.Polearm}" +
                        $" {t.Proficiencies.Archery} {t.Proficiencies.Crossbow} {t.Proficiencies.Throwing} {t.Proficiencies.Firearm}\r\n");

            // Line 5: skills
            writer.Write($"{t.Skills[0]} {t.Skills[1]} {t.Skills[2]} {t.Skills[3]} {t.Skills[4]} {t.Skills[5]} \r\n");

            // Line 6: face codes
            writer.Write($"  {t.Face[0]} {t.Face[1]} {t.Face[2]} {t.Face[3]} {t.Face[4]} {t.Face[5]} {t.Face[6]} {t.Face[7]} \r\n");

            // Blank separator line
            writer.Write("\r\n");
        }
    }
}
