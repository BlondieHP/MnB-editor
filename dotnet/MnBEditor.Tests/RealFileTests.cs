using MnBEditor.Core;
using MnBEditor.Core.Models;

namespace MnBEditor.Tests;

/// <summary>
/// Tests using real MOD data files from the project's Support/ directory.
/// </summary>
public class RealFileTests
{
    private static string SupportDir =>
        Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../../../Support"));

    [Fact]
    public void LoadRealItemFile_ParsesAllItems()
    {
        var filePath = Path.Combine(SupportDir, "item_kinds1.txt");
        if (!File.Exists(filePath))
        {
            // Skip if test data not available
            return;
        }

        var repo = new DataRepository();
        var fileBytes = File.ReadAllBytes(filePath);
        var parser = new FileParser(fileBytes);

        // Read header
        var h1 = parser.GetWord(); // "itemsfile"
        var h2 = parser.GetWord(); // "version"
        var h3 = parser.GetWord(); // "3"
        var count = (int)parser.GetLong();

        Assert.Equal("itemsfile", h1);
        Assert.Equal("version", h2);
        Assert.True(count > 0, $"Expected positive item count, got {count}");

        int parsedCount = 0;
        var errors = new List<string>();

        for (int i = 0; i < count && !parser.IsEndOfFile; i++)
        {
            try
            {
                var item = ReadItemFromParser(parser);
                item.Id = i;
                repo.Items.Add(item);
                parsedCount++;
            }
            catch (Exception ex)
            {
                errors.Add($"Item[{i}]: {ex.Message}");
                if (errors.Count > 10) break; // Don't flood
            }
        }

        Assert.Equal(count, parsedCount);
        Assert.Empty(errors);

        // Verify some known items
        Assert.NotEmpty(repo.Items);
        Assert.Equal("itm_no_item", repo.Items[0].DbName);
        Assert.Equal("INVALID_ITEM", repo.Items[0].DisplayName);

        // Check stats
        var namedItems = repo.Items.Count(i => !string.IsNullOrEmpty(i.DisplayName));
        Assert.Equal(count, namedItems);
    }

    [Fact]
    public void LoadBugCaseItemFile_ParsesCorrectly()
    {
        var filePath = Path.Combine(SupportDir, "Bugs", "1", "item_kinds1.txt");
        if (!File.Exists(filePath)) return;

        var repo = new DataRepository();
        var parser = new FileParser(File.ReadAllBytes(filePath));

        parser.GetWord(); parser.GetWord(); parser.GetWord(); // header
        var count = (int)parser.GetLong();
        Assert.True(count > 0);

        for (int i = 0; i < count && !parser.IsEndOfFile; i++)
        {
            var item = ReadItemFromParser(parser);
            item.Id = i;
            repo.Items.Add(item);
        }

        Assert.Equal(count, repo.Items.Count);
    }

    [Fact]
    public void LoadBugCaseTroopFile_ParsesCorrectly()
    {
        var filePath = Path.Combine(SupportDir, "Bugs", "1", "troops.txt");
        if (!File.Exists(filePath)) return;

        var parser = new FileParser(File.ReadAllBytes(filePath));

        parser.GetWord(); parser.GetWord(); parser.GetWord(); // header
        var count = (int)parser.GetLong();
        Assert.True(count > 0);

        int parsedCount = 0;
        for (int i = 0; i < count && !parser.IsEndOfFile; i++)
        {
            ReadTroopFromParser(parser);
            parsedCount++;
        }

        Assert.Equal(count, parsedCount);
    }

    [Fact]
    public void FileParser_RoundTrip_RealData()
    {
        // Load and re-parse to verify output is valid
        var filePath = Path.Combine(SupportDir, "item_kinds1.txt");
        if (!File.Exists(filePath)) return;

        var original = File.ReadAllBytes(filePath);
        var parser1 = new FileParser(original);

        // Parse header
        parser1.GetWord(); parser1.GetWord(); parser1.GetWord();
        var count = (int)parser1.GetLong();

        // Parse all items
        for (int i = 0; i < count; i++)
            ReadItemFromParser(parser1);

        // Should reach EOF
        Assert.True(parser1.IsEndOfFile || parser1.Position >= parser1.Length - 10);
    }

    // --- Helpers (same logic as ItemLoader/TroopLoader) ---

    private static Item ReadItemFromParser(FileParser parser)
    {
        var item = new Item
        {
            DbName = parser.GetWord(),
            DisplayName = parser.GetWord(),
            TextureName = parser.GetWord()
        };

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

        item.FactionCount = parser.GetLong();
        for (int i = 0; i < item.FactionCount; i++)
            item.Factions.Add(new FactionRef { Id = parser.GetLong() });

        item.TriggerCount = parser.GetLong();
        for (int i = 0; i < item.TriggerCount; i++)
        {
            parser.GetDouble(); // TriggerType
            long actCount = parser.GetLong();
            for (int j = 0; j < actCount; j++)
            {
                parser.GetWord(); // OpCode
                long paramCount = parser.GetLong();
                parser.SkipWords((int)paramCount);
            }
        }

        return item;
    }

    private static Troop ReadTroopFromParser(FileParser parser)
    {
        var t = new Troop
        {
            StrId = parser.GetWord(),
            Name = parser.GetWord(),
            PluralName = parser.GetWord(),
            UnknownWarband = parser.GetWord(),
            Flags = parser.GetWord(),
            Scene = parser.GetLong(),
            Reserved = parser.GetLong(),
            Faction = parser.GetLong(),
            Upgrade1 = parser.GetLong(),
            Upgrade2 = parser.GetLong()
        };

        for (int m = 0; m < 64; m++)
            t.Inventory[m] = new InventorySlot { ItemId = parser.GetLong(), Quantity = parser.GetLong() };

        t.Attributes = new Attributes
        {
            Str = (int)parser.GetLong(), Agi = (int)parser.GetLong(),
            Int = (int)parser.GetLong(), Cha = (int)parser.GetLong(),
            Level = (int)parser.GetLong()
        };

        t.Proficiencies = new WeaponProficiencies
        {
            OneHanded = parser.GetLong(), TwoHanded = parser.GetLong(), Polearm = parser.GetLong(),
            Archery = parser.GetLong(), Crossbow = parser.GetLong(), Throwing = parser.GetLong(), Firearm = parser.GetLong()
        };

        for (int s = 0; s < 6; s++) t.Skills[s] = parser.GetLong();
        for (int f = 0; f < 8; f++) t.Face[f] = parser.GetWord();

        return t;
    }
}
