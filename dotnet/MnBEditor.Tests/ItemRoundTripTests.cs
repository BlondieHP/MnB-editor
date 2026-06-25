using MnBEditor.Core;
using MnBEditor.Core.Models;

namespace MnBEditor.Tests;

public class ItemRoundTripTests
{
    [Fact]
    public void ParseSimpleItems_ParsesCorrectFields()
    {
        // Simplified test: 2 items with no factions or triggers
        var data = string.Join("\r\n",
            "itemsfile version 3 2",
            " itm_sword Sword sword_tex 1 mesh1 0 1 655370 1000 0 2.5 100 0 0 0 0 100 90 0 95 0 10 30 0 0",
            " itm_shield Shield shield_tex 1 mesh2 0 7 167772194 500 0 3.0 50 0 0 0 0 80 0 0 0 0 0 0 0 0"
        );

        var repo = new DataRepository();
        var parser = new FileParser(System.Text.Encoding.ASCII.GetBytes(data));

        // Read header
        parser.GetWord(); // itemsfile
        parser.GetWord(); // version
        parser.GetWord(); // 3
        var count = (int)parser.GetLong();
        Assert.Equal(2, count);

        // Parse items
        for (int i = 0; i < count; i++)
        {
            var item = ReadItemFromParser(parser);
            item.Id = i;
            repo.Items.Add(item);
        }

        Assert.Equal(2, repo.Items.Count);
        Assert.Equal("itm_sword", repo.Items[0].DbName);
        Assert.Equal("Sword", repo.Items[0].DisplayName);
        Assert.Equal(1000, repo.Items[0].Price);
        Assert.Equal("2.5", repo.Items[0].Weight);

        Assert.Equal("itm_shield", repo.Items[1].DbName);
        Assert.Equal("Shield", repo.Items[1].DisplayName);
        Assert.Equal(500, repo.Items[1].Price);
        Assert.Equal("7", repo.Items[1].ItemType);
    }

    [Fact]
    public void ParseItemWithFaction_ReadsFactionData()
    {
        var data = string.Join("\r\n",
            "itemsfile version 3 1",
            " itm_faction_sword Sword tex 1 m 0 1 655370 1000 0 2.5 100 0 0 0 0 100 90 0 95 0 10 30 1 3 0"
        );

        var parser = new FileParser(System.Text.Encoding.ASCII.GetBytes(data));
        parser.GetWord(); parser.GetWord(); parser.GetWord(); // header
        parser.GetLong(); // count

        var item = ReadItemFromParser(parser);
        Assert.Equal("itm_faction_sword", item.DbName);
        Assert.Equal(1, item.FactionCount);
        Assert.Single(item.Factions);
        Assert.Equal(3, item.Factions[0].Id);
        Assert.Equal(0, item.TriggerCount);
    }

    [Fact]
    public void GetWord_SkipsWhitespace_ReturnsTokens()
    {
        var parser = new FileParser("hello  world   123\r\nfoo"u8.ToArray());

        Assert.Equal("hello", parser.GetWord());
        Assert.Equal("world", parser.GetWord());
        Assert.Equal("123", parser.GetWord());
        Assert.Equal("foo", parser.GetWord());
        Assert.Equal("", parser.GetWord()); // end of file
    }

    [Fact]
    public void GetLine_ReadsUntilCrLf_ReturnsContent()
    {
        var parser = new FileParser("line1 content\r\nline2 more text\r\n"u8.ToArray());

        Assert.Equal("line1 content", parser.GetLine());
        Assert.Equal("line2 more text", parser.GetLine());
    }

    [Fact]
    public void CsvLoader_LoadsItemNames_Correctly()
    {
        // Arrange
        var repo = new DataRepository { Language = "cns" };
        repo.Items.Add(new Item { DbName = "itm_test_sword" });
        repo.Items.Add(new Item { DbName = "itm_test_shield" });

        // Create temp CSV
        var csvPath = Path.GetTempFileName();
        File.WriteAllText(csvPath, "itm_test_sword|Test Sword\nitm_test_shield|Test Shield\n");

        try
        {
            // Act
            CsvLoader.LoadItemCsv(repo, csvPath);

            // Assert
            Assert.Equal("Test Sword", repo.Items[0].CsvName);
            Assert.Equal("Test Shield", repo.Items[1].CsvName);
        }
        finally
        {
            File.Delete(csvPath);
        }
    }

    [Fact]
    public void DataRepository_Events_FireOnItemChange()
    {
        var repo = new DataRepository();
        repo.Items.Add(new Item { DbName = "test" });

        int? changedIndex = null;
        repo.ItemChanged += (idx) => changedIndex = idx;

        repo.NotifyItemChanged(0);

        Assert.Equal(0, changedIndex);
    }

    // Helper: replicates the ReadItem logic without depending on ItemLoader static class
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
            var trigger = new Trigger { TriggerType = parser.GetDouble() };
            trigger.ActionCount = parser.GetLong();
            for (int j = 0; j < trigger.ActionCount; j++)
            {
                var op = new Operation { OpCode = parser.GetWord() };
                op.ParamCount = parser.GetLong();
                for (int k = 0; k < op.ParamCount; k++)
                    op.Parameters.Add(new OperationParam { Value = parser.GetWord() });
                trigger.Actions.Add(op);
            }
            item.Triggers.Add(trigger);
        }

        return item;
    }
}
