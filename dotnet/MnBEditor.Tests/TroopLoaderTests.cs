using MnBEditor.Core;
using MnBEditor.Core.Models;

namespace MnBEditor.Tests;

public class TroopLoaderTests
{
    [Fact]
    public void ParseTroops_ParsesCoreFields()
    {
        // Arrange: minimal troops.txt with 2 troops
        var data = string.Join("\r\n",
            "troopsfile version 3 2",
            " trp_player Player Players 0 0 0 0 0 0 0 " +
                string.Join(" ", Enumerable.Repeat("0 0", 64)) + " " +  // 64 empty inventory slots
                "10 10 10 10 5 " +    // attributes: str,agi,int,cha,level
                "100 100 100 80 80 80 0 " +  // proficiencies
                "0 0 0 0 0 0 " +      // skills
                "0 0 0 0 0 0 0 0",    // face codes
            " trp_bandit Bandit Bandits 0 0 0 0 0 0 0 " +
                string.Join(" ", Enumerable.Repeat("0 0", 64)) + " " +
                "5 8 3 2 3 " +
                "60 50 40 30 20 50 0 " +
                "0 0 0 0 0 0 " +
                "0 0 0 0 0 0 0 0"
        );

        var repo = new DataRepository();
        // Add empty faction for reference
        repo.Factions.Add(new Faction { StrId = "fac_neutral", Name = "Neutral" });

        var parser = new FileParser(System.Text.Encoding.ASCII.GetBytes(data));

        // Act: parse header + troops
        parser.GetWord(); parser.GetWord(); parser.GetWord(); // "troopsfile", "version", "3"
        var count = (int)parser.GetLong();
        Assert.Equal(2, count);

        for (int i = 0; i < count; i++)
        {
            var t = ReadTroop(parser);
            t.Id = i;
            repo.Troops.Add(t);
        }

        // Assert
        Assert.Equal(2, repo.Troops.Count);
        Assert.Equal("trp_player", repo.Troops[0].StrId);
        Assert.Equal("Player", repo.Troops[0].Name);
        Assert.Equal("Players", repo.Troops[0].PluralName);
        Assert.Equal(10, repo.Troops[0].Attributes.Str);
        Assert.Equal(5, repo.Troops[0].Attributes.Level);
        Assert.Equal(100, repo.Troops[0].Proficiencies.OneHanded);

        Assert.Equal("trp_bandit", repo.Troops[1].StrId);
        Assert.Equal("Bandit", repo.Troops[1].Name);
        Assert.Equal(5, repo.Troops[1].Attributes.Str);
        Assert.Equal(8, repo.Troops[1].Attributes.Agi);
        Assert.Equal(60, repo.Troops[1].Proficiencies.OneHanded);
    }

    [Fact]
    public void ParseTroops_WithInventory_ParsesInventorySlots()
    {
        var inventoryFields = new List<string>();
        for (int i = 0; i < 64; i++)
        {
            if (i < 4)
            {
                inventoryFields.Add($"{(i + 1) * 10}");
                inventoryFields.Add($"{i + 1}");
            }
            else
            {
                inventoryFields.Add("0");
                inventoryFields.Add("0");
            }
        }

        var data = string.Join("\r\n",
            "troopsfile version 3 1",
            " trp_test Test Tests 0 0 0 0 0 0 0 " +
                string.Join(" ", inventoryFields) + " " +
                "10 10 10 10 1 " +
                "50 50 50 30 30 30 0 " +
                "0 0 0 0 0 0 " +
                "0 0 0 0 0 0 0 0"
        );

        var parser = new FileParser(System.Text.Encoding.ASCII.GetBytes(data));
        parser.GetWord(); parser.GetWord(); parser.GetWord(); parser.GetLong();

        var troop = ReadTroop(parser);

        Assert.Equal(10, troop.Inventory[0].ItemId);
        Assert.Equal(1, troop.Inventory[0].Quantity);
        Assert.Equal(20, troop.Inventory[1].ItemId);
        Assert.Equal(2, troop.Inventory[1].Quantity);
        Assert.Equal(40, troop.Inventory[3].ItemId);
        Assert.Equal(4, troop.Inventory[3].Quantity);
        // Slot 4 should be empty
        Assert.Equal(0, troop.Inventory[4].ItemId);
    }

    // Helper: replicates ReadTroop from TroopLoader
    private static Troop ReadTroop(FileParser parser)
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
