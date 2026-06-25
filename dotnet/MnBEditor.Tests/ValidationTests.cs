using MnBEditor.Core;
using MnBEditor.Core.Models;

namespace MnBEditor.Tests;

public class ValidationTests
{
    [Fact]
    public void Validate_ValidData_ReturnsTrue()
    {
        var repo = CreateValidRepo();
        var validator = new DataValidator();

        var result = validator.Validate(repo);

        Assert.True(result);
        Assert.Empty(validator.Errors);
    }

    [Fact]
    public void Validate_TroopWithEmptyName_ReportsError()
    {
        var repo = CreateValidRepo();
        repo.Troops[0].Name = "";
        var validator = new DataValidator();

        var result = validator.Validate(repo);

        Assert.False(result);
        Assert.Contains(validator.Errors, e => e.Contains("empty Name"));
    }

    [Fact]
    public void Validate_ItemWithNegativePrice_ReportsError()
    {
        var repo = CreateValidRepo();
        repo.Items[0].Price = -1;
        var validator = new DataValidator();

        var result = validator.Validate(repo);

        Assert.False(result);
        Assert.Contains(validator.Errors, e => e.Contains("negative"));
    }

    [Fact]
    public void Validate_ItemWithBadFactionRef_ReportsError()
    {
        var repo = CreateValidRepo();
        repo.Items[0].Factions.Add(new FactionRef { Id = 999 });
        var validator = new DataValidator();

        var result = validator.Validate(repo);

        Assert.False(result);
        Assert.Contains(validator.Errors, e => e.Contains("non-existent faction"));
    }

    [Fact]
    public void Validate_PartyWithBadTroopRef_ReportsError()
    {
        var repo = CreateValidRepo();
        repo.Parties[0].Stacks.Add(new PartyStack { TroopId = 9999, Min = 1, Max = 5 });
        var validator = new DataValidator();

        var result = validator.Validate(repo);

        Assert.False(result);
        Assert.Contains(validator.Errors, e => e.Contains("non-existent troop"));
    }

    [Fact]
    public void Validate_WrongRelationshipCount_ReportsError()
    {
        var repo = CreateValidRepo();
        repo.Factions[0].Relationships.Clear(); // Should have N_Faction entries
        var validator = new DataValidator();

        var result = validator.Validate(repo);

        Assert.False(result);
        Assert.Contains(validator.Errors, e => e.Contains("relationships"));
    }

    private static DataRepository CreateValidRepo()
    {
        var repo = new DataRepository();

        // Add 2 factions (needed for relationships)
        repo.Factions.Add(new Faction { StrId = "fac_1", Name = "Faction 1" });
        repo.Factions.Add(new Faction { StrId = "fac_2", Name = "Faction 2" });
        // Fix relationship counts
        foreach (var f in repo.Factions)
        {
            for (int i = 0; i < repo.Factions.Count; i++)
                f.Relationships.Add(new Relationship { FactionId = i, Value = 0 });
        }

        // Add 3 troops
        repo.Troops.Add(new Troop { StrId = "trp_player", Name = "Player", Faction = 0 });
        repo.Troops.Add(new Troop { StrId = "trp_guard", Name = "Guard", Faction = 0 });
        repo.Troops.Add(new Troop { StrId = "trp_bandit", Name = "Bandit", Faction = 1 });

        // Add 2 items
        repo.Items.Add(new Item { DbName = "itm_sword", DisplayName = "Sword", Price = 100, Abundance = 50, HitPoints = 80 });
        repo.Items.Add(new Item { DbName = "itm_shield", DisplayName = "Shield", Price = 50, Abundance = 30, HitPoints = 60 });

        // Add party template
        repo.PartyTemplates.Add(new PartyTemplate { PtId = "pt_bandits", Name = "Bandits" });

        // Add 1 party
        repo.Parties.Add(new Party { StrId = "p_bandit_lair", Name = "Bandit Lair", Faction = 0, Template = 0 });

        // Add 1 scene
        repo.Scenes.Add(new Scene { StrId = "scn_test", Name = "Test Scene" });

        return repo;
    }
}
