using MnBEditor.Core.Models;

namespace MnBEditor.Core;

/// <summary>
/// Loaders for Scene, PartyTemplate, and Party entity types.
/// </summary>
public static class EntityLoaders
{
    // ====== SCENE ======
    // Format: strID, strName, Flags, MeshName, BodyName, p0.X, p0.Y, p1.X, p1.Y,
    //         WaterLevel, TerrainCode, AccessCount, [Accesses], ChestCount, [Chests],
    //         OuterTerrainType

    public static List<Scene> LoadScenes(string filePath)
    {
        var p = new FileParser(File.ReadAllBytes(filePath));
        p.GetWord(); p.GetWord(); p.GetWord();
        int count = (int)p.GetLong();
        var list = new List<Scene>(count);
        for (int i = 0; i < count; i++)
        {
            var s = new Scene
            {
                Id = i, StrId = p.GetWord(), Name = p.GetWord(),
                Flags = p.GetWord(), MeshName = p.GetWord(), BodyName = p.GetWord(),
                Point1 = new Point2D(p.GetDouble(), p.GetDouble()),
                Point2 = new Point2D(p.GetDouble(), p.GetDouble()),
                WaterLevel = p.GetDouble(), TerrainCode = p.GetWord()
            };
            s.AccessCount = p.GetLong();
            for (int j = 0; j < s.AccessCount; j++) s.Accesses.Add(p.GetWord());
            s.ChestCount = p.GetLong();
            for (int j = 0; j < s.ChestCount; j++) s.Chests.Add(new FactionRef { Id = p.GetLong() });
            s.OuterTerrainType = p.GetWord();
            list.Add(s);
        }
        return list;
    }

    // ====== PARTY TEMPLATE ======
    // Format: ptID, ptName, Flags, Menu, Faction, Personality,
    //         6×(troopID, [min, max, flags if troopID>=0])

    public static List<PartyTemplate> LoadPartyTemplates(string filePath)
    {
        var p = new FileParser(File.ReadAllBytes(filePath));
        p.GetWord(); p.GetWord(); p.GetWord();
        int count = (int)p.GetLong();
        var list = new List<PartyTemplate>(count);
        for (int i = 0; i < count; i++)
        {
            var pt = new PartyTemplate
            {
                Id = i, PtId = p.GetWord(), Name = p.GetWord(),
                Flags = p.GetWord(), Menu = p.GetWord(),
                Faction = p.GetLong(), Personality = p.GetLong()
            };
            for (int j = 0; j < 6; j++)
            {
                long tid = p.GetLong();
                pt.Stacks[j] = tid >= 0
                    ? new PartyStack { TroopId = tid, Min = p.GetLong(), Max = p.GetLong(), Flags = p.GetLong() }
                    : new PartyStack { TroopId = -1, Min = -1, Max = -1, Flags = 0 };
            }
            list.Add(pt);
        }
        return list;
    }

    // ====== PARTY ======
    // Format: UnknownTitle, ID, id2, strID, strName, Flags, Menu, Template,
    //         Faction, Personality(1), Personality(2), AI_Behavior, AI_Target,
    //         reserved, 3×(posX, posY), UnknownStr, StacksCount, [Stacks], Degree

    public static List<Party> LoadParties(string filePath)
    {
        var p = new FileParser(File.ReadAllBytes(filePath));
        p.GetWord(); p.GetWord(); p.GetWord();
        int count = (int)p.GetLong();
        int count2 = (int)p.GetLong(); // N_Party2
        var list = new List<Party>(count);
        for (int i = 0; i < count; i++)
        {
            var party = new Party
            {
                Id = i, UnknownTitle = p.GetLong(), Id2 = p.GetLong(),
                StrId = p.GetWord(), Name = p.GetWord(),
                Flags = p.GetWord(), Menu = p.GetWord(),
                Template = p.GetLong(), Faction = p.GetLong()
            };
            party.Personality[0] = p.GetLong();
            party.Personality[1] = p.GetLong();
            party.AiBehavior = p.GetWord();
            party.AiTarget = p.GetWord();
            party.Reserved = p.GetLong();

            for (int j = 0; j < 3; j++)
                party.InitPositions[j] = new Point2D(p.GetDouble(), p.GetDouble());

            party.UnknownStr = p.GetWord();

            party.StacksCount = p.GetLong();
            for (int j = 0; j < party.StacksCount; j++)
                party.Stacks.Add(new PartyStack
                {
                    TroopId = p.GetLong(), Min = p.GetLong(),
                    Max = p.GetLong(), Flags = p.GetLong()
                });

            party.Degree = p.GetWord();
            list.Add(party);
        }
        return list;
    }
}
