namespace MnBEditor.Core;

/// <summary>
/// All game constants ported from Constants.bas.
/// </summary>
public static class Constants
{
    // Item types
    public const uint ItpTypeHorse = 0x1;
    public const uint ItpTypeOneHandedWpn = 0x2;
    public const uint ItpTypeTwoHandedWpn = 0x3;
    public const uint ItpTypePolearm = 0x4;
    public const uint ItpTypeArrows = 0x5;
    public const uint ItpTypeBolts = 0x6;
    public const uint ItpTypeShield = 0x7;
    public const uint ItpTypeBow = 0x8;
    public const uint ItpTypeCrossbow = 0x9;
    public const uint ItpTypeThrown = 0xA;
    public const uint ItpTypeGoods = 0xB;
    public const uint ItpTypeHeadArmor = 0xC;
    public const uint ItpTypeBodyArmor = 0xD;
    public const uint ItpTypeFootArmor = 0xE;
    public const uint ItpTypeHandArmor = 0xF;
    public const uint ItpTypePistol = 0x10;
    public const uint ItpTypeMusket = 0x11;
    public const uint ItpTypeBullets = 0x12;
    public const uint ItpTypeAnimal = 0x13;
    public const uint ItpTypeBook = 0x14;
    public const uint ItpTypeMask = 0x1F;

    // Item property flag bits
    public const int ItpUnique = 12;
    public const int ItpAlwaysLoot = 13;
    public const int ItpNoParry = 14;
    public const int ItpDefaultAmmo = 15;
    public const int ItpMerchandise = 16;
    public const int ItpWoodenAttack = 17;
    public const int ItpWoodenParry = 18;
    public const int ItpFood = 19;
    public const int ItpCantReloadOnHorseback = 20;
    public const int ItpTwoHanded = 21;
    public const int ItpPrimary = 22;
    public const int ItpSecondary = 23;
    public const int ItpCoversLegs = 24;
    public const int ItpConsumable = 25;
    public const int ItpBonusAgainstShield = 26;
    public const int ItpPenaltyWithShield = 27;
    public const int ItpCantUseOnHorseback = 28;
    public const int ItpCivilian = 29;
    public const int ItpFitToHead = 30;
    public const int ItpCoversHead = 31;
    public const int ItpCrushThrough = 32;
    public const int ItpKnockBack = 33;
    public const int ItpRemoveItemOnUse = 34;
    public const int ItpUnbalanced = 35;
    public const int ItpCoversBeard = 36;
    public const int ItpNoPickUpFromGround = 37;
    public const int ItpCanKnockDown = 38;

    // Tag types
    public const int TagRegister = 1;
    public const int TagVariable = 2;
    public const int TagString = 3;
    public const int TagItem = 4;
    public const int TagTroop = 5;
    public const int TagFaction = 6;
    public const int TagQuest = 7;
    public const int TagPartyTpl = 8;
    public const int TagParty = 9;
    public const int TagScene = 10;
    public const int TagMissionTpl = 11;
    public const int TagMenu = 12;
    public const int TagScript = 13;
    public const int TagParticleSys = 14;
    public const int TagSceneProp = 15;
    public const int TagSound = 16;
    public const int TagLocalVariable = 17;
    public const int TagMapIcon = 18;
    public const int TagSkill = 19;
    public const int TagMesh = 20;
    public const int TagPresentation = 21;
    public const int TagQuickString = 22;
    public const int TagTrack = 23;
    public const int TagTableau = 24;
    public const int TagAnimation = 25;

    // Skill constants
    public const int SklTrade = 0;
    public const int SklLeadership = 1;
    public const int SklPrisonerManagement = 2;
    public const int SklPersuasion = 7;
    public const int SklEngineer = 8;
    public const int SklFirstAid = 9;
    public const int SklSurgery = 10;
    public const int SklWoundTreatment = 11;
    public const int SklInventoryManagement = 12;
    public const int SklSpotting = 13;
    public const int SklPathfinding = 14;
    public const int SklTactics = 15;
    public const int SklTracking = 16;
    public const int SklTrainer = 17;
    public const int SklLooting = 22;
    public const int SklHorseArchery = 23;
    public const int SklRiding = 24;
    public const int SklAthletics = 25;
    public const int SklShield = 26;
    public const int SklWeaponMaster = 27;
    public const int SklPowerDraw = 33;
    public const int SklPowerThrow = 34;
    public const int SklPowerStrike = 35;
    public const int SklIronflesh = 36;

    // Edit info types
    public const int EditInfoTroopsCount = 0;
    public const int EditInfoItemsCount = 1;
    public const int EditInfoScenesCount = 2;
    public const int EditInfoFactionsCount = 3;
    public const int EditInfoPartyTemplatesCount = 4;
    public const int EditInfoPartiesCount = 5;
    public const int EditInfoMapIconsCount = 6;
    public const int EditInfoSoundsCount = 7;
    public const int EditInfoSoundRessCount = 8;
    public const int EditInfoPSysCount = 9;
    public const int EditInfoTabMatCount = 10;
    public const int EditInfoMeshCount = 11;
    public const int EditInfoTimeTrgCount = 12;
    public const int EditInfoStringsCount = 13;

    // Trigger events
    public const int TiOnInitItem = -50;
    public const int TiOnWeaponAttack = -51;
    public const int TiOnMissileHit = -52;
    public const int TiOnInitMapIcon = -70;

    // Max constants
    public const int NPos = 65;
    public const int MaxIndentation = 16;
    public const int MaxInventorySlots = 64;
    public const int MaxPartyStacks = 6;
    public const int MaxSkillCount = 42;
    public const int MaxTagCount = 26;
}
