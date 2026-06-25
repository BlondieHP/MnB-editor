Attribute VB_Name = "Models"

Option Explicit

Public Tags(1 To 26) As String

Type type_Attrib
    '#  9) Attributes (int): Example usage:
    '#    str_6|agi_6|int_4|cha_5|level(5)
    strPoint As Integer
    agiPoint As Integer
    intPoint As Integer
    chaPoint As Integer
    level As Integer
End Type

Type Type_WeaponProf
    '# 10) Weapon proficiencies (int): Example usage:
    '# wp_one_handed(55)|wp_two_handed(90)|wp_polearm(36)|wp_archery(80)|wp_crossbow(24)|wp_throwing(45)
    one_handed As Long
    two_handed As Long
    polearm As Long
    archery As Long
    crossbow As Long
    throwing As Long
    firearm As Long
End Type

Type Type_XY
    X As Long
    Y As Long
End Type

Type Type_XY_Index
    X As Long
    strX As String
    Y As Long
End Type

Type Type_strXY
    X As String
    Y As String
End Type

Type Type_strXYZ
    X As String
    Y As String
    Z As String
End Type

Type Type_dblXYZ
    X As Double
    Y As Double
    Z As Double
End Type

Public Type Type_RelationShip
    ID As Long
    strID As String
    Value As Double
End Type

Public Type Type_Chest
    ID As Long
    strID As String
End Type

Public Type Type_Param
    Value As String
    strID As String
End Type

Public Type Type_ResourceInSound
    ID As Long
    strID As String
    Unknown As Long
End Type

Type Type_Sound
    ID As Long
    sndName As String
    Flags As String
    
    ResourceCount As Long
    Resource() As Type_ResourceInSound
    
    Edit As Boolean
End Type

Type Type_SoundResource
    ID As Long
    sndName As String
    Flags As String
    
    Edit As Boolean
End Type

Type Type_Troops
    '#  Each troop contains the following fields:
    '#  1) Troop id (string): used for referencing troops in other files. The prefix trp_ is automatically added before each troop-id .
    '#  2) Toop name (string).
    '#  3) Plural troop name (string).
    '#  4) Troop flags (int). See header_troops.py for a list of available flags
    '#  5) Scene (int) (only applicable to heroes) For example: scn_reyvadin_castle|entry(1) puts troop in reyvadin castle's first entry point
    '#  6) Reserved (int). Put constant "reserved" or 0.
    '#  7) Faction (int)
    '#  8) Inventory (list): Must be a list of items
    '#  9) Attributes (int): Example usage:
    '#           str_6|agi_6|int_4|cha_5|level(5)
    '# 10) Weapon proficiencies (int): Example usage:
    '#           wp_one_handed(55)|wp_two_handed(90)|wp_polearm(36)|wp_archery(80)|wp_crossbow(24)|wp_throwing(45)
    '#     The function wp(x) will create random weapon proficiencies close to Value x.
    '#     To make an expert archer with other weapon proficiencies close to 60 you can use something like:
    '#           wp_archery(160) | wp(60)
    '# 11) Skills (int): See header_skills.py to see a list of skills. Example:
    '#           knows_ironflesh_3|knows_power_strike_2|knows_athletics_2|knows_riding_2
    '# 12) Face code (int): You can obtain the face code by pressing ctrl+E in face generator screen
    '# 13) Face code (int)(2) (only applicable to regular troops, can be omitted for heroes):
    '#     The game will create random faces between Face code 1 and face code 2 for generated troops
    ID As Long
    strID As String
    strName As String
    strPtName As String
    
    '*1.1x 比1.011多的部分*
    unknown_warband(1 To 1) As String
    
    Flags As String
    
    Scene As Long
    SceneID As Long
    Scene_strID As String
    Entry As Long
    
    reserved As Long
    
    Faction As Long
    Faction_strID As String     '外引
    
    Upgrade1 As String
    Upgrade1_strID As String
    Upgrade2 As String
    Upgrade2_strID As String
    
    lstInventory(1 To 64) As Type_XY_Index
    
    tAttrib As type_Attrib
    WP As Type_WeaponProf
    Skills(1 To 6) As String
    Face(1 To 8) As String
    
    Edit As Boolean
    csvName As String
    csvName_pl As String
    
End Type

Type Type_Stacks
    ID As Long
    strID As String      '外引
    Min As Long
    Max As Long
    Flags As Long '0:member; 1:俘虏.
End Type

Public Type Type_Op_Block
    Op As String
    ParaNum As Long
    Para() As Type_Param
End Type

Type Type_Trigger
    tiOn As Double
    ActNum As Long
    tiAct() As Type_Op_Block
End Type

Type Type_PT
    '#  Each party template record contains the following fields:
    '#  1) Party-template id: used for referencing party-templates in other files.
    '#     The prefix pt_ is automatically added before each party-template id.
    '#  2) Party-template name.
    '#  3) Party flags. See header_parties.py for a list of available flags
    '#  4) Menu. ID of the menu to use when this party is met. The Value 0 uses the default party encounter system.
    '#  5) Faction
    '#  6) Personality. See header_parties.py for an explanation of personality flags.
    '#  7) List of stacks. Each stack record is a tuple that contains the following fields:
    '#    7.1) Troop-id.
    '#    7.2) Minimum number of troops in the stack.
    '#    7.3) Maximum number of troops in the stack.
    '#    7.4) Member flags(optional). Use pmf_is_prisoner to note that this member is a prisoner.
    '#     Note: There can be at most 6 stacks.
    ID As Long
    ptID As String
    ptName As String
    Flags As String
    Menu As String
    Faction As Long
    Faction_strID As String
    
    Personality As Long
    
    Stacks(1 To 6) As Type_Stacks
    
    Edit As Boolean
    csvName As String
    
End Type

Type Type_Party
'####################################################################################################################
'#  Each party record contains the following fields:
'#  1) Party id: used for referencing parties in other files.
'#     The prefix p_ is automatically added before each party id.
'#  2) Party name.
'#  3) Party flags. See header_parties.py for a list of available flags
'#  4) Menu. ID of the menu to use when this party is met. The value 0 uses the default party encounter system.
'#  5) Party-template. ID of the party template this party belongs to. Use pt_none as the default value.
'#  6) Faction.
'#  7) Personality. See header_parties.py for an explanation of personality flags.
'#  8) Ai-behavior
'#  9) Ai-target party
'# 10) Initial coordinates.
'# 11) List of stacks. Each stack record is a triple that contains the following fields:
'#   11.1) Troop-id.
'#   11.2) Number of troops in this stack.
'#   11.3) Member flags. Use pmf_is_prisoner to note that this member is a prisoner.
'# 12) Party direction in degrees [optional]
'####################################################################################################################
    UnknownTitle As Long
    ID As Long
    id2 As Long
    strID As String
    strName As String
    
    Flags As String
    MapIcon_strID As String
    
    Menu As String
    
    Template As Long
    Template_strID As String
    
    Faction As Long
    Faction_strID As String
    
    Personality(1 To 2) As Long
    AI_Behavior As String
    AI_Target As String
    AI_Target_strID As String
    
    reserved As Long
    InitPos(1 To 3) As Type_strXY
    
    UnknownStr As String
    StacksCount As Long
    Stacks() As Type_Stacks
    Degree As String        '单位:度
    
    Edit As Boolean
    csvName As String
    
End Type

Type Type_Item
    ID As Long '物品编号
    dbName As String '物品在数据库中的名字
    disname As String '物品在游戏中显示的名字
    texname As String '贴图的名字
    nmdl As Long '模型的数量，通常第一个模型是本体，其他是剑鞘
    mdlname() As String '模型的名字
    mdl_b() As String '模型参数
    'container(3) As String     '剑鞘的名字
    'container_binary(3) As String '剑鞘位置的代码
    itmType As String '物品性质
    Action As String '攻击动作
    price As Long '价格
    'prefix As Long  '前缀限制 v0.951
    Prefix As String  '前缀限制 v0.952
    'weight As Double '重量  'v0.951
    weight As String '重量  'v0.952
    abundance As Long '充裕度
    head_armor As Long
    body_armor As Long
    leg_armor As Long
    difficulty As Long
    hit_points As Long
    speed_rating As Long
    missile_speed As Long
    weapon_length As Long
    max_ammo As Long
    thrust_damage As Long
    swing_damage As Long
    
    FactionCount As Long
    Faction() As Type_Chest      '出售阵营
    
    TriggerCount As Long
    Trigger() As Type_Trigger  '对于火枪的处理
    
    Edit As Boolean
    csvName As String
    csvName_pl As String
End Type

Type Type_Faction
    ID As Long
    strID As String
    strName As String
    Flags As String
    lColor As String
    'RelationShip() As Double
    RelationShip() As Type_RelationShip
    
    reserved As Long
    
    csvName As String
    Edit As Boolean
End Type

Type Type_ImodBits
    ID As String
    csvName As String
End Type

Public Type Type_MapIcon
    ID As Long
    strID As String
    Flags As Long
    MeshName As String
    mScale As String
    Sound As Long
    Sound_sndName As String
    
    Offset(0 To 2) As String
    
    TriggerCount As Long
    Triggers() As Type_Trigger  '触发器
    
    Edit As Boolean
End Type

Public Type Type_TerrainInfo
    Code As String
    Length As Long
End Type

Type Type_Scene
     ID As Long
     strID As String
     strName As String
     Flags As String
     MeshName As String
     BodyName As String
     p(0 To 1) As tPoint
     WaterLevel As Double
     TerrainCode As String
     AccessCount As Long
     Accesses() As String
     
     ChestCount As Long
     Chests() As Type_Chest
     Outer_Terrain_Type As String
     
     Edit As Boolean
End Type

Type Double_XY
     X As Double
     Y As Double
End Type

Type Type_Particle_System
'  1) Particle system id (string)
'  2) Particle system flags (int). See header_particle_systems.py for a list of available flags
'  3) mesh-name.
''''
'  4) Num particles per second:    Number of particles emitted per second.
'  5) Particle Life:    Each particle lives this long (in seconds).
'  6) Damping:          How much particle's speed is lost due to friction.
'  7) Gravity strength: Effect of gravity. (Negative values make the particles float upwards.)
'  8) Turbulance size:  Size of random turbulance (in meters)
'  9) Turbulance strength: How much a particle is affected by turbulance.
''''
' 10,11) Alpha keys :    Each attribute is controlled by two keys and
' 12,13) Red keys   :    each key has two fields: (time, magnitude)
' 14,15) Green keys :    For example scale key (0.3,0.6) means
' 16,17) Blue keys  :    scale of each particle will be 0.6 at the
' 18,19) Scale keys :    time 0.3 (where time=0 means creation and time=1 means end of the particle)
'
' The magnitudes are interpolated in between the two keys and remain constant beyond the keys.
' Except the alpha always starts from 0 at time 0.
''''
' 20) Emit Box Size :   The dimension of the box particles are emitted from.
' 21) Emit velocity :   Particles are initially shot with this velocity.
' 22) Emit dir randomness
' 23) Particle rotation speed: Particles start to rotate with this (angular) speed (degrees per second).
' 24) Particle rotation damping: How quickly particles stop their rotation
    ID As Long
    strID As String
    Flags As String
    Mesh_Name As String
    Particles_Num As Long
    Life As Double
    Damping As Double
    Gravity As Double
    Turbulance_SZ As Double
    Turbulance_Str As Double
    
    Alphak(1) As Double_XY
    Redk(1) As Double_XY
    Greenk(1) As Double_XY
    Bluek(1) As Double_XY
    Scalek(1) As Double_XY
    
    EBSZ(2) As Double
    EV(2) As Double
    EDR As Double
    PRS As Double
    PRD As Double
    
    Edit As Boolean
End Type

Type Type_Tableau_Material
'#######################################################################################################################
'#  1) Tableau id (string)
'#  2) Tableau flags (int)
'#  3) Tableau sample material name (string).
'#  4) Tableau width (int).
'#  5) Tableau height (int).
'#  6) Tableau mesh min x (int): divided by 1000 and used when a mesh is auto-generated using the tableau material
'#  7) Tableau mesh min y (int): divided by 1000 and used when a mesh is auto-generated using the tableau material
'#  8) Tableau mesh max x (int): divided by 1000 and used when a mesh is auto-generated using the tableau material
'#  9) Tableau mesh max y (int): divided by 1000 and used when a mesh is auto-generated using the tableau material
'#  10) Operations block (list): A list of operations
'#     The operations block is executed when the tableau is activated.
'#######################################################################################################################
   ID As Long
   strID As String
   Flags As String
   Sample As String
   Width As Long
   Height As Long
   Min As Type_XY
   Max As Type_XY
   OpCount As Long
   OpBlock() As Type_Op_Block
   Edit As Boolean
End Type

Type Type_Mesh
'####################################################################################################################
'#  Each mesh record contains the following fields:
'#  1) Mesh id: used for referencing meshes in other files. The prefix mesh_ is automatically added before each mesh id.
'#  2) Mesh flags. See header_meshes.py for a list of available flags
'#  3) Mesh resource name: Resource name of the mesh
'#  4) Mesh translation on x axis: Will be done automatically when the mesh is loaded
'#  5) Mesh translation on y axis: Will be done automatically when the mesh is loaded
'#  6) Mesh translation on z axis: Will be done automatically when the mesh is loaded
'#  7) Mesh rotation angle over x axis: Will be done automatically when the mesh is loaded
'#  8) Mesh rotation angle over y axis: Will be done automatically when the mesh is loaded
'#  9) Mesh rotation angle over z axis: Will be done automatically when the mesh is loaded
'#  10) Mesh x scale: Will be done automatically when the mesh is loaded
'#  11) Mesh y scale: Will be done automatically when the mesh is loaded
'#  12) Mesh z scale: Will be done automatically when the mesh is loaded
'####################################################################################################################
    ID As Long
    strID As String
    Flags As String
    Resource_Name As String
    Translation As Type_dblXYZ
    Rotation_Angle As Type_dblXYZ
    Scale As Type_dblXYZ
    Edit As Boolean
End Type

Type Type_Time_Trigger
'####################################################################################################################
'#  Each trigger contains the following fields:
'# 1) Check interval: How frequently this trigger will be checked
'# 2) Delay interval: Time to wait before applying the consequences of the trigger
'#    After its conditions have been evaluated as true.
'# 3) Re-arm interval. How much time must pass after applying the consequences of the trigger for the trigger to become active again.
'#    You can put the constant ti_once here to make sure that the trigger never becomes active again after it fires once.
'# 4) Conditions block (list). This must be a valid operation block. See header_operations.py for reference.
'#    Every time the trigger is checked, the conditions block will be executed.
'#    If the conditions block returns true, the consequences block will be executed.
'#    If the conditions block is empty, it is assumed that it always evaluates to true.
'# 5) Consequences block (list). This must be a valid operation block. See header_operations.py for reference.
'####################################################################################################################
    ID As Long
    Check_Interval As Double
    Delay_Interval As Double
    Rearm_Interval As Double
    Condition() As Type_Op_Block
    ConditionsCount As Long
    Consequence() As Type_Op_Block
    ConsequencesCount As Long
    Edit As Boolean
End Type

Type Type_Global_Variable
    ID As Long
    VarName As String
    'Uses as Integer
End Type

Type Type_String
    ID As Long
    Name As String
    Str As String
    CSV As String
    Edit As Boolean
End Type


Public ShortTags(1 To 26) As Type_XY

Public Trps() As Type_Troops
Public PTs() As Type_PT
Public Parties() As Type_Party
Public itm() As Type_Item
Public Factions() As Type_Faction
Public Scenes() As Type_Scene
Public itmID() As Long
Public IMod() As Type_ImodBits
Public PSys() As Type_Particle_System
Public MapIcons() As Type_MapIcon
Public Sounds() As Type_Sound
Public SoundRess() As Type_SoundResource
Public TabMat() As Type_Tableau_Material
Public Mesh() As Type_Mesh
Public TimeTrg() As Type_Time_Trigger
Public gVars() As Type_Global_Variable
Public qStrs() As Type_String
Public Strs() As Type_String

Public ItmVersionInform(2) As String
Public PTVersionInform(2) As String
Public PartyVersionInform(2) As String
Public TrpsVersionInform(2) As String
Public FactionVersionInform(2) As String
Public SceneVersionInform(2) As String
Public PSysVersionInform(2) As String
Public MapIconVersionInform(2) As String
Public SoundVersionInform(2) As String
Public TabMatVersionInform(2) As String
Public MeshVersionInform(2) As String
Public TimeTrgVersionInform(2) As String
Public StringVersionInform(0) As String

Public N_Item As Long '装备总数
Public N_Troop As Long 'troop总数
Public N_PT As Long 'pt总数
Public N_Party As Long 'party总数
Public N_Party2 As Long 'party总数
Public N_Faction As Long '阵营总数
Public N_Scene As Long '场景总数
Public N_IMod As Long '物品前缀总数
Public N_PSys As Long '粒子系统总数
Public N_MapIcon As Long '大地图图标总数
Public N_Sound As Long '声音总数
Public N_SoundRes As Long '声音资源总数
Public N_TabMat As Long '可变素材总数
Public N_Mesh As Long '网格模型总数
Public N_TimeTrg As Long   '触发器总数
Public N_gVar As Long   '全局变量总数
Public N_qStr As Long   '快速字符串总数
Public N_Str As Long    '字符串总数

Public CurrentItmID As Long '当前装备号
Public CurrentItm As Type_Item '当前item拷贝
Public CurrentTrpID As Long '当前troop号
Public CurrentTrp As Type_Troops '当前troop拷贝
Public CurrentFactionID As Long '当前faction号
Public CurrentFaction As Type_Faction  '当前faction拷贝
Public CurPartyTemplateID As Long '当前PartyTemplate号.
Public CurPartyTemplate As Type_PT   '当前PartyTemplate拷贝
Public CurrentPartyID As Long '当前Party号.
Public CurrentParty As Type_Party   '当前Party拷贝
Public CurrentSceneID As Long '当前Scene号.
Public CurrentScene As Type_Scene   '当前Scene拷贝
Public CurrentMapIconID As Long '当前MapIcon号.
Public CurrentMapIcon As Type_MapIcon  '当前MapIcon拷贝
Public CurrentSoundID As Long '当前Sound号.
Public CurrentSound As Type_Sound  '当前Sound拷贝
Public CurrentSoundResID As Long '当前SoundResource号.
Public CurrentSoundRes As Type_SoundResource   '当前SoundResource拷贝
Public CurrentPSysID As Long    '当前Particles System号
Public CurrentPSys As Type_Particle_System   '当前Particles System拷贝
Public CurrentTabMatID As Long    '当前可变素材号
Public CurrentTabMat As Type_Tableau_Material    '当前可变素材拷贝
Public CurrentMeshID As Long    '当前网格模型号
Public CurrentMesh As Type_Mesh    '当前网格模型拷贝
Public CurrentTimeTrgID As Long    '当前触发器号
Public CurrentTimeTrg As Type_Time_Trigger     '当前触发器拷贝
'Public CurrentqStrID As Long    '当前快速字符串号
'Public CurrentqStr As Type_Quick_String     '当前快速字符串拷贝
Public CurrentStrID As Long    '当前字符串号
Public CurrentStr As Type_String     '当前字符串拷贝

Public TempItemTrigger As Type_Trigger    '临时物品trigger,供复制用
Public TempAct As Type_Op_Block    '临时act,供复制用

Public itmEditCount As Long
Public tpsEditCount As Long
Public tpsItmTabsFrameIndex As Integer '当前显示的

Public i64b_num1 As Integer64b
Public i64b_num2 As Integer64b

Public CountDream As Integer
Public DreamTeam(0 To 99) As Integer

Public gBackupFileFlag As Boolean

Public gSkillName(0 To 41) As String
Public gTroopsFlagChk(0 To 63) As String * 12

Public Type saveFlag
    changeAllPartyNumber As Boolean
    ptMax As Double
    ptMin As Double
End Type
Public gSaveFlag As saveFlag

Public IsLoadString As Boolean

Type Type_Global_Variable_Symbol
    iniSetting As String
    iniFileName As String
    ModIniFileName As String
    ModPath As String
    ModName As String
    ModBackUp As String
    CVSPath As String
    LastError As String
    Version As String
    Language As String
    Language_Edit As String
    Op_Set As String
    MBHome As String
    MBsaves As String
    MBsets As String
    
    FirstTimeEdit As Boolean
    EditInfo(13) As Long    'A_E
    InfoFinished As Boolean
    InitFinished As Boolean
End Type
Public MnBInfo As Type_Global_Variable_Symbol

Type kiss_type_run_error_define
    MissCSV As Boolean
    MissMod As Boolean
    MissINI As Boolean
End Type
Public RunERR As kiss_type_run_error_define

Public DisplayMode As Long
Public TriggerBoard As Type_Trigger
Public TriggerCopied As Boolean
Public isShowTip As Boolean
