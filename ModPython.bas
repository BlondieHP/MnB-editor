Attribute VB_Name = "ModPython"
'Option Explicit
'Public Const GENERAL = 0
'Public Const ITEM_TRIGGER = 1
'Public Const MISSION_TEMPLATE_TRIGGER = 2
'Public Const SIMPLE_TRIGGER = 3
'Public Const SCENE_PROPS_TRIGGER = 4
'Public Const PRESENTATION_TRIGGER = 5
'Public Const MAPICON_TRIGGER = 6

Public Const skl_trade = 0
Public Const skl_leadership = 1
Public Const skl_prisoner_management = 2
Public Const skl_reserved_1 = 3
Public Const skl_reserved_2 = 4
Public Const skl_reserved_3 = 5
Public Const skl_reserved_4 = 6
Public Const skl_persuasion = 7
Public Const skl_engineer = 8
Public Const skl_first_aid = 9
Public Const skl_surgery = 10
Public Const skl_wound_treatment = 11
Public Const skl_inventory_management = 12
Public Const skl_spotting = 13
Public Const skl_pathfinding = 14
Public Const skl_tactics = 15
Public Const skl_tracking = 16
Public Const skl_trainer = 17
Public Const skl_reserved_5 = 18
Public Const skl_reserved_6 = 19
Public Const skl_reserved_7 = 20
Public Const skl_reserved_8 = 21
Public Const skl_looting = 22
Public Const skl_horse_archery = 23
Public Const skl_riding = 24
Public Const skl_athletics = 25
Public Const skl_shield = 26
Public Const skl_weapon_master = 27
Public Const skl_reserved_9 = 28
Public Const skl_reserved_10 = 29
Public Const skl_reserved_11 = 30
Public Const skl_reserved_12 = 31
Public Const skl_reserved_13 = 32
Public Const skl_power_draw = 33
Public Const skl_power_throw = 34
Public Const skl_power_strike = 35
Public Const skl_ironflesh = 36
Public Const skl_reserved_14 = 37
Public Const skl_reserved_15 = 38
Public Const skl_reserved_16 = 39
Public Const skl_reserved_17 = 40
Public Const skl_reserved_18 = 41

Public Type Type_Flag
    strName As String
    csvName As String
    Value As Integer64b
End Type

Public Type Type_Trigger_Function
    FunctionName As String
    Description As String
    Opblock1() As Type_Op_Block
    OpBlock2() As Type_Op_Block
End Type

Public Type Type_tiOn
  Value As Double
  csvName As String
  dbName As String
  Occation As String
  Tip As String
End Type

Public Itcf(66) As Type_strXYZ
Public Itc(18) As Type_strXY
Public Itp(31) As Type_strXY
Public Tf(19) As Type_Flag
Public Pf(23) As Type_Flag
Public AI_Bhvr(11) As Type_strXY
Public Item_Type(1 To 20) As Type_strXY
Public IModC(16) As Type_strXY
Public tiOn(2) As Type_strXYZ
Public tiOn_General(0) As Type_strXYZ
Public tiAct(16) As Type_strXY
Public PSf(9) As Type_strXYZ
Public MeshFlag(0) As Type_strXYZ
Public Negations(1)  As Type_strXY
Public TrgFunc(14) As Type_Trigger_Function
Public tiOns(4) As Type_tiOn
Public BoolSwitch(1) As Type_strXY
Public PlayOption(2) As Type_strXY
Public AccessPrivilege(1) As Type_strXY
Public AbsSwitch(1) As Type_strXY

Public Sub InitPy()
'逻辑操作符
InitNegations

'触发器功能
InitTrgFunc

'兵种
RegTfs

'部队
RegAI_Bhvr
RegPfs

'物品
RegItps
RegItemType
InitItcf
InitItc
InitIModCombines
InittiOn
InittiOn_General

'粒子系统
InitPSf

'网格模型
InitMeshFlags

'布尔值开关
initBoolSwitch

'播放选项
initPlayOption

'访问权限
initAccessPrivilege

'返回值选项
initAbsSwitch
End Sub

Private Sub InitNegations()

Negations(0).X = "80000000"
Negations(0).Y = "neg"

Negations(1).X = "40000000"
Negations(1).Y = "this_or_next"

End Sub

Private Sub initBoolSwitch()

BoolSwitch(0).X = "clear"
BoolSwitch(0).Y = "清除"

BoolSwitch(1).X = "set"
BoolSwitch(1).Y = "设置"

End Sub

Private Sub initAbsSwitch()

AbsSwitch(0).X = "relative"
AbsSwitch(0).Y = "相对数值(0-100)"

AbsSwitch(1).X = "absolute"
AbsSwitch(1).Y = "绝对数值"

End Sub

Private Sub initPlayOption()

PlayOption(0).X = "default"
PlayOption(0).Y = "默认"

PlayOption(1).X = "fade out current track"
PlayOption(1).Y = "淡出当前曲目"

PlayOption(2).X = "stop current track"
PlayOption(2).Y = "停止当前曲目"
End Sub

Private Sub initAccessPrivilege()
AccessPrivilege(0).X = "local"
AccessPrivilege(0).Y = "本地"

AccessPrivilege(1).X = "global"
AccessPrivilege(1).Y = "全局"
End Sub

Private Sub InitTrgFunc()

With TrgFunc(0)
  .FunctionName = "Unknown"
  .Description = "未知"
ReDim .Opblock1(0)
ReDim .OpBlock2(0)
End With

With TrgFunc(1)
  .FunctionName = "Tutorial"
  .Description = "教程"
ReDim .Opblock1(1 To 1)
ReDim .OpBlock2(1 To 1)
  With .Opblock1(1)
      .Op = CStr(map_free)
      .ParaNum = 0
       ReDim .Para(0)
  End With
  With .OpBlock2(1)
      .Op = CStr(dialog_box)
      .ParaNum = 1
       ReDim .Para(1 To .ParaNum)
      .Para(1).Value = "216172782113783828"    '待改进
  End With
End With

With TrgFunc(2)
  .FunctionName = "Refresh Merchants"
  .Description = "刷新杂货商人商品"
ReDim .Opblock1(0)
ReDim .OpBlock2(1 To 1)
  With .OpBlock2(1)
      .Op = CStr(troop_add_merchandise)
      .ParaNum = 3
      ReDim .Para(1 To .ParaNum)
      .Para(2).Value = CStr(itp_type_goods)
  End With
End With

With TrgFunc(3)
  .FunctionName = "Refresh Armor sellers"
  .Description = "刷新防具商人商品"
ReDim .Opblock1(0)
ReDim .OpBlock2(1 To 1)
  With .OpBlock2(1)
      .Op = CStr(troop_add_merchandise_with_faction)
      .ParaNum = 4
      ReDim .Para(1 To .ParaNum)
      .Para(3).Value = CStr(itp_type_body_armor) & "#" & CStr(itp_type_head_armor) & "#" & CStr(itp_type_foot_armor) & "#" & CStr(itp_type_hand_armor)
  End With
End With

With TrgFunc(4)
  .FunctionName = "Refresh Weapon sellers"
  .Description = "刷新武器商人商品"
ReDim .Opblock1(0)
ReDim .OpBlock2(1 To 1)
  With .OpBlock2(1)
      .Op = CStr(troop_add_merchandise_with_faction)
      .ParaNum = 4
      ReDim .Para(1 To .ParaNum)
      .Para(3).Value = CStr(itp_type_one_handed_wpn) & "#" & CStr(itp_type_two_handed_wpn) & "#" & CStr(itp_type_polearm) & "#" & CStr(itp_type_shield) & "#" & CStr(itp_type_bow) & "#" & CStr(itp_type_crossbow) & "#" & CStr(itp_type_thrown) & "#" & CStr(itp_type_arrows) & "#" & CStr(itp_type_bolts) & "#" & CStr(itp_type_pistol) & "#" & CStr(itp_type_musket) & "#" & CStr(itp_type_bullets)
  End With
End With

With TrgFunc(5)
  .FunctionName = "Refresh Horse sellers"
  .Description = "刷新马匹商人商品"
ReDim .Opblock1(0)
ReDim .OpBlock2(1 To 1)
  With .OpBlock2(1)
      .Op = CStr(troop_add_merchandise_with_faction)
      .ParaNum = 4
      ReDim .Para(1 To .ParaNum)
      .Para(3).Value = CStr(itp_type_horse)
  End With
End With

With TrgFunc(6)
  .FunctionName = "Respawn Parties"
  .Description = "重生部队"
ReDim .Opblock1(0)
ReDim .OpBlock2(1 To 1)
  With .OpBlock2(1)
      .Op = CStr(spawn_around_party)
      .ParaNum = 2
      ReDim .Para(1 To .ParaNum)
  End With
End With

With TrgFunc(7)
  .FunctionName = "Quest Bandits Trackdown Trigger"
  .Description = "任务:追捕强盗 相关触发器"
ReDim .Opblock1(1 To 1)
ReDim .OpBlock2(0)
  With .Opblock1(1)
      .Op = CStr(check_quest_active)
      .ParaNum = 1
      ReDim .Para(1 To .ParaNum)
      .Para(1).Value = "504403158265495605"
  End With
End With

With TrgFunc(8)
  .FunctionName = "Quest Incriminate Loyal Advisor Trigger"
  .Description = "任务:控告指挥官 相关触发器"
ReDim .Opblock1(1 To 1)
ReDim .OpBlock2(0)
  With .Opblock1(1)
      .Op = CStr(check_quest_active)
      .ParaNum = 1
      ReDim .Para(1 To .ParaNum)
      .Para(1).Value = "504403158265495565"
  End With
End With

With TrgFunc(9)
  .FunctionName = "Quest Runaway Peasants Trigger"
  .Description = "任务:追回逃跑农奴 相关触发器"
ReDim .Opblock1(1 To 1)
ReDim .OpBlock2(0)
  With .Opblock1(1)
      .Op = CStr(check_quest_active)
      .ParaNum = 1
      ReDim .Para(1 To .ParaNum)
      .Para(1).Value = "504403158265495560"
  End With
End With

With TrgFunc(10)
  .FunctionName = "Quest Follow Spy Trigger"
  .Description = "任务:跟踪间谍 相关触发器"
ReDim .Opblock1(1 To 1)
ReDim .OpBlock2(0)
  With .Opblock1(1)
      .Op = CStr(check_quest_active)
      .ParaNum = 1
      ReDim .Para(1 To .ParaNum)
      .Para(1).Value = "504403158265495561"
  End With
End With

With TrgFunc(11)
  .FunctionName = "Apply interest to merchants guild debt Trigger"
  .Description = "增加商会债务利息触发器"
ReDim .Opblock1(0)
ReDim .OpBlock2(1 To 1)
  With .OpBlock2(1)
      .Op = CStr(val_mul)
      .ParaNum = 2
      ReDim .Para(1 To .ParaNum)
      .Para(1).Value = "144115188075855881"
  End With
End With

With TrgFunc(12)
  .FunctionName = "Apply interest to merchants guild debt Trigger"
  .Description = "增加商会债务利息触发器"
ReDim .Opblock1(0)
ReDim .OpBlock2(1 To 1)
  With .OpBlock2(1)
      .Op = CStr(val_div)
      .ParaNum = 2
      ReDim .Para(1 To .ParaNum)
      .Para(1).Value = "144115188075855881"
  End With
End With

With TrgFunc(13)
  .FunctionName = "Quest Escort merchant caravan Trigger"
  .Description = "任务:护送商队 相关触发器"
ReDim .Opblock1(1 To 1)
ReDim .OpBlock2(0)
  With .Opblock1(1)
      .Op = CStr(check_quest_active)
      .ParaNum = 1
      ReDim .Para(1 To .ParaNum)
      .Para(1).Value = "504403158265495581"
  End With
End With

With TrgFunc(14)
  .FunctionName = "Quest involved Trigger"
  .Description = "任务相关触发器"
ReDim .Opblock1(1 To 1)
ReDim .OpBlock2(0)
  With .Opblock1(1)
      .Op = CStr(check_quest_active)
      .ParaNum = 1
      ReDim .Para(1 To .ParaNum)
  End With
End With

End Sub

Private Sub RegTfs()
Tf(0).strName = "tf_male"
Tf(0).Value = HexStrToI64(tf_male)
Tf(0).csvName = "男性"

Tf(1).strName = "tf_female"
Tf(1).Value = HexStrToI64(tf_female)
Tf(1).csvName = "女性"

Tf(2).strName = "tf_undead"
Tf(2).Value = HexStrToI64(tf_undead)
Tf(2).csvName = "僵尸"

Tf(3).strName = "troop_type_mask"
Tf(3).Value = HexStrToI64(troop_type_mask)
Tf(3).csvName = "兵种类型mask"

Tf(4).strName = "tf_hero"
Tf(4).Value = HexStrToI64(tf_hero)
Tf(4).csvName = "NPC英雄"

Tf(5).strName = "tf_inactive"
Tf(5).Value = HexStrToI64(tf_inactive)
Tf(5).csvName = "无生命的"

Tf(6).strName = "tf_unkillable"
Tf(6).Value = HexStrToI64(tf_unkillable)
Tf(6).csvName = "只能击晕"

Tf(7).strName = "tf_allways_fall_dead"
Tf(7).Value = HexStrToI64(tf_allways_fall_dead)
Tf(7).csvName = "倒下必死"

Tf(8).strName = "tf_no_capture_alive"
Tf(8).Value = HexStrToI64(tf_no_capture_alive)
Tf(8).csvName = "无法活捉"

Tf(9).strName = "tf_mounted"
Tf(9).Value = HexStrToI64(tf_mounted)
Tf(9).csvName = "骑马"

Tf(10).strName = "tf_is_merchant"
Tf(10).Value = HexStrToI64(tf_is_merchant)
Tf(10).csvName = "商队"

Tf(11).strName = "tf_randomize_face"
Tf(11).Value = HexStrToI64(tf_randomize_face)
Tf(11).csvName = "随机相貌"

Tf(12).strName = "tf_guarantee_boots"
Tf(12).Value = HexStrToI64(tf_guarantee_boots)
Tf(12).csvName = "保证穿鞋子"

Tf(13).strName = "tf_guarantee_armor"
Tf(13).Value = HexStrToI64(tf_guarantee_armor)
Tf(13).csvName = "保证穿盔甲"

Tf(14).strName = "tf_guarantee_helmet"
Tf(14).Value = HexStrToI64(tf_guarantee_helmet)
Tf(14).csvName = "保证戴头盔"

Tf(15).strName = "tf_guarantee_gloves"
Tf(15).Value = HexStrToI64(tf_guarantee_gloves)
Tf(15).csvName = "保证戴手套"

Tf(16).strName = "tf_guarantee_horse"
Tf(16).Value = HexStrToI64(tf_guarantee_horse)
Tf(16).csvName = "保证有马"

Tf(17).strName = "tf_guarantee_shield"
Tf(17).Value = HexStrToI64(tf_guarantee_shield)
Tf(17).csvName = "保证有盾"

Tf(18).strName = "tf_guarantee_ranged"
Tf(18).Value = HexStrToI64(tf_guarantee_ranged)
Tf(18).csvName = "保证有远程武器"

Tf(19).strName = "tf_unmoveable_in_party_window"
Tf(19).Value = HexStrToI64(tf_unmoveable_in_party_window)
Tf(19).csvName = "不能作为驻兵"
End Sub

Private Sub RegPfs()
Pf(0).strName = "pf_icon_mask"
Pf(0).Value = HexStrToI64(pf_icon_mask)
Pf(0).csvName = "图标mask"

Pf(1).strName = "pf_disabled"
Pf(1).Value = HexStrToI64(pf_disabled)
Pf(1).csvName = "不可用"

Pf(2).strName = "pf_is_ship"
Pf(2).Value = HexStrToI64(pf_is_ship)
Pf(2).csvName = "船"

Pf(3).strName = "pf_is_static"
Pf(3).Value = HexStrToI64(pf_is_static)
Pf(3).csvName = "静态"

Pf(4).strName = "pf_label_small"
Pf(4).Value = HexStrToI64(pf_label_small)
Pf(4).csvName = "小标签"

Pf(5).strName = "pf_label_medium"
Pf(5).Value = HexStrToI64(pf_label_medium)
Pf(5).csvName = "中标签"

Pf(6).strName = "pf_label_large"
Pf(6).Value = HexStrToI64(pf_label_large)
Pf(6).csvName = "大标签"

Pf(7).strName = "pf_label_mask"
Pf(7).Value = HexStrToI64(pf_label_mask)
Pf(7).csvName = "标签mask"

Pf(8).strName = "pf_always_visible"
Pf(8).Value = HexStrToI64(pf_always_visible)
Pf(8).csvName = "总是可见"

Pf(9).strName = "pf_default_behavior"
Pf(9).Value = HexStrToI64(pf_default_behavior)
Pf(9).csvName = "行为默认"

Pf(10).strName = "pf_auto_remove_in_town"
Pf(10).Value = HexStrToI64(pf_auto_remove_in_town)
Pf(10).csvName = "在城镇中自动去除"

Pf(11).strName = "pf_quest_party"
Pf(11).Value = HexStrToI64(pf_quest_party)
Pf(11).csvName = "任务特设"

Pf(12).strName = "pf_no_label"
Pf(12).Value = HexStrToI64(pf_no_label)
Pf(12).csvName = "无标签"

Pf(13).strName = "pf_limit_members"
Pf(13).Value = HexStrToI64(pf_limit_members)
Pf(13).csvName = "人数有限"

Pf(14).strName = "pf_hide_defenders"
Pf(14).Value = HexStrToI64(pf_hide_defenders)
Pf(14).csvName = "藏身处"

Pf(15).strName = "pf_show_faction"
Pf(15).Value = HexStrToI64(pf_show_faction)
Pf(15).csvName = "显示阵营"

Pf(16).strName = "pf_is_hidden"
Pf(16).Value = HexStrToI64(pf_is_hidden)
Pf(16).csvName = "不可见"

Pf(17).strName = "pf_dont_attack_civilians"
Pf(17).Value = HexStrToI64(pf_dont_attack_civilians)
Pf(17).csvName = "不攻击平民"

Pf(18).strName = "pf_civilian"
Pf(18).Value = HexStrToI64(pf_civilian)
Pf(18).csvName = "平民"

Pf(19).strName = "pf_carry_goods_bits"
Pf(19).Value = HexStrToI64(pf_carry_goods_bits)
Pf(19).csvName = "货物运输位"

Pf(20).strName = "pf_carry_gold_bits"
Pf(20).Value = HexStrToI64(pf_carry_gold_bits)
Pf(20).csvName = "金钱运输位"

Pf(21).strName = "pf_carry_gold_multiplier"
Pf(21).Value = HexStrToI64(pf_carry_gold_multiplier)
Pf(21).csvName = "金钱运输乘数"

Pf(22).strName = "pf_carry_goods_mask"
Pf(22).Value = HexStrToI64(pf_carry_goods_mask)
Pf(22).csvName = "货物运输mask"

Pf(23).strName = "pf_carry_gold_mask"
Pf(23).Value = HexStrToI64(pf_carry_gold_mask)
Pf(23).csvName = "金钱运输mask"

End Sub

Private Sub RegAI_Bhvr()
AI_Bhvr(0).X = "ai_bhvr_hold"
AI_Bhvr(0).Y = "坚守"
AI_Bhvr(1).X = "ai_bhvr_travel_to_party"
AI_Bhvr(1).Y = "旅行到部队"
AI_Bhvr(2).X = "ai_bhvr_patrol_location"
AI_Bhvr(2).Y = "在某地巡逻"
AI_Bhvr(3).X = "ai_bhvr_patrol_party"
AI_Bhvr(3).Y = "在某部队处巡逻"
AI_Bhvr(4).X = "ai_bhvr_attack_party"
AI_Bhvr(4).Y = "遇到部队攻击"
AI_Bhvr(5).X = "ai_bhvr_avoid_party"
AI_Bhvr(5).Y = "遇到部队躲避"
AI_Bhvr(6).X = "ai_bhvr_travel_to_point"
AI_Bhvr(6).Y = "旅行到某地"
AI_Bhvr(7).X = "ai_bhvr_negotiate_party"
AI_Bhvr(7).Y = "遇到部队讲和"
AI_Bhvr(8).X = "ai_bhvr_in_town"
AI_Bhvr(8).Y = "在城镇"
AI_Bhvr(9).X = "ai_bhvr_travel_to_ship"
AI_Bhvr(9).Y = "旅行到船"
AI_Bhvr(10).X = "ai_bhvr_escort_party"
AI_Bhvr(10).Y = "护送部队"
AI_Bhvr(11).X = "ai_bhvr_driven_by_party"
AI_Bhvr(11).Y = "被部队驱动"
End Sub

Private Sub RegItps()
'12+
Itp(0).X = "itp_unique"
Itp(1).X = "itp_always_loot"
Itp(2).X = "itp_no_parry"
Itp(3).X = "itp_default_ammo"
Itp(4).X = "itp_merchandise"
Itp(5).X = "itp_wooden_attack"
Itp(6).X = "itp_wooden_parry"
Itp(7).X = "itp_food"
Itp(0).Y = "独一无二"
Itp(1).Y = "总是战利品"
Itp(2).Y = "不能格挡"
Itp(3).Y = "默认弹药"
Itp(4).Y = "商品"
Itp(5).Y = "木质攻击"
Itp(6).Y = "木质格挡"
Itp(7).Y = "食物"

Itp(8).X = "itp_cant_reload_on_horseback"
Itp(9).X = "itp_two_handed"
Itp(10).X = "itp_primary"
Itp(11).X = "itp_secondary"
Itp(12).X = "itp_covers_legs/itp_doesnt_cover_hair/itp_can_penetrate_shield"
Itp(13).X = "itp_consumable"
Itp(14).X = "itp_bonus_against_shield"
Itp(15).X = "itp_penalty_with_shield"
Itp(8).Y = "不能在马上装填"
Itp(9).Y = "不能持盾"
Itp(10).Y = "主要武器"
Itp(11).Y = "次要武器"
Itp(12).Y = "遮盖脚部/不覆盖头发/可以穿透盾牌"
Itp(13).Y = "消耗品"
Itp(14).Y = "对盾加成"
Itp(15).Y = "持盾时伤害减少"

Itp(16).X = "itp_cant_use_on_horseback"
Itp(17).X = "itp_civilian/itp_next_item_as_melee"
Itp(18).X = "itp_fit_to_head/itp_offset_lance"
Itp(19).X = "itp_covers_head/itp_couchable"
Itp(20).X = "itp_crush_through"
Itp(21).X = "itp_knock_back"
Itp(22).X = "itp_remove_item_on_use"
Itp(23).X = "itp_unbalanced"
Itp(16).Y = "不能在马上使用"
Itp(17).Y = "民用/下个物品作为第二攻击模式"
Itp(18).Y = "吻合头部/冲刺时握持武器位置偏后"
Itp(19).Y = "遮盖头部/可以发动长枪冲刺"
Itp(20).Y = "破格挡"
Itp(21).Y = "Knock_Back(不可用)"
Itp(22).Y = "卸下正在使用的物品"
Itp(23).Y = "非平衡武器"

Itp(24).X = "itp_covers_beard"
Itp(25).X = "itp_no_pick_up_from_ground"
Itp(26).X = "itp_can_knock_down"
Itp(24).Y = "遮盖面部"
Itp(25).Y = "不能拾起"
Itp(26).Y = "可以击倒敌人"

'17+
Itp(27).X = "itp_extra_penetration"
Itp(28).X = "itp_has_bayonet"
Itp(29).X = "itp_cant_reload_while_moving"
Itp(30).X = "itp_ignore_gravity"
Itp(31).X = "itp_ignore_friction"
Itp(27).Y = "更多穿透力"
Itp(28).Y = "有刺刀"
Itp(29).Y = "移动时不能装填"
Itp(30).Y = "忽视重力"
Itp(31).Y = "忽视阻力"

End Sub

Private Sub RegItemType()
Item_Type(1).X = "itp_type_horse"
Item_Type(2).X = "itp_type_one_handed_wpn"
Item_Type(3).X = "itp_type_two_handed_wpn"
Item_Type(4).X = "itp_type_polearm"
Item_Type(1).Y = "马匹"
Item_Type(2).Y = "单手武器"
Item_Type(3).Y = "双手武器"
Item_Type(4).Y = "长杆武器"

Item_Type(5).X = "itp_type_arrows"
Item_Type(6).X = "itp_type_bolts"
Item_Type(7).X = "itp_type_shield"
Item_Type(8).X = "itp_type_bow"
Item_Type(9).X = "itp_type_crossbow"
Item_Type(10).X = "itp_type_thrown"
Item_Type(5).Y = "箭"
Item_Type(6).Y = "弩箭"
Item_Type(7).Y = "盾牌"
Item_Type(8).Y = "弓"
Item_Type(9).Y = "弩"
Item_Type(10).Y = "投掷武器"

Item_Type(11).X = "itp_type_goods"
Item_Type(12).X = "itp_type_head_armor"
Item_Type(13).X = "itp_type_body_armor"
Item_Type(14).X = "itp_type_foot_armor"
Item_Type(15).X = "itp_type_hand_armor"
Item_Type(11).Y = "货物"
Item_Type(12).Y = "头盔"
Item_Type(13).Y = "铠甲"
Item_Type(14).Y = "鞋子"
Item_Type(15).Y = "手套"

Item_Type(16).X = "itp_type_pistol"
Item_Type(17).X = "itp_type_musket"
Item_Type(18).X = "itp_type_bullets"
Item_Type(19).X = "itp_type_animal"
Item_Type(20).X = "itp_type_book"
Item_Type(16).Y = "手枪"
Item_Type(17).Y = "步枪"
Item_Type(18).Y = "子弹"
Item_Type(19).Y = "动物"
Item_Type(20).Y = "书籍"

End Sub

Private Sub InitItcf()

Itcf(0).X = itcf_Thrust_onehanded
Itcf(0).Y = "单手武器直刺"
Itcf(0).Z = "itcf_thrust_onehanded"
Itcf(1).X = itcf_Overswing_onehanded
Itcf(1).Y = "单手武器竖劈"
Itcf(1).Z = "itcf_overswing_onehanded"
Itcf(2).X = itcf_Slashright_onehanded
Itcf(2).Y = "单手武器右挥"
Itcf(2).Z = "itcf_slashright_onehanded"
Itcf(3).X = itcf_Slashleft_onehanded
Itcf(3).Y = "单手武器左挥"
Itcf(3).Z = "itcf_slashleft_onehanded"

Itcf(4).X = itcf_Thrust_twohanded
Itcf(5).X = itcf_Overswing_twohanded
Itcf(6).X = itcf_Slashright_twohanded
Itcf(7).X = itcf_Slashleft_twohanded
Itcf(4).Y = "双手武器直刺"
Itcf(5).Y = "双手武器竖劈"
Itcf(6).Y = "双手武器右挥"
Itcf(7).Y = "双手武器左挥"
Itcf(4).Z = "itcf_thrust_twohanded"
Itcf(5).Z = "itcf_overswing_twohanded"
Itcf(6).Z = "itcf_slashright_twohanded"
Itcf(7).Z = "itcf_slashleft_twohanded"

Itcf(8).X = itcf_Thrust_polearm
Itcf(9).X = itcf_Overswing_polearm
Itcf(10).X = itcf_Slashright_polearm
Itcf(11).X = itcf_Slashleft_polearm
Itcf(8).Y = "长杆武器直刺"
Itcf(9).Y = "长杆武器竖劈"
Itcf(10).Y = "长杆武器右挥"
Itcf(11).Y = "长杆武器左挥"
Itcf(8).Z = "itcf_thrust_polearm"
Itcf(9).Z = "itcf_overswing_polearm"
Itcf(10).Z = "itcf_slashright_polearm"
Itcf(11).Z = "itcf_slashleft_polearm"

Itcf(12).X = itcf_Horseback_thrust_onehanded
Itcf(13).X = itcf_Horseback_overswing_right_onehanded
Itcf(14).X = itcf_Horseback_overswing_left_onehanded
Itcf(15).X = itcf_Horseback_slashright_onehanded
Itcf(16).X = itcf_Horseback_slashleft_onehanded
Itcf(17).X = itcf_Horseback_slash_polearm
Itcf(18).X = itcf_Thrust_onehanded_lance
Itcf(19).X = itcf_Thrust_onehanded_lance_horseback
Itcf(12).Y = "在马背上单手直刺"
Itcf(13).Y = "在马背上右边单手竖劈"
Itcf(14).Y = "在马背上左边单手竖劈"
Itcf(15).Y = "在马背上单手右挥"
Itcf(16).Y = "在马背上单手左挥"
Itcf(17).Y = "在马背上挥舞长杆"
Itcf(18).Y = "骑枪单手直刺"
Itcf(19).Y = "在马背上骑枪单手直刺"
Itcf(12).Z = "itcf_horseback_thrust_onehanded"
Itcf(13).Z = "itcf_horseback_overswing_right_onehanded"
Itcf(14).Z = "itcf_horseback_overswing_left_onehanded"
Itcf(15).Z = "itcf_horseback_slashright_onehanded"
Itcf(16).Z = "itcf_horseback_slashleft_onehanded"
Itcf(17).Z = "itcf_horseback_slash_polearm"
Itcf(18).Z = "itcf_thrust_onehanded_lance"
Itcf(19).Z = "itcf_thrust_onehanded_lance_horseback"

Itcf(20).X = itcf_Parry_forward_onehanded
Itcf(21).X = itcf_Parry_up_onehanded
Itcf(22).X = itcf_Parry_right_onehanded
Itcf(23).X = itcf_Parry_left_onehanded
Itcf(20).Y = "单手格挡直刺"
Itcf(21).Y = "单手格挡竖劈"
Itcf(22).Y = "单手右格挡"
Itcf(23).Y = "单手左格挡"
Itcf(20).Z = "itcf_parry_forward_onehanded"
Itcf(21).Z = "itcf_parry_up_onehanded"
Itcf(22).Z = "itcf_parry_right_onehanded"
Itcf(23).Z = "itcf_parry_left_onehanded"

Itcf(24).X = itcf_Parry_forward_twohanded
Itcf(25).X = itcf_Parry_up_twohanded
Itcf(26).X = itcf_Parry_right_twohanded
Itcf(27).X = itcf_Parry_left_twohanded
Itcf(24).Y = "双手格挡直刺"
Itcf(25).Y = "双手格挡竖劈"
Itcf(26).Y = "双手右格挡"
Itcf(27).Y = "双手左格挡"
Itcf(24).Z = "itcf_parry_forward_twohanded"
Itcf(25).Z = "itcf_parry_up_twohanded"
Itcf(26).Z = "itcf_parry_right_twohanded"
Itcf(27).Z = "itcf_parry_left_twohanded"

Itcf(28).X = itcf_Parry_forward_polearm
Itcf(29).X = itcf_Parry_up_polearm
Itcf(30).X = itcf_Parry_right_polearm
Itcf(31).X = itcf_Parry_left_polearm
Itcf(28).Y = "长杆格挡直刺"
Itcf(29).Y = "长杆格挡竖劈"
Itcf(30).Y = "长杆右格挡"
Itcf(31).Y = "长杆左格挡"
Itcf(28).Z = "itcf_parry_forward_polearm"
Itcf(29).Z = "itcf_parry_up_polearm"
Itcf(30).Z = "itcf_parry_right_polearm"
Itcf(31).Z = "itcf_parry_left_polearm"

Itcf(32).X = itcf_Show_holster_when_drawn
Itcf(32).Y = "显示剑鞘"
Itcf(32).Z = "itcf_Show_holster_when_drawn"

'shoot mask
Itcf(33).X = itcf_Shoot_bow
Itcf(34).X = itcf_Shoot_javelin
Itcf(35).X = itcf_Shoot_crossbow
Itcf(33).Y = "弓射击"
Itcf(34).Y = "标枪射击"
Itcf(35).Y = "弩射击"
Itcf(33).Z = "itcf_shoot_bow"
Itcf(34).Z = "itcf_shoot_javelin"
Itcf(35).Z = "itcf_shoot_crossbow"

Itcf(36).X = itcf_Throw_stone
Itcf(37).X = itcf_Throw_knife
Itcf(38).X = itcf_Throw_axe
Itcf(39).X = itcf_Throw_javelin
Itcf(40).X = itcf_Shoot_pistol
Itcf(41).X = itcf_Shoot_musket
Itcf(36).Y = "投掷石子"
Itcf(37).Y = "投掷飞刀"
Itcf(38).Y = "投掷飞斧"
Itcf(39).Y = "投掷标枪"
Itcf(40).Y = "手枪射击"
Itcf(41).Y = "步枪射击"
Itcf(36).Z = "itcf_throw_stone"
Itcf(37).Z = "itcf_throw_knife"
Itcf(38).Z = "itcf_throw_axe"
Itcf(39).Z = "itcf_throw_javelin"
Itcf(40).Z = "itcf_shoot_pistol"
Itcf(41).Z = "itcf_shoot_musket"
'shoot end

'carry mask
Itcf(42).X = itcf_Carry_sword_left_hip
Itcf(43).X = itcf_Carry_sword_back
Itcf(42).Y = "将剑背在左臀部"
Itcf(43).Y = "将剑背在背后"
Itcf(42).Z = "itcf_carry_sword_left_hip"
Itcf(43).Z = "itcf_carry_sword_back"

Itcf(44).X = itcf_Carry_axe_left_hip
Itcf(45).X = itcf_Carry_axe_back
Itcf(44).Y = "将斧头背在左臀部"
Itcf(45).Y = "将斧头背在背后"
Itcf(44).Z = "itcf_carry_axe_left_hip"
Itcf(45).Z = "itcf_carry_axe_back"

Itcf(46).X = itcf_Carry_spear
Itcf(46).Y = "携带长矛"
Itcf(46).Z = "itcf_carry_spear"

Itcf(47).X = itcf_Carry_dagger_front_left
Itcf(48).X = itcf_Carry_dagger_front_right
Itcf(47).Y = "将飞刀戴在左前"
Itcf(48).Y = "将飞刀戴在右前"
Itcf(47).Z = "itcf_carry_dagger_front_left"
Itcf(48).Z = "itcf_carry_dagger_front_right"

Itcf(49).X = itcf_Carry_quiver_front_right
Itcf(50).X = itcf_Carry_quiver_back_right
Itcf(51).X = itcf_Carry_quiver_right_vertical
Itcf(52).X = itcf_Carry_quiver_back
Itcf(49).Y = "将箭筒戴在右前"
Itcf(50).Y = "将箭筒戴在右后"
Itcf(51).Y = "将箭筒垂直戴在右边"
Itcf(52).Y = "将箭筒戴在背后"
Itcf(49).Z = "itcf_carry_quiver_front_right"
Itcf(50).Z = "itcf_carry_quiver_back_right"
Itcf(51).Z = "itcf_carry_quiver_right_vertical"
Itcf(52).Z = "itcf_carry_quiver_back"

Itcf(53).X = itcf_Carry_revolver_right
Itcf(54).X = itcf_Carry_pistol_front_left
Itcf(55).X = itcf_Carry_bowcase_left
Itcf(56).X = itcf_Carry_mace_left_hip
Itcf(53).Y = "将左轮手枪戴在右边"
Itcf(54).Y = "将手枪戴在左前"
Itcf(55).Y = "将弓袋戴在左边"
Itcf(56).Y = "将锤头戴在左臀部"
Itcf(53).Z = "itcf_carry_revolver_right"
Itcf(54).Z = "itcf_carry_pistol_front_left"
Itcf(55).Z = "itcf_carry_bowcase_left"
Itcf(56).Z = "itcf_carry_mace_left_hip"

Itcf(57).X = itcf_Carry_kite_shield
Itcf(58).X = itcf_Carry_round_shield
Itcf(59).X = itcf_Carry_buckler_left
Itcf(60).X = itcf_Carry_board_shield
Itcf(57).Y = "携带筝形盾"
Itcf(58).Y = "携带圆盾"
Itcf(59).Y = "将小圆盾戴在左边"
Itcf(60).Y = "携带阔盾"
Itcf(57).Z = "itcf_carry_kite_shield"
Itcf(58).Z = "itcf_carry_round_shield"
Itcf(59).Z = "itcf_carry_buckler_left"
Itcf(60).Z = "itcf_carry_board_shield"

Itcf(61).X = itcf_Carry_crossbow_back
Itcf(62).X = itcf_Carry_bow_back
Itcf(61).Y = "将弩背在背后"
Itcf(62).Y = "将弓背在背后"
Itcf(61).Z = "itcf_carry_crossbow_back"
Itcf(62).Z = "itcf_carry_bow_back"

Itcf(63).X = itcf_Carry_katana
Itcf(64).X = itcf_Carry_wakizashi
Itcf(63).Y = "携带太刀"
Itcf(64).Y = "携带胁指"
Itcf(63).Z = "itcf_carry_katana"
Itcf(64).Z = "itcf_carry_wakizashi"
'carry end

'reload mask
Itcf(65).X = itcf_Reload_pistol
Itcf(66).X = itcf_Reload_musket
Itcf(65).Y = "手枪装填"
Itcf(66).Y = "步枪装填"
Itcf(65).Z = "itcf_reload_pistol"
Itcf(66).Z = "itcf_reload_musket"
'reload end

End Sub

Private Sub InittiOn()

tiOn(0).X = "-50"
tiOn(0).Y = "初始化物品"
tiOn(0).Z = "ti_on_init_item"

tiOn(1).X = "-51"
tiOn(1).Y = "武器攻击"
tiOn(1).Z = "ti_on_weapon_attack"

tiOn(2).X = "-52"
tiOn(2).Y = "箭矢击中"
tiOn(2).Z = "ti_on_missile_hit"

With tiOns(0)
  .Value = 0
  .csvName = "(无)"
  .dbName = "0"
  .Occation = "gnl"
  .Tip = ""
End With

With tiOns(1)
  .Value = 100000000
  .csvName = "激发一次"
  .dbName = "ti_once"
  .Occation = "gnl"
  .Tip = ""
End With

With tiOns(2)
  .Value = -50
  .csvName = "初始化物品"
  .dbName = "ti_on_init_item"
  .Occation = "itm"
  .Tip = "[触发器参数1:当前角色, 触发器参数2:角色兵种]"
End With

With tiOns(3)
  .Value = -51
  .csvName = "武器攻击"
  .dbName = "ti_on_weapon_attack"
  .Occation = "itm"
  .Tip = "[触发器参数1:攻击方角色, 位置1:武器位置]"
End With

With tiOns(4)
  .Value = -52
  .csvName = "箭矢击中"
  .dbName = "ti_on_missile_hit"
  .Occation = "itm"
  .Tip = "[触发器参数1:射手角色, 位置1:箭矢位置]"
End With

End Sub

Private Sub InittiOn_General()

tiOn_General(0).X = "100000000"
tiOn_General(0).Y = "ti_once"
tiOn_General(0).Z = "ti_once"

End Sub
Private Sub InitItc()
Itc(0).X = itc_cleaver
Itc(1).X = itc_dagger
Itc(0).Y = "屠宰刀"
Itc(1).Y = "匕首=屠宰刀+单手直刺"

Itc(2).X = itc_parry_onehanded
Itc(3).X = itc_longsword
Itc(4).X = itc_scimitar
Itc(2).Y = "单手格挡"
Itc(3).Y = "长剑(包含匕首)"
Itc(4).Y = "弯刀(包含屠宰刀)"

Itc(5).X = itc_parry_two_handed
Itc(6).X = itc_cut_two_handed
Itc(5).Y = "双手格挡"
Itc(6).Y = "双手挥砍"

Itc(7).X = itc_greatsword
Itc(8).X = itc_nodachi
Itc(7).Y = "巨剑(包含双手挥砍+双手格挡)"
Itc(8).Y = "太刀(包含双手挥砍+双手格挡)"

Itc(9).X = itc_bastardsword
Itc(10).X = itc_morningstar
Itc(9).Y = "手半剑(包含双手挥砍+双手格挡+匕首)"
Itc(10).Y = "晨星(包含双手挥砍+双手格挡+屠宰刀)"

Itc(11).X = itc_parry_polearm
Itc(12).X = itc_poleaxe
Itc(13).X = itc_staff
Itc(14).X = itc_spear
Itc(15).X = itc_cutting_spear
Itc(16).X = itc_pike
Itc(17).X = itc_guandao
Itc(11).Y = "长杆格挡"
Itc(12).Y = "长斧(包含长杆格挡)"
Itc(13).Y = "棍子(包含长杆格挡)"
Itc(14).Y = "长枪(包含长杆格挡)"
Itc(15).Y = "长枪挥舞(包含长枪)"
Itc(16).Y = "长矛"
Itc(17).Y = "关刀(包含长杆格挡)"

Itc(18).X = itc_greatlance
Itc(18).Y = "巨型骑枪"
End Sub

Private Sub InitIModCombines()
 IModC(0).X = "0"
 IModC(1).X = imodbits_horse_basic
 IModC(2).X = imodbits_horse_good
 IModC(3).X = imodbits_cloth
 IModC(4).X = imodbits_armor
 IModC(5).X = imodbits_plate
 IModC(6).X = imodbits_shield
 IModC(7).X = imodbits_polearm
 IModC(8).X = imodbits_axe
 IModC(9).X = imodbits_sword
 IModC(10).X = imodbits_sword_high
 IModC(11).X = imodbits_pick
 IModC(12).X = imodbits_bow
 IModC(13).X = imodbits_crossbow
 IModC(14).X = imodbits_missile
 IModC(15).X = imodbits_thrown_minus_heavy
 IModC(16).X = imodbits_thrown
 
  IModC(0).Y = "imodbits_none"
 IModC(1).Y = "imodbits_horse_basic"
 IModC(2).Y = "imodbits_horse_good"
 IModC(3).Y = "imodbits_cloth"
 IModC(4).Y = "imodbits_armor"
 IModC(5).Y = "imodbits_plate"
 IModC(6).Y = "imodbits_shield"
 IModC(7).Y = "imodbits_polearm"
 IModC(8).Y = "imodbits_axe"
 IModC(9).Y = "imodbits_sword"
 IModC(10).Y = "imodbits_sword_high"
 IModC(11).Y = "imodbits_pick"
 IModC(12).Y = "imodbits_bow"
 IModC(13).Y = "imodbits_crossbow"
 IModC(14).Y = "imodbits_missile"
 IModC(15).Y = "imodbits_thrown_minus_heavy"
 IModC(16).Y = "imodbits_thrown"

End Sub

Public Sub InitPSf()

PSf(0).X = psf_global_emit_dir
PSf(0).Y = "psf_always_emit"
PSf(0).Z = "总是存在"

PSf(1).X = psf_global_emit_dir
PSf(1).Y = "psf_global_emit_dir"
PSf(1).Z = "全场景存在"

PSf(2).X = psf_emit_at_water_level
PSf(2).Y = "psf_emit_at_water_level"
PSf(2).Z = "水面爆出"

'billboard  mask
PSf(3).X = psf_billboard_drop
PSf(3).Y = "psf_billboard_drop"
PSf(3).Z = "Billboard Drop"

PSf(4).X = psf_billboard_2d               ' up_vec = dir, front rotated towards camera
PSf(4).Y = "psf_billboard_2d"
PSf(4).Z = "2D平面效果"

PSf(5).X = psf_billboard_3d              '# front_vec point to camera.
PSf(5).Y = "psf_billboard_3d"
PSf(5).Z = "3D平面效果"
'________________________________

PSf(6).X = "0000000400"
PSf(6).Y = "psf_turn_to_velocity"
PSf(6).Z = "Turn to Velocity"

PSf(7).X = psf_randomize_rotation
PSf(7).Y = "psf_randomize_rotation"
PSf(7).Z = "随机旋转"

PSf(8).X = psf_randomize_size
PSf(8).Y = "psf_randomize_size"
PSf(8).Z = "随机尺寸"

PSf(9).X = psf_2d_turbulance
PSf(9).Y = "psf_2d_turbulance"
PSf(9).Z = "2D震荡"

End Sub
Public Sub InitMeshFlags()

MeshFlag(0).X = render_order_plus_1
MeshFlag(0).Y = "render_order_plus_1"
MeshFlag(0).Z = "Render_Order_Plus_1"

End Sub

'*************************************************************************
'**函 数 名：ExportTroopPYCode
'**输    入：(Type_Troops)trp
'**输    出：(String)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：
'**日    期：
'**修 改 人：Ser_Charles
'**日    期：2010-12-22 00:04:21
'**版    本：V1.1321
'*************************************************************************
Public Function ExportTroopPYCode(Trp As Type_Troops, Optional ifEnt As Boolean = True) As String
Dim sq As String, i As Integer, j As Integer, n As Byte, Ent As String, tI(1) As Integer64b, strTem As String, q As Boolean, Discrabe() As Variant
Dim strFace(1) As String

sq = Chr(34)
If ifEnt Then
    Ent = vbCrLf
Else
    Ent = ""
End If

With Trp
'ID,名称
ExportTroopPYCode = "[" & sq & Right(.strID, Len(.strID) - 4) & sq & "," & sq & .strName & sq & "," & sq & .strPtName & sq & ", "

'Flags
tI(0) = StrToI64(.Flags)
strTem = ""
For i = 1 To UBound(Tf)
       tI(1) = And64b(tI(0), Tf(i).Value)
       If IsEqual64b(tI(1), Tf(i).Value) Then
          strTem = strTem & Tf(i).strName & "|"
       End If
Next i

If Len(strTem) > 0 Then
   strTem = Left(strTem, Len(strTem) - 1)
Else
   strTem = Tf(0).strName
End If

ExportTroopPYCode = ExportTroopPYCode & strTem & ", "

'场景
If .SceneID > 0 Then
   strTem = .Scene_strID & "|entry(" & .Entry & "),reserved, "
Else
   strTem = "no_scene,reserved,"
End If

'阵营
strTem = strTem & .Faction_strID & ","

ExportTroopPYCode = ExportTroopPYCode & strTem

'物品
strTem = ""
For i = 1 To 64
    If .lstInventory(i).X > -1 Then
       strTem = strTem & .lstInventory(i).strX & ","
    End If
Next i

If Len(strTem) > 0 Then
   strTem = Left(strTem, Len(strTem) - 1)
End If

ExportTroopPYCode = ExportTroopPYCode & Ent & " [" & strTem & "]," & Ent

'属性
If .tAttrib.strPoint > 0 Then
   strTem = "str_" & .tAttrib.strPoint & "|"
End If

If .tAttrib.agiPoint > 0 Then
   strTem = strTem & "agi_" & .tAttrib.agiPoint & "|"
End If

If .tAttrib.intPoint > 0 Then
   strTem = strTem & "int_" & .tAttrib.intPoint & "|"
End If

If .tAttrib.chaPoint > 0 Then
   strTem = strTem & "cha_" & .tAttrib.chaPoint & "|"
End If

strTem = strTem & "level(" & .tAttrib.level & "),"
ExportTroopPYCode = ExportTroopPYCode & strTem

'熟练度
strTem = ExportWeaponProficiencies(.WP)
ExportTroopPYCode = ExportTroopPYCode & strTem & ","

'技能
strTem = ""
Discrabe = Array("trade", "leadership", "prisoner_management", "reserved_1", "reserved_2", _
               "reserved_3", "reserved_4", "persuasion", "engineer", "first_aid", _
               "surgery", "wound_treatment", "inventory_management", "spotting", "pathfinding", _
               "tactics", "tracking", "trainer", "reserved_5", "reserved_6", _
               "reserved_7", "reserved_8", "looting", "horse_archery", "riding", _
               "athletics", "shield", "weapon_master", "reserved_9", "reserved_10", _
               "reserved_11", "reserved_12", "reserved_13", "power_draw", "power_throw", _
               "power_strike", "ironflesh", "reserved_14", "reserved_15", "reserved_16", _
               "reserved_17", "reserved_18")

For i = 0 To UBound(Discrabe)
    n = GetSkill(i)
    If n > 0 Then
       strTem = strTem & "knows_" & Discrabe(i) & "_" & n & "|"
    End If
Next i

If Len(strTem) > 1 Then
   strTem = Left(strTem, Len(strTem) - 1)
Else
   strTem = "0"
End If

ExportTroopPYCode = ExportTroopPYCode & strTem & ","

'脸部代码
strFace(0) = "0x"
strFace(1) = "0x"
For i = 1 To 4
 strFace(0) = strFace(0) & LCase$(I64toHexStr(StrToI64(.Face(i))))
 strFace(1) = strFace(1) & LCase$(I64toHexStr(StrToI64(.Face(i + 4))))
Next i
          
If strFace(1) <> "0x0000000000000000000000000000000000000000000000000000000000000000" Then
  ExportTroopPYCode = ExportTroopPYCode & strFace(0) & "," & strFace(1) & "],"
Else
  ExportTroopPYCode = ExportTroopPYCode & strFace(0) & "],"
End If

End With
End Function
'*************************************************************************
'**函 数 名：GettiOnIndex
'**输    入：-(Double)ti_On
'**输    出：-
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-6-7 15:38:12
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Function GettiOnIndex(ti_On As Double, Optional Failed As Boolean) As Integer
Dim i As Integer

Failed = True
GettiOnIndex = -1
For i = 0 To UBound(tiOns)
    If CDbl(tiOns(i).Value) = ti_On Then
       GettiOnIndex = i
       Failed = False
       Exit For
    End If
Next i

End Function
'*************************************************************************
'**函 数 名：ExportItemPYCode
'**输    入：(Type_Item)itm
'**输    出：(String)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：
'**日    期：
'**修 改 人：Ser_Charles
'**日    期：2010-12-22 00:04:21
'**版    本：V1.1321
'*************************************************************************
Public Function ExportItemPYCode(itm As Type_Item, Optional ifEnt As Boolean = True) As String
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
Dim sq As String, i As Integer, j As Integer, n As Byte, Ent As String, temStr As String
Dim ErrStr As String

sq = Chr(34)
If ifEnt Then
    Ent = vbCrLf
Else
    Ent = ""
End If
'dbname,disname
'[("dbname","disname",[
ExportItemPYCode = "[" & sq & Right(itm.dbName, Len(itm.dbName) - 4) & sq & ", " & sq & itm.disname & sq & ", ["

'mdlname
Dim tI As Integer64b, tmdl_b As String, tImod As String, tI2 As Integer64b

For i = 1 To itm.nmdl
     tI = StrToI64(itm.mdl_b(i))
     'ixmesh
     If ChkBit64b(tI, ixmesh_Inventory_bit) And ChkBit64b(tI, ixmesh_Flying_Ammo_bit) Then
         tmdl_b = "ixmesh_carry"
     ElseIf ChkBit64b(tI, ixmesh_Inventory_bit) Then
         tmdl_b = "ixmesh_inventory"
     ElseIf ChkBit64b(tI, ixmesh_Flying_Ammo_bit) Then
         tmdl_b = "ixmesh_flying_ammo"
     Else
         tmdl_b = "0"
     End If
     
     'mimod
     For n = 0 To N_IMod - 1
         If ChkBit64b(tI, n) Then
             tImod = IMod(n).ID
             Exit For
         End If
     Next n
     If tImod <> "" Then tImod = "|" & tImod
     ExportItemPYCode = ExportItemPYCode & "(" & sq & itm.mdlname(i) & sq & ", " & tmdl_b & tImod & ")"
     
     If i < itm.nmdl Then
     ExportItemPYCode = ExportItemPYCode & ", "
     End If
Next i
ExportItemPYCode = ExportItemPYCode & "], "

'itp
   'Type
   ExportItemPYCode = ExportItemPYCode & Item_Type(GetItmType(itm.itmType)).X & "|"
   'attachments
   Dim Att As Integer
   Att = GetAttachment(itm.itmType)
   Select Case Att
       Case 1
            ExportItemPYCode = ExportItemPYCode & "itp_force_attach_left_hand|"
       Case 2
            ExportItemPYCode = ExportItemPYCode & "itp_force_attach_right_hand|"
       Case 3
            ExportItemPYCode = ExportItemPYCode & "itp_force_attach_left_forearm|"
       Case 4
            ExportItemPYCode = ExportItemPYCode & "itp_attach_armature|"
   End Select
   
   'itp
   tI = StrToI64(itm.itmType)
   
   
   For n = 0 To UBound(Itp)
      If n <= 26 Then
        If ChkBit64b(tI, n + 12) Then
          If n = 12 Then
             If IsWeapon(itm.itmType) Or IsAmmo(itm.itmType) Then
                ExportItemPYCode = ExportItemPYCode & GetFlagName(Itp(n).X, 3) & "|"
             ElseIf IsHeadArmor(itm.itmType) Then
                ExportItemPYCode = ExportItemPYCode & GetFlagName(Itp(n).X, 2) & "|"
             Else
                ExportItemPYCode = ExportItemPYCode & GetFlagName(Itp(n).X, 1) & "|"
             End If
          ElseIf n = 17 Then
             If IsWeapon(itm.itmType) Then
                ExportItemPYCode = ExportItemPYCode & GetFlagName(Itp(n).X, 2) & "|"
             Else
                ExportItemPYCode = ExportItemPYCode & GetFlagName(Itp(n).X, 1) & "|"
             End If
          ElseIf n = 18 Then
             If IsWeapon(itm.itmType) Then
                ExportItemPYCode = ExportItemPYCode & GetFlagName(Itp(n).X, 2) & "|"
             Else
                ExportItemPYCode = ExportItemPYCode & GetFlagName(Itp(n).X, 1) & "|"
             End If
          ElseIf n = 19 Then
             If IsWeapon(itm.itmType) Then
                ExportItemPYCode = ExportItemPYCode & GetFlagName(Itp(n).X, 2) & "|"
             Else
                ExportItemPYCode = ExportItemPYCode & GetFlagName(Itp(n).X, 1) & "|"
             End If
          Else
            ExportItemPYCode = ExportItemPYCode & Itp(n).X & "|"
          End If
        End If
      Else
         If ChkBit64b(tI, n + 17) Then
            ExportItemPYCode = ExportItemPYCode & Itp(n).X & "|"
         End If
      End If
   Next n
   
'末尾处理
   If Right(ExportItemPYCode, 1) = "|" Then
        ExportItemPYCode = Left(ExportItemPYCode, Len(ExportItemPYCode) - 1)
   End If
   
   If itm.itmType = "0" Then
        ExportItemPYCode = ExportItemPYCode & "0"
   End If
   
   ExportItemPYCode = ExportItemPYCode & "," & Ent & " "
   
'itcf
'no mask
tI = StrToI64(itm.Action)

For i = 0 To 32
    If IsZero64b(And64b(tI, HexStrToI64(Itcf(i).X))) = False Then
          ExportItemPYCode = ExportItemPYCode & Itcf(i).Z & "|"
    End If
Next i

'shoot
tI2 = And64b(tI, HexStrToI64(itcf_Shoot_mask))
For i = 32 To 41
    If I64ToBinStr(tI2) = HexToBin(Itcf(i).X) Then
          ExportItemPYCode = ExportItemPYCode & Itcf(i).Z & "|"
    End If
Next i

'carry
tI2 = And64b(tI, HexStrToI64(itcf_Carry_mask))
For i = 42 To 64
    If I64ToBinStr(tI2) = HexToBin(Itcf(i).X) Then
          ExportItemPYCode = ExportItemPYCode & Itcf(i).Z & "|"
    End If
Next i

'reload
tI2 = And64b(tI, HexStrToI64(itcf_Reload_mask))
For i = 64 To 66
    If I64ToBinStr(tI2) = HexToBin(Itcf(i).X) Then
          ExportItemPYCode = ExportItemPYCode & Itcf(i).Z & "|"
    End If
Next i

'末尾处理
   If Right(ExportItemPYCode, 1) = "|" Then
        ExportItemPYCode = Left(ExportItemPYCode, Len(ExportItemPYCode) - 1)
   End If
   
   If itm.Action = "0" Then
        ExportItemPYCode = ExportItemPYCode & "0"
   End If
   
   ExportItemPYCode = ExportItemPYCode & "," & Ent & " "

'价格
   ExportItemPYCode = ExportItemPYCode & itm.price & ", "

'属性
   Dim Dam1 As Long, Dam2 As Long, tpD1 As Integer, tpD2 As Integer
   If IsMeleeWeapon(itm.itmType) Then
        ExportItemPYCode = ExportItemPYCode & "weight(" & Round(Val(itm.weight), 2) & ")|"
        ExportItemPYCode = ExportItemPYCode & "weapon_length(" & itm.weapon_length & ")|"
        ExportItemPYCode = ExportItemPYCode & "difficulty(" & itm.difficulty & ")|"
        ExportItemPYCode = ExportItemPYCode & "spd_rtng(" & itm.speed_rating & ")|"
        If itm.abundance > 0 Then ExportItemPYCode = ExportItemPYCode & "abundance(" & itm.abundance & ")|"
        If itm.swing_damage > 0 Then
             Dam1 = GetDamage(itm.swing_damage, tpD1)
             ExportItemPYCode = ExportItemPYCode & "swing_damage(" & Dam1 & ", "
             If tpD1 = 0 Then
                ExportItemPYCode = ExportItemPYCode & "cut)|"
             ElseIf tpD1 = 1 Then
                ExportItemPYCode = ExportItemPYCode & "pierce)|"
             ElseIf tpD1 = 2 Then
                ExportItemPYCode = ExportItemPYCode & "blunt)|"
             End If
        End If
        If itm.thrust_damage > 0 Then
              Dam2 = GetDamage(itm.thrust_damage, tpD2)
             ExportItemPYCode = ExportItemPYCode & "thrust_damage(" & Dam2 & ", "
             If tpD2 = 0 Then
                ExportItemPYCode = ExportItemPYCode & "cut)|"
             ElseIf tpD2 = 1 Then
                ExportItemPYCode = ExportItemPYCode & "pierce)|"
             ElseIf tpD2 = 2 Then
                ExportItemPYCode = ExportItemPYCode & "blunt)|"
             End If
        End If
   ElseIf IsRangedWeapon(itm.itmType) Or IsFireArm(itm.itmType) Then
        ExportItemPYCode = ExportItemPYCode & "weight(" & Round(Val(itm.weight), 2) & ")|"
        If itm.weapon_length > 0 Then ExportItemPYCode = ExportItemPYCode & "weapon_length(" & itm.weapon_length & ")|"
        ExportItemPYCode = ExportItemPYCode & "difficulty(" & itm.difficulty & ")|"
        ExportItemPYCode = ExportItemPYCode & "spd_rtng(" & itm.speed_rating & ")|"
        ExportItemPYCode = ExportItemPYCode & "shoot_speed(" & itm.missile_speed & ")|"
        If itm.leg_armor <> 0 Then ExportItemPYCode = ExportItemPYCode & "accuracy(" & itm.leg_armor & ")|"
        If itm.abundance > 0 Then ExportItemPYCode = ExportItemPYCode & "abundance(" & itm.abundance & ")|"
             Dam2 = GetDamage(itm.thrust_damage, tpD2)
             ExportItemPYCode = ExportItemPYCode & "thrust_damage(" & Dam2 & ", "
             If tpD2 = 0 Then
                ExportItemPYCode = ExportItemPYCode & "cut)|"
             ElseIf tpD2 = 1 Then
                ExportItemPYCode = ExportItemPYCode & "pierce)|"
             ElseIf tpD2 = 2 Then
                ExportItemPYCode = ExportItemPYCode & "blunt)|"
        End If
        ExportItemPYCode = ExportItemPYCode & "max_ammo(" & itm.max_ammo & ")|"
   ElseIf IsAmmo(itm.itmType) Then
        ExportItemPYCode = ExportItemPYCode & "weight(" & Round(Val(itm.weight), 2) & ")|"
        ExportItemPYCode = ExportItemPYCode & "weapon_length(" & itm.weapon_length & ")|"
        If itm.abundance > 0 Then ExportItemPYCode = ExportItemPYCode & "abundance(" & itm.abundance & ")|"
             Dam2 = GetDamage(itm.thrust_damage, tpD2)
             ExportItemPYCode = ExportItemPYCode & "thrust_damage(" & Dam2 & ", "
             If tpD2 = 0 Then
                ExportItemPYCode = ExportItemPYCode & "cut)|"
             ElseIf tpD2 = 1 Then
                ExportItemPYCode = ExportItemPYCode & "pierce)|"
             ElseIf tpD2 = 2 Then
                ExportItemPYCode = ExportItemPYCode & "blunt)|"
        End If
        ExportItemPYCode = ExportItemPYCode & "max_ammo(" & itm.max_ammo & ")|"
   ElseIf IsArmor(itm.itmType) Then
        ExportItemPYCode = ExportItemPYCode & "weight(" & Round(Val(itm.weight), 2) & ")|"
        If itm.abundance > 0 Then ExportItemPYCode = ExportItemPYCode & "abundance(" & itm.abundance & ")|"
        If itm.difficulty > 0 Then ExportItemPYCode = ExportItemPYCode & "difficulty(" & itm.difficulty & ")|"
        ExportItemPYCode = ExportItemPYCode & "head_armor(" & itm.head_armor & ")|"
        ExportItemPYCode = ExportItemPYCode & "body_armor(" & itm.body_armor & ")|"
        ExportItemPYCode = ExportItemPYCode & "leg_armor(" & itm.leg_armor & ")|"
   ElseIf IsShield(itm.itmType) Then
        ExportItemPYCode = ExportItemPYCode & "weight(" & Round(Val(itm.weight), 2) & ")|"
        ExportItemPYCode = ExportItemPYCode & "shield_width(" & itm.weapon_length & ")|"
        If itm.abundance > 0 Then ExportItemPYCode = ExportItemPYCode & "abundance(" & itm.abundance & ")|"
        ExportItemPYCode = ExportItemPYCode & "hit_points(" & itm.hit_points & ")|"
        ExportItemPYCode = ExportItemPYCode & "body_armor(" & itm.body_armor & ")|"
        ExportItemPYCode = ExportItemPYCode & "spd_rtng(" & itm.speed_rating & ")|"
   ElseIf IsHorse(itm.itmType) Then
        ExportItemPYCode = ExportItemPYCode & "difficulty(" & itm.difficulty & ")|"
        If itm.abundance > 0 Then ExportItemPYCode = ExportItemPYCode & "abundance(" & itm.abundance & ")|"
        ExportItemPYCode = ExportItemPYCode & "hit_points(" & itm.hit_points & ")|"
        ExportItemPYCode = ExportItemPYCode & "body_armor(" & itm.body_armor & ")|"
        ExportItemPYCode = ExportItemPYCode & "horse_charge(" & itm.thrust_damage & ")|"
        ExportItemPYCode = ExportItemPYCode & "horse_maneuver(" & itm.speed_rating & ")|"
        ExportItemPYCode = ExportItemPYCode & "horse_speed(" & itm.missile_speed & ")|"
        If itm.weapon_length > 0 Then ExportItemPYCode = ExportItemPYCode & "horse_scale(" & itm.weapon_length & ")|"
   ElseIf IsFood(itm.itmType) Then
        ExportItemPYCode = ExportItemPYCode & "weight(" & Round(Val(itm.weight), 2) & ")|"
        ExportItemPYCode = ExportItemPYCode & "food_quality(" & itm.head_armor & ")|"
        ExportItemPYCode = ExportItemPYCode & "max_ammo(" & itm.max_ammo & ")|"
   Else
        ExportItemPYCode = ExportItemPYCode & "weight(" & Round(Val(itm.weight), 2) & ")|"
        If itm.abundance > 0 Then ExportItemPYCode = ExportItemPYCode & "abundance(" & itm.abundance & ")|"

   End If
'末尾处理
   If Right(ExportItemPYCode, 1) = "|" Then
        ExportItemPYCode = Left(ExportItemPYCode, Len(ExportItemPYCode) - 1)
   End If
   
   ExportItemPYCode = ExportItemPYCode & "," & " "
   
'IMod
Dim tB As Boolean
   For i = 0 To UBound(IModC)
      If itm.Prefix = IModC(i).X Then
         ExportItemPYCode = ExportItemPYCode & IModC(i).Y
         tB = True
         Exit For
      End If
   Next i
   If tB = False Then
       tI = StrToI64(itm.Prefix)
       For n = 0 To N_IMod - 1
           If ChkBit64b(tI, n) Then
           ExportItemPYCode = ExportItemPYCode & IMod(n).ID & "|"
           End If
       Next n
   End If

'末尾处理
   If Right(ExportItemPYCode, 1) = "|" Then
        ExportItemPYCode = Left(ExportItemPYCode, Len(ExportItemPYCode) - 1)
   ElseIf Right(ExportItemPYCode, 1) = "," Then
        ExportItemPYCode = ExportItemPYCode & "0"
   End If
   
   ExportItemPYCode = ExportItemPYCode & "," & " "

'触发器
   If itm.TriggerCount > 0 Then
     ExportItemPYCode = ExportItemPYCode & vbCrLf & "["
     For i = 1 To itm.TriggerCount
       ExportItemPYCode = ExportItemPYCode & "(" & tiOns(GettiOnIndex(itm.Trigger(i).tiOn)).dbName & ","
       If itm.Trigger(i).ActNum > 0 Then
         ExportItemPYCode = ExportItemPYCode & "[" & vbCrLf
         ExportPYCodefromTXT_OpBlocks itm.Trigger(i).tiAct(), temStr, , , , "   "
         ExportItemPYCode = ExportItemPYCode & temStr
         ExportItemPYCode = ExportItemPYCode & "])"
       Else
         ExportItemPYCode = ExportItemPYCode & "[])"
       End If
       
       If i < itm.TriggerCount Then
         ExportItemPYCode = ExportItemPYCode & ", " & vbCrLf
       Else
         ExportItemPYCode = ExportItemPYCode & vbCrLf & "], " & vbCrLf
       End If
     Next i
   Else
     ExportItemPYCode = ExportItemPYCode & "[], "
   End If
   
'阵营
   If itm.FactionCount > 0 Then
     ExportItemPYCode = ExportItemPYCode & "["
     For i = 1 To itm.FactionCount
        ExportItemPYCode = ExportItemPYCode & Factions(itm.Faction(i).ID).strID
        If i < itm.FactionCount Then
          ExportItemPYCode = ExportItemPYCode & ","
        End If
     Next i
     ExportItemPYCode = ExportItemPYCode & "]"
   End If
   
'最后的末尾处理
   If Right(ExportItemPYCode, 2) = ", " Then
        ExportItemPYCode = Left(ExportItemPYCode, Len(ExportItemPYCode) - 2) & " "
   End If
   ExportItemPYCode = ExportItemPYCode & "],"
   
    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModPython", "ExportItemPYCode", Err.Number, Err.Description)
End Function

Private Function GetFlagName(FlagStr As String, Optional FlagNo As Integer = 1)
Dim temStr As Variant
temStr = Split(FlagStr, "/")

If FlagNo - 1 <= UBound(temStr) Then
    GetFlagName = temStr(FlagNo - 1)
Else
    GetFlagName = ""
End If

End Function

Public Function GetTriggerFunctionIndex(Opblocks1() As Type_Op_Block, Opblocks2() As Type_Op_Block) As Integer
Dim i As Integer, n As Integer, H As Integer, j As Integer, temArr As Variant, IsMatch1 As Boolean, IsMatch2 As Boolean, Match As Boolean

For i = 1 To UBound(TrgFunc)             '遍历特征库
    If UBound(TrgFunc(i).Opblock1) = 0 Then
         IsMatch1 = True
    ElseIf UBound(Opblocks1) = 0 Then
         IsMatch1 = False
    Else
         H = 0
         For j = 1 To UBound(TrgFunc(i).Opblock1)     '遍历特征库操作块1中的所有operation
             For n = 1 To UBound(Opblocks1)
                  If IsOperationMatch(Opblocks1(n), TrgFunc(i).Opblock1(j)) Then
                       H = H + 1
                       Exit For
                  End If
             Next n
         Next j
         If H = UBound(TrgFunc(i).Opblock1) Then
            IsMatch1 = True
         Else
            IsMatch1 = False
         End If
    End If

    If UBound(TrgFunc(i).OpBlock2) = 0 Then
         IsMatch2 = True
    ElseIf UBound(Opblocks2) = 0 Then
         IsMatch2 = False
    Else
         H = 0
         For j = 1 To UBound(TrgFunc(i).OpBlock2)     '遍历特征库操作块2中的所有operation
             For n = 1 To UBound(Opblocks2)
                  If IsOperationMatch(Opblocks2(n), TrgFunc(i).OpBlock2(j)) Then
                       H = H + 1
                       Exit For
                  End If
             Next n
         Next j
         If H = UBound(TrgFunc(i).OpBlock2) Then
            IsMatch2 = True
         Else
            IsMatch2 = False
         End If
    End If
    
If IsMatch1 And IsMatch2 Then
    GetTriggerFunctionIndex = i
    Exit Function
End If

Next i

GetTriggerFunctionIndex = 0

End Function

Public Function IsOperationMatch(Opblock1 As Type_Op_Block, OpblockPack As Type_Op_Block) As Boolean
Dim i As Integer, j As Integer

If Opblock1.Op = OpblockPack.Op Then
     If Opblock1.ParaNum = 0 Or OpblockPack.ParaNum = 0 Then
         IsOperationMatch = True
     Else
         IsOperationMatch = True
         If Opblock1.ParaNum < OpblockPack.ParaNum Then
             IsOperationMatch = False
         Else
             For i = 1 To OpblockPack.ParaNum
                If Not (IsParamMatch(Opblock1.Para(i), OpblockPack.Para(i))) Then
                    IsOperationMatch = False
                    Exit For
                End If
             Next i
         End If
     End If
Else
     IsOperationMatch = False
End If

End Function

Public Function IsParamMatch(Param As Type_Param, ParamPack As Type_Param) As Boolean
Dim i As Integer, temArr As Variant

If ParamPack.Value = "" And ParamPack.strID = "" Then
    IsParamMatch = True
    Exit Function
End If

If ParamPack.strID <> "" Then
    If InStr(1, ParamPack.strID, "#") Then
       temArr = Split(ParamPack.strID, "#")
       For i = 0 To UBound(temArr)
           If CStr(temArr(i)) = Param.strID Then
               IsParamMatch = True
               Exit Function
           End If
       Next i
       IsParamMatch = False
    Else
       If Param.strID = ParamPack.strID Then
           IsParamMatch = True
       Else
           IsParamMatch = False
       End If
    End If
Else
    If InStr(1, ParamPack.Value, "#") Then
       temArr = Split(ParamPack.Value, "#")
       For i = 0 To UBound(temArr)
           If CStr(temArr(i)) = Param.Value Then
               IsParamMatch = True
               Exit Function
           End If
       Next i
       IsParamMatch = False
    Else
       If Param.Value = ParamPack.Value Then
           IsParamMatch = True
       Else
           IsParamMatch = False
       End If
    End If
End If

End Function

'*************************************************************************
'**函 数 名：ExportWeaponProficiencies
'**输    入：(Type_Item)itm
'**输    出：(String)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：
'**日    期：
'**修 改 人：SSgt_Edward
'**日    期：2011-03-25 22:44:45
'**版    本：V1.1321
'*************************************************************************
Public Function ExportWeaponProficiencies(WeaponProf As Type_WeaponProf) As String
Dim q As Boolean, resTem As String, lngTem(6) As Long, i As Integer

'熟练度
With WeaponProf

lngTem(0) = .one_handed
lngTem(1) = .two_handed
lngTem(2) = .polearm
lngTem(3) = .archery
lngTem(4) = .crossbow
lngTem(5) = .throwing
lngTem(6) = .firearm

'wp_melee(x)
q = False
If lngTem(0) = lngTem(1) + 20 Then
    If lngTem(2) = lngTem(1) + 10 Then
        q = True
    End If
End If

If q Then
  resTem = "wp_melee(" & .two_handed & ") | "
  GoTo WP_Melee
End If

'wp(x)
q = True
If lngTem(0) > 0 Then
   For i = 1 To 5
       If lngTem(0) <> lngTem(i) Then
          q = False
          Exit For
       End If
   Next i
Else
   q = False
End If

If q Then
  resTem = "wp(" & .one_handed & ") | "
  GoTo WP
End If

'wpe(m,a,c,t)          暂且认为前三项相同,且每项不为零即匹配(每项相同已被wp排除)
q = True
If lngTem(0) > 0 Then
   For i = 3 To 5
       If lngTem(i) <= 0 Then
          q = False
          Exit For
       End If
   Next i
Else
   q = False
End If

If q Then
   For i = 1 To 2
       If lngTem(0) <> lngTem(i) Then
          q = False
          Exit For
       End If
   Next i
End If

If q Then
  resTem = "wpe(" & .one_handed & "," & .archery & "," & .crossbow & "," & .throwing & ") | "
  GoTo WP
End If

'wpex(o,w,p,a,c,t)
q = True
   For i = 0 To 5
       If lngTem(i) <= 0 Then
          q = False
          Exit For
       End If
   Next i

If q Then
   resTem = "wpex(" & .one_handed & "," & .two_handed & "," & .polearm & "," & .archery & "," & .crossbow & "," & .throwing & ") | "
   GoTo WP
End If

'wps
If .one_handed > 0 Then
   resTem = "wp_one_handed (" & .one_handed & ") | "
End If

If .two_handed > 0 Then
   resTem = resTem & "wp_two_handed (" & .two_handed & ") | "
End If

If .polearm > 0 Then
   resTem = resTem & "wp_polearm (" & .polearm & ") | "
End If

WP_Melee:
If .archery > 0 Then
   resTem = resTem & "wp_archery (" & .archery & ") | "
End If

If .crossbow > 0 Then
   resTem = resTem & "wp_crossbow (" & .crossbow & ") | "
End If

If .throwing > 0 Then
   resTem = resTem & "wp_throwing (" & .throwing & ") | "
End If

WP:
If .firearm > 0 Then
   resTem = resTem & "wp_firearm (" & .firearm & ") | "
End If

If Len(resTem) > 1 Then
   resTem = Left(resTem, Len(resTem) - 2)
Else
   resTem = "0"
End If

End With

ExportWeaponProficiencies = resTem
End Function

'*************************************************************************
'**函 数 名：ExportPYCodefromTXT_OpBlocks
'**输    入：(Type_Item)itm
'**输    出：(String)
'**功能描述：反编译操作块
'**全局变量：
'**调用模块：
'**作    者：
'**日    期：
'**修 改 人：Ser_Charles
'**日    期：2011-06-07 13:31:04
'**版    本：V1.1321
'*************************************************************************
Public Function ExportPYCodefromTXT_OpBlocks(OpBlocks() As Type_Op_Block, ExportStr As String, Optional ErrOp As Integer, Optional ErrParam As Integer, Optional ErrReason As Integer, Optional IndentationStr As String = "") As Boolean
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
Dim i As Integer, n As Integer
Dim neg As Integer, negs(3) As String, OpID As Long, OpIndex As Long, TagNo As Integer, Pid As String
Dim temStr As String, temLine As String, temParam As String, Indentation() As String
Dim notFail As Boolean

negs(0) = ""
negs(1) = "neg|"
negs(2) = "this_or_next|"
negs(3) = "neg|this_or_next|"

ExportStr = ""

If IndentationStr <> "" Then
   CalcOperationIndentation OpBlocks(), Indentation()
Else
   ReDim Indentation(1 To UBound(OpBlocks))
   For i = 1 To UBound(Indentation)
      Indentation(i) = ""
   Next i
End If

For i = 1 To UBound(OpBlocks)
  With OpBlocks(i)
    GetOpCodeInfo .Op, neg, OpID
    OpIndex = GetOpIndex(OpID)
    If OpIndex >= 0 Then
       temLine = Indentation(i) & "(" & negs(neg) & Operation(OpIndex).Op_name
    Else
       temLine = Indentation(i) & "(" & negs(neg) & OpID
    End If
    If .ParaNum > 0 Then
      temLine = temLine & ","
      For n = 1 To UBound(.Para)
        GetParamCodeInfo .Para(n).Value, TagNo, Pid
        If TagNo = 0 Then
           If OpIndex >= 0 Then
             If n <= Operation(OpIndex).ParaNum Then
                notFail = GetParamPYCode(Pid, Operation(OpIndex).Para(n).Para_Type, temParam, ErrReason)
                If Not notFail Then GoTo errorHandle
             Else
                notFail = GetParamPYCode(Pid, "0", temParam, ErrReason)
                If Not notFail Then GoTo errorHandle
             End If
           Else
             notFail = GetParamPYCode(Pid, "0", temParam, ErrReason)
             If Not notFail Then GoTo errorHandle
           End If
        Else
           notFail = GetParamPYCode(Pid, CStr(TagNo), temParam, ErrReason)
           If Not notFail Then GoTo errorHandle
        End If
      
        If n < .ParaNum Then
          temLine = temLine & temParam & ","
        Else
          temLine = temLine & temParam & "),"
        End If
      Next n
    Else
      temLine = temLine & "),"
    End If
    ExportStr = ExportStr & temLine & vbCrLf
  End With
Next i
    ExportPYCodefromTXT_OpBlocks = True
    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    ExportPYCodefromTXT_OpBlocks = False
    ErrOp = i
    ErrParam = n
    ErrReason = Err.Number
    'MsgBox "error!"
End Function
'*************************************************************************
'**函 数 名：GetParamPYCode
'**输    入：(Long)PID, (String)TagNo , (String)ReturnValue , (Int)Optional ErrReason
'**输    出：(Boolean)
'**功能描述：获得参数PY代码
'**全局变量：
'**调用模块：
'**作    者：
'**日    期：
'**修 改 人：Ser_Charles
'**日    期：2011-06-27 14:13:22
'**版    本：V1.1321
'*************************************************************************
Public Function GetParamPYCode(Pid As String, TagNo As String, ReturnValue As String, Optional ErrReason As Integer) As Boolean
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    Dim tem As String, i As Integer

    GetParamPYCode = True
    ErrStr = ""
    
    Select Case TagNo      'ends_add
         Case ""
            tem = Pid
         Case "bs"
            tem = Pid
         Case "ap"
            tem = Pid
         Case "as"
            tem = Pid
         Case "po"
            tem = Pid
         Case "itp"
            tem = Item_Type(Val(Pid)).X
         Case "tf"
            For i = 0 To UBound(Tf)
               If I64toStrNZ(Tf(i).Value) = Pid Then Exit For
            Next i
            tem = Tf(i).strName
         Case "pf"
            For i = 0 To UBound(Pf)
               If I64toStrNZ(Pf(i).Value) = Pid Then Exit For
            Next i
            tem = Pf(i).strName
         Case "ai_bhvr"
            tem = AI_Bhvr(Val(Pid)).X
         Case "pos"
            tem = "pos" & Pid
         Case "s"
            tem = "s" & Pid
         Case "1"
            tem = GetPYPrefix(CInt(TagNo)) & Pid
         Case "2"
            tem = GetPYPrefix(CInt(TagNo)) & GetVariablePYCode(Val(Pid), True) & Chr(34)
         Case "3"
            tem = GetPYPrefix(CInt(TagNo)) & Pid & Chr(34)
         Case "4"
            tem = GetPYPrefix(CInt(TagNo)) & RemoveTag(itm(CLng(Pid)).dbName) & Chr(34)
         Case "5"
            tem = GetPYPrefix(CInt(TagNo)) & RemoveTag(Trps(CLng(Pid)).strID) & Chr(34)
         Case "6"
            tem = GetPYPrefix(CInt(TagNo)) & RemoveTag(Factions(CLng(Pid)).strID) & Chr(34)
         Case "7"
            tem = GetPYPrefix(CInt(TagNo)) & Pid & Chr(34)
         Case "8"
            tem = GetPYPrefix(CInt(TagNo)) & RemoveTag(PTs(CLng(Pid)).ptID) & Chr(34)
         Case "9"
            tem = GetPYPrefix(CInt(TagNo)) & RemoveTag(Parties(CLng(Pid)).strID) & Chr(34)
         Case "10"
            tem = GetPYPrefix(CInt(TagNo)) & RemoveTag(Scenes(CLng(Pid)).strID) & Chr(34)
         Case "11"
            tem = GetPYPrefix(CInt(TagNo)) & Pid & Chr(34)
         Case "12"
            tem = GetPYPrefix(CInt(TagNo)) & Pid & Chr(34)
         Case "13"    'script
            tem = GetPYPrefix(CInt(TagNo)) & Pid & Chr(34)
         Case "14"
            tem = GetPYPrefix(CInt(TagNo)) & RemoveTag(PSys(CLng(Pid)).strID) & Chr(34)
         Case "15"
            tem = GetPYPrefix(CInt(TagNo)) & Pid & Chr(34)
         Case "16"
            tem = GetPYPrefix(CInt(TagNo)) & RemoveTag(Sounds(CLng(Pid)).sndName) & Chr(34)
         Case "17"
            tem = GetPYPrefix(CInt(TagNo)) & GetVariablePYCode(Val(Pid), False) & Chr(34)
         Case "18"
            tem = GetPYPrefix(CInt(TagNo)) & MapIcons(Pid).strID & Chr(34)
         Case "19"
            tem = GetPYPrefix(CInt(TagNo)) & Pid & Chr(34)
         Case "20"
            tem = GetPYPrefix(CInt(TagNo)) & RemoveTag(Mesh(CLng(Pid)).strID) & Chr(34)
         Case "21"
            tem = GetPYPrefix(CInt(TagNo)) & Pid & Chr(34)
         Case "22"
            tem = Chr(34) & "@qstr_" & Pid & Chr(34)
         Case "23"
            tem = GetPYPrefix(CInt(TagNo)) & Pid & Chr(34)
         Case "24"
            tem = GetPYPrefix(CInt(TagNo)) & RemoveTag(TabMat(CLng(Pid)).strID) & Chr(34)
         Case "25"
            tem = GetPYPrefix(CInt(TagNo)) & Pid & Chr(34)
         Case Else
            tem = Pid
    End Select
    ReturnValue = tem
    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    GetParamPYCode = False
    ErrReason = Err.Number
End Function

'*************************************************************************
'**函 数 名：GetNegWord
'**输    入：(Int)neg
'**输    出：(String)
'**功能描述：获得操作符描述
'**全局变量：
'**调用模块：
'**作    者：
'**日    期：
'**修 改 人：SSgt_Edward
'**日    期：2011-07-21 16:48:12
'**版    本：V1.1321
'*************************************************************************
Public Function GetNegWord(neg As Integer) As String
Select Case neg
  Case 0
    GetNegWord = PublicMsgs(157)
  Case 1
    GetNegWord = PublicMsgs(153)
  Case 2
    GetNegWord = PublicMsgs(154)
  Case 3
    GetNegWord = PublicMsgs(161)
End Select
End Function

'*************************************************************************
'**函 数 名：CalcOperationIndentation
'**输    入：(Type_Op_Block)OpBlock(),(String)Indentation()
'**输    出：-
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-2-27 12:48:11
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub CalcOperationIndentation(OpBlock() As Type_Op_Block, Indentation() As String)   'remained to be removed to ModPython or ModOperation
Dim i As Integer, n As Integer, H As Integer, tStr As String, temOp() As Type_Op_Block

ReDim temOp(LBound(OpBlock) To UBound(OpBlock))

For i = LBound(temOp) To UBound(temOp)
    temOp(i) = OpBlock(i)
Next i

ReDim Indentation(LBound(OpBlock) To UBound(OpBlock))

For i = LBound(temOp) To UBound(temOp)
    If temOp(i).Op = CStr(try_end) Then
         For n = i To LBound(temOp) Step -1
              tStr = RemoveOperationNegations(temOp(n).Op)
              If Val(tStr) = try_begin Or Val(tStr) = try_for_range Or Val(tStr) = try_for_range_backwards Or Val(tStr) = try_for_parties Or Val(tStr) = try_for_agents Then
                  For H = n + 1 To i - 1
                      If Val(temOp(H).Op) = else_try Then
                          temOp(H).Op = "0"
                      Else
                          Indentation(H) = Indentation(H) & "    "
                      End If
                  Next H
                  temOp(n).Op = "0"
                  Exit For
              End If
         Next n
    End If
Next i

End Sub

Public Function tiOnIndex(Value As Double) As Integer
Dim i As Integer

tiOnIndex = -1
For i = 0 To UBound(tiOns)
  If tiOns(i).Value = Value Then
    tiOnIndex = i
    Exit For
  End If
Next i

End Function
