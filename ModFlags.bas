Attribute VB_Name = "ModFlags"
Option Explicit

'Troop flags
Public Const tf_male = "0"
Public Const tf_female = "1"
Public Const tf_undead = "2"
Public Const troop_type_mask = "F"
Public Const tf_hero = "10"             '16
Public Const tf_inactive = "20"
Public Const tf_unkillable = "40"
Public Const tf_allways_fall_dead = "80"
Public Const tf_no_capture_alive = "100"
Public Const tf_mounted = "400"                'Troop's movement speed on map is determined by riding skill.
Public Const tf_is_merchant = "1000"           'When set, troop does not equip stuff he owns
Public Const tf_randomize_face = "8000"        'randomize face at the beginning of the game.

Public Const tf_guarantee_boots = "100000"
Public Const tf_guarantee_armor = "200000"
Public Const tf_guarantee_helmet = "400000"
Public Const tf_guarantee_gloves = "800000"
Public Const tf_guarantee_horse = "1000000"
Public Const tf_guarantee_shield = "2000000"
Public Const tf_guarantee_ranged = "4000000"               '2 ^ 26
Public Const tf_unmoveable_in_party_window = "10000000"    '2 ^ 28

'Scene Flags
Public Const sf_indoors = "1"                    'The scene shouldn't have a skybox and lighting by sun.
Public Const sf_force_skybox = "2"               'Force adding a skybox even if indoors flag is set.
Public Const sf_generate = "100"                 'Generate terrain by terran-generator
Public Const sf_randomize = "200"                'Randomize terrain generator key
Public Const sf_auto_entry_points = "400"        'Automatically create entry points
Public Const sf_no_horses = "800"                'Horses are not avaible
Public Const sf_muddy_water = "1000"             'Changes the shader of the river mesh

'Faction Flags
Public Const ff_always_hide_label = &H1
Public Const ff_max_rating_bits = 8
Public Const ff_max_rating_mask = 65280

'Sound
Public Const sf_2d = "00000001"
Public Const sf_looping = "00000002"
Public Const sf_start_at_random_pos = "00000004"

Public Const sf_priority_10 = "000000A0"
Public Const sf_priority_9 = "00000090"
Public Const sf_priority_8 = "00000080"
Public Const sf_priority_7 = "00000070"
Public Const sf_priority_6 = "00000060"
Public Const sf_priority_5 = "00000050"
Public Const sf_priority_4 = "00000040"
Public Const sf_priority_3 = "00000030"
Public Const sf_priority_2 = "00000020"
Public Const sf_priority_1 = "00000010"

Public Const sf_vol_15 = "00000F00"
Public Const sf_vol_14 = "00000E00"
Public Const sf_vol_13 = "00000D00"
Public Const sf_vol_12 = "00000C00"
Public Const sf_vol_11 = "00000B00"
Public Const sf_vol_10 = "00000A00"
Public Const sf_vol_9 = "00000900"
Public Const sf_vol_8 = "00000800"
Public Const sf_vol_7 = "00000700"
Public Const sf_vol_6 = "00000600"
Public Const sf_vol_5 = "00000500"
Public Const sf_vol_4 = "00000400"
Public Const sf_vol_3 = "00000300"
Public Const sf_vol_2 = "00000200"
Public Const sf_vol_1 = "00000100"

'Map Icon
Public Const mcn_no_shadow = &H1

'Party
Public Const pf_icon_mask = "FF"
Public Const pf_disabled = "100"
Public Const pf_is_ship = "200"
Public Const pf_is_static = "400"

Public Const pf_label_small = "0"
Public Const pf_label_medium = "1000"
Public Const pf_label_large = "2000"
Public Const pf_label_mask = "3000"

Public Const pf_always_visible = "4000"
Public Const pf_default_behavior = "10000"
Public Const pf_auto_remove_in_town = "20000"
Public Const pf_quest_party = "40000"
Public Const pf_no_label = "80000"
Public Const pf_limit_members = "100000"
Public Const pf_hide_defenders = "200000"
Public Const pf_show_faction = "400000"
Public Const pf_is_hidden = "1000000"                 'used in the engine, do not overwrite this flag
Public Const pf_dont_attack_civilians = "2000000"
Public Const pf_civilian = "4000000"
'plus
Public Const pf_others = "0"
Public Const pf_town = "406400"
Public Const pf_castle = "405400"
Public Const pf_village = "204400"
Public Const pf_bridge = "84400"
Public Const pf_respawnpoint = "500"

Public Const pf_carry_goods_bits = 48
Public Const pf_carry_gold_bits = 56
Public Const pf_carry_gold_multiplier = 20
Public Const pf_carry_goods_mask = "00ff000000000000"
Public Const pf_carry_gold_mask = "ff00000000000000"

Public Const pmf_is_prisoner = &H1

Public Const no_faction = -1

Public Const ai_bhvr_hold = 0
Public Const ai_bhvr_travel_to_party = 1
Public Const ai_bhvr_patrol_location = 2
Public Const ai_bhvr_patrol_party = 3
Public Const ai_bhvr_track_party = 4     'deprecated, use the alias ai_bhvr_attack_party instead.
Public Const ai_bhvr_attack_party = 4
Public Const ai_bhvr_avoid_party = 5
Public Const ai_bhvr_travel_to_point = 6
Public Const ai_bhvr_negotiate_party = 7
Public Const ai_bhvr_in_town = 8
Public Const ai_bhvr_travel_to_ship = 9
Public Const ai_bhvr_escort_party = 10
Public Const ai_bhvr_driven_by_party = 11

'experience constants
Public Const player_loot_share = 10
Public Const hero_loot_share = 3


'personality modifiers:
' courage 8 means neutral
Public Const courage_4 = &H4
Public Const courage_5 = &H5
Public Const courage_6 = &H6
Public Const courage_7 = &H7
Public Const courage_8 = &H8
Public Const courage_9 = &H9
Public Const courage_10 = &HA
Public Const courage_11 = &HB
Public Const courage_12 = &HC
Public Const courage_13 = &HD
Public Const courage_14 = &HE
Public Const courage_15 = &HF

Public Const aggressiveness_0 = &H0
Public Const aggressiveness_1 = &H10
Public Const aggressiveness_2 = &H20
Public Const aggressiveness_3 = &H30
Public Const aggressiveness_4 = &H40
Public Const aggressiveness_5 = &H50
Public Const aggressiveness_6 = &H60
Public Const aggressiveness_7 = &H70
Public Const aggressiveness_8 = &H80
Public Const aggressiveness_9 = &H90
Public Const aggressiveness_10 = &HA0
Public Const aggressiveness_11 = &HB0
Public Const aggressiveness_12 = &HC0
Public Const aggressiveness_13 = &HD0
Public Const aggressiveness_14 = &HE0
Public Const aggressiveness_15 = &HF0

Public Const banditness = &H100
'soldier_personality = aggressiveness_8 | courage_9
'merchant_personality = aggressiveness_0 | courage_7
'escorted_merchant_personality = aggressiveness_0 | courage_11
'bandit_personality   = aggressiveness_3 | courage_8 | banditness

'Map Icons
Public Const icon_player = 0
Public Const icon_player_horseman = 1
Public Const icon_gray_knight = 2
Public Const icon_vaegir_knight = 3
Public Const icon_flagbearer_a = 4
Public Const icon_flagbearer_b = 5
Public Const icon_peasant = 6
Public Const icon_khergit = 7
Public Const icon_khergit_horseman_b = 8
Public Const icon_axeman = 9
Public Const icon_woman = 10
Public Const icon_woman_b = 11
Public Const icon_town = 12
Public Const icon_town_steppe = 13
Public Const icon_town_desert = 14
Public Const icon_village_a = 15
Public Const icon_village_b = 16
Public Const icon_village_c = 17
Public Const icon_village_burnt_a = 18
Public Const icon_village_deserted_a = 19
Public Const icon_village_burnt_b = 20
Public Const icon_village_deserted_b = 21
Public Const icon_village_burnt_c = 22
Public Const icon_village_deserted_c = 23
Public Const icon_village_snow_a = 24
Public Const icon_village_snow_burnt_a = 25
Public Const icon_village_snow_deserted_a = 26
Public Const icon_camp = 27
Public Const icon_ship = 28
Public Const icon_ship_on_land = 29
Public Const icon_castle_a = 30
Public Const icon_castle_b = 31
Public Const icon_castle_c = 32
Public Const icon_castle_d = 33
Public Const icon_town_snow = 34
Public Const icon_castle_snow_a = 35
Public Const icon_castle_snow_b = 36
Public Const icon_mule = 37
Public Const icon_cattle = 38
Public Const icon_training_ground = 39
Public Const icon_bridge_a = 40
Public Const icon_bridge_b = 41
Public Const icon_bridge_snow_a = 42
Public Const icon_custom_banner_01 = 43
Public Const icon_custom_banner_02 = 44
Public Const icon_custom_banner_03 = 45
Public Const icon_banner_01 = 46
Public Const icon_banner_02 = 47
Public Const icon_banner_03 = 48
Public Const icon_banner_04 = 49
Public Const icon_banner_05 = 50
Public Const icon_banner_06 = 51
Public Const icon_banner_07 = 52
Public Const icon_banner_08 = 53
Public Const icon_banner_09 = 54
Public Const icon_banner_10 = 55
Public Const icon_banner_11 = 56
Public Const icon_banner_12 = 57
Public Const icon_banner_13 = 58
Public Const icon_banner_14 = 59
Public Const icon_banner_15 = 60
Public Const icon_banner_16 = 61
Public Const icon_banner_17 = 62
Public Const icon_banner_18 = 63
Public Const icon_banner_19 = 64
Public Const icon_banner_20 = 65
Public Const icon_banner_21 = 66
Public Const icon_banner_22 = 67
Public Const icon_banner_23 = 68
Public Const icon_banner_24 = 69
Public Const icon_banner_25 = 70
Public Const icon_banner_26 = 71
Public Const icon_banner_27 = 72
Public Const icon_banner_28 = 73
Public Const icon_banner_29 = 74
Public Const icon_banner_30 = 75
Public Const icon_banner_31 = 76
Public Const icon_banner_32 = 77
Public Const icon_banner_33 = 78
Public Const icon_banner_34 = 79
Public Const icon_banner_35 = 80
Public Const icon_banner_36 = 81
Public Const icon_banner_37 = 82
Public Const icon_banner_38 = 83
Public Const icon_banner_39 = 84
Public Const icon_banner_40 = 85
Public Const icon_banner_41 = 86
Public Const icon_banner_42 = 87
Public Const icon_banner_43 = 88
Public Const icon_banner_44 = 89
Public Const icon_banner_45 = 90
Public Const icon_banner_46 = 91
Public Const icon_banner_47 = 92
Public Const icon_banner_48 = 93
Public Const icon_banner_49 = 94
Public Const icon_banner_50 = 95
Public Const icon_banner_51 = 96
Public Const icon_banner_52 = 97
Public Const icon_banner_53 = 98
Public Const icon_banner_54 = 99
Public Const icon_banner_55 = 100
Public Const icon_banner_56 = 101
Public Const icon_banner_57 = 102
Public Const icon_banner_58 = 103
Public Const icon_banner_59 = 104
Public Const icon_banner_60 = 105
Public Const icon_banner_61 = 106
Public Const icon_banner_62 = 107
Public Const icon_banner_63 = 108
Public Const icon_banner_64 = 109
Public Const icon_banner_65 = 110
Public Const icon_banner_66 = 111
Public Const icon_banner_67 = 112
Public Const icon_banner_68 = 113
Public Const icon_banner_69 = 114
Public Const icon_banner_70 = 115
Public Const icon_banner_71 = 116
Public Const icon_banner_72 = 117
Public Const icon_banner_73 = 118
Public Const icon_banner_74 = 119
Public Const icon_banner_75 = 120
Public Const icon_banner_76 = 121
Public Const icon_banner_77 = 122
Public Const icon_banner_78 = 123
Public Const icon_banner_79 = 124
Public Const icon_banner_80 = 125
Public Const icon_banner_81 = 126
Public Const icon_banner_82 = 127
Public Const icon_banner_83 = 128
Public Const icon_banner_84 = 129
Public Const icon_banner_85 = 130
Public Const icon_banner_86 = 131
Public Const icon_banner_87 = 132
Public Const icon_banner_88 = 133
Public Const icon_banner_89 = 134
Public Const icon_banner_90 = 135
Public Const icon_banner_91 = 136
Public Const icon_banner_92 = 137
Public Const icon_banner_93 = 138
Public Const icon_banner_94 = 139
Public Const icon_banner_95 = 140
Public Const icon_banner_96 = 141
Public Const icon_banner_97 = 142
Public Const icon_banner_98 = 143
Public Const icon_banner_99 = 144
Public Const icon_banner_100 = 145
Public Const icon_banner_101 = 146
Public Const icon_banner_102 = 147
Public Const icon_banner_103 = 148
Public Const icon_banner_104 = 149
Public Const icon_banner_105 = 150
Public Const icon_banner_106 = 151
Public Const icon_banner_107 = 152
Public Const icon_banner_108 = 153
Public Const icon_banner_109 = 154
Public Const icon_banner_110 = 155
Public Const icon_banner_111 = 156
Public Const icon_banner_112 = 157
Public Const icon_banner_113 = 158
Public Const icon_banner_114 = 159
Public Const icon_banner_115 = 160
Public Const icon_banner_116 = 161
Public Const icon_banner_117 = 162
Public Const icon_banner_118 = 163
Public Const icon_banner_119 = 164
Public Const icon_banner_120 = 165
Public Const icon_banner_121 = 166
Public Const icon_banner_122 = 167
Public Const icon_banner_123 = 168
Public Const icon_banner_124 = 169
Public Const icon_banner_125 = 170
Public Const icon_banner_126 = 171
Public Const icon_banner_127 = 172
Public Const icon_banner_128 = 173
Public Const icon_banner_129 = 174
Public Const icon_banner_130 = 175
Public Const icon_banner_131 = 176
Public Const icon_banner_132 = 177
Public Const icon_banner_133 = 178
Public Const icon_banner_134 = 179
Public Const icon_banner_135 = 180
Public Const icon_map_flag_kingdom_a = 181
Public Const icon_map_flag_kingdom_b = 182
Public Const icon_map_flag_kingdom_c = 183
Public Const icon_map_flag_kingdom_d = 184
Public Const icon_map_flag_kingdom_e = 185
Public Const icon_map_flag_kingdom_f = 186
Public Const icon_banner_136 = 187
Public Const icon_bandit_lair = 188

'*************************************************************************
'**函 数 名：AddFlags
'**输    入：(Double)NewFlags
'**输    出：(Double) -
'**功能描述：增加Flags
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-11-25 17:13:57
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function AddFlags(ByVal Flags As Double, ByVal NewFlags As Double) As Double
AddFlags = Flags
'If Not (flags And NewFlags) Then
    AddFlags = Or_28(Flags, NewFlags)
'End If
End Function

'*************************************************************************
'**函 数 名：DeleteFlags
'**输    入：(Double)NewFlags
'**输    出：(Double) -
'**功能描述：增加Flags
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-11-25 17:16:46
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function DeleteFlags(ByVal Flags As Double, ByVal NewFlags As Double) As Double
DeleteFlags = Flags
'If flags Or NewFlags Then
    DeleteFlags = And_28(Flags, Not_28(NewFlags))
'End If
End Function

'*************************************************************************
'**函 数 名：DetoBinString_28
'**输    入：(Double)Num_De
'**输    出：(String) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-11-25 17:16:46
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function DetoBinString_28(ByVal Num_De As Double) As String
Dim i As Integer, Dec As Double
Dec = Num_De

For i = 0 To 28 Step 1
       If Dec Mod 2 <> 0 Then
         DetoBinString_28 = DetoBinString_28 & "1"
       Else
         DetoBinString_28 = DetoBinString_28 & "0"
       End If
         Dec = Dec \ 2
      DoEvents
Next i
End Function

'*************************************************************************
'**函 数 名：DetoBinString_15
'**输    入：(Double)Num_De
'**输    出：(String) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-04 08:37:48
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function DetoBinString_15(ByVal Num_De As Long) As String
Dim i As Integer, Dec As Long
Dec = Num_De

For i = 0 To 15 Step 1
       If Dec Mod 2 <> 0 Then
         DetoBinString_15 = DetoBinString_15 & "1"
       Else
         DetoBinString_15 = DetoBinString_15 & "0"
       End If
         Dec = Dec \ 2
      DoEvents
Next i
End Function

'*************************************************************************
'**函 数 名：DetoBinString
'**输    入：(Double)Num_De,(Byte)Maxbit
'**输    出：(String) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-04 08:37:48
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function DetoBinString(ByVal Num_De As Double) As String
Dim i As Integer, Dec As Double
Dec = Num_De

Do While Dec <> 0
       If Dec Mod 2 <> 0 Then
         DetoBinString = DetoBinString & "1"
       Else
         DetoBinString = DetoBinString & "0"
       End If
         Dec = Dec \ 2
      DoEvents
Loop
End Function

'*************************************************************************
'**函 数 名：BinStringtoDe
'**输    入：(Double)Num_De
'**输    出：(String) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-11-25 17:16:46
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function BinStringtoDe(ByVal StrBin As String) As Double
Dim i As Integer, K As String

If Trim(StrBin) = "" Then
   Exit Function
End If

For i = 1 To Len(StrBin) Step 1
      K = Mid(StrBin, i, 1)
      
      If K = "1" Then
          BinStringtoDe = BinStringtoDe + 2 ^ (i - 1)
      End If
Next i
End Function

'*************************************************************************
'**函 数 名：BinAnd
'**输    入：(Double)Num_De1,(Double)Num_De2
'**输    出：(String) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-04 09:05:14
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function BinAnd(ByVal Num_De1 As Double, ByVal Num_De2 As Double) As Double
Dim sB(1 To 2) As String, i As Integer, K(1 To 2) As String, j As Integer, res As Integer, TemResult As String, MaxLen As Long
sB(1) = DetoBinString(Num_De1)
sB(2) = DetoBinString(Num_De2)

MaxLen = GetMaxLen(sB(1), sB(2))

For i = 1 To MaxLen
     For j = 1 To 2
          If i > Len(sB(j)) Then
              K(j) = 0
          Else
              K(j) = Mid(sB(j), i, 1)
          End If
     Next j
     
     res = Val(K(1)) And Val(K(2))
     
     TemResult = TemResult & res
Next i

BinAnd = BinStringtoDe(TemResult)
End Function

'*************************************************************************
'**函 数 名：BinOr
'**输    入：(Double)Num_De1,(Double)Num_De2
'**输    出：(String) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-04 09:07:17
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function BinOr(ByVal Num_De1 As Double, ByVal Num_De2 As Double) As Double
Dim sB(1 To 2) As String, i As Integer, K(1 To 2) As String, j As Integer, res As Integer, TemResult As String, MaxLen As Long
sB(1) = DetoBinString(Num_De1)
sB(2) = DetoBinString(Num_De2)

MaxLen = GetMaxLen(sB(1), sB(2))

For i = 1 To MaxLen
     For j = 1 To 2
          If i > Len(sB(j)) Then
              K(j) = 0
          Else
              K(j) = Mid(sB(j), i, 1)
          End If
     Next j
     
     res = Val(K(1)) Or Val(K(2))
     
     TemResult = TemResult & res
Next i

BinOr = BinStringtoDe(TemResult)
End Function

'*************************************************************************
'**函 数 名：BinNot
'**输    入：(Double)Num_De
'**输    出：(String) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-04 09:14:22
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function BinNot(ByVal Num_De As Double, ByVal MaxBit As Byte) As Double
Dim sB As String, i As Integer, K As String, res As String, TemResult As String
sB = DetoBinString(Num_De)

For i = 1 To MaxBit
     If Len(sB) > MaxBit Then
          K = "0"
     Else
          K = Mid(sB, i, 1)
     End If
     
     If K = "1" Then
     res = "0"
     Else
     res = "1"
     End If
     
     TemResult = TemResult & res
Next i

BinNot = BinStringtoDe(TemResult)
End Function


'*************************************************************************
'**函 数 名：GetMaxLen
'**输    入：(String)Str1,(String)Str2
'**输    出：(Long) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-04 09:05:14
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function GetMaxLen(ByVal Str1 As String, ByVal Str2 As String) As Long
If Len(Str1) > Len(Str2) Then
   GetMaxLen = Len(Str1)
Else
   GetMaxLen = Len(Str2)
End If
End Function
'*************************************************************************
'**函 数 名：InvertString
'**输    入：(String)StrBin
'**输    出：(String) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-04 08:48:22
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function InvertString(ByVal StrBin As String) As String
Dim i As Integer, K As String * 1

For i = 1 To Len(StrBin)
      K = Mid(StrBin, i, 1)
      InvertString = K & InvertString
      DoEvents
Next i
End Function

'*************************************************************************
'**函 数 名：And_28
'**输    入：(Double)Num_De1,(Double)Num_De2
'**输    出：(String) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-11-25 17:16:46
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function And_28(ByVal Num_De1 As Double, ByVal Num_De2 As Double) As Double
Dim sB(1 To 2) As String * 29, i As Integer, K(1 To 2) As String, j As Integer, res As Integer, TemResult As String
sB(1) = DetoBinString_28(Num_De1)
sB(2) = DetoBinString_28(Num_De2)

For i = 1 To 29
     For j = 1 To 2
          K(j) = Mid(sB(j), i, 1)
     Next j
     
     res = Val(K(1)) And Val(K(2))
     
     TemResult = TemResult & res
Next i

And_28 = BinStringtoDe(TemResult)
End Function

'*************************************************************************
'**函 数 名：And_15
'**输    入：(Double)Num_De1,(Double)Num_De2
'**输    出：(String) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-04 08:41:02
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function And_15(ByVal Num_De1 As Long, ByVal Num_De2 As Long) As Double
Dim sB(1 To 2) As String * 16, i As Integer, K(1 To 2) As String, j As Integer, res As Integer, TemResult As String
sB(1) = DetoBinString_15(Num_De1)
sB(2) = DetoBinString_15(Num_De2)

For i = 1 To 16
     For j = 1 To 2
          K(j) = Mid(sB(j), i, 1)
     Next j
     
     res = Val(K(1)) And Val(K(2))
     
     TemResult = TemResult & res
Next i

And_15 = BinStringtoDe(TemResult)
End Function
'*************************************************************************
'**函 数 名：Or_28
'**输    入：(Double)Num_De1,(Double)Num_De2
'**输    出：(String) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-11-25 17:16:46
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function Or_28(ByVal Num_De1 As Double, ByVal Num_De2 As Double) As Double
Dim sB(1 To 2) As String * 29, i As Integer, K(1 To 2) As String, j As Integer, res As Integer, TemResult As String
sB(1) = DetoBinString_28(Num_De1)
sB(2) = DetoBinString_28(Num_De2)

For i = 1 To 29
     For j = 1 To 2
          K(j) = Mid(sB(j), i, 1)
     Next j
     
     res = Val(K(1)) Or Val(K(2))
     
     TemResult = TemResult & res
Next i

Or_28 = BinStringtoDe(TemResult)
End Function

'*************************************************************************
'**函 数 名：Not_28
'**输    入：(Double)Num_De
'**输    出：(String) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-11-25 17:16:46
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function Not_28(ByVal Num_De As Double) As Double
Dim sB As String * 29, i As Integer, K As String, res As String, TemResult As String
sB = DetoBinString_28(Num_De)

For i = 1 To 29
          K = Mid(sB, i, 1)
     
     If K = "1" Then
     res = "0"
     Else
     res = "1"
     End If
     
     TemResult = TemResult & res
Next i

Not_28 = BinStringtoDe(TemResult)
End Function

'*************************************************************************
'**函 数 名：RepFlags
'**输    入：(String)Num_De,(Integer)Bit,(Integer)Rep
'**输    出：(String) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-11-30 11:54:40
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function RepFlags(Num_De As String, Bit As Integer, Rep As Integer) As String
Dim tI As Integer64b, tB As String, tH As String

tI = StrToI64(Num_De)
tB = I64ToBinStr(tI)
tB = ReplaceBinStr(tB, Bit, Rep)
tH = BinToHex(tB)
tI = HexStrToI64(tH)
RepFlags = I64toStrNZ(tI)

End Function



'*************************************************************************
'**函 数 名：DeleteFlagsBinStr
'**输    入：(String)Bin,(String)SetBin
'**输    出：(String) -
'**功能描述：删除二进制字符串的FLAGS
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-11-30 14:14:33
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function DeleteFlagsBinStr(Bin As String, SetBin As String) As String

DeleteFlagsBinStr = BinStrAnd(Bin, BinStrNot(SetBin))

End Function

'*************************************************************************
'**函 数 名：DeleteFlagsI64
'**输    入：(integer64b)Flags,(integer64b)SetFlags
'**输    出：(integer64b) -
'**功能描述：删除二进制字符串的FLAGS
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-11-30 14:43:21
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function DeleteFlagsI64(Flags As Integer64b, SetFlags As Integer64b) As Integer64b
Dim Bin As String, SetBin As String, NewBin As String, tHex As String
Bin = I64ToBinStr(Flags)
SetBin = I64ToBinStr(SetFlags)

NewBin = BinStrAnd(Bin, BinStrNot(SetBin))
tHex = BinToHex(NewBin)
DeleteFlagsI64 = HexStrToI64(tHex)

End Function

'*************************************************************************
'**函 数 名：AddFlagsI64
'**输    入：(integer64b)Flags,(integer64b)SetFlags
'**输    出：(integer64b) -
'**功能描述：添加二进制字符串的FLAGS
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-11-30 14:57:10
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function AddFlagsI64(Flags As Integer64b, SetFlags As Integer64b) As Integer64b
Dim Bin As String, SetBin As String, NewBin As String, tHex As String
Bin = I64ToBinStr(Flags)
SetBin = I64ToBinStr(SetFlags)

NewBin = BinStrOr(Bin, SetBin)
tHex = BinToHex(NewBin)
AddFlagsI64 = HexStrToI64(tHex)

End Function
