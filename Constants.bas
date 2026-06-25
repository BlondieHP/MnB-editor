Attribute VB_Name = "Constants"

Option Explicit

Public Const itp_type_horse = &H1
Public Const itp_type_one_handed_wpn = &H2
Public Const itp_type_two_handed_wpn = &H3
Public Const itp_type_polearm = &H4
Public Const itp_type_arrows = &H5
Public Const itp_type_bolts = &H6
Public Const itp_type_shield = &H7
Public Const itp_type_bow = &H8
Public Const itp_type_crossbow = &H9
Public Const itp_type_thrown = &HA
Public Const itp_type_goods = &HB
Public Const itp_type_head_armor = &HC
Public Const itp_type_body_armor = &HD
Public Const itp_type_foot_armor = &HE
Public Const itp_type_hand_armor = &HF
Public Const itp_type_pistol = &H10
Public Const itp_type_musket = &H11
Public Const itp_type_bullets = &H12
Public Const itp_type_animal = &H13
Public Const itp_type_book = &H14
Public Const itp_type_mask = &H1F

Public Const itp_force_attach_left_hand = "0000000000000100"
Public Const itp_force_attach_right_hand = "0000000000000200"
Public Const itp_force_attach_left_forearm = "0000000000000300"
Public Const itp_attach_armature = "0000000000000f00"
Public Const itp_attachment_mask = "0000000000000f00"
Public Const itp_Attachment_Left_bit = 8
Public Const itp_Attachment_Right_bit = 9
Public Const itp_Attachment_Armature_bit1 = 10
Public Const itp_Attachment_Armature_bit2 = 11

Public Const itp_unique = 12
Public Const itp_always_loot = 13
Public Const itp_no_parry = 14
Public Const itp_default_ammo = 15
Public Const itp_merchandise = 16
Public Const itp_wooden_attack = 17
Public Const itp_wooden_parry = 18
Public Const itp_food = 19

Public Const itp_cant_reload_on_horseback = 20
Public Const itp_two_handed = 21
Public Const itp_primary = 22
Public Const itp_secondary = 23
Public Const itp_covers_legs = 24
Public Const itp_doesnt_cover_hair = 24
Public Const itp_can_penetrate_shield = 24
Public Const itp_consumable = 25
Public Const itp_bonus_against_shield = 26
Public Const itp_penalty_with_shield = 27
Public Const itp_cant_use_on_horseback = 28
Public Const itp_civilian = 29
Public Const itp_next_item_as_melee = 29
Public Const itp_fit_to_head = 30
Public Const itp_offset_lance = 30
Public Const itp_covers_head = 31
Public Const itp_couchable = 31
Public Const itp_crush_through = 32
Public Const itp_knock_back = 33
Public Const itp_remove_item_on_use = 34
Public Const itp_unbalanced = 35

Public Const itp_covers_beard = 36
Public Const itp_no_pick_up_from_ground = 37
Public Const itp_can_knock_down = 38

Public Const itcf_Thrust_onehanded = "0000000000000001"
Public Const itcf_Overswing_onehanded = "0000000000000002"
Public Const itcf_Slashright_onehanded = "0000000000000004"
Public Const itcf_Slashleft_onehanded = "0000000000000008"

Public Const itcf_Thrust_twohanded = "0000000000000010"
Public Const itcf_Overswing_twohanded = "0000000000000020"
Public Const itcf_Slashright_twohanded = "0000000000000040"
Public Const itcf_Slashleft_twohanded = "0000000000000080"

Public Const itcf_Thrust_polearm = "0000000000000100"
Public Const itcf_Overswing_polearm = "0000000000000200"
Public Const itcf_Slashright_polearm = "0000000000000400"
Public Const itcf_Slashleft_polearm = "0000000000000800"

Public Const itcf_Shoot_bow = "0000000000001000"
Public Const itcf_Shoot_javelin = "0000000000002000"
Public Const itcf_Shoot_crossbow = "0000000000004000"

Public Const itcf_Throw_stone = "0000000000010000"
Public Const itcf_Throw_knife = "0000000000020000"
Public Const itcf_Throw_axe = "0000000000030000"
Public Const itcf_Throw_javelin = "0000000000040000"
Public Const itcf_Shoot_pistol = "0000000000070000"
Public Const itcf_Shoot_musket = "0000000000080000"
Public Const itcf_Shoot_mask = "00000000000ff000"

Public Const itcf_Horseback_thrust_onehanded = "0000000000100000"
Public Const itcf_Horseback_overswing_right_onehanded = "0000000000200000"
Public Const itcf_Horseback_overswing_left_onehanded = "0000000000400000"
Public Const itcf_Horseback_slashright_onehanded = "0000000000800000"
Public Const itcf_Horseback_slashleft_onehanded = "0000000001000000"
Public Const itcf_Thrust_onehanded_lance = "0000000004000000"
Public Const itcf_Thrust_onehanded_lance_horseback = "0000000008000000"

Public Const itcf_Carry_mask = "00000007f0000000"
Public Const itcf_Carry_sword_left_hip = "0000000010000000"
Public Const itcf_Carry_axe_left_hip = "0000000020000000"
Public Const itcf_Carry_dagger_front_left = "0000000030000000"
Public Const itcf_Carry_dagger_front_right = "0000000040000000"
Public Const itcf_Carry_quiver_front_right = "0000000050000000"
Public Const itcf_Carry_quiver_back_right = "0000000060000000"
Public Const itcf_Carry_quiver_right_vertical = "0000000070000000"
Public Const itcf_Carry_quiver_back = "0000000080000000"
Public Const itcf_Carry_revolver_right = "0000000090000000"
Public Const itcf_Carry_pistol_front_left = "00000000a0000000"
Public Const itcf_Carry_bowcase_left = "00000000b0000000"
Public Const itcf_Carry_mace_left_hip = "00000000c0000000"

Public Const itcf_Carry_axe_back = "0000000100000000"
Public Const itcf_Carry_sword_back = "0000000110000000"
Public Const itcf_Carry_kite_shield = "0000000120000000"
Public Const itcf_Carry_round_shield = "0000000130000000"
Public Const itcf_Carry_buckler_left = "0000000140000000"
Public Const itcf_Carry_crossbow_back = "0000000150000000"
Public Const itcf_Carry_bow_back = "0000000160000000"
Public Const itcf_Carry_spear = "0000000170000000"
Public Const itcf_Carry_board_shield = "0000000180000000"

Public Const itcf_Carry_katana = "0000000210000000"
Public Const itcf_Carry_wakizashi = "0000000220000000"


Public Const itcf_Show_holster_when_drawn = "0000000800000000"

Public Const itcf_Reload_pistol = "0000007000000000"
Public Const itcf_Reload_musket = "0000008000000000"
Public Const itcf_Reload_mask = "000000f000000000"

Public Const itcf_Parry_forward_onehanded = "0000010000000000"
Public Const itcf_Parry_up_onehanded = "0000020000000000"
Public Const itcf_Parry_right_onehanded = "0000040000000000"
Public Const itcf_Parry_left_onehanded = "0000080000000000"

Public Const itcf_Parry_forward_twohanded = "0000100000000000"
Public Const itcf_Parry_up_twohanded = "0000200000000000"
Public Const itcf_Parry_right_twohanded = "0000400000000000"
Public Const itcf_Parry_left_twohanded = "0000800000000000"

Public Const itcf_Parry_forward_polearm = "0001000000000000"
Public Const itcf_Parry_up_polearm = "0002000000000000"
Public Const itcf_Parry_right_polearm = "0004000000000000"
Public Const itcf_Parry_left_polearm = "0008000000000000"

Public Const itcf_Horseback_slash_polearm = "0010000000000000"

Public Const itcf_Force_64_bits = "8000000000000000"

Public Const imodbits_polearm = "8202"
Public Const imodbits_axe = "262164"
Public Const imodbits_missile = "4398046511112"
Public Const imodbits_bow = "655370"
Public Const imodbits_crossbow = "131082"
Public Const imodbits_horse_basic = "41876193280"
Public Const imodbits_shield = "167772194"
Public Const imodbits_sword = "24596"
Public Const imodbits_cloth = "123731968"
Public Const imodbits_plate = "704643238"
Public Const imodbits_horse_good = "110595670016"
Public Const imodbits_sword_high = "155668"
Public Const imodbits_pick = "270356"
Public Const imodbits_mace = "262164"
Public Const imodbits_thrown = "4398046781448"
Public Const imodbits_thrown_minus_heavy = "4398046519304"
Public Const imodbits_armor = "704643236"

Public Const ixmesh_inventory = "1000000000000000"
Public Const ixmesh_flying_ammo = "2000000000000000"
Public Const ixmesh_Carry = "3000000000000000"
Public Const ixmesh_Inventory_bit = 60
Public Const ixmesh_Flying_Ammo_bit = 61

Public Const ti_on_init_item = -50           'can only be used in module_items triggers
Public Const ti_on_weapon_attack = -51       'can only be used in module_items triggers
' Position Register 1: Weapon Item Position
Public Const ti_on_missile_hit = -52         'can only be used in module_items triggers
' Position Register 1: Missile Position
' Trigger Param 1: shooter agent id
Public Const ti_on_init_map_icon = -70                   'can only be used in module_map_icons triggers
' Trigger Param 1: id of the owner party

'粒子系统
Public Const psf_always_emit = "0000000002"
Public Const psf_global_emit_dir = "0000000010"
Public Const psf_emit_at_water_level = "0000000020"
Public Const psf_billboard_2d = "0000000100"                  '# up_vec = dir, front rotated towards camera
Public Const psf_billboard_3d = "0000000200"                  '  # front_vec point to camera.
Public Const psf_billboard_drop = "0000000300"
Public Const psf_turn_to_velocity = "0000000400"
Public Const psf_randomize_rotation = "0000001000"
Public Const psf_randomize_size = "0000002000"
Public Const psf_2d_turbulance = "0000010000"

'网格模型
Public Const render_order_plus_1 = "00000001"

'Tags
Public Const Tag_Register = 1
Public Const Tag_Variable = 2
Public Const Tag_String = 3
Public Const Tag_Item = 4
Public Const Tag_Troop = 5
Public Const Tag_Faction = 6
Public Const Tag_Quest = 7
Public Const Tag_Party_Tpl = 8
Public Const Tag_Party = 9
Public Const Tag_Scene = 10
Public Const Tag_Mission_tpl = 11
Public Const Tag_Menu = 12
Public Const Tag_Script = 13
Public Const Tag_Particle_Sys = 14
Public Const Tag_Scene_Prop = 15
Public Const Tag_Sound = 16
Public Const Tag_Local_Variable = 17
Public Const Tag_Map_Icon = 18
Public Const Tag_Skill = 19
Public Const Tag_Mesh = 20
Public Const Tag_Presentation = 21
Public Const Tag_Quick_String = 22
Public Const Tag_Track = 23
Public Const Tag_Tableau = 24
Public Const Tag_Animation = 25
Public Const Tags_End = 26
Public Const N_Pos = 65

Public Const Max_Indentation = 16
Public Const EditInfo_TroopsCount = 0
Public Const EditInfo_ItemsCount = 1
Public Const EditInfo_ScenesCount = 2
Public Const EditInfo_FactionsCount = 3
Public Const EditInfo_PartyTemplatesCount = 4
Public Const EditInfo_PartiesCount = 5
Public Const EditInfo_MapIconsCount = 6
Public Const EditInfo_SoundsCount = 7
Public Const EditInfo_SoundRessCount = 8
Public Const EditInfo_PSysCount = 9
Public Const EditInfo_TabMatCount = 10
Public Const EditInfo_MeshCount = 11
Public Const EditInfo_TimeTrgCount = 12
Public Const EditInfo_StringsCount = 13
