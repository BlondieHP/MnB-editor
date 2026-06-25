Attribute VB_Name = "ModOperation"
Option Explicit

Public Type Type_Para
    Value As String
    Para_Type As String
End Type

Public Type Type_Operation
     OpID As Long
     Op_name As String
     Op_CSVname As String
     Pseudo As String     'Add
     
     ParaNum As Integer
     Para() As Type_Para
     Type As Integer
End Type


'-----------------OperationTypeConsts---------------------
Public Const OPT_NONE = 0
Public Const OPT_Lhs = 1
Public Const OPT_Global_Lhs = 2
Public Const OPT_Can_Fail = 3
'---------------------------------------------------------

Public Max_Op_Len As Long
Public Operation() As Type_Operation

'-------------OperationGroups------------------
Public ControlOperationGroup(6) As Integer    '流程控制
Public OptionalParamGroup(1) As Integer       '具可选参数
'----------------------------------------------

Public ChangeTag As String         'TV当前选中项的类型,以帮助ChangeFrm确定封装类型

'--------------------------------------------------------------------------
' CONTROL OPERATIONS
'--------------------------------------------------------------------------
Public Const Call_Script = 1       ' (call_script,<script_id>), @
Public Const end_try = 3           ' deprecated, use try_end instead
Public Const try_end = 3           ' (try_end), @
Public Const try_begin = 4         ' (try_begin), @
Public Const else_try_begin = 5    ' deprecated, use else_try instead
Public Const else_try = 5          ' (else_try), @

Public Const try_for_range = 6     ' Works like a for loop from lower-bound up to (upper-bound - 1) @
              ' (try_for_range,<destination>,<lower_bound>,<upper_bound>),

Public Const try_for_range_backwards = 7 ' Same as above but starts from (upper-bound - 1) down-to lower bound. @
                ' (try_for_range_backwards,<destination>,<upper_bound>,<lower_bound>),
Public Const try_for_parties = 11            ' (try_for_parties,<destination>), @
Public Const try_for_agents = 12         ' (try_for_agents,<destination>), @

Public Const store_script_param_1 = 21       ' (store_script_param_1,<destination>),  --(Within a script) stores the first script parameter.@
Public Const store_script_param_2 = 22       ' (store_script_param_2,<destination>),  --(Within a script) stores the second script parameter.@
Public Const store_script_param = 23         ' (store_script_param,<destination>,<script_param_no>), --(Within a script) stores <script_param_no>th script parameter.@

'--------------------------------------------------------------------------
' CONDITION OPERATIONS
'--------------------------------------------------------------------------

Public Const ge = 30            ' greater than or equal to -- (ge,<value>,<value>), @
Public Const eq = 31            ' equal to             -- (eq,<value>,<value>), @
Public Const gt = 32            ' greater than         -- (gt,<value>,<value>), @

Public Const is_between = 33    ' (is_between,<value>,<lower_bound>,<upper_bound>), 'greater than or equal to lower bound and less than upper bound @

Public Const entering_town = 36   ' (entering_town,<town_id>),@
Public Const map_free = 37         ' (map_free),@
Public Const encountered_party_is_attacker = 39      ' (encountered_party_is_attacker),@
Public Const conversation_screen_is_active = 42      ' (conversation_screen_active), 'used in mission template triggers only@

Public Const in_meta_mission = 44 ' deprecated, do not use.

Public Const set_player_troop = 47                ' (set_player_troop,<troop_id>),

Public Const store_repeat_object = 50              ' stores the index of a repeated dialog option for repeat_for_factions, etc...

Public Const set_result_string = 60                ' sets the result string for game scripts that need one (set_result_string, <string_id>),

Public Const key_is_down = 70                      ' fails if the key is not currently down (key_is_down, <key_id>),
Public Const key_clicked = 71                      ' fails if the key is not clicked on the specific frame (key_clicked, <key_id>),
Public Const game_key_is_down = 72                 ' fails if the game key is not currently down (key_is_down, <game_key_id>),
Public Const game_key_clicked = 73                 ' fails if the game key is not clicked on the specific frame (key_clicked, <game_key_id>),
Public Const mouse_get_position = 75                ' (mouse_get_position, <position_no>), 'x and y values of position are filled
Public Const omit_key_once = 77                    ' game omits any bound action for the key once (omit_key_once, <key_id>),
Public Const clear_omitted_keys = 78               ' (clear_omitted_keys),

Public Const get_global_cloud_amount = 90          ' (get_global_cloud_amount, <destination>), 'returns a value between 0-100
Public Const set_global_cloud_amount = 91          ' (set_global_cloud_amount, <value>), 'value is clamped to 0-100
Public Const get_global_haze_amount = 92           ' (get_global_haze_amount, <destination>), 'returns a value between 0-100
Public Const set_global_haze_amount = 93           ' (set_global_haze_amount, <value>), 'value is clamped to 0-100

Public Const hero_can_join = 101                   ' (hero_can_join, [party_id]),
Public Const hero_can_join_as_prisoner = 102       ' (hero_can_join_as_prisoner, [party_id]),
Public Const party_can_join = 103                  ' (party_can_join),
Public Const party_can_join_as_prisoner = 104      ' (party_can_join_as_prisoner),
Public Const troops_can_join = 105                 ' (troops_can_join,<value>),
Public Const troops_can_join_as_prisoner = 106     ' (troops_can_join_as_prisoner,<value>),
Public Const party_can_join_party = 107            ' (party_can_join_party, <joiner_party_id>, <host_party_id>,[flip_prisoners]),
Public Const main_party_has_troop = 110            ' (main_party_has_troop,<troop_id>),@
Public Const party_is_in_town = 130                ' (party_is_in_town,<party_id_1>,<party_id_2>),@
Public Const party_is_in_any_town = 131            ' (party_is_in_any_town,<party_id>),@
Public Const party_is_active = 132                 ' (party_is_active,<party_id>),@
Public Const player_has_item = 150                 ' (player_has_item,<item_id>),@
Public Const troop_has_item_equipped = 151         ' (troop_has_item_equipped,<troop_id>,<item_id>),@
Public Const troop_is_mounted = 152                ' (troop_is_mounted,<troop_id>),@
Public Const troop_is_guarantee_ranged = 153       ' (troop_is_guarantee_ranged, <troop_id>),@
Public Const troop_is_guarantee_horse = 154        ' (troop_is_guarantee_horse, <troop_id>),@

Public Const check_quest_active = 200              ' (check_quest_active,<quest_id>),@
Public Const check_quest_finished = 201            ' (check_quest_finished,<quest_id>),@
Public Const check_quest_succeeded = 202           ' (check_quest_succeeded,<quest_id>),@
Public Const check_quest_failed = 203              ' (check_quest_failed,<quest_id>),@
Public Const check_quest_concluded = 204           ' (check_quest_concluded,<quest_id>),@

Public Const is_trial_version = 250                ' (is_trial_version),

Public Const profile_get_banner_id = 350                ' (profile_get_banner_id, <destination>),
Public Const profile_set_banner_id = 351                ' (profile_set_banner_id, <value>),

Public Const get_achievement_stat = 370                 ' (get_achievement_stat, <destination>, <achievement_id>, <stat_index>),
Public Const set_achievement_stat = 371                 ' (set_achievement_stat, <achievement_id>, <stat_index>, <value>),
Public Const unlock_achievement = 372                   ' (unlock_achievement, <achievement_id>),

Public Const send_message_to_url = 380                  ' (send_message_to_url, <string_id>), 'result will be returned to script_game_receive_url_response

' multiplayer
Public Const multiplayer_send_message_to_server = 388   ' (multiplayer_send_int_to_server, <message_type>),
Public Const multiplayer_send_int_to_server = 389       ' (multiplayer_send_int_to_server, <message_type>, <value>),
Public Const multiplayer_send_2_int_to_server = 390     ' (multiplayer_send_2_int_to_server, <message_type>, <value>, <value>),
Public Const multiplayer_send_3_int_to_server = 391     ' (multiplayer_send_3_int_to_server, <message_type>, <value>, <value>, <value>),
Public Const multiplayer_send_4_int_to_server = 392     ' (multiplayer_send_4_int_to_server, <message_type>, <value>, <value>, <value>, <value>),
Public Const multiplayer_send_string_to_server = 393    ' (multiplayer_send_string_to_server, <message_type>, <string_id>),
Public Const multiplayer_send_message_to_player = 394   ' (multiplayer_send_message_to_player, <player_id>, <message_type>),
Public Const multiplayer_send_int_to_player = 395       ' (multiplayer_send_int_to_player, <player_id>, <message_type>, <value>),
Public Const multiplayer_send_2_int_to_player = 396     ' (multiplayer_send_2_int_to_player, <player_id>, <message_type>, <value>, <value>),
Public Const multiplayer_send_3_int_to_player = 397     ' (multiplayer_send_3_int_to_player, <player_id>, <message_type>, <value>, <value>, <value>),
Public Const multiplayer_send_4_int_to_player = 398     ' (multiplayer_send_4_int_to_player, <player_id>, <message_type>, <value>, <value>, <value>, <value>),
Public Const multiplayer_send_string_to_player = 399    ' (multiplayer_send_string_to_player, <player_id>, <message_type>, <string_id>),
Public Const get_max_players = 400                      ' (get_max_players, <destination>),
Public Const player_is_active = 401                     ' (player_is_active, <player_id>),
Public Const player_get_team_no = 402                   ' (player_get_team_no,  <destination>, <player_id>),
Public Const player_set_team_no = 403                   ' (player_get_team_no,  <destination>, <player_id>),
Public Const player_get_troop_id = 404                  ' (player_get_troop_id, <destination>, <player_id>),
Public Const player_set_troop_id = 405                  ' (player_get_troop_id, <destination>, <player_id>),
Public Const player_get_agent_id = 406                  ' (player_get_agent_id, <destination>, <player_id>),
Public Const player_get_gold = 407                      ' (player_get_gold, <destination>, <player_id>),
Public Const player_set_gold = 408                      ' (player_set_gold, <player_id>, <value>, <max_value>), 'set max_value to 0 if no limit is wanted
Public Const player_spawn_new_agent = 409               ' (player_spawn_new_agent, <player_id>),
Public Const player_add_spawn_item = 410                ' (player_add_spawn_item, <player_id>, <item_slot_no>, <item_id>),
Public Const multiplayer_get_my_team = 411              ' (multiplayer_get_my_team, <destination>),
Public Const multiplayer_get_my_troop = 412             ' (multiplayer_get_my_troop, <destination>),
Public Const multiplayer_set_my_troop = 413             ' (multiplayer_get_my_troop, <destination>),
Public Const multiplayer_get_my_gold = 414              ' (multiplayer_get_my_gold, <destination>),
Public Const multiplayer_get_my_player = 415            ' (multiplayer_get_my_player, <destination>),
Public Const multiplayer_clear_scene = 416              ' (multiplayer_clear_scene),
Public Const multiplayer_is_server = 417                ' (multiplayer_is_server),
Public Const multiplayer_is_dedicated_server = 418      ' (multiplayer_is_dedicated_server),
Public Const game_in_multiplayer_mode = 419             ' (game_in_multiplayer_mode),
Public Const multiplayer_make_everyone_enemy = 420      ' (multiplayer_make_everyone_enemy),
Public Const player_control_agent = 421                 ' (player_control_agent, <player_id>, <agent_id>),
Public Const player_get_item_id = 422                   ' (player_get_item_id, <destination>, <player_id>, <item_slot_no>) Only for server
Public Const player_get_banner_id = 423                 ' (player_get_banner_id, <destination>, <player_id>),
Public Const game_get_reduce_campaign_ai = 424          ' (game_get_reduce_campaign_ai, <destination>),
Public Const multiplayer_find_spawn_point = 425         ' (multiplayer_find_spawn_point, <destination>, <team_no>, <examine_all_spawn_points>, <is_horseman>),
Public Const set_spawn_effector_scene_prop_kind = 426   ' (set_spawn_effector_scene_prop_kind <team_no> <scene_prop_kind_no>)
Public Const set_spawn_effector_scene_prop_id = 427     ' (set_spawn_effector_scene_prop_id <scene_prop_id>)

Public Const player_is_admin = 430                      ' (player_is_admin, <player_id>),
Public Const player_get_score = 431                     ' (player_get_score, <destination>, <player_id>),
Public Const player_set_score = 432                     ' (player_set_score,<player_id>, <value>),
Public Const player_get_kill_count = 433                ' (player_get_kill_count, <destination>, <player_id>),
Public Const player_set_kill_count = 434                ' (player_set_kill_count,<player_id>, <value>),
Public Const player_get_death_count = 435               ' (player_get_death_count, <destination>, <player_id>),
Public Const player_set_death_count = 436               ' (player_set_death_count, <player_id>, <value>),
Public Const player_get_ping = 437                      ' (player_get_ping, <destination>, <player_id>),
Public Const player_is_busy_with_menus = 438            ' (player_is_busy_with_menus, <player_id>),
Public Const player_get_is_muted = 439                  ' (player_get_is_muted, <destination>, <player_id>),
Public Const player_set_is_muted = 440                  ' (player_set_is_muted, <player_id>, <value>),
Public Const player_get_unique_id = 441                 ' (player_get_unique_id, <destination>, <player_id>), 'can only bew used on server side

Public Const team_get_bot_kill_count = 450              ' (team_get_bot_kill_count, <destination>, <team_id>),
Public Const team_set_bot_kill_count = 451              ' (team_get_bot_kill_count, <destination>, <team_id>),
Public Const team_get_bot_death_count = 452             ' (team_get_bot_death_count, <destination>, <team_id>),
Public Const team_set_bot_death_count = 453             ' (team_get_bot_death_count, <destination>, <team_id>),
Public Const team_get_kill_count = 454                  ' (team_get_kill_count, <destination>, <team_id>),
Public Const team_get_score = 455                       ' (team_get_score, <destination>, <team_id>),
Public Const team_set_score = 456                       ' (team_set_score, <team_id>, <value>),
Public Const team_set_faction = 457                     ' (team_set_faction, <team_id>, <faction_id>),
Public Const team_get_faction = 458                     ' (team_get_faction, <destination>, <team_id>),
Public Const player_save_picked_up_items_for_next_spawn = 459  ' (player_save_picked_up_items_for_next_spawn, <player_id>),
Public Const player_get_value_of_original_items = 460   ' (player_get_value_of_original_items, <player_id>), 'this operation returns values of the items, but default troop items will be counted as zero (except horse)
Public Const player_item_slot_is_picked_up = 461        ' (player_item_slot_is_picked_up, <player_id>, <item_slot_no>), 'item slots are overriden when player picks up an item and stays alive until the next round

Public Const kick_player = 465                          ' (kick_player, <player_id>),
Public Const ban_player = 466                           ' (ban_player, <player_id>, <value>, <player_id>), 'set value = 1 for banning temporarily, assign 2nd player id as the administrator player id if banning is permanent
Public Const save_ban_info_of_player = 467              ' (save_ban_info_of_player, <player_id>),
Public Const ban_player_using_saved_ban_info = 468      ' (ban_player_using_saved_ban_info),

Public Const start_multiplayer_mission = 470            ' (start_multiplayer_mission, <mission_template_id>, <scene_id>, <started_manually>),

Public Const server_add_message_to_log = 473            ' (server_add_message_to_log, <string_id>),

Public Const server_get_renaming_server_allowed = 475   ' (server_get_renaming_server_allowed, <destination>), '0-1
Public Const server_get_changing_game_type_allowed = 476 ' (server_get_changing_game_type_allowed, <destination>), '0-1
''477 used for: server_set_anti_cheat                = 477 ' (server_set_anti_cheat, <value>), '0 = off, 1 = on
Public Const server_get_combat_speed = 478              ' (server_get_combat_speed, <destination>), '0-2
Public Const server_set_combat_speed = 479              ' (server_set_combat_speed, <value>), '0-2
Public Const server_get_friendly_fire = 480             ' (server_get_friendly_fire, <destination>),
Public Const server_set_friendly_fire = 481             ' (server_set_friendly_fire, <value>), '0 = off, 1 = on
Public Const server_get_control_block_dir = 482         ' (server_get_control_block_dir, <destination>),
Public Const server_set_control_block_dir = 483         ' (server_set_control_block_dir, <value>), '0 = automatic, 1 = by mouse movement
Public Const server_set_password = 484                  ' (server_set_password, <string_id>),
Public Const server_get_add_to_game_servers_list = 485  ' (server_get_add_to_game_servers_list, <destination>),
Public Const server_set_add_to_game_servers_list = 486  ' (server_set_add_to_game_servers_list, <value>),
Public Const server_get_ghost_mode = 487                ' (server_get_ghost_mode, <destination>),
Public Const server_set_ghost_mode = 488                ' (server_set_ghost_mode, <value>),
Public Const server_set_name = 489                      ' (server_set_name, <string_id>),
Public Const server_get_max_num_players = 490           ' (server_get_max_num_players, <destination>),
Public Const server_set_max_num_players = 491           ' (server_set_max_num_players, <value>),
Public Const server_set_welcome_message = 492           ' (server_set_welcome_message, <string_id>),
Public Const server_get_melee_friendly_fire = 493       ' (server_get_melee_friendly_fire, <destination>),
Public Const server_set_melee_friendly_fire = 494       ' (server_set_melee_friendly_fire, <value>), '0 = off, 1 = on
Public Const server_get_friendly_fire_damage_self_ratio = 495   ' (server_get_friendly_fire_damage_self_ratio, <destination>),
Public Const server_set_friendly_fire_damage_self_ratio = 496   ' (server_set_friendly_fire_damage_self_ratio, <value>), '0-100
Public Const server_get_friendly_fire_damage_friend_ratio = 497 ' (server_get_friendly_fire_damage_friend_ratio, <destination>),
Public Const server_set_friendly_fire_damage_friend_ratio = 498 ' (server_set_friendly_fire_damage_friend_ratio, <value>), '0-100
Public Const server_get_anti_cheat = 499                ' (server_get_anti_cheat, <destination>),
Public Const server_set_anti_cheat = 477                ' (server_set_anti_cheat, <value>), '0 = off, 1 = on

'' Set_slot operations. These assign a value to a slot.
Public Const troop_set_slot = 500                  ' (troop_set_slot,<troop_id>,<slot_no>,<value>),
Public Const party_set_slot = 501                  ' (party_set_slot,<party_id>,<slot_no>,<value>),
Public Const faction_set_slot = 502                ' (faction_set_slot,<faction_id>,<slot_no>,<value>),
Public Const scene_set_slot = 503                  ' (scene_set_slot,<scene_id>,<slot_no>,<value>),
Public Const party_template_set_slot = 504         ' (party_template_set_slot,<party_template_id>,<slot_no>,<value>),
Public Const agent_set_slot = 505                  ' (agent_set_slot,<agent_id>,<slot_no>,<value>),
Public Const quest_set_slot = 506                  ' (quest_set_slot,<quest_id>,<slot_no>,<value>),
Public Const item_set_slot = 507                   ' (item_set_slot,<item_id>,<slot_no>,<value>),
Public Const player_set_slot = 508                 ' (player_set_slot,<player_id>,<slot_no>,<value>),
Public Const team_set_slot = 509                   ' (team_set_slot,<team_id>,<slot_no>,<value>),
Public Const scene_prop_set_slot = 510             ' (scene_prop_set_slot,<scene_prop_instance_id>,<slot_no>,<value>),

'' Get_slot operations. These retrieve the value of a slot.
Public Const troop_get_slot = 520                  ' (troop_get_slot,<destination>,<troop_id>,<slot_no>),@
Public Const party_get_slot = 521                  ' (party_get_slot,<destination>,<party_id>,<slot_no>),@
Public Const faction_get_slot = 522                ' (faction_get_slot,<destination>,<faction_id>,<slot_no>),@
Public Const scene_get_slot = 523                  ' (scene_get_slot,<destination>,<scene_id>,<slot_no>),@
Public Const party_template_get_slot = 524         ' (party_template_get_slot,<destination>,<party_template_id>,<slot_no>),@
Public Const agent_get_slot = 525                  ' (agent_get_slot,<destination>,<agent_id>,<slot_no>),@
Public Const quest_get_slot = 526                  ' (quest_get_slot,<destination>,<quest_id>,<slot_no>),@
Public Const item_get_slot = 527                   ' (item_get_slot,<destination>,<item_id>,<slot_no>),@
Public Const player_get_slot = 528                 ' (player_get_slot,<destination>,<player_id>,<slot_no>),@
Public Const team_get_slot = 529                   ' (team_get_slot,<destination>,<player_id>,<slot_no>),@
Public Const scene_prop_get_slot = 530             ' (scene_prop_get_slot,<destination>,<scene_prop_instance_id>,<slot_no>),@

'' slot_eq operations. These check whether the value of a slot is equal to a given value.
Public Const troop_slot_eq = 540                   ' (troop_slot_eq,<troop_id>,<slot_no>,<value>),
Public Const party_slot_eq = 541                   ' (party_slot_eq,<party_id>,<slot_no>,<value>),
Public Const faction_slot_eq = 542                 ' (faction_slot_eq,<faction_id>,<slot_no>,<value>),
Public Const scene_slot_eq = 543                   ' (scene_slot_eq,<scene_id>,<slot_no>,<value>),
Public Const party_template_slot_eq = 544          ' (party_template_slot_eq,<party_template_id>,<slot_no>,<value>),
Public Const agent_slot_eq = 545                   ' (agent_slot_eq,<agent_id>,<slot_no>,<value>),
Public Const quest_slot_eq = 546                   ' (quest_slot_eq,<quest_id>,<slot_no>,<value>),
Public Const item_slot_eq = 547                    ' (item_slot_eq,<item_id>,<slot_no>,<value>),
Public Const player_slot_eq = 548                  ' (player_slot_eq,<player_id>,<slot_no>,<value>),
Public Const team_slot_eq = 549                    ' (team_slot_eq,<team_id>,<slot_no>,<value>),
Public Const scene_prop_slot_eq = 550              ' (scene_prop_slot_eq,<scene_prop_instance_id>,<slot_no>,<value>),

'' slot_ge operations. These check whether the value of a slot is greater than or equal to a given value.
Public Const troop_slot_ge = 560                   ' (troop_slot_ge,<troop_id>,<slot_no>,<value>),
Public Const party_slot_ge = 561                   ' (party_slot_ge,<party_id>,<slot_no>,<value>),
Public Const faction_slot_ge = 562                 ' (faction_slot_ge,<faction_id>,<slot_no>,<value>),
Public Const scene_slot_ge = 563                   ' (scene_slot_ge,<scene_id>,<slot_no>,<value>),
Public Const party_template_slot_ge = 564          ' (party_template_slot_ge,<party_template_id>,<slot_no>,<value>),
Public Const agent_slot_ge = 565                   ' (agent_slot_ge,<agent_id>,<slot_no>,<value>),
Public Const quest_slot_ge = 566                   ' (quest_slot_ge,<quest_id>,<slot_no>,<value>),
Public Const item_slot_ge = 567                    ' (item_slot_ge,<item_id>,<slot_no>,<value>),
Public Const player_slot_ge = 568                  ' (player_slot_ge,<player_id>,<slot_no>,<value>),
Public Const team_slot_ge = 569                    ' (team_slot_ge,<team_id>,<slot_no>,<value>),
Public Const scene_prop_slot_ge = 570              ' (scene_prop_slot_ge,<scene_prop_instance_id>,<slot_no>,<value>),

Public Const Play_Sound = 600                      ' (play_sound,<sound_id>,[options]),@
Public Const Play_Track = 601       '播放曲目      ' (play_track,<track_id>, [options]), ' 0 = default, 1 = fade out current track, 2 = stop current track@
Public Const Play_Cue_Track = 602   '播放提示曲目  ' (play_cue_track,<track_id>), 'starts immediately@
Public Const music_set_situation = 603             ' (music_set_situation, <situation_type>),
Public Const music_set_culture = 604               ' (music_set_culture, <culture_type>),
Public Const stop_all_sounds = 609                 ' (stop_all_sounds, [options]), ' 0 = default, 1 = fade out current track, 2 = stop current track

Public Const copy_position = 700                   ' copies position_no_2 to position_no_1
                      ' (copy_position,<position_no_1>,<position_no_2>),@
Public Const init_position = 701                   ' (init_position,<position_no>),@
Public Const get_trigger_object_position = 702     ' (get_trigger_object_position,<position_no>),

Public Const get_angle_between_positions = 705     ' (get_angle_between_positions, <destination_fixed_point>, <position_no_1>, <position_no_2>),@
Public Const position_has_line_of_sight_to_position = 707 ' (position_has_line_of_sight_to_position, <position_no_1>, <position_no_2>),@
Public Const get_distance_between_positions = 710  ' gets distance in centimeters. ' (get_distance_between_positions,<destination>,<position_no_1>,<position_no_2>),@
Public Const get_distance_between_positions_in_meters = 711  ' gets distance in meters. ' (get_distance_between_positions_in_meters,<destination>,<position_no_1>,<position_no_2>),@
Public Const get_sq_distance_between_positions = 712 ' gets squared distance in centimeters ' (get_sq_distance_between_positions,<destination>,<position_no_1>,<position_no_2>),@
Public Const get_sq_distance_between_positions_in_meters = 713 ' gets squared distance in meters ' (get_sq_distance_between_positions_in_meters,<destination>,<position_no_1>,<position_no_2>),@
Public Const position_is_behind_position = 714     ' (position_is_behind_position,<position_no_1>,<position_no_2>),
Public Const get_sq_distance_between_position_heights = 715 ' gets squared distance in centimeters ' (get_sq_distance_between_position_heights,<destination>,<position_no_1>,<position_no_2>),@

Public Const position_transform_position_to_parent = 716 ' (position_transform_position_to_parent,<dest_position_no>,<position_no>,<position_no_to_be_transformed>),
Public Const position_transform_position_to_local = 717  ' (position_transform_position_to_local, <dest_position_no>,<position_no>,<position_no_to_be_transformed>),

Public Const position_copy_rotation = 718          ' (position_copy_rotation,<position_no_1>,<position_no_2>), copies rotation of position_no_2 to position_no_1
Public Const position_copy_origin = 719            ' (position_copy_origin,<position_no_1>,<position_no_2>), copies origin of position_no_2 to position_no_1
Public Const position_move_x = 720                 ' movement is in cms, [0 = local; 1=global]
                      ' (position_move_x,<position_no>,<movement>,[value]),
Public Const position_move_y = 721                 ' (position_move_y,<position_no>,<movement>,[value]),
Public Const position_move_z = 722                 ' (position_move_z,<position_no>,<movement>,[value]),

Public Const position_rotate_x = 723               ' (position_rotate_x,<position_no>,<angle>),
Public Const position_rotate_y = 724               ' (position_rotate_y,<position_no>,<angle>),
Public Const position_rotate_z = 725               ' (position_rotate_z,<position_no>,<angle>),

Public Const position_get_x = 726                  ' (position_get_x,<destination_fixed_point>,<position_no>), 'x position in meters * fixed point multiplier is returned
Public Const position_get_y = 727                  ' (position_get_y,<destination_fixed_point>,<position_no>), 'y position in meters * fixed point multiplier is returned
Public Const position_get_z = 728                  ' (position_get_z,<destination_fixed_point>,<position_no>), 'z position in meters * fixed point multiplier is returned

Public Const position_set_x = 729                  ' (position_set_x,<position_no>,<value_fixed_point>), 'meters / fixed point multiplier is set
Public Const position_set_y = 730                  ' (position_set_y,<position_no>,<value_fixed_point>), 'meters / fixed point multiplier is set
Public Const position_set_z = 731                  ' (position_set_z,<position_no>,<value_fixed_point>), 'meters / fixed point multiplier is set

Public Const position_get_scale_x = 735            ' (position_get_scale_x,<destination_fixed_point>,<position_no>), 'x scale in meters * fixed point multiplier is returned
Public Const position_get_scale_y = 736            ' (position_get_scale_y,<destination_fixed_point>,<position_no>), 'y scale in meters * fixed point multiplier is returned
Public Const position_get_scale_z = 737            ' (position_get_scale_z,<destination_fixed_point>,<position_no>), 'z scale in meters * fixed point multiplier is returned

Public Const position_rotate_x_floating = 738      ' (position_rotate_x_floating,<position_no>,<angle>), 'angle in degree * fixed point multiplier
Public Const position_rotate_y_floating = 739      ' (position_rotate_y_floating,<position_no>,<angle>), 'angle in degree * fixed point multiplier

Public Const position_get_rotation_around_z = 740  ' (position_get_rotation_around_z,<destination>,<position_no>), 'rotation around z axis is returned as angle
Public Const position_normalize_origin = 741       ' (position_normalize_origin,<destination_fixed_point>,<position_no>),
                                                                      ' destination = convert_to_fixed_point(length(position.origin))
                                                                      ' position.origin *= 1/length(position.origin) 'so it normalizes the origin vector

Public Const position_get_screen_projection = 750  ' (position_get_screen_projection, <position_no_1>, <position_no_2>), returns screen projection of position_no_2 to position_no_1

Public Const position_set_z_to_ground_level = 791  ' (position_set_z_to_ground_level, <position_no>), Only works during a mission
Public Const position_get_distance_to_terrain = 792 ' (position_get_distance_to_terrain, <position_no>), Only works during a mission
Public Const position_get_distance_to_ground_level = 793 ' (position_get_distance_to_ground_level, <position_no>), Only works during a mission

Public Const start_presentation = 900                            ' (start_presentation, <presentation_id>),
Public Const start_background_presentation = 901               ' (start_background_presentation, <presentation_id>), 'can only be used in game menus
Public Const presentation_set_duration = 902                   ' (presentation_set_duration, <duration-in-1/100-seconds>), 'there must be an active presentation
Public Const is_presentation_active = 903                    ' (is_presentation_active, <presentation_id),
Public Const create_text_overlay = 910                       ' (create_text_overlay, <destination>, <string_id>), 'returns overlay id
Public Const create_mesh_overlay = 911                          ' (create_mesh_overlay, <destination>, <mesh_id>), 'returns overlay id
Public Const create_button_overlay = 912                     ' (create_button_overlay, <destination>, <string_id>), 'returns overlay id
Public Const create_image_button_overlay = 913               ' (create_image_button_overlay, <destination>, <mesh_id>, <mesh_id>), 'returns overlay id. second mesh is the pressed button mesh
Public Const create_slider_overlay = 914                     ' (create_slider_overlay, <destination>, <min_value>, <max_value>), 'returns overlay id
Public Const create_progress_overlay = 915                      ' (create_progress_overlay, <destination>, <min_value>, <max_value>), 'returns overlay id
Public Const create_combo_button_overlay = 916                 ' (create_combo_button_overlay, <destination>), 'returns overlay id
Public Const create_text_box_overlay = 917                   ' (create_text_box_overlay, <destination>), 'returns overlay id
Public Const create_check_box_overlay = 918                  ' (create_check_box_overlay, <destination>), 'returns overlay id
Public Const create_simple_text_box_overlay = 919            ' (create_simple_text_box_overlay, <destination>), 'returns overlay id
Public Const overlay_set_text = 920                                ' (overlay_set_text, <overlay_id>, <string_id>),
Public Const overlay_set_color = 921                           ' (overlay_set_color, <overlay_id>, <color>), 'color in RGB format like 0xRRGGBB (put hexadecimal values for RR GG and BB parts)
Public Const overlay_set_alpha = 922                         ' (overlay_set_alpha, <overlay_id>, <alpha>), 'alpha in A format like 0xAA (put hexadecimal values for AA part)
Public Const overlay_set_hilight_color = 923                   ' (overlay_set_hilight_color, <overlay_id>, <color>), 'color in RGB format like 0xRRGGBB (put hexadecimal values for RR GG and BB parts)
Public Const overlay_set_hilight_alpha = 924                 ' (overlay_set_hilight_alpha, <overlay_id>, <alpha>), 'alpha in A format like 0xAA (put hexadecimal values for AA part)
Public Const overlay_set_size = 925                                ' (overlay_set_size, <overlay_id>, <position_no>), 'position's x and y values are used
Public Const overlay_set_position = 926                        ' (overlay_set_position, <overlay_id>, <position_no>), 'position's x and y values are used
Public Const overlay_set_val = 927                             ' (overlay_set_val, <overlay_id>, <value>), 'can be used for sliders, combo buttons and check boxes
Public Const overlay_set_boundaries = 928                      ' (overlay_set_boundaries, <overlay_id>, <min_value>, <max_value>),
Public Const overlay_set_area_size = 929                           ' (overlay_set_area_size, <overlay_id>, <position_no>), 'position's x and y values are used
Public Const overlay_set_mesh_rotation = 930                 ' (overlay_set_mesh_rotation, <overlay_id>, <position_no>), 'position's rotation values are used for rotations around x, y and z axis
Public Const overlay_add_item = 931                             ' (overlay_add_item, <overlay_id>, <string_id>), ' adds an item to the combo box
Public Const overlay_animate_to_color = 932                  ' (overlay_animate_to_color, <overlay_id>, <duration-in-1/1000-seconds>, <color>), 'alpha value will not be used
Public Const overlay_animate_to_alpha = 933                     ' (overlay_animate_to_alpha, <overlay_id>, <duration-in-1/1000-seconds>, <color>), Only alpha value will be used
Public Const overlay_animate_to_highlight_color = 934        ' (overlay_animate_to_highlight_color, <overlay_id>, <duration-in-1/1000-seconds>, <color>), 'alpha value will not be used
Public Const overlay_animate_to_highlight_alpha = 935        ' (overlay_animate_to_highlight_alpha, <overlay_id>, <duration-in-1/1000-seconds>, <color>), Only alpha value will be used
Public Const overlay_animate_to_size = 936                      ' (overlay_animate_to_size, <overlay_id>, <duration-in-1/1000-seconds>, <position_no>), 'position's x and y values are used as
Public Const overlay_animate_to_position = 937                 ' (overlay_animate_to_position, <overlay_id>, <duration-in-1/1000-seconds>, <position_no>), 'position's x and y values are used as
Public Const create_image_button_overlay_with_tableau_material = 938 ' (create_image_button_overlay_with_tableau_material, <destination>, <mesh_id>, <tableau_material_id>, <value>), 'returns overlay id. value is passed to tableau_material
                                                        ' when mesh_id is -1, a default mesh is generated automatically
Public Const create_mesh_overlay_with_tableau_material = 939         ' (create_mesh_overlay_with_tableau_material, <destination>, <mesh_id>, <tableau_material_id>, <value>), 'returns overlay id. value is passed to tableau_material
                                                        ' when mesh_id is -1, a default mesh is generated automatically
Public Const create_game_button_overlay = 940                ' (create_game_button_overlay, <destination>, <string_id>), 'returns overlay id
Public Const create_in_game_button_overlay = 941             ' (create_in_game_button_overlay, <destination>, <string_id>), 'returns overlay id
Public Const create_number_box_overlay = 942                 ' (create_number_box_overlay, <destination>, <min_value>, <max_value>), 'returns overlay id
Public Const create_listbox_overlay = 943                    ' (create_list_box_overlay, <destination> 'returns overlay id
Public Const create_mesh_overlay_with_item_id = 944          ' (create_mesh_overlay_with_item_id, <destination>, <item_id>), 'returns overlay id.
Public Const set_container_overlay = 945                     ' (set_container_overlay, <overlay_id>), 'sets the container overlay that new overlays will attach to. give -1 to reset
Public Const overlay_get_position = 946                      ' (overlay_get_position, <destination>, <overlay_id>)
Public Const overlay_set_display = 947                       ' (overlay_set_display, <overlay_id>, <value>), 'shows/hides overlay (1 = show, 0 = hide)
Public Const create_combo_label_overlay = 948                ' (create_combo_label_overlay, <destination>), 'returns overlay id
Public Const overlay_obtain_focus = 949                      ' (overlay_obtain_focus, <overlay_id>), 'works for textboxes only

Public Const overlay_set_tooltip = 950                       ' (overlay_set_tooltip, <overlay_id>, <string_id>),

Public Const show_object_details_overlay = 960               ' (show_object_details_overlay, <value>), '0 = hide, 1 = show

Public Const show_item_details = 970      ' (show_item_details, <item_id>, <position_no>, <show_default_text_or_not>) 'show_default_text_or_not should be 1 for showing "default" for default item costs
Public Const close_item_details = 971     ' (close_item_details)

Public Const context_menu_add_item = 980       ' (right_mouse_menu_add_item, <string_id>, <value>), 'must be called only inside script_game_right_mouse_menu_get_buttons

Public Const get_average_game_difficulty = 990 ' (get_average_game_difficulty, <destination>),
Public Const get_level_boundary = 991 ' (get_level_boundary, <level_no>),


'-------------------------
' Mission Condition types
'-------------------------
Public Const all_enemies_defeated = 1003      ' (all_enemies_defeated),
Public Const race_completed_by_player = 1004  ' (race_completed_by_player),
Public Const num_active_teams_le = 1005       ' (num_active_teams_le,<value>),
Public Const main_hero_fallen = 1006          ' (main_hero_fallen),


'----------------------------
' NEGATIONS
'----------------------------
Public Const neg = "80000000"             ' (neg|<operation>),
Public Const this_or_next = "40000000"    ' (this_or_next|<operation>),
Public Const this_or_next_Offset = "400000"

'lt           = neg | ge ' less than     -- (lt,<value>,<value>),
'neq          = neg | eq ' not equal to      -- (neq,<value>,<value>),
'le           = neg | gt ' less or equal to  -- (le,<value>,<value>),
'Public Const lt = -30

'-------------------------------------------------------------------------------------------
' CONSEQUENCE OPERATIONS                                                                   -
'-------------------------------------------------------------------------------------------
Public Const finish_party_battle_mode = 1019        ' (finish_party_battle_mode),
Public Const set_party_battle_mode = 1020           ' (set_party_battle_mode),

Public Const set_camera_follow_party = 1021         ' (set_camera_follow_party,<party_id>), 'Works on map only.
Public Const start_map_conversation = 1025          ' (start_map_conversation,<troop_id>),
Public Const rest_for_hours = 1030                  ' (rest_for_hours,<rest_period>,[time_speed],[remain_attackable]),
Public Const rest_for_hours_interactive = 1031      ' (rest_for_hours_interactive,<rest_period>,[time_speed],[remain_attackable]),

Public Const add_xp_to_troop = 1062                 ' (add_xp_to_troop,<value>,[troop_id]),
Public Const add_gold_as_xp = 1063                  ' (add_gold_as_xp,<value>,[troop_id]),
Public Const add_xp_as_reward = 1064                ' (add_xp_as_reward,<value>),

Public Const add_gold_to_party = 1070               ' party_id should be different from 0
                   ' (add_gold_to_party,<value>,<party_id>),

Public Const set_party_creation_random_limits = 1080 ' (set_party_creation_random_limits, <min_value>, <max_value>), (values should be between 0, 100)

Public Const troop_set_note_available = 1095        ' (troop_set_note_available, <troop_id>, <value>), '1 = available, 0 = not available
Public Const faction_set_note_available = 1096      ' (faction_set_note_available, <faction_id>, <value>), '1 = available, 0 = not available
Public Const party_set_note_available = 1097        ' (party_set_note_available, <party_id>, <value>), '1 = available, 0 = not available
Public Const quest_set_note_available = 1098        ' (quest_set_note_available, <quest_id>, <value>), '1 = available, 0 = not available



'1090-1091-1092 is taken, see below (info_page)
Public Const spawn_around_party = 1100              ' ID of spawned party is put into reg(0)
                   ' (spawn_around_party,<party_id>,<party_template_id>),@
Public Const set_spawn_radius = 1103                ' (set_spawn_radius,<value>),@

Public Const display_debug_message = 1104           ' (display_debug_message,<string_id>,[hex_colour_code]), 'displays message only in debug mode, but writes to rgl_log.txt in both release and debug modes when edit mode is enabled
Public Const display_log_message = 1105             ' (display_log_message,<string_id>,[hex_colour_code]),
Public Const display_message = 1106                 ' (display_message,<string_id>,[hex_colour_code]),
Public Const set_show_messages = 1107               ' (set_show_messages,<value>), '0 disables window messages 1 re-enables them.

Public Const add_troop_note_tableau_mesh = 1108     ' (add_troop_note_tableau_mesh,<troop_id>,<tableau_material_id>),
Public Const add_faction_note_tableau_mesh = 1109   ' (add_faction_note_tableau_mesh,<faction_id>,<tableau_material_id>),
Public Const add_party_note_tableau_mesh = 1110     ' (add_party_note_tableau_mesh,<party_id>,<tableau_material_id>),
Public Const add_quest_note_tableau_mesh = 1111     ' (add_quest_note_tableau_mesh,<quest_id>,<tableau_material_id>),
Public Const add_info_page_note_tableau_mesh = 1090 ' (add_info_page_note_tableau_mesh,<info_page_id>,<tableau_material_id>),
Public Const add_troop_note_from_dialog = 1114      ' (add_troop_note_from_dialog,<troop_id>,<note_slot_no>, <value>), 'There are maximum of 8 slots. value = 1 -> shows when the note is added
Public Const add_faction_note_from_dialog = 1115    ' (add_faction_note_from_dialog,<faction_id>,<note_slot_no>, <value>), 'There are maximum of 8 slots value = 1 -> shows when the note is added
Public Const add_party_note_from_dialog = 1116      ' (add_party_note_from_dialog,<party_id>,<note_slot_no>, <value>), 'There are maximum of 8 slots value = 1 -> shows when the note is added
Public Const add_quest_note_from_dialog = 1112      ' (add_quest_note_from_dialog,<quest_id>,<note_slot_no>, <value>), 'There are maximum of 8 slots value = 1 -> shows when the note is added
Public Const add_info_page_note_from_dialog = 1091  ' (add_info_page_note_from_dialog,<info_page_id>,<note_slot_no>, <value>), 'There are maximum of 8 slots value = 1 -> shows when the note is added
Public Const add_troop_note_from_sreg = 1117        ' (add_troop_note_from_sreg,<troop_id>,<note_slot_no>,<string_id>, <value>), 'There are maximum of 8 slots value = 1 -> shows when the note is added
Public Const add_faction_note_from_sreg = 1118      ' (add_faction_note_from_sreg,<faction_id>,<note_slot_no>,<string_id>, <value>), 'There are maximum of 8 slots value = 1 -> shows when the note is added
Public Const add_party_note_from_sreg = 1119        ' (add_party_note_from_sreg,<party_id>,<note_slot_no>,<string_id>, <value>), 'There are maximum of 8 slots value = 1 -> shows when the note is added
Public Const add_quest_note_from_sreg = 1113        ' (add_quest_note_from_sreg,<quest_id>,<note_slot_no>,<string_id>, <value>), 'There are maximum of 8 slots value = 1 -> shows when the note is added
Public Const add_info_page_note_from_sreg = 1092    ' (add_info_page_note_from_sreg,<info_page_id>,<note_slot_no>,<string_id>, <value>), 'There are maximum of 8 slots value = 1 -> shows when the note is added

Public Const tutorial_box = 1120                    ' (tutorial_box,<string_id>,<string_id>), 'deprecated use dialog_box instead.@
Public Const dialog_box = 1120                      ' (tutorial_box,<text_string_id>,<title_string_id>),@
Public Const question_box = 1121                    ' (question_box,<string_id>, [<yes_string_id>], [<no_string_id>]),@
Public Const tutorial_message = 1122                ' (tutorial_message,<string_id>, <color>), 'set string_id = -1 for hiding the message@
Public Const tutorial_message_set_position = 1123   ' (tutorial_message_set_position, <position_x>, <position_y>),@
Public Const tutorial_message_set_size = 1124       ' (tutorial_message_set_size, <size_x>, <size_y>),@
Public Const tutorial_message_set_center_justify = 1125 ' (tutorial_message_set_center_justify, <val>), 'set not 0 for center justify, 0 for not center justify@
Public Const tutorial_message_set_background = 1126 ' (tutorial_message_set_background, <value>), '1 = on, 0 = off, default is off@

Public Const set_tooltip_text = 1130                '  (set_tooltip_text, <string_id>),


Public Const reset_price_rates = 1170               ' (reset_price_rates),
Public Const set_price_rate_for_item = 1171         ' (set_price_rate_for_item,<item_id>,<value_percentage>),
Public Const set_price_rate_for_item_type = 1172    ' (set_price_rate_for_item_type,<item_type_id>,<value_percentage>),

Public Const party_join = 1201                      ' (party_join),
Public Const party_join_as_prisoner = 1202          ' (party_join_as_prisoner),
Public Const troop_join = 1203                      ' (troop_join,<troop_id>),
Public Const troop_join_as_prisoner = 1204          ' (troop_join_as_prisoner,<troop_id>),

Public Const remove_member_from_party = 1210        ' (remove_member_from_party,<troop_id>,[party_id]),
Public Const remove_regular_prisoners = 1211        ' (remove_regular_prisoners,<party_id>),
Public Const remove_troops_from_companions = 1215   ' (remove_troops_from_companions,<troop_id>,<value>),
Public Const remove_troops_from_prisoners = 1216    ' (remove_troops_from_prisoners,<troop_id>,<value>),

Public Const heal_party = 1225                      ' (heal_party,<party_id>),

Public Const disable_party = 1230                   ' (disable_party,<party_id>),
Public Const enable_party = 1231                    ' (enable_party,<party_id>),
Public Const remove_party = 1232                    ' (remove_party,<party_id>),
Public Const add_companion_party = 1233             ' (add_companion_party,<troop_id_hero>),

Public Const add_troop_to_site = 1250               ' (add_troop_to_site,<troop_id>,<scene_id>,<entry_no>),
Public Const remove_troop_from_site = 1251          ' (remove_troop_from_site,<troop_id>,<scene_id>),
Public Const modify_visitors_at_site = 1261         ' (modify_visitors_at_site,<scene_id>),
Public Const reset_visitors = 1262                  ' (reset_visitors),
Public Const set_visitor = 1263                     ' (set_visitor,<entry_no>,<troop_id>,[<dna>]),
Public Const set_visitors = 1264                    ' (set_visitors,<entry_no>,<troop_id>,<number_of_troops>),
Public Const add_visitors_to_current_scene = 1265   ' (add_visitors_to_current_scene,<entry_no>,<troop_id>,<number_of_troops>, <team_no>, <group_no>), 'team no and group no are used in multiplayer mode only. default team in entry is used in single player mode
Public Const scene_set_day_time = 1266              ' (scene_set_day_time, <value>), 'value in hours (0-23), must be called within ti_before_mission_start triggers

Public Const set_relation = 1270                    ' (set_relation,<faction_id>,<faction_id>,<value>),
Public Const faction_set_name = 1275                ' (faction_set_name, <faction_id>, <string_id>),
Public Const faction_set_color = 1276               ' (faction_set_color, <faction_id>, <value>),
Public Const faction_get_color = 1277               ' (faction_get_color, <color>, <faction_id>)

'Quest stuff
Public Const start_quest = 1280              ' (start_quest,<quest_id>),
Public Const complete_quest = 1281           ' (complete_quest,<quest_id>),
Public Const succeed_quest = 1282            ' (succeed_quest,<quest_id>), 'also concludes the quest
Public Const fail_quest = 1283               ' (fail_quest,<quest_id>), 'also concludes the quest
Public Const cancel_quest = 1284             ' (cancel_quest,<quest_id>),

Public Const set_quest_progression = 1285    ' (set_quest_progression,<quest_id>,<value>),

Public Const conclude_quest = 1286           ' (conclude_quest,<quest_id>),

Public Const setup_quest_text = 1290         ' (setup_quest_text,<quest_id>),
Public Const setup_quest_giver = 1291        ' (setup_quest_giver,<quest_id>, <string_id>),


'encounter outcomes.
Public Const start_encounter = 1300            ' (start_encounter,<party_id>),
Public Const leave_encounter = 1301            ' (leave_encounter),
Public Const encounter_attack = 1302           ' (encounter_attack),
Public Const select_enemy = 1303               ' (select_enemy,<value>),
Public Const set_passage_menu = 1304           ' (set_passage_menu,<value>),
Public Const auto_set_meta_mission_at_end_commited = 1305 ' (start_auto_encounter,<value>),

'simulate_battle            = 1305 ' (simulate_battle,<value>),
Public Const end_current_battle = 1307         ' (end_current_battle),



Public Const set_mercenary_source_party = 1320 ' selects party from which to buy mercenaries
                   ' (set_mercenary_source_party,<party_id>),
                   
Public Const set_merchandise_modifier_quality = 1490         ' Quality rate in percentage (average quality = 100),@
                        ' (set_merchandise_modifier_quality,<value>),@
Public Const set_merchandise_max_value = 1491        ' (set_merchandise_max_value,<value>),@
Public Const reset_item_probabilities = 1492             ' (reset_item_probabilities),@
Public Const set_item_probability_in_merchandise = 1493  ' (set_item_probability_in_merchandise,<itm_id>,<value>),@

'active Troop
'set_active_troop                       = 1050
Public Const troop_set_name = 1501                           ' (troop_set_name, <troop_id>, <string_no>),
Public Const troop_set_plural_name = 1502                    ' (troop_set_plural_name, <troop_id>, <string_no>),
Public Const troop_set_face_key_from_current_profile = 1503  ' (troop_set_face_key_from_current_profile, <troop_id>),
Public Const troop_set_type = 1505                           ' (troop_set_type,<troop_id>,<gender>),
Public Const troop_get_type = 1506                           ' (troop_get_type,<destination>,<troop_id>),
Public Const troop_is_hero = 1507                            ' (troop_is_hero,<troop_id>),
Public Const troop_is_wounded = 1508                         ' (troop_is_wounded,<troop_id>), Only for heroes!
Public Const troop_set_auto_equip = 1509                     ' (troop_set_auto_equip,<troop_id>,<value>),'disables otr enables auto-equipping
Public Const troop_ensure_inventory_space = 1510             ' (troop_ensure_inventory_space,<troop_id>,<value>),@
Public Const troop_sort_inventory = 1511                     ' (troop_sort_inventory,<troop_id>),@
Public Const troop_add_merchandise = 1512                    ' (troop_add_merchandise,<troop_id>,<item_type_id>,<value>),
Public Const troop_add_merchandise_with_faction = 1513       ' (troop_add_merchandise_with_faction,<troop_id>,<faction_id>,<item_type_id>,<value>), 'faction_id is given to check if troop is eligible to produce that item
Public Const troop_get_xp = 1515                             ' (troop_get_xp, <destination>, <troop_id>),
Public Const troop_get_class = 1516                          ' (troop_get_class, <destination>, <troop_id>),

Public Const troop_raise_attribute = 1520                    ' (troop_raise_attribute,<troop_id>,<attribute_id>,<value>),
Public Const troop_raise_skill = 1521                        ' (troop_raise_skill,<troop_id>,<skill_id>,<value>),
Public Const troop_raise_proficiency = 1522                  ' (troop_raise_proficiency,<troop_id>,<proficiency_no>,<value>),
Public Const troop_raise_proficiency_linear = 1523           ' raises weapon proficiencies linearly without being limited by weapon master skill
                        ' (troop_raise_proficiency,<troop_id>,<proficiency_no>,<value>),

Public Const troop_add_proficiency_points = 1525             ' (troop_add_proficiency_points,<troop_id>,<value>),
Public Const troop_add_gold = 1528                           ' (troop_add_gold,<troop_id>,<value>),
Public Const troop_remove_gold = 1529                        ' (troop_remove_gold,<troop_id>,<value>),
Public Const troop_add_item = 1530                           ' (troop_add_item,<troop_id>,<item_id>,[modifier]),
Public Const troop_remove_item = 1531                        ' (troop_remove_item,<troop_id>,<item_id>),
Public Const troop_clear_inventory = 1532                    ' (troop_clear_inventory,<troop_id>),
Public Const troop_equip_items = 1533                ' (troop_equip_items,<troop_id>), 'equips the items in the inventory automatically
Public Const troop_inventory_slot_set_item_amount = 1534     ' (troop_inventory_slot_set_item_amount,<troop_id>,<inventory_slot_no>,<value>),
Public Const troop_inventory_slot_get_item_amount = 1537     ' (troop_inventory_slot_get_item_amount,<destination>,<troop_id>,<inventory_slot_no>),
Public Const troop_inventory_slot_get_item_max_amount = 1538 ' (troop_inventory_slot_get_item_max_amount,<destination>,<troop_id>,<inventory_slot_no>),

Public Const troop_add_items = 1535                          ' (troop_add_items,<troop_id>,<item_id>,<number>),
Public Const troop_remove_items = 1536                       ' puts cost of items to reg0
                                                ' (troop_remove_items,<troop_id>,<item_id>,<number>),
Public Const troop_loot_troop = 1539                         ' (troop_loot_troop,<target_troop>,<source_troop_id>,<probability>),

Public Const troop_get_inventory_capacity = 1540             ' (troop_get_inventory_capacity,<destination>,<troop_id>),
Public Const troop_get_inventory_slot = 1541                 ' (troop_get_inventory_slot,<destination>,<troop_id>,<inventory_slot_no>),
Public Const troop_get_inventory_slot_modifier = 1542        ' (troop_get_inventory_slot_modifier,<destination>,<troop_id>,<inventory_slot_no>),
Public Const troop_set_inventory_slot = 1543                 ' (troop_set_inventory_slot,<troop_id>,<inventory_slot_no>,<value>),
Public Const troop_set_inventory_slot_modifier = 1544        ' (troop_set_inventory_slot_modifier,<troop_id>,<inventory_slot_no>,<value>),
Public Const troop_set_faction = 1550                      ' (troop_set_faction,<troop_id>,<faction_id>),
Public Const troop_set_age = 1555                          ' (troop_set_age, <troop_id>, <age_slider_pos>),  'Enter a value between 0..100
Public Const troop_set_health = 1560                         ' (troop_set_health,<troop_id>,<relative health (0-100)>),

Public Const troop_get_upgrade_troop = 1561                  ' (troop_get_upgrade_troop,<destination>,<troop_id>,<upgrade_path>), 'upgrade_path can be: 0 = get first node, 1 = get second node (returns -1 if not available)

'Items...
Public Const item_get_type = 1570                            ' (item_get_type, <destination>, <item_id>), 'returned values are listed at header_items.py (values starting with itp_type_)

'Parties...
Public Const party_get_num_companions = 1601                 ' (party_get_num_companions,<destination>,<party_id>),@
Public Const party_get_num_prisoners = 1602                  ' (party_get_num_prisoners,<destination>,<party_id>),@
Public Const party_set_flags = 1603                          ' (party_set_flag, <party_id>, <flag>, <clear_or_set>), 'sets flags like pf_default_behavior. see header_parties.py for flags.@
Public Const party_set_marshall = 1604                       ' (party_set_marshall, <party_id>, <value>)@
Public Const party_set_extra_text = 1605                     ' (party_set_extra_text,<party_id>, <string>)@
Public Const party_set_aggressiveness = 1606                 ' (party_set_aggressiveness, <party_id>, <number>),@
Public Const party_set_courage = 1607                        ' (party_set_courage, <party_id>, <number>),@
Public Const party_get_current_terrain = 1608                ' (party_get_current_terrain,<destination>,<party_id>),@
Public Const party_get_template_id = 1609                    ' (party_get_template_id,<destination>,<party_id>),@

Public Const party_add_members = 1610                        ' (party_add_members,<party_id>,<troop_id>,<number>), 'returns number added in reg0@
Public Const party_add_prisoners = 1611                      ' (party_add_prisoners,<party_id>,<troop_id>,<number>),'returns number added in reg0@
Public Const party_add_leader = 1612                         ' (party_add_leader,<party_id>,<troop_id>,[<number>]),@
Public Const party_force_add_members = 1613                  ' (party_force_add_members,<party_id>,<troop_id>,<number>),@
Public Const party_force_add_prisoners = 1614                ' (party_force_add_prisoners,<party_id>,<troop_id>,<number>),@

Public Const party_remove_members = 1615                     ' stores number removed to reg0
                        ' (party_remove_members,<party_id>,<troop_id>,<number>),@
Public Const party_remove_prisoners = 1616                   ' stores number removed to reg0
                        ' (party_remove_members,<party_id>,<troop_id>,<number>),@
Public Const party_clear = 1617                              ' (party_clear,<party_id>),@
Public Const party_wound_members = 1618                      ' (party_wound_members,<party_id>,<troop_id>,<number>),@
Public Const party_remove_members_wounded_first = 1619       ' stores number removed to reg0
                        ' (party_remove_members_wounded_first,<party_id>,<troop_id>,<number>),@

Public Const party_set_faction = 1620                        ' (party_set_faction,<party_id>,<faction_id>),@
Public Const party_relocate_near_party = 1623                ' (party_relocate_near_party,<party_id>,<target_party_id>,<value_spawn_radius>),@

Public Const party_get_position = 1625                       ' (party_get_position,<position_no>,<party_id>),@
Public Const party_set_position = 1626                       ' (party_set_position,<party_id>,<position_no>),@
Public Const map_get_random_position_around_position = 1627  ' (map_get_random_position_around_position,<dest_position_no>,<source_position_no>,<radius>),@
Public Const map_get_land_position_around_position = 1628    ' (map_get_land_position_around_position,<dest_position_no>,<source_position_no>,<radius>),@
Public Const map_get_water_position_around_position = 1629   ' (map_get_water_position_around_position,<dest_position_no>,<source_position_no>,<radius>),@


Public Const party_count_members_of_type = 1630              ' (party_count_members_of_type,<destination>,<party_id>,<troop_id>),@
Public Const party_count_companions_of_type = 1631           ' (party_count_companions_of_type,<destination>,<party_id>,<troop_id>),@
Public Const party_count_prisoners_of_type = 1632            ' (party_count_prisoners_of_type,<destination>,<party_id>,<troop_id>),@

Public Const party_get_free_companions_capacity = 1633       ' (party_get_free_companions_capacity,<destination>,<party_id>),
Public Const party_get_free_prisoners_capacity = 1634        ' (party_get_free_prisoners_capacity,<destination>,<party_id>),

Public Const party_get_ai_initiative = 1638                  ' (party_get_ai_initiative,<destination>,<party_id>), 'result is between 0-100@
Public Const party_set_ai_initiative = 1639                  ' (party_set_ai_initiative,<party_id>,<value>), 'value is between 0-100@
Public Const party_set_ai_behavior = 1640                    ' (party_set_ai_behavior,<party_id>,<ai_bhvr>),@
Public Const party_set_ai_object = 1641                      ' (party_set_ai_object,<party_id>,<party_id>),@
Public Const party_set_ai_target_position = 1642             ' (party_set_ai_target_position,<party_id>,<position_no>),@
Public Const party_set_ai_patrol_radius = 1643               ' (party_set_ai_patrol_radius,<party_id>,<radius_in_km>),@
Public Const party_ignore_player = 1644                      ' (party_ignore_player, <party_id>,<duration_in_hours>), 'don't pursue player party for this duration@
Public Const party_set_bandit_attraction = 1645              ' (party_set_bandit_attraction, <party_id>,<attaraction>), 'set how attractive a target the party is for bandits (0..100)@
Public Const party_get_helpfulness = 1646                    ' (party_get_helpfulness,<destination>,<party_id>),@
Public Const party_set_helpfulness = 1647                    ' (party_set_helpfulness, <party_id>, <number>), 'tendency to help friendly parties under attack. (0-10000, 100 default.)@

Public Const party_get_num_companion_stacks = 1650           ' (party_get_num_companion_stacks,<destination>,<party_id>),
Public Const party_get_num_prisoner_stacks = 1651            ' (party_get_num_prisoner_stacks, <destination>,<party_id>),
Public Const party_stack_get_troop_id = 1652                 ' (party_stack_get_troop_id,      <destination>,<party_id>,<stack_no>),
Public Const party_stack_get_size = 1653                     ' (party_stack_get_size,          <destination>,<party_id>,<stack_no>),
Public Const party_stack_get_num_wounded = 1654              ' (party_stack_get_num_wounded,   <destination>,<party_id>,<stack_no>),
Public Const party_stack_get_troop_dna = 1655                ' (party_stack_get_troop_dna,     <destination>,<party_id>,<stack_no>),
Public Const party_prisoner_stack_get_troop_id = 1656        ' (party_get_prisoner_stack_troop,<destination>,<party_id>,<stack_no>),
Public Const party_prisoner_stack_get_size = 1657            ' (party_get_prisoner_stack_size, <destination>,<party_id>,<stack_no>),
Public Const party_prisoner_stack_get_troop_dna = 1658       ' (party_prisoner_stack_get_troop_dna, <destination>,<party_id>,<stack_no>),

Public Const party_attach_to_party = 1660                    ' (party_attach_to_party, <party_id>, <party_id to attach to>),
Public Const party_detach = 1661                             ' (party_detach, <party_id>),
Public Const party_collect_attachments_to_party = 1662       ' (party_collect_attachments_to_party, <party_id>, <destination party_id>),
Public Const party_quick_attach_to_current_battle = 1663     ' (party_quick_attach_to_current_battle, <party_id>, <side (0:players side, 1:enemy side)>),

Public Const party_get_cur_town = 1665                       ' (party_get_cur_town, <destination>, <party_id>),

Public Const party_leave_cur_battle = 1666                   ' (party_leave_cur_battle, <party_id>),
Public Const party_set_next_battle_simulation_time = 1667    ' (party_set_next_battle_simulation_time,<party_id>,<next_simulation_time_in_hours>),

Public Const party_set_name = 1669                           ' (party_set_name, <party_id>, <string_no>),

Public Const party_add_xp_to_stack = 1670                    ' (party_add_xp_to_stack, <party_id>, <stack_no>, <xp_amount>),

Public Const party_get_morale = 1671                         ' (party_get_morale, <destination>,<party_id>),
Public Const party_set_morale = 1672                         ' (party_set_morale, <party_id>, <value>), 'value is clamped to range [0...100].

Public Const party_upgrade_with_xp = 1673                    ' (party_upgrade_with_xp, <party_id>, <xp_amount>, <upgrade_path>), 'upgrade_path can be:
                                                                                                                    '0 = choose random, 1 = choose first, 2 = choose second
Public Const party_add_xp = 1674                             ' (party_add_xp, <party_id>, <xp_amount>),

Public Const party_add_template = 1675                       ' (party_add_template, <party_id>, <party_template_id>, [reverse_prisoner_status]),

Public Const party_set_icon = 1676                           ' (party_set_icon, <party_id>, <map_icon_id>),
Public Const party_set_banner_icon = 1677                    ' (party_set_banner_icon, <party_id>, <map_icon_id>),
Public Const party_add_particle_system = 1678                ' (party_add_particle_system, <party_id>, <particle_system_id>),
Public Const party_clear_particle_systems = 1679             ' (party_clear_particle_systems, <party_id>),

Public Const party_get_battle_opponent = 1680                ' (party_get_battle_opponent, <destination>, <party_id>)
Public Const party_get_icon = 1681                           ' (party_get_icon, <destination>, <party_id>),

Public Const party_get_skill_level = 1685                    ' (party_get_skill_level, <destination>, <party_id>, <skill_no>),
Public Const get_battle_advantage = 1690                     ' (get_battle_advantage, <destination>),
Public Const set_battle_advantage = 1691                     ' (set_battle_advantage, <value>),

'1693 is used, agent_is_in_special_mode = 1693
Public Const party_get_attached_to = 1694                    ' (party_get_attached_to, <destination>, <party_id>),
Public Const party_get_num_attached_parties = 1695           ' (party_get_num_attached_parties, <destination>, <party_id>),
Public Const party_get_attached_party_with_rank = 1696       ' (party_get_attached_party_with_rank, <destination>, <party_id>, <attached_party_no>),
Public Const inflict_casualties_to_party_group = 1697        ' (inflict_casualties_to_party, <parent_party_id>, <attack_rounds>, <party_id_to_add_causalties_to>),
Public Const distribute_party_among_party_group = 1698       ' (distribute_party_among_party_group, <party_to_be_distributed>, <group_root_party>),
'1699 is used, agent_is_routed         = 1699

'Agents

'store_distance_between_positions,
'position_is_behind_poisiton,
Public Const get_player_agent_no = 1700                      ' (get_player_agent_no,<destination>),
Public Const get_player_agent_kill_count = 1701              ' (get_player_agent_kill_count,<destination>,[get_wounded]), 'Set second value to non-zero to get wounded count. returns lifetime kill counts
Public Const agent_is_alive = 1702                           ' (agent_is_alive,<agent_id>),
Public Const agent_is_wounded = 1703                         ' (agent_is_wounded,<agent_id>),
Public Const agent_is_human = 1704                           ' (agent_is_human,<agent_id>),
Public Const get_player_agent_own_troop_kill_count = 1705    ' (get_player_agent_own_troop_kill_count,<destination>,[get_wounded]), 'Set second value to non-zero to get wounded count
Public Const agent_is_ally = 1706                            ' (agent_is_ally,<agent_id>),
Public Const agent_is_non_player = 1707                      ' (agent_is_non_player, <agent_id>),
Public Const agent_is_defender = 1708                        ' (agent_is_defender,<agent_id>),
Public Const agent_is_active = 1712                          ' (agent_is_active,<agent_id>),
Public Const agent_is_routed = 1699                          ' (agent_is_routed,<agent_id>),
Public Const agent_is_in_special_mode = 1693                 ' (agent_is_in_special_mode,<agent_id>),

Public Const agent_get_look_position = 1709                  ' (agent_get_look_position, <position_no>, <agent_id>),@
Public Const agent_get_position = 1710                       ' (agent_get_position,<position_no>,<agent_id>),@
Public Const agent_set_position = 1711                       ' (agent_set_position,<agent_id>,<position_no>),@
Public Const agent_set_look_target_agent = 1713            ' (agent_set_look_target_agent, <agent_id>, <agent_id>), 'second agent_id is the target@
Public Const agent_get_horse = 1714                          ' (agent_get_horse,<destination>,<agent_id>),@
Public Const agent_get_rider = 1715                          ' (agent_get_rider,<destination>,<agent_id>),@
Public Const agent_get_party_id = 1716                       ' (agent_get_party_id,<destination>,<agent_id>),@
Public Const agent_get_entry_no = 1717                       ' (agent_get_entry_no,<destination>,<agent_id>),@
Public Const agent_get_troop_id = 1718                       ' (agent_get_troop_id,<destination>, <agent_id>),@
Public Const agent_get_item_id = 1719                        ' (agent_get_item_id,<destination>, <agent_id>), (works only for horses, returns -1 otherwise)@

Public Const store_agent_hit_points = 1720                   ' set absolute to 1 to retrieve actual hps, otherwise will return relative hp in range [0..100]
                        ' (store_agent_hit_points,<destination>,<agent_id>,[absolute]),@
Public Const agent_set_hit_points = 1721                     ' set absolute to 1 if value is absolute, otherwise value will be treated as relative number in range [0..100]
                        ' (agent_set_hit_points,<agent_id>,<value>,[absolute]),@
Public Const agent_deliver_damage_to_agent = 1722            ' (agent_deliver_damage_to_agent,<agent_id_deliverer>,<agent_id>,<value>), 'if value <= 0, then damage will be calculated using the weapon item@
Public Const agent_get_kill_count = 1723                     ' (agent_get_kill_count,<destination>,<agent_id>,[get_wounded]), 'Set second value to non-zero to get wounded count
Public Const agent_get_player_id = 1724                      ' (agent_get_player_id,<destination>,<agent_id>),
Public Const agent_set_invulnerable_shield = 1725          ' (agent_set_invulnerable_shield, <agent_id>),
Public Const agent_get_wielded_item = 1726                   ' (agent_get_wielded_item,<destination>,<agent_id>,<hand_no>),
Public Const agent_get_ammo = 1727                           ' (agent_get_ammo,<destination>,<agent_id>, <value>), 'value = 1 gets ammo for wielded item, value = 0 gets ammo for all items
Public Const agent_refill_ammo = 1728                        ' (agent_refill_ammo,<agent_id>),
Public Const agent_has_item_equipped = 1729                  ' (agent_has_item_equipped,<agent_id>,<item_id>),

Public Const agent_set_scripted_destination = 1730           ' (agent_set_scripted_destination,<agent_id>,<position_no>,<auto_set_z_to_ground_level>), 'auto_set_z_to_ground_level can be 0 (false) or 1 (true)
Public Const agent_get_scripted_destination = 1731           ' (agent_get_scripted_destination,<position_no>,<agent_id>),
Public Const agent_force_rethink = 1732                    ' (agent_force_rethink, <agent_id>),
Public Const agent_set_no_death_knock_down_only = 1733     ' (agent_set_no_death_knock_down_only, <agent_id>, <value>), '0 for disable, 1 for enable
Public Const agent_set_horse_speed_factor = 1734           ' (agent_set_horse_speed_factor, <agent_id>, <speed_multiplier-in-1/100>),
Public Const agent_clear_scripted_mode = 1735                ' (agent_clear_scripted_mode,<agent_id>),
Public Const agent_set_speed_limit = 1736                    ' (agent_set_speed_limit,<agent_id>,<speed_limit(kilometers/hour)>), 'Affects AI only
Public Const agent_ai_set_always_attack_in_melee = 1737      ' (agent_ai_set_always_attack_in_melee, <agent_id>,<value>), 'to be used in sieges so that agents don't wait on the ladder.
Public Const agent_get_simple_behavior = 1738                ' (agent_get_simple_behavior, <destination>, <agent_id>), 'constants are written in header_mission_templates.py, starting with aisb_
Public Const agent_get_combat_state = 1739                   ' (agent_get_combat_state, <destination>, <agent_id>),

Public Const agent_set_animation = 1740                      ' (agent_set_animation, <agent_id>, <anim_id>),
Public Const agent_set_stand_animation = 1741                ' (agent_set_stand_action, <agent_id>, <anim_id>),
Public Const agent_set_walk_forward_animation = 1742         ' (agent_set_walk_forward_action, <agent_id>, <anim_id>),
Public Const agent_set_animation_progress = 1743             ' (agent_set_animation_progress, <agent_id>, <value_fixed_point>), 'value should be between 0-1 (as fixed point)
Public Const agent_set_look_target_position = 1744           ' (agent_set_look_target_position, <agent_id>, <position_no>),
Public Const agent_set_attack_action = 1745                  ' (agent_set_attack_action, <agent_id>, <value>, <value>), 'value: 0 = thrust, 1 = slashright, 2 = slashleft, 3 = overswing - second value 0 = ready and release, 1 = ready and hold
Public Const agent_set_defend_action = 1746                  ' (agent_set_defend_action, <agent_id>, <value>, <duration-in-1/1000-seconds>), 'value_1: 0 = defend_down, 1 = defend_right, 2 = defend_left, 3 = defend_up
Public Const agent_set_wielded_item = 1747                   ' (agent_set_wielded_item, <agent_id>, <item_id>),
Public Const agent_set_scripted_destination_no_attack = 1748 ' (agent_set_scripted_destination_no_attack,<agent_id>,<position_no>,<auto_set_z_to_ground_level>), 'auto_set_z_to_ground_level can be 0 (false) or 1 (true)
Public Const agent_fade_out = 1749                           ' (agent_fade_out, <agent_id>),
Public Const agent_play_sound = 1750                         ' (agent_play_sound, <agent_id>, <sound_id>),
Public Const agent_start_running_away = 1751                 ' (agent_start_running_away, <agent_id>),
Public Const agent_stop_running_away = 1752                  ' (agent_stop_run_away, <agent_id>),
Public Const agent_ai_set_aggressiveness = 1753              ' (agent_ai_set_aggressiveness, <agent_id>, <value>), '100 is the default aggressiveness. higher the value, less likely to run back
Public Const agent_set_kick_allowed = 1754                   ' (agent_set_kick_allowed, <agent_id>, <value>), '0 for disable, 1 for allow

Public Const remove_agent = 1755                             ' (remove_agent, <agent_id>),

Public Const agent_get_attached_scene_prop = 1756            ' (agent_get_attached_scene_prop, <agent_id>, <scene_prop_id>)
Public Const agent_set_attached_scene_prop = 1757            ' (agent_set_attached_scene_prop, <destination>, <agent_id>)
Public Const agent_set_attached_scene_prop_x = 1758          ' (agent_set_attached_scene_prop_x, <agent_id>, <value>)
Public Const agent_set_attached_scene_prop_z = 1759          ' (agent_set_attached_scene_prop_z, <agent_id>, <value>)

Public Const agent_get_time_elapsed_since_removed = 1760     ' (agent_get_time_elapsed_since_dead, <destination>, <agent_id>),
Public Const agent_get_number_of_enemies_following = 1761    ' (agent_get_number_of_enemies_following, <destination>, <agent_id>),

Public Const agent_set_no_dynamics = 1762                    ' (agent_set_no_dynamics, <agent_id>, <value>), '0 = turn dynamics off, 1 = turn dynamics on (required for cut-scenes)

Public Const agent_get_attack_action = 1763                  ' (agent_get_attack_action, <destination>, <agent_id>), 'returned values: free = 0, readying_attack = 1, releasing_attack = 2, completing_attack_after_hit = 3, attack_parried = 4, reloading = 5, after_release = 6, cancelling_attack = 7
Public Const agent_get_defend_action = 1764                  ' (agent_get_defend_action, <destination>, <agent_id>), 'returned values: free = 0, parrying = 1, blocking = 2

Public Const agent_get_group = 1765                          ' (agent_get_group, <destination>, <agent_id>),
Public Const agent_set_group = 1766                          ' (agent_set_group, <agent_id>, <value>),

Public Const agent_get_action_dir = 1767                     ' (agent_get_action_dir, <destination>, <agent_id>), 'invalid = -1, down = 0, right = 1, left = 2, up = 3
Public Const agent_get_animation = 1768                      ' (agent_get_animation, <destination>, <agent_id>, <body_part), '0 = lower body part, 1 = upper body part
Public Const agent_is_in_parried_animation = 1769            ' (agent_is_in_parried_animation, <agent_id>),

Public Const agent_get_team = 1770                           ' (agent_get_team  ,<destination>, <agent_id>),
Public Const agent_set_team = 1771                           ' (agent_set_team  , <agent_id>, <value>),

Public Const agent_get_class = 1772                          ' (agent_get_class ,<destination>, <agent_id>),
Public Const agent_get_division = 1773                       ' (agent_get_division ,<destination>, <agent_id>),
Public Const agent_unequip_item = 1774                         ' (agent_unequip_item,<agent_id>,<item_id>),

Public Const class_is_listening_order = 1775                 ' (class_is_listening_order, <team_no>, <sub_class>),
Public Const agent_set_ammo = 1776                           ' (agent_set_ammo,<agent_id>,<item_id>,<value>), 'value = a number between 0 and maximum ammo

Public Const agent_add_offer_with_timeout = 1777             ' (agent_add_offer_with_timeout, <agent_id>, <agent_id>, <duration-in-1/1000-seconds>), 'second agent_id is offerer, 0 value for duration is an infinite offer
Public Const agent_check_offer_from_agent = 1778             ' (agent_check_offer_from_agent, <agent_id>, <agent_id>), 'second agent_id is offerer

Public Const agent_equip_item = 1779                           ' (agent_equip_item,<agent_id>,<item_id>), 'for weapons, agent needs to have an empty weapon slot

Public Const entry_point_get_position = 1780                 ' (entry_point_get_position, <position_no>, <entry_no>),
Public Const entry_point_set_position = 1781                 ' (entry_point_set_position, <entry_no>, <position_no>),
Public Const entry_point_is_auto_generated = 1782            ' (entry_point_is_auto_generated, <entry_no>),

Public Const team_get_hold_fire_order = 1784                 ' (team_get_hold_fire_order, <destination>, <team_no>, <sub_class>),
Public Const team_get_movement_order = 1785                  ' (team_get_movement_order, <destination>, <team_no>, <sub_class>),
Public Const team_get_riding_order = 1786                    ' (team_get_riding_order, <destination>, <team_no>, <sub_class>),
Public Const team_get_weapon_usage_order = 1787              ' (team_get_weapon_usage_order, <destination>, <team_no>, <sub_class>),
Public Const teams_are_enemies = 1788                        ' (teams_are_enemies, <team_no>, <team_no_2>),
Public Const team_give_order = 1790                          ' (team_give_order, <team_no>, <sub_class>, <order_id>),
Public Const team_set_order_position = 1791                  ' (team_set_order_position, <team_no>, <sub_class>, <position_no>),
Public Const team_get_leader = 1792                          ' (team_get_leader, <destination>, <team_no>),
Public Const team_set_leader = 1793                          ' (team_set_leader, <team_no>, <new_leader_agent_id>),
Public Const team_get_order_position = 1794                  ' (team_get_order_position, <position_no>, <team_no>, <sub_class>),
Public Const team_set_order_listener = 1795                  ' (team_set_order_listener, <team_no>, <sub_class>, <value>), 'merge with old listeners if value is non-zero 'clear listeners if sub_class is less than zero
Public Const team_set_relation = 1796                        ' (team_set_relation, <team_no>, <team_no_2>, <value>), ' -1 for enemy, 1 for friend, 0 for neutral

Public Const set_rain = 1797                                 ' (set_rain,<rain-type>,<strength>), (rain_type: 1= rain, 2=snow ; strength: 0 - 100)
Public Const set_fog_distance = 1798                         ' (set_fog_distance, <distance_in_meters>, [fog_color]),
Public Const get_scene_boundaries = 1799                     ' (get_scene_boundaries, <position_min>, <position_max>),

Public Const scene_prop_enable_after_time = 1800             ' (scene_prop_enable_after_time, <scene_prop_id>, <value>)
Public Const scene_prop_has_agent_on_it = 1801               ' (scene_prop_has_agent_on_it, <scene_prop_id>, <agent_id>)

Public Const agent_clear_relations_with_agents = 1802        ' (agent_clear_relations_with_agents, <agent_id>),
Public Const agent_add_relation_with_agent = 1803            ' (agent_add_relation_with_agent, <agent_id>, <agent_id>, <value>), '-1 = enemy, 0 = neutral (no friendly fire at all), 1 = ally

Public Const ai_mesh_face_group_show_hide = 1805             ' (ai_mesh_face_group_show_hide, <group_no>, <value>), ' 1 for enable, 0 for disable

Public Const agent_is_alarmed = 1806                         ' (agent_is_alarmed, <agent_id>),
Public Const agent_set_is_alarmed = 1807                     ' (agent_set_is_alarmed, <agent_id>, <value>), ' 1 for enable, 0 for disable

Public Const scene_prop_get_num_instances = 1810             ' (scene_prop_get_num_instances, <destination>, <scene_prop_id>),
Public Const scene_prop_get_instance = 1811                  ' (scene_prop_get_instance, <destination>, <scene_prop_id>, <instance_no>),
Public Const scene_prop_get_visibility = 1812                ' (scene_prop_get_visibility, <destination>, <scene_prop_id>),
Public Const scene_prop_set_visibility = 1813                ' (scene_prop_set_visibility, <destination>, <scene_prop_id>),
Public Const scene_prop_set_hit_points = 1814                ' (scene_prop_set_hit_points, <destination>, <scene_prop_id>),
Public Const scene_prop_get_hit_points = 1815                ' (scene_prop_get_hit_points, <scene_prop_id>, <value>),
Public Const scene_prop_get_max_hit_points = 1816            ' (scene_prop_get_max_hit_points, <scene_prop_id>, <value>),
Public Const scene_prop_get_team = 1817                      ' (scene_prop_get_team, <value>, <scene_prop_id>),
Public Const scene_prop_set_team = 1818                      ' (scene_prop_set_team, <scene_prop_id>, <value>),

Public Const scene_item_get_num_instances = 1830             ' (scene_item_get_num_instances, <destination>, <item_id>),
Public Const scene_item_get_instance = 1831                  ' (scene_item_get_instance, <destination>, <item_id>, <instance_no>),
Public Const scene_spawned_item_get_num_instances = 1832     ' (scene_spawned_item_get_num_instances, <destination>, <item_id>),
Public Const scene_spawned_item_get_instance = 1833          ' (scene_spawned_item_get_instance, <destination>, <item_id>, <instance_no>),
Public Const scene_allows_mounted_units = 1834               ' (scene_allows_mounted_units),

Public Const prop_instance_get_variation_id = 1840           ' (prop_instance_get_variation_id, <destination>, <scene_prop_id>),
Public Const prop_instance_get_variation_id_2 = 1841         ' (prop_instance_get_variation_id_2, <destination>, <scene_prop_id>),

Public Const prop_instance_get_position = 1850               ' (prop_instance_get_position, <position_no>, <scene_prop_id>),
Public Const prop_instance_get_starting_position = 1851      ' (prop_instance_get_starting_position, <position_no>, <scene_prop_id>),
Public Const prop_instance_get_scale = 1852                  ' (prop_instance_get_scale, <position_no>, <scene_prop_id>),
Public Const prop_instance_get_scene_prop_kind = 1853        ' (prop_instance_get_scene_prop_type, <destination>, <scene_prop_id>)

Public Const prop_instance_set_position = 1855               ' (prop_instance_set_position, <scene_prop_id>, position),
Public Const prop_instance_animate_to_position = 1860        ' (prop_instance_animate_to_position, <scene_prop_id>, position, <duration-in-1/100-seconds>),
Public Const prop_instance_stop_animating = 1861             ' (prop_instance_stop_animating, <scene_prop_id>),
Public Const prop_instance_is_animating = 1862               ' (prop_instance_is_animating, <destination>, <scene_prop_id>),
Public Const prop_instance_get_animation_target_position = 1863    ' (prop_instance_get_animation_target_position, <pos>, <scene_prop_id>)
Public Const prop_instance_enable_physics = 1864             ' (prop_instance_enable_physics, <scene_prop_id>, <value>) '0 for disable, 1 for enable
Public Const prop_instance_rotate_to_position = 1865         ' (prop_instance_rotate_to_position, <scene_prop_id>, position, <duration-in-1/100-seconds>, <total_rotate_angle>),
Public Const prop_instance_initialize_rotation_angles = 1866   ' (prop_instance_initialize_rotation_angles, <scene_prop_id>),
Public Const prop_instance_refill_hit_points = 1870        ' (prop_instance_refill_hit_points, <scene_prop_id>),

Public Const prop_instance_dynamics_set_properties = 1871  ' (prop_instance_dynamics_set_properties,<scene_prop_id>,mass_friction),
Public Const prop_instance_dynamics_set_velocity = 1872    ' (prop_instance_dynamics_set_velocity,<scene_prop_id>,linear_velocity),
Public Const prop_instance_dynamics_set_omega = 1873       ' (prop_instance_dynamics_set_omega,<scene_prop_id>,angular_velocity),
Public Const prop_instance_dynamics_apply_impulse = 1874   ' (prop_instance_dynamics_apply_impulse,<scene_prop_id>,impulse_force),

Public Const prop_instance_intersects_with_prop_instance = 1880 ' (prop_instance_intersects_with_prop_instance, <scene_prop_id>, <scene_prop_id>),


Public Const replace_scene_props = 1890                      ' (replace_scene_props, <old_scene_prop_id>,<new_scene_prop_id>),
Public Const replace_scene_items_with_scene_props = 1891     ' (replace_scene_items_with_scene_props, <old_item_id>,<new_scene_prop_id>),
'---------------------------
' Mission Consequence types
'---------------------------

Public Const set_mission_result = 1906                       ' (set_mission_result,<value>),
Public Const finish_mission = 1907                           ' (finish_mission),
Public Const jump_to_scene = 1910                            ' (jump_to_scene,<scene_id>,<entry_no>),
Public Const set_jump_mission = 1911                         ' (set_jump_mission,<mission_template_id>),
Public Const set_jump_entry = 1912                           ' (set_jump_entry,<entry_no>),
Public Const start_mission_conversation = 1920               ' (start_mission_conversation,<troop_id>),
Public Const add_reinforcements_to_entry = 1930              ' (add_reinforcements_to_entry,<mission_template_entry_no>,<value>),

Public Const mission_enable_talk = 1935                      ' (mission_enable_talk), 'can talk with troops during battles
Public Const mission_disable_talk = 1936                     ' (mission_disable_talk), 'disables talk option for the mission

Public Const mission_tpl_entry_set_override_flags = 1940     ' (mission_entry_set_override_flags, <mission_template_id>, <entry_no>, <value>),
Public Const mission_tpl_entry_clear_override_items = 1941   ' (mission_entry_clear_override_items, <mission_template_id>, <entry_no>),
Public Const mission_tpl_entry_add_override_item = 1942      ' (mission_entry_add_override_item, <mission_template_id>, <entry_no>, <item_kind_id>),

Public Const Set_Current_Color = 1950                        ' red, green, blue: a value of 255 means 100%
                              ' (set_current_color,<value>,<value>,<value>),
Public Const set_position_delta = 1955                       ' x, y, z
                                                  ' (set_position_delta,<value>,<value>,<value>),
Public Const add_point_light = 1960                          ' (add_point_light,[flicker_magnitude],[flicker_interval]), 'flicker_magnitude between 0 and 100, flicker_interval is in 1/100 seconds
Public Const add_point_light_to_entity = 1961                ' (add_point_light_to_entity,[flicker_magnitude],[flicker_interval]), 'flicker_magnitude between 0 and 100, flicker_interval is in 1/100 seconds
Public Const particle_system_add_new = 1965                  ' (particle_system_add_new,<par_sys_id>,[position_no]),
Public Const particle_system_emit = 1968                     ' (particle_system_emit,<par_sys_id>,<value_num_particles>,<value_period>),
Public Const particle_system_burst = 1969                    ' (particle_system_burst,<par_sys_id>,<position_no>,[percentage_burst_strength]),

Public Const set_spawn_position = 1970                       ' (set_spawn_position, <position_no>)
Public Const spawn_item = 1971                               ' (spawn_item, <item_kind_id>, <item_modifier>)
Public Const spawn_agent = 1972                              ' (spawn_agent,<troop_id>), (stores agent_id in reg0)
Public Const spawn_horse = 1973                              ' (spawn_horse,<item_kind_id>, <item_modifier>)  (stores agent_id in reg0)
Public Const spawn_scene_prop = 1974                         ' (spawn_scene_prop, <scene_prop_id>)  (stores prop_instance_id in reg0) not yet.
Public Const cur_tableau_add_tableau_mesh = 1980             ' (cur_tableau_add_tableau_mesh, <tableau_material_id>, <value>, <position_register_no>), 'value is passed to tableau_material
Public Const cur_item_set_tableau_material = 1981            ' (cur_item_set_tableu_material, <tableau_material_id>, <instance_code>), Only call inside ti_on_init_item in module_items
Public Const cur_scene_prop_set_tableau_material = 1982      ' (cur_scene_prop_set_tableau_material, <tableau_material_id>, <instance_code>), Only call inside ti_on_init_scene_prop in module_scene_props
Public Const cur_map_icon_set_tableau_material = 1983        ' (cur_map_icon_set_tableau_material, <tableau_material_id>, <instance_code>), Only call inside ti_on_init_map_icon in module_scene_props
Public Const cur_tableau_render_as_alpha_mask = 1984         ' (cur_tableau_render_as_alpha_mask)
Public Const cur_tableau_set_background_color = 1985         ' (cur_tableau_set_background_color, <value>),@
Public Const cur_agent_set_banner_tableau_material = 1986    ' (cur_agent_set_banner_tableau_material, <tableau_material_id>)
Public Const cur_tableau_set_ambient_light = 1987            ' (cur_tableau_set_ambient_light, <red_fixed_point>, <green_fixed_point>, <blue_fixed_point>),
Public Const cur_tableau_set_camera_position = 1988          ' (cur_tableau_set_camera_position, <position_no>),
Public Const cur_tableau_set_camera_parameters = 1989        ' (cur_tableau_set_camera_parameters, <is_perspective>, <camera_width_times_1000>, <camera_height_times_1000>, <camera_near_times_1000>, <camera_far_times_1000>),
Public Const cur_tableau_add_point_light = 1990              ' (cur_tableau_add_point_light, <map_icon_id>, <position_no>, <red_fixed_point>, <green_fixed_point>, <blue_fixed_point>),
Public Const cur_tableau_add_sun_light = 1991                ' (cur_tableau_add_sun_light, <map_icon_id>, <position_no>, <red_fixed_point>, <green_fixed_point>, <blue_fixed_point>),
Public Const cur_tableau_add_mesh = 1992                     ' (cur_tableau_add_mesh, <mesh_id>, <position_no>, <value_fixed_point>, <value_fixed_point>),
                                                ' first value fixed point is the scale factor, second value fixed point is alpha. use 0 for default values@
Public Const cur_tableau_add_mesh_with_vertex_color = 1993   ' (cur_tableau_add_mesh_with_vertex_color, <mesh_id>, <position_no>, <value_fixed_point>, <value_fixed_point>, <value>),
                                                ' first value fixed point is the scale factor, second value fixed point is alpha. value is vertex color. use 0 for default values. vertex_color has no default value.@

Public Const cur_tableau_add_map_icon = 1994                 ' (cur_tableau_add_map_icon, <map_icon_id>, <position_no>, <value_fixed_point>),
                                                ' value fixed point is the scale factor
                                                
Public Const cur_tableau_add_troop = 1995                    ' (cur_tableau_add_troop, <troop_id>, <position_no>, <animation_id>, <instance_no>), 'if instance_no value is 0 or less, then the face is not generated randomly (important for heroes)
Public Const cur_tableau_add_horse = 1996                    ' (cur_tableau_add_horse, <item_id>, <position_no>, <animation_id>),
Public Const cur_tableau_set_override_flags = 1997           ' (cur_tableau_set_override_flags, <value>),
Public Const cur_tableau_clear_override_items = 1998         ' (cur_tableau_clear_override_items),
Public Const cur_tableau_add_override_item = 1999            ' (cur_tableau_add_override_item, <item_kind_id>),
Public Const cur_tableau_add_mesh_with_scale_and_vertex_color = 2000   ' (cur_tableau_add_mesh_with_scale_and_vertex_color, <mesh_id>, <position_no>, <position_no>, <value_fixed_point>, <value>),
                                                ' second position_no is x,y,z scale factors (with fixed point values). value fixed point is alpha. value is vertex color. use 0 for default values. scale and vertex_color has no default values.
 
Public Const mission_cam_set_mode = 2001                     ' (mission_cam_set_mode, <mission_cam_mode>, <duration-in-1/1000-seconds>, <value>) ' when leaving manual mode, duration defines the animation time from the initial position to the new position. set as 0 for instant camera position update
                                                                                                                                    ' if value = 0, then camera velocity will be linear. else it will be non-linear
Public Const mission_get_time_speed = 2002                   ' (mission_get_time_speed, <destination_fixed_point>),
Public Const mission_set_time_speed = 2003                   ' (mission_set_time_speed, <value_fixed_point>) 'this works only when cheat mode is enabled
Public Const mission_time_speed_move_to_value = 2004         ' (mission_speed_move_to_value, <value_fixed_point>, <duration-in-1/1000-seconds>) 'this works only when cheat mode is enabled
Public Const mission_set_duel_mode = 2006                    ' (mission_set_duel_mode, <value>), 'value: 0 = off, 1 = on

Public Const mission_cam_set_screen_color = 2008             '(mission_cam_set_screen_color, <value>), 'value is color together with alpha
Public Const mission_cam_animate_to_screen_color = 2009      '(mission_cam_animate_to_screen_color, <value>, <duration-in-1/1000-seconds>), 'value is color together with alpha

Public Const mission_cam_get_position = 2010                 ' (mission_cam_get_position, <position_register_no>)
Public Const mission_cam_set_position = 2011                 ' (mission_cam_set_position, <position_register_no>)
Public Const mission_cam_animate_to_position = 2012          ' (mission_cam_animate_to_position, <position_register_no>, <duration-in-1/1000-seconds>, <value>) ' if value = 0, then camera velocity will be linear. else it will be non-linear
Public Const mission_cam_get_aperture = 2013                 ' (mission_cam_get_aperture, <destination>)
Public Const mission_cam_set_aperture = 2014                 ' (mission_cam_set_aperture, <value>)
Public Const mission_cam_animate_to_aperture = 2015          ' (mission_cam_animate_to_aperture, <value>, <duration-in-1/1000-seconds>, <value>) ' if value = 0, then camera velocity will be linear. else it will be non-linear
Public Const mission_cam_animate_to_position_and_aperture = 2016   ' (mission_cam_animate_to_position_and_aperture, <position_register_no>, <value>, <duration-in-1/1000-seconds>, <value>) ' if value = 0, then camera velocity will be linear. else it will be non-linear
Public Const mission_cam_set_target_agent = 2017             ' (mission_cam_set_target_agent, <agent_id>, <value>) 'if value = 0 then do not use agent's rotation, else use agent's rotation
Public Const mission_cam_clear_target_agent = 2018           ' (mission_cam_clear_target_agent)
Public Const mission_cam_set_animation = 2019                ' (mission_cam_set_animation, <anim_id>),

Public Const talk_info_show = 2020                           ' (talk_info_show, <hide_or_show>) :0=hide 1=show
Public Const talk_info_set_relation_bar = 2021               ' (talk_info_set_relation_bar, <value>) :set relation bar to a value between -100 to 100, enter an invalid value to hide the bar.
Public Const talk_info_set_line = 2022                       ' (talk_info_set_line, <line_no>, <string_no>)

'mesh related
Public Const set_background_mesh = 2031                      ' (set_background_mesh, <mesh_id>),
Public Const set_game_menu_tableau_mesh = 2032               ' (set_game_menu_tableau_mesh, <tableau_material_id>, <value>, <position_register_no>), 'value is passed to tableau_material
                                                ' position contains the following information: x = x position of the mesh, y = y position of the mesh, z = scale of the mesh

'change_window types.
Public Const change_screen_return = 2040                     ' (change_screen_return),
Public Const change_screen_loot = 2041                       ' (change_screen_loot, <troop_id>),
Public Const change_screen_trade = 2042                      ' (change_screen_trade),
Public Const change_screen_exchange_members = 2043         ' (change_screen_exchange_members, [0,1 = exchange_leader], [party_id]), 'if party id is not given, current party will be used
Public Const change_screen_trade_prisoners = 2044            ' (change_screen_trade_prisoners),
Public Const change_screen_buy_mercenaries = 2045            ' (change_screen_buy_mercenaries),
Public Const change_screen_view_character = 2046             ' (change_screen_view_character),
Public Const change_screen_training = 2047                   ' (change_screen_training),
Public Const change_screen_mission = 2048                    ' (change_screen_mission),
Public Const change_screen_map_conversation = 2049           ' (change_screen_map_conversation),
Public Const change_screen_exchange_with_party = 2050        ' (change_screen_exchange_with_party, <party_id>),
Public Const change_screen_equip_other = 2051                ' (change_screen_equip_other, <troop_id>),
Public Const change_screen_map = 2052
Public Const change_screen_notes = 2053                      ' (change_screen_notes, <note_type>, <object_id>), 'Note type can be 1 = troops, 2 = factions, 3 = parties, 4 = quests, 5 = info_pages
Public Const change_screen_quit = 2055                       ' (change_screen_quit),
Public Const change_screen_give_members = 2056               ' (change_screen_give_members, [party_id]), 'if party id is not given, current party will be used
Public Const change_screen_controls = 2057                   ' (change_screen_controls),
Public Const change_screen_options = 2058                    ' (change_screen_options),


Public Const jump_to_menu = 2060                             ' (jump_to_menu,<menu_id>),
Public Const disable_menu_option = 2061                      ' (disable_menu_option),

Public Const store_trigger_param_1 = 2071   ' (store_trigger_param_1,<destination>),@
Public Const store_trigger_param_2 = 2072   ' (store_trigger_param_2,<destination>),@
Public Const store_trigger_param_3 = 2073   ' (store_trigger_param_3,<destination>),@
Public Const set_trigger_result = 2075      ' (set_trigger_result, <value>),

Public Const val_add = 2105                  'dest, operand ::       dest = dest + operand@
                ' (val_add,<destination>,<value>),
Public Const val_sub = 2106                  'dest, operand ::       dest = dest + operand@
                ' (val_sub,<destination>,<value>),
Public Const val_mul = 2107                  'dest, operand ::       dest = dest * operand@
                ' (val_mul,<destination>,<value>),
Public Const val_div = 2108                  'dest, operand ::       dest = dest / operand@
                ' (val_div,<destination>,<value>),
Public Const val_mod = 2109                  'dest, operand ::       dest = dest mod operand@
                ' (val_mod,<destination>,<value>),
Public Const val_min = 2110                  'dest, operand ::       dest = min(dest, operand)@
                ' (val_min,<destination>,<value>),
Public Const val_max = 2111                  'dest, operand ::       dest = max(dest, operand)@
                ' (val_max,<destination>,<value>),
Public Const val_clamp = 2112                'dest, operand ::       dest = max(min(dest,<upper_bound> - 1),<lower_bound>)@
                ' (val_clamp,<destination>,<lower_bound>, <upper_bound>),
Public Const val_abs = 2113                 'dest          ::       dest = abs(dest)@
                                ' (val_abs,<destination>),
Public Const val_or = 2114                   'dest, operand ::       dest = dest | operand@
                ' (val_or,<destination>,<value>),
Public Const val_and = 2115                  'dest, operand ::       dest = dest & operand@
                ' (val_and,<destination>,<value>),
Public Const store_or = 2116                 'dest, op1, op2 :      dest = op1 | op2@
                                ' (store_or,<destination>,<value>,<value>),
Public Const store_and = 2117                'dest, op1, op2 :      dest = op1 & op2@
                                ' (store_or,<destination>,<value>,<value>),

Public Const store_mod = 2119                'dest, op1, op2 :      dest = op1 % op2@
                ' (store_mod,<destination>,<value>,<value>),
Public Const store_add = 2120                'dest, op1, op2 :      dest = op1 + op2@
                ' (store_add,<destination>,<value>,<value>),
Public Const store_sub = 2121                'dest, op1, op2 :      dest = op1 - op2@
                ' (store_sub,<destination>,<value>,<value>),
Public Const store_mul = 2122                'dest, op1, op2 :      dest = op1 * op2@
                ' (store_mul,<destination>,<value>,<value>),
Public Const store_div = 2123                'dest, op1, op2 :      dest = op1 / op2@
                ' (store_div,<destination>,<value>,<value>),

Public Const set_fixed_point_multiplier = 2124      ' (set_fixed_point_multiplier, <value>),
                                        ' sets the precision of the values that are named as value_fixed_point or destination_fixed_point.
                                        ' Default is 1 (every fixed point value will be regarded as an integer)@

Public Const store_sqrt = 2125              ' (store_sqrt, <destination_fixed_point>, <value_fixed_point>), takes square root of the value@
Public Const store_pow = 2126               ' (store_pow, <destination_fixed_point>, <value_fixed_point>, <value_fixed_point), takes square root of the value@
                                'dest, op1, op2 :      dest = op1 ^ op2
Public Const store_sin = 2127               ' (store_sin, <destination_fixed_point>, <value_fixed_point>), takes sine of the value that is in degrees
Public Const store_cos = 2128               ' (store_cos, <destination_fixed_point>, <value_fixed_point>), takes cosine of the value that is in degrees
Public Const store_tan = 2129               ' (store_tan, <destination_fixed_point>, <value_fixed_point>), takes tangent of the value that is in degrees

Public Const convert_to_fixed_point = 2130  ' (convert_to_fixed_point, <destination_fixed_point>), multiplies the value with the fixed point multiplier
Public Const convert_from_fixed_point = 2131 ' (convert_from_fixed_point, <destination>), divides the value with the fixed point multiplier

Public Const assign = 2133                   ' had to put this here so that it can be called from conditions.@
                ' (assign,<destination>,<value>),
Public Const shuffle_range = 2134            ' (shuffle_range,<reg_no>,<reg_no>),

Public Const store_random = 2135             ' deprecated, use store_random_in_range instead.
Public Const store_random_in_range = 2136    ' gets random number in range [range_low,range_high] excluding range_high@
                ' (store_random_in_range,<destination>,<range_low>,<range_high>),
Public Const store_troop_gold = 2149         ' (store_troop_gold,<destination>,<troop_id>),@

Public Const store_num_free_stacks = 2154           ' (store_num_free_stacks,<destination>,<party_id>),
Public Const store_num_free_prisoner_stacks = 2155  ' (store_num_free_prisoner_stacks,<destination>,<party_id>),

Public Const store_party_size = 2156                 ' (store_party_size,<destination>,[party_id]),
Public Const store_party_size_wo_prisoners = 2157    ' (store_party_size_wo_prisoners,<destination>,[party_id]),
Public Const store_troop_kind_count = 2158          ' deprecated, use party_count_members_of_type instead
Public Const store_num_regular_prisoners = 2159      ' (store_mum_regular_prisoners,<destination>,<party_id>),

Public Const store_troop_count_companions = 2160     ' (store_troop_count_companions,<destination>,<troop_id>,[party_id]),
Public Const store_troop_count_prisoners = 2161      ' (store_troop_count_prisoners,<destination>,<troop_id>,[party_id]),
Public Const store_item_kind_count = 2165            ' (store_item_kind_count,<destination>,<item_id>,[troop_id]),

Public Const store_free_inventory_capacity = 2167    ' (store_free_inventory_capacity,<destination>,[troop_id]),

Public Const store_skill_level = 2170                ' (store_skill_level,<destination>,<skill_id>,[troop_id]),
Public Const store_character_level = 2171            ' (store_character_level,<destination>,[troop_id]),
Public Const store_attribute_level = 2172            ' (store_attribute_level,<destination>,<troop_id>,<attribute_id>),

Public Const store_troop_faction = 2173              ' (store_troop_faction,<destination>,<troop_id>),@
Public Const store_faction_of_troop = 2173           ' (store_troop_faction,<destination>,<troop_id>),@
Public Const store_troop_health = 2175               ' (store_troop_health,<destination>,<troop_id>,[absolute]),
                                        ' set absolute to 1 to get actual health; otherwise this will return percentage health in range (0-100)

Public Const store_proficiency_level = 2176          ' (store_proficiency_level,<destination>,<troop_id>,<attribute_id>),

                    ' (store_troop_health,<destination>,<troop_id>,[absolute]),
Public Const store_relation = 2190                   ' (store_relation,<destination>,<faction_id_1>,<faction_id_2>),
Public Const set_conversation_speaker_troop = 2197   ' (set_conversation_speaker_troop, <troop_id>),
Public Const set_conversation_speaker_agent = 2198   ' (set_conversation_speaker_troop, <agent_id>),
Public Const store_conversation_agent = 2199        ' (store_conversation_agent,<destination>),
Public Const store_conversation_troop = 2200        ' (store_conversation_troop,<destination>),
Public Const store_partner_faction = 2201           ' (store_partner_faction,<destination>),
Public Const store_encountered_party = 2202         ' (store_encountered_party,<destination>),
Public Const store_encountered_party2 = 2203        ' (store_encountered_party2,<destination>),
Public Const store_faction_of_party = 2204          ' (store_faction_of_party,<destination>),@
Public Const set_encountered_party = 2205           ' (set_encountered_party,<destination>),


'store_current_town              = 2210 ' deprecated, use store_current_scene instead
'store_current_site              = 2211 ' deprecated, use store_current_scene instead
Public Const store_current_scene = 2211             ' (store_current_scene,<destination>),

Public Const store_item_value = 2230                ' (store_item_value,<destination>,<item_id>),
Public Const store_troop_value = 2231               ' (store_troop_value,<destination>,<troop_id>),

Public Const store_partner_quest = 2240             ' (store_partner_quest,<destination>),
Public Const store_random_quest_in_range = 2250     ' (store_random_quest_in_range,<destination>,<lower_bound>,<upper_bound>),
Public Const store_random_troop_to_raise = 2251     ' (store_random_troop_to_raise,<destination>,<lower_bound>,<upper_bound>),
Public Const store_random_troop_to_capture = 2252    ' (store_random_troop_to_capture,<destination>,<lower_bound>,<upper_bound>),
Public Const store_random_party_in_range = 2254      ' (store_random_party_in_range,<destination>,<lower_bound>,<upper_bound>),
Public Const store01_random_parties_in_range = 2255 ' stores two random, different parties in a range to reg0 and reg1.
                    ' (store01_random_parties_in_range,<lower_bound>,<upper_bound>),
Public Const store_random_horse = 2257               ' (store_random_horse,<destination>)
Public Const store_random_equipment = 2258           ' (store_random_equipment,<destination>)
Public Const store_random_armor = 2259               ' (store_random_armor,<destination>)
Public Const store_quest_number = 2261              ' (store_quest_number,<destination>,<quest_id>),
Public Const store_quest_item = 2262                 ' (store_quest_item,<destination>,<item_id>),
Public Const store_quest_troop = 2263                ' (store_quest_troop,<destination>,<troop_id>),

Public Const store_current_hours = 2270             ' (store_current_hours,<destination>),
Public Const store_time_of_day = 2271                ' (store_time_of_day,<destination>),
Public Const store_current_day = 2272                ' (store_current_day,<destination>),
Public Const is_currently_night = 2273               ' (is_currently_night),

Public Const store_distance_to_party_from_party = 2281   ' (store_distance_to_party_from_party,<destination>,<party_id>,<party_id>),

Public Const get_party_ai_behavior = 2290                    ' (get_party_ai_behavior,<destination>,<party_id>),
Public Const get_party_ai_object = 2291                      ' (get_party_ai_object,<destination>,<party_id>),
Public Const party_get_ai_target_position = 2292             ' (party_get_ai_target_position,<position_no>,<party_id>),
Public Const get_party_ai_current_behavior = 2293           ' (get_party_ai_current_behavior,<destination>,<party_id>),
Public Const get_party_ai_current_object = 2294              ' (get_party_ai_current_object,<destination>,<party_id>),


Public Const store_num_parties_created = 2300                ' (store_num_parties_created,<destination>,<party_template_id>),
Public Const store_num_parties_destroyed = 2301              ' (store_num_parties_destroyed,<destination>,<party_template_id>),
Public Const store_num_parties_destroyed_by_player = 2302    ' (store_num_parties_destroyed_by_player,<destination>,<party_template_id>),


' Searching operations.
Public Const store_num_parties_of_template = 2310    ' (store_num_parties_of_template,<destination>,<party_template_id>),
Public Const store_random_party_of_template = 2311   ' fails if no party exists with tempolate_id (expensive)
                    ' (store_random_party_of_template,<destination>,<party_template_id>),

Public Const str_is_empty = 2318                    ' (str_is_empty, <string_register>),
Public Const str_clear = 2319                       ' (str_clear, <string_register>)
Public Const str_store_string = 2320                 ' (str_store_string,<string_register>,<string_id>),
Public Const str_store_string_reg = 2321             ' (str_store_string,<string_register>,<string_no>), 'copies one string register to another.
Public Const str_store_troop_name = 2322             ' (str_store_troop_name,<string_register>,<troop_id>),
Public Const str_store_troop_name_plural = 2323      ' (str_store_troop_name_plural,<string_register>,<troop_id>),
Public Const str_store_troop_name_by_count = 2324    ' (str_store_troop_name_by_count,<string_register>,<troop_id>,<number>),
Public Const str_store_item_name = 2325              ' (str_store_item_name,<string_register>,<item_id>),
Public Const str_store_item_name_plural = 2326       ' (str_store_item_name_plural,<string_register>,<item_id>),
Public Const str_store_item_name_by_count = 2327     ' (str_store_item_name_by_count,<string_register>,<item_id>),
Public Const str_store_party_name = 2330             ' (str_store_party_name,<string_register>,<party_id>),
Public Const str_store_agent_name = 2332             ' (str_store_agent_name,<string_register>,<agent_id>),
Public Const str_store_faction_name = 2335           ' (str_store_faction_name,<string_register>,<faction_id>),
Public Const str_store_quest_name = 2336             ' (str_store_quest_name,<string_register>,<quest_id>),
Public Const str_store_info_page_name = 2337         ' (str_store_info_page_name,<string_register>,<info_page_id>),
Public Const str_store_date = 2340                  ' (str_store_date,<string_register>,<number_of_hours_to_add_to_the_current_date>),
Public Const str_store_troop_name_link = 2341       ' (str_store_troop_name_link,<string_register>,<troop_id>),
Public Const str_store_party_name_link = 2342       ' (str_store_party_name_link,<string_register>,<party_id>),
Public Const str_store_faction_name_link = 2343     ' (str_store_faction_name_link,<string_register>,<faction_id>),
Public Const str_store_quest_name_link = 2344       ' (str_store_quest_name_link,<string_register>,<quest_id>),
Public Const str_store_info_page_name_link = 2345   ' (str_store_info_page_name_link,<string_register>,<info_page_id>),
Public Const str_store_class_name = 2346            ' (str_store_class_name,<stribg_register>,<class_id>)
Public Const str_store_player_username = 2350       ' (str_store_player_username,<string_register>,<player_id>), 'used in multiplayer mode only
Public Const str_store_server_password = 2351       ' (str_store_server_password, <string_register>),
Public Const str_store_server_name = 2352           ' (str_store_server_name, <string_register>),
Public Const str_store_welcome_message = 2353       ' (str_store_welcome_message, <string_register>),

'mission ones:
Public Const store_remaining_team_no = 2360          ' (store_remaining_team_no,<destination>),

Public Const store_mission_timer_a_msec = 2365   ' (store_mission_timer_a_msec,<destination>),
Public Const store_mission_timer_b_msec = 2366   ' (store_mission_timer_b_msec,<destination>),
Public Const store_mission_timer_c_msec = 2367   ' (store_mission_timer_c_msec,<destination>),

Public Const store_mission_timer_a = 2370    ' (store_mission_timer_a,<destination>),
Public Const store_mission_timer_b = 2371    ' (store_mission_timer_b,<destination>),
Public Const store_mission_timer_c = 2372    ' (store_mission_timer_c,<destination>),

Public Const reset_mission_timer_a = 2375    ' (reset_mission_timer_a),
Public Const reset_mission_timer_b = 2376    ' (reset_mission_timer_b),
Public Const reset_mission_timer_c = 2377    ' (reset_mission_timer_c),

Public Const store_enemy_count = 2380     ' (store_enemy_count,<destination>),
Public Const store_friend_count = 2381    ' (store_friend_count,<destination>),
Public Const store_ally_count = 2382      ' (store_ally_count,<destination>),
Public Const store_defender_count = 2383  ' (store_defender_count,<destination>),
Public Const store_attacker_count = 2384  ' (store_attacker_count,<destination>),
Public Const store_normalized_team_count = 2385 '(store_normalized_team_count,<destination>, <team_no>), 'Counts the number of agents belonging to a team
                                                                                            ' and normalizes the result regarding battle_size and advantage.
Public Const set_postfx = 2386
Public Const set_river_shader_to_mud = 2387 'changes river material for muddy env

'___________________________参数类型常量_______________________________________   'A_P

Public Const Tag_Mask_Hex = "00000000000000"

'______________________________________________________________________________
'*************************************************************************
'**函 数 名：InitOperations
'**输    入：无
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-12-13 22:29:35
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub InitOperations()
Dim i As Long

ReDim Operation(0 To 186)

With Operation(0)
    .Op_name = "call_script"
    .Op_CSVname = "运行脚本"
    .Pseudo = "运行<脚本>,<更多参数...>"
    .OpID = Call_Script
    .Type = OPT_NONE
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Script)
    .Para(1).Value = "脚本"
End With

With Operation(1)
    .Op_name = "play_sound"
    .Op_CSVname = "播放声音"
    .Pseudo = "播放<声音>"
    .OpID = Play_Sound
    .Type = OPT_NONE
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Sound)
    .Para(1).Value = "声音"
End With

With Operation(2)
    .Op_name = "play_track"       '(play_track,<track_id>, [options]), ' 0 = default, 1 = fade out current track, 2 = stop current track
    .Op_CSVname = "播放曲目"
    .Pseudo = "播放<曲目>,选项:<播放选项(可选参数)>"
    .OpID = Play_Track
    .Type = OPT_NONE
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Track)
    .Para(1).Value = "曲目"
    .Para(2).Para_Type = "po#"
    .Para(2).Value = "播放选项(可选参数)"
End With

With Operation(3)
    .Op_name = "play_cue_track"        ' (play_cue_track,<track_id>), 'starts immediately
    .Op_CSVname = "播放提示音乐"
    .Pseudo = "播放提示<曲目>"
    .OpID = Play_Cue_Track
    .Type = OPT_NONE
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Track)
    .Para(1).Value = "曲目"
End With

With Operation(4)
    .Op_name = "position_move_x"        ' movement is in cms, [0 = local; 1=global]' (position_move_x,<position_no>,<movement>,[value]),
    .Op_CSVname = "Position_Move_X"
    .Pseudo = "将<位置>沿 X 轴移动<位移量>,访问权限:<访问权限(可选参数)>"
    .OpID = position_move_x
    .Type = OPT_NONE
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "pos"
    .Para(1).Value = "位置"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "位移量"
    .Para(3).Para_Type = "ap#"
    .Para(3).Value = "访问权限(可选参数)"
End With

With Operation(5)
    .Op_name = "position_move_y"        ' movement is in cms, [0 = local; 1=global]' (position_move_x,<position_no>,<movement>,[value]),
    .Op_CSVname = "Position_Move_Y"
    .Pseudo = "将<位置>沿 Y 轴移动<位移量>,访问权限:<访问权限(可选参数)>"
    .OpID = position_move_y
    .Type = OPT_NONE
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "pos"
    .Para(1).Value = "位置"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "位移量"
    .Para(3).Para_Type = "ap#"
    .Para(3).Value = "访问权限(可选参数)"
End With

With Operation(6)
    .Op_name = "position_move_z"        ' movement is in cms, [0 = local; 1=global]' (position_move_x,<position_no>,<movement>,[value]),
    .Op_CSVname = "Position_Move_Z"
    .Pseudo = "将<位置>沿 Z 轴移动<位移量>,访问权限:<访问权限(可选参数)>"
    .OpID = position_move_z
    .Type = OPT_NONE
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "pos"
    .Para(1).Value = "位置"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "位移量"
    .Para(3).Para_Type = "ap#"
    .Para(3).Value = "访问权限(可选参数)"
End With

With Operation(7)
    .Op_name = "set_current_color"        ' red, green, blue: a value of 255 means 100%' (set_current_color,<value>,<value>,<value>),
    .Op_CSVname = "设置当前颜色"
    .Pseudo = "设置当前颜色(<红>,<绿>,<蓝>)"
    .OpID = Set_Current_Color
    .Type = OPT_NONE
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "0"
    .Para(1).Value = "红"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "绿"
    .Para(3).Para_Type = "0"
    .Para(3).Value = "蓝"
End With

With Operation(8)
    .Op_name = "set_position_delta"        ' x, y, z' (set_position_delta,<value>,<value>,<value>),
    .Op_CSVname = "Set_Position_Delta"
    .Pseudo = "设置位置(<X>,<Y>,<Z>)"
    .OpID = set_position_delta
    .Type = OPT_NONE
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "0"
    .Para(1).Value = "X"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "Y"
    .Para(3).Para_Type = "0"
    .Para(3).Value = "Z"
End With

With Operation(9)
    .Op_name = "add_point_light"       ' (add_point_light,[flicker_magnitude],[flicker_interval]), 'flicker_magnitude between 0 and 100, flicker_interval is in 1/100 seconds
    .Op_CSVname = "创建点光源"
    .Pseudo = "创建点光源(<闪烁幅度(0-100,可选参数)>,<闪烁间隔(1/100sec,可选参数)>)"
    .OpID = add_point_light
    .Type = OPT_NONE
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "0#"
    .Para(1).Value = "闪烁幅度(0-100,可选参数)"
    .Para(2).Para_Type = "0#"
    .Para(2).Value = "闪烁间隔(1/100sec,可选参数)"

End With

With Operation(10)
    .Op_name = "add_point_light_to_entity"       ' (add_point_light_to_entity,[flicker_magnitude],[flicker_interval]), 'flicker_magnitude between 0 and 100, flicker_interval is in 1/100 seconds
    .Op_CSVname = "创建点光源到实体"
    .Pseudo = "创建点光源到实体(<闪烁幅度(0-100,可选参数)>,<闪烁间隔(1/100sec,可选参数)>)"
    .OpID = add_point_light_to_entity
    .Type = OPT_NONE
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "0#"
    .Para(1).Value = "闪烁幅度(0-100,可选参数)"
    .Para(2).Para_Type = "0#"
    .Para(2).Value = "闪烁间隔(1/100sec,可选参数)"

End With

With Operation(11)
    .Op_name = "particle_system_add_new"       ' (particle_system_add_new,<par_sys_id>,[position_no]),
    .Op_CSVname = "粒子系统(新建)"
    .Pseudo = "在<位置(可选参数)>新建<粒子系统>"
    .OpID = particle_system_add_new
    .Type = OPT_NONE
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Particle_Sys)
    .Para(1).Value = "粒子系统"
    .Para(2).Para_Type = "0#"
    .Para(2).Value = "位置(可选参数)"

End With

With Operation(12)
    .Op_name = "particle_system_emit"      ' (particle_system_emit,<par_sys_id>,<value_num_particles>,<value_period>),
    .Op_CSVname = "粒子系统(emit)"
    .Pseudo = "在<时间>内喷射<粒子数量>个<粒子系统>"
    .OpID = particle_system_emit
    .Type = OPT_NONE
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Particle_Sys)
    .Para(1).Value = "粒子系统"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "粒子数量"
    .Para(3).Para_Type = "0"
    .Para(3).Value = "时间"
End With

With Operation(13)
    .Op_name = "particle_system_burst"     ' (particle_system_burst,<par_sys_id>,<position_no>,[percentage_burst_strength]),
    .Op_CSVname = "粒子系统(Burst)"
    .Pseudo = "在<位置>迸发<粒子数量(可选参数)>个<粒子系统>"
    .OpID = particle_system_burst
    .Type = OPT_NONE
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Particle_Sys)
    .Para(1).Value = "粒子系统"
    .Para(2).Para_Type = "pos"
    .Para(2).Value = "位置"
    .Para(3).Para_Type = "0#"
    .Para(3).Value = "粒子数量(可选参数)"
End With

With Operation(14)
    .Op_name = "store_trigger_param_1"    ' (store_trigger_param_1,<destination>),
    .Op_CSVname = "储存触发器参数1"
    .Pseudo = "<变量> = 触发器参数1"
    .OpID = store_trigger_param_1
    .Type = OPT_Lhs
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
End With

With Operation(15)
    .Op_name = "store_trigger_param_2"    ' (store_trigger_param_1,<destination>),
    .Op_CSVname = "储存触发器参数2"
    .Pseudo = "<变量> = 触发器参数2"
    .OpID = store_trigger_param_2
    .Type = OPT_Lhs
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
End With

With Operation(16)
    .Op_name = "store_trigger_param_3"    ' (store_trigger_param_1,<destination>),
    .Op_CSVname = "储存触发器参数3"
    .Pseudo = "<变量> = 触发器参数3"
    .OpID = store_trigger_param_3
    .Type = OPT_Lhs
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
End With

With Operation(17)
    .Op_name = "try_begin"
    .Op_CSVname = "Try_Begin"
    .Pseudo = "如果"
    .OpID = try_begin
    .Type = OPT_NONE
    .ParaNum = 0
    ReDim .Para(0)

End With

With Operation(18)
    .Op_name = "try_end"
    .Op_CSVname = "Try_End"
    .Pseudo = "结束判断/循环"
    .OpID = try_end
    .Type = OPT_NONE
    .ParaNum = 0
    ReDim .Para(0)

End With

With Operation(19)
    .Op_name = "else_try"
    .Op_CSVname = "Else_Try"
    .Pseudo = "否则"
    .OpID = else_try
    .Type = OPT_NONE
    .ParaNum = 0
    ReDim .Para(0)

End With

With Operation(20)
    .Op_name = "try_for_range"         ' Works like a for loop from lower-bound up to (upper-bound - 1)
                                       ' (try_for_range,<destination>,<lower_bound>,<upper_bound>),
    .Op_CSVname = "Try_for_Range"
    .Pseudo = "令<变量>从<下界>到<上界>循环"
    .OpID = try_for_range
    .Type = OPT_Lhs
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = ""
    .Para(2).Value = "下界"
    .Para(3).Para_Type = ""
    .Para(3).Value = "上界"
End With

With Operation(21)
    .Op_name = "try_for_range_backwards"         ' Same as above but starts from (upper-bound - 1) down-to lower bound.
                                                 ' (try_for_range_backwards,<destination>,<upper_bound>,<lower_bound>),
    .Op_CSVname = "Try_for_Range_Backwards"
    .Pseudo = "令<变量>从<上界>到<下界>逆向循环"
    .OpID = try_for_range_backwards
    .Type = OPT_Lhs
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = ""
    .Para(2).Value = "上界"
    .Para(3).Para_Type = ""
    .Para(3).Value = "下界"
    
End With

With Operation(22)
    .Op_name = "ge"                     ' greater than or equal to -- (ge,<value>,<value>),
    .Op_CSVname = "大于或等于"
    .Pseudo = "<逻辑符><值1>大于或等于<值2>"
    .OpID = ge
    .Type = OPT_Can_Fail
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ""
    .Para(1).Value = "值1"
    .Para(2).Para_Type = ""
    .Para(2).Value = "值2"
End With

With Operation(23)
    .Op_name = "eq"          ' equal to             -- (eq,<value>,<value>),
    .Op_CSVname = "等于"
    .Pseudo = "<逻辑符><值1>等于<值2>"
    .OpID = eq
    .Type = OPT_Can_Fail
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ""
    .Para(1).Value = "值1"
    .Para(2).Para_Type = ""
    .Para(2).Value = "值2"

End With

With Operation(24)
    .Op_name = "gt"         ' greater than         -- (gt,<value>,<value>),
    .Op_CSVname = "大于"
    .Pseudo = "<逻辑符><值1>大于<值2>"
    .OpID = gt
    .Type = OPT_Can_Fail
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ""
    .Para(1).Value = "值1"
    .Para(2).Para_Type = ""
    .Para(2).Value = "值2"
End With

With Operation(25)
    .Op_name = "is_between"           ' (is_between,<value>,<lower_bound>,<upper_bound>), 'greater than or equal to lower bound and less than upper bound
    .Op_CSVname = "在...之间"
    .Pseudo = "<逻辑符><值>在<下界>和<上界>之间"
    .OpID = is_between
    .Type = OPT_Can_Fail
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ""
    .Para(1).Value = "值"
    .Para(2).Para_Type = ""
    .Para(2).Value = "下界"
    .Para(3).Para_Type = ""
    .Para(3).Value = "上界"
End With

With Operation(26)
    .Op_name = "get_player_agent_no"        ' (get_player_agent_no,<destination>),
    .Op_CSVname = "获得玩家角色序号"
    .Pseudo = "<变量> = 玩家角色序号"
    .OpID = get_player_agent_no
    .Type = OPT_Lhs
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
End With

With Operation(27)
    .Op_name = "val_add"                  'dest, operand ::       dest = dest + operand
                                          ' (val_add,<destination>,<value>),
    .Op_CSVname = "相加+(值)"
    .Pseudo = "<变量>  = 变量 + <值>"
    .OpID = val_add
    .Type = OPT_Global_Lhs
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = ""
    .Para(2).Value = "值"
End With

With Operation(28)
    .Op_name = "val_sub"         'dest, operand ::       dest = dest + operand
                                 ' (val_sub,<destination>,<value>),
    .Op_CSVname = "相减-(值)"
    .Pseudo = "<变量>  = 变量 - <值>"
    .OpID = val_sub
    .Type = OPT_Global_Lhs
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = ""
    .Para(2).Value = "值"
End With

With Operation(29)
    .Op_name = "val_mul"        'dest, operand ::       dest = dest * operand
                                ' (val_mul,<destination>,<value>),
    .Op_CSVname = "相乘×(值)"
    .Pseudo = "<变量>  = 变量 × <值>"
    .OpID = val_mul
    .Type = OPT_Global_Lhs
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = ""
    .Para(2).Value = "值"
End With

With Operation(30)
    .Op_name = "val_div"        'dest, operand ::       dest = dest / operand
                                ' (val_div,<destination>,<value>),
    .Op_CSVname = "相除÷(值)"
    .Pseudo = "<变量>  = 变量 ÷ <值>"
    .OpID = val_div
    .Type = OPT_Global_Lhs
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = ""
    .Para(2).Value = "值"
End With

With Operation(31)
    .Op_name = "val_mod"        'dest, operand ::       dest = dest mod operand
                                ' (val_mod,<destination>,<value>),
    .Op_CSVname = "取模(值)"
    .Pseudo = "<变量>  = 变量 mod <值>"
    .OpID = val_mod
    .Type = OPT_Global_Lhs
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = ""
    .Para(2).Value = "值"
End With

With Operation(32)
    .Op_name = "val_min"        'dest, operand ::       dest = min(dest, operand)
                                ' (val_min,<destination>,<value>),
    .Op_CSVname = "最小(值)"
    .Pseudo = "<变量>  = 取变量和<值>中的较小数"
    .OpID = val_min
    .Type = OPT_Global_Lhs
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = ""
    .Para(2).Value = "值"
End With

With Operation(33)
    .Op_name = "val_max"        'dest, operand ::       dest = max(dest, operand)
                                ' (val_max,<destination>,<value>),
    .Op_CSVname = "最大(值)"
    .Pseudo = "<变量>  = 取变量和<值>中的较大数"
    .OpID = val_max
    .Type = OPT_Global_Lhs
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = ""
    .Para(2).Value = "值"
End With

With Operation(34)
    .Op_name = "val_clamp"        'dest, operand ::       dest = max(min(dest,<upper_bound> - 1),<lower_bound>)
                                  ' (val_clamp,<destination>,<lower_bound>, <upper_bound>),
    .Op_CSVname = "钳(值)"
    .Pseudo = "将<变量>的值修正到<下界>和<上界>之间"
    .OpID = val_clamp
    .Type = OPT_NONE
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = ""
    .Para(2).Value = "上界"
    .Para(3).Para_Type = ""
    .Para(3).Value = "下界"
End With

With Operation(35)
    .Op_name = "val_abs"         'dest          ::       dest = abs(dest)
                                ' (val_abs,<destination>),
    .Op_CSVname = "绝对值(值)"
    .Pseudo = "<变量> = |变量|"
    .OpID = val_abs
    .Type = OPT_NONE
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
End With

With Operation(36)
    .Op_name = "val_or"         'dest, operand ::       dest = dest | operand
                                ' (val_or,<destination>,<value>),
    .Op_CSVname = "或(值)"
    .Pseudo = "<变量> = 变量 或 <值>"
    .OpID = val_or
    .Type = OPT_NONE
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = ""
    .Para(2).Value = "值"
End With

With Operation(37)
    .Op_name = "val_and"         'dest, operand ::       dest = dest & operand
                                ' (val_and,<destination>,<value>),
    .Op_CSVname = "与(值)"
    .Pseudo = "<变量> = 变量 与 <值>"
    .OpID = val_and
    .Type = OPT_NONE
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = ""
    .Para(2).Value = "操作数"
End With

With Operation(38)
    .Op_name = "store_sqrt"        ' (store_sqrt, <destination_fixed_point>, <value_fixed_point>), takes square root of the value
    .Op_CSVname = "开根(储存)"
    .Pseudo = "<变量(浮点)> = 对<值(浮点)>开根"
    .OpID = store_sqrt
    .Type = OPT_Lhs
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量(浮点)"
    .Para(2).Para_Type = ""
    .Para(2).Value = "值(浮点)"
End With

With Operation(39)
    .Op_name = "store_or"         'dest, op1, op2 :      dest = op1 | op2
                                ' (store_or,<destination>,<value>,<value>),
    .Op_CSVname = "或(储存)"
    .Pseudo = "<变量> = <值1> 或 <值2>"
    .OpID = store_or
    .Type = OPT_Lhs
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = ""
    .Para(2).Value = "值1"
    .Para(3).Para_Type = ""
    .Para(3).Value = "值2"
End With

With Operation(40)
    .Op_name = "store_and"         'dest, op1, op2 :      dest = op1 & op2
                                ' (store_and,<destination>,<value>,<value>),
    .Op_CSVname = "与(储存)"
    .Pseudo = "<变量> = <值1> 与 <值2>"
    .OpID = store_and
    .Type = OPT_Lhs
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = ""
    .Para(2).Value = "值1"
    .Para(3).Para_Type = ""
    .Para(3).Value = "值2"
End With

With Operation(41)
    .Op_name = "store_mod"        'dest, op1, op2 :      dest = op1 % op2
                                  ' (store_mod,<destination>,<value>,<value>),
    .Op_CSVname = "取模(储存)"
    .Pseudo = "<变量> = <值1> mod <值2>"
    .OpID = store_mod
    .Type = OPT_Lhs
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = ""
    .Para(2).Value = "值1"
    .Para(3).Para_Type = ""
    .Para(3).Value = "值2"
End With

With Operation(42)
    .Op_name = "store_add"        'dest, op1, op2 :      dest = op1 + op2
                                  ' (store_add,<destination>,<value>,<value>),
    .Op_CSVname = "相加+(储存)"
    .Pseudo = "<变量> = <值1> + <值2>"
    .OpID = store_add
    .Type = OPT_Lhs
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = ""
    .Para(2).Value = "值1"
    .Para(3).Para_Type = ""
    .Para(3).Value = "值2"
End With

With Operation(43)
    .Op_name = "store_sub"        'dest, op1, op2 :      dest = op1 - op2
                                 ' (store_sub,<destination>,<value>,<value>),
    .Op_CSVname = "相减-(储存)"
    .Pseudo = "<变量> = <值1> - <值2>"
    .OpID = store_sub
    .Type = OPT_Lhs
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = ""
    .Para(2).Value = "值1"
    .Para(3).Para_Type = ""
    .Para(3).Value = "值2"
End With

With Operation(44)
    .Op_name = "store_mul"        'dest, op1, op2 :      dest = op1 * op2
                                 ' (store_mul,<destination>,<value>,<value>),
    .Op_CSVname = "相乘×(储存)"
    .Pseudo = "<变量> = <值1> × <值2>"
    .OpID = store_mul
    .Type = OPT_Lhs
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = ""
    .Para(2).Value = "值1"
    .Para(3).Para_Type = ""
    .Para(3).Value = "值2"
End With

With Operation(45)
    .Op_name = "store_div"       'dest, op1, op2 :      dest = op1 / op2
                                  ' (store_div,<destination>,<value>,<value>),
    .Op_CSVname = "相除÷(储存)"
    .Pseudo = "<变量> = <值1> ÷ <值2>"
    .OpID = store_div
    .Type = OPT_Lhs
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = ""
    .Para(2).Value = "值1"
    .Para(3).Para_Type = ""
    .Para(3).Value = "值2"
End With

With Operation(46)
    .Op_name = "cur_tableau_set_background_color"       ' (cur_tableau_set_background_color, <value>),
    .Op_CSVname = "当前可变素材设定背景色"
    .Pseudo = "将当前可变材质的背景色设定为<颜色值>"
    .OpID = cur_tableau_set_background_color
    .Type = OPT_NONE
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "0"
    .Para(1).Value = "颜色值"
End With

With Operation(47)
    .Op_name = "cur_tableau_add_mesh"       ' (cur_tableau_add_mesh, <mesh_id>, <position_no>, <value_fixed_point>, <value_fixed_point>),
                                                ' first value fixed point is the scale factor, second value fixed point is alpha. use 0 for default values
    .Op_CSVname = "当前可变素材增加网格"
    .Pseudo = "为当前可变材质在<位置>处添加<网格模型>,其尺寸为<尺寸比例系数(浮点)>,Alpha为<Alpha值(浮点)>"
    .OpID = cur_tableau_add_mesh
    .Type = OPT_NONE
    .ParaNum = 4
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Mesh)
    .Para(1).Value = "网格模型"
    .Para(2).Para_Type = "pos"
    .Para(2).Value = "位置"
    .Para(3).Para_Type = "0"
    .Para(3).Value = "尺寸比例系数(浮点)"
    .Para(4).Para_Type = "0"
    .Para(4).Value = "Alpha值(浮点)"
End With

With Operation(48)
    .Op_name = "cur_tableau_add_mesh_with_vertex_color"       ' (cur_tableau_add_mesh_with_vertex_color, <mesh_id>, <position_no>, <value_fixed_point>, <value_fixed_point>, <value>),
                                                ' first value fixed point is the scale factor, second value fixed point is alpha. value is vertex color. use 0 for default values. vertex_color has no default value.
    .Op_CSVname = "当前可变素材增加带顶点颜色的网格"
    .Pseudo = "为当前可变材质在<位置>处添加顶点颜色为<颜色值>的<网格模型>,其尺寸为<尺寸比例系数(浮点)>,Alpha为<Alpha值(浮点)>"
    .OpID = cur_tableau_add_mesh_with_vertex_color
    .Type = OPT_NONE
    .ParaNum = 5
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Mesh)
    .Para(1).Value = "网格模型"
    .Para(2).Para_Type = "pos"
    .Para(2).Value = "位置"
    .Para(3).Para_Type = "0"
    .Para(3).Value = "尺寸比例系数(浮点)"
    .Para(4).Para_Type = "0"
    .Para(4).Value = "Alpha值(浮点)"
    .Para(5).Para_Type = "0"
    .Para(5).Value = "颜色值"
End With

With Operation(49)
    .Op_name = "position_set_x"        ' (position_set_x,<position_no>,<value_fixed_point>), 'meters / fixed point multiplier is set
    .Op_CSVname = "Position_Set_X"
    .Pseudo = "设置<位置>的 X轴 坐标值为<值(浮点)>"
    .OpID = position_set_x
    .Type = OPT_NONE
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "pos"
    .Para(1).Value = "位置"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "值(浮点)"
End With

With Operation(50)
    .Op_name = "position_set_y"        ' (position_set_y,<position_no>,<value_fixed_point>), 'meters / fixed point multiplier is set
    .Op_CSVname = "Position_Set_Y"
    .Pseudo = "设置<位置>的 Y轴 坐标值为<值(浮点)>"
    .OpID = position_set_y
    .Type = OPT_NONE
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "pos"
    .Para(1).Value = "位置"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "值(浮点)"
End With

With Operation(51)
    .Op_name = "position_set_z"        ' (position_set_z,<position_no>,<value_fixed_point>), 'meters / fixed point multiplier is set
    .Op_CSVname = "Position_Set_Z"
    .Pseudo = "设置<位置>的 Z轴 坐标值为<值(浮点)>"
    .OpID = position_set_z
    .Type = OPT_NONE
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "pos"
    .Para(1).Value = "位置"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "值(浮点)"
End With

With Operation(52)
    .Op_name = "try_for_parties"        ' (try_for_parties,<destination>),
    .Op_CSVname = "Try_for_Parties"
    .Pseudo = "令<变量>在所有部队中循环"
    .OpID = try_for_parties
    .Type = OPT_Lhs
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
End With

With Operation(53)
    .Op_name = "try_for_agents"        ' (try_for_agents,<destination>),
    .Op_CSVname = "Try_for_Agents"
    .Pseudo = "令<变量>在所有角色中循环"
    .OpID = try_for_agents
    .Type = OPT_Lhs
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
End With

With Operation(54)
    .Op_name = "entering_town"        ' (entering_town,<town_id>),
    .Op_CSVname = "进入城镇"
    .Pseudo = "<逻辑符>进入<城镇>"
    .OpID = entering_town
    .Type = OPT_Can_Fail
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Party)
    .Para(1).Value = "城镇"
End With

With Operation(55)
    .Op_name = "map_free"       ' (map_free),
    .Op_CSVname = "Map_Free"
    .Pseudo = "<逻辑符>玩家可在大地图上自由移动,<值(可选)>"
    .OpID = map_free
    .Type = OPT_Can_Fail
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "0#"
    .Para(1).Value = "值(可选)"
End With

With Operation(56)
    .Op_name = "encountered_party_is_attacker"       ' (encountered_party_is_attacker),
    .Op_CSVname = "遭遇的部队是攻击方"
    .Pseudo = "<逻辑符>遭遇的部队是攻击方"
    .OpID = encountered_party_is_attacker
    .Type = OPT_Can_Fail
    .ParaNum = 0
    ReDim .Para(0)
End With

With Operation(57)
    .Op_name = "conversation_screen_is_active"        ' (conversation_screen_active), 'used in mission template triggers only
    .Op_CSVname = "对话窗口是激活的"
    .OpID = conversation_screen_is_active
    .Type = OPT_Can_Fail
    .ParaNum = 0
    ReDim .Para(0)
End With

With Operation(58)
    .Op_name = "dialog_box"        ' (tutorial_box,<text_string_id>,<title_string_id>),
    .Op_CSVname = "显示对话框"
    .Pseudo = "显示对话框:<内容>,<标题(可选)>"
    .OpID = dialog_box
    .Type = OPT_NONE
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_String)
    .Para(1).Value = "内容"
    .Para(2).Para_Type = CStr(Tag_String) & "#"
    .Para(2).Value = "标题(可选)"
End With

With Operation(59)
    .Op_name = "question_box"       ' (question_box,<string_id>, [<yes_string_id>], [<no_string_id>]),
    .Op_CSVname = "显示提问框"
    .Pseudo = "显示提问框:<内容>,<是>,<否>"
    .OpID = question_box
    .Type = OPT_NONE
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_String)
    .Para(1).Value = "内容"
    .Para(2).Para_Type = CStr(Tag_String) & "#"
    .Para(2).Value = "是"
    .Para(3).Para_Type = CStr(Tag_String) & "#"
    .Para(3).Value = "否"
End With

With Operation(60)
    .Op_name = "tutorial_message"      ' (tutorial_message,<string_id>, <color>), 'set string_id = -1 for hiding the message
    .Op_CSVname = "显示教学信息"
    .Pseudo = "显示<颜色>的教学信息<内容>"
    .OpID = tutorial_message
    .Type = OPT_NONE
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_String)
    .Para(1).Value = "内容"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "颜色"
End With

With Operation(61)
    .Op_name = "tutorial_message_set_position"      ' (tutorial_message_set_position, <position_x>, <position_y>),
    .Op_CSVname = "设置教学信息位置"
    .Pseudo = "设置教学信息位置坐标:<X>,<Y>"
    .OpID = tutorial_message_set_position
    .Type = OPT_NONE
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "0"
    .Para(1).Value = "X"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "Y"
End With

With Operation(62)
    .Op_name = "tutorial_message_set_size"     ' (tutorial_message_set_size, <size_x>, <size_y>),
    .Op_CSVname = "设置教学信息大小"
    .Pseudo = "设置教学信息大小:<长度>,<高度>"
    .OpID = tutorial_message_set_size
    .Type = OPT_NONE
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "0"
    .Para(1).Value = "长度"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "高度"
End With

With Operation(63)
    .Op_name = "tutorial_message_set_center_justify"     ' (tutorial_message_set_center_justify, <val>), 'set not 0 for center justify, 0 for not center justify
    .Op_CSVname = "设置教学信息居中"
    .Pseudo = "设置教学信息居中<值(非0为居中,0为不居中)>"
    .OpID = tutorial_message_set_center_justify
    .Type = OPT_NONE
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "0"
    .Para(1).Value = "值(非0为居中,0为不居中)"
End With

With Operation(64)
    .Op_name = "tutorial_message_set_background"    ' (tutorial_message_set_background, <value>), '1 = on, 0 = off, default is off
    .Op_CSVname = "设置教学信息背景"
    .Pseudo = "设置教学信息背景<开关(1=开,2=关)>"
    .OpID = tutorial_message_set_background
    .Type = OPT_NONE
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "0"
    .Para(1).Value = "开关(1=开,2=关)"
End With

With Operation(65)
    .Op_name = "troop_add_merchandise"    ' (troop_add_merchandise,<troop_id>,<item_type_id>,<value>),
    .Op_CSVname = "部队增加商品"
    .Pseudo = "<部队>增加<数目>个<物品类型>类的商品"
    .OpID = troop_add_merchandise
    .Type = OPT_NONE
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Troop)
    .Para(1).Value = "部队"
    .Para(2).Para_Type = "itp"
    .Para(2).Value = "物品类型"
    .Para(3).Para_Type = "0"
    .Para(3).Value = "数目"
End With

With Operation(66)
    .Op_name = "troop_add_merchandise_with_faction"     ' (troop_add_merchandise_with_faction,<troop_id>,<faction_id>,<item_type_id>,<value>), 'faction_id is given to check if troop is eligible to produce that item
    .Op_CSVname = "部队增加商品(阵营)"
    .Pseudo = "<阵营>的<部队>增加<数目>个<商品>(<物品类型>)"
    .OpID = troop_add_merchandise_with_faction
    .Type = OPT_NONE
    .ParaNum = 4
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Troop)
    .Para(1).Value = "部队"
    .Para(2).Para_Type = CStr(Tag_Faction)
    .Para(2).Value = "阵营"
    .Para(3).Para_Type = "itp"
    .Para(3).Value = "物品类型"
    .Para(4).Para_Type = "0"
    .Para(4).Value = "数目"
End With

With Operation(67)
    .Op_name = "set_merchandise_modifier_quality"     ' Quality rate in percentage (average quality = 100),
                                                      ' (set_merchandise_modifier_quality,<value>),
    .Op_CSVname = "设置商品前缀品质"
    .Pseudo = "设置商品前缀为<品质百分比(一般品质100)>"
    .OpID = set_merchandise_modifier_quality
    .Type = OPT_NONE
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "0"
    .Para(1).Value = "品质百分比(一般品质100)"
End With

With Operation(68)
    .Op_name = "set_merchandise_max_value"     ' (set_merchandise_max_value,<value>),
    .Op_CSVname = "设置商品最大值"
    .Pseudo = "设置商品最大值为<值>"
    .OpID = set_merchandise_max_value
    .Type = OPT_NONE
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "0"
    .Para(1).Value = "值"
End With

With Operation(69)
    .Op_name = "reset_item_probabilities"     ' (reset_item_probabilities),
    .Op_CSVname = "重置物品出现几率"
    .Pseudo = "重置物品在商店出现的概率为<值(可选)>"
    .OpID = reset_item_probabilities
    .Type = OPT_NONE
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "0#"
    .Para(1).Value = "值(可选)"
End With

With Operation(70)
    .Op_name = "set_item_probability_in_merchandise"     ' (set_item_probability_in_merchandise,<itm_id>,<value>),
    .Op_CSVname = "设置商品物品概率"
    .Pseudo = "设置<物品>在商店出现的<概率>"
    .OpID = set_item_probability_in_merchandise
    .Type = OPT_NONE
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Item)
    .Para(1).Value = "物品"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "概率"
End With

With Operation(71)
    .Op_name = "troop_get_slot"       ' (troop_get_slot,<destination>,<troop_id>,<slot_no>),
    .Op_CSVname = "获得槽(兵种)"
    .Pseudo = "<变量> = <兵种>的槽<槽序号>"
    .OpID = troop_get_slot
    .Type = OPT_Lhs
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = CStr(Tag_Troop)
    .Para(2).Value = "兵种"
    .Para(3).Para_Type = "0"
    .Para(3).Value = "槽序号"
End With

With Operation(72)
    .Op_name = "party_get_slot"       ' (party_get_slot,<destination>,<party_id>,<slot_no>),
    .Op_CSVname = "获得槽(部队)"
    .Pseudo = "<变量> = <部队>的槽<槽序号>"
    .OpID = party_get_slot
    .Type = OPT_Lhs
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = CStr(Tag_Party)
    .Para(2).Value = "部队"
    .Para(3).Para_Type = "0"
    .Para(3).Value = "槽序号"
End With

With Operation(73)
    .Op_name = "faction_get_slot"       ' (faction_get_slot,<destination>,<faction_id>,<slot_no>),
    .Op_CSVname = "获得槽(阵营)"
    .Pseudo = "<变量> = <阵营>的槽<槽序号>"
    .OpID = faction_get_slot
    .Type = OPT_Lhs
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = CStr(Tag_Faction)
    .Para(2).Value = "阵营"
    .Para(3).Para_Type = "0"
    .Para(3).Value = "槽序号"
End With

With Operation(74)
    .Op_name = "scene_get_slot"       ' (scene_get_slot,<destination>,<scene_id>,<slot_no>),
    .Op_CSVname = "获得槽(场景)"
    .Pseudo = "<变量> = <场景>的槽<槽序号>"
    .OpID = scene_get_slot
    .Type = OPT_Lhs
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = CStr(Tag_Scene)
    .Para(2).Value = "场景"
    .Para(3).Para_Type = "0"
    .Para(3).Value = "槽序号"
End With

With Operation(75)
    .Op_name = "party_template_get_slot"      ' (party_template_get_slot,<destination>,<party_template_id>,<slot_no>),
    .Op_CSVname = "获得槽(部队模板)"
    .Pseudo = "<变量> = <部队模板>的槽<槽序号>"
    .OpID = party_template_get_slot
    .Type = OPT_Lhs
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = CStr(Tag_Party_Tpl)
    .Para(2).Value = "部队模板"
    .Para(3).Para_Type = "0"
    .Para(3).Value = "槽序号"
End With

With Operation(76)
    .Op_name = "agent_get_slot"      ' (agent_get_slot,<destination>,<agent_id>,<slot_no>),
    .Op_CSVname = "获得槽(角色)"
    .Pseudo = "<变量> = (角色)<角色>的槽<槽序号>"
    .OpID = agent_get_slot
    .Type = OPT_Lhs
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "角色"
    .Para(3).Para_Type = "0"
    .Para(3).Value = "槽序号"
End With

With Operation(77)
    .Op_name = "quest_get_slot"       ' (quest_get_slot,<destination>,<quest_id>,<slot_no>),
    .Op_CSVname = "获得槽(任务)"
    .Pseudo = "<变量> = <任务>的槽<槽序号>"
    .OpID = quest_get_slot
    .Type = OPT_Lhs
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = CStr(Tag_Quest)
    .Para(2).Value = "任务"
    .Para(3).Para_Type = "0"
    .Para(3).Value = "槽序号"
End With

With Operation(78)
    .Op_name = "item_get_slot"       ' (item_get_slot,<destination>,<item_id>,<slot_no>),
    .Op_CSVname = "获得槽(物品)"
    .Pseudo = "<变量> = <物品>的槽<槽序号>"
    .OpID = item_get_slot
    .Type = OPT_Lhs
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = CStr(Tag_Item)
    .Para(2).Value = "物品"
    .Para(3).Para_Type = "0"
    .Para(3).Value = "槽序号"
End With

With Operation(79)
    .Op_name = "player_get_slot"       ' (player_get_slot,<destination>,<player_id>,<slot_no>),
    .Op_CSVname = "获得槽(玩家)"
    .Pseudo = "<变量> = 玩家:<玩家>的槽<槽序号>"
    .OpID = player_get_slot
    .Type = OPT_Lhs
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "玩家"
    .Para(3).Para_Type = "0"
    .Para(3).Value = "槽序号"
End With

With Operation(80)
    .Op_name = "team_get_slot"       ' (team_get_slot,<destination>,<player_id>,<slot_no>),
    .Op_CSVname = "获得槽(队伍)"
    .Pseudo = "<变量> = (队伍)<队伍>的槽<槽序号>"
    .OpID = team_get_slot
    .Type = OPT_Lhs
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "队伍"
    .Para(3).Para_Type = "0"
    .Para(3).Value = "槽序号"
End With

With Operation(81)
    .Op_name = "scene_prop_get_slot"        ' (scene_prop_get_slot,<destination>,<scene_prop_instance_id>,<slot_no>),
    .Op_CSVname = "获得槽(道具)"
    .Pseudo = "<变量> = <道具>的槽<槽序号>"
    .OpID = scene_prop_get_slot
    .Type = OPT_Lhs
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = CStr(Tag_Scene_Prop)
    .Para(2).Value = "道具"
    .Para(3).Para_Type = "0"
    .Para(3).Value = "槽序号"
End With

With Operation(82)
    .Op_name = "troop_ensure_inventory_space"         ' (troop_ensure_inventory_space,<troop_id>,<value>),
    .Op_CSVname = "确保物品栏空间"
    .Pseudo = "确保<兵种>的物品栏有<数目>个空余空间"
    .OpID = troop_ensure_inventory_space
    .Type = OPT_NONE
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Troop)
    .Para(1).Value = "兵种"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "数目"
End With

With Operation(83)
    .Op_name = "troop_sort_inventory"          ' (troop_sort_inventory,<troop_id>),
    .Op_CSVname = "物品栏排序"
    .Pseudo = "将<兵种>的物品栏排序"
    .OpID = troop_sort_inventory
    .Type = OPT_NONE
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Troop)
    .Para(1).Value = "兵种"
End With

With Operation(84)
    .Op_name = "store_random_in_range"          ' gets random number in range [range_low,range_high] excluding range_high
                                                ' (store_random_in_range,<destination>,<range_low>,<range_high>),
    .Op_CSVname = "产生随机数"
    .Pseudo = "<变量> = <下界>到<上界>中的随机数"
    .OpID = store_random_in_range
    .Type = OPT_Lhs
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = ""
    .Para(2).Value = "下界"
    .Para(3).Para_Type = ""
    .Para(3).Value = "上界"
End With

With Operation(85)
    .Op_name = "store_troop_gold"           ' (store_troop_gold,<destination>,<troop_id>),
    .Op_CSVname = "储存兵种金钱"
    .Pseudo = "<变量> = <兵种>所值金钱数"
    .OpID = store_troop_gold
    .Type = OPT_Lhs
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = CStr(Tag_Troop)
    .Para(2).Value = "兵种"
End With

With Operation(86)
    .Op_name = "assign"          ' had to put this here so that it can be called from conditions.
                                 ' (assign,<destination>,<value>),
    .Op_CSVname = "赋值"
    .Pseudo = "<变量> = <值>"
    .OpID = assign
    .Type = OPT_Lhs
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = ""
    .Para(2).Value = "值"
End With

With Operation(87)
    .Op_name = "spawn_around_party"          ' ID of spawned party is put into reg(0)
                                             ' (spawn_around_party,<party_id>,<party_template_id>),
    .Op_CSVname = "在部队附近出生"
    .Pseudo = "在<部队>附近产生一个套用<部队模板>的部队,新产生的部队ID被存在寄存器0中"
    .OpID = spawn_around_party
    .Type = OPT_Lhs
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Party)
    .Para(1).Value = "部队"
    .Para(2).Para_Type = CStr(Tag_Party_Tpl)
    .Para(2).Value = "部队模板"
End With

With Operation(88)
    .Op_name = "set_spawn_radius"          ' (set_spawn_radius,<value>),
    .Op_CSVname = "设置出生半径"
    .Pseudo = "设置出生半径为<半径>"
    .OpID = set_spawn_radius
    .Type = OPT_Lhs
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "0"
    .Para(1).Value = "半径"
End With

With Operation(89)
    .Op_name = "store_num_parties_of_template"          ' (store_num_parties_of_template,<destination>,<party_template_id>),
    .Op_CSVname = "储存部队模板部队数"
    .Pseudo = "<变量> = 套用了<部队模板>的部队数"
    .OpID = store_num_parties_of_template
    .Type = OPT_Lhs
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = CStr(Tag_Party_Tpl)
    .Para(2).Value = "部队模板"
End With

With Operation(90)
    .Op_name = "store_random_party_of_template"         ' fails if no party exists with tempolate_id (expensive)
                                                        ' (store_random_party_of_template,<destination>,<party_template_id>),
    .Op_CSVname = "储存部队模板的随机部队"
    .Pseudo = "<变量> = 随机从套用了<部队模板>的所有部队中取出的一个部队"
    .OpID = store_random_party_of_template
    .Type = OPT_Lhs
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = CStr(Tag_Party_Tpl)
    .Para(2).Value = "部队模板"
End With

With Operation(91)
    .Op_name = "check_quest_active"         ' (check_quest_active,<quest_id>),
    .Op_CSVname = "检查任务是否激活"
    .Pseudo = "<逻辑符><任务>已激活"
    .OpID = check_quest_active
    .Type = OPT_Can_Fail
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Quest)
    .Para(1).Value = "任务"
End With

With Operation(92)
    .Op_name = "check_quest_finished"        ' (check_quest_finished,<quest_id>),
    .Op_CSVname = "检查任务是否结束"
    .Pseudo = "<逻辑符><任务>已结束"
    .OpID = check_quest_finished
    .Type = OPT_Can_Fail
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Quest)
    .Para(1).Value = "任务"
End With

With Operation(93)
    .Op_name = "check_quest_succeeded"        ' (check_quest_succeeded,<quest_id>),
    .Op_CSVname = "检查任务是否成功"
    .Pseudo = "<逻辑符><任务>已成功"
    .OpID = check_quest_succeeded
    .Type = OPT_Can_Fail
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Quest)
    .Para(1).Value = "任务"
End With

With Operation(94)
    .Op_name = "check_quest_failed"        ' (check_quest_failed,<quest_id>),
    .Op_CSVname = "检查任务是否失败"
    .Pseudo = "<逻辑符><任务>已失败"
    .OpID = check_quest_failed
    .Type = OPT_Can_Fail
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Quest)
    .Para(1).Value = "任务"
End With

With Operation(95)
    .Op_name = "check_quest_concluded"         ' (check_quest_concluded,<quest_id>),
    .Op_CSVname = "检查任务是否结束"
    .Pseudo = "<逻辑符><任务>已结束"
    .OpID = check_quest_concluded
    .Type = OPT_Can_Fail
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Quest)
    .Para(1).Value = "任务"
End With

With Operation(96)
    .Op_name = "store_troop_faction"         ' (store_troop_faction,<destination>,<troop_id>),
    .Op_CSVname = "储存兵种阵营"
    .Pseudo = "<变量> = <兵种>的阵营"
    .OpID = store_troop_faction
    .Type = OPT_Lhs
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = CStr(Tag_Troop)
    .Para(2).Value = "兵种"
End With

With Operation(97)
    .Op_name = "store_faction_of_party"        ' (store_faction_of_party,<destination>,<party_id>),
    .Op_CSVname = "储存部队阵营"
    .Pseudo = "<变量> = <部队>的阵营"
    .OpID = store_faction_of_party
    .Type = OPT_Lhs
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = CStr(Tag_Party)
    .Para(2).Value = "部队"
End With

With Operation(98)
    .Op_name = "agent_deliver_damage_to_agent"    ' (agent_deliver_damage_to_agent,<agent_id_deliverer>,<agent_id>,<value>), 'if value <= 0, then damage will be calculated using the weapon item
    .Op_CSVname = "给予伤害"
    .Pseudo = "(角色)<攻击方>施加<伤害值(可选参数,若小于0则按武器伤害计算)>点伤害给(角色)<被攻击方>"
    .OpID = agent_deliver_damage_to_agent
    .Type = OPT_NONE
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "0"
    .Para(1).Value = "攻击方"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "被攻击方"
    .Para(3).Para_Type = "0#"
    .Para(3).Value = "伤害值(可选参数,若小于0则按武器伤害计算)"
End With

With Operation(99)
    .Op_name = "agent_get_look_position"    ' (agent_get_look_position, <position_no>, <agent_id>),
    .Op_CSVname = "获得角色视线方向"
    .Pseudo = "<位置> = (角色)<角色>的视线方向"
    .OpID = agent_get_look_position
    .Type = OPT_NONE
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "pos"
    .Para(1).Value = "位置"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "角色"
End With

With Operation(100)
    .Op_name = "agent_get_position"    ' (agent_get_position,<position_no>,<agent_id>),
    .Op_CSVname = "获得角色位置"
    .Pseudo = "<位置> = (角色)<角色>所在位置"
    .OpID = agent_get_position
    .Type = OPT_Lhs
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "pos"
    .Para(1).Value = "位置"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "角色"
End With

With Operation(101)
    .Op_name = "agent_set_position"    ' (agent_set_position,<agent_id>,<position_no>),
    .Op_CSVname = "设置角色位置"
    .Pseudo = "移动(角色)<角色>到<位置>"
    .OpID = agent_set_position
    .Type = OPT_NONE
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "0"
    .Para(1).Value = "角色"
    .Para(2).Para_Type = "pos"
    .Para(2).Value = "位置"
End With

With Operation(102)
    .Op_name = "agent_set_look_target_agent"    ' (agent_set_look_target_agent, <agent_id>, <agent_id>), 'second agent_id is the target
    .Op_CSVname = "设置角色视点目标"
    .Pseudo = "使(角色)<角色>看着<目标角色>"
    .OpID = agent_set_look_target_agent
    .Type = OPT_NONE
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "0"
    .Para(1).Value = "角色"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "目标角色"
End With

With Operation(103)
    .Op_name = "agent_get_horse"     ' (agent_get_horse,<destination>,<agent_id>),
    .Op_CSVname = "获得角色的马"
    .Pseudo = "<变量> = (角色)<角色>所骑的马"
    .OpID = agent_get_horse
    .Type = OPT_Lhs
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "角色"
End With

With Operation(104)
    .Op_name = "agent_get_rider"     ' (agent_get_rider,<destination>,<agent_id>),
    .Op_CSVname = "获得角色的骑手"
    .Pseudo = "<变量> = 骑着(角色)<角色(马)>的骑手"
    .OpID = agent_get_rider
    .Type = OPT_Lhs
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "角色(马)"
End With

With Operation(105)
    .Op_name = "agent_get_party_id"     ' (agent_get_party_id,<destination>,<agent_id>),
    .Op_CSVname = "获得角色所属部队"
    .Pseudo = "<变量> = (角色)<角色>所属的部队"
    .OpID = agent_get_party_id
    .Type = OPT_Lhs
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "角色"
End With

With Operation(106)
    .Op_name = "agent_get_entry_no"    ' (agent_get_entry_no,<destination>,<agent_id>),
    .Op_CSVname = "获得角色入口"
    .Pseudo = "<变量> = (角色)<角色>的入口"
    .OpID = agent_get_entry_no
    .Type = OPT_Lhs
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "角色"
End With

With Operation(107)
    .Op_name = "agent_get_troop_id"    ' (agent_get_troop_id,<destination>, <agent_id>),
    .Op_CSVname = "获得角色兵种"
    .Pseudo = "<变量> = (角色)<角色>所属的兵种"
    .OpID = agent_get_troop_id
    .Type = OPT_Lhs
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "角色"
End With

With Operation(108)
    .Op_name = "agent_get_item_id"     ' (agent_get_item_id,<destination>, <agent_id>), (works only for horses, returns -1 otherwise)
    .Op_CSVname = "获得角色(马)的物品ID"
    .Pseudo = "<变量> = (角色)<角色(马)>的物品ID"
    .OpID = agent_get_item_id
    .Type = OPT_Lhs
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "角色(马)"
End With

With Operation(109)
    .Op_name = "store_agent_hit_points"      ' set absolute to 1 to retrieve actual hps, otherwise will return relative hp in range [0..100]
                                           ' (store_agent_hit_points,<destination>,<agent_id>,[absolute]),
    .Op_CSVname = "储存角色生命值"
    .Pseudo = "<变量> = (角色)<角色>的生命值,返回:<返回值设定(可选参数)>"
    .OpID = store_agent_hit_points
    .Type = OPT_Lhs
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "角色"
    .Para(3).Para_Type = "as#"
    .Para(3).Value = "返回值设定(可选参数)"
End With

With Operation(110)
    .Op_name = "agent_set_hit_points"      ' set absolute to 1 if value is absolute, otherwise value will be treated as relative number in range [0..100]
                                           ' (agent_set_hit_points,<agent_id>,<value>,[absolute]),
    .Op_CSVname = "设置角色生命值"
    .Pseudo = "设置(角色)<角色>的生命值为<值>,返回:<返回值设定(可选参数)>"
    .OpID = agent_set_hit_points
    .Type = OPT_NONE
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "0"
    .Para(1).Value = "角色"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "值"
    .Para(3).Para_Type = "as#"
    .Para(3).Value = "返回值设定(可选参数)"
End With

With Operation(111)
    .Op_name = "get_angle_between_positions"      ' (get_angle_between_positions, <destination_fixed_point>, <position_no_1>, <position_no_2>),
    .Op_CSVname = "获得位置之间的夹角"
    .Pseudo = "<变量> = <位置1>与<位置2>之间的夹角"
    .OpID = get_angle_between_positions
    .Type = OPT_Lhs
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = "pos"
    .Para(2).Value = "位置1"
    .Para(3).Para_Type = "pos"
    .Para(3).Value = "位置2"
End With

With Operation(112)
    .Op_name = "position_has_line_of_sight_to_position"      ' (position_has_line_of_sight_to_position, <position_no_1>, <position_no_2>),
    .Op_CSVname = "位置之间无遮挡物直接可视"
    .Pseudo = "<逻辑符><位置1>与<位置2>之间无遮挡物直接可视"
    .OpID = position_has_line_of_sight_to_position
    .Type = OPT_Can_Fail
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "pos"
    .Para(1).Value = "位置1"
    .Para(2).Para_Type = "pos"
    .Para(2).Value = "位置2"
End With

With Operation(113)
    .Op_name = "get_distance_between_positions"      ' gets distance in centimeters. ' (get_distance_between_positions,<destination>,<position_no_1>,<position_no_2>),
    .Op_CSVname = "获得位置之间的距离(厘米)"
    .Pseudo = "<变量> = <位置1>与<位置2>之间距离(厘米)"
    .OpID = get_distance_between_positions
    .Type = OPT_Lhs
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = "pos"
    .Para(2).Value = "位置1"
    .Para(3).Para_Type = "pos"
    .Para(3).Value = "位置2"
End With

With Operation(114)
    .Op_name = "get_distance_between_positions_in_meters"      ' gets distance in meters. ' (get_distance_between_positions_in_meters,<destination>,<position_no_1>,<position_no_2>),
    .Op_CSVname = "获得位置之间的距离(米)"
    .Pseudo = "<变量> = <位置1>与<位置2>之间距离(米)"
    .OpID = get_distance_between_positions_in_meters
    .Type = OPT_Lhs
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = "pos"
    .Para(2).Value = "位置1"
    .Para(3).Para_Type = "pos"
    .Para(3).Value = "位置2"
End With

With Operation(115)
    .Op_name = "get_sq_distance_between_positions"      ' gets squared distance in centimeters ' (get_sq_distance_between_positions,<destination>,<position_no_1>,<position_no_2>),
    .Op_CSVname = "获得位置之间的距离的平方(厘米)"
    .Pseudo = "<变量> = <位置1>与<位置2>之间距离的平方(厘米)"
    .OpID = get_sq_distance_between_positions
    .Type = OPT_Lhs
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = "pos"
    .Para(2).Value = "位置1"
    .Para(3).Para_Type = "pos"
    .Para(3).Value = "位置2"
End With

With Operation(116)
    .Op_name = "get_sq_distance_between_positions_in_meters"      ' gets squared distance in meters ' (get_sq_distance_between_positions_in_meters,<destination>,<position_no_1>,<position_no_2>),
    .Op_CSVname = "获得位置之间的距离的平方(米)"
    .Pseudo = "<变量> = <位置1>与<位置2>之间距离的平方(米)"
    .OpID = get_sq_distance_between_positions_in_meters
    .Type = OPT_Lhs
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = "pos"
    .Para(2).Value = "位置1"
    .Para(3).Para_Type = "pos"
    .Para(3).Value = "位置2"
End With

With Operation(117)
    .Op_name = "position_is_behind_position"       ' (position_is_behind_position,<position_no_1>,<position_no_2>),
    .Op_CSVname = "位置1在位置2后面"
    .Pseudo = "<逻辑符><位置1>在<位置2>后面"
    .OpID = position_is_behind_position
    .Type = OPT_Can_Fail
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "pos"
    .Para(1).Value = "位置1"
    .Para(2).Para_Type = "pos"
    .Para(2).Value = "位置2"
End With

With Operation(118)
    .Op_name = "get_sq_distance_between_position_heights"      ' gets squared distance in centimeters ' (get_sq_distance_between_position_heights,<destination>,<position_no_1>,<position_no_2>),
    .Op_CSVname = "获得位置间高度的平方(厘米)"
    .Pseudo = "<变量> = <位置1>与<位置2>之间高度差的平方(厘米)"
    .OpID = get_sq_distance_between_position_heights
    .Type = OPT_Lhs
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = "pos"
    .Para(2).Value = "位置1"
    .Para(3).Para_Type = "pos"
    .Para(3).Value = "位置2"
End With

With Operation(119)
    .Op_name = "copy_position"      ' copies position_no_2 to position_no_1
                                    ' (copy_position,<position_no_1>,<position_no_2>),@
    .Op_CSVname = "复制位置2到位置1"
    .Pseudo = "复制<位置2>到<位置1>"
    .OpID = copy_position
    .Type = OPT_NONE
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "pos"
    .Para(1).Value = "位置1"
    .Para(2).Para_Type = "pos"
    .Para(2).Value = "位置2"
End With

With Operation(120)
    .Op_name = "party_get_num_companions"      ' (party_get_num_companions,<destination>,<party_id>),
    .Op_CSVname = "获得部队中同伴的数目"
    .Pseudo = "<变量> = <部队>中同伴的数目"
    .OpID = party_get_num_companions
    .Type = OPT_Lhs
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = CStr(Tag_Party)
    .Para(2).Value = "部队"
End With

With Operation(121)
    .Op_name = "party_get_num_prisoners"      ' (party_get_num_prisoners,<destination>,<party_id>),
    .Op_CSVname = "获得部队所带俘虏数"
    .Pseudo = "<变量> = 获得<部队>所带俘虏数"
    .OpID = party_get_num_prisoners
    .Type = OPT_Lhs
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = CStr(Tag_Party)
    .Para(2).Value = "部队"
End With

With Operation(122)
    .Op_name = "party_set_flags"      ' (party_set_flag, <party_id>, <flag>, <clear_or_set>), 'sets flags like pf_default_behavior. see header_parties.py for flags.
    .Op_CSVname = "设置部队标签"
    .Pseudo = "<设置><部队>的标签:<部队标签>"
    .OpID = party_set_flags
    .Type = OPT_Lhs
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Party)
    .Para(1).Value = "部队"
    .Para(2).Para_Type = "pf"
    .Para(2).Value = "部队标签"
    .Para(3).Para_Type = "bs"
    .Para(3).Value = "设置"
End With

With Operation(123)
    .Op_name = "party_set_marshall"      ' (party_set_marshall, <party_id>, <value>)
    .Op_CSVname = "设置部队的元帅"
    .Pseudo = "设置<部队>的元帅为<值>"
    .OpID = party_set_marshall
    .Type = OPT_Lhs
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Party)
    .Para(1).Value = "部队"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "值"   'remained to be improved
End With

With Operation(124)
    .Op_name = "party_set_extra_text"      ' (party_set_extra_text,<party_id>, <string>)
    .Op_CSVname = "设置部队的额外文本"
    .Pseudo = "设置<部队>的额外文本为<文本>"
    .OpID = party_set_extra_text
    .Type = OPT_Lhs
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Party)
    .Para(1).Value = "部队"
    .Para(2).Para_Type = CStr(Tag_String)
    .Para(2).Value = "文本"
End With

With Operation(125)
    .Op_name = "party_set_aggressiveness"     ' (party_set_aggressiveness, <party_id>, <number>),
    .Op_CSVname = "设置部队的攻击性"
    .Pseudo = "设置<部队>的攻击性为<值>"
    .OpID = party_set_aggressiveness
    .Type = OPT_Lhs
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Party)
    .Para(1).Value = "部队"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "值"
End With

With Operation(126)
    .Op_name = "party_set_courage"     ' (party_set_courage, <party_id>, <number>),
    .Op_CSVname = "设置部队的勇气值"
    .Pseudo = "设置<部队>的勇气值为<值>"
    .OpID = party_set_courage
    .Type = OPT_Lhs
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Party)
    .Para(1).Value = "部队"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "值"
End With

With Operation(127)
    .Op_name = "party_get_current_terrain"     ' (party_get_current_terrain,<destination>,<party_id>),
    .Op_CSVname = "获得部队所在地形"
    .Pseudo = "<变量> = <部队>所处的地形"
    .OpID = party_get_current_terrain
    .Type = OPT_Lhs
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = CStr(Tag_Party)
    .Para(2).Value = "部队"
End With

With Operation(128)
    .Op_name = "party_get_template_id"     ' (party_get_current_terrain,<destination>,<party_id>),
    .Op_CSVname = "获得部队的部队模板"
    .Pseudo = "<变量> = <部队>的部队模板"
    .OpID = party_get_template_id
    .Type = OPT_Lhs
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = CStr(Tag_Party)
    .Para(2).Value = "部队"
End With

With Operation(129)
    .Op_name = "party_add_members"     ' (party_add_members,<party_id>,<troop_id>,<number>), 'returns number added in reg0 'remained to be improved
    .Op_CSVname = "为部队增加成员"
    .Pseudo = "为<部队>增加<数目>个<兵种>作为成员,增加的数目被存在寄存器0中"
    .OpID = party_add_members
    .Type = OPT_NONE
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Party)
    .Para(1).Value = "部队"
    .Para(2).Para_Type = CStr(Tag_Troop)
    .Para(2).Value = "兵种"
    .Para(3).Para_Type = "0"
    .Para(3).Value = "数目"
End With

With Operation(130)
    .Op_name = "party_add_prisoners"     ' (party_add_prisoners,<party_id>,<troop_id>,<number>),'returns number added in reg0
    .Op_CSVname = "为部队增加俘虏"
    .Pseudo = "为<部队>增加<数目>个<兵种>作为俘虏,增加的数目被存在寄存器0中"
    .OpID = party_add_prisoners
    .Type = OPT_NONE
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Party)
    .Para(1).Value = "部队"
    .Para(2).Para_Type = CStr(Tag_Troop)
    .Para(2).Value = "兵种"
    .Para(3).Para_Type = "0"
    .Para(3).Value = "数目"
End With

With Operation(131)
    .Op_name = "party_add_leader"      ' (party_add_leader,<party_id>,<troop_id>,[<number>]),
    .Op_CSVname = "为部队增加领导"
    .Pseudo = "为<部队>增加<数目(可选参数)>个<兵种>作为领导"
    .OpID = party_add_leader
    .Type = OPT_NONE
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Party)
    .Para(1).Value = "部队"
    .Para(2).Para_Type = CStr(Tag_Troop)
    .Para(2).Value = "兵种"
    .Para(3).Para_Type = "0#"
    .Para(3).Value = "数目(可选参数)"
End With

With Operation(132)
    .Op_name = "party_force_add_members"       ' (party_force_add_members,<party_id>,<troop_id>,<number>),
    .Op_CSVname = "强制为部队增加成员"
    .Pseudo = "强制为<部队>增加<数目>个<兵种>作为成员"
    .OpID = party_force_add_members
    .Type = OPT_NONE
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Party)
    .Para(1).Value = "部队"
    .Para(2).Para_Type = CStr(Tag_Troop)
    .Para(2).Value = "兵种"
    .Para(3).Para_Type = "0"
    .Para(3).Value = "数目"
End With

With Operation(133)
    .Op_name = "party_force_add_prisoners"       ' (party_force_add_prisoners,<party_id>,<troop_id>,<number>),
    .Op_CSVname = "强制为部队增加俘虏"
    .Pseudo = "强制为<部队>增加<数目>个<兵种>作为俘虏"
    .OpID = party_force_add_prisoners
    .Type = OPT_NONE
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Party)
    .Para(1).Value = "部队"
    .Para(2).Para_Type = CStr(Tag_Troop)
    .Para(2).Value = "兵种"
    .Para(3).Para_Type = "0"
    .Para(3).Value = "数目"
End With

With Operation(134)
    .Op_name = "party_remove_members"       ' stores number removed to reg0
                                            ' (party_remove_members,<party_id>,<troop_id>,<number>),
    .Op_CSVname = "移除部队成员"
    .Pseudo = "移除<部队>中<数目>个<兵种>作为成员,移除的数目被存在寄存器0中"
    .OpID = party_remove_members
    .Type = OPT_NONE
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Party)
    .Para(1).Value = "部队"
    .Para(2).Para_Type = CStr(Tag_Troop)
    .Para(2).Value = "兵种"
    .Para(3).Para_Type = "0"
    .Para(3).Value = "数目"
End With

With Operation(135)
    .Op_name = "party_remove_prisoners"      ' stores number removed to reg0
                                             ' (party_remove_prisoners,<party_id>,<troop_id>,<number>),
    .Op_CSVname = "移除部队俘虏"
    .Pseudo = "移除<部队>中<数目>个<兵种>作为俘虏,移除的数目被存在寄存器0中"
    .OpID = party_remove_prisoners
    .Type = OPT_NONE
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Party)
    .Para(1).Value = "部队"
    .Para(2).Para_Type = CStr(Tag_Troop)
    .Para(2).Value = "兵种"
    .Para(3).Para_Type = "0"
    .Para(3).Value = "数目"
End With

With Operation(136)
    .Op_name = "party_clear"      ' (party_clear,<party_id>),
    .Op_CSVname = "清除部队成员"
    .Pseudo = "清除<部队>中所有成员及俘虏"
    .OpID = party_clear
    .Type = OPT_NONE
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Party)
    .Para(1).Value = "部队"
End With

With Operation(137)
    .Op_name = "party_wound_members"       ' (party_wound_members,<party_id>,<troop_id>,<number>),
    .Op_CSVname = "使部队中某些成员变为受伤状态"
    .Pseudo = "使<部队>中<数目>个<兵种>变为受伤状态"
    .OpID = party_wound_members
    .Type = OPT_NONE
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Party)
    .Para(1).Value = "部队"
    .Para(2).Para_Type = CStr(Tag_Troop)
    .Para(2).Value = "兵种"
    .Para(3).Para_Type = "0"
    .Para(3).Value = "数目"
End With

With Operation(138)
    .Op_name = "party_remove_members_wounded_first"        ' stores number removed to reg0
                                                        ' (party_remove_members_wounded_first,<party_id>,<troop_id>,<number>),
    .Op_CSVname = "优先移除部队中的伤者"
    .Pseudo = "优先移除<部队>中<数目>个受伤的<兵种>,移除的数目被存在寄存器0中"
    .OpID = party_remove_members_wounded_first
    .Type = OPT_NONE
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Party)
    .Para(1).Value = "部队"
    .Para(2).Para_Type = CStr(Tag_Troop)
    .Para(2).Value = "兵种"
    .Para(3).Para_Type = "0"
    .Para(3).Value = "数目"
End With

With Operation(139)
    .Op_name = "party_set_faction"       ' (party_set_faction,<party_id>,<faction_id>),
    .Op_CSVname = "设置部队阵营"
    .Pseudo = "设置<部队>所属阵营为<阵营>"
    .OpID = party_set_faction
    .Type = OPT_NONE
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Party)
    .Para(1).Value = "部队"
    .Para(2).Para_Type = CStr(Tag_Faction)
    .Para(2).Value = "阵营"
End With

With Operation(140)
    .Op_name = "party_relocate_near_party"       ' (party_relocate_near_party,<party_id>,<target_party_id>,<value_spawn_radius>),
    .Op_CSVname = "将部队重定位到目标部队附近"
    .Pseudo = "将<部队>传送到<目标部队>附近<半径>的范围内"
    .OpID = party_relocate_near_party
    .Type = OPT_NONE
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Party)
    .Para(1).Value = "部队"
    .Para(2).Para_Type = CStr(Tag_Party)
    .Para(2).Value = "目标部队"
    .Para(3).Para_Type = "0"
    .Para(3).Value = "半径"
End With

With Operation(141)
    .Op_name = "party_get_position"       ' (party_get_position,<position_no>,<party_id>),
    .Op_CSVname = "获得部队位置"
    .Pseudo = "<位置> = <部队>所处位置"
    .OpID = party_get_position
    .Type = OPT_NONE
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "pos"
    .Para(1).Value = "位置"
    .Para(2).Para_Type = CStr(Tag_Party)
    .Para(2).Value = "部队"
End With

With Operation(142)
    .Op_name = "party_set_position"       ' (party_set_position,<party_id>,<position_no>),
    .Op_CSVname = "设置部队位置"
    .Pseudo = "移动<部队>到<位置>"
    .OpID = party_set_position
    .Type = OPT_NONE
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Party)
    .Para(1).Value = "部队"
    .Para(2).Para_Type = "pos"
    .Para(2).Value = "位置"
End With

With Operation(143)
    .Op_name = "map_get_random_position_around_position"    ' (map_get_random_position_around_position,<dest_position_no>,<source_position_no>,<radius>),
    .Op_CSVname = "获得源位置附近的随机位置"
    .Pseudo = "<位置> = <源位置>附近<半径>范围内的随机位置"
    .OpID = map_get_random_position_around_position
    .Type = OPT_NONE
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "pos"
    .Para(1).Value = "位置"
    .Para(2).Para_Type = "pos"
    .Para(2).Value = "源位置"
    .Para(3).Para_Type = "0"
    .Para(3).Value = "半径"
End With

With Operation(144)
    .Op_name = "map_get_land_position_around_position"   ' (map_get_land_position_around_position,<dest_position_no>,<source_position_no>,<radius>),
    .Op_CSVname = "获得源位置附近在陆地上的随机位置"
    .Pseudo = "<逻辑符><位置> = <源位置>附近<半径>范围内在陆地上的随机位置"
    .OpID = map_get_land_position_around_position
    .Type = OPT_Can_Fail
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "pos"
    .Para(1).Value = "位置"
    .Para(2).Para_Type = "pos"
    .Para(2).Value = "源位置"
    .Para(3).Para_Type = "0"
    .Para(3).Value = "半径"
End With

With Operation(145)
    .Op_name = "map_get_water_position_around_position"   ' (map_get_water_position_around_position,<dest_position_no>,<source_position_no>,<radius>),
    .Op_CSVname = "获得源位置附近在水面上的随机位置"
    .Pseudo = "<逻辑符><位置> = <源位置>附近<半径>范围内在水面上的随机位置"
    .OpID = map_get_water_position_around_position
    .Type = OPT_Can_Fail
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "pos"
    .Para(1).Value = "位置"
    .Para(2).Para_Type = "pos"
    .Para(2).Value = "源位置"
    .Para(3).Para_Type = "0"
    .Para(3).Value = "半径"
End With

With Operation(146)
    .Op_name = "party_count_members_of_type"      ' (party_count_members_of_type,<destination>,<party_id>,<troop_id>),
    .Op_CSVname = "获得部队中某一兵种的所有成员的数目"
    .Pseudo = "<变量> = <部队>中所有类型为<兵种>的成员数目"
    .OpID = party_count_members_of_type
    .Type = OPT_Lhs
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = CStr(Tag_Party)
    .Para(2).Value = "部队"
    .Para(3).Para_Type = CStr(Tag_Troop)
    .Para(3).Value = "兵种"
End With

With Operation(147)
    .Op_name = "party_count_companions_of_type"       ' (party_count_companions_of_type,<destination>,<party_id>,<troop_id>),
    .Op_CSVname = "获得部队中某一兵种的同伴的数目"
    .Pseudo = "<变量> = <部队>中类型为<兵种>的同伴数目"
    .OpID = party_count_companions_of_type
    .Type = OPT_Lhs
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = CStr(Tag_Party)
    .Para(2).Value = "部队"
    .Para(3).Para_Type = CStr(Tag_Troop)
    .Para(3).Value = "兵种"
End With

With Operation(148)
    .Op_name = "party_count_prisoners_of_type"       ' (party_count_prisoners_of_type,<destination>,<party_id>,<troop_id>),
    .Op_CSVname = "获得部队中某一兵种的俘虏的数目"
    .Pseudo = "<变量> = <部队>中类型为<兵种>的俘虏数目"
    .OpID = party_count_prisoners_of_type
    .Type = OPT_Lhs
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = CStr(Tag_Party)
    .Para(2).Value = "部队"
    .Para(3).Para_Type = CStr(Tag_Troop)
    .Para(3).Value = "兵种"
End With

With Operation(149)
    .Op_name = "party_get_free_companions_capacity"      ' (party_get_free_companions_capacity,<destination>,<party_id>),
    .Op_CSVname = "获得部队可携带同伴的剩余数目"
    .Pseudo = "<变量> = <部队>中可携带同伴的剩余数目"
    .OpID = party_get_free_companions_capacity
    .Type = OPT_Lhs
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = CStr(Tag_Party)
    .Para(2).Value = "部队"
End With

With Operation(150)
    .Op_name = "party_get_free_prisoners_capacity"      ' (party_get_free_prisoners_capacity,<destination>,<party_id>),
    .Op_CSVname = "获得部队可携带俘虏的剩余数目"
    .Pseudo = "<变量> = <部队>中可携带俘虏的剩余数目"
    .OpID = party_get_free_prisoners_capacity
    .Type = OPT_Lhs
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = CStr(Tag_Party)
    .Para(2).Value = "部队"
End With

With Operation(151)
    .Op_name = "party_get_ai_initiative"      ' (party_get_ai_initiative,<destination>,<party_id>), 'result is between 0-100
    .Op_CSVname = "获得部队AI的主动性"
    .Pseudo = "<变量> = <部队>的主动性(0-100)"
    .OpID = party_get_ai_initiative
    .Type = OPT_Lhs
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = CStr(Tag_Party)
    .Para(2).Value = "部队"
End With

With Operation(152)
    .Op_name = "party_set_ai_initiative"      ' (party_set_ai_initiative,<party_id>,<value>), 'value is between 0-100
    .Op_CSVname = "设置部队AI的主动性"
    .Pseudo = "设置<部队>的主动性为<值(0-100)>"
    .OpID = party_set_ai_initiative
    .Type = OPT_NONE
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Party)
    .Para(1).Value = "部队"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "值(0-100)"
End With

With Operation(153)
    .Op_name = "party_set_ai_behavior"     ' (party_set_ai_behavior,<party_id>,<ai_bhvr>),
    .Op_CSVname = "设置部队AI的行为"
    .Pseudo = "设置<部队>的行为为<部队AI行为>"
    .OpID = party_set_ai_behavior
    .Type = OPT_NONE
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Party)
    .Para(1).Value = "部队"
    .Para(2).Para_Type = "ai_bhvr"
    .Para(2).Value = "部队AI行为"
End With

With Operation(154)
    .Op_name = "party_set_ai_object"      ' (party_set_ai_object,<party_id>,<party_id>),
    .Op_CSVname = "设置部队的目标"
    .Pseudo = "设置<部队>的目标为<目标部队>"
    .OpID = party_set_ai_object
    .Type = OPT_NONE
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Party)
    .Para(1).Value = "部队"
    .Para(2).Para_Type = CStr(Tag_Party)
    .Para(2).Value = "目标部队"
End With

With Operation(155)
    .Op_name = "party_set_ai_target_position"       ' (party_set_ai_target_position,<party_id>,<position_no>),
    .Op_CSVname = "设置部队的目标位置"
    .Pseudo = "设置<部队>的目标位置为<位置>"
    .OpID = party_set_ai_target_position
    .Type = OPT_NONE
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Party)
    .Para(1).Value = "部队"
    .Para(2).Para_Type = "pos"
    .Para(2).Value = "位置"
End With

With Operation(156)
    .Op_name = "party_set_ai_patrol_radius"       ' (party_set_ai_patrol_radius,<party_id>,<radius_in_km>),
    .Op_CSVname = "设置部队的巡逻半径"
    .Pseudo = "设置<部队>的巡逻半径为<半径(km)>千米"
    .OpID = party_set_ai_patrol_radius
    .Type = OPT_NONE
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Party)
    .Para(1).Value = "部队"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "半径(km)"
End With

With Operation(157)
    .Op_name = "party_ignore_player"        ' (party_ignore_player, <party_id>,<duration_in_hours>), 'don't pursue player party for this duration
    .Op_CSVname = "使部队忽视玩家"
    .Pseudo = "使<部队>在<时间段(h)>小时内忽视玩家"
    .OpID = party_ignore_player
    .Type = OPT_NONE
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Party)
    .Para(1).Value = "部队"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "时间段(h)"
End With

With Operation(158)
    .Op_name = "party_set_bandit_attraction"       ' (party_set_bandit_attraction, <party_id>,<attaraction>), 'set how attractive a target the party is for bandits (0..100)
    .Op_CSVname = "设置部队对匪徒的吸引力"
    .Pseudo = "设置<部队>对匪徒的吸引力为<吸引力(0-100)>"
    .OpID = party_set_bandit_attraction
    .Type = OPT_NONE
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Party)
    .Para(1).Value = "部队"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "吸引力(0-100)"
End With

With Operation(159)
    .Op_name = "party_get_helpfulness"       ' (party_get_helpfulness,<destination>,<party_id>),
    .Op_CSVname = "获得部队帮助受攻击友军的倾向(0-10000)"
    .Pseudo = "<变量> = <部队>帮助受攻击友军的倾向(0-10000)"
    .OpID = party_get_helpfulness
    .Type = OPT_NONE
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = CStr(Tag_Party)
    .Para(2).Value = "部队"
End With

With Operation(160)
    .Op_name = "party_set_helpfulness"        ' (party_set_helpfulness, <party_id>, <number>), 'tendency to help friendly parties under attack. (0-10000, 100 default.)
    .Op_CSVname = "设置部队帮助受攻击友军的倾向(0-10000,默认100)"
    .Pseudo = "设置<部队>帮助受攻击友军的倾向为<倾向值(0-10000,默认100)>"
    .OpID = party_set_helpfulness
    .Type = OPT_NONE
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Party)
    .Para(1).Value = "部队"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "倾向值(0-10000,默认100)"
End With

With Operation(161)
    .Op_name = "set_fixed_point_multiplier"         ' (set_fixed_point_multiplier, <value>),
                                        ' sets the precision of the values that are named as value_fixed_point or destination_fixed_point.
                                        ' Default is 1 (every fixed point value will be regarded as an integer)
    .Op_CSVname = "设置浮点乘数"
    .Pseudo = "设置浮点乘数为<乘数值(整数,默认1)>"
    .OpID = set_fixed_point_multiplier
    .Type = OPT_NONE
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "0"
    .Para(1).Value = "乘数值(整数,默认1)"
End With

With Operation(162)
    .Op_name = "convert_to_fixed_point"         ' (convert_to_fixed_point, <destination_fixed_point>), multiplies the value with the fixed point multiplier
    .Op_CSVname = "使变量转换为浮点数"
    .Pseudo = "<变量(浮点数)> = 变量 × 浮点乘数"
    .OpID = convert_to_fixed_point
    .Type = OPT_NONE
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量(浮点数)"
End With

With Operation(163)
    .Op_name = "convert_from_fixed_point"         ' (convert_from_fixed_point, <destination>), divides the value with the fixed point multiplier
    .Op_CSVname = "使变量转换为整数"
    .Pseudo = "<变量> = 变量 ÷ 浮点乘数"
    .OpID = convert_from_fixed_point
    .Type = OPT_NONE
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
End With

With Operation(164)
    .Op_name = "store_script_param_1"          ' (store_script_param_1,<destination>),  --(Within a script) stores the first script parameter.
    .Op_CSVname = "储存脚本参数1"
    .Pseudo = "<变量> = 脚本参数1"
    .OpID = store_script_param_1
    .Type = OPT_Lhs
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
End With

With Operation(165)
    .Op_name = "store_script_param_2"          ' (store_script_param_2,<destination>),  --(Within a script) stores the second script parameter.
    .Op_CSVname = "储存脚本参数2"
    .Pseudo = "<变量> = 脚本参数2"
    .OpID = store_script_param_2
    .Type = OPT_Lhs
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
End With

With Operation(166)
    .Op_name = "store_script_param"          ' (store_script_param,<destination>,<script_param_no>), --(Within a script) stores <script_param_no>th script parameter.
    .Op_CSVname = "储存脚本参数"
    .Pseudo = "<变量> = 脚本参数<脚本参数序号>"
    .OpID = store_script_param
    .Type = OPT_Lhs
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = ":"
    .Para(1).Value = "变量"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "脚本参数序号"
End With

With Operation(167)
    .Op_name = "init_position"           ' (init_position,<position_no>),
    .Op_CSVname = "初始化位置"
    .Pseudo = "初始化<位置>"
    .OpID = init_position
    .Type = OPT_NONE
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "pos"
    .Para(1).Value = "位置"
End With

With Operation(168)
    .Op_name = "main_party_has_troop"           ' (main_party_has_troop,<troop_id>),
    .Op_CSVname = "Main_Party是否拥有兵种"
    .Pseudo = "<逻辑符>Main_Party拥有<兵种>"
    .OpID = main_party_has_troop
    .Type = OPT_Can_Fail
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Troop)
    .Para(1).Value = "兵种"
End With

With Operation(169)
    .Op_name = "party_is_in_town"           ' (party_is_in_town,<party_id_1>,<party_id_2>),
    .Op_CSVname = "部队是否在城镇里"
    .Pseudo = "<逻辑符><部队>驻扎在<城镇>里"
    .OpID = party_is_in_town
    .Type = OPT_Can_Fail
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Party)
    .Para(1).Value = "部队"
    .Para(2).Para_Type = CStr(Tag_Party)
    .Para(2).Value = "城镇"
End With

With Operation(170)
    .Op_name = "party_is_in_any_town"           ' (party_is_in_any_town,<party_id>),
    .Op_CSVname = "部队是否在任何城镇里"
    .Pseudo = "<逻辑符><部队>驻扎在城镇里"
    .OpID = party_is_in_any_town
    .Type = OPT_Can_Fail
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Party)
    .Para(1).Value = "部队"
End With

With Operation(171)
    .Op_name = "party_is_active"           ' (party_is_active,<party_id>),
    .Op_CSVname = "地图上存在该部队"
    .Pseudo = "<逻辑符>地图上存在<部队>"
    .OpID = party_is_active
    .Type = OPT_Can_Fail
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Party)
    .Para(1).Value = "部队"
End With

With Operation(172)
    .Op_name = "player_has_item"            ' (player_has_item,<item_id>),
    .Op_CSVname = "玩家拥有物品"
    .Pseudo = "<逻辑符>玩家拥有<物品>"
    .OpID = player_has_item
    .Type = OPT_Can_Fail
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Item)
    .Para(1).Value = "物品"
End With

With Operation(173)
    .Op_name = "troop_has_item_equipped"             ' (troop_has_item_equipped,<troop_id>,<item_id>),
    .Op_CSVname = "兵种装备了某物品"
    .Pseudo = "<逻辑符><兵种>装备了<物品>"
    .OpID = troop_has_item_equipped
    .Type = OPT_Can_Fail
    .ParaNum = 2
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Troop)
    .Para(1).Value = "兵种"
    .Para(2).Para_Type = CStr(Tag_Item)
    .Para(2).Value = "物品"
End With

With Operation(174)
    .Op_name = "troop_is_mounted"             ' (troop_is_mounted,<troop_id>),
    .Op_CSVname = "兵种是否是骑兵"
    .Pseudo = "<逻辑符><兵种>为骑兵"
    .OpID = troop_is_mounted
    .Type = OPT_Can_Fail
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Troop)
    .Para(1).Value = "兵种"
End With

With Operation(175)
    .Op_name = "troop_is_guarantee_ranged"              ' (troop_is_guarantee_ranged, <troop_id>),
    .Op_CSVname = "兵种是否保证装备远程武器"
    .Pseudo = "<逻辑符><兵种>保证装备远程武器"
    .OpID = troop_is_guarantee_ranged
    .Type = OPT_Can_Fail
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Troop)
    .Para(1).Value = "兵种"
End With

With Operation(176)
    .Op_name = "troop_is_guarantee_horse"              ' (troop_is_guarantee_horse, <troop_id>),
    .Op_CSVname = "兵种是否保证有马"
    .Pseudo = "<逻辑符><兵种>保证有马"
    .OpID = troop_is_guarantee_horse
    .Type = OPT_Can_Fail
    .ParaNum = 1
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Troop)
    .Para(1).Value = "兵种"
End With

With Operation(177)
    .Op_name = "troop_set_slot"              ' (troop_set_slot,<troop_id>,<slot_no>,<value>),
    .Op_CSVname = "设置兵种的槽"
    .Pseudo = "设置<兵种>的槽:<槽序号>为<值>"
    .OpID = troop_set_slot
    .Type = OPT_NONE
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Troop)
    .Para(1).Value = "兵种"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "槽序号"
    .Para(3).Para_Type = ""
    .Para(3).Value = "值"
End With

With Operation(178)
    .Op_name = "party_set_slot"               ' (party_set_slot,<party_id>,<slot_no>,<value>),
    .Op_CSVname = "设置部队的槽"
    .Pseudo = "设置<部队>的槽:<槽序号>为<值>"
    .OpID = party_set_slot
    .Type = OPT_NONE
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Party)
    .Para(1).Value = "部队"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "槽序号"
    .Para(3).Para_Type = ""
    .Para(3).Value = "值"
End With

With Operation(179)
    .Op_name = "faction_set_slot"               ' (faction_set_slot,<faction_id>,<slot_no>,<value>),
    .Op_CSVname = "设置阵营的槽"
    .Pseudo = "设置<阵营>的槽:<槽序号>为<值>"
    .OpID = faction_set_slot
    .Type = OPT_NONE
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Faction)
    .Para(1).Value = "阵营"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "槽序号"
    .Para(3).Para_Type = ""
    .Para(3).Value = "值"
End With

With Operation(180)
    .Op_name = "scene_set_slot"               ' (scene_set_slot,<scene_id>,<slot_no>,<value>),
    .Op_CSVname = "设置场景的槽"
    .Pseudo = "设置<场景>的槽:<槽序号>为<值>"
    .OpID = scene_set_slot
    .Type = OPT_NONE
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Scene)
    .Para(1).Value = "场景"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "槽序号"
    .Para(3).Para_Type = ""
    .Para(3).Value = "值"
End With

With Operation(181)
    .Op_name = "party_template_set_slot"               ' (party_template_set_slot,<party_template_id>,<slot_no>,<value>),
    .Op_CSVname = "设置部队模板的槽"
    .Pseudo = "设置<部队模板>的槽:<槽序号>为<值>"
    .OpID = party_template_set_slot
    .Type = OPT_NONE
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Party_Tpl)
    .Para(1).Value = "部队模板"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "槽序号"
    .Para(3).Para_Type = ""
    .Para(3).Value = "值"
End With

With Operation(182)
    .Op_name = "agent_set_slot"                ' (agent_set_slot,<agent_id>,<slot_no>,<value>),
    .Op_CSVname = "设置角色的槽"
    .Pseudo = "设置角色:<角色>的槽:<槽序号>为<值>"
    .OpID = agent_set_slot
    .Type = OPT_NONE
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "0"
    .Para(1).Value = "角色"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "槽序号"
    .Para(3).Para_Type = ""
    .Para(3).Value = "值"
End With

With Operation(183)
    .Op_name = "quest_set_slot"                ' (quest_set_slot,<quest_id>,<slot_no>,<value>),
    .Op_CSVname = "设置任务的槽"
    .Pseudo = "设置<任务>的槽:<槽序号>为<值>"
    .OpID = quest_set_slot
    .Type = OPT_NONE
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Quest)
    .Para(1).Value = "任务"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "槽序号"
    .Para(3).Para_Type = ""
    .Para(3).Value = "值"
End With

With Operation(184)
    .Op_name = "item_set_slot"                ' (item_set_slot,<item_id>,<slot_no>,<value>),
    .Op_CSVname = "设置物品的槽"
    .Pseudo = "设置<物品>的槽:<槽序号>为<值>"
    .OpID = item_set_slot
    .Type = OPT_NONE
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Item)
    .Para(1).Value = "物品"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "槽序号"
    .Para(3).Para_Type = ""
    .Para(3).Value = "值"
End With

With Operation(185)
    .Op_name = "player_set_slot"               ' (player_set_slot,<player_id>,<slot_no>,<value>),
    .Op_CSVname = "设置玩家的槽"
    .Pseudo = "设置玩家:<玩家>的槽:<槽序号>为<值>"
    .OpID = player_set_slot
    .Type = OPT_NONE
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "0"
    .Para(1).Value = "玩家"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "槽序号"
    .Para(3).Para_Type = ""
    .Para(3).Value = "值"
End With

With Operation(185)
    .Op_name = "team_set_slot"                ' (team_set_slot,<team_id>,<slot_no>,<value>),
    .Op_CSVname = "设置队伍的槽"
    .Pseudo = "设置队伍:<队伍>的槽:<槽序号>为<值>"
    .OpID = team_set_slot
    .Type = OPT_NONE
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = "0"
    .Para(1).Value = "队伍"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "槽序号"
    .Para(3).Para_Type = ""
    .Para(3).Value = "值"
End With

With Operation(186)
    .Op_name = "scene_prop_set_slot"                ' (scene_prop_set_slot,<scene_prop_instance_id>,<slot_no>,<value>),
    .Op_CSVname = "设置场景道具的槽"
    .Pseudo = "设置<场景道具>的槽:<槽序号>为<值>"
    .OpID = scene_prop_set_slot
    .Type = OPT_NONE
    .ParaNum = 3
    ReDim .Para(1 To .ParaNum)
    .Para(1).Para_Type = CStr(Tag_Scene_Prop)
    .Para(1).Value = "场景道具"
    .Para(2).Para_Type = "0"
    .Para(2).Value = "槽序号"
    .Para(3).Para_Type = ""
    .Para(3).Value = "值"
End With

Max_Op_Len = 42
End Sub
'*************************************************************************
'**函 数 名：InitControlOperationGroup
'**输    入：无
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-03-15 12:10:11
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub InitControlOperationGroup()
ControlOperationGroup(0) = 17
ControlOperationGroup(1) = 18
ControlOperationGroup(2) = 19
ControlOperationGroup(3) = 20
ControlOperationGroup(4) = 21
ControlOperationGroup(5) = 52
ControlOperationGroup(6) = 53
End Sub
'*************************************************************************
'**函 数 名：InitOptionalParamGroup
'**输    入：无
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-03-16 22:43:34
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub InitOptionalParamGroup()
OptionalParamGroup(0) = 0
'OptionalParamGroup(1) = 55
End Sub
'*************************************************************************
'**函 数 名：GetNegation
'**输    入：-(String)Op
'**输    出：-(Integer)0为无,1为非,2为或,3为都有
'**功能描述：判断逻辑操作符
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-03-18 22:49:11
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function GetNegation(Op As String) As Integer
Dim tI As Integer64b, neg As Integer, this_or_next As Integer

tI = StrToI64(Op)
neg = 128
this_or_next = 64

If tI.by(3) >= neg + this_or_next Then   'neg + this_or_next
     GetNegation = 3
ElseIf tI.by(3) < neg + this_or_next And tI.by(3) >= neg Then    'neg
     GetNegation = 1
ElseIf tI.by(3) < neg And tI.by(3) >= this_or_next Then     'this_or_next
     GetNegation = 2
Else
     GetNegation = 0
End If

End Function
'*************************************************************************
'**函 数 名：GetOpCodeInfo
'**输    入：-(String)Op,(Int,0为无,1为非,2为或,3为都有)Negation,(Long)OpID
'**输    出：-
'**功能描述：判断逻辑操作符并获得操作ID
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-03-18 22:49:11
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub GetOpCodeInfo(Op As String, Negation As Integer, OpID As Long)
Dim tI As Integer64b, neg As Integer, this_or_next As Integer

tI = StrToI64(Op)
neg = 128
this_or_next = 64

If tI.by(3) >= neg + this_or_next Then
     Negation = 3
ElseIf tI.by(3) < neg + this_or_next And tI.by(3) >= neg Then
     Negation = 1
ElseIf tI.by(3) < neg And tI.by(3) >= this_or_next Then
     Negation = 2
Else
     Negation = 0
End If

tI.by(3) = 0
OpID = CLng(Val(I64toStrNZ(tI)))

End Sub
'*************************************************************************
'**函 数 名：QuickGetOpCodeInfo
'**输    入：-(String)Op,(Int,0为无,1为非,2为或,3为都有)Negation,(Int)OpID
'**输    出：-
'**功能描述：快速判断逻辑操作符并获得操作ID
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-03-18 22:49:11
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub QuickGetOpCodeInfo(Op As String, Negation As Integer, OpID As Long)
Dim neg As Integer, this_or_next As Integer, Both As Integer, neg2 As Long, this_or_next2 As Long, Both2 As Long
Dim Top As String, tOp1 As String, tOp2 As String

Negation = 0

Top = Right("0000000000" & Op, 10)
tOp1 = Left(Top, 3)
tOp2 = Right(Top, 7)

neg = 214
this_or_next = 107
Both = 322

neg2 = 7483648
this_or_next2 = 3741824
Both2 = 1225472

If tOp1 >= Both Then
     Negation = 3
     OpID = Val(tOp2) - Both2
ElseIf tOp1 < Both And tOp1 >= neg Then
     Negation = 1
     OpID = Val(tOp2) - neg2
ElseIf tOp1 < neg And tOp1 >= this_or_next Then
     Negation = 2
     OpID = Val(tOp2) - this_or_next2
Else
     Negation = 0
     OpID = Val(Op)
End If

End Sub
'*************************************************************************
'**函 数 名：HaveNot
'**输    入：-(String)Op
'**输    出：-(Boolean)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-02-13 22:46:23
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function HaveNot(Op As String) As Boolean
Dim tB As String, negBin As String

tB = I64ToBinStr(StrToI64(Op))
negBin = HexToBin(neg)

If BinToHex(BinStrAnd(tB, negBin)) = neg Then
     HaveNot = True
Else
     HaveNot = False
End If

End Function
'*************************************************************************
'**函 数 名：HaveOr
'**输    入：-(String)Op
'**输    出：-(Boolean)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-02-13 22:48:46
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function HaveOr(Op As String) As Boolean
Dim tB As String, this_or_next_Bin As String

tB = I64ToBinStr(StrToI64(Op))
this_or_next_Bin = HexToBin(this_or_next)

If BinToHex(BinStrAnd(tB, this_or_next_Bin)) = this_or_next Then
     HaveOr = True
Else
     HaveOr = False
End If

End Function
'*************************************************************************
'**函 数 名：FixOpId
'**输    入：-(String)opId
'**输    出：-(Long)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-12-27 23:06:12
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function FixOpId(OpID As String) As Long
Dim tB As String, negBin As String, NOTBin As String, tStr As String, tB2 As String

tB = I64ToBinStr(StrToI64(OpID))
negBin = HexToBin(neg)
NOTBin = HexToBin(this_or_next)

't = I64toHexStr(And64b(tI, HexStrToI64(neg)))
If BinToHex(BinStrAnd(tB, negBin)) = neg Then
     FixOpId = -CLng(Val(I64toStrNZ(HexStrToI64(Right(I64toHexStr(StrToI64(OpID)), 7)))))
ElseIf BinToHex(BinStrAnd(tB, NOTBin)) = this_or_next Then      '未测试
     tStr = Right(I64toHexStr(StrToI64(OpID)), 7)
     tB2 = HexToBin(tStr)
     tStr = BinStrOr(tB2, HexToBin(this_or_next_Offset))
     FixOpId = CLng(Val(I64toStrNZ(HexStrToI64(BinToHex(tStr)))))
Else
     FixOpId = CLng(Val(OpID))
End If

End Function
'*************************************************************************
'**函 数 名：RemoveOperationNegations
'**输    入：-(String)Op
'**输    出：-(String)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-02-13 23:26:31
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function RemoveOperationNegations(Op As String) As String
Dim tI As Integer64b

tI = StrToI64(Op)
tI.by(3) = 0
RemoveOperationNegations = I64toStrNZ(tI)

End Function
'*************************************************************************
'**函 数 名：RestoreOpId
'**输    入：-(Long)opId
'**输    出：-(String)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-1-18 21:31:43
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function RestoreOpId(OpID As Long) As String
Dim tB As String, negBin As String, NOTBin As String, tStr As String, tB2 As String

tB = I64ToBinStr(StrToI64(CStr(OpID)))
negBin = HexToBin(neg)
NOTBin = HexToBin(this_or_next)

't = I64toHexStr(And64b(tI, HexStrToI64(neg)))
If OpID < 0 Then
     RestoreOpId = I64toStrNZ(HexStrToI64(BinToHex(BinStrOr(tB, negBin))))
ElseIf BinToHex(BinStrAnd(tB, NOTBin)) = this_or_next_Offset Then      '未测试
     tStr = Right(I64toHexStr(StrToI64(CStr(OpID))), 5)
     tB2 = HexToBin(tStr)
     tStr = BinStrOr(tB2, HexToBin(this_or_next))
     RestoreOpId = I64toStrNZ(HexStrToI64(BinToHex(tStr)))
Else
     RestoreOpId = CStr(OpID)
End If

End Function

'*************************************************************************
'**函 数 名：GetOpStr
'**输    入：-(String)Op
'**输    出：-
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-12-13 21:36:50
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function GetOpStr(Op As String) As String
Dim t As Integer, i As Integer

i = GetOpIndex(CInt(Val(RemoveOperationNegations(Op))))

If i >= 0 And i <= UBound(Operation) Then
     GetOpStr = Operation(i).Op_CSVname
Else
     GetOpStr = RemoveOperationNegations(Op)
End If

End Function

'*************************************************************************
'**函 数 名：GetOpStrWithoutNeg
'**输    入：-(Long)Op
'**输    出：-
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-03-23 15:37:23
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function GetOpStrWithoutNeg(Op As Long) As String
Dim t As Integer, i As Integer

i = GetOpIndex(Op)

If i >= 0 And i <= UBound(Operation) Then
     GetOpStrWithoutNeg = Operation(i).Op_CSVname
Else
     GetOpStrWithoutNeg = CStr(Op)
End If

End Function
'*************************************************************************
'**函 数 名：GetOpIndex
'**输    入：-(Long)Op_ID
'**输    出：-
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-12-15 10:29:50
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function GetOpIndex(Op_ID As Long) As Integer
Dim i As Integer

GetOpIndex = GetID(CStr(Op_ID), False, CStr(Op_ID), -1)

End Function

'*************************************************************************
'**函 数 名：FixParam
'**输    入：(Str)Param_Value,(Str)Hint
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-12-30 17:00:00
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub FixParam(Param_Value As String, Hint As String)
Dim Pid As String, Tag_No As Integer, IsOverFlow As Boolean             'A_P

IsOverFlow = False
GetParamCodeInfo Param_Value, Tag_No, Pid

      Select Case Tag_No
            Case Tag_Item
                If CLng(Pid) > N_Item - 1 Then         '物品
                   IsOverFlow = True
                End If
            Case Tag_Party_Tpl                   '部队模板
                If CLng(Pid) > N_PT - 1 Then
                   IsOverFlow = True
                End If
            Case Tag_Party                       '部队
                If CLng(Pid) > N_Party - 1 Then
                   IsOverFlow = True
                End If
            Case Tag_Troop                       '兵种
                If CLng(Pid) > N_Troop - 1 Then
                   IsOverFlow = True
                End If
            Case Tag_Scene                       '场景
                If CLng(Pid) > N_Scene - 1 Then
                   IsOverFlow = True
                End If
            Case Tag_Map_Icon                    '大地图图标
                If CLng(Pid) > N_MapIcon - 1 Then
                   IsOverFlow = True
                End If
            Case Tag_Sound                       '声音
                If CLng(Pid) > N_Sound - 1 Then
                   IsOverFlow = True
                End If
            Case Tag_Particle_Sys                '粒子系统
                If CLng(Pid) > N_PSys - 1 Then
                   IsOverFlow = True
                End If
            Case Tag_Tableau                     '可变材质
                If CLng(Pid) > N_TabMat - 1 Then
                   IsOverFlow = True
                End If
            Case Tag_Mesh                        '网格模型
                If CLng(Pid) > N_Mesh - 1 Then
                   IsOverFlow = True
                End If
      End Select

If IsOverFlow Then
    Param_Value = getTXTID(Tag_No, 0)
    Hint = PublicTags(Tag_No)
Else
    Hint = ""
End If

End Sub

'*************************************************************************
'**函 数 名：GetParamID
'**输    入：-(String)Value
'**输    出：-(String)
'**功能描述：获得参数所指的ID
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-12-15 10:49:12
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function GetParamID(Value As String) As String
Dim tI As Integer64b

tI = StrToI64(Value)
tI.by(7) = 0
GetParamID = I64toStrNZ(tI)

End Function

'*************************************************************************
'**函 数 名：GetParamType
'**输    入：(Str)Value
'**输    出：(Int)-
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-1-7 16:53:01
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function GetParamType(Value As String) As Integer
Dim tI As Integer64b

tI = StrToI64(Value)
GetParamType = tI.by(7)

End Function

'*************************************************************************
'**函 数 名：QuickGetParamType
'**输    入：(Str)Value
'**输    出：(Int)-
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-03-22 23:46:12
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function QuickGetParamType(Value As String) As Integer
Dim i As Integer, tP As String

tP = Right("0000000000000000000" & Value, 19)
tP = Left(tP, 5)

For i = 1 To 25
    If Val(tP) >= ShortTags(i).X And Val(tP) < ShortTags(i + 1).X Then
        QuickGetParamType = i
        Exit For
    End If
Next i

End Function

'*************************************************************************
'**函 数 名：QuickGetParamCodeInfo
'**输    入：(Str)Value,(Int)Tag_no,(Str)Idx
'**输    出：(Int)-
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-03-23 13:52:16
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub QuickGetParamCodeInfo(Value As String, Tag_No As Integer, Idx As String)
Dim i As Integer, tP As String, tP1 As String, n As Integer, tP2 As String

tP = Right("0000000000000000000" & Value, 19)
tP1 = Left(tP, 5)
tP2 = Right(tP, 7)

Tag_No = 0

For i = 1 To 25
    If Val(tP1) >= ShortTags(i).X And Val(tP1) < ShortTags(i + 1).X Then
        Tag_No = i
        Idx = CStr(Val(tP2) - ShortTags(i).Y)
        Exit For
    End If
Next i

If Tag_No = 0 Then Idx = Value

End Sub
'*************************************************************************
'**函 数 名：GetParamCodeInfo
'**输    入：(Str)Value,(Int)Tag_No,(Str)ID
'**输    出：-
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-3-21 22:33:43
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub GetParamCodeInfo(ByVal Value As String, Tag_No As Integer, ID As String)
Dim tI As Integer64b

tI = StrToI64(Value)
Tag_No = tI.by(7)
tI.by(7) = 0
ID = I64toStrNZ(tI)

End Sub

'*************************************************************************
'**函 数 名：FixFormula
'**输    入：-(String)Formula
'**输    出：-
'**功能描述：修正表达式
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-1-18 22:53:22
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub FixFormula(Formula As String)
Dim i As Integer

Formula = UCase(Formula)
'符号修正
Formula = Replace(Formula, "“", Chr(34))
Formula = Replace(Formula, "”", Chr(34))
Formula = Replace(Formula, "：", ":")

'参数标签修正
'Formula = Replace(Formula, ":", "lvar_")
'Formula = Replace(Formula, "$", "var_")

    If UCase(Left(Formula, 3)) = "REG" Then
         If Mid(Formula, 4) <> "_" Then
              Formula = Left(Formula, 3) & "_" & Right(Formula, Len(Formula) - 4)
         End If
    End If

'Formula = Replace(Formula, Chr(34) & "lvar_" & Chr(34), Chr(34) & "lvar_0" & Chr(34))
'Formula = Replace(Formula, Chr(34) & "var_" & Chr(34), Chr(34) & "var_0" & Chr(34))
'Formula = Replace(Formula, Chr(34) & "reg_" & Chr(34), Chr(34) & "reg_0" & Chr(34))

End Sub

'*************************************************************************
'**函 数 名：RemoveTag
'**输    入：(Str)Formula
'**输    出：(Str)-
'**功能描述：去除参数表达式Tag
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-02-07 11:34:11
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function RemoveTag(Formula As String) As String
Dim i As Integer

RemoveTag = Formula

For i = 1 To Len(Formula)
     If Mid(Formula, i, 1) = "_" Then
         RemoveTag = Right(Formula, Len(Formula) - i)
         Exit For
     End If
Next i

End Function

'*************************************************************************
'**函 数 名：ExportParam
'**输    入：(String)FixedParam
'**输    出：(Str)-
'**功能描述：包装参数
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-2-6 23:44:12
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function ExportParam(FixedParam As String) As Type_Param     'A_P
Dim Tag_No As Integer, Idx As Integer

Tag_No = GetParamTagNo(FixedParam, Idx)

If Tag_No = 0 Then
     ExportParam.Value = FixedParam
     ExportParam.strID = ""
Else
     ExportParam.Value = getTXTID(Tag_No, CLng(Idx), ExportParam.strID)
End If
    
End Function

'*************************************************************************
'**函 数 名：IsOptional
'**输    入：(Long)opId,(Integer)Param_no
'**输    出：(Boolean)-
'**功能描述：判断参数是否可选
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-01-26 19:34:11
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function IsOptional(OpID As Long, Param_No As Integer) As Boolean
Dim OpIndex As Integer

IsOptional = False
OpIndex = GetOpIndex(OpID)

If OpIndex >= 0 Then
   If Operation(OpIndex).ParaNum >= Param_No Then
        If InStr(1, Operation(OpIndex).Para(Param_No).Para_Type, "#") Then
           IsOptional = True
        End If
   End If
End If

End Function

'*************************************************************************
'**函 数 名：SwapAct
'**输    入：(Type_Op_Block)Act1,(Type_Op_Block)Act2
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-1-19 13:43:12
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub SwapAct(Act1 As Type_Op_Block, Act2 As Type_Op_Block)
Dim tAct As Type_Op_Block

tAct = Act1
Act1 = Act2
Act2 = tAct

End Sub

'*************************************************************************
'**函 数 名：GetParamTagNo
'**输    入：(Str)ParamFormula,(Int)Idx
'**输    出：(Int)-
'**功能描述：从Tag_Idx形式的参数中获得Tag及参数idx
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-02-07 12:39:53
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function GetParamTagNo(ParamFormula As String, Idx As Integer) As Integer
Dim i As Integer, Tag As String

For i = 1 To Len(ParamFormula)
     If Mid(ParamFormula, i, 1) = "_" Then
         Tag = Left(ParamFormula, i - 1)
         Idx = CInt(Val(Right(ParamFormula, Len(ParamFormula) - i)))
         Exit For
     End If
Next i

For i = 1 To UBound(Tags)
     If Tags(i) = Tag Then GetParamTagNo = i
Next i

End Function

'*************************************************************************
'**函 数 名：GetTagNo
'**输    入：(Str)Fml
'**输    出：(Int)-
'**功能描述：获得参数Tag的序号
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-02-07 18:12:33
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function GetTagNo(Fml As String) As Integer
Dim i As Integer, Tag As String

GetTagNo = 0
For i = 1 To Len(Fml)
    If Mid(Fml, i, 1) = "_" Then
         Tag = Left(Fml, i - 1)
         Exit For
    End If
Next i

For i = 1 To UBound(Tags)
     If Tags(i) = Tag Then
         GetTagNo = i
         Exit For
     End If
Next i

End Function

'*************************************************************************
'**函 数 名：StandardizeParam
'**输    入： TagType As Integer, Idx As Integer
'**输    出：(Str)-
'**功能描述：将参数转换成Tag_Idx形式的标准参数形式
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-02-07 18:12:33
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function StandardizeParam(TagType As Integer, Idx As Integer) As String

If TagType > 0 And TagType < UBound(Tags) Then
     StandardizeParam = Tags(TagType) & "_" & CStr(Idx)
ElseIf TagType = 0 Then
     StandardizeParam = CStr(Idx)
End If

End Function

'*************************************************************************
'**函 数 名：RemoveParamBrackets
'**输    入：(Str)ParamFml
'**输    出：(Str)-
'**功能描述：将引用参数前的括号部分去掉
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-02-07 23:17:22
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function RemoveParamBrackets(ParamFml As String) As String
Dim i As Integer, n As Integer

RemoveParamBrackets = ParamFml
    For i = 1 To Len(ParamFml)
        If Mid(ParamFml, i, 1) = "(" Then
             For n = i To Len(ParamFml)
                  If Mid(ParamFml, n, 1) = ")" Then
                     RemoveParamBrackets = Left(ParamFml, i - 1) & Right(ParamFml, Len(ParamFml) - n)
                     Exit For
                  End If
             Next n
        Exit For
        End If
    Next i
    
End Function

'*************************************************************************
'**函 数 名：IsInt
'**输    入：(Str)ParamFml
'**输    出：(Boolean)-
'**功能描述：判断参数是否为整型
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-02-12 22:17:41
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function IsInt(ParmFml As String) As Boolean
Dim i As Integer

IsInt = True
For i = 1 To Len(ParmFml)
     If Not (IsNumeric(Mid(ParmFml, i, 1)) Or Mid(ParmFml, i, 1) = "-") Then
          IsInt = False
          Exit For
     End If
Next i
                          
End Function

'*************************************************************************
'**函 数 名：IsBin
'**输    入：(Str)ParamFml
'**输    出：(Boolean)-
'**功能描述：判断参数是否为整型
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-02-12 22:20:30
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function IsBin(ParmFml As String) As Boolean
Dim i As Integer

IsBin = True
For i = 1 To Len(ParmFml)
     If Not (Mid(ParmFml, i, 1) = "0" Or Mid(ParmFml, i, 1) = "1" Or Mid(ParmFml, i, 1) = "-") Then
          IsBin = False
          Exit For
     End If
Next i
                          
End Function

'*************************************************************************
'**函 数 名：IsHex
'**输    入：(Str)ParamFml
'**输    出：(Boolean)-
'**功能描述：判断参数是否为十六进制整型
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-02-16 20:17:12
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function IsHex(ParmFml As String) As Boolean
Dim i As Integer

IsHex = True
For i = 1 To Len(ParmFml)
     If Not (IsNumeric(Mid(ParmFml, i, 1)) Or Mid(ParmFml, i, 1) = "-" Or Mid(ParmFml, i, 1) = "A" Or Mid(ParmFml, i, 1) = "B" _
     Or Mid(ParmFml, i, 1) = "C" Or Mid(ParmFml, i, 1) = "D" Or Mid(ParmFml, i, 1) = "E" Or Mid(ParmFml, i, 1) = "F") Then
          IsHex = False
          Exit For
     End If
Next i
                          
End Function

'*************************************************************************
'**函 数 名：SwapTrigger
'**输    入：(Type_Trigger)Trg1,(Type_Trigger)Trg2
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-1-19 13:43:19
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub SwapTrigger(Trg1 As Type_Trigger, Trg2 As Type_Trigger)
Dim tTrg As Type_Trigger

tTrg = Trg1
Trg1 = Trg2
Trg2 = tTrg

End Sub


'*************************************************************************
'**函 数 名：IsInGroup
'**输    入：(Int)Group(),(Int)OpIndex
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-03-15 13:17:49
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function IsInGroup(Group() As Integer, OpIndex As Integer) As Boolean
Dim i As Integer

For i = LBound(Group) To UBound(Group)
    If Group(i) = OpIndex Then
       IsInGroup = True
       Exit For
    End If
Next i

End Function

'*************************************************************************
'**函 数 名：CheckOverFlow
'**输    入：(String)Value
'**输    出：无
'**功能描述：检测数值是否溢出
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-03-18 0:12:44
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function CheckOverFlow(Value As String) As Boolean

If Len(Value) > 14 Then
     CheckOverFlow = True
Else
     CheckOverFlow = False
End If

End Function

'*************************************************************************
'**函 数 名：RegisterOperations
'**输    入：(String)FileName
'**输    出：无
'**功能描述：注册操作
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-05-26 20:33:52
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub RegisterOperations(ByVal FileName As String)
Dim strTem As String, lngTem As Long, intTem As Integer, i As Long, Op_Count As Long, j As Long
Op_Count = ReadInt(FileName, "Info", "Op_Count")

If Op_Count > 0 Then
  ReDim Operation(0 To Op_Count - 1)
  For i = 0 To Op_Count - 1
    With Operation(i)
      .OpID = ReadInt(FileName, "Op_" & i, "ID")
      
      .Op_name = ReadString(FileName, "Op_" & i, "Name", 255)
      If Len(.Op_name) > Max_Op_Len Then Max_Op_Len = Len(.Op_name)
      .Pseudo = ReadString(FileName, "Op_" & i, "Pseudo", 255)
      .Op_CSVname = ReadString(FileName, "Op_" & i, "Description", 255)
     ' If .Op_CSVname = "" Then .Op_CSVname = .Op_name
      lngTem = ReadInt(FileName, "Op_" & i, "Type")
      .Type = CInt(lngTem)
      
      .ParaNum = ReadInt(FileName, "Op_" & i, "Param_Count")
      
      If .ParaNum > 0 Then
        ReDim .Para(1 To .ParaNum)
        For j = 1 To .ParaNum
          .Para(j).Para_Type = ReadString(FileName, "Op_" & i, "Param_" & j & ".Type", 255)
          .Para(j).Value = ReadString(FileName, "Op_" & i, "Param_" & j & ".Description", 255)
        Next j
      Else
        ReDim .Para(0)
      End If
    End With
  Next i
End If

End Sub

'*************************************************************************
'**函 数 名：OutputOperations
'**输    入：(String)FileName
'**输    出：无
'**功能描述：存储操作
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-05-26 21:02:34
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub OutputOperations(ByVal FileName As String)
Dim strTem As String, lngTem As Long, intTem As Integer, i As Long, Op_Count As Long, j As Long
WriteString FileName, "Info", "Op_Count", CStr(UBound(Operation) + 1)

If UBound(Operation) >= 0 Then
  For i = 0 To UBound(Operation)
    With Operation(i)
      WriteString FileName, "Op_" & i, "ID", CStr(.OpID)
      WriteString FileName, "Op_" & i, "Name", .Op_name
      WriteString FileName, "Op_" & i, "Pseudo", .Pseudo
      WriteString FileName, "Op_" & i, "Type", CStr(.Type)
      WriteString FileName, "Op_" & i, "Description", CStr(.Op_CSVname)
      
      WriteString FileName, "Op_" & i, "Param_Count", CStr(.ParaNum)
      
      If .ParaNum > 0 Then
        For j = 1 To .ParaNum
          WriteString FileName, "Op_" & i, "Param_" & j & ".Type", .Para(j).Para_Type
          WriteString FileName, "Op_" & i, "Param_" & j & ".Description", .Para(j).Value
        Next j

      End If
    End With
  Next i
End If

End Sub

'*************************************************************************
'**函 数 名：LoadOperations
'**输    入：-
'**输    出：无
'**功能描述：载入操作
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-05-26 22:06:23
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub LoadOperations()
Dim strTem As String, i As Long
    'Load Operations
    strTem = ReadString(MnBInfo.iniSetting, "Settings", "Op_Set", 250)
    If strTem = "" Then strTem = "default"
    If strTem <> "default" And FileExists(App.Path & "\" & strTem & ".op.ini") Then
      RegisterOperations App.Path & "\" & strTem & ".op.ini"
    Else
      InitOperations
    End If
    
    MnBInfo.Op_Set = strTem
    frmData.LstOp.ListItems.Clear
    For i = 0 To UBound(Operation)
      AddIndex i, CStr(Operation(i).OpID)
      frmData.LstOp.ListItems.Add , "Op_" & Operation(i).Op_name, CStr(i)
    Next i
    
    InitControlOperationGroup
    InitOptionalParamGroup
End Sub

Public Function GetOpIndexbyName(strID As String) As Long
On Error GoTo EL
Dim Index_Now As String

Index_Now = frmData.LstOp.ListItems("Op_" & strID).Text

GetOpIndexbyName = CLng(Val(Index_Now))
Exit Function

EL:

If Err.Number = ERROR_NOELEMENT Then
    GetOpIndexbyName = -1
Else
    GetOpIndexbyName = -1
    logErr "ModData", "GetOpIndexbyName", Err.Number, Err.Description
End If

End Function

Public Function IsControlOp(ByVal OpID As Long) As Long
If OpID = try_begin Or OpID = try_for_range Or OpID = try_for_range_backwards Or OpID = try_for_parties Or OpID = try_for_agents Or OpID = else_try Then
  IsControlOp = 2
ElseIf OpID = try_end Then
  IsControlOp = 1
Else
  IsControlOp = 0
End If
End Function

