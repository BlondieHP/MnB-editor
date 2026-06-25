VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmModule 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "剧本："
   ClientHeight    =   4635
   ClientLeft      =   5820
   ClientTop       =   3750
   ClientWidth     =   10575
   ControlBox      =   0   'False
   Icon            =   "frmModule.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   309
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   705
   Begin VB.PictureBox ModPic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4125
      Left            =   240
      ScaleHeight     =   273
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   338
      TabIndex        =   0
      Top             =   240
      Width           =   5100
      Begin MSComctlLib.ImageList IL1 
         Left            =   1320
         Top             =   2760
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule.frx":000C
               Key             =   "edit_1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule.frx":08E6
               Key             =   "edit_2"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule.frx":11C0
               Key             =   "edit_3"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule.frx":14DA
               Key             =   "edit_4"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule.frx":17F4
               Key             =   "edit_5"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule.frx":20CE
               Key             =   "edit_6"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule.frx":29A8
               Key             =   "tool_1"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule.frx":3282
               Key             =   "edit_7"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ListView LstEditors 
      Height          =   4095
      Left            =   5520
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   7223
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "IL1"
      ForeColor       =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frmModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

publicInit

Me.Caption = "剧本：" & MnBInfo.ModName

frmTip.ShowTip "载入中,请稍后..."

ModPic.Picture = LoadPicture(MnBInfo.ModPath & "\Main.bmp")

'Load Sounds
LoadSoundFile MnBInfo.ModPath & "\sounds.txt"

'Load Items
LoadItemFile MnBInfo.ModPath & "\item_kinds1.txt"

LoadItemCSVFile MnBInfo.ModPath & "\languages\" & MnBInfo.Language & "\item_kinds.csv"

'Load Troops
LoadTroopFile MnBInfo.ModPath & "\troops.txt"

LoadTroopCSVFile MnBInfo.ModPath & "\languages\" & MnBInfo.Language & "\troops.csv"

'Load Party_Templates
LoadPTFile MnBInfo.ModPath & "\party_templates.txt"

LoadPartyTemplateCSVFile MnBInfo.ModPath & "\languages\" & MnBInfo.Language & "\party_templates.csv"

'Load Parties
LoadPartyFile MnBInfo.ModPath & "\parties.txt"

LoadPartyCSVFile MnBInfo.ModPath & "\languages\" & MnBInfo.Language & "\parties.csv"

'Load Factions
LoadFactionFile MnBInfo.ModPath & "\factions.txt"

LoadFactionCSVFile MnBInfo.ModPath & "\languages\" & MnBInfo.Language & "\factions.csv"

'Load IModifiers
LoadIModFile MnBInfo.MBHome & "\Data\item_modifiers.txt"
LoadIModCSVFile MnBInfo.ModPath & "\languages\" & MnBInfo.Language & "\item_modifiers.csv"

'Load Scenes
LoadSceneFile MnBInfo.ModPath & "\scenes.txt"

'Load Particle System
LoadPSysFile MnBInfo.ModPath & "\particle_systems.txt"

'Load Map Icons
LoadMapIconFile MnBInfo.ModPath & "\map_icons.txt"

'Load Editors
InitEditorsListView

'Load Operations
InitOperations

frmTip.HideTip

frmMain.mSave.Enabled = True
frmMain.mSaveAs.Enabled = True
frmMain.mEditor.Enabled = True
frmMain.mTool.Enabled = True

Me.Show
End Sub

Private Sub InitEditorsListView()
Dim n As Integer

With LstEditors
  .View = lvwIcon
  .Sorted = False
  .ListItems.Clear
  .ColumnHeaders.Clear
  .LabelEdit = lvwManual
  .Checkboxes = False
  .MultiSelect = False
  .HideSelection = False
  .Visible = True
  .HideColumnHeaders = True
End With

LstEditors.ColumnHeaders.Add , , "编辑器名称"

LstEditors.ListItems.Add , "edit_1", "兵种", "edit_1"
LstEditors.ListItems.Add , "edit_2", "物品", "edit_2"
LstEditors.ListItems.Add , "edit_3", "部队", "edit_3"
LstEditors.ListItems.Add , "edit_4", "部队模板", "edit_4"
LstEditors.ListItems.Add , "edit_5", "阵营", "edit_5"
LstEditors.ListItems.Add , "edit_6", "场景", "edit_6"
LstEditors.ListItems.Add , "edit_7", "大地图图标", "edit_7"
LstEditors.ListItems.Add , "tool_1", "卡拉迪亚地图", "tool_1"
End Sub

Private Sub LstEditors_DblClick()

Select Case LstEditors.SelectedItem.Key
      Case "edit_1"
          frmTroops.Show
      Case "edit_2"
          frmItems.Show
      Case "edit_3"
          frmParties.Show
      Case "edit_4"
          frmParty_Templates.Show
      Case "edit_5"
          frmFactions.Show
      Case "edit_6"
          frmScenes.Show
      Case "edit_7"
          frmMap_Icons.Show
      Case "tool_1"
          frmMap.Show
End Select

End Sub

