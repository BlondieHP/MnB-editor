VERSION 5.00
Begin VB.Form frmLine 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "文本行"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Tag             =   "debg_2"
   Begin VB.PictureBox PicToolBar 
      Align           =   2  'Align Bottom
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   10365
      TabIndex        =   1
      Top             =   3045
      Width           =   10425
      Begin VB.CommandButton CCopy 
         BackColor       =   &H00C0FFC0&
         Caption         =   "复制到剪切板(&C)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   30
         Width           =   2055
      End
      Begin VB.CommandButton CApply 
         BackColor       =   &H008080FF&
         Caption         =   "导入到编辑器(&A)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   30
         Width           =   2055
      End
   End
   Begin VB.TextBox txtLine 
      Height          =   5655
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   10335
   End
End
Attribute VB_Name = "frmLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public frmTag As String, Idx As Long

Private Sub CApply_Click() '-A_E

If MsgBox(PublicMsgs(99), vbInformation + vbYesNo + vbDefaultButton2, PublicMsgs(100)) = vbYes Then
Select Case frmTag
      Case "edit_1"
         If CheckExist(EditInfo_TroopsCount, Idx) Then
           ReadTroopLine txtLine.Text, Trps(Idx)
         Else
           ReadTroopLine txtLine.Text, CurrentTrp
         End If
         
      Case "edit_2"
         If CheckExist(EditInfo_ItemsCount, Idx) Then
           ReadItemLine txtLine.Text, itm(Idx)
         Else
           ReadItemLine txtLine.Text, CurrentItm
         End If
         
      Case "edit_3"
         If CheckExist(EditInfo_PartiesCount, Idx) Then
           ReadPartyLine txtLine.Text, Parties(Idx)
         Else
           ReadPartyLine txtLine.Text, CurrentParty
         End If
         
      Case "edit_4"
          If CheckExist(EditInfo_PartyTemplatesCount, Idx) Then
           ReadPTLine txtLine.Text, PTs(Idx)
          Else
           ReadPTLine txtLine.Text, CurPartyTemplate
          End If

      Case "edit_5"
          If CheckExist(EditInfo_FactionsCount, Idx) Then
           ReadFactionLine txtLine.Text, Factions(Idx)
          Else
           ReadFactionLine txtLine.Text, CurrentFaction
          End If

      Case "edit_6"
          If CheckExist(EditInfo_ScenesCount, Idx) Then
           ReadSceneLine txtLine.Text, Scenes(Idx)
          Else
           ReadSceneLine txtLine.Text, CurrentScene
          End If

      Case "edit_7"
          If CheckExist(EditInfo_MapIconsCount, Idx) Then
           ReadMapIconLine txtLine.Text, MapIcons(Idx)
          Else
           ReadMapIconLine txtLine.Text, CurrentMapIcon
          End If
          
      Case "edit_8"
          If CheckExist(EditInfo_PSysCount, Idx) Then
           ReadPSysLine txtLine.Text, PSys(Idx)
          Else
           ReadPSysLine txtLine.Text, CurrentPSys
          End If
          
      Case "edit_9"
          If CheckExist(EditInfo_SoundsCount, Idx) Then
           ReadSoundLine txtLine.Text, Sounds(Idx)
          Else
           ReadSoundLine txtLine.Text, CurrentSound
          End If
          
      Case "edit_10"
          If CheckExist(EditInfo_SoundRessCount, Idx) Then
           ReadSoundResLine txtLine.Text, SoundRess(Idx)
          Else
           ReadSoundResLine txtLine.Text, CurrentSoundRes
          End If
          
      Case "edit_11"
          If CheckExist(EditInfo_TabMatCount, Idx) Then
           ReadTabMatLine txtLine.Text, TabMat(Idx)
          Else
           ReadTabMatLine txtLine.Text, CurrentTabMat
          End If
          
      Case "edit_12"
          If CheckExist(EditInfo_MeshCount, Idx) Then
           ReadMeshLine txtLine.Text, Mesh(Idx)
          Else
           ReadMeshLine txtLine.Text, CurrentMesh
          End If
      
      Case "edit_13"
          If CheckExist(EditInfo_TimeTrgCount, Idx) Then
           readTriggerLine txtLine.Text, TimeTrg(Idx)
          Else
           readTriggerLine txtLine.Text, CurrentTimeTrg
          End If
          
      Case Else
           UnLoad Me
End Select

If frmMain.HaveForm Then frmMain.Display.ReLoadInfo

MsgBox PublicMsgs(101), vbInformation, PublicMsgs(100)

End If

End Sub

Private Sub CCopy_Click()
Clipboard.Clear
Clipboard.SetText txtLine.Text
End Sub

Private Sub Form_Deactivate()
Me.ZOrder
End Sub

Public Sub ShowTxtLine(Tag As String, Index As Long)

frmTag = Tag
Idx = Index

LoadTxtLine

Me.Show
End Sub

Private Sub Form_Load()

TranslateForm Me
End Sub

Private Sub Form_Resize()
On Error Resume Next

txtLine.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - PicToolBar.Height

End Sub

Private Sub LoadTxtLine() '-A_E
Dim s As String

Select Case frmTag
      Case "edit_1"
          txtLine.Text = OutputTroopLine(Idx)
          Me.Caption = PublicEditors(1) & ":" & strIIf(Idx <= -1, CurrentTrp.strID, Trps(Abs(Idx)).strID)
      Case "edit_2"
          txtLine.Text = OutputItemLine(Idx)
          Me.Caption = PublicEditors(2) & ":" & strIIf(Idx <= -1, CurrentItm.dbName, itm(Abs(Idx)).dbName)
      Case "edit_3"
          txtLine.Text = OutputPartyLine(Idx)
          Me.Caption = PublicEditors(3) & ":" & strIIf(Idx <= -1, CurrentParty.strID, Parties(Abs(Idx)).strID)
      Case "edit_4"
          txtLine.Text = OutputPTLine(Idx)
          Me.Caption = PublicEditors(4) & ":" & strIIf(Idx <= -1, CurPartyTemplate.ptID, PTs(Abs(Idx)).ptID)
      Case "edit_5"
          txtLine.Text = OutputFactionLine(Idx)
          Me.Caption = PublicEditors(5) & ":" & strIIf(Idx <= -1, CurrentFaction.strID, Factions(Abs(Idx)).strID)
      Case "edit_6"
          txtLine.Text = OutputSceneLine(Idx)
          Me.Caption = PublicEditors(6) & ":" & strIIf(Idx <= -1, CurrentScene.strID, Scenes(Abs(Idx)).strID)
      Case "edit_7"
          txtLine.Text = OutputMapIconLine(Idx)
          Me.Caption = PublicEditors(7) & ":" & strIIf(Idx <= -1, CurrentMapIcon.strID, MapIcons(Abs(Idx)).strID)
      Case "edit_8"
          txtLine.Text = OutputPSysLine(Idx)
          Me.Caption = PublicEditors(8) & ":" & strIIf(Idx <= -1, CurrentPSys.strID, PSys(Abs(Idx)).strID)
      Case "edit_9"
          txtLine.Text = OutputSoundLine(Idx)
          Me.Caption = PublicEditors(9) & ":" & strIIf(Idx <= -1, CurrentSound.sndName, Sounds(Abs(Idx)).sndName)
      Case "edit_10"
          txtLine.Text = OutputSoundResLine(Idx)
          Me.Caption = PublicEditors(10) & ":" & strIIf(Idx <= -1, CurrentSoundRes.sndName, SoundRess(Abs(Idx)).sndName)
      Case "edit_11"
          txtLine.Text = OutputTabMatLine(Idx)
          Me.Caption = PublicEditors(11) & ":" & strIIf(Idx <= -1, CurrentTabMat.strID, TabMat(Abs(Idx)).strID)
      Case "edit_12"
          txtLine.Text = OutputMeshLine(Idx)
          Me.Caption = PublicEditors(12) & ":" & strIIf(Idx <= -1, CurrentMesh.strID, Mesh(Abs(Idx)).strID)
      Case "edit_13"
          txtLine.Text = OutputTimeTriggerLine(Idx)
          Me.Caption = PublicEditors(13) & ":" & strIIf(Idx <= -1, "", "")
      Case Else
           UnLoad Me
End Select
End Sub
