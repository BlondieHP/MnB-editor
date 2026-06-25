VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSounds 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "声音编辑器"
   ClientHeight    =   9375
   ClientLeft      =   1755
   ClientTop       =   180
   ClientWidth     =   14355
   Icon            =   "frmSounds.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   14355
   ShowInTaskbar   =   0   'False
   Tag             =   "edit_9"
   Begin VB.Frame FraProps 
      Caption         =   "声音属性"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   8535
      Index           =   0
      Left            =   6120
      TabIndex        =   13
      Top             =   120
      Width           =   8055
      Begin VB.Frame Frame1 
         Caption         =   "声音资源"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   4215
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   4080
         Width           =   7455
         Begin MSComctlLib.ListView LstSoundRess 
            Height          =   3135
            Left            =   240
            TabIndex        =   22
            Top             =   840
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   5530
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.ComboBox cbSoundRess 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   885
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   360
            Width           =   6135
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "资源:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   6
            Left            =   360
            TabIndex        =   21
            Top             =   480
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "特性"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   2295
         Index           =   2
         Left            =   240
         TabIndex        =   17
         Top             =   1680
         Width           =   7455
         Begin MSComctlLib.Slider sldPriority 
            Height          =   375
            Left            =   3720
            TabIndex        =   27
            Top             =   480
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   661
            _Version        =   393216
            LargeChange     =   2
         End
         Begin VB.CheckBox chkFlags 
            Caption         =   "chkFlags"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   2
            Left            =   480
            TabIndex        =   24
            Top             =   1440
            Width           =   2175
         End
         Begin VB.CheckBox chkFlags 
            Caption         =   "chkFlags"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   23
            Top             =   960
            Width           =   2175
         End
         Begin VB.CheckBox chkFlags 
            Caption         =   "chkFlags"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   18
            Top             =   480
            Width           =   2175
         End
         Begin MSComctlLib.Slider sldVolume 
            Height          =   375
            Left            =   3720
            TabIndex        =   28
            Top             =   1200
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   661
            _Version        =   393216
            LargeChange     =   3
            Max             =   15
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "音量:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   2
            Left            =   3120
            TabIndex        =   26
            Top             =   1245
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "优先级:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   180
            Index           =   1
            Left            =   3000
            TabIndex        =   25
            Top             =   525
            Width           =   690
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "识别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   1215
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   7455
         Begin VB.TextBox txtPropBag 
            Height          =   375
            Index           =   0
            Left            =   1320
            TabIndex        =   15
            Top             =   480
            Width           =   5655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "声音名:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   180
            Index           =   0
            Left            =   600
            TabIndex        =   16
            Top             =   560
            Width           =   690
         End
      End
   End
   Begin VB.CommandButton CApply 
      BackColor       =   &H000000FF&
      Caption         =   "套用(&A)"
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
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8805
      Width           =   2175
   End
   Begin VB.CommandButton CReset 
      BackColor       =   &H00C0C0C0&
      Caption         =   "重置(&R)"
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
      Left            =   9720
      TabIndex        =   4
      Top             =   8805
      Width           =   2175
   End
   Begin VB.TextBox txtQuery 
      Height          =   330
      Left            =   600
      TabIndex        =   3
      Top             =   120
      Width           =   4575
   End
   Begin VB.CommandButton CQuery 
      Caption         =   "↓"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   5160
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton CQuery 
      Caption         =   "↑"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   5520
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton COutputLine 
      BackColor       =   &H0080FF80&
      Caption         =   "输出当前声音(&O)"
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8760
      Width           =   2415
   End
   Begin MSComctlLib.ListView LstSounds 
      Height          =   8415
      Left            =   120
      TabIndex        =   6
      Top             =   525
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   14843
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
      Picture         =   "frmSounds.frx":08CA
   End
   Begin VB.Label cCMD 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "删除(&E)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   0
      Left            =   5160
      MouseIcon       =   "frmSounds.frx":4F37
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   9060
      Width           =   705
   End
   Begin VB.Label cCMD 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "创建(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   180
      Index           =   1
      Left            =   4320
      MouseIcon       =   "frmSounds.frx":5241
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   9060
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "声音数:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   135
      TabIndex        =   10
      Top             =   9060
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "查询:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   120
      TabIndex        =   9
      Top             =   195
      Width           =   495
   End
   Begin VB.Label cCMD 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "上移(&U)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   2
      Left            =   2640
      MouseIcon       =   "frmSounds.frx":554B
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   9060
      Width           =   705
   End
   Begin VB.Label cCMD 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "下移(&D)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   3
      Left            =   3480
      MouseIcon       =   "frmSounds.frx":5855
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   9060
      Width           =   705
   End
End
Attribute VB_Name = "frmSounds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CustomActive As Boolean

Private Sub CApply_Click()
Dim q As Boolean
If MsgBox(PublicMsgs(1), vbYesNo + vbDefaultButton2 + vbInformation, PublicMsgs(0)) = vbYes Then

    If UCase(Sounds(CurrentSoundID).sndName) <> UCase(CurrentSound.sndName) Then             '外引
        q = ChangeStrID(Sounds(CurrentSoundID).sndName, CurrentSound.sndName)
        If Not q Then
           MsgBox ActiveString(PublicMsgs(91), CurrentSound.sndName), vbCritical, PublicMsgs(19)
           Exit Sub
        End If
    End If
    
    Sounds(CurrentSoundID) = CurrentSound
    LstSounds.ListItems(CurrentSoundID + 1).SubItems(1) = Sounds(CurrentSoundID).sndName
    
    CurrentSound = Sounds(CurrentSoundID)
    LoadSoundInfo
End If
End Sub

Private Sub cbSoundRess_Click()

If CustomActive Then

  If LstSoundRess.SelectedItem.Index < LstSoundRess.ListItems.Count Then
    With CurrentSound.Resource(LstSoundRess.SelectedItem.Index)
         .ID = cbSoundRess.ListIndex - 1
    
         If .ID = -1 Then
            DeleteCurrentResource LstSoundRess.SelectedItem.Index
            LoadSoundRessListView
         Else
             LstSoundRess.SelectedItem.Text = .ID
             LstSoundRess.SelectedItem.SubItems(1) = SoundRess(.ID).sndName
             .strID = SoundRess(.ID).sndName
         End If
    End With
  Else
    CreateCurrentResource cbSoundRess.ListIndex - 1
    If cbSoundRess.ListIndex - 1 >= 0 Then
       LoadSoundRessListView
    End If
  End If
  
End If
End Sub

Private Sub cbSoundRess_Scroll()
Call cbSoundRess_Click
End Sub

Private Sub cCMD_Click(Index As Integer)
Dim i As Long, oItem As ListItem, j As Long
Select Case Index
       Case 0
           If Sounds(CurrentSoundID).Edit Then
           
           If MsgBox(ActiveString(PublicMsgs(2), Sounds(CurrentSoundID).sndName), vbYesNo + vbExclamation, ActiveString(PublicMsgs(3), PublicEditors(GetEditorIndex(Me.Tag)))) = vbYes Then
              
              DelIndex Sounds(CurrentSoundID).sndName
              If CurrentSoundID < N_Sound - 1 Then
                For i = CurrentSoundID To N_Sound - 2 Step 1
                    ChangeID Sounds(i + 1).sndName, Sounds(i + 1).ID - 1
                    j = Sounds(i).ID
                    Sounds(i) = Sounds(i + 1)
                    Sounds(i).ID = j
                    LstSounds.ListItems(i + 1).SubItems(1) = LstSounds.ListItems(i + 2).SubItems(1)
                Next i
                
                ReDim Preserve Sounds(N_Sound - 2)
                LstSounds.ListItems.Remove N_Sound
                N_Sound = N_Sound - 1
                
              Else
                ReDim Preserve Sounds(N_Sound - 2)
                LstSounds.ListItems.Remove N_Sound
                
                N_Sound = N_Sound - 1
                CurrentSoundID = N_Sound - 1
                
              End If
               
               LstSounds_ItemClick LstSounds.ListItems(CurrentSoundID + 1)
               LstSounds.ListItems(CurrentSoundID + 1).Selected = True
               LstSounds.ListItems(CurrentSoundID + 1).EnsureVisible
           
           End If
           Else
              MsgBox ActiveString(PublicMsgs(4), Sounds(CurrentSoundID).sndName), vbCritical, ActiveString(PublicMsgs(3), PublicEditors(GetEditorIndex(Me.Tag)))

           End If
       Case 1

         If MsgBox(ActiveString(PublicMsgs(5), Sounds(CurrentSoundID).sndName, PublicEditors(GetEditorIndex(Me.Tag))), vbYesNo + vbInformation, ActiveString(PublicMsgs(9), PublicEditors(GetEditorIndex(Me.Tag)))) = vbYes Then
           
           If AddIndex(N_Sound, Sounds(CurrentSoundID).sndName & "_New") Then
           ReDim Preserve Sounds(N_Sound)
           N_Sound = N_Sound + 1
           Sounds(N_Sound - 1) = Sounds(CurrentSoundID)
           With Sounds(N_Sound - 1)
                 .ID = N_Sound - 1
                 .sndName = .sndName & "_New"
                 .Edit = True
           End With
           
           Set oItem = LstSounds.ListItems.Add(, "mis_" & Sounds(N_Sound - 1).ID, Sounds(N_Sound - 1).ID)
      
                 With oItem
                    .SubItems(1) = Sounds(N_Sound - 1).sndName
                 End With
           LstSounds_ItemClick LstSounds.ListItems(N_Sound)
           LstSounds.ListItems(N_Sound).Selected = True
           LstSounds.ListItems(N_Sound).EnsureVisible
           
           Else
             MsgBox ActiveString(PublicMsgs(90), Sounds(CurrentSoundID).sndName & "_New"), vbCritical, PublicMsgs(19)
           End If
         End If

      Case 2
         If CurrentSoundID > 0 Then
           If Sounds(CurrentSoundID - 1).Edit And Sounds(CurrentSoundID).Edit Then
           
                SwapID Sounds(CurrentSoundID - 1).sndName, Sounds(CurrentSoundID).sndName
                SwapSounds CurrentSoundID - 1, CurrentSoundID
                SwapListItem LstSounds.ListItems(CurrentSoundID), LstSounds.ListItems(CurrentSoundID + 1), 1, True
                
               LstSounds_ItemClick LstSounds.ListItems(CurrentSoundID)
               LstSounds.ListItems(CurrentSoundID + 1).Selected = True
           Else
            MsgBox ActiveString(PublicMsgs(7), Sounds(CurrentSoundID - 1).sndName), vbCritical, ActiveString(PublicMsgs(8), PublicEditors(GetEditorIndex(Me.Tag)))

           End If
         End If
      Case 3
        If CurrentSoundID + 1 <= N_Sound - 1 Then
           If Sounds(CurrentSoundID).Edit And Sounds(CurrentSoundID + 1).Edit Then
           
                SwapID Sounds(CurrentSoundID).sndName, Sounds(CurrentSoundID + 1).sndName
                SwapSounds CurrentSoundID, CurrentSoundID + 1
                SwapListItem LstSounds.ListItems(CurrentSoundID + 1), LstSounds.ListItems(CurrentSoundID + 2), 1, True
                
                LstSounds_ItemClick LstSounds.ListItems(CurrentSoundID + 2)
                LstSounds.ListItems(CurrentSoundID + 1).Selected = True
           Else
              MsgBox ActiveString(PublicMsgs(7), Sounds(CurrentSoundID).sndName), vbCritical, ActiveString(PublicMsgs(8), PublicEditors(GetEditorIndex(Me.Tag)))

           End If
           
        End If
End Select
End Sub




Private Sub chkFlags_Click(Index As Integer)
Dim tI(2) As Integer64b

If CustomActive Then

With CurrentSound
    tI(1) = StrToI64(.Flags)
    tI(2) = HexStrToI64(chkFlags(Index).Tag)
    If chkFlags(Index).Value = 0 Then
      tI(0) = DeleteFlags64b(tI(1), tI(2))
    Else
      tI(0) = AddFlags64b(tI(1), tI(2))
    End If
    .Flags = I64toStrNZ(tI(0))
End With

End If
End Sub




Private Sub COutputLine_Click()
frmLine.ShowTxtLine Me.Tag, -1
End Sub

Private Sub CQuery_Click(Index As Integer)
Dim q As Boolean
q = QueryItem(LstSounds, LstSounds.SelectedItem.Index, txtQuery.Text, CBool(Index))

If q Then
   Call LstSounds_ItemClick(LstSounds.SelectedItem)
End If
End Sub

'*************************************************************************
'**函 数 名：QueryItem
'**输    入：(ListItem)oLV,(Long)Start,(String)QueryString,(Boolean)bReverse
'**输    出：
'**功能描述：进行ListView查询功能
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-11-27 21:19:17
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Private Function QueryItem(oLV As ListView, ByVal Start As Long, ByVal QueryString As String, Optional bReverse As Boolean = False) As Boolean
Dim oItem As ListItem
  With oLV
    Set oItem = FindItem(oLV, Start, "0|1", QueryString, True, vbTextCompare, bReverse)
       If Not oItem Is Nothing Then
       .ListItems(oItem.Index).Selected = True
       .ListItems(oItem.Index).EnsureVisible
       QueryItem = True
       Else
        MsgBox PublicMsgs(11), vbInformation, PublicMsgs(12)
        QueryItem = False
       End If
  End With
    Set oItem = Nothing
End Function

Private Sub CReset_Click()
If MsgBox(PublicMsgs(10), vbYesNo + vbDefaultButton2 + vbInformation, PublicMsgs(0)) = vbYes Then
LstSounds_ItemClick LstSounds.ListItems(CurrentSoundID + 1)
LstSounds.ListItems(CurrentSoundID + 1).Selected = True
LstSounds.ListItems(CurrentSoundID + 1).EnsureVisible
End If
End Sub

Private Sub Form_Load()
CustomActive = False

InitSoundRessCombo
InitSoundsListView
InitSoundRessListView
LoadSoundsList
InitFlagsList

CurrentSoundID = 0
CurrentSound = Sounds(CurrentSoundID)
LoadSoundInfo

TranslateForm Me
InitCMDs

Label2.Caption = Label2.Caption & N_Sound

CustomActive = True
End Sub

Private Sub InitCMDs()
Dim i As Integer

For i = 1 To cCMD.UBound
    cCMD(i).Left = cCMD(i - 1).Left - cCMD(i).Width - 135
Next i

End Sub

Private Sub InitSoundsListView()
Dim n As Integer
n = 2
LstSounds.View = lvwReport
LstSounds.Sorted = False
LstSounds.ListItems.Clear
LstSounds.ColumnHeaders.Clear
LstSounds.SortOrder = lvwAscending
LstSounds.FullRowSelect = True
LstSounds.AllowColumnReorder = False
LstSounds.LabelEdit = lvwManual
LstSounds.Checkboxes = False
LstSounds.GridLines = True
LstSounds.MultiSelect = False
LstSounds.HideSelection = False

LstSounds.ColumnHeaders.Add , , PublicMsgs(13), LstSounds.Width / n / 3.6
LstSounds.ColumnHeaders.Add , , PublicEditors(9) & PublicMsgs(14), LstSounds.Width / n * 1.5

End Sub

Private Sub InitSoundRessListView()
Dim n As Integer
n = 2
LstSoundRess.View = lvwReport
LstSoundRess.Sorted = False
LstSoundRess.ListItems.Clear
LstSoundRess.ColumnHeaders.Clear
LstSoundRess.SortOrder = lvwAscending
LstSoundRess.FullRowSelect = True
LstSoundRess.AllowColumnReorder = False
LstSoundRess.LabelEdit = lvwManual
LstSoundRess.Checkboxes = False
LstSoundRess.GridLines = True
LstSoundRess.MultiSelect = False
LstSoundRess.HideSelection = False

LstSoundRess.ColumnHeaders.Add , , PublicMsgs(13), LstSoundRess.Width / n / 3.6
LstSoundRess.ColumnHeaders.Add , , PublicEditors(10) & PublicMsgs(14), LstSoundRess.Width / n * 1.5

End Sub


'*************************************************************************
'**函 数 名：LoadSoundsList
'**输    入：无
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-11 00:50:24
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub LoadSoundsList()
Dim n As Long, oItem As ListItem

LstSounds.ListItems.Clear
For n = 0 To N_Sound - 1

      Set oItem = LstSounds.ListItems.Add(, "snds_" & Sounds(n).ID, Sounds(n).ID)
      
      With oItem
         .SubItems(1) = Sounds(n).sndName
      End With

Next n

End Sub

'*************************************************************************
'**函 数 名：LoadSoundInfo
'**输    入：-
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-02 14:34:47
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub LoadSoundInfo()
Dim i As Integer, oItem As ListItem, strTem As String, lngTem As Long, fI64 As Integer64b, fI64_NOW As Integer64b
CustomActive = False

With CurrentSound
       txtPropBag(0).Text = .sndName
     
       LoadFlagsList
       
       LstSoundRess.ListItems.Clear
       
       LoadSoundRessListView
End With

CustomActive = True
End Sub

'*************************************************************************
'**函 数 名：InitFlagsList
'**输    入：
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-11 00:48:02
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub InitFlagsList()
Dim i As Integer, strTemArray() As Variant, TemArray() As Variant
strTemArray = Array("2D声效", "循环播放", "随机播放")
TemArray = Array(sf_2d, sf_looping, sf_start_at_random_pos)
         For i = 0 To UBound(strTemArray)
             chkFlags(i).Caption = strTemArray(i)
             chkFlags(i).Tag = TemArray(i)
         Next i

End Sub

Private Sub LstSoundRess_ItemClick(ByVal Item As MSComctlLib.ListItem)

With CurrentSound
 If Item.Index < LstSoundRess.ListItems.Count Then
   If .Resource(Item.Index).ID + 1 > -1 And .Resource(Item.Index).ID + 1 < cbSoundRess.ListCount Then
     cbSoundRess.ListIndex = .Resource(Item.Index).ID + 1
   Else
       .Resource(Item.Index).ID = -1
   End If
   
   If .Resource(Item.Index).ID >= 0 Then

   End If
 Else
   cbSoundRess.ListIndex = 0
 End If
End With

End Sub

Private Sub LstSounds_ItemClick(ByVal Item As MSComctlLib.ListItem)
CurrentSoundID = Val(Item.Text)

CurrentSound = Sounds(CurrentSoundID)
LoadSoundInfo
End Sub


Private Sub sldPriority_Click()
Dim fI64_NOW As Integer64b, lngTem As Long, lngTem2 As Long

If CustomActive Then
fI64_NOW = StrToI64(CurrentSound.Flags)
lngTem = sldPriority.Value
lngTem2 = fI64_NOW.by(0) Mod 16

fI64_NOW.by(0) = lngTem * 16 + lngTem2
CurrentSound.Flags = I64toStrNZ(fI64_NOW)
End If

End Sub

Private Sub sldPriority_Scroll()
Call sldPriority_Click
End Sub

Private Sub sldVolume_Click()
Dim fI64_NOW As Integer64b, lngTem As Long

If CustomActive Then
fI64_NOW = StrToI64(CurrentSound.Flags)
lngTem = sldVolume.Value

fI64_NOW.by(1) = lngTem
CurrentSound.Flags = I64toStrNZ(fI64_NOW)
End If

End Sub

Private Sub txtPropBag_LostFocus(Index As Integer)
If CustomActive Then
With CurrentSound

If Index < 2 Then
   txtPropBag(Index).Text = Replace(txtPropBag(Index).Text, " ", "_")
End If

Select Case Index
      Case 0
         .sndName = txtPropBag(Index).Text
End Select

End With
End If
End Sub

'*************************************************************************
'**函 数 名：LoadFlagsList
'**输    入：
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-04 07:45:31
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub LoadFlagsList()
On Error GoTo EL
Dim i As Integer, fI64_NOW As Integer64b, fI64 As Integer64b, res64b As Integer64b, lngTem As Long

With CurrentSound
    fI64_NOW = StrToI64(.Flags)
    For i = 0 To 2
        fI64 = HexStrToI64(chkFlags(i).Tag)
        res64b = And64b(fI64_NOW, fI64)
        
        If IsEqual64b(res64b, fI64) Then
           chkFlags(i).Value = 1
        Else
           chkFlags(i).Value = 0
        End If
    Next i
End With

lngTem = fI64_NOW.by(0) \ 16
lngTem = IIf(lngTem > 10, 10, lngTem)
lngTem = IIf(lngTem < 0, 0, lngTem)
sldPriority.Value = lngTem

lngTem = fI64_NOW.by(1)
lngTem = IIf(lngTem > 15, 15, lngTem)
lngTem = IIf(lngTem < 0, 0, lngTem)
sldVolume.Value = lngTem

Exit Sub

EL:
    Call logErr("frmSounds", "LoadFlagsList", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**函 数 名：InitSoundRessCombo
'**输    入：
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-01-02 22:10:43
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub InitSoundRessCombo()
Dim j As Long

   cbSoundRess.Clear
   
   cbSoundRess.AddItem PublicTips(0)
   
   For j = 0 To N_SoundRes - 1
         cbSoundRess.AddItem "(" & j & ")" & SoundRess(j).sndName
   Next j
   
End Sub


Private Sub LoadSoundRessListView()
Dim i As Integer, oItem As ListItem, q As Boolean

CustomActive = False
LstSoundRess.ListItems.Clear
With CurrentSound

   If .ResourceCount > 0 Then
     For i = 1 To .ResourceCount
           .Resource(i).ID = GetID(.Resource(i).strID, True, "", -1)
           If .Resource(i).ID = -1 Then q = True
           
           If .Resource(i).ID > -1 Then
                Set oItem = LstSoundRess.ListItems.Add(, , .Resource(i).ID)
                oItem.SubItems(1) = SoundRess(.Resource(i).ID).sndName
           End If
     Next i
     
     If q Then StructureSoundRes -1
     
     LstSoundRess.ListItems.Add , , PublicTips(0)
     
     LstSoundRess.ListItems(1).Selected = True
     LstSoundRess_ItemClick LstSoundRess.ListItems(1)
   Else
     LstSoundRess.ListItems.Add , , PublicTips(0)
     cbSoundRess.ListIndex = 0
     
   End If
End With

CustomActive = True
End Sub

'*************************************************************************
'**函 数 名：DeleteCurrentResource
'**输    入：(Long)ResourceIndex
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-01-03 12:30:26
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub DeleteCurrentResource(ByVal ResourceIndex As Long)
Dim i As Long
With CurrentSound
     If .ResourceCount > 0 Then
          .ResourceCount = .ResourceCount - 1
          If .ResourceCount = 0 Then
             'ReDim .Resource(0)
             ClearResource .Resource(1)
          Else
              For i = ResourceIndex To .ResourceCount
                   .Resource(i) = .Resource(i + 1)
              Next i
              
              ReDim Preserve .Resource(1 To .ResourceCount)
              'ClearResource .Resource(.ResourceCount + 1)
          End If
          
     End If
End With
End Sub

'*************************************************************************
'**函 数 名：CreateCurrentResource
'**输    入：(Long)Trp_ID
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-01-03 12:32:33
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub CreateCurrentResource(ByVal Trp_ID As Long)
Dim i As Long
If Trp_ID > -1 Then
With CurrentSound
     If .ResourceCount > 0 Then
          .ResourceCount = .ResourceCount + 1
          ReDim Preserve .Resource(1 To .ResourceCount)
     Else
          .ResourceCount = 1
          ReDim .Resource(1 To .ResourceCount)
     End If
     
     .Resource(.ResourceCount).ID = Trp_ID
     .Resource(.ResourceCount).strID = SoundRess(Trp_ID).sndName
End With
End If
End Sub

Public Sub ReLoadInfo()
LoadSoundInfo
End Sub
