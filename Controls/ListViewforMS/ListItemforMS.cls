VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListItemforMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private p_Text As String
Public Value As String    'Reserved
Private OpLabel As Integer
Public Locked As Boolean
Private p_Index As Integer
  
Public SubItems As New ListSubItemsforMS
  
Public Indention As Integer

Public Label As Label
Attribute Label.VB_VarHelpID = -1
Public ListView As ListViewforMS

Public Initialized As Boolean

Public Property Get Index() As Integer
Index = p_Index
End Property

Public Property Let Index(ByVal vNewValue As Integer)

p_Index = vNewValue
End Property

Public Property Get ForeColor() As Long
ForeColor = Label.ForeColor
End Property

Public Property Let ForeColor(ByVal vNewValue As Long)
Label.ForeColor = vNewValue

End Property

Public Property Get FontBold() As Boolean
FontBold = Label.FontBold
End Property

Public Property Let FontBold(ByVal vNewValue As Boolean)
Label.FontBold = vNewValue

End Property

Public Property Get Text() As String
Text = p_Text
End Property

Public Property Let Text(ByVal vNewValue As String)
p_Text = vNewValue

End Property

Public Sub Initialize(lIndex As Integer, UserList As ListViewforMS, LabelBox As Label)
Dim i As Integer

Set ListView = UserList
Set Label = LabelBox
Label.Visible = True
p_Index = lIndex
Label.Tag = p_Index & ",0"
    
SubItems.Initialize UserList, Me

Initialized = True
End Sub

Private Sub Class_Terminate()
Set ListView = Nothing
Set Label = Nothing

Set SubItems = Nothing
Initialized = False
End Sub


Private Function FreeLabel() As Label
Dim i As Integer

  For i = 1 To ListView.Label.UBound
    With ListView.Label(i)
      If Not .Visible Then
        Exit For
      End If
    End With
  Next i

  If i > ListView.Label.UBound Then
    Load ListView.Label(i)
    With ListView.Label(i)
      .AutoSize = True
    
    End With
  End If

  FreeLabel = ListView.Label(i)
End Function

Private Sub PurseText()
Dim i As Integer, KeyWord As String, n As Integer

  For i = 1 To SubItems.Count
    With SubItems(i)
      KeyWord = IIf(.KeyWord = "", "{arg_" & i & "}", .KeyWord)
      n = InStr(1, p_Text, KeyWord)
      .Start = ListView.LenW(Left(p_Text, n))
      If n > 0 Then
        .Label.Caption = IIf(.Text <> "", .Text, .KeyWord)
      Else
        .Label.Caption = ""
      End If
    End With
  Next i

End Sub

Private Sub ShowLabel()
Dim i As Integer, j As Integer, KeyWord As String, LastCaption As String, ParamLabel As Label, MainLabel As Label
Dim SubText As String, Selected As Long

Set MainLabel = Label
MainLabel.Caption = p_Text
Selected = MainLabel.BorderStyle
If Selected = 1 Then MainLabel.BorderStyle = 0
If Indention > 0 Then
  MainLabel.Left = ListView.TextWidth(Space(Indention))
Else
  MainLabel.Left = 0
End If
MainLabel.Top = (p_Index - 1) * ListView.Spacing * MainLabel.Height
If Selected = 1 Then MainLabel.BorderStyle = Selected
MainLabel.Visible = True
'--------------------------
'i = MainLabel.Index
'KeyWord = MainLabel.Tag
'--------------------------

For i = 1 To SubItems.Count
  Set ParamLabel = SubItems(i).Label
  SubText = SubItems(i).Label.Caption
  KeyWord = IIf(SubItems(i).KeyWord = "", "{arg_" & i & "}", SubItems(i).KeyWord)
  
  MainLabel.Caption = Replace(MainLabel.Caption, KeyWord, Space(ListView.LenW(SubText)), , 1)

  For j = 1 To SubItems.Count
    If SubItems(j).Start > SubItems(i).Start Then
      SubItems(j).Start = SubItems(j).Start + ListView.LenW(SubText) - ListView.LenW(KeyWord)
    End If
  Next j
Next i

For i = 1 To SubItems.Count
  Set ParamLabel = SubItems(i).Label
    ParamLabel.Top = MainLabel.Top
    
    If SubItems(i).Start > 0 Then
      ParamLabel.Left = MainLabel.Left + ListView.TextWidth(Space(SubItems(i).Start - 1))
      ParamLabel.Visible = True
      ParamLabel.ZOrder
      'ParamLabel.ForeColor = vbRed
    Else
      'ParamLabel.Visible = False
      ParamLabel.Top = -1000
      ParamLabel.Left = -1000
    End If
Next i

End Sub

Public Sub Refresh(Optional RedrawControl As Boolean = True)
PurseText

ShowLabel

If RedrawControl Then ListView.MeasureFrame
End Sub

Public Function AssignValue(vNewValue As String) As Boolean
Dim Default() As String, OpID As Long, OpIndex As Integer, TagNo As Integer, Pid As String, temStr As String
Dim temLine As String, neg As Integer, ParaCount As Integer, Ub As Integer, i As Integer, OpType As Integer
Dim q As Boolean, TemType As String
'GetOpCodeInfo vNewValue, neg, OpID
OpID = Val(vNewValue)

AssignValue = CStr(OpID) <> Value
If AssignValue Then
    OpIndex = GetOpIndex(OpID)
    If OpIndex >= 0 Then
       Ub = Operation(OpIndex).ParaNum
       ReDim Default(Ub + 1)
       temLine = Operation(OpIndex).Pseudo
       
       For i = 1 To Operation(OpIndex).ParaNum
          Default(i) = "<" & Operation(OpIndex).Para(i).Value & ">"
       Next i
       
       ParaCount = Ub
       If Operation(OpIndex).Type = OPT_Can_Fail Then
         ParaCount = ParaCount + 1
         Default(ParaCount) = PublicMsgs(158)
         OpType = 1
       End If
       
       If Operation(OpIndex).OpID = Call_Script Then
         ParaCount = ParaCount + 1
         Default(ParaCount) = PublicMsgs(159)
         OpType = 2
       End If
    Else
       Ub = 0
       ReDim Default(Ub)
       temLine = ActiveString(PublicMsgs(155), CStr(OpID))
       ParaCount = 0
    End If

SubItems.Count = ParaCount

For i = 1 To ParaCount
  SubItems(i).KeyWord = Default(i)
  SubItems(i).Value = ""
  SubItems(i).Text = ""
  SubItems(i).ForeColor = MENU_COLOR_DEFAULT
  If i <= Ub Then
    SubItems(i).Locked = True
    If OpIndex > -1 Then
      q = IsOptionalParam(Operation(OpIndex).Para(i).Para_Type, TemType)
      If q Then
         SubItems(i).ForeColor = MENU_COLOR_OPTIONAL
      End If
      SubItems(i).ParaType = TemType
    Else
      SubItems(i).ParaType = ""
    End If
    SubItems(i).TemType = ""
  Else
    SubItems(i).Locked = True
    SubItems(i).Negation = False
    SubItems(i).ParaType = ""
    SubItems(i).TemType = ""
    If OpType = 1 Then
      SubItems(i).Negation = True
      SubItems(i).Value = CStr(neg)
      SubItems(i).Text = GetNegWord(neg)
      SubItems(i).ParaType = "NEG"
      SubItems(i).ForeColor = MENU_COLOR_NEG
    ElseIf OpType = 2 Then
      SubItems(i).ParaType = "MORE"
      SubItems(i).ForeColor = MENU_COLOR_BUTTON
    End If
  End If
Next i

Value = CStr(OpID)
p_Text = temLine
ListView.Event_Change
End If

End Function

