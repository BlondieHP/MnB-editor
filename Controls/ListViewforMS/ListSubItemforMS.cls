VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListSubItemforMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private p_Text As String
Public Value As String    'Reserved
Public Start As Integer
Public Index As Integer
Public KeyWord As String
Public Negation As Boolean
Public Locked As Boolean
Public ParaType As String
Public TemType As String

Public Label As Label
Public ListView As ListViewforMS

Public Initialized As Boolean
Public Parent As ListSubItemsforMS

Public Property Get ForeColor() As Long
ForeColor = Label.ForeColor
End Property

Public Property Let ForeColor(ByVal vNewValue As Long)
Label.ForeColor = vNewValue

End Property

Public Property Get Text() As String
Text = p_Text
End Property

Public Property Let Text(ByVal vNewValue As String)
p_Text = vNewValue

End Property

Public Sub Initialize(lIndex As Integer, sIndex As Integer, UserList As ListViewforMS, LabelBox As Label, lParent As ListSubItemsforMS)

Set ListView = UserList
Set Label = LabelBox
Set Parent = lParent

Label.Visible = True
Index = sIndex

Label.Tag = lIndex & "," & Index

Initialized = True
End Sub


Private Sub Class_Terminate()
Set ListView = Nothing
Set Label = Nothing

Initialized = False
End Sub

Public Property Get FontBold() As Boolean
FontBold = Label.FontBold
End Property

Public Property Let FontBold(ByVal vNewValue As Boolean)
Label.FontBold = vNewValue

End Property

Public Sub AssignValue(vNewValue As String, Optional PType As String = "")
Dim Default As String, OpID As Long, OpIndex As Integer, TagNo As Integer, Pid As String, temStr As String
Dim strType As String, q As Boolean

If vNewValue = "" Then
  Value = ""
  p_Text = ""
  If Label.ForeColor <> MENU_COLOR_OPTIONAL Then Label.ForeColor = MENU_COLOR_DEFAULT
  ListView.Event_Change
  Exit Sub
End If

OpID = Val(Parent.Parent.Value)
OpIndex = GetOpIndex(OpID)

If OpIndex >= 0 Then
  If Index <= Operation(OpIndex).ParaNum Then
    Default = "<" & Operation(OpIndex).Para(Index).Value & ">"
  Else
    Default = ActiveString(PublicMsgs(156), Index)
  End If
Else
  Default = ActiveString(PublicMsgs(156), Index)
End If
  'KeyWord = Default
  Value = vNewValue
  GetParamCodeInfo Value, TagNo, Pid
  If TagNo = 0 Then      '无Tag
     If OpIndex >= 0 Then    '已注册Op
        If Index <= Operation(OpIndex).ParaNum Then    '可能有指定数据类型
           If PType = "" Then    '有指定数据类型
             strType = Operation(OpIndex).Para(Index).Para_Type
           Else      '没有指定数据类型
             strType = PType
           End If
           temStr = ShowNoTagParam(strType, Pid)
        Else      '没有指定数据类型
           temStr = ShowNoTagParam(PType, Pid)
        End If
     Else     '未注册Op
        temStr = ShowNoTagParam(PType, Pid)
     End If
  Else   '有Tag
    If Index <= Operation(OpIndex).ParaNum Then    '可能有指定数据类型
       q = IsOptionalParam(Operation(OpIndex).Para(Index).Para_Type, strType)
    Else     '没有指定数据类型
       strType = ""
    End If
    temStr = ShowParam_Type(TagNo, Val(Pid), strType)
  End If
TemType = PType
p_Text = temStr
If Label.ForeColor <> MENU_COLOR_OPTIONAL Then Label.ForeColor = MENU_COLOR_PARAM
ListView.Event_Change

End Sub
