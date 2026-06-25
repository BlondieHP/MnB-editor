VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListItemsforMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private ListView As ListViewforMS
Private ListItem() As New ListItemforMS
Private AutoRef As Boolean
Private DefColor As Long
Private DefFont As StdFont
Private DefLocked As Boolean
Public ErrPack As New clsEncodeErr

Public Sub Initialize(UserList As ListViewforMS)

Set ListView = UserList

End Sub

Public Function AddItem(Text As String, Optional ListIndex As Integer, Optional ParamCount As Integer = 0) As ListItemforMS
Dim New_Index As Integer, i As Integer, j As Integer, MainLabel As Label, ParamLabel As Label

If ListIndex > 0 And ListIndex <= UBound(ListItem) Then
   New_Index = ListIndex
Else
   New_Index = UBound(ListItem) + 1
End If

ReDim Preserve ListItem(UBound(ListItem) + 1)

ListItem(UBound(ListItem)).Index = UBound(ListItem)
For i = UBound(ListItem) - 1 To New_Index Step -1

  CopyListItem ListItem(i + 1), ListItem(i)
  
  If i = New_Index Then
    Set ListItem(i) = Nothing
  End If
Next i

With ListItem(New_Index)
  .Initialize New_Index, ListView, ListView.FreeLabel()
  .SubItems.Count = ParamCount
  .Locked = DefLocked
End With

With ListItem(New_Index).Label
  ListItem(New_Index).Text = Text
  .Caption = Text
  .Visible = True
  .Tag = CStr(New_Index) & ",0"
  .Left = 0
  .Top = (New_Index - 1) * .Height
  .ForeColor = DefColor
End With

For i = New_Index + 1 To UBound(ListItem)
  Set MainLabel = ListItem(i).Label
  MainLabel.Left = 0
  MainLabel.Top = (i - 1) * ListView.Spacing * MainLabel.Height
  'MainLabel.Tag = TagPlus(MainLabel.Tag, 0, 1)
  For j = 1 To ListItem(i).SubItems.Count
    Set ParamLabel = ListItem(i).SubItems(j).Label
    ParamLabel.Top = MainLabel.Top
  Next j
Next i

Set AddItem = ListItem(New_Index)

If AutoRef Then
  ListItem(New_Index).Refresh True
End If

End Function


Public Sub Clear()
Dim i As Integer

For i = 1 To ListView.Label.Count - 1
  UnLoad ListView.Label(i)
Next i

ReDim ListItem(0)

End Sub


Public Property Get Item(ByVal Index As Integer) As ListItemforMS
Attribute Item.VB_UserMemId = 0
Set Item = ListItem(Index)
End Property

Public Property Let Item(ByVal Index As Integer, ByVal vNewValue As ListItemforMS)
Dim i As Integer, l As Label
Set ListItem(Index) = vNewValue

Set l = ListItem(Index).Label
l.Tag = TagAssign(l.Tag, 0, ListItem(Index).Index)

For i = 1 To UBound(ListItem)
  Set l = ListItem(Index).SubItems(i).Label
  l.Tag = TagAssign(l.Tag, 0, ListItem(Index).Index)
Next i

End Property

Public Property Get Count() As Integer
Count = UBound(ListItem)
End Property

Private Sub Class_Initialize()
ReDim ListItem(0)
End Sub

Public Sub Remove(Index As Integer)
Dim i As Integer, j As Integer, Label(1) As Label

ListItem(Index).Label.Visible = False
For j = 1 To ListItem(Index).SubItems.Count
   ListItem(Index).SubItems(j).Label.Visible = False
Next j
  
For i = UBound(ListItem) - 1 To Index Step -1
    Set Label(0) = ListItem(i).Label
    Set Label(1) = ListItem(i + 1).Label
    Label(1).Top = Label(0).Top
    'Label(1).Tag = TagPlus(Label(1).Tag, 0, -1)
  
    For j = 1 To ListItem(i).SubItems.Count
      ListItem(i).SubItems(j).Label.Top = ListItem(i).Label.Top
    Next j
Next i

For i = Index To UBound(ListItem) - 1
  CopyListItem ListItem(i), ListItem(i + 1)
Next i

ReDim Preserve ListItem(UBound(ListItem) - 1)

If AutoRef Then
  ListView.MeasureFrame
End If

End Sub

Public Sub Swap(Item1 As Integer, Item2 As Integer)
Dim tItem As New ListItemforMS

CopyListItem tItem, ListItem(Item1), True

CopyListItem ListItem(Item1), ListItem(Item2), False

CopyListItem ListItem(Item2), tItem, False


End Sub

Public Property Get Locked() As Boolean
Locked = DefLocked

End Property

Public Property Let Locked(ByVal vNewValue As Boolean)
Dim i As Integer

DefLocked = vNewValue
For i = 1 To UBound(ListItem)
  ListItem(i).Locked = DefLocked
Next i

End Property

Public Property Get ForeColor() As Long
ForeColor = DefColor

End Property

Public Property Let ForeColor(ByVal vNewValue As Long)
Dim i As Integer

DefColor = vNewValue
For i = 1 To UBound(ListItem)
  ListItem(i).ForeColor = DefColor
Next i

End Property

Public Property Get FontBold() As Boolean
FontBold = DefFont.Bold
End Property

Public Property Let FontBold(ByVal vNewValue As Boolean)
Dim i As Integer
DefFont.Bold = vNewValue

For i = 1 To UBound(ListItem)
  ListItem(i).FontBold = vNewValue
Next i
End Property

Private Sub Class_Terminate()
Dim i As Integer

For i = 1 To UBound(ListItem)
  Set ListItem(i) = Nothing
Next i
End Sub

'*************************************************************************
'**函 数 名：AddOpBlock
'**输    入：(String)Operation,(Array)Param[0为操作值]
'**输    出：(String)
'**功能描述：增加操作块列表
'**全局变量：
'**调用模块：
'**作    者：
'**日    期：
'**修 改 人：SSgt_Edward
'**日    期：2011-07-15 10:39:24
'**版    本：V1.1321
'*************************************************************************
Public Function AddOpBlock(ByVal ListIndex As Integer, Param() As String) As ListItemforMS
On Error GoTo EL

Dim i As Integer, n As Integer
Dim neg As Integer, OpID As Long, OpIndex As Long, TagNo As Integer, Pid As String
Dim temStr As String, temLine As String, temParam As String, ParaCount As Integer
Dim oItem As ListItemforMS
Dim notFail As Boolean, Op As String
Dim Default() As String, OpType As Long, Ub As Integer, TemType As String, q As Boolean
Dim l As String, r As String, j As Integer

Op = Param(0)

    GetOpCodeInfo Op, neg, OpID
    OpIndex = GetOpIndex(OpID)
    If OpIndex >= 0 Then
       Ub = Max_Int(Operation(OpIndex).ParaNum, UBound(Param))
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

         For i = Operation(OpIndex).ParaNum + 1 To UBound(Param)
           If OpType <> 2 Then
             Default(i) = ActiveString(PublicMsgs(156), i)
             temLine = temLine & "["
             temLine = temLine & Default(i)
             temLine = temLine & "]"
           Else
             Default(i) = ActiveString(PublicMsgs(156), i - Operation(OpIndex).ParaNum)
             j = InStrRev(temLine, Default(ParaCount))
             If j > 0 Then
               l = Left(temLine, Len(temLine) - Len(Default(ParaCount)))
               r = Right(temLine, Len(Default(ParaCount)))
               temLine = l & Default(i) & "," & r
             End If
           End If
         Next i
    Else
       Ub = UBound(Param)
       ReDim Default(UBound(Param))
       
       temLine = ActiveString(PublicMsgs(155), CStr(OpID))
       For i = 1 To UBound(Param)
         Default(i) = ActiveString(PublicMsgs(156), i)
         temLine = temLine & Default(i) & IIf(i < UBound(Param), ",", "")
       Next i
       
       ParaCount = UBound(Param)
    End If
    
    Set oItem = AddItem(temLine, ListIndex, ParaCount)
    oItem.Value = OpID
    oItem.Locked = True 'False
    oItem.ForeColor = IIf(IsControlOp(OpID), MENU_COLOR_CONTROL, MENU_COLOR_MAIN)
    
      For n = 1 To ParaCount
        oItem.SubItems(n).KeyWord = Default(n)
        oItem.SubItems(n).Locked = False
        oItem.SubItems(n).ForeColor = MENU_COLOR_DEFAULT
        
        If n <= Ub Then
          If n <= UBound(Param) Then '1.有传参
            oItem.SubItems(n).ForeColor = MENU_COLOR_PARAM
            oItem.SubItems(n).Value = Param(n)
            GetParamCodeInfo Param(n), TagNo, Pid
            If TagNo = 0 Then '1.1无Tag
               If OpIndex >= 0 Then '1.1.1操作已注册
                 If n <= Operation(OpIndex).ParaNum Then '1.1.1.1参数序号在注册范围内
                    temStr = ShowNoTagParam(Operation(OpIndex).Para(n).Para_Type, Pid)
                    q = IsOptionalParam(Operation(OpIndex).Para(n).Para_Type, TemType)
                    If q Then
                       oItem.SubItems(n).ForeColor = MENU_COLOR_OPTIONAL
                    End If
                    oItem.SubItems(n).ParaType = TemType
                 Else '1.1.1.2参数序号不在注册范围内
                    temStr = ShowNoTagParam("", Pid)
                    oItem.SubItems(n).ParaType = "" ' "NA"
                 End If
               Else '1.1.2操作未注册
                 temStr = ShowNoTagParam("", Pid)
                 oItem.SubItems(n).ParaType = "" ' "NA"
               End If
            Else '1.2.有Tag
               'temStr = ShowParam(TagNo, Val(Pid), PublicTags(TagNo) & ":([index])[csvname]")
               If OpIndex >= 0 Then
                 If n <= Operation(OpIndex).ParaNum Then '1.2.1参数序号在注册范围内
                   q = IsOptionalParam(Operation(OpIndex).Para(n).Para_Type, TemType)
                   If q Then
                      oItem.SubItems(n).ForeColor = MENU_COLOR_OPTIONAL
                   End If
                    oItem.SubItems(n).ParaType = TemType
                 Else '1.2.2参数序号不在注册范围内
                   oItem.SubItems(n).ParaType = "" ' "NA"
                 End If
               Else
                 oItem.SubItems(n).ParaType = "" ' "NA"
               End If
               temStr = ShowParam_Type(TagNo, Val(Pid), oItem.SubItems(n).ParaType)
            End If
          Else '2.无有传参
            If OpIndex >= 0 Then '2.1操作已注册
              If n <= Operation(OpIndex).ParaNum Then '2.1.1参数序号在注册范围内
                q = IsOptionalParam(Operation(OpIndex).Para(n).Para_Type, TemType)
                If q Then
                   oItem.SubItems(n).ForeColor = MENU_COLOR_OPTIONAL
                End If
                oItem.SubItems(n).ParaType = TemType
              Else '2.1.2参数序号不在注册范围内
                oItem.SubItems(n).ParaType = "" ' "NA"
              End If
            Else '2.2操作未注册
              oItem.SubItems(n).ParaType = "" ' "NA"
            End If
            temStr = ""
          End If
          
          oItem.SubItems(n).Locked = True 'IsMenuType(oItem.SubItems(n).ParaType)
          oItem.SubItems(n).Text = temStr
          oItem.SubItems(n).Negation = False
        Else
          If OpType = 1 Then
            oItem.SubItems(n).Locked = True
            oItem.SubItems(n).Negation = True
            oItem.SubItems(n).Value = CStr(neg)
            oItem.SubItems(n).Text = GetNegWord(neg)
            oItem.SubItems(n).ForeColor = MENU_COLOR_NEG
            oItem.SubItems(n).ParaType = "NEG"
          ElseIf OpType = 2 Then
            oItem.SubItems(n).Locked = True
            oItem.SubItems(n).ForeColor = MENU_COLOR_BUTTON
            oItem.SubItems(n).ParaType = "MORE"
          End If
        End If
      Next n
Set AddOpBlock = oItem
Exit Function
EL:
    Call logErr("ListItemsforMS", "AddOpBlock", Err.Number, Err.Description)
End Function

