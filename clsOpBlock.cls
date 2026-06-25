VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOpBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents OpbEd As OpBlockEditor
Attribute OpbEd.VB_VarHelpID = -1
Private IsAttached As Boolean
Private OpBlocks() As Type_Op_Block
Private LastChange As Integer
Private ReEncode(1 To 3) As Boolean
Private Lct As String

Public Code_TXT As String
Public Code_PY As String
Public ErrPack As New clsEncodeErr

Public Sub Attach(iOpBlockEditor As OpBlockEditor)
Set OpbEd = iOpBlockEditor
OpbEd.Initialize Me
IsAttached = True
End Sub

Public Sub UnLoad(iOpBlockEditor As OpBlockEditor)
Set OpbEd = Nothing
IsAttached = False
End Sub

Friend Sub AssignOpBlock(iBlocks() As Type_Op_Block, Optional AutoEncode As Boolean = True, Optional strLocation As String = "", Optional intCheckListTrgIdx As Long = 0, Optional NeedLocate As Boolean = True)
Dim TabIndex As Long

If NeedLocate Then
  If strLocation <> "" Then Location = strLocation
  If intCheckListTrgIdx > -1 Then CheckListTrgIdx = intCheckListTrgIdx
End If

OpBlocks = iBlocks
OpbEd.Pseudo.ListIndex = 0
OpbEd.Pseudo.SubIndex = 0

ReEncode(1) = True
ReEncode(2) = True
ReEncode(3) = True

If AutoEncode And IsAttached Then
  TabIndex = OpbEd.TabStripIndex
  
  'ReEncode(TabIndex) = False
  
  Refresh
End If
End Sub

Friend Sub GetOpBlock(oBlocks() As Type_Op_Block, Optional SaveCheck As Boolean = False)
Dim q As Boolean
If LastChange = OpbEd.TabStripIndex Then
  Select Case LastChange
     Case 1
       q = ReadOpBlockFromPseudo
       If Not q Then
         'OpbEd.SelStart_PY = ErrPack.Start
         'OpbEd.SelLength_PY = ErrPack.Length
         ErrPack.ShowEncodeErrMsg
         Exit Sub
       Else
         ReEncode(1) = False
         ReEncode(2) = True
         ReEncode(3) = True
       End If
     Case 2
       Code_PY = OpbEd.RichPY.Text
       q = ReadOpBlockFromPY
       If Not q Then
         OpbEd.SelStart_PY = ErrPack.Start
         OpbEd.SelLength_PY = ErrPack.Length
         ErrPack.ShowEncodeErrMsg
         Exit Sub
       Else
         ReEncode(1) = True
         ReEncode(2) = False
         ReEncode(3) = True
       End If
     Case 3
       Code_TXT = OpbEd.RichTXT.Text
       q = ReadOpBlockFromTXT
       If Not q Then
         OpbEd.SelStart_TXT = ErrPack.Start
         ErrPack.ShowEncodeErrMsg
         Exit Sub
       Else
         ReEncode(1) = True
         ReEncode(2) = True
         ReEncode(3) = False
       End If
  End Select
End If
oBlocks = OpBlocks

Dim i As Long, j As Long
If SaveCheck Then
  i = GetVarNameCheckListNo(Lct)

  If i > 0 Then
    VarNameLists(i) = CurVarNameList
  Else
    j = CreateVarNameCheckList(Lct)
    VarNameLists(j) = CurVarNameList
    LocateVarNameCheckList Lct, j
  End If

  VarNameLists(0) = TemGVarNameList
End If
End Sub

Public Function ReadOpBlockFromTXT() As Boolean
On Error GoTo EL

Dim values() As String, j As Long, strTem As String
Dim TemBlocks() As Type_Op_Block, b As Boolean, i As Integer

ReadOpBlockFromTXT = False
If Trim(Code_TXT) = "" Then
  ReDim OpBlocks(0)
  ReadOpBlockFromTXT = True
  ErrPack.Reason = ERR_NULL_STRING
  Exit Function
End If

txtLine = Code_TXT
LinePointer = 1
b = False

ReDim TemBlocks(1 To 1)
i = 1
Do While Not b
   j = 0
   strTem = GetWordL(b)
   If b Then
     Exit Do
   End If
   ReDim Preserve TemBlocks(1 To i)
   With TemBlocks(UBound(TemBlocks))
     
     .Op = strTem
     
     .ParaNum = Val(GetWordL(b))
     If b Then GoTo Unfinished
     
     ReDim .Para(.ParaNum)
     For j = 1 To .ParaNum
       .Para(j).Value = GetWordL(b)
       If b And (j < .ParaNum) Then GoTo Unfinished
     Next j
     
   End With
   
   i = i + 1
Loop
OpBlocks = TemBlocks
ReadOpBlockFromTXT = True

Exit Function

EL:
ErrPack.Source = "ReadOpBlockFromTXT"
ErrPack.Reason = Err.Number
ErrPack.Block = UBound(TemBlocks)
ErrPack.Start = LinePointer
ErrPack.Param = j

Exit Function

Unfinished:
ErrPack.Source = "ReadOpBlockFromTXT"
ErrPack.Reason = ERR_INCOMPLETE_SOURCE_CODE
ErrPack.Block = UBound(TemBlocks)
ErrPack.Start = LinePointer
ErrPack.Param = j
End Function

Public Function ReadOpBlockFromPY() As Boolean
On Error GoTo EL

Dim values() As String, i As Long, j As Long, iPlus As Long, n As Long, Op_Index As Integer
Dim Remark As String, Remark_Start As Long, Params() As String, Params_Start() As Long, head As Integer
Dim TemBlocks() As Type_Op_Block

ReDim TemBlocks(1 To 1)
ReadOpBlockFromPY = False

If Trim(Code_PY) = "" Then
  ReDim OpBlocks(0)
  ReadOpBlockFromPY = True
  ErrPack.Reason = ERR_NULL_STRING
  Exit Function
End If
  
values = Split(Code_PY, vbCrLf)

Do While i <= UBound(values)
   n = PursePYLine(values(i), 1, Params(), Params_Start(), Remark, Remark_Start)
   head = head + Len(values(i))
   j = 0
   If n > -1 Then
     ReDim Preserve TemBlocks(1 To i + 1)
     With TemBlocks(UBound(TemBlocks))
       .Op = EncodeOperation(Params(0))
       
       If .Op = "" Then
         ErrPack.Reason = ERR_BAD_OPERATION
         Err.Clear
         GoTo EL
       End If
       
       .ParaNum = UBound(Params) - 1
       ReDim .Para(.ParaNum)
       For j = 1 To UBound(Params) - 1
         .Para(j).Value = EncodeParam(Params(j))
       Next j
     End With
   End If
   
   i = i + 1
Loop

OpBlocks = TemBlocks
ReadOpBlockFromPY = True

Exit Function

EL:
If Err.Number <> 0 Then ErrPack.Reason = Err.Number
ErrPack.Source = "ReadOpBlockFromPY"
ErrPack.Block = UBound(TemBlocks)
ErrPack.Start = head + Params_Start(j)
ErrPack.Length = Len(Params(j))
ErrPack.Line = i
ErrPack.Param = j
End Function

Public Function OutputOpBlockAsTXT(Optional AutoEnt As Boolean = False) As Boolean
On Error GoTo EL

Dim i As Long, j As Long

Code_TXT = ""
For i = 1 To UBound(OpBlocks)
   With OpBlocks(i)
     Code_TXT = Code_TXT & .Op & " " & .ParaNum & " "
     
     For j = 1 To .ParaNum
       Code_TXT = Code_TXT & .Para(j).Value & " "
     Next j

   End With
   
   If AutoEnt And i < UBound(OpBlocks) Then Code_TXT = Code_TXT & vbCrLf
Next i
OutputOpBlockAsTXT = True

If IsAttached And OutputOpBlockAsTXT Then
  'OpbEd.TextTXT = Code_TXT
  OpbEd.RichTXT.DrawText Code_TXT
End If

Exit Function

EL:
ErrPack.Source = "OutputOpBlockAsTXT"
ErrPack.Block = i
ErrPack.Param = j
ErrPack.Reason = Err.Number

End Function

Public Function OutputOpBlockAsPY() As Boolean
Dim EB As Integer, EP As Integer, ER As Integer

Code_PY = ""
OutputOpBlockAsPY = ExportPYCodefromTXT_OpBlocks(OpBlocks, Code_PY, EB, EP, ER, " ")

ErrPack.Source = "OutputOpBlockAsPY"
ErrPack.Block = EB
ErrPack.Param = EP
ErrPack.Reason = ER

If IsAttached And OutputOpBlockAsPY Then
  OpbEd.TextPY = Code_PY
End If
End Function

Private Sub OpBEd_Change(TabIndex As Integer)
LastChange = TabIndex
End Sub


Public Sub OpBEd_TabClick(Previous As Integer, TabIndex As Integer)

Dim q As Boolean
q = True
If LastChange = Previous Then
  Select Case Previous
     Case 1
       q = ReadOpBlockFromPseudo
       If Not q Then
         'OpbEd.SelStart_PY = ErrPack.Start
         'OpbEd.SelLength_PY = ErrPack.Length
         ErrPack.ShowEncodeErrMsg
         Exit Sub
       Else
         ReEncode(1) = False
         ReEncode(2) = True
         ReEncode(3) = True
       End If
     Case 2
       Code_PY = OpbEd.RichPY.Text
       q = ReadOpBlockFromPY
       If Not q Then
         OpbEd.SelStart_PY = ErrPack.Start
         OpbEd.SelLength_PY = ErrPack.Length
         ErrPack.ShowEncodeErrMsg
         Exit Sub
       Else
         ReEncode(1) = True
         ReEncode(2) = False
         ReEncode(3) = True
       End If
     Case 3
       Code_TXT = OpbEd.RichTXT.Text
       q = ReadOpBlockFromTXT
       If Not q Then
         OpbEd.SelStart_TXT = ErrPack.Start
         ErrPack.ShowEncodeErrMsg
         Exit Sub
       Else
         ReEncode(1) = True
         ReEncode(2) = True
         ReEncode(3) = False
       End If
  End Select
End If
LastChange = 0

  Select Case TabIndex
     Case 1
       If ReEncode(TabIndex) Then
         q = OutputOpBlockAsPseudo
         ReEncode(TabIndex) = False
       End If
     Case 2
       If ReEncode(TabIndex) Then
         OpbEd.NoChangeEvent = True
         q = OutputOpBlockAsPY
         OpbEd.NoChangeEvent = False
         ReEncode(TabIndex) = False
       End If
     Case 3
       If ReEncode(TabIndex) Then
         OpbEd.NoChangeEvent = True
         q = OutputOpBlockAsTXT(True)
         OpbEd.NoChangeEvent = False
         ReEncode(TabIndex) = False
       End If
  End Select
  
  If Not q Then
    ErrPack.ShowEncodeErrMsg
    Exit Sub
  End If


End Sub

Public Sub Refresh()
Dim q As Boolean
q = True
  Select Case OpbEd.TabStripIndex
     Case 1
       If ReEncode(OpbEd.TabStripIndex) Then
         q = OutputOpBlockAsPseudo
         ReEncode(OpbEd.TabStripIndex) = False
       End If
     Case 2
       If ReEncode(OpbEd.TabStripIndex) Then
         q = OutputOpBlockAsPY
         ReEncode(OpbEd.TabStripIndex) = False
       End If
     Case 3
       If ReEncode(OpbEd.TabStripIndex) Then
         q = OutputOpBlockAsTXT
         ReEncode(OpbEd.TabStripIndex) = False
       End If
  End Select
  
If Not q Then
  ErrPack.ShowEncodeErrMsg
  Exit Sub
End If
End Sub

Public Function OutputOpBlockAsPseudo() As Boolean
On Error GoTo EL
Dim EB As Integer, EP As Integer, ER As Integer, oItem As ListItemforMS, ParaAry() As String
Dim i As Long, j As Long, Indention As Integer, neg As Integer, OpID As Long, OpIndex As Integer

If IsAttached Then
OpbEd.Pseudo.ListItems.Clear
Indention = 0

  For i = 1 To UBound(OpBlocks)
     With OpBlocks(i)
       'GetOpCodeInfo .Op, neg, OpID
     
       ReDim ParaAry(0 To .ParaNum)
       ParaAry(0) = .Op
       For j = 1 To .ParaNum
         ParaAry(j) = OpBlocks(i).Para(j).Value
       Next j

       Set oItem = OpbEd.Pseudo.ListItems.AddOpBlock(0, ParaAry())

     End With

  Next i

  CalcIndention
  OpbEd.Pseudo.Refresh


End If

OutputOpBlockAsPseudo = True

Exit Function

EL:
ErrPack.Source = "OutputOpBlockAsPseudo"
ErrPack.Block = i
ErrPack.Param = j
ErrPack.Reason = Err.Number

End Function

Public Sub CalcIndention()
Dim i As Integer, neg As Integer, OpID As Long, Indention As Long

Indention = 0
For i = 1 To OpbEd.Pseudo.ListItems.Count
   With OpbEd.Pseudo.ListItems(i)
     GetOpCodeInfo .Value, neg, OpID
     If OpID = else_try Then
       .Indention = Indention - 4
     ElseIf OpID = try_begin Or OpID = try_for_range Or OpID = try_for_range_backwards Or OpID = try_for_parties Or OpID = try_for_agents Then
       .Indention = Indention
       Indention = Indention + 4
     ElseIf OpID = try_end Then
       Indention = Indention - 4
       .Indention = Indention
     Else
       .Indention = Indention
     End If
   End With
Next i
End Sub

Public Function ReadOpBlockFromPseudo() As Boolean
On Error GoTo EL

Dim Op As String, i As Long, j As Long, iPlus As Long
Dim TemBlocks() As Type_Op_Block, ParaCount As Integer
Dim q As Boolean

ReadOpBlockFromPseudo = False

If OpbEd.Pseudo.ListItems.Count = 0 Then
  ReDim OpBlocks(0)
  ErrPack.Reason = ERR_NULL_STRING
  ReadOpBlockFromPseudo = True
  Exit Function
End If

ReDim TemBlocks(1 To 1)
i = 1
Do While i <= OpbEd.Pseudo.ListItems.Count
   ReDim Preserve TemBlocks(1 To i)
   j = 0
   With TemBlocks(UBound(TemBlocks))
     
     ParaCount = OpbEd.Pseudo.ListItems(i).SubItems.Count
     If ParaCount > 0 And OpbEd.Pseudo.ListItems(i).SubItems(ParaCount).ParaType = "MORE" Then
       .Op = OpbEd.Pseudo.ListItems(i).Value
       .ParaNum = ParaCount - 1
     ElseIf ParaCount > 0 And OpbEd.Pseudo.ListItems(i).SubItems(ParaCount).ParaType = "NEG" Then
       .Op = EncodeOperationNeg(Val(OpbEd.Pseudo.ListItems(i).Value), Val(OpbEd.Pseudo.ListItems(i).SubItems(ParaCount).Value))
       .ParaNum = ParaCount - 1
     Else
       .Op = OpbEd.Pseudo.ListItems(i).Value
       .ParaNum = ParaCount
     End If

     ReDim .Para(.ParaNum)
     For j = 1 To .ParaNum
       If OpbEd.Pseudo.ListItems(i).SubItems(j).Value <> "" Then
         .Para(j).Value = OpbEd.Pseudo.ListItems(i).SubItems(j).Value
       Else
         If OpbEd.Pseudo.ListItems(i).SubItems(j).ForeColor = MENU_COLOR_OPTIONAL Then
           ReDim Preserve .Para(j - 1)
           .ParaNum = j - 1
           Exit For
         Else
           '.Para(j).Value = "0"  'Unfinished
           GoTo Unfinished
         End If
       End If
     Next j
     
     i = i + 1
   End With
Loop
OpBlocks = TemBlocks
ReadOpBlockFromPseudo = True

Exit Function

EL:
ErrPack.Source = "ReadOpBlockFromPseudo"
ErrPack.Reason = Err.Number
ErrPack.Block = UBound(TemBlocks)
ErrPack.Start = i
ErrPack.Param = j

Exit Function

Unfinished:
ErrPack.Source = "ReadOpBlockFromPseudo"
ErrPack.Reason = ERR_INCOMPLETE_SOURCE_CODE
ErrPack.Block = UBound(TemBlocks)
ErrPack.Start = i
ErrPack.Param = j
End Function

Public Property Get Visible() As Boolean
Visible = OpbEd.Visible
End Property

Public Property Let Visible(ByVal vNewValue As Boolean)
OpbEd.Visible = vNewValue
End Property


Public Property Get Location() As String
Location = Lct
End Property

Public Property Let Location(ByVal vNewValue As String)  '先赋值Location,再给Triggers赋值
Dim i As Long
Lct = vNewValue
i = GetVarNameCheckListNo(Lct)

If i > 0 Then
  CurVarNameList = VarNameLists(i)
Else
  UpdateCurVarNameCheckList Lct, 0, True
End If

TemGVarNameList = VarNameLists(0)

End Property
