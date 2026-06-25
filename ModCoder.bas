Attribute VB_Name = "ModCoder"
Option Explicit

Public Const ERR_SUCCESS = 0
Public Const ERR_FAIL = 1
Public Const ERR_BAD_QUETO = 2

Public Const ERR_BAD_OPERATION = -1
Public Const ERR_NULL_STRING = 0
Public Const ERR_INCOMPLETE_SOURCE_CODE = 9

Public Const MNU_EMPTY = 0
Public Const MNU_ACTIVE = 1
Public Const MNU_CONST = 2
Public Type Type_Text_Param
  Content As String
  Start As Long
  Length As Long
End Type

Public Type Type_Variable_Name_Check_Sub_List
  Checks() As String
End Type

Public Type Type_Variable_Name_Check_List
  Triggers() As Type_Variable_Name_Check_Sub_List
  Location As String
End Type

Public PYTags(1 To 26) As String
Public PYQueto(1 To 26) As Boolean
Public PYUnderLine(1 To 26) As Boolean
Public PYMenu(1 To 26) As Long
Public TagReg(1 To 26) As Boolean
Public VarNameLists() As Type_Variable_Name_Check_List
Public CurVarNameList As Type_Variable_Name_Check_List
Public TemGVarNameList As Type_Variable_Name_Check_List
Public CheckListTrgIdx As Long

Public Function SplitParam(ByVal CMD As String, Params() As String) As Long
Dim i As Long, strTem() As String, tCMD As String, strParam() As String, j As Long, tParam() As String, k As String, q As Boolean

strTem = Split(CMD, Chr(34))

If UBound(strTem) > 0 Then
  tCMD = ""
  ReDim strParam((UBound(strTem) + 1) \ 2 - 1)
  For i = 1 To UBound(strTem) Step 2
    j = (i + 1) \ 2 - 1
    strParam(j) = strTem(i)
    strTem(i) = "{str" & j & "}"
  Next i
  
  For i = 0 To UBound(strTem)
    tCMD = tCMD & strTem(i)
  Next i
  
  q = True
Else
  tCMD = CMD
  q = False
End If

tParam = Split(tCMD, ",")
If UBound(tParam) >= 0 Then
  ReDim Params(UBound(tParam))
  j = 0
  For i = 0 To UBound(tParam)
    Params(i) = tParam(i)
    
    If q Then
      For j = 0 To UBound(strParam)
        k = "{str" & j & "}"
        If InStr(1, Params(i), k) > 0 Then
          Params(i) = Replace(Params(i), k, Chr(34) & strParam(j) & Chr(34))
        End If
      Next j
    End If
  Next i
  SplitParam = UBound(tParam) + 1
Else
  ReDim Params(0)
  Params(0) = CMD
  SplitParam = 0
End If
End Function

Public Function IsEven(ByVal Value As Long) As Boolean
IsEven = Value / 2 = Value \ 2
End Function

Public Function PurseParams(strPara As String, Pointer As Long, Params() As String, Param_Start() As Long) As Long
Dim i As Long, len_now As Long, n As Long

n = SplitParam(strPara, Params())

If n = 0 Then
  ReDim Param_Start(0)
  PurseParams = -1
  Exit Function
End If

ReDim Param_Start(UBound(Params))
For i = 0 To UBound(Params)

  Param_Start(i) = len_now + 1
  len_now = len_now + Len(Params(i)) + 1
  
  If len_now > Pointer Then
    PurseParams = i
    Exit For
  End If
  
Next i

End Function

Public Function StandardizeParams(Params() As String, Param_Start() As Long) As Boolean
Dim i As Long, q(1) As Boolean, Ub As Long, Last As Long, TL As Long

Ub = UBound(Params)

For i = 0 To Ub
  TL = Len(Params(i)) - Len(LTrim(Params(i)))
  Param_Start(i) = Param_Start(i) + TL
  Params(i) = Trim(Params(i))
Next i

If Len(Params(0)) > 0 Then
  If Left(Params(0), 1) = "(" Then
    Params(0) = Right(Params(0), Len(Params(0)) - 1)
    Param_Start(0) = Param_Start(0) + 1
    
    TL = Len(Params(0)) - Len(LTrim(Params(0)))
    Param_Start(0) = Param_Start(0) + TL
    Params(0) = Trim(Params(0))
    q(0) = True
  End If
End If

If Ub > 0 Then
  If Trim(Params(Ub)) = "" Then
    Last = Ub - 1
  Else
    Last = -1
  End If
Else
  Last = Ub
End If

If Last > -1 Then
  If Len(Params(Last)) > 1 Then
    If Right(Params(Last), 1) = ")" Then
      Params(Last) = Left(Params(Last), Len(Params(Last)) - 1)
      Params(Last) = Trim(Params(Last))
      q(1) = True
    End If
  End If
End If

StandardizeParams = q(0) And q(1)
End Function

Public Function SplitCodeLine(ByVal strCodes As String, Pointer As Long, Optional head As Long, Optional Tail As Long) As String
Dim n As Long, m As Long

If strCodes = "" Then Exit Function

If Pointer = 0 Then Pointer = 1
n = InStr(Pointer, strCodes, Chr(13))
m = InStrRev(strCodes, Chr(10), Pointer)

If n = 0 Then
  n = Len(strCodes)
Else
  n = n - 1
End If

m = m + 1

SplitCodeLine = Mid(strCodes, m, n - m + 1)
Pointer = Pointer - m + 1

If Not IsMissing(head) Then
head = m
End If

If Not IsMissing(Tail) Then
Tail = n - m + 1
End If

End Function

Public Function PursePYLine(ByVal PYLine As String, Pointer As Long, Params() As String, Params_Start() As Long, Remark As String, Remark_Start As Long) As Long
Dim strTem() As String, i As Long, j As Long, n As Long, tLine As String, CMDLine As String

PursePYLine = -1
Remark_Start = 0
Remark = ""
If Trim(PYLine) = "" Then Exit Function

strTem = Split(PYLine, "#")

If UBound(strTem) = 0 Then
  CMDLine = strTem(0)
Else
  For i = 0 To UBound(strTem)
   
    If i = 0 Then
      tLine = strTem(0)
    Else
      tLine = tLine & "#" & strTem(i)
    End If
   
    n = GetCharCount(tLine, Chr(34))
    If IsEven(n) Then
      Remark_Start = Len(tLine) + 1
      CMDLine = tLine
     
      For j = i + 1 To UBound(strTem)
        Remark = Remark & strTem(j)
      Next j
      Exit For
    End If

  Next i
End If

n = PurseParams(CMDLine, Pointer, Params(), Params_Start())

StandardizeParams Params(), Params_Start()

PursePYLine = n

End Function

Public Function GetCharCount(ByVal Str As String, ByVal Char As String, Optional Ignore_Case As Boolean = False) As Long
Dim CompareMod As Long

If Ignore_Case Then
  CompareMod = vbTextCompare
Else
  CompareMod = vbBinaryCompare
End If

GetCharCount = Len(Str) - Len(Replace(Str, Char, "", , , CompareMod))
End Function

Public Function TrimQueto(ByVal Param As String) As String
Dim strTem As String
If Param = "" Then Exit Function

If Left(Param, 1) = Chr(34) Then
  strTem = Right(Param, Len(Param) - 1)
End If

If strTem = "" Then Exit Function

If Right(strTem, 1) = Chr(34) Then
  TrimQueto = Left(strTem, Len(strTem) - 1)
End If

TrimQueto = strTem

End Function

'*************************************************************************
'**函 数 名：GetCaretColLine
'**输    入：(long)TextHwnd,(long)LineNo,(long)ColNo
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Zephyrous' tickler
'**日    期：2006-08-08 16:19
'**修 改 人：SSgt_Edward
'**日    期：2011-05-27 23:00:23
'**版    本：V1.1321
'*************************************************************************
Public Sub GetCaretColLine(ByVal TextHwnd As Long, LineNo As Long, ColNo As Long)

Dim i As Long, j As Long
Dim lParam As Long, wParam As Long
Dim k As Long

'首先向文本框传递EM_GETSEL消息以获取从起始位置到
'光标所在位置的字符数

i = SendMessage(TextHwnd, EM_GETSEL, wParam, lParam)
j = i / 2 ^ 16

'再向文本框传递EM_LINEFROMCHAR消息根据获得的字符
'数确定光标以获取所在行数

LineNo = SendMessage(TextHwnd, EM_LINEFROMCHAR, j, 0)
LineNo = LineNo + 1

'向文本框传递EM_LINEINDEX消息以获取所在列数

k = SendMessage(TextHwnd, EM_LINEINDEX, -1, 0)
ColNo = j - k
End Sub

Public Function GetPYTag(ByVal Param As String) As Integer
Dim i As Integer, p As String, n As Long

If Param = "" Then Exit Function

  For i = 1 To 26
      p = GetPYPrefix(i)
      n = InStr(1, Param, p, vbTextCompare)
      If n = 1 Then
        GetPYTag = i
        Exit For
      End If
  Next i

End Function

'*************************************************************************
'**函 数 名：InitPYTags
'**输    入：无
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-06-02 21:56:11
'**修 改 人：
'**日    期：
'**版    本：V1.132.1
'*************************************************************************
Sub InitPYTags()
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    PYTags(Tag_Register) = "reg"  '自定义tag
    PYQueto(Tag_Register) = False
    PYUnderLine(Tag_Register) = False
    PYMenu(Tag_Register) = MNU_EMPTY
    TagReg(Tag_Register) = False    'A-E
    
    
    PYTags(Tag_Variable) = "$"  '自定义tag
    PYQueto(Tag_Variable) = True
    PYUnderLine(Tag_Variable) = False
    PYMenu(Tag_Variable) = MNU_ACTIVE
    TagReg(Tag_Variable) = True
    
    PYTags(Tag_String) = "str"
    PYQueto(Tag_String) = True
    PYUnderLine(Tag_String) = True
    PYMenu(Tag_String) = MNU_CONST
    TagReg(Tag_String) = False
    
    PYTags(Tag_Item) = "itm"
    PYQueto(Tag_Item) = True
    PYUnderLine(Tag_Item) = True
    PYMenu(Tag_Item) = MNU_CONST
    TagReg(Tag_Item) = True
    
    PYTags(Tag_Troop) = "trp"
    PYQueto(Tag_Troop) = True
    PYUnderLine(Tag_Troop) = True
    PYMenu(Tag_Troop) = MNU_CONST
    TagReg(Tag_Troop) = True
    
    PYTags(Tag_Faction) = "fac"
    PYQueto(Tag_Faction) = True
    PYUnderLine(Tag_Faction) = True
    PYMenu(Tag_Faction) = MNU_CONST
    TagReg(Tag_Faction) = True
    
    PYTags(Tag_Quest) = "qst"
    PYQueto(Tag_Quest) = True
    PYUnderLine(Tag_Quest) = True
    PYMenu(Tag_Quest) = MNU_CONST
    TagReg(Tag_Quest) = False
    
    PYTags(Tag_Party_Tpl) = "pt"
    PYQueto(Tag_Party_Tpl) = True
    PYUnderLine(Tag_Party_Tpl) = True
    PYMenu(Tag_Party_Tpl) = MNU_CONST
    TagReg(Tag_Party_Tpl) = True
    
    PYTags(Tag_Party) = "p"
    PYQueto(Tag_Party) = True
    PYUnderLine(Tag_Party) = True
    PYMenu(Tag_Party) = MNU_CONST
    TagReg(Tag_Party) = True
    
    PYTags(Tag_Scene) = "scn"
    PYQueto(Tag_Scene) = True
    PYUnderLine(Tag_Scene) = True
    PYMenu(Tag_Scene) = MNU_CONST
    TagReg(Tag_Scene) = True
    
    PYTags(Tag_Mission_tpl) = "mst"
    PYQueto(Tag_Mission_tpl) = True
    PYUnderLine(Tag_Mission_tpl) = True
    PYMenu(Tag_Mission_tpl) = MNU_CONST
    TagReg(Tag_Mission_tpl) = False
    
    PYTags(Tag_Menu) = "mnu"
    PYQueto(Tag_Menu) = True
    PYUnderLine(Tag_Menu) = True
    PYMenu(Tag_Menu) = MNU_CONST
    TagReg(Tag_Menu) = False
    
    PYTags(Tag_Script) = "script"
    PYQueto(Tag_Script) = True
    PYUnderLine(Tag_Script) = True
    PYMenu(Tag_Script) = MNU_CONST
    TagReg(Tag_Script) = False
    
    PYTags(Tag_Particle_Sys) = "psys"
    PYQueto(Tag_Particle_Sys) = True
    PYUnderLine(Tag_Particle_Sys) = True
    PYMenu(Tag_Particle_Sys) = MNU_CONST
    TagReg(Tag_Particle_Sys) = True
    
    PYTags(Tag_Scene_Prop) = "spr"
    PYQueto(Tag_Scene_Prop) = True
    PYUnderLine(Tag_Scene_Prop) = True
    PYMenu(Tag_Scene_Prop) = MNU_CONST
    TagReg(Tag_Scene_Prop) = False
    
    PYTags(Tag_Sound) = "snd"
    PYQueto(Tag_Sound) = True
    PYUnderLine(Tag_Sound) = True
    PYMenu(Tag_Sound) = MNU_CONST
    TagReg(Tag_Sound) = True
    
    PYTags(Tag_Local_Variable) = ":" '自定义tag
    PYQueto(Tag_Local_Variable) = True
    PYUnderLine(Tag_Local_Variable) = False
    PYMenu(Tag_Local_Variable) = MNU_ACTIVE
    TagReg(Tag_Local_Variable) = True
    
    PYTags(Tag_Map_Icon) = "icon"    '自定义tag
    PYQueto(Tag_Map_Icon) = True
    PYUnderLine(Tag_Map_Icon) = True
    PYMenu(Tag_Map_Icon) = MNU_CONST
    TagReg(Tag_Map_Icon) = True
    
    PYTags(Tag_Skill) = "skl"
    PYQueto(Tag_Skill) = True
    PYUnderLine(Tag_Skill) = True
    PYMenu(Tag_Skill) = MNU_CONST
    TagReg(Tag_Skill) = True
    
    PYTags(Tag_Mesh) = "mesh"
    PYQueto(Tag_Mesh) = True
    PYUnderLine(Tag_Mesh) = True
    PYMenu(Tag_Mesh) = MNU_CONST
    TagReg(Tag_Mesh) = True
    
    PYTags(Tag_Presentation) = "prsnt"
    PYQueto(Tag_Presentation) = True
    PYUnderLine(Tag_Presentation) = True
    PYMenu(Tag_Presentation) = MNU_CONST
    TagReg(Tag_Presentation) = False
    
    PYTags(Tag_Quick_String) = "qstr"
    PYQueto(Tag_Quick_String) = True
    PYUnderLine(Tag_Quick_String) = True
    PYMenu(Tag_Quick_String) = MNU_ACTIVE
    TagReg(Tag_Quick_String) = False
    
    PYTags(Tag_Track) = "track"     '自定义tag
    PYQueto(Tag_Track) = True
    PYUnderLine(Tag_Track) = True
    PYMenu(Tag_Track) = MNU_CONST
    TagReg(Tag_Track) = False
    
    PYTags(Tag_Tableau) = "tab"
    PYQueto(Tag_Tableau) = True
    PYUnderLine(Tag_Tableau) = True
    PYMenu(Tag_Tableau) = MNU_CONST
    TagReg(Tag_Tableau) = True
    
    PYTags(Tag_Animation) = "anim"   '自定义tag
    PYQueto(Tag_Animation) = True
    PYUnderLine(Tag_Animation) = True
    PYMenu(Tag_Animation) = MNU_CONST
    TagReg(Tag_Animation) = False
    
    PYTags(Tags_End) = "end"
    PYQueto(Tags_End) = False
    PYUnderLine(Tags_End) = False
    TagReg(Tags_End) = True

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "InitTags", Err.Number, Err.Description)
End Sub

Public Function GetPYPrefix(ByVal Index As Long) As String
GetPYPrefix = IIf(PYQueto(Index), Chr(34), "") & PYTags(Index) & IIf(PYUnderLine(Index), "_", "")

End Function


Public Function EncodeParam(ByVal Param As String) As String
Dim i As Long, TagNo As Integer, tP As String, n As Long, strTem As String, l As Long

If Trim(Param) = "" Then Exit Function
tP = Trim(Param)
TagNo = GetPYTag(tP)

If TagNo > 0 Then
   If PYQueto(TagNo) Then
     If Left(tP, 1) = Chr(34) Then
       tP = Right(tP, Len(tP) - 1)
       If Right(tP, 1) = Chr(34) Then
         tP = Left(tP, Len(tP) - 1)
         
         If TagNo = Tag_Local_Variable Or TagNo = Tag_Variable Then
           n = GetVariableID(tP)
         Else
           n = GetID(tP, False, , -1)
         End If
         
         If n = -1 Then
           l = Len(PYTags(TagNo) & IIf(PYUnderLine(TagNo), "_", ""))
           strTem = Right(tP, Len(tP) - l)
           
           If IsNumeric(strTem) Then
             n = Val(strTem)
           End If
         End If
         
         If n >= 0 Then
           EncodeParam = getTXTID(TagNo, n)
         Else
           Exit Function
         End If
       Else
         Exit Function
       End If
     Else
       Exit Function
     End If
   Else
      tP = Replace(tP, PYTags(TagNo), "", , , vbTextCompare)
      EncodeParam = getTXTID(TagNo, Val(tP))
   End If
Else
  Dim q As Boolean
  EncodeParam = getNoTagParamValue(tP, q)
  
  If q Then
      Call logErr("ModCoder", "EncodeParam", "0", "Param[" + tP + "] do not exist!")
  End If
End If
End Function


Public Function EncodeOperation(ByVal strOp As String) As String
Dim i As Long, strTem() As String, tP As String, n As Long, Op_Index As Integer, OpID As Long, k As String
Dim Neg64b As Integer64b, I64 As Integer64b

If Trim(strOp) = "" Then Exit Function
tP = Trim(strOp)

strTem = Split(tP, "|")

Op_Index = GetOpIndexbyName(strTem(UBound(strTem)))

If Op_Index >= 0 Then
  OpID = Operation(Op_Index).OpID
Else
  If IsNumeric(strTem(UBound(strTem))) Then
    OpID = Val(strTem(UBound(strTem)))
  Else
    Exit Function
  End If
End If

I64 = StrToI64(CStr(OpID))

For i = 0 To UBound(strTem) - 1
  k = LCase(Trim(strTem(i)))
  
  If k = "neg" Then
    Neg64b = HexStrToI64(neg)
    I64 = Or64b(I64, Neg64b)
  ElseIf k = "this_or_next" Then
    Neg64b = HexStrToI64(this_or_next)
    I64 = Or64b(I64, Neg64b)
  End If
Next i

EncodeOperation = I64toStrNZ(I64)
End Function


Public Function EncodeOperationNeg(ByVal OpID As Long, ByVal neg As Integer) As String
Dim Neg64b As Integer64b, I64 As Integer64b, strNeg As String

I64 = StrToI64(CStr(OpID))

If neg = 0 Then
  strNeg = "0"
ElseIf neg = 1 Then
  strNeg = "80000000"
ElseIf neg = 2 Then
  strNeg = "40000000"
ElseIf neg = 3 Then
  strNeg = "C0000000"
End If

Neg64b = HexStrToI64(strNeg)
I64 = Or64b(I64, Neg64b)

EncodeOperationNeg = I64toStrNZ(I64)
End Function


'*************************************************************************
'**函 数 名：ReadVarNameCheckLists
'**输    入：无
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-06-16 22:33:11
'**修 改 人：
'**日    期：
'**版    本：V1.132.1
'*************************************************************************
Public Sub ReadVarNameCheckLists()
Dim FileName As String, Count As Long, Section As String, i As Long, j As Long, m As Long, Ub As Long, Ub2 As Long

FileName = MnBInfo.iniFileName
Count = ReadInt(FileName, "EDITINFO", "Variable_Name_Check_Lists_Count")
ReDim VarNameLists(Count)

'------Global Variables-----
ReDim VarNameLists(0).Triggers(1)
ReDim VarNameLists(0).Triggers(1).Checks(UBound(gVars))
For i = 0 To UBound(gVars)
   VarNameLists(0).Triggers(1).Checks(i) = gVars(i).VarName
Next i
'---------------------------

For i = 1 To Count  '0 To Count
   With VarNameLists(i)
     Section = "Variable_Name_Check_List_" & i
     .Location = ReadString(FileName, Section, "Location")
     
     'If .Location = "" Then GoTo NextFor
     LocateVarNameCheckList .Location, i
     
     Ub = ReadInt(FileName, Section, "Ubound")
     
     ReDim .Triggers(Ub)
     For j = 1 To Ub
       Ub2 = ReadInt(FileName, Section, CStr(j) & "_Ubound")
       
       ReDim .Triggers(j).Checks(Ub2)   'check redim
       For m = 0 To Ub2
         .Triggers(j).Checks(m) = ReadString(FileName, Section, j & "_" & m)
       Next m
     Next j
   End With
   
NextFor:
Next i

End Sub

'*************************************************************************
'**函 数 名：SetVarNameCheck
'**输    入：无
'**输    出：无
'**功能描述：设置变量名对照
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-06-16 22:33:11
'**修 改 人：
'**日    期：
'**版    本：V1.132.1
'*************************************************************************
Public Sub SetVarNameCheck(CheckList As Long, Index As Long, Param_No As Integer, ParaName As String)
Dim i As Long, Ub As Long

With VarNameLists(CheckList)
  Ub = UBound(.Triggers)
  If Index > Ub Then
    ReDim Preserve .Triggers(Index)
    For i = Ub + 1 To Index
      ReDim .Triggers(i).Checks(0)    'check redim
    Next i
  End If
  
  If Param_No > UBound(.Triggers(Index).Checks) Then
    ReDim Preserve .Triggers(UBound(.Triggers)).Checks(Param_No)  'check redim
  End If
  .Triggers(Index).Checks(Param_No) = ParaName
End With

End Sub

'*************************************************************************
'**函 数 名：GetVarNameCheckListNo
'**输    入：(String)Location,(Long)Check_No
'**输    出：(Boolean)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-06-20 14:06:19
'**修 改 人：
'**日    期：
'**版    本：V1.132.1
'*************************************************************************
Public Function GetVarNameCheckListNo(Location As String) As Long

GetVarNameCheckListNo = GetSlot(Location, SLOT_VARIABLE_NAME_CHECK_LIST, False)

End Function

'*************************************************************************
'**函 数 名：LocateVarNameCheckList
'**输    入：(String)Location,(Long)Check_No
'**输    出：(Boolean)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-06-19 23:15:38
'**修 改 人：
'**日    期：
'**版    本：V1.132.1
'*************************************************************************
Public Function LocateVarNameCheckList(Location As String, Check_No As Long) As Boolean

If Location = "" Then Exit Function
LocateVarNameCheckList = SetSlot(Location, SLOT_VARIABLE_NAME_CHECK_LIST, CStr(Check_No), False)

End Function

'*************************************************************************
'**函 数 名：CreateVarNameCheckList
'**输    入：(String)Location
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-06-19 23:15:38
'**修 改 人：
'**日    期：
'**版    本：V1.132.1
'*************************************************************************
Public Function CreateVarNameCheckList(Location As String, Optional Trigger As Long = 0) As Long
Dim i As Long, j As Long, q As Boolean, Ub As Long

For i = 1 To UBound(VarNameLists)
  With VarNameLists(i)
    If .Location = "" Then
      q = False
      Exit For
    ElseIf LCase(.Location) = LCase(Location) Then
      q = True
      Exit For
    End If
  End With
Next i

If i > UBound(VarNameLists) Or Not q Then
  ReDim Preserve VarNameLists(i)
  ReDim VarNameLists(i).Triggers(Trigger)
  
  For j = 0 To Trigger
    ReDim VarNameLists(i).Triggers(j).Checks(0)   'check redim
  Next j
Else
  If q Then
    Ub = UBound(VarNameLists(i).Triggers)
    If Ub < Trigger Then
      ReDim Preserve VarNameLists(i).Triggers(Trigger)
      For j = Ub + 1 To Trigger
        ReDim VarNameLists(i).Triggers(j).Checks(0)   'check redim
      Next j
    End If
  End If
End If

VarNameLists(i).Location = Location

CreateVarNameCheckList = i
End Function

'*************************************************************************
'**函 数 名：UpdateCurVarNameCheckList
'**输    入：(String)Location
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-09-14 21:14:15
'**修 改 人：
'**日    期：
'**版    本：V1.132.1
'*************************************************************************
Public Sub UpdateCurVarNameCheckList(Location As String, Optional Trigger As Long = 0, Optional Reset As Boolean = False)
On Error GoTo EL
Dim j As Long, Ub As Long

If Reset Then
  ReDim CurVarNameList.Triggers(Trigger)
  
  For j = 0 To Trigger
    ReDim CurVarNameList.Triggers(j).Checks(0)   'check redim
  Next j
Else
    Ub = UBound(CurVarNameList.Triggers)
    ReDim Preserve CurVarNameList.Triggers(Trigger)
  
    For j = Ub + 1 To Trigger
      ReDim Preserve CurVarNameList.Triggers(j).Checks(0)   'check redim
    Next j
End If

CurVarNameList.Location = Location
Exit Sub

EL:
If Err.Number = 9 Then
  ReDim CurVarNameList.Triggers(Trigger)
    
    For j = 0 To Trigger
      ReDim CurVarNameList.Triggers(j).Checks(0)   'check redim
    Next j
    
    CurVarNameList.Location = Location
End If
End Sub

'*************************************************************************
'**函 数 名：SaveVarNameCheckList
'**输    入：(Long)Check_No
'**输    出：(Boolean)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-06-20 13:56:47
'**修 改 人：
'**日    期：
'**版    本：V1.132.1
'*************************************************************************
Public Function SaveVarNameCheckList(Check_No As Long) As Boolean
On Error GoTo EL
Dim q As Boolean, res As Boolean, Count As Long, i As Long, j As Long
res = True

With VarNameLists(Check_No)

  Count = ReadInt(MnBInfo.iniFileName, "EDITINFO", "Variable_Name_Check_Lists_Count")
  If Count < Check_No Then
     q = WriteString(MnBInfo.iniFileName, "EDITINFO", "Variable_Name_Check_Lists_Count", CStr(Check_No))
     res = res And q
  End If
  q = WriteString(MnBInfo.iniFileName, "Variable_Name_Check_List_" & Check_No, "Location", .Location)
  res = res And q
  
  q = WriteString(MnBInfo.iniFileName, "Variable_Name_Check_List_" & Check_No, "Ubound", UBound(.Triggers))
  res = res And q
  
  For i = 1 To UBound(.Triggers)
    q = WriteString(MnBInfo.iniFileName, "Variable_Name_Check_List_" & Check_No, i & "_Ubound", UBound(.Triggers(i).Checks))
    res = res And q
    
    For j = 0 To UBound(.Triggers(i).Checks)
      q = WriteString(MnBInfo.iniFileName, "Variable_Name_Check_List_" & Check_No, i & "_" & j, .Triggers(i).Checks(j))
      res = res And q
    Next j
  Next i
End With

SaveVarNameCheckList = res

Exit Function

EL:
SaveVarNameCheckList = False
End Function


'*************************************************************************
'**函 数 名：DeleteVarNameCheckList
'**输    入：(Long)Check_No
'**输    出：(Boolean)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-06-20 14:00:14
'**修 改 人：
'**日    期：
'**版    本：V1.132.1
'*************************************************************************
Public Function DeleteVarNameCheckList(Check_No As Long) As Boolean
On Error GoTo EL
Dim q As Boolean, Count As Long, i As Long, j As Long

With VarNameLists(Check_No)
  .Location = ""
  ReDim .Triggers(0)
  ReDim .Triggers(0).Checks(0)   'check redim
End With

DeleteVarNameCheckList = True

Exit Function

EL:
DeleteVarNameCheckList = False
End Function


'*************************************************************************
'**函 数 名：GetVariableID
'**输    入：(String)VarName
'**输    出：(Long)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-06-20 14:00:14
'**修 改 人：
'**日    期：
'**版    本：V1.132.1
'*************************************************************************
Public Function GetVariableID(VarName As String, Optional DefaultIndex As Long = -1, Optional DefaultSet As Boolean = True) As Long
Dim tN As String, Chk_Ves As Type_Variable_Name_Check_List, Index As Long, i As Long, tIndex As String, q As Boolean
Dim strPre As String, Ub As Long, trgUb As Long

tN = Trim(VarName)
If tN = "" Then Exit Function

If Left(tN, 1) = PYTags(Tag_Variable) Then
  Chk_Ves = TemGVarNameList
  Index = 1
  tN = Right(tN, Len(tN) - Len(PYTags(Tag_Variable)))
  q = True
  strPre = PublicTags(Tag_Variable)
ElseIf Left(tN, 1) = PYTags(Tag_Local_Variable) Then
  Chk_Ves = CurVarNameList
  Index = CheckListTrgIdx
  tN = Right(tN, Len(tN) - Len(PYTags(Tag_Local_Variable)))
  q = False
  strPre = PublicTags(Tag_Local_Variable)
Else
  Exit Function
End If

If tN = "" Then Exit Function
  If DefaultSet Then
    If Len(tN) > 4 Then
      If LCase(Left(tN, 4)) = "var_" Then
        tIndex = Right(tN, Len(tN) - 4)
        If IsNumeric(tIndex) Then
          GetVariableID = Val(tIndex)
          Exit Function
        End If
      End If
    End If
    
    If Len(tN) > Len(strPre) Then
      If LCase(Left(tN, Len(strPre))) = strPre Then
        tIndex = Right(tN, Len(tN) - Len(strPre))
        If IsNumeric(tIndex) Then
          GetVariableID = Val(tIndex)
          Exit Function
        End If
      End If
    End If
  End If
If UBound(Chk_Ves.Triggers) >= Index Then
With Chk_Ves.Triggers(Index)
  
  For i = 0 To UBound(.Checks)
    If LCase(.Checks(i)) = LCase(tN) Then
      GetVariableID = i
      Exit Function
    End If
  Next i
  
  If DefaultIndex <= -1 Then
    For i = 0 To UBound(.Checks)
      If .Checks(i) = "" Then
        GetVariableID = i
        .Checks(i) = tN
        GoTo Finally
      End If
    Next i
  Else
    GetVariableID = DefaultIndex
    If UBound(.Checks) < DefaultIndex Then
      ReDim Preserve .Checks(DefaultIndex)
    End If
    .Checks(DefaultIndex) = tN
    GetVariableID = DefaultIndex
    
    GoTo Finally
  End If
  
  ReDim Preserve .Checks(UBound(.Checks) + 1)
  .Checks(UBound(.Checks)) = tN
  GetVariableID = i
  
End With
Else
  trgUb = UBound(Chk_Ves.Triggers)
  ReDim Preserve Chk_Ves.Triggers(Index)
  For i = trgUb + 1 To Index - 1
    ReDim Chk_Ves.Triggers(i).Checks(0)
  Next i
  
  Ub = IIf(DefaultIndex > -1, DefaultIndex, 0)
  ReDim Chk_Ves.Triggers(Index).Checks(Ub)   'check redim
  Chk_Ves.Triggers(Index).Checks(Ub) = tN
  GetVariableID = Ub
End If

Finally:
If q Then
  TemGVarNameList = Chk_Ves
Else
  CurVarNameList = Chk_Ves
End If



End Function

'*************************************************************************
'**函 数 名：GetVariablePYCode
'**输    入：(Long)Pid
'**输    出：(Long)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-06-20 20:32:52
'**修 改 人：
'**日    期：
'**版    本：V1.132.1
'*************************************************************************
Public Function GetVariablePYCode(Pid As Long, Optional IsGlobal As Boolean = False) As String
Dim tN As String, Chk_Ves As Type_Variable_Name_Check_List, Index As Long, tIndex As String

If IsGlobal Then
  Chk_Ves = TemGVarNameList
  Index = 1
  
Else
  Chk_Ves = CurVarNameList
  Index = CheckListTrgIdx
End If

If UBound(Chk_Ves.Triggers) >= Index Then
With Chk_Ves.Triggers(Index)
   If Pid <= UBound(.Checks) Then
     If .Checks(Pid) <> "" Then
       GetVariablePYCode = .Checks(Pid)
     Else
       GetVariablePYCode = "var_" & Pid
     End If
   Else
     GetVariablePYCode = "var_" & Pid
   End If
End With
Else
  GetVariablePYCode = "var_" & Pid
End If
End Function

'*************************************************************************
'**函 数 名：IsVariableIDExist
'**输    入：(String)VarName
'**输    出：(Long)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-06-20 14:00:14
'**修 改 人：
'**日    期：
'**版    本：V1.132.1
'*************************************************************************
Public Function IsVariableIDExist(VarName As String, Optional DefaultSet As Boolean = True) As Long
Dim tN As String, Chk_Ves As Type_Variable_Name_Check_List, Index As Long, i As Long, tIndex As String, q As Boolean
Dim strPre As String

IsVariableIDExist = -1
tN = Trim(VarName)
If tN = "" Then Exit Function

If Left(tN, 1) = PYTags(Tag_Variable) Then
  Chk_Ves = TemGVarNameList
  Index = 1
  tN = Right(tN, Len(tN) - Len(PYTags(Tag_Variable)))
  q = True
  strPre = PublicTags(Tag_Variable)
ElseIf Left(tN, 1) = PYTags(Tag_Local_Variable) Then
  Chk_Ves = CurVarNameList
  Index = CheckListTrgIdx
  tN = Right(tN, Len(tN) - Len(PYTags(Tag_Local_Variable)))
  q = False
  strPre = PublicTags(Tag_Local_Variable)
Else
  Exit Function
End If

If tN = "" Then Exit Function
  If DefaultSet Then
    If Len(tN) > 4 Then
      If LCase(Left(tN, 4)) = "var_" Then
        tIndex = Right(tN, Len(tN) - 4)
        If IsNumeric(tIndex) Then
          IsVariableIDExist = Val(tIndex)
          Exit Function
        End If
      End If
    End If
    
    If Len(tN) > Len(strPre) Then
      If LCase(Left(tN, Len(strPre))) = strPre Then
        tIndex = Right(tN, Len(tN) - Len(strPre))
        If IsNumeric(tIndex) Then
          IsVariableIDExist = Val(tIndex)
          Exit Function
        End If
      End If
    End If
  End If
  
If UBound(Chk_Ves.Triggers) >= Index Then
  With Chk_Ves.Triggers(Index)
    For i = 0 To UBound(.Checks)
      If LCase(.Checks(i)) = LCase(tN) Then
        IsVariableIDExist = i
        Exit Function
      End If
    Next i
  End With
End If

End Function

'*************************************************************************
'**函 数 名：getNoTagParamValue
'**输    入：(String)Param
'**输    出：(String)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2012-03-10 21:15:26
'**修 改 人：
'**日    期：
'**版    本：V1.132.1
'*************************************************************************
Public Function getNoTagParamValue(Param As String, Optional isFailed As Boolean = False) As String
Dim i As Integer

isFailed = True
If Left(Param, 3) = "pos" Then
   getNoTagParamValue = Replace(Param, "pos", "")
   isFailed = False
ElseIf Left(Param, 8) = "itp_type" Then
   For i = 1 To UBound(Item_Type)
      If Item_Type(i).X = Param Then
         getNoTagParamValue = CStr(i)
         isFailed = False
         Exit For
      End If
   Next i
ElseIf Left(Param, 3) = "tf_" Then
   For i = 0 To UBound(Tf)
      If Tf(i).strName = Param Then
         getNoTagParamValue = I64toStrNZ(Tf(i).Value)
         isFailed = False
         Exit For
      End If
   Next i
ElseIf Left(Param, 3) = "pf_" Then
   For i = 0 To UBound(Pf)
      If Pf(i).strName = Param Then
         getNoTagParamValue = I64toStrNZ(Pf(i).Value)
         isFailed = False
         Exit For
      End If
   Next i
ElseIf Left(Param, 7) = "ai_bhvr" Then
   For i = 0 To UBound(AI_Bhvr)
      If AI_Bhvr(i).X = Param Then
         getNoTagParamValue = CStr(i)
         isFailed = False
         Exit For
      End If
   Next i
ElseIf Left(Param, 1) = "s" Then
   getNoTagParamValue = Replace(Param, "s", "")
   isFailed = False
Else
   getNoTagParamValue = Param
   isFailed = False
End If

End Function
