VERSION 5.00
Begin VB.Form frmParamType 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "选择参数类型"
   ClientHeight    =   3285
   ClientLeft      =   8235
   ClientTop       =   4350
   ClientWidth     =   3285
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   3285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton CMDCancel 
      Caption         =   "取消(&C)"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定(&O)"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   1455
   End
   Begin VB.ListBox LstParamTypes 
      Height          =   2400
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmParamType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FormTag As String
Public Trg_No As Integer
Public Act_No As Integer
Public Param_No As Integer
Public Op_ID As Long

'*************************************************************************
'**函 数 名：LoadLstParamTypes
'**输    入：无
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-01-26 20:31:22
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub LoadLstParamTypes()
Dim i As Integer

LstParamTypes.Clear   'A_P
LstParamTypes.AddItem PublicMsgs(88)

For i = 1 To UBound(Tags) - 1
      LstParamTypes.AddItem PublicTags(i)
Next i

End Sub

Private Sub CMDCancel_Click()
Select Case FormTag
    Case "edit_2"
       frmItems.Enabled = True
    Case "edit_11"
       frmTabMat.Enabled = True
    Case "edit_13_Condition"
       frmTrigger.Enabled = True
    Case "edit_13_Consequence"
       frmTrigger.Enabled = True
End Select

Unload Me
End Sub

Private Sub CmdOK_Click() 'A_P
Dim Indentation As String, tStr As String

Select Case FormTag
    Case "edit_2"
       frmItems.Enabled = True
       With CurrentItm.Trigger(Trg_No).tiAct(Act_No)
             .ParaNum = .ParaNum + 1
             If .ParaNum <= 1 Then
                ReDim .Para(1 To .ParaNum)
             Else
                ReDim Preserve .Para(1 To .ParaNum)
             End If
       If LstParamTypes.ListIndex = 0 Then
            .Para(Param_No).Value = "0"
            .Para(Param_No).strID = ""
       Else
            .Para(Param_No).Value = getTXTID(LstParamTypes.ListIndex, 0, .Para(Param_No).strID)     '根据参数类型创建参数初始值
       End If
          frmItems.TVTrigger.Nodes.Add "Op(" & Trg_No & "," & Act_No & ",0)", tvwChild, "Op(" & Trg_No & "," & Act_No & "," & Param_No & ")", GetParaEntity(Op_ID, Param_No, .Para(Param_No))
       End With
   Case "edit_11"
       frmTabMat.Enabled = True
       With CurrentTabMat.OpBlock(Act_No)
             .ParaNum = .ParaNum + 1
             If .ParaNum <= 1 Then
                ReDim .Para(1 To .ParaNum)
             Else
                ReDim Preserve .Para(1 To .ParaNum)
             End If
       If LstParamTypes.ListIndex = 0 Then
            .Para(Param_No).Value = "0"
            .Para(Param_No).strID = ""
       Else
            .Para(Param_No).Value = getTXTID(LstParamTypes.ListIndex, 0, .Para(Param_No).strID)    '根据参数类型创建参数初始值
       End If
          If Param_No >= .ParaNum Then
              Prefix = Chr(3) & Chr(6)
          Else
              Prefix = Chr(25) & Chr(6)
          End If
          
          Indentation = GetIndentationStr(GetIndentation(frmTabMat.TVOpBlocks.Nodes("Op(0," & Act_No & ",0)").Text))
          If Len(Indentation) >= 1 Then
              Indentation = Left(Indentation, Len(Indentation) - 1)
          End If
          
          If Param_No > 1 Then
               tStr = frmTabMat.TVOpBlocks.Nodes("Op(0," & Act_No & "," & Param_No - 1 & ")").Text
               ChangeTVParamPrefix tStr, False
               frmTabMat.TVOpBlocks.Nodes("Op(0," & Act_No & "," & Param_No - 1 & ")").Text = tStr
          End If
          frmTabMat.TVOpBlocks.Nodes.Add "Op(0," & Act_No & ",0)", tvwChild, "Op(0," & Act_No & "," & Param_No & ")", Indentation & Prefix & GetParaEntity(Op_ID, Param_No, .Para(Param_No))
         
       End With
   Case "edit_13_Condition"
       frmTrigger.Enabled = True
       With CurrentTimeTrg.Condition(Act_No)
             .ParaNum = .ParaNum + 1
             If .ParaNum <= 1 Then
                ReDim .Para(1 To .ParaNum)
             Else
                ReDim Preserve .Para(1 To .ParaNum)
             End If
       If LstParamTypes.ListIndex = 0 Then
            .Para(Param_No).Value = "0"
            .Para(Param_No).strID = ""
       Else
            .Para(Param_No).Value = getTXTID(LstParamTypes.ListIndex, 0, .Para(Param_No).strID)    '根据参数类型创建参数初始值
       End If
          If Param_No >= .ParaNum Then
              Prefix = Chr(3) & Chr(6)
          Else
              Prefix = Chr(25) & Chr(6)
          End If
          
          Indentation = GetIndentationStr(GetIndentation(frmTrigger.TVConditions.Nodes("Op(0," & Act_No & ",0)").Text))
          If Len(Indentation) >= 1 Then
              Indentation = Left(Indentation, Len(Indentation) - 1)
          End If
          
          If Param_No > 1 Then
               tStr = frmTrigger.TVConditions.Nodes("Op(0," & Act_No & "," & Param_No - 1 & ")").Text
               ChangeTVParamPrefix tStr, False
               frmTrigger.TVConditions.Nodes("Op(0," & Act_No & "," & Param_No - 1 & ")").Text = tStr
          End If
          frmTrigger.TVConditions.Nodes.Add "Op(0," & Act_No & ",0)", tvwChild, "Op(0," & Act_No & "," & Param_No & ")", Indentation & Prefix & GetParaEntity(Op_ID, Param_No, .Para(Param_No))
       End With
   Case "edit_13_Consequence"
       frmTrigger.Enabled = True
       With CurrentTimeTrg.Consequence(Act_No)
             .ParaNum = .ParaNum + 1
             If .ParaNum <= 1 Then
                ReDim .Para(1 To .ParaNum)
             Else
                ReDim Preserve .Para(1 To .ParaNum)
             End If
       If LstParamTypes.ListIndex = 0 Then
            .Para(Param_No).Value = "0"
            .Para(Param_No).strID = ""
       Else
            .Para(Param_No).Value = getTXTID(LstParamTypes.ListIndex, 0, .Para(Param_No).strID)    '根据参数类型创建参数初始值
       End If
          If Param_No >= .ParaNum Then
              Prefix = Chr(3) & Chr(6)
          Else
              Prefix = Chr(25) & Chr(6)
          End If
          
          Indentation = GetIndentationStr(GetIndentation(frmTrigger.TVConsequences.Nodes("Op(0," & Act_No & ",0)").Text))
          If Len(Indentation) >= 1 Then
              Indentation = Left(Indentation, Len(Indentation) - 1)
          End If
          
          If Param_No > 1 Then
               tStr = frmTrigger.TVConsequences.Nodes("Op(0," & Act_No & "," & Param_No - 1 & ")").Text
               ChangeTVParamPrefix tStr, False
               frmTrigger.TVConsequences.Nodes("Op(0," & Act_No & "," & Param_No - 1 & ")").Text = tStr
          End If
          frmTrigger.TVConsequences.Nodes.Add "Op(0," & Act_No & ",0)", tvwChild, "Op(0," & Act_No & "," & Param_No & ")", Indentation & Prefix & GetParaEntity(Op_ID, Param_No, .Para(Param_No))
       End With
End Select

Unload Me

End Sub

Private Sub Form_Deactivate()
Me.ZOrder
End Sub

Private Sub Form_Load()
LoadLstParamTypes
TranslateForm Me

LstParamTypes.ListIndex = 0
End Sub

