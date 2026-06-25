Attribute VB_Name = "ModWizard"
Option Explicit

Public Type Type_Wizard_Node
     ID As Integer
     NextID() As Integer
     Frame_Idx As Integer
     AllowDefault As Boolean
End Type

Public Type Type_Wizard_Progress
     ForeNode As String
     NodeNow As Integer
     DefaultChoice As Integer
End Type

Public Type Type_ErrReport
     Number As Long
     Description As String
End Type

Public Pro As Type_Wizard_Progress
Public ItemNodes(5) As Type_Wizard_Node
Public TemItems() As Type_Item
Public TemItemCount As Long


Public Sub InitItemNodes()

TemItemCount = 0

With ItemNodes(0)       '选择模型
     .ID = 0
     
     ReDim .NextID(0)
     .NextID(0) = 1
     
     .Frame_Idx = 0
     .AllowDefault = False
End With

With ItemNodes(1)        '添加txt行
     .ID = 1
     
     ReDim .NextID(1)
     .NextID(0) = 2     '如果有txt行
     .NextID(1) = 3     '没有
     
     .Frame_Idx = 1
     .AllowDefault = False
End With

With ItemNodes(2)        '添加语言信息
     .ID = 2
     
     ReDim .NextID(0)
     .NextID(0) = 3
     
     .Frame_Idx = 2
     .AllowDefault = False
End With


With ItemNodes(3)        '确定导入
     .ID = 3
     
     ReDim .NextID(1)
     .NextID(0) = 4
     .NextID(1) = 5
     
     .Frame_Idx = 3
     .AllowDefault = True
End With

With ItemNodes(4)        '导入成功,有txt行
     .ID = 4
     
     ReDim .NextID(0)
     .NextID(0) = -1
     
     .Frame_Idx = 4
     .AllowDefault = False
End With

With ItemNodes(5)        '导入成功,无txt行
     .ID = 5
     
     ReDim .NextID(0)
     .NextID(0) = -1
     
     .Frame_Idx = 4
     .AllowDefault = False
End With

Pro.NodeNow = 0
Pro.ForeNode = ""
End Sub


Public Function GoFore() As Boolean
Dim TemS() As String, i As Integer

If Trim(Pro.ForeNode) <> "" Then

TemS = Split(Pro.ForeNode, "|")

Pro.NodeNow = Val(TemS(UBound(TemS)))

Pro.ForeNode = "0"
For i = 1 To UBound(TemS) - 1
      Pro.ForeNode = "|" & Pro.ForeNode
Next i

GoFore = True

If Pro.ForeNode = "0" Then
   GoFore = False
End If

End If

End Function

Public Function IsEmptyStr(ByVal Str As String) As Boolean
Dim i As Integer, K As Long

IsEmptyStr = True
For i = 1 To Len(Str)
     K = Asc(Mid(Str, i, 1))
     
     If K <> 13 And K <> 10 And K <> 32 Then
        IsEmptyStr = False
        Exit For
     End If
Next i

End Function

Public Sub SetItemsDefaultCSV()
Dim i As Long

For i = 0 To TemItemCount - 1
    TemItems(i).csvName = TemItems(i).disname
    TemItems(i).csvName_pl = TemItems(i).disname
Next i

End Sub
