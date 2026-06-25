Attribute VB_Name = "ModINI"
Public Type Type_ModINI
    ModResourceCount As Long
    ModResource() As String
End Type

Public ModSets As Type_ModINI

'*************************************************************************
'**函 数 名：LoadModINI
'**输    入： -
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-01-27 16:53:33
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub LoadModINI()
On Error GoTo errorHandle

Dim F As Integer, temStr As String, TemU() As String

ReDim ModSets.ModResource(0)
F = FreeFile

Open MnBInfo.ModIniFileName For Input As #F
     Do While Not EOF(F)
        Line Input #F, temStr
        TemU = Split(temStr, " = ")
        
      If UBound(TemU) > 0 Then
        If LCase(TemU(0)) = "load_mod_resource" Or LCase(TemU(0)) = "load_module_resource" Then
           
            ModSets.ModResourceCount = ModSets.ModResourceCount + 1
            ReDim Preserve ModSets.ModResource(ModSets.ModResourceCount)
            ModSets.ModResource(ModSets.ModResourceCount) = TemU(1)
        End If
      End If
      
     Loop
Close #F

Exit Sub
'----------------
errorHandle:
    Call logErr("ModINI", "LoadModINI", Err.Number, Err.Description)
End Sub
