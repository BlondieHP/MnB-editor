Attribute VB_Name = "ModMap"
Option Explicit

Public Type Type_Actions
      Scroll As tPoint
      dSCL As Single
End Type

Public Type Type_Custom
      ViewP As tPoint
      SCL As Single
      Action As Type_Actions
End Type

Public Type Type_MapObject
      PartyID As Long
      Label As String
      LabelSize As Long
      Body As tPoint
      lColor As Long
      Degree As Single
      Flags As String
      Visible As Boolean
      MOType As Long
End Type

Public MapObj() As Type_MapObject
Public CntP As tPoint
Public Po As tPoint
Public Custom As Type_Custom
Public N_MOs As Long       'MapObject总数

Public Const MO_Radius_Medium = 10

Public Const MO_LabelSize_Small = 9
Public Const MO_LabelSize_Medium = 12
Public Const MO_LabelSize_Large = 15

Public Const MO_Type_TemTroop = 0
Public Const MO_Type_Town = 1
Public Const MO_Type_Castle = 2
Public Const MO_Type_Village = 3
Public Const MO_Type_Bridge = 4
Public Const MO_Type_RespawnPoint = 5


'*************************************************************************
'**函 数 名：InitModMap
'**输    入：-
'**输    出：-
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：
'**日    期：
'**修 改 人：SSgt_Edward
'**日    期：2010-12-09 15:43:52
'**版    本：V1.1321
'*************************************************************************
Public Sub InitModMap()
    With Custom
        .ViewP.X = 0
        .ViewP.Y = 0
        .SCL = 1
        .Action.Scroll.X = 0
        .Action.Scroll.Y = 0
        .Action.dSCL = 1
    End With
    
    Po.X = 0: Po.Y = 0
    
    InitMapObjects
End Sub

'*************************************************************************
'**函 数 名：InitMapObjects
'**输    入：-
'**输    出：-
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：
'**日    期：
'**修 改 人：SSgt_Edward
'**日    期：2010-12-09 15:43:52
'**版    本：V1.1321
'*************************************************************************
Public Sub InitMapObjects()
Dim i As Long, fI64 As Integer64b, fI64_NOW As Integer64b, resI64 As Integer64b, TemFlags(2) As String, j As Integer

ReDim MapObj(0 To N_Party - 1)
N_MOs = N_Party
For i = 0 To N_Party - 1
  With Parties(i)
    MapObj(i).PartyID = i
    MapObj(i).Label = .csvName
    MapObj(i).Body = SetPoint(.InitPos(1).X, .InitPos(1).Y)
    MapObj(i).lColor = vbBlack ' "&H" & MnBtoRGBColor(Factions(.Faction).lColor)
    MapObj(i).Degree = Val(.Degree)
    MapObj(i).Flags = .Flags
    MapObj(i).Visible = True
  End With

fI64_NOW = StrToI64(MapObj(i).Flags)

TemFlags(0) = pf_label_small
TemFlags(1) = pf_label_medium
TemFlags(2) = pf_label_large

MapObj(i).LabelSize = MO_LabelSize_Small
For j = 1 To 2
       fI64 = HexStrToI64(TemFlags(j))
       resI64 = And64b(fI64_NOW, fI64)
               
       If IsEqual64b(resI64, fI64) Then
         If j = 1 Then
           MapObj(i).LabelSize = MO_LabelSize_Medium
         Else
           MapObj(i).LabelSize = MO_LabelSize_Large
         End If
       End If
Next j

MapObj(i).MOType = GetMapObjectsType(MapObj(i).Flags)

MapObj(i).lColor = GetMapObjectsColor(MapObj(i).MOType)
Next i

End Sub

'*************************************************************************
'**函 数 名：GetMapObjectsType
'**输    入：(String)Flags,(Long)LabelSize
'**输    出：(Long)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：
'**日    期：
'**修 改 人：SSgt_Edward
'**日    期：2010-12-11 11:41:50
'**版    本：V1.1321
'*************************************************************************
Public Function GetMapObjectsType(ByVal Flags As String) As Long
Dim fI64 As Integer64b, fI64_NOW As Integer64b, resI64 As Integer64b, TemFlags() As Variant, j As Integer
fI64_NOW = StrToI64(Flags)

TemFlags = Array(pf_others, pf_town, pf_castle, pf_village, pf_bridge, pf_respawnpoint)

For j = 1 To 5
       fI64 = HexStrToI64(CStr(TemFlags(j)))
       resI64 = And64b(fI64_NOW, fI64)
               
       If IsEqual64b(resI64, fI64) Then
            GetMapObjectsType = j
               Exit For
       End If
       
Next j

End Function

'*************************************************************************
'**函 数 名：GetMapObjectsColor
'**输    入：(Long)mType
'**输    出：(Long)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：
'**日    期：
'**修 改 人：SSgt_Edward
'**日    期：2010-12-11 11:41:50
'**版    本：V1.1321
'*************************************************************************
Public Function GetMapObjectsColor(ByVal mType As Long) As Long
Select Case mType
      Case MO_Type_TemTroop
           GetMapObjectsColor = vbBlack
      Case MO_Type_Town
           GetMapObjectsColor = vbBlue
      Case MO_Type_Castle
           GetMapObjectsColor = &HC000C0
      Case MO_Type_Village
           GetMapObjectsColor = &H8000&
      Case MO_Type_Bridge
           GetMapObjectsColor = &HFF8080
      Case MO_Type_RespawnPoint
           GetMapObjectsColor = vbRed
End Select
End Function

'*************************************************************************
'**函 数 名：ApplyMapObjects
'**输    入：-
'**输    出：-
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：
'**日    期：
'**修 改 人：SSgt_Edward
'**日    期：2010-12-11 16:05:38
'**版    本：V1.1321
'*************************************************************************
Public Sub ApplyMapObjects()
Dim i As Long, j As Integer

For i = 0 To N_MOs - 1
   With MapObj(i)
      For j = 1 To 3
        Parties(i).InitPos(j).X = Format(.Body.X, "0.000000")
        Parties(i).InitPos(j).Y = Format(.Body.Y, "0.000000")
        Parties(i).Degree = Format(.Degree, "0.000000")
      Next j
   End With
Next i

End Sub
