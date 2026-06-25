Attribute VB_Name = "CsvLoader"
Option Explicit

' CSV Entity Type enumeration for unified CSV loading
Private Enum CSVEntityType
    CSVET_Party = 1
    CSVET_PartyTemplate = 2
    CSVET_Item = 3
    CSVET_Troop = 4
    CSVET_ItemModifier = 5
    CSVET_Faction = 6
    CSVET_QuickString = 7
    CSVET_String = 8
End Enum

Public Sub clearCSVName()

    Dim n As Integer
    For n = 0 To N_Troop - 1
        Trps(n).csvName = ""
    Next

    For n = 0 To N_Item - 1
        itm(n).csvName = ""
    Next

    For n = 0 To N_PT - 1
        PTs(n).csvName = ""
    Next
End Sub

'==========================================================================
' Unified CSV file loader - O(n) optimized with Dictionary lookup
'
' Builds a Dictionary (hash map) of entity ID -> index in O(n) time,
' then looks up each CSV line in O(1) instead of the original O(n*m)
' nested loop. For a large mod with 2000 items and 2000 CSV lines,
' this reduces ~4,000,000 comparisons to ~4,000 operations.
'==========================================================================
Private Function LoadCSVFileGeneric(ByVal FileName As String, ByVal EntityType As CSVEntityType) As Boolean
    On Error GoTo errorHandle

    LoadCSVFileGeneric = False

    ' Language check: English uses internal names, no CSV needed
    If LCase$(MnBInfo.Language) = "en" Then
        LoadCSVFileGeneric = True
        Exit Function
    End If

    ' Missing CSV flag check
    If RunERR.MissCSV Then
        Exit Function
    End If

    ' File existence check
    Dim tmpFileName As String
    tmpFileName = FileName

    If FileLen(tmpFileName) = 0 Then
        MsgBox "Missing file: " & tmpFileName
        Exit Function
    End If

    ' Load CSV file
    Dim TemP As String
    Dim arrTmp() As String
    Dim arrFileBuff() As String
    Dim n As Long, i As Long

    TemP = UEFLoadTextFile(tmpFileName, UEF_UTF8)
    If TemP = vbNullString Then
        MsgBox "Null string from file"
        Exit Function
    End If
    arrFileBuff = Split(TemP, vbCrLf)

    ' Build Dictionary: LCase(key) -> index for O(1) lookup
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 0  ' Binary compare (case-sensitive; we use LCase keys)

    Select Case EntityType
        Case CSVET_Party
            For n = 0 To N_Party - 1
                dict.Add LCase$(Parties(n).strID), n
            Next
            For i = 0 To UBound(arrFileBuff)
                TemP = arrFileBuff(i)
                If Len(Trim$(TemP)) >= 1 Then
                    arrTmp = Split(TemP, "|")
                    If dict.Exists(LCase$(arrTmp(0))) Then
                        n = dict(LCase$(arrTmp(0)))
                        Parties(n).csvName = arrTmp(1)
                    End If
                End If
            Next

        Case CSVET_PartyTemplate
            For n = 0 To N_PT - 1
                dict.Add LCase$(PTs(n).ptID), n
            Next
            For i = 0 To UBound(arrFileBuff)
                TemP = arrFileBuff(i)
                If Len(Trim$(TemP)) >= 1 Then
                    arrTmp = Split(TemP, "|")
                    If dict.Exists(LCase$(arrTmp(0))) Then
                        n = dict(LCase$(arrTmp(0)))
                        PTs(n).csvName = arrTmp(1)
                    End If
                End If
            Next

        Case CSVET_Item
            For n = 0 To N_Item - 1
                dict.Add LCase$(itm(n).dbName), n
            Next
            For i = 0 To UBound(arrFileBuff)
                TemP = arrFileBuff(i)
                If Len(Trim$(TemP)) >= 1 Then
                    arrTmp = Split(TemP, "|")
                    ' Handle plural form: "xxx_pl" -> csvName_pl
                    If Right$(arrTmp(0), 3) = "_pl" Then
                        Dim baseKey As String
                        baseKey = LCase$(Left$(arrTmp(0), Len(arrTmp(0)) - 3))
                        If dict.Exists(baseKey) Then
                            itm(dict(baseKey)).csvName_pl = arrTmp(1)
                        End If
                    ElseIf dict.Exists(LCase$(arrTmp(0))) Then
                        itm(dict(LCase$(arrTmp(0)))).csvName = arrTmp(1)
                    End If
                End If
            Next

        Case CSVET_Troop
            For n = 0 To N_Troop - 1
                dict.Add LCase$(Trps(n).strID), n
            Next
            For i = 0 To UBound(arrFileBuff)
                TemP = arrFileBuff(i)
                If Len(Trim$(TemP)) >= 1 Then
                    arrTmp = Split(TemP, "|")
                    ' Handle plural form: "xxx_pl" -> csvName_pl
                    If Right$(arrTmp(0), 3) = "_pl" Then
                        Dim baseKeyTrp As String
                        baseKeyTrp = LCase$(Left$(arrTmp(0), Len(arrTmp(0)) - 3))
                        If dict.Exists(baseKeyTrp) Then
                            Trps(dict(baseKeyTrp)).csvName_pl = arrTmp(1)
                        End If
                    ElseIf dict.Exists(LCase$(arrTmp(0))) Then
                        Trps(dict(LCase$(arrTmp(0)))).csvName = arrTmp(1)
                    End If
                End If
            Next

        Case CSVET_ItemModifier
            ' Build reverse dictionary from CSV lines, then match IMod entries
            For i = 0 To UBound(arrFileBuff)
                TemP = arrFileBuff(i)
                If Len(Trim$(TemP)) >= 1 Then
                    arrTmp = Split(TemP, "|")
                    ' Store pre-processed value (strip %s, trim)
                    Dim cleanVal As String
                    cleanVal = Trim$(Replace(arrTmp(1), "%s", ""))
                    dict.Add LCase$(Trim$(arrTmp(0))), cleanVal
                End If
            Next
            For i = 0 To N_IMod - 1
                If dict.Exists(LCase$(Trim$(IMod(i).ID))) Then
                    IMod(i).csvName = dict(LCase$(Trim$(IMod(i).ID)))
                End If
            Next

        Case CSVET_Faction
            For n = 0 To N_Faction - 1
                dict.Add LCase$(Factions(n).strID), n
            Next
            For i = 0 To UBound(arrFileBuff)
                TemP = arrFileBuff(i)
                If Len(Trim$(TemP)) >= 1 Then
                    arrTmp = Split(TemP, "|")
                    If dict.Exists(LCase$(arrTmp(0))) Then
                        Factions(dict(LCase$(arrTmp(0)))).csvName = arrTmp(1)
                    End If
                End If
            Next

        Case CSVET_QuickString
            For n = 0 To N_qStr - 1
                dict.Add LCase$(qStrs(n).Name), n
            Next
            For i = 0 To UBound(arrFileBuff)
                TemP = arrFileBuff(i)
                If Len(Trim$(TemP)) >= 1 Then
                    arrTmp = Split(TemP, "|")
                    If dict.Exists(LCase$(arrTmp(0))) Then
                        qStrs(dict(LCase$(arrTmp(0)))).CSV = arrTmp(1)
                    End If
                End If
            Next

        Case CSVET_String
            For n = 0 To N_Str - 1
                dict.Add LCase$(Strs(n).Name), n
            Next
            For i = 0 To UBound(arrFileBuff)
                TemP = arrFileBuff(i)
                If Len(Trim$(TemP)) >= 1 Then
                    arrTmp = Split(TemP, "|")
                    If dict.Exists(LCase$(arrTmp(0))) Then
                        Strs(dict(LCase$(arrTmp(0)))).CSV = arrTmp(1)
                    End If
                End If
            Next

    End Select

    Set dict = Nothing
    LoadCSVFileGeneric = True

    Exit Function

errorHandle:
    Set dict = Nothing
    Call logErr("CsvLoader", "LoadCSVFile:[" & FileName & "]", Err.Number, Err.Description)
End Function

'==========================================================================
' Public wrapper functions - preserve original API for backward compatibility
'==========================================================================

Public Function LoadPartyCSVFile(FileName As String) As Boolean
    LoadPartyCSVFile = LoadCSVFileGeneric(FileName, CSVET_Party)
End Function

Public Function LoadPartyTemplateCSVFile(FileName As String) As Boolean
    LoadPartyTemplateCSVFile = LoadCSVFileGeneric(FileName, CSVET_PartyTemplate)
End Function

Public Function LoadItemCSVFile(FileName As String) As Boolean
    LoadItemCSVFile = LoadCSVFileGeneric(FileName, CSVET_Item)
End Function

Public Function LoadTroopCSVFile(FileName As String) As Boolean
    LoadTroopCSVFile = LoadCSVFileGeneric(FileName, CSVET_Troop)
End Function

Public Function LoadIModCSVFile(FileName As String) As Boolean
    LoadIModCSVFile = LoadCSVFileGeneric(FileName, CSVET_ItemModifier)
End Function

Public Function LoadFactionCSVFile(FileName As String) As Boolean
    LoadFactionCSVFile = LoadCSVFileGeneric(FileName, CSVET_Faction)
End Function

Public Function LoadQuickStringCSVFile(FileName As String) As Boolean
    LoadQuickStringCSVFile = LoadCSVFileGeneric(FileName, CSVET_QuickString)
End Function

Public Function LoadStringCSVFile(FileName As String) As Boolean
    LoadStringCSVFile = LoadCSVFileGeneric(FileName, CSVET_String)
End Function
