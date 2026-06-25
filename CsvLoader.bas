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
' Unified CSV file loader
' Replaces 8 separate nearly-identical Load*CSVFile functions
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

    ' Load and parse CSV file
    Dim TemP As String
    Dim arrTmp() As String
    Dim arrFileBuff() As String
    Dim n As Long, i As Long
    Dim H As Long, F As Boolean

    TemP = UEFLoadTextFile(tmpFileName, UEF_UTF8)
    If TemP = vbNullString Then
        MsgBox "Null string from file"
        Exit Function
    End If
    arrFileBuff = Split(TemP, vbCrLf)

    ' Entity-specific matching logic
    Select Case EntityType
        Case CSVET_Party
            For i = 0 To UBound(arrFileBuff)
                TemP = arrFileBuff(i)
                If Len(Trim$(TemP)) >= 1 Then
                    arrTmp = Split(TemP, "|")
                    For n = 0 To N_Party - 1
                        If LCase$(Parties(n).strID) = LCase$(arrTmp(0)) Then
                            Parties(n).csvName = arrTmp(1)
                        End If
                    Next
                End If
            Next

        Case CSVET_PartyTemplate
            For i = 0 To UBound(arrFileBuff)
                TemP = arrFileBuff(i)
                If Len(Trim$(TemP)) >= 1 Then
                    arrTmp = Split(TemP, "|")
                    For n = 0 To N_PT - 1
                        If LCase$(PTs(n).ptID) = LCase$(arrTmp(0)) Then
                            PTs(n).csvName = arrTmp(1)
                        End If
                    Next
                End If
            Next

        Case CSVET_Item
            For i = 0 To UBound(arrFileBuff)
                TemP = arrFileBuff(i)
                If Len(Trim$(TemP)) >= 1 Then
                    arrTmp = Split(TemP, "|")
                    For n = 0 To N_Item - 1
                        If LCase$(itm(n).dbName) = LCase$(arrTmp(0)) Then
                            itm(n).csvName = arrTmp(1)
                        End If
                        ' Handle plural form (_pl suffix)
                        If Right(arrTmp(0), 3) = "_pl" Then
                            If LCase$(itm(n).dbName) = LCase$(Left(arrTmp(0), Len(arrTmp(0)) - 3)) Then
                                itm(n).csvName_pl = arrTmp(1)
                            End If
                        End If
                    Next
                End If
            Next

        Case CSVET_Troop
            For i = 0 To UBound(arrFileBuff)
                TemP = arrFileBuff(i)
                If Len(Trim$(TemP)) >= 1 Then
                    arrTmp = Split(TemP, "|")
                    For n = 0 To N_Troop - 1
                        If LCase$(Trps(n).strID) = LCase$(arrTmp(0)) Then
                            Trps(n).csvName = arrTmp(1)
                        End If
                        ' Handle plural form (_pl suffix)
                        If Right(arrTmp(0), 3) = "_pl" Then
                            If LCase$(Trps(n).strID) = LCase$(Left(arrTmp(0), Len(arrTmp(0)) - 3)) Then
                                Trps(n).csvName_pl = arrTmp(1)
                            End If
                        End If
                    Next
                End If
            Next

        Case CSVET_ItemModifier
            ' Nested loop: outer over IMod entries, inner over file lines
            For i = 0 To N_IMod - 1
                For n = 0 To UBound(arrFileBuff)
                    TemP = arrFileBuff(n)
                    If Len(Trim$(TemP)) >= 1 Then
                        arrTmp = Split(TemP, "|")
                        If LCase(Trim(IMod(i).ID)) = LCase(Trim(arrTmp(0))) Then
                            arrTmp(1) = Replace(arrTmp(1), "%s", "")
                            arrTmp(1) = Trim(arrTmp(1))
                            IMod(i).csvName = arrTmp(1)
                            Exit For
                        End If
                    End If
                Next n
            Next i

        Case CSVET_Faction
            For i = 0 To UBound(arrFileBuff)
                TemP = arrFileBuff(i)
                If Len(Trim$(TemP)) >= 1 Then
                    arrTmp = Split(TemP, "|")
                    For n = 0 To N_Faction - 1
                        If LCase$(Factions(n).strID) = LCase$(arrTmp(0)) Then
                            Factions(n).csvName = arrTmp(1)
                        End If
                    Next
                End If
            Next

        Case CSVET_QuickString
            For i = 0 To UBound(arrFileBuff)
                TemP = arrFileBuff(i)
                If Len(Trim$(TemP)) >= 1 Then
                    arrTmp = Split(TemP, "|")
                    For n = 0 To N_qStr - 1
                        If LCase$(qStrs(n).Name) = LCase$(arrTmp(0)) Then
                            qStrs(n).CSV = arrTmp(1)
                        End If
                    Next
                End If
            Next

        Case CSVET_String
            ' Cursor-optimized: H tracks last matched position to skip already-matched entries
            H = 0
            For i = 0 To UBound(arrFileBuff)
                TemP = arrFileBuff(i)
                If Len(Trim$(TemP)) >= 1 Then
                    arrTmp = Split(TemP, "|")
                    F = False
                    For n = H To N_Str - 1
                        If LCase$(Strs(n).Name) = LCase$(arrTmp(0)) Then
                            Strs(n).CSV = arrTmp(1)
                            H = n + 1
                            F = True
                            Exit For
                        End If
                    Next n

                    ' Fallback: scan from beginning if not found after cursor
                    If Not F Then
                        For n = 0 To H - 1
                            If LCase$(Strs(n).Name) = LCase$(arrTmp(0)) Then
                                Strs(n).CSV = arrTmp(1)
                                Exit For
                            End If
                        Next n
                    End If
                End If
            Next

    End Select

    LoadCSVFileGeneric = True

    Exit Function

errorHandle:
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
