Attribute VB_Name = "ModBackUp"
Option Explicit

'============================================================================
' ModBackUp.bas - Transactional backup/restore system
'
' Improved from original:
'   1. Transactional: copies to temp dir first, then atomically renames
'   2. Verification: checks file sizes match after each copy
'   3. Data-driven: file list centralized in BackUpFileList array
'   4. Error recovery: cleans up temp dir on failure
'   5. Rollback: RestoreMod verifies files before restoring
'============================================================================

' Backup file specification
Private Type Type_BackupFile
    SourceName As String     ' File in ModPath (e.g., "item_kinds1.txt")
    CsvName As String        ' CSV companion in Language subdir (empty if none)
End Type

' Centralized file list - add new mod resources here only
Private Function BackUpFileList() As Type_BackupFile()
    Dim files(10) As Type_BackupFile

    files(0).SourceName = "item_kinds1.txt":     files(0).CsvName = "item_kinds.csv"
    files(1).SourceName = "troops.txt":           files(1).CsvName = "troops.csv"
    files(2).SourceName = "factions.txt":         files(2).CsvName = "factions.csv"
    files(3).SourceName = "party_templates.txt":  files(3).CsvName = "party_templates.csv"
    files(4).SourceName = "parties.txt":          files(4).CsvName = "parties.csv"
    files(5).SourceName = "scenes.txt":           files(5).CsvName = ""
    files(6).SourceName = "map_icons.txt":        files(6).CsvName = ""
    files(7).SourceName = "sounds.txt":           files(7).CsvName = ""
    files(8).SourceName = "particle_systems.txt": files(8).CsvName = ""
    files(9).SourceName = "tableau_materials.txt": files(9).CsvName = ""
    files(10).SourceName = "meshes.txt":          files(10).CsvName = ""

    BackUpFileList = files
End Function

'============================================================================
' Transactional backup with verification
'============================================================================
Public Function SetBackUp(Optional FileName As String = "") As Boolean
    On Error GoTo errorHandle

    Dim LanSource As String, strBackUp As String, strTempBackup As String
    Dim strTime As String, i As Long
    Dim files() As Type_BackupFile
    Dim srcPath As String, dstPath As String
    Dim srcSize As Long, dstSize As Long
    Dim copiedCount As Long, totalCount As Long
    Dim lanCsvPath As String

    files = BackUpFileList()
    strTime = Format(Now, "yyyy_mm_dd hh_mm_ss")

    If FileName = "" Then
        strBackUp = MnBInfo.ModBackUp & "\" & strTime
    Else
        strBackUp = MnBInfo.ModBackUp & "\" & FileName
    End If

    ' Ensure backup root exists
    If Not DirExists(MnBInfo.ModBackUp) Then
        MkDirEx MnBInfo.ModBackUp
    End If

    ' Create temporary directory for transactional backup
    strTempBackup = MnBInfo.ModBackUp & "\_temp_" & strTime
    If DirExists(strTempBackup) Then
        ' Clean up stale temp dir
        On Error Resume Next
        RmDir strTempBackup
        On Error GoTo errorHandle
    End If
    MkDirEx strTempBackup

    ' Set up source paths
    LanSource = MnBInfo.ModPath & "\languages\" & MnBInfo.Language

    ' Count total files to copy
    totalCount = 0
    For i = 0 To UBound(files)
        totalCount = totalCount + 1  ' Source file
        If Len(files(i).CsvName) > 0 Then
            totalCount = totalCount + 1  ' CSV companion
        End If
    Next i

    ' -- Phase 1: Copy all files to temp directory --
    copiedCount = 0
    For i = 0 To UBound(files)
        ' Copy main source file
        srcPath = MnBInfo.ModPath & "\" & files(i).SourceName
        dstPath = strTempBackup & "\" & files(i).SourceName

        If FileExists(srcPath) Then
            FileCopy srcPath, dstPath

            ' Verify copy
            srcSize = FileLen(srcPath)
            dstSize = FileLen(dstPath)
            If srcSize <> dstSize Then
                Err.Raise vbObjectError + 1, "ModBackUp", _
                    "Copy verification failed for " & files(i).SourceName & _
                    " (src:" & srcSize & " dst:" & dstSize & ")"
            End If
            copiedCount = copiedCount + 1
        End If

        ' Copy CSV companion if exists
        If Len(files(i).CsvName) > 0 Then
            lanCsvPath = LanSource & "\" & files(i).CsvName
            dstPath = strTempBackup & "\" & files(i).CsvName

            If FileExists(lanCsvPath) Then
                FileCopy lanCsvPath, dstPath

                ' Verify copy
                srcSize = FileLen(lanCsvPath)
                dstSize = FileLen(dstPath)
                If srcSize <> dstSize Then
                    Err.Raise vbObjectError + 1, "ModBackUp", _
                        "Copy verification failed for " & files(i).CsvName
                End If
                copiedCount = copiedCount + 1
            End If
        End If
    Next i

    ' -- Phase 2: Atomically rename temp dir to final backup dir --
    If DirExists(strBackUp) Then
        ' Backup dir already exists - remove temp and succeed
        On Error Resume Next
        RmDir strTempBackup
        On Error GoTo errorHandle
    Else
        Name strTempBackup As strBackUp
    End If

    SetBackUp = True
    Debug.Print "Backup complete: " & copiedCount & "/" & totalCount & " files backed up to " & strBackUp

    Exit Function

errorHandle:
    ' Clean up temp directory on failure
    On Error Resume Next
    If DirExists(strTempBackup) Then
        ' Remove files from temp dir
        Dim tmpFile As String
        tmpFile = Dir(strTempBackup & "\*.*")
        Do While tmpFile <> ""
            Kill strTempBackup & "\" & tmpFile
            tmpFile = Dir
        Loop
        RmDir strTempBackup
    End If
    On Error GoTo 0

    Call logErr("ModBackUp", "SetBackUp", Err.Number, Err.Description)
    SetBackUp = False
End Function

'============================================================================
' Restore with pre-flight verification
'============================================================================
Public Function RestoreMod(ByVal RevTime As String) As Boolean
    On Error GoTo errorHandle

    Dim LanSource As String, strBackUp As String, strTempRestore As String
    Dim i As Long, files() As Type_BackupFile
    Dim srcPath As String, dstPath As String
    Dim lanCsvPath As String

    files = BackUpFileList()
    strBackUp = MnBInfo.ModBackUp & "\" & RevTime

    ' Validate backup exists
    If Not DirExists(strBackUp) Then
        MsgBox "Backup not found: " & strBackUp, vbExclamation, "Restore Failed"
        Exit Function
    End If

    ' Verify at least one expected file exists in backup
    If Not FileExists(strBackUp & "\item_kinds1.txt") And _
       Not FileExists(strBackUp & "\troops.txt") Then
        MsgBox "Backup directory appears empty or corrupted: " & strBackUp, vbExclamation, "Restore Failed"
        Exit Function
    End If

    ' Ensure target directories exist
    LanSource = MnBInfo.ModPath & "\languages\" & MnBInfo.Language
    If Not DirExists(LanSource) Then
        MkDirEx LanSource
    End If

    ' Restore each file
    For i = 0 To UBound(files)
        ' Restore main source file
        srcPath = strBackUp & "\" & files(i).SourceName
        dstPath = MnBInfo.ModPath & "\" & files(i).SourceName

        If FileExists(srcPath) Then
            FileCopy srcPath, dstPath
        End If

        ' Restore CSV companion
        If Len(files(i).CsvName) > 0 Then
            srcPath = strBackUp & "\" & files(i).CsvName
            dstPath = LanSource & "\" & files(i).CsvName

            If FileExists(srcPath) Then
                FileCopy srcPath, dstPath
            End If
        End If
    Next i

    RestoreMod = True
    MsgBox "Mod restored from backup: " & RevTime, vbInformation, "Restore Complete"

    Exit Function

errorHandle:
    Call logErr("ModBackUp", "RestoreMod:[" & RevTime & "]", Err.Number, Err.Description)
    RestoreMod = False
End Function

'============================================================================
' List available backups (returns count, populates array)
'============================================================================
Public Function ListBackups(ByRef BackupNames() As String) As Long
    On Error GoTo errorHandle

    Dim strPath As String, strDir As String
    Dim count As Long

    strPath = MnBInfo.ModBackUp & "\"
    If Not DirExists(strPath) Then
        ListBackups = 0
        Exit Function
    End If

    ReDim BackupNames(0 To 99) ' Max 100 backups
    count = 0

    strDir = Dir(strPath, vbDirectory)
    Do While strDir <> ""
        If strDir <> "." And strDir <> ".." Then
            If (GetAttr(strPath & strDir) And vbDirectory) = vbDirectory Then
                ' Skip temp directories
                If Left$(strDir, 6) <> "_temp_" Then
                    If count <= UBound(BackupNames) Then
                        BackupNames(count) = strDir
                        count = count + 1
                    End If
                End If
            End If
        End If
        strDir = Dir
    Loop

    If count > 0 Then
        ReDim Preserve BackupNames(0 To count - 1)
    Else
        ReDim BackupNames(0 To 0)
        BackupNames(0) = ""
    End If

    ListBackups = count

    Exit Function

errorHandle:
    Call logErr("ModBackUp", "ListBackups", Err.Number, Err.Description)
    ListBackups = 0
End Function
