Attribute VB_Name = "ModError"
Option Explicit

'============================================================================
' ModError.bas - Standardized error handling utilities
'
' Provides consistent error handling patterns across all modules.
' Replaces scattered On Error Resume Next / Debug.Print patterns.
'
' Usage:
'   On Error GoTo EH
'   ' ... code ...
'   Exit Sub/Function
' EH:
'   HandleError "ModuleName", "FunctionName"
'============================================================================

' Centralized error handler - logs the error and optionally shows a message
Public Sub HandleError(ByVal ModuleName As String, ByVal FuncName As String, _
                        Optional ByVal ShowMessage As Boolean = False)
    Dim errMsg As String

    errMsg = ModuleName & ":" & FuncName & _
             ", Err.Number=" & Err.Number & _
             " : Error=" & Err.Description

    ' Always log to debug
    Debug.Print "[" & Format(Now, "YYYY-MM-DD HH:MM:SS") & "] " & errMsg

    ' Log to debug form if available
    On Error Resume Next
    Call OutAsDebugTex(errMsg)
    On Error GoTo 0

    ' Optionally show to user
    If ShowMessage Then
        MsgBox errMsg, vbExclamation, "Error in " & ModuleName
    End If

    ' Reset mouse cursor
    SetMouseDefault
End Sub

' Safe file existence check - handles errors internally
Public Function FileExistsSafe(ByVal FilePath As String) As Boolean
    On Error GoTo ErrHandler
    FileExistsSafe = (Len(Dir(FilePath)) > 0)
    Exit Function
ErrHandler:
    FileExistsSafe = False
End Function

' Safe directory existence check
Public Function DirExistsSafe(ByVal DirPath As String) As Boolean
    On Error GoTo ErrHandler
    DirExistsSafe = (Len(Dir(DirPath, vbDirectory)) > 0)
    Exit Function
ErrHandler:
    DirExistsSafe = False
End Function

' Safe numeric conversion - returns 0 on failure instead of crashing
Public Function SafeVal(ByVal str As String) As Long
    On Error GoTo ErrHandler
    SafeVal = Val(str)
    Exit Function
ErrHandler:
    SafeVal = 0
End Function

' Safe hex conversion
Public Function SafeHex(ByVal Value As Variant) As String
    On Error GoTo ErrHandler
    SafeHex = Hex(Val(Value))
    Exit Function
ErrHandler:
    SafeHex = "0"
End Function

' Try to execute a statement; returns True if no error occurred
Public Function TryExecute(ByRef Success As Boolean) As Boolean
    Success = (Err.Number = 0)
    TryExecute = Success
End Function
