VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEncodeErr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Source As String
Public Reason As Integer
Public Start As Integer
Public Length As Integer
Public Line As Integer
Public Block As Integer
Public Param As Integer

Public Function ShowEncodeErrMsg()
Dim Msg As String

Msg = ActiveString(PublicMsgs(152), Reason, GetReasonDescription(Reason), Block, Param)

MsgBox Msg, vbCritical, Source
End Function

Private Function GetReasonDescription(Reason As Integer) As String

Select Case Reason
  Case ERR_BAD_OPERATION
    GetReasonDescription = PublicMsgs(149)
  Case ERR_NULL_STRING
    GetReasonDescription = PublicMsgs(150)
  Case ERR_INCOMPLETE_SOURCE_CODE
    GetReasonDescription = PublicMsgs(151)
End Select

End Function
