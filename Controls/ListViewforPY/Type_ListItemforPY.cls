VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Type_ListItemforPY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Text As String
Public Value As String    'Reserved
Public OpLabel As Integer
Public Locked As Boolean
  
Public ParamCount As Integer
Public Params() As New Type_Param_ListItemforPY
  
Public Indention As Integer

Private MainLabel As Label


