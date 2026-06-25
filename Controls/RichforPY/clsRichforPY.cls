VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRichforPY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents Rich As RichTextBox
Attribute Rich.VB_VarHelpID = -1
Private WithEvents Delayer As Timer
Attribute Delayer.VB_VarHelpID = -1

Public Sub Attach(RichText As RichTextBox, DelayTimer As Timer)
Set Rich = RichText
Set Delayer = DelayTimer
End Sub

Public Sub Unload()
Set Rich = Nothing
Set Delayer = Nothing
End Sub

