VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl RichforTXT 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8655
   ScaleHeight     =   6000
   ScaleWidth      =   8655
   Begin RichTextLib.RichTextBox txtMain 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   10398
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      TextRTF         =   $"RichforTXT.ctx":0000
   End
End
Attribute VB_Name = "RichforTXT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private pColor(2) As Long
Private CustomActive As Boolean

Public Event Change()

Private Sub txtMain_Change()
If CustomActive Then
  RaiseEvent Change
End If
End Sub

Private Sub UserControl_Initialize()
CustomActive = True
AutoSwitchLine txtMain, True
End Sub

Private Sub UserControl_Resize()
txtMain.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub

Public Property Get Text() As String
Text = txtMain.Text
End Property

Public Property Let Text(ByVal vNewValue As String)
DrawText vNewValue

PropertyChanged Text

End Property

Public Property Get OperationColor() As Long
OperationColor = pColor(0)
End Property

Public Property Let OperationColor(ByVal vNewValue As Long)
pColor(0) = vNewValue

PropertyChanged OperationColor
End Property

Public Property Get CountColor() As Long
CountColor = pColor(1)
End Property

Public Property Let CountColor(ByVal vNewValue As Long)
pColor(1) = vNewValue

PropertyChanged CountColor
End Property

Public Property Get ParamColor() As Long
ParamColor = pColor(2)
End Property

Public Property Let ParamColor(ByVal vNewValue As Long)
pColor(2) = vNewValue

PropertyChanged ParamColor
End Property

Private Sub SetRichTextColor(ByVal Start As Long, ByVal Length As Long, ByVal lColor As Long, Optional Restore As Boolean = True, Optional DefColor As Long)
Dim i As Long, j As Long, l As Long, c As Long

CustomActive = False

If Restore Then
  j = txtMain.SelStart
  l = txtMain.SelLength
  c = txtMain.SelColor
End If

txtMain.SelStart = Start
txtMain.SelLength = Length
txtMain.SelColor = lColor

If Restore Then
  txtMain.SelStart = j
  txtMain.SelLength = l

  If Not IsMissing(DefColor) Then
     txtMain.SelColor = DefColor
  Else
     txtMain.SelColor = c
  End If
End If

CustomActive = True
End Sub

Public Sub RefreshColor()
Dim strTem() As String, i As Long, H As Long, V As Long, m As Long, l As Long, n As Long, TemVal(1) As Long

If Trim(txtMain.Text) = "" Then Exit Sub

CustomActive = False
TemVal(0) = txtMain.SelStart
TemVal(1) = txtMain.SelLength

strTem = Split(txtMain.Text, " ")
m = 0

Do While i <= UBound(strTem)
  
  If m = 2 And l = 0 Then
    m = 0
  End If
  
  SetRichTextColor H, Len(strTem(i)), pColor(m), False

  If m = 0 Then
    m = 1
  ElseIf m = 1 Then
    V = Val(strTem(i))
    m = 2
    l = V
    n = 0
  ElseIf m = 2 Then
    n = n + 1
    If n >= l Then
      m = 0
    End If
  End If
  
  H = H + Len(strTem(i)) + 1
  
  i = i + 1
Loop

txtMain.SelStart = TemVal(0)
txtMain.SelLength = TemVal(1)

CustomActive = True
End Sub

Public Sub DrawText(Text As String)
Dim strTem() As String, i As Long, H As Long, V As Long, m As Long, l As Long, n As Long

txtMain = ""
If Trim(Text) = "" Then Exit Sub

CustomActive = False

strTem = Split(Text, " ")
m = 0

Do While i <= UBound(strTem)
  
  If m = 2 And l = 0 Then
    m = 0
  End If
  
  txtMain.SelColor = pColor(m)
  txtMain.SelStart = Len(txtMain.Text)
  txtMain.SelText = IIf(i = 0, "", " ") & strTem(i)
  
  If m = 0 Then
    m = 1
  ElseIf m = 1 Then
    V = Val(strTem(i))
    m = 2
    l = V
    n = 0
  ElseIf m = 2 Then
    n = n + 1
    If n >= l Then
      m = 0
    End If
  End If
  
  i = i + 1
Loop

txtMain.SelLength = 0

CustomActive = True

RaiseEvent Change
End Sub


Public Property Get SelStart() As Long
SelStart = txtMain.SelStart
End Property

Public Property Let SelStart(ByVal vNewValue As Long)
txtMain.SelStart = vNewValue

PropertyChanged (SelStart)
End Property


Public Property Get SelLength() As Long
SelLength = txtMain.SelLength
End Property

Public Property Let SelLength(ByVal vNewValue As Long)
txtMain.SelLength = vNewValue

PropertyChanged (SelLength)
End Property

Public Property Get SelColor() As Long
SelColor = txtMain.SelColor
End Property

Public Property Let SelColor(ByVal vNewValue As Long)
txtMain.SelColor = vNewValue

PropertyChanged (SelColor)
End Property
