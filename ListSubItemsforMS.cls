VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListSubItemsforMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ListSubItem() As New ListSubItemforMS
Private ListView As ListViewforMS
Private lIndex As Integer
Private DefColor As Long
Private DefFont As New StdFont
Private DefLocked As Boolean
Public Parent As ListItemforMS

Public Property Get Item(ByVal Index As Integer) As ListSubItemforMS
Attribute Item.VB_UserMemId = 0
Set Item = ListSubItem(Index)
End Property

Public Property Let Item(ByVal Index As Integer, ByVal vNewValue As ListSubItemforMS)
Dim l As Label
Set ListSubItem(Index) = vNewValue

Set l = ListSubItem(Index).Label
l.Tag = TagAssign(l.Tag, 1, ListSubItem(Index).Index)

End Property

Public Sub Initialize(UserList As ListViewforMS, lParent As ListItemforMS)

Set ListView = UserList
Set Parent = lParent

lIndex = lParent.Index

End Sub

Private Sub Class_Initialize()
ReDim ListSubItem(0)
End Sub

Public Sub Clear()
Dim i As Integer

For i = 1 To UBound(ListSubItem)
  If ListSubItem(i).Initialized Then
    UnLoad ListSubItem(i).Label
  End If
Next i

ReDim ListSubItem(0)

End Sub

Public Property Get Count() As Integer
Count = UBound(ListSubItem)
End Property

Public Property Let Count(ByVal vNewValue As Integer)
Dim i As Integer, OldBound As Integer, l As Label

  OldBound = UBound(ListSubItem)

  For i = vNewValue + 1 To OldBound
    If ListSubItem(i).Initialized Then
      ListSubItem(i).Label.Visible = False
    End If
  Next i


  ReDim Preserve ListSubItem(vNewValue)

  For i = OldBound + 1 To UBound(ListSubItem)
    If Not ListSubItem(i).Initialized Then
      Set l = ListView.FreeLabel()
      l.ForeColor = DefColor
      l.Font = DefFont
      l.Visible = True
      ListSubItem(i).Locked = DefLocked
      ListSubItem(i).Initialize lIndex, i, ListView, l, Me
    End If
  Next i


End Property

Public Sub SetCount(ByVal NewCount As Integer)
Attribute SetCount.VB_MemberFlags = "40"
ReDim ListSubItem(NewCount)
End Sub

Public Property Get ForeColor() As Long
ForeColor = DefColor

End Property

Public Property Let ForeColor(ByVal vNewValue As Long)
Dim i As Integer

DefColor = vNewValue
For i = 1 To UBound(ListSubItem)
  If ListSubItem(i).Initialized Then
    ListSubItem(i).ForeColor = DefColor
  End If
Next i

End Property

Public Property Get FontBold() As Boolean
FontBold = DefFont.Bold
End Property

Public Property Let FontBold(ByVal vNewValue As Boolean)
Dim i As Integer
DefFont.Bold = vNewValue

For i = 1 To UBound(ListSubItem)
  If ListSubItem(i).Initialized Then
    ListSubItem(i).FontBold = DefFont.Bold
  End If
Next i
End Property

Public Sub KeyWord(ParamArray KeyWords() As Variant)
Dim i As Integer

For i = 0 To UBound(KeyWords)
  If i + 1 > UBound(ListSubItem) Then Exit For
  ListSubItem(i + 1).KeyWord = KeyWords(i)
Next i

End Sub

Private Sub Class_Terminate()
Dim i As Integer

Set ListView = Nothing
For i = 1 To UBound(ListSubItem)
  Set ListSubItem(i) = Nothing
Next i

End Sub

Public Property Get Locked() As Boolean
Locked = DefLocked

End Property

Public Property Let Locked(ByVal vNewValue As Boolean)
Dim i As Integer

DefLocked = vNewValue
For i = 1 To UBound(ListSubItem)
  ListSubItem(i).Locked = DefLocked
Next i

End Property


Public Sub Swap(Item1 As Integer, Item2 As Integer)
Dim tItem As New ListSubItemforMS

CopySubItem tItem, ListSubItem(Item1), lIndex, True

CopySubItem ListSubItem(Item1), ListSubItem(Item2), lIndex, False

CopySubItem ListSubItem(Item2), tItem, lIndex, False


End Sub

Public Sub Remove(Item As Integer)   'Call_Script Only
Dim i As Integer, k As String, l As String, r As String, j As Integer

For i = Item + 1 To UBound(ListSubItem)
  Swap i - 1, i
  If i < UBound(ListSubItem) Then ListSubItem(i - 1).KeyWord = ActiveString(PublicMsgs(156), i - 2)
Next i

k = ListSubItem(UBound(ListSubItem) - 1).KeyWord
Parent.Text = Operation(0).Pseudo

For j = 2 To UBound(ListSubItem) - 2
  i = InStrRev(Parent.Text, k)
  l = Left(Parent.Text, i - 1)
  r = Right(Parent.Text, Len(k))
  Parent.Text = l & ListSubItem(j).KeyWord & "," & r
Next j

ListSubItem(UBound(ListSubItem)).Label.Visible = False
ReDim Preserve ListSubItem(UBound(ListSubItem) - 1)


End Sub
