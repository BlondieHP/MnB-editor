VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim a As Double, b As String, c As Double, d As Double, temI64(2) As Integer64b

Me.Show

b = "4"
temI64(0) = StrToI64(b)
b = "8"
temI64(1) = StrToI64(b)
temI64(2) = And64b(temI64(0), temI64(1))

Print IsZero64b(temI64(2))


End Sub

Function max_player_rating(rating As Long) As Long
Dim r As Long, a As Integer64b
r = 100 - rating
a = StrToI64(CStr(r))
 LeftMv64bEx a, 8
 max_player_rating = Val(I64toStrNZ(a)) And ff_max_rating_mask
End Function
  
