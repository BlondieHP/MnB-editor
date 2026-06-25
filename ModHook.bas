Attribute VB_Name = "ModHook"
'模块代码
Option Explicit
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'设置窗口信息
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'调用窗口处理
Private Const WM_RBUTTONDOWN = &H204
Private Const GWL_WNDPROC = (-4)
Private lpOldWndFunc As Long
'--------------------------------------------------------------------------------------
'函 数 名: WindowProcedure
'描    述: 窗口消息处理函数
'--------------------------------------------------------------------------------------
Private Function WindowProcedure(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        Select Case wMsg
               Case WM_RBUTTONDOWN
               Case Else
                    WindowProcedure = CallWindowProc(lpOldWndFunc, hWnd, wMsg, wParam, _
                    lParam)                                                             '原窗口消息处理
        End Select
End Function
'--------------------------------------------------------------------------------------
'函 数 名: Hook
'描    述: 子类化窗口
'--------------------------------------------------------------------------------------
Public Sub Hook(ByVal hWnd As Long)
       lpOldWndFunc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProcedure)       '新窗口消息处理
End Sub
'--------------------------------------------------------------------------------------
'函 数 名: UnHook
'描    述: 取消子类化
'--------------------------------------------------------------------------------------
Public Sub UnHook(ByVal hWnd As Long)
       Call SetWindowLong(hWnd, GWL_WNDPROC, lpOldWndFunc)                              '恢复原窗口消息处理
End Sub


