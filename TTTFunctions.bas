Attribute VB_Name = "ApiBas"
'Thanks for downloading my source from www.planetsourcecode.com

'If you have any Questions,Problems,or Comments
'E-mail me at: MiKE_3D@hotmail.com
'Or visit my website at: http://ww.8op.com/mike3d

'-MiKE 3D
Option Explicit

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Public Declare Sub ReleaseCapture Lib "user32" ()

Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const SWP_NOSIZE = 1
Global Const SWP_NOMOVE = &H2

Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const EM_SETREADONLY = &HCF
Public Const FLAGS = SWP_NOSIZE Or SWP_NOMOVE

Global TuRn As String
Global startTuRn As Integer
Global NickName As String
Global Counter As Integer
Global Counter2 As String

Type FILETIME
lLowDateTime    As Long
lHighDateTime   As Long
End Type


Public Function TimeOUT(HesitateTime)
Dim Hesitator As Long
Hesitator& = Timer
Do While Timer - Hesitator& < Val(HesitateTime)
DoEvents
Loop
End Function
Function ButtonDOWN(TheHandle, Times)
If Times > 500 Then Times = 500
Dim X As Integer
For X = 1 To Times
SendMessage TheHandle, WM_LBUTTONDOWN, 0, 0
Next X
End Function
Function ButtonUP(TheHandle, Times)
If Times > 500 Then Times = 500
Dim X As Integer
For X = 1 To Times
SendMessage TheHandle, WM_LBUTTONUP, 0, 0
Next X
End Function

Sub StayOnTop(TheForm As Form)
Dim SetWinOnTop As Long
SetWinOnTop = SetWindowPos(TheForm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Sub NotOnTop(frm As Form)
Dim SetWinOnTop As Long
SetWinOnTop = SetWindowPos(frm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Sub CenterForm(TENProg As Form)
TENProg.Move (Screen.Width) / 2 - (TENProg.Width) / 2, (Screen.Height) / 2 - (TENProg.Height) / 2
End Sub
Sub DisableTextbox(txtValue As TextBox)
SendMessage txtValue.hWnd, EM_SETREADONLY, 1, 0
End Sub
Sub EnableTextbox(txtValue As TextBox)
SendMessage txtValue.hWnd, EM_SETREADONLY, 0, 0
End Sub

Public Sub FormDrag(TheForm As Form)
ReleaseCapture
Call SendMessage(TheForm.hWnd, &HA1, 2, 0&)
End Sub
