VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMainClient 
   BackColor       =   &H00000000&
   Caption         =   "TicTacToe Server - ßy: MiKE 3D"
   ClientHeight    =   3975
   ClientLeft      =   2160
   ClientTop       =   2040
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3975
   ScaleWidth      =   7710
   Begin VB.Timer GetWinner 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2880
      Top             =   120
   End
   Begin VB.Timer NickSend 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3360
      Top             =   120
   End
   Begin RichTextLib.RichTextBox MainTextBox 
      Height          =   2910
      Left            =   3840
      TabIndex        =   16
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   5133
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      TextRTF         =   $"frmMainServer.frx":0000
   End
   Begin VB.CommandButton cSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   6960
      TabIndex        =   14
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox SendTextBox 
      Height          =   375
      Left            =   3840
      MaxLength       =   200
      TabIndex        =   13
      Top             =   3120
      Width           =   3015
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3720
      Top             =   4080
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H00000000&
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3735
      Begin VB.PictureBox ClearPic 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   13200
         ScaleHeight     =   591.509
         ScaleMode       =   0  'User
         ScaleWidth      =   855
         TabIndex        =   29
         Top             =   1320
         Width           =   855
      End
      Begin VB.PictureBox cPicture9 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   855
         Left            =   2400
         ScaleHeight     =   591.509
         ScaleMode       =   0  'User
         ScaleWidth      =   855
         TabIndex        =   28
         Top             =   2400
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.PictureBox cPicture8 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   855
         Left            =   1320
         ScaleHeight     =   591.509
         ScaleMode       =   0  'User
         ScaleWidth      =   855
         TabIndex        =   27
         Top             =   2400
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.PictureBox cPicture7 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   855
         Left            =   240
         ScaleHeight     =   591.509
         ScaleMode       =   0  'User
         ScaleWidth      =   855
         TabIndex        =   26
         Top             =   2400
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.PictureBox cPicture6 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   855
         Left            =   2400
         ScaleHeight     =   591.509
         ScaleMode       =   0  'User
         ScaleWidth      =   855
         TabIndex        =   25
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.PictureBox cPicture5 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   855
         Left            =   1320
         ScaleHeight     =   591.509
         ScaleMode       =   0  'User
         ScaleWidth      =   855
         TabIndex        =   24
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.PictureBox cPicture4 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   855
         Left            =   240
         ScaleHeight     =   591.509
         ScaleMode       =   0  'User
         ScaleWidth      =   855
         TabIndex        =   23
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.PictureBox cPicture3 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   855
         Left            =   2400
         ScaleHeight     =   591.509
         ScaleMode       =   0  'User
         ScaleWidth      =   855
         TabIndex        =   22
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.PictureBox cPicture2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   855
         Left            =   1320
         ScaleHeight     =   591.509
         ScaleMode       =   0  'User
         ScaleWidth      =   855
         TabIndex        =   21
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.PictureBox cPicture1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   855
         Left            =   240
         ScaleHeight     =   591.509
         ScaleMode       =   0  'User
         ScaleWidth      =   855
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.PictureBox OPic 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   1320
         Picture         =   "frmMainServer.frx":00C9
         ScaleHeight     =   591.509
         ScaleMode       =   0  'User
         ScaleWidth      =   855
         TabIndex        =   12
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.PictureBox XPic 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   1320
         Picture         =   "frmMainServer.frx":1B36
         ScaleHeight     =   591.509
         ScaleMode       =   0  'User
         ScaleWidth      =   855
         TabIndex        =   11
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin RichTextLib.RichTextBox TExtHold 
         Height          =   135
         Left            =   27600
         TabIndex        =   10
         Top             =   240
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   238
         _Version        =   393217
         TextRTF         =   $"frmMainServer.frx":33A4
      End
      Begin VB.PictureBox Picture9 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   2400
         ScaleHeight     =   591.509
         ScaleMode       =   0  'User
         ScaleWidth      =   855
         TabIndex        =   9
         Top             =   2400
         Width           =   855
      End
      Begin VB.PictureBox Picture8 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   1320
         ScaleHeight     =   591.509
         ScaleMode       =   0  'User
         ScaleWidth      =   855
         TabIndex        =   8
         Top             =   2400
         Width           =   855
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   240
         ScaleHeight     =   591.509
         ScaleMode       =   0  'User
         ScaleWidth      =   855
         TabIndex        =   7
         Top             =   2400
         Width           =   855
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   2400
         ScaleHeight     =   591.509
         ScaleMode       =   0  'User
         ScaleWidth      =   855
         TabIndex        =   6
         Top             =   1320
         Width           =   855
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   1320
         ScaleHeight     =   591.509
         ScaleMode       =   0  'User
         ScaleWidth      =   855
         TabIndex        =   5
         Top             =   1320
         Width           =   855
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   240
         ScaleHeight     =   591.509
         ScaleMode       =   0  'User
         ScaleWidth      =   855
         TabIndex        =   4
         Top             =   1320
         Width           =   855
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   2400
         ScaleHeight     =   591.509
         ScaleMode       =   0  'User
         ScaleWidth      =   855
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   1320
         ScaleHeight     =   591.509
         ScaleMode       =   0  'User
         ScaleWidth      =   855
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   240
         ScaleHeight     =   591.509
         ScaleMode       =   0  'User
         ScaleWidth      =   855
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
      Begin MSWinsockLib.Winsock Wsck 
         Left            =   3240
         Top             =   3240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         X1              =   2280
         X2              =   2280
         Y1              =   0
         Y2              =   3480
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         X1              =   1200
         X2              =   1200
         Y1              =   0
         Y2              =   3480
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         X1              =   0
         X2              =   3480
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         X1              =   0
         X2              =   3480
         Y1              =   1200
         Y2              =   1200
      End
   End
   Begin VB.Label NickScore2 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H008080FF&
      Height          =   375
      Left            =   7140
      TabIndex        =   19
      Top             =   3675
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label NickScore1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H008080FF&
      Height          =   375
      Left            =   5400
      TabIndex        =   18
      Top             =   3670
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label NickName2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   5760
      TabIndex        =   17
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label NickName1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   3960
      TabIndex        =   15
      Top             =   3600
      Width           =   1215
   End
End
Attribute VB_Name = "frmMainClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cSend_Click()
On Error Resume Next
Wsck.SendData "T" & NickName & ":" & vbTab & SendTextBox
MainTextBox.SelColor = vbBlue
MainTextBox.SelBold = True
MainTextBox.SelText = NickName & ":" & vbTab
MainTextBox.SelBold = False
MainTextBox.SelText = SendTextBox & vbCrLf
SendTextBox = ""
End Sub

Private Sub Form_Load()
StayOnTop Me
CenterForm Me
cPicture1 = OPic
cPicture2 = OPic
cPicture3 = OPic
cPicture4 = OPic
cPicture5 = OPic
cPicture6 = OPic
cPicture7 = OPic
cPicture8 = OPic
cPicture9 = OPic
Counter = 0
TuRn = True
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState = 2 Then Me.WindowState = 0
If Me.Width <> 7830 Then Me.Width = 7830
If Me.Height <> 4425 Then Me.Height = 4425
End Sub

Private Sub Form_Unload(Cancel As Integer)
GetWinner.Enabled = False
X& = MsgBox("Are you sure you want to exit?", vbSystemModal + vbYesNo + vbInformation, "Exit TicTacToe")
GetWinner.Enabled = True
If X& = vbYes Then
Wsck.SendData "C" & "Disconnected"
Wsck.Close
End
End If
Cancel = 1
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub MainTextBox_Change()
On Error Resume Next
MainTextBox.SelLength = 0
If Len(MainTextBox.Text) > 0 Then
If Right$(MainTextBox.Text, 1) = vbCrLf Then
MainTextBox.SelStart = Len(MainTextBox.Text) - 1
Exit Sub
End If
MainTextBox.SelStart = Len(MainTextBox.Text)
End If
End Sub
Private Sub NickScore1_Change()
If NickScore1.Caption = "99" Then NickScore1.Caption = "0"
End Sub

Private Sub NickSend_Timer()
NickSend.Enabled = False
Wsck.SendData "N" & NickName
NickSend.Enabled = False
End Sub

Function GetNickName()
Nick$ = NickName1.Caption
GetNick$ = Mid(Nick$, InStr(Nick$, ":"), Len(Nick$))
GetNick3$ = Replace(Nick$, GetNick$, "")
GetNickName = GetNick3$
End Function

Private Sub Picture1_Click()
If TuRn = False Then
GetWinner.Enabled = False
MsgBox "Wait for " & GetNickName & "'s turn to finish!", vbSystemModal + vbCritical, "Impatient"
GetWinner.Enabled = True
Exit Sub
End If
If TuRn = True Then
TuRn = False
Counter = Val(Counter) + 1
Picture1 = XPic
Picture1.Enabled = False
Wsck.SendData "P" & "Picture1_Click"
End If
End Sub

Private Sub Picture2_Click()
If TuRn = False Then
GetWinner.Enabled = False
MsgBox "Wait for " & GetNickName & "'s turn to finish!", vbSystemModal + vbCritical, "Impatient"
GetWinner.Enabled = True
Exit Sub
End If
If TuRn = True Then
TuRn = False
Counter = Val(Counter) + 1
Picture2 = XPic
Picture2.Enabled = False
Wsck.SendData "P" & "Picture2_Click"
End If
End Sub

Private Sub Picture3_Click()
If TuRn = False Then
GetWinner.Enabled = False
MsgBox "Wait for " & GetNickName & "'s turn to finish!", vbSystemModal + vbCritical, "Impatient"
GetWinner.Enabled = True
Exit Sub
End If
If TuRn = True Then
TuRn = False
Counter = Val(Counter) + 1
Picture3 = XPic
Picture3.Enabled = False
Wsck.SendData "P" & "Picture3_Click"
End If
End Sub

Private Sub Picture4_Click()
If TuRn = False Then
GetWinner.Enabled = False
MsgBox "Wait for " & GetNickName & "'s turn to finish!", vbSystemModal + vbCritical, "Impatient"
GetWinner.Enabled = True
Exit Sub
End If
If TuRn = True Then
TuRn = False
Counter = Val(Counter) + 1
Picture4 = XPic
Picture4.Enabled = False
Wsck.SendData "P" & "Picture4_Click"
End If
End Sub

Private Sub Picture5_Click()
If TuRn = False Then
GetWinner.Enabled = False
MsgBox "Wait for " & GetNickName & "'s turn to finish!", vbSystemModal + vbCritical, "Impatient"
GetWinner.Enabled = True
Exit Sub
End If
If TuRn = True Then
TuRn = False
Counter = Val(Counter) + 1
Picture5 = XPic
Picture5.Enabled = False
Wsck.SendData "P" & "Picture5_Click"
End If
End Sub

Private Sub Picture6_Click()
If TuRn = False Then
GetWinner.Enabled = False
MsgBox "Wait for " & GetNickName & "'s turn to finish!", vbSystemModal + vbCritical, "Impatient"
GetWinner.Enabled = True
Exit Sub
End If
If TuRn = True Then
TuRn = False
Counter = Val(Counter) + 1
Picture6 = XPic
Picture6.Enabled = False
Wsck.SendData "P" & "Picture6_Click"
End If
End Sub

Private Sub Picture7_Click()
If TuRn = False Then
GetWinner.Enabled = False
MsgBox "Wait for " & GetNickName & "'s turn to finish!", vbSystemModal + vbCritical, "Impatient"
GetWinner.Enabled = True
Exit Sub
End If
If TuRn = True Then
TuRn = False
Counter = Val(Counter) + 1
Picture7 = XPic
Picture7.Enabled = False
Wsck.SendData "P" & "Picture7_Click"
End If
End Sub

Private Sub Picture8_Click()
If TuRn = False Then
GetWinner.Enabled = False
MsgBox "Wait for " & GetNickName & "'s turn to finish!", vbSystemModal + vbCritical, "Impatient"
GetWinner.Enabled = True
Exit Sub
End If
If TuRn = True Then
TuRn = False
Counter = Val(Counter) + 1
Picture8 = XPic
Picture8.Enabled = False
Wsck.SendData "P" & "Picture8_Click"
End If
End Sub

Private Sub Picture9_Click()
If TuRn = False Then
GetWinner.Enabled = False
MsgBox "Wait for " & GetNickName & "'s turn to finish!", vbSystemModal + vbCritical, "Impatient"
GetWinner.Enabled = True
Exit Sub
End If
If TuRn = True Then
TuRn = False
Counter = Val(Counter) + 1
Picture9 = XPic
Picture9.Enabled = False
Wsck.SendData "P" & "Picture9_Click"
End If
End Sub


Private Sub SendTextBox_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyShift Then cSend_Click
End Sub

Private Sub Timer1_Timer()
If frmSelect.Visible = False Then
frmMainClient.Visible = True
Timer1.Enabled = False
End If
End Sub

Private Sub GetWinner_Timer()
If Picture1.Enabled = False And Picture2.Enabled = False And Picture3.Enabled = False Then
GoTo YouWon
Exit Sub
End If
If Picture1.Enabled = False And Picture5.Enabled = False And Picture9.Enabled = False Then
GoTo YouWon
Exit Sub
End If
If Picture1.Enabled = False And Picture4.Enabled = False And Picture7.Enabled = False Then
GoTo YouWon
Exit Sub
End If
If Picture2.Enabled = False And Picture5.Enabled = False And Picture8.Enabled = False Then
GoTo YouWon
Exit Sub
End If
If Picture3.Enabled = False And Picture5.Enabled = False And Picture7.Enabled = False Then
GoTo YouWon
Exit Sub
End If
If Picture3.Enabled = False And Picture6.Enabled = False And Picture9.Enabled = False Then
GoTo YouWon
Exit Sub
End If
If Picture4.Enabled = False And Picture5.Enabled = False And Picture6.Enabled = False Then
GoTo YouWon
Exit Sub
End If
If Picture7.Enabled = False And Picture8.Enabled = False And Picture9.Enabled = False Then
GoTo YouWon
Exit Sub
End If
If cPicture1.Visible = True And cPicture2.Visible = True And cPicture3.Visible = True Then
GoTo YouWon
Exit Sub
End If
If cPicture1.Visible = True And cPicture5.Visible = True And cPicture9.Visible = True Then
GoTo YouWon
Exit Sub
End If
If cPicture1.Visible = True And cPicture4.Visible = True And cPicture7.Visible = True Then
GoTo YouLost
Exit Sub
End If
If cPicture2.Visible = True And cPicture5.Visible = True And cPicture8.Visible = True Then
GoTo YouLost
Exit Sub
End If
If cPicture3.Visible = True And cPicture5.Visible = True And cPicture7.Visible = True Then
GoTo YouLost
Exit Sub
End If
If cPicture3.Visible = True And cPicture6.Visible = True And cPicture9.Visible = True Then
GoTo YouLost
Exit Sub
End If
If cPicture4.Visible = True And cPicture5.Visible = True And cPicture6.Visible = True Then
GoTo YouLost
Exit Sub
End If
If cPicture7.Visible = True And cPicture8.Visible = True And cPicture9.Visible = True Then
GoTo YouLost
Exit Sub
End If
If Counter = 9 Then
Counter = 0
GoTo GameOver
Exit Sub
End If
Exit Sub
YouWon:
NickScore2.Caption = Val(NickScore2.Caption) + 1
GoTo GameOver
Exit Sub
YouLost:
NickScore1.Caption = Val(NickScore1.Caption) + 1
GoTo GameOver
Exit Sub
GameOver:
cPicture1.Visible = False
cPicture2.Visible = False
cPicture3.Visible = False
cPicture4.Visible = False
cPicture5.Visible = False
cPicture6.Visible = False
cPicture7.Visible = False
cPicture8.Visible = False
cPicture9.Visible = False
Picture1.Visible = True
Picture2.Visible = True
Picture3.Visible = True
Picture4.Visible = True
Picture5.Visible = True
Picture6.Visible = True
Picture7.Visible = True
Picture8.Visible = True
Picture9.Visible = True
Picture1.Enabled = True
Picture1 = ClearPic
Picture2.Enabled = True
Picture2 = ClearPic
Picture3.Enabled = True
Picture3 = ClearPic
Picture4.Enabled = True
Picture4 = ClearPic
Picture5.Enabled = True
Picture5 = ClearPic
Picture6.Enabled = True
Picture6 = ClearPic
Picture7.Enabled = True
Picture7 = ClearPic
Picture8.Enabled = True
Picture8 = ClearPic
Picture9.Enabled = True
Picture9 = ClearPic

End Sub

Private Sub Wsck_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim Data As String
Dim Data2 As String
Wsck.GetData Data, vbString
Data2 = Left(Data, 1)
Data = Mid(Data, 2)
If Data2 = "N" Then
If NickName1.Caption = Data Then Exit Sub
NickName1.Caption = Data & ":"
NickName2.Caption = NickName & ":"
NickScore1.Visible = True
NickScore2.Visible = True
End If
If Data2 = "C" Then
Select Case (Data)
Case "Connection Accepted"
frmSelect.Hide
NickSend.Enabled = True
GetWinner.Enabled = True
Case "Disconnected"
MsgBox GetNickName & " was disconnected from you.", vbSystemModal + vbOKOnly + vbCritical, "Notice"
Wsck.Close
Shell App.Path & "\" & App.EXEName, vbNormalFocus
End
End Select
ElseIf Data2 = "T" Then
On Error Resume Next
MainTextBox.SelStart = Len(MainTextBox.Text)
MainTextBox.SelColor = vbRed
GetNick$ = Mid(Data, InStr(Data, ":"), Len(Data))
GetNick2$ = Mid(Data, InStr(Data, ":") + 1, Len(Data))
GetNick3$ = Replace(Data, GetNick$, "")
MainTextBox.SelBold = True
MainTextBox.SelText = GetNick3$ & ":"
MainTextBox.SelBold = False
MainTextBox.SelText = GetNick2$ & vbCrLf
ElseIf Data2 = "P" Then
Select Case (Data)
Case "Picture1_Click"
cPicture1.Visible = True
Picture1.Visible = False
Counter = Val(Counter) + 1
TuRn = True
Case "Picture2_Click"
cPicture2.Visible = True
Picture2.Visible = False
Counter = Val(Counter) + 1
TuRn = True
Case "Picture3_Click"
cPicture3.Visible = True
Picture3.Visible = False
Counter = Val(Counter) + 1
TuRn = True
Case "Picture4_Click"
cPicture4.Visible = True
Picture4.Visible = False
Counter = Val(Counter) + 1
TuRn = True
Case "Picture5_Click"
cPicture5.Visible = True
Picture5.Visible = False
Counter = Val(Counter) + 1
TuRn = True
Case "Picture6_Click"
cPicture6.Visible = True
Picture6.Visible = False
Counter = Val(Counter) + 1
TuRn = True
Case "Picture7_Click"
cPicture7.Visible = True
Picture7.Visible = False
Counter = Val(Counter) + 1
TuRn = True
Case "Picture8_Click"
cPicture8.Visible = True
Picture8.Visible = False
Counter = Val(Counter) + 1
TuRn = True
Case "Picture9_Click"
cPicture9.Visible = True
Picture9.Visible = False
Counter = Val(Counter) + 1
TuRn = True
End Select
End If
End Sub

