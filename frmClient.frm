VERSION 5.00
Begin VB.Form frmSelect 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TicTacToe ßy: MiKE 3D"
   ClientHeight    =   1815
   ClientLeft      =   4170
   ClientTop       =   5535
   ClientWidth     =   3615
   ForeColor       =   &H00000000&
   Icon            =   "frmClient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1815
   ScaleWidth      =   3615
   Begin VB.CommandButton Command2 
      Caption         =   "Disconnect!"
      Height          =   375
      Left            =   1110
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Height          =   375
      Left            =   630
      TabIndex        =   3
      Text            =   "Guest"
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect!"
      Height          =   375
      Left            =   1110
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   720
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Waiting for host to accept connection..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Made by:  MiKE 3D"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   1440
      Width           =   1575
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
On Error Resume Next
If Text1 = "" Then
MsgBox "Please enter a Name to play as", vbSystemModal + vbOKOnly + vbInformation, "Error"
Exit Sub
End If
Me.Height = 2235
Label2.Visible = True
Text1.Top = 480
Command1.Top = 960
Label1.Top = 1440
Command1.Visible = False
Command2.Visible = True
NickName = Text1
Text1.Enabled = False
With frmMainServer.Wsck
.Protocol = sckUDPProtocol
.LocalPort = 11111
.Bind
End With
End Sub
Private Sub Command2_Click()
Command1.Visible = True
Command2.Visible = False
Text1.Top = 120
Command1.Top = 600
Label1.Top = 1080
Me.Height = 1800
Label2.Visible = False
frmMainServer.Wsck.Close
Text1.Enabled = True
End Sub

Private Sub Command3_Click()
MsgBox vbKeyReturn
End Sub

Private Sub Form_Load()
StayOnTop Me
CenterForm Me
frmMainServer.Show
frmMainServer.Hide
Text1.Top = 120
Command1.Top = 600
Label1.Top = 1080
Me.Height = 1800
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmMainServer.Wsck.Close
End
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
If Label2.Caption = "Connection Granted!" Then
Me.Hide
frmMainServer.Show
frmMainServer.NickSend.Enabled = True
frmMainServer.GetWinner.Enabled = True
End If
Timer1.Enabled = False
End Sub
