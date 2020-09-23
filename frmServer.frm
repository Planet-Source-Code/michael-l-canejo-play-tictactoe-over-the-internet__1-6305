VERSION 5.00
Begin VB.Form frmSelect 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TicTacToe ßy: MiKE 3D"
   ClientHeight    =   1890
   ClientLeft      =   4170
   ClientTop       =   5535
   ClientWidth     =   3120
   ForeColor       =   &H00000000&
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1890
   ScaleWidth      =   3120
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Height          =   375
      Left            =   400
      TabIndex        =   4
      Text            =   "Guest"
      Top             =   630
      Width           =   2295
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   120
      Top             =   720
   End
   Begin VB.CommandButton cConnect 
      Caption         =   "Connect!"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Height          =   375
      Left            =   400
      TabIndex        =   1
      Text            =   "Client's IP Here"
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Made by:  MiKE 3D"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   1560
      Width           =   1575
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cConnect_Click()
On Error Resume Next
If Text1 = "" Then
MsgBox "Please enter the Client's IP to connect to!", vbSystemModal + vbOKOnly + vbInformation, "Error"
Exit Sub
ElseIf Text2 = "" Then
MsgBox "Please enter a Name to play as", vbSystemModal + vbOKOnly + vbInformation, "Error"
Exit Sub
End If
frmMainClient.Wsck.RemoteHost = Text1
frmMainClient.Wsck.SendData "C" & "Connection Requested"
NickName = Text2
Exit Sub
End Sub

Private Sub cmdO_Click()

End Sub

Private Sub cmdX_Click()

End Sub


Private Sub Option1_Click()

End Sub

Private Sub Option2_Click()

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
StayOnTop Me
CenterForm Me
With frmMainClient.Wsck
     .Protocol = sckUDPProtocol
     .RemotePort = 11111
     .Bind
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmMainClient.Wsck.Close
End
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text1 = frmMainClient.Wsck.LocalIP
End Sub


Private Sub Timer2_Timer()
Timer2.Enabled = False
frmMainClient.Show
frmMainClient.Hide
Timer2.Enabled = False
End Sub
