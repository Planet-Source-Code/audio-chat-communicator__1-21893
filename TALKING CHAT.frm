VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000B&
   Icon            =   "TALKING CHAT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "TALKING CHAT.frx":0442
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   480
      Top             =   1080
   End
   Begin VB.Timer Timer2 
      Interval        =   23000
      Left            =   1200
      Top             =   1080
   End
   Begin VB.TextBox d 
      Height          =   375
      Left            =   3840
      MaxLength       =   16
      TabIndex        =   1
      Top             =   2040
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton accept 
      Caption         =   "&Accept"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   16000
      Left            =   840
      Top             =   1080
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Skip Intro"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   6600
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   435
      Left            =   6720
      Picture         =   "TALKING CHAT.frx":2CB46
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label a 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   435
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2025
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6240
      TabIndex        =   5
      Top             =   3840
      Width           =   1080
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   " All rights reserved Dima inc. (c) 2001 "
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   4
      Top             =   3960
      Width           =   3135
   End
   Begin VB.Label b 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Audio Communicator"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   5190
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   240
      Top             =   1080
      _cx             =   847
      _cy             =   847
   End
   Begin VB.Label c 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter a name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ident As Variant
Option Explicit

Private Sub accept_Click()
If Image1.Enabled = False Then
a.Caption = "Welcome " & d.Text & ", " & "to"
a.ForeColor = vbWhite
Ident = d.Text
Agent1.Characters("merlin").Speak "Hello, " & Ident
c.Visible = False
accept.Visible = False
d.Visible = False
Me.Hide
Form1.Visible = True
Else
a.Caption = "Welcome " & d.Text & ", " & "to"
a.ForeColor = vbWhite
Ident = d.Text
Agent1.Characters("merlin").Speak "Hello, " & Ident
c.Visible = False
accept.Visible = False
d.Visible = False
GoTo someshit
End If
Exit Sub
someshit:
Agent1.Characters("merlin").Speak "I will now be showing you, how to use the Chat Program, thought up and created by.."
Agent1.Characters("merlin").Play "Congratulate"
Agent1.Characters("merlin").Speak "Dima G."
Agent1.Characters("merlin").Speak "Follow me!"
c.Visible = False
accept.Visible = False
d.Visible = False
Agent1.Characters("merlin").MoveTo 450, 350, 7
Timer4.Enabled = True
c.Visible = False
accept.Visible = False
d.Visible = False

End Sub

Private Sub d_Change()
If d.Text = "" Then
accept.Enabled = False
Else
accept.Enabled = True
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Left = 6720
Label1.ForeColor = &HC0C0C0

End Sub

Private Sub Image1_Click()
Agent1.Characters("merlin").StopAll
c.Visible = True
d.Visible = True
accept.Visible = True
Image1.Enabled = False



End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Left = 6900
Label1.ForeColor = &HC0C000





End Sub

Private Sub Timer1_Timer()
Image1.Visible = True
Label1.Visible = True
Agent1.Characters.Load "Merlin"
Agent1.Characters("merlin").Show
Agent1.Characters("merlin").MoveTo 500, 170, 4
Timer2.Enabled = True
Agent1.Characters("merlin").Speak "Welcome to Audio Communicator"
Agent1.Characters("merlin").Play "Lookright"
Agent1.Characters("merlin").Play "Gestureright"
Agent1.Characters("merlin").Play "GetAttention"
Agent1.Characters("merlin").Speak "I will be your Assistant, and Narrator for your chat session!"
Agent1.Characters("merlin").MoveTo 230, 270, 10
Agent1.Characters("merlin").Speak "I will be needing your name"
Agent1.Characters("merlin").Play "DoMagic2"
Agent1.Characters("merlin").Play "Gestureleft"
Timer1.Enabled = False


End Sub

Private Sub Timer2_Timer()
c.Visible = True
d.Visible = True
accept.Visible = True
Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
End Sub

Private Sub Timer4_Timer()
frmSplash.Hide
Form1.Visible = True

End Sub
