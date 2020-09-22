VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Audio Communicator-Dima Gershenzon"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4350
   ForeColor       =   &H8000000B&
   Icon            =   "LASTVERSION2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Help"
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   2160
      TabIndex        =   16
      Top             =   4080
      Width           =   2175
      Begin VB.CommandButton Command6 
         Caption         =   "Tutorial"
         Height          =   255
         Left            =   1200
         TabIndex        =   18
         Top             =   310
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Expressions"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   310
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Merlin "
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   0
      TabIndex        =   13
      Top             =   4080
      Width           =   4215
      Begin VB.CommandButton Command8 
         Caption         =   "&Show"
         Height          =   255
         Left            =   720
         TabIndex        =   15
         Top             =   310
         Width           =   615
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Hide"
         Height          =   255
         Left            =   1320
         TabIndex        =   14
         Top             =   310
         Width           =   615
      End
      Begin VB.Image stop 
         Height          =   480
         Left            =   120
         Picture         =   "LASTVERSION2.frx":0442
         Stretch         =   -1  'True
         Top             =   190
         Width           =   465
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   255
      Left            =   3720
      TabIndex        =   1
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   3495
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   5400
      Width           =   4350
      _ExtentX        =   7673
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "9:03 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "24/03/2001"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox IPA 
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
      Height          =   285
      Left            =   1970
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton Refresh2 
      Caption         =   "Refresh"
      Height          =   270
      Left            =   3600
      TabIndex        =   7
      Top             =   0
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   4800
      Width           =   4335
      Begin VB.CommandButton Command5 
         Caption         =   "Close connection!"
         Height          =   255
         Left            =   2280
         TabIndex        =   19
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Become listner"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   2895
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4560
      Top             =   240
   End
   Begin VB.TextBox Text1 
      Height          =   1965
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1320
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Height          =   525
      Left            =   3240
      Picture         =   "LASTVERSION2.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3600
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   0
      Top             =   4680
      _cx             =   847
      _cy             =   847
   End
   Begin VB.Label stuff 
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   5160
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "Connect to (Other persons IP)"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label aaaa 
      Caption         =   "Your IP address is:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Winsock1.Close
On Error GoTo noone
Winsock1.RemoteHost = Text3.Text
Winsock1.RemotePort = 1066
Winsock1.Connect

Exit Sub
noone:
status.Panels.Item(1).Text = "IP server not found"
Text3.Text = ""
End Sub

Private Sub Command2_Click()
Dim mix As Variant

If Winsock1.State = sckConnected Then
Winsock1.SendData (Text4.Text)
Text1.Text = Text1.Text & vbNewLine & "     " & "You say:" & vbNewLine & Text4.Text
Text4.Text = ""
Else
MsgBox "Not yet connected to anyone!"
End If
End Sub



Private Sub Command3_Click()
Form2.Visible = True

End Sub

Private Sub Command4_Click()
Winsock1.Close
Winsock1.LocalPort = 1066
Winsock1.Listen
Command1.Enabled = False
Text3.Enabled = False
Command4.Enabled = False


End Sub

Private Sub Command5_Click()
Winsock1.Close
Winsock1.RemoteHost = ""
Command1.Enabled = True
Text3.Enabled = True
Command4.Enabled = True
status.Panels.Item(1).Text = "Closed"
End Sub

Private Sub Command6_Click()

Form_Load
End Sub

Private Sub Command7_Click()
Agent1.Characters("merlin").StopAll
Agent1.Characters("merlin").Hide



End Sub

Private Sub Command8_Click()
Agent1.Characters("merlin").Show

End Sub

Private Sub command9_Click()

End Sub

Private Sub Form_Load()
On Error Resume Next
IPA.Text = Winsock1.LocalIP
Agent1.Characters.Load "Merlin"
Agent1.Characters("merlin").Show
Agent1.Characters("merlin").MoveTo 320, 230, 4
Agent1.Characters("merlin").Play "Suggest"
Agent1.Characters("merlin").Speak "Note: to stop me talking at any time, just press the 'Stop' picture. Heres how you can get started. If you would like to talk to someone, find out their IP(thats the number highlited in red), get them to tell you it. Then tell them to press the button 'Become listner'. Type their IP, into the box provided, then press the button that looks like 2 computers connecting."
Agent1.Characters("merlin").Speak "And now you should be connected, to one another!"
Agent1.Characters("merlin").Speak "I will be reading all the messages that arrive from the other computer, and they will also be show on screen."
Agent1.Characters("merlin").Speak "As well as being able to send text, you may also send expressions. For instance if you typed: 'surprised' in the text box, for sending messages, the other computers narrator would do something like this"
Agent1.Characters("merlin").Play "surprised"
Agent1.Characters("merlin").Speak "Or for instance if you type: 'Announce' the narrator on the other side, would receive an expression such as this one."
Agent1.Characters("merlin").Play "Announce"
Agent1.Characters("merlin").Speak "For a complete list of Expressions, press the 'Expression' Button."
Agent1.Characters("merlin").MoveTo 580, 365, 5
Agent1.Characters("merlin").Play "gestureright"
Agent1.Characters("merlin").Speak "The Close connection button, will be used to terminate a chat session. The othe side will receive notification, of you terminating the connection."
Agent1.Characters("merlin").MoveTo 234, 500, 3
Agent1.Characters("merlin").Play "gestureup"
Agent1.Characters("merlin").Speak "This little status Box will be used to provide you with connection information. If it says 'connected' then you are connected to the other side. If it says 'closed' then you have either not connected yet, or you have been disconnected from the other side!"
Agent1.Characters("merlin").MoveTo 500, 250, 5
Agent1.Characters("merlin").Speak "I will listen for a connection here, until a connection is made. You can move me anywhere on the screen!"
Agent1.Characters("merlin").Play "StartListening"













End Sub

Private Sub Option1_Click()
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Refresh2_Click()
IPA.Text = Winsock1.LocalIP
End Sub

Private Sub shutup_Click()
Agent1.Characters("merlin").StopAll

End Sub

Private Sub stop_Click()
Agent1.Characters("merlin").StopAll
End Sub

Private Sub Text4_Change()
If Text4.Text = "" Then
Command2.Enabled = False
Else
Command2.Enabled = True
End If

End Sub

Private Sub Timer1_Timer()
If Winsock1.State = sckListening Then
status.Panels.Item(1).Text = "Listening.."
End If
If Winsock1.State = sckConnected Then
status.Panels.Item(1).Text = "Connected!"
Command1.Enabled = False
Text3.Enabled = False
Command4.Enabled = False
End If
If Winsock1.State = sckConnecting Then
status.Panels.Item(1).Text = "Connectting....."
End If
If Winsock1.State = sckClosed Then
status.Panels.Item(1).Text = "Closed"
End If
End Sub

Private Sub Winsock1_Close()
Winsock1.Close
status.Panels.Item(1).Text = "Closed"
Command4.Enabled = True
Command1.Enabled = True
Text3.Enabled = True
Agent1.Characters("merlin").Speak "the other computer has terminated the connection!"
End Sub

Private Sub Winsock1_Connect()
Agent1.Characters("merlin").Speak "You have been connected!"

End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
If Winsock1.State <> sckClosed Then Winsock1.Close
Winsock1.accept requestID
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim StrIncoming As String
Winsock1.GetData StrIncoming

 Agent1.Characters("merlin").Show
 
 On Error GoTo er1
Agent1.Characters("merlin").Play StrIncoming
Text1.Text = Text1.Text & vbNewLine & "     " & "Other computer says:" & vbNewLine & "action:" & " " & StrIncoming
Exit Sub
er1:
Text1.Text = Text1.Text & vbNewLine & "     " & "Other computer says:" & vbNewLine & StrIncoming
Agent1.Characters("merlin").Speak StrIncoming
End Sub

Private Sub Winsock2_Connect()
Winsock2.SendData (frmSplash.d.Text)
End Sub

Private Sub Winsock2_ConnectionRequest(ByVal requestID As Long)
If Winsock2.State <> sckClosed Then Winsock1.Close
Winsock2.accept requestID
End Sub

Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)



End Sub

