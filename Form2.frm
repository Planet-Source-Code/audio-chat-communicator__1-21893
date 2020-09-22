VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Begin VB.Form Form2 
   BackColor       =   &H8000000D&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Expressions"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   3030
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox list1 
      Height          =   4350
      ItemData        =   "Form2.frx":0000
      Left            =   120
      List            =   "Form2.frx":0002
      TabIndex        =   2
      Top             =   240
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Get Merlin to perform and read "
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   4680
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Stop Merlin"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   5280
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   0
      Width           =   855
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   960
      Top             =   3360
      _cx             =   847
      _cy             =   847
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
Form1.Show

End Sub

Private Sub Command2_Click()
Agent1.Characters("merlin").StopAll


Agent1.Characters("merlin").MoveTo 420, 250
Dim i As Integer
For i = 1 To list1.ListCount - 1
Agent1.Characters("merlin").Speak list1.List(i)
Next i


End Sub

Private Sub Command3_Click()
Agent1.Characters("merlin").StopAll
End Sub

Private Sub Command4_Click()
Agent1.Characters("merlin").StopAll
Agent1.Characters("merlin").MoveTo 420, 250
If list1.ListIndex = -1 Then
On Error Resume Next
For i = 1 To list1.ListCount
Agent1.Characters("merlin").Speak list1.List(i)
Agent1.Characters("merlin").Play list1.List(i)
Next i
Else
On Error Resume Next
idx2 = list1.ListIndex
Agent1.Characters("merlin").Speak list1.List(idx2)
Agent1.Characters("merlin").Play list1.List(idx2)
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Agent1.Characters("merlin").StopAll
Agent1.Characters.Load ("merlin")
Agent1.Characters("merlin").Show
Agent1.Characters("merlin").MoveTo 420, 250
Agent1.Characters("merlin").Speak "You can add an expression to your chat program by Dboule clicking on it."
Agent1.Characters("merlin").Speak "If you would like to see what the epression does, just select it, and then press 'Get Merlin to perform and read' button."
Agent1.Characters("merlin").Speak "To hear and see all of the expressions, do not make any selections, just press the button."
list1.AddItem "Acknowledge"
list1.AddItem "Alert"

list1.AddItem "Announce"

list1.AddItem "Blink"
list1.AddItem "Confused"

list1.AddItem "Congratulate"

list1.AddItem "Congratulate_2"
list1.AddItem "Decline"

list1.AddItem "DoMagic1"
list1.AddItem "DoMagic2"

list1.AddItem "DontRecognize"

list1.AddItem "Explain"

list1.AddItem "GestureDown"

list1.AddItem "GestureLeft"

list1.AddItem "GestureRight"

list1.AddItem "GestureUp"

list1.AddItem "GetAttention"

list1.AddItem "Greet"
list1.AddItem "Hide"
list1.AddItem "Idle1_1"
list1.AddItem "Idle1_2"
list1.AddItem "Idle1_3"

list1.AddItem "LookDown"
list1.AddItem "LookDownBlink"
list1.AddItem "LookLeft"
list1.AddItem "LookLeftBlink"

list1.AddItem "LookRight"
list1.AddItem "LookRightBlink"

list1.AddItem "LookUp"
list1.AddItem "LookUpBlink"

list1.AddItem "MoveDown"

list1.AddItem "MoveLeft"

list1.AddItem "MoveRight"

list1.AddItem "MoveUp"

list1.AddItem "Pleased"


list1.AddItem "Read"
list1.AddItem "ReadContinued"


list1.AddItem "RestPose"
list1.AddItem "Sad"


list1.AddItem "Show"
list1.AddItem "StartListening"

list1.AddItem "StopListening"

list1.AddItem "Suggest"

list1.AddItem "Surprised"

list1.AddItem "Think"

list1.AddItem "Uncertain"

list1.AddItem "Wave"

list1.AddItem "Write"
list1.AddItem "WriteContinued"

Label1.Caption = list1.ListCount


End Sub

Private Sub list1_DblClick()
idx = list1.ListIndex
Form1.Text4.Text = list1.List(idx)
End Sub
