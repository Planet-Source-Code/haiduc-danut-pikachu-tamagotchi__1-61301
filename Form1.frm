VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tamagotchi Pikachu"
   ClientHeight    =   1800
   ClientLeft      =   2190
   ClientTop       =   2565
   ClientWidth     =   3240
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   3240
   Begin VB.CommandButton Command3 
      Caption         =   "Dry"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   1200
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2880
      Top             =   240
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   10
      Scrolling       =   1
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      _Version        =   393216
      BackStyle       =   1
      FullWidth       =   33
      FullHeight      =   33
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Wash"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Give Apple"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Timer eat 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   4560
      Top             =   720
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   10
      Scrolling       =   1
   End
   Begin VB.Label Label2 
      Caption         =   "Cleanness"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Hunger"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo 1
Animation1.Open App.Path & "/eating.avi"
Animation1.Play 1
ProgressBar1.Value = ProgressBar1.Value + 3
Exit Sub
1
ProgressBar1.Value = 10
End Sub

Private Sub Command2_Click()
On Error GoTo 1
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = True
Animation1.Open App.Path & "/washing.avi"
Animation1.Play 1
ProgressBar2.Value = ProgressBar2.Value + 3
Exit Sub
1
ProgressBar2.Value = 10
End Sub

Private Sub Command3_Click()
Animation1.Open App.Path & "/wait.avi"
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = False
End Sub

Private Sub Form_Load()
Animation1.Open App.Path & "/wait.avi"
ProgressBar1.Value = 10
ProgressBar2.Value = 10
End Sub

Private Sub Timer1_Timer()
If ProgressBar1.Value < 0.05 Then
MsgBox "Your Pikachu has died hungry. Sorry for him."
End
ElseIf ProgressBar2.Value < 0.05 Then
MsgBox "Your Pikachu STINKS! Worms and bacteria attacked him!"
End
Else
ProgressBar1.Value = ProgressBar1.Value - 0.02
ProgressBar2.Value = ProgressBar2.Value - 0.02
End If
End Sub
