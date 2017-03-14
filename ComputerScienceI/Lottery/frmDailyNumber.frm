VERSION 5.00
Begin VB.Form frmDailyNumber 
   Caption         =   "Daily Number"
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   8430
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   195
      Left            =   8040
      TabIndex        =   19
      Top             =   3240
      Width           =   255
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   735
      Left            =   3360
      TabIndex        =   18
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Frame fraResults 
      Caption         =   "Results"
      Height          =   2415
      Left            =   4320
      TabIndex        =   7
      Top             =   120
      Width           =   3975
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Net Winnings"
         Height          =   255
         Left            =   1320
         TabIndex        =   17
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Money Won"
         Height          =   255
         Left            =   2280
         TabIndex        =   16
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Money Spent"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Number of Wins"
         Height          =   255
         Left            =   2280
         TabIndex        =   14
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Number of Plays"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblNet 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   12
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lblWon 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2280
         TabIndex        =   11
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblSpent 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblWins 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2280
         TabIndex        =   9
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblPlays 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Frame fraDaily 
      Caption         =   "Daily Number"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.TextBox txtPlayer 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   2
         Left            =   2760
         TabIndex        =   3
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtPlayer 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   1
         Left            =   1560
         TabIndex        =   2
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtPlayer 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblWinning 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   2760
         TabIndex        =   6
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblWinning 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   1560
         TabIndex        =   5
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblWinning 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   1440
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmDailyNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Austin Herring
'November 11, 2009
'Lottery Project
'frmDaily Number

Private Sub cmdPlay_Click()

lblPlays = Val(lblPlays) + 1
lblSpent = Val(lblSpent) + 1

Dim w1, w2, w3 As Integer ' winning digits
Dim p1, p2, p3 As Integer ' player's numbers

Randomize ' generates a new random seed

w1 = Int(10 * Rnd) ' rnd : random
w2 = Int(10 * Rnd)
w3 = Int(10 * Rnd)

lblWinning(0) = w1: lblWinning(1) = w2: lblWinning(2) = w3

p1 = Val(txtPlayer(0)): p2 = Val(txtPlayer(1)): p3 = Val(txtPlayer(2))

If p1 = w1 And p2 = w2 And p3 = w3 Then
    llbWins = Val(lblWins) + 1 ' add to numb. of wins
    lblWon = Val(lblWon) + 500
End If

lblNet = Val(lblWon) - Val(lblSpent)


End Sub

Private Sub Command1_Click()
frmPowerBall.Show
End Sub

Private Sub txtPlayer_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn And Index < 2 Then txtPlayer(Index + 1).SetFocus
If KeyAscii = vbKeyReturn And Index = 2 Then cmdPlay_Click
End Sub
