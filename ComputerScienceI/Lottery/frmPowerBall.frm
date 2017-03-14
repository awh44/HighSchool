VERSION 5.00
Begin VB.Form frmPowerBall 
   Caption         =   "Power Ball"
   ClientHeight    =   9930
   ClientLeft      =   2310
   ClientTop       =   1875
   ClientWidth     =   11085
   LinkTopic       =   "Form1"
   ScaleHeight     =   9930
   ScaleWidth      =   11085
   Begin VB.CommandButton Command3 
      Caption         =   "Daily Number"
      Height          =   735
      Left            =   8040
      TabIndex        =   71
      Top             =   8640
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Play to Win"
      Height          =   795
      Left            =   6240
      TabIndex        =   70
      Top             =   8640
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play FOREVER"
      Height          =   795
      Left            =   5040
      TabIndex        =   69
      Top             =   8640
      Width           =   975
   End
   Begin VB.CommandButton cmdPlay4 
      Caption         =   "Play x 10000"
      Height          =   795
      Left            =   3840
      TabIndex        =   68
      Top             =   8640
      Width           =   975
   End
   Begin VB.CommandButton cmdPlay3 
      Caption         =   "Play x 1000"
      Height          =   795
      Left            =   6240
      TabIndex        =   57
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton cmdPlay2 
      Caption         =   "Play x 100"
      Height          =   795
      Left            =   5040
      TabIndex        =   56
      Top             =   7680
      Width           =   975
   End
   Begin VB.Frame fraMoney 
      Caption         =   "Money"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   240
      TabIndex        =   45
      Top             =   7560
      Width           =   3375
      Begin VB.Label Label21 
         Caption         =   "Net Winnings"
         Height          =   255
         Left            =   1920
         TabIndex        =   53
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblNet 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1920
         TabIndex        =   52
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label20 
         Caption         =   "Money Won"
         Height          =   255
         Left            =   240
         TabIndex        =   51
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblWon 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   240
         TabIndex        =   50
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label19 
         Caption         =   "Money Spent"
         Height          =   255
         Left            =   1920
         TabIndex        =   49
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblSpent 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1920
         TabIndex        =   48
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "Number of Plays"
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblPlays 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame fraResults 
      Caption         =   "Results"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   240
      TabIndex        =   18
      Top             =   3000
      Width           =   7335
      Begin VB.Label lblMoneyPB5 
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
         Left            =   5640
         TabIndex        =   67
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label lblMoneyPB4 
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
         Left            =   4560
         TabIndex        =   66
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label lblMoneyPB3 
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
         Left            =   3480
         TabIndex        =   65
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label lblMoneyPB2 
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
         Left            =   2400
         TabIndex        =   64
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label lblMoneyPB1 
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
         Left            =   1320
         TabIndex        =   63
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label lblMoneyPB 
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
         Left            =   240
         TabIndex        =   62
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label lblMoney5 
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
         Left            =   5640
         TabIndex        =   61
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblMoney4 
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
         Left            =   4560
         TabIndex        =   60
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblMoney3 
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
         Left            =   3480
         TabIndex        =   59
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "PB + 5"
         Height          =   255
         Left            =   5640
         TabIndex        =   44
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label lblPB5 
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
         Left            =   5640
         TabIndex        =   43
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "PB + 4"
         Height          =   255
         Left            =   4560
         TabIndex        =   42
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label lblPB4 
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
         Left            =   4560
         TabIndex        =   41
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "PB + 3"
         Height          =   255
         Left            =   3480
         TabIndex        =   40
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label lblPB3 
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
         Left            =   3480
         TabIndex        =   39
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "PB + 2"
         Height          =   255
         Left            =   2400
         TabIndex        =   38
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label lblPB2 
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
         Left            =   2400
         TabIndex        =   37
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "PB + 1"
         Height          =   255
         Left            =   1320
         TabIndex        =   36
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label lblPB1 
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
         Left            =   1320
         TabIndex        =   35
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "PB"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Power Ball Matches"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label11 
         Caption         =   "Player Matches"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblPBMatch 
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
         Left            =   240
         TabIndex        =   31
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "5"
         Height          =   255
         Left            =   5640
         TabIndex        =   30
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "4"
         Height          =   255
         Left            =   4560
         TabIndex        =   29
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "3"
         Height          =   255
         Left            =   3480
         TabIndex        =   28
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Left            =   2400
         TabIndex        =   27
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "1"
         Height          =   255
         Left            =   1320
         TabIndex        =   26
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lbl5 
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
         Left            =   5640
         TabIndex        =   24
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lbl4 
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
         Left            =   4560
         TabIndex        =   23
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lbl3 
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
         Left            =   3480
         TabIndex        =   22
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lbl2 
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
         Left            =   2400
         TabIndex        =   21
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lbl1 
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
         Left            =   1320
         TabIndex        =   20
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lbl0 
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
         Left            =   240
         TabIndex        =   19
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   795
      Left            =   3840
      TabIndex        =   15
      Top             =   7680
      Width           =   975
   End
   Begin VB.Frame fraPowerball 
      Caption         =   "Power Ball"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7335
      Begin VB.TextBox txtPB 
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
         Left            =   5640
         TabIndex        =   6
         Top             =   840
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
         Index           =   4
         Left            =   4560
         TabIndex        =   5
         Top             =   840
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
         Index           =   3
         Left            =   3480
         TabIndex        =   4
         Top             =   840
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
         Index           =   2
         Left            =   2400
         TabIndex        =   3
         Top             =   840
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
         Left            =   1320
         TabIndex        =   2
         Top             =   840
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
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblPB 
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
         Left            =   5640
         TabIndex        =   14
         Top             =   1800
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
         Index           =   4
         Left            =   4560
         TabIndex        =   13
         Top             =   1800
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
         Index           =   3
         Left            =   3480
         TabIndex        =   12
         Top             =   1800
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
         Left            =   2400
         TabIndex        =   11
         Top             =   1800
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
         Left            =   1320
         TabIndex        =   10
         Top             =   1800
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
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Power Ball Number, 1 - 39"
         Height          =   495
         Left            =   5640
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Player Numbers, 1 - 59"
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.Line Line1 
         X1              =   5400
         X2              =   5400
         Y1              =   360
         Y2              =   2400
      End
   End
   Begin VB.Label Label23 
      Height          =   255
      Left            =   240
      TabIndex        =   58
      Top             =   8280
      Width           =   1815
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Caption         =   "Power Ball?"
      Height          =   255
      Left            =   9240
      TabIndex        =   55
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label lblYN 
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
      Left            =   9720
      TabIndex        =   54
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblMatches 
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
      Left            =   7920
      TabIndex        =   17
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Number of Matches"
      Height          =   255
      Left            =   7560
      TabIndex        =   16
      Top             =   960
      Width           =   1695
   End
End
Attribute VB_Name = "frmPowerBall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Austin Herring
'November 13, 2009
'Lottery Project
'frmPowerBall
'Powerball payouts:
'PB = $3
'Match 1 + PB = $4
'Match 2 + PB = $7
'Match 3 = $7
'Match 3 + PB = $100
'Match 4 = $100
'Match 4 + PB = $10,000
'Match 5 = $200,000
'Match all = jackpot


Private Sub cmdPlay_Click()
Dim w(4) As Integer 'list of winning numbers
Dim p(4) As Integer ' list of player's numbers
Dim player_pb As Integer
Dim winning_pb As Integer

If Val(txtPlayer(0)) = 0 Or Val(txtPlayer(1)) = 0 Or Val(txtPlayer(2)) = 0 Or Val(txtPlayer(3)) = 0 Or Val(txtPlayer(4)) = 0 Or Val(txtPB) = 0 Then
    j = 1
Else

lblMatches = 0

For c = 0 To 4
    p(c) = Val(txtPlayer(c))
Next c

Randomize ' generate new sequence

w(0) = Int(59 * Rnd) + 1
'w(0) = 1

Do                              ' an event controlled do loop
    w(1) = Int(59 * Rnd) + 1
Loop Until w(1) <> w(0)
'w(1) = 2

Do
    w(2) = Int(59 * Rnd) + 1
Loop Until w(2) <> w(0) And w(2) <> w(1)
'w(2) = 3

Do
    w(3) = Int(59 * Rnd) + 1
Loop Until w(3) <> w(0) And w(3) <> w(1) And w(3) <> w(2)
'w(3) = 4

Do
    w(4) = Int(59 * Rnd) + 1
Loop Until w(4) <> w(0) And w(4) <> w(1) And w(4) <> w(2) And w(4) <> w(3)
'w(4) = 5

winning_pb = Int(39 * Rnd) + 1
'winning_pb = 6


lblWinning(0) = w(0)
lblWinning(1) = w(1)
lblWinning(2) = w(2)
lblWinning(3) = w(3)
lblWinning(4) = w(4)
lblPB = winning_pb
player_pb = Val(txtPB)

If w(0) = p(0) Or w(0) = p(1) Or w(0) = p(2) Or w(0) = p(3) Or w(0) = p(4) Then lblMatches = Val(lblMatches) + 1
If w(1) = p(0) Or w(1) = p(1) Or w(1) = p(2) Or w(1) = p(3) Or w(1) = p(4) Then lblMatches = Val(lblMatches) + 1
If w(2) = p(0) Or w(2) = p(1) Or w(2) = p(2) Or w(2) = p(3) Or w(2) = p(4) Then lblMatches = Val(lblMatches) + 1
If w(3) = p(0) Or w(3) = p(1) Or w(3) = p(2) Or w(3) = p(3) Or w(3) = p(4) Then lblMatches = Val(lblMatches) + 1
If w(4) = p(0) Or w(4) = p(1) Or w(4) = p(2) Or w(4) = p(3) Or w(4) = p(4) Then lblMatches = Val(lblMatches) + 1

If winning_pb <> player_pb Then
    lblYN = "No"
    txtPB.BackColor = &H80000005
    lblPB.BackColor = &H8000000F
    If lblMatches = 0 Then lbl0 = Val(lbl0) + 1
    If lblMatches = 1 Then lbl1 = Val(lbl1) + 1
    If lblMatches = 2 Then lbl2 = Val(lbl2) + 1
    If lblMatches = 3 Then
        lbl3 = Val(lbl3) + 1
        lblWon = Val(lblWon) + 7
        lblMoney3 = Val(lblMoney3) + 7
    End If
    If lblMatches = 4 Then
        lbl4 = Val(lbl4) + 1
        lblWon = Val(lblWon) + 100
        lblMoney4 = Val(lblMoney4) + 100
    End If
    If lblMatches = 5 Then
        lbl5 = Val(lbl5) + 1
        lblWon = Val(lblWon) + 200000
        lblMoney5 = Val(lblMoney5) + 200000
    End If
Else
    lblYN = "Yes"
    txtPB.BackColor = RGB(255, 0, 0)
    lblPB.BackColor = RGB(255, 0, 0)
    
    
    If lblMatches = 0 Then
        lblPBMatch = Val(lblPBMatch) + 1
        lblWon = Val(lblWon) + 3
        lblMoneyPB = Val(lblMoneyPB) + 3
    End If
    If lblMatches = 1 Then
        lblPB1 = Val(lblPB1) + 1
        lblWon = Val(lblWon) + 4
        lblMoneyPB1 = Val(lblMoneyPB1) + 4
    End If
    If lblMatches = 2 Then
        lblPB2 = Val(lblPB2) + 1
        lblWon = Val(lblWon) + 7
        lblMoneyPB2 = Val(lblMoneyPB2) + 7
    End If
    If lblMatches = 3 Then
        lblPB3 = Val(lblPB3) + 1
        lblWon = Val(lblWon) + 100
        lblMoneyPB3 = Val(lblMoneyPB3) + 100
    End If
    If lblMatches = 4 Then
        lblPB4 = Val(lblPB4) + 1
        lblWon = Val(lblWon) + 10000
        lblMoneyPB4 = Val(lblMoneyPB4) + 10000
    End If
    If lblMatches = 5 Then
        lblPB5 = Val(lblPB5) + 1
        lblWon = Val(lblWon) + 200000
        lblMoneyPB5 = Val(lblMoneyPB5) + 200000
    End If
End If

lblPlays = Val(lblPlays) + 1
lblSpent = Val(lblSpent) + 1
lblNet = Val(lblWon) - Val(lblSpent)





If txtPlayer(0) = lblWinning(0) Then
    txtPlayer(0).BackColor = RGB(0, 150, 255)
    lblWinning(0).BackColor = RGB(0, 150, 255)
End If
    
If txtPlayer(0) = lblWinning(1) Then
    txtPlayer(0).BackColor = RGB(0, 150, 255)
    lblWinning(1).BackColor = RGB(0, 150, 255)
End If
    
If txtPlayer(0) = lblWinning(2) Then
    txtPlayer(0).BackColor = RGB(0, 150, 255)
    lblWinning(2).BackColor = RGB(0, 150, 255)
End If

If txtPlayer(0) = lblWinning(3) Then
    txtPlayer(0).BackColor = RGB(0, 150, 255)
    lblWinning(3).BackColor = RGB(0, 150, 255)
End If
    
If txtPlayer(0) = lblWinning(4) Then
    txtPlayer(0).BackColor = RGB(0, 150, 255)
    lblWinning(4).BackColor = RGB(0, 150, 255)
End If



If txtPlayer(1) = lblWinning(0) Then
    txtPlayer(1).BackColor = RGB(0, 150, 255)
    lblWinning(0).BackColor = RGB(0, 150, 255)
End If
    
If txtPlayer(1) = lblWinning(1) Then
    txtPlayer(1).BackColor = RGB(0, 150, 255)
    lblWinning(1).BackColor = RGB(0, 150, 255)
End If
    
If txtPlayer(1) = lblWinning(2) Then
    txtPlayer(1).BackColor = RGB(0, 150, 255)
    lblWinning(2).BackColor = RGB(0, 150, 255)
End If

If txtPlayer(1) = lblWinning(3) Then
    txtPlayer(1).BackColor = RGB(0, 150, 255)
    lblWinning(3).BackColor = RGB(0, 150, 255)
End If
    
If txtPlayer(1) = lblWinning(4) Then
    txtPlayer(1).BackColor = RGB(0, 150, 255)
    lblWinning(4).BackColor = RGB(0, 150, 255)
End If



If txtPlayer(2) = lblWinning(0) Then
    txtPlayer(2).BackColor = RGB(0, 150, 255)
    lblWinning(0).BackColor = RGB(0, 150, 255)
End If
    
If txtPlayer(2) = lblWinning(1) Then
    txtPlayer(2).BackColor = RGB(0, 150, 255)
    lblWinning(1).BackColor = RGB(0, 150, 255)
End If
    
If txtPlayer(2) = lblWinning(2) Then
    txtPlayer(2).BackColor = RGB(0, 150, 255)
    lblWinning(2).BackColor = RGB(0, 150, 255)
End If

If txtPlayer(2) = lblWinning(3) Then
    txtPlayer(2).BackColor = RGB(0, 150, 255)
    lblWinning(3).BackColor = RGB(0, 150, 255)
End If
    
If txtPlayer(2) = lblWinning(4) Then
    txtPlayer(2).BackColor = RGB(0, 150, 255)
    lblWinning(4).BackColor = RGB(0, 150, 255)
End If




If txtPlayer(3) = lblWinning(0) Then
    txtPlayer(3).BackColor = RGB(0, 150, 255)
    lblWinning(0).BackColor = RGB(0, 150, 255)
End If
    
If txtPlayer(3) = lblWinning(1) Then
    txtPlayer(3).BackColor = RGB(0, 150, 255)
    lblWinning(1).BackColor = RGB(0, 150, 255)
End If
    
If txtPlayer(3) = lblWinning(2) Then
    txtPlayer(3).BackColor = RGB(0, 150, 255)
    lblWinning(2).BackColor = RGB(0, 150, 255)
End If

If txtPlayer(3) = lblWinning(3) Then
    txtPlayer(3).BackColor = RGB(0, 150, 255)
    lblWinning(3).BackColor = RGB(0, 150, 255)
End If
    
If txtPlayer(3) = lblWinning(4) Then
    txtPlayer(3).BackColor = RGB(0, 150, 255)
    lblWinning(4).BackColor = RGB(0, 150, 255)
End If





If txtPlayer(4) = lblWinning(0) Then
    txtPlayer(4).BackColor = RGB(0, 150, 255)
    lblWinning(0).BackColor = RGB(0, 150, 255)
End If
    
If txtPlayer(4) = lblWinning(1) Then
    txtPlayer(4).BackColor = RGB(0, 150, 255)
    lblWinning(1).BackColor = RGB(0, 150, 255)
End If
    
If txtPlayer(4) = lblWinning(2) Then
    txtPlayer(4).BackColor = RGB(0, 150, 255)
    lblWinning(2).BackColor = RGB(0, 150, 255)
End If

If txtPlayer(4) = lblWinning(3) Then
    txtPlayer(4).BackColor = RGB(0, 150, 255)
    lblWinning(3).BackColor = RGB(0, 150, 255)
End If
    
If txtPlayer(4) = lblWinning(4) Then
    txtPlayer(4).BackColor = RGB(0, 150, 255)
    lblWinning(4).BackColor = RGB(0, 150, 255)
End If



If txtPlayer(0) <> lblWinning(0) And txtPlayer(0) <> lblWinning(1) And txtPlayer(0) <> lblWinning(2) And txtPlayer(0) <> lblWinning(3) And txtPlayer(0) <> lblWinning(4) Then
    txtPlayer(0).BackColor = &H80000005
End If
    
If txtPlayer(1) <> lblWinning(0) And txtPlayer(1) <> lblWinning(1) And txtPlayer(1) <> lblWinning(2) And txtPlayer(1) <> lblWinning(3) And txtPlayer(1) <> lblWinning(4) Then
    txtPlayer(1).BackColor = &H80000005
End If

If txtPlayer(2) <> lblWinning(0) And txtPlayer(2) <> lblWinning(1) And txtPlayer(2) <> lblWinning(2) And txtPlayer(2) <> lblWinning(3) And txtPlayer(2) <> lblWinning(4) Then
    txtPlayer(2).BackColor = &H80000005
End If
    
If txtPlayer(3) <> lblWinning(0) And txtPlayer(3) <> lblWinning(1) And txtPlayer(3) <> lblWinning(2) And txtPlayer(3) <> lblWinning(3) And txtPlayer(3) <> lblWinning(4) Then
    txtPlayer(3).BackColor = &H80000005
End If
   
If txtPlayer(4) <> lblWinning(0) And txtPlayer(4) <> lblWinning(1) And txtPlayer(4) <> lblWinning(2) And txtPlayer(4) <> lblWinning(3) And txtPlayer(4) <> lblWinning(4) Then
    txtPlayer(4).BackColor = &H80000005
End If



If lblWinning(0) <> txtPlayer(0) And lblWinning(0) <> txtPlayer(1) And lblWinning(0) <> txtPlayer(2) And lblWinning(0) <> txtPlayer(3) And lblWinning(0) <> txtPlayer(4) Then
    lblWinning(0).BackColor = &H80000005
End If
    
If lblWinning(1) <> txtPlayer(0) And lblWinning(1) <> txtPlayer(1) And lblWinning(1) <> txtPlayer(2) And lblWinning(1) <> txtPlayer(3) And lblWinning(1) <> txtPlayer(4) Then
    lblWinning(1).BackColor = &H80000005
End If

If lblWinning(2) <> txtPlayer(0) And lblWinning(2) <> txtPlayer(1) And lblWinning(2) <> txtPlayer(2) And lblWinning(2) <> txtPlayer(3) And lblWinning(2) <> txtPlayer(4) Then
    lblWinning(2).BackColor = &H80000005
End If
    
If lblWinning(3) <> txtPlayer(0) And lblWinning(3) <> txtPlayer(1) And lblWinning(3) <> txtPlayer(2) And lblWinning(3) <> txtPlayer(3) And lblWinning(3) <> txtPlayer(4) Then
    lblWinning(3).BackColor = &H80000005
End If
   
If lblWinning(4) <> txtPlayer(0) And lblWinning(4) <> txtPlayer(1) And lblWinning(4) <> txtPlayer(2) And lblWinning(4) <> txtPlayer(3) And lblWinning(4) <> txtPlayer(4) Then
    lblWinning(4).BackColor = &H80000005
End If
    

End If

End Sub


Private Sub cmdPlay2_Click()

Label23 = 0

Do
cmdPlay_Click
Label23 = Val(Label23) + 1
Loop Until Label23 = 100


End Sub

Private Sub cmdPlay3_Click()
Label23 = 0

Do
cmdPlay_Click
Label23 = Val(Label23) + 1
Loop Until Label23 = 1000

End Sub

Private Sub cmdPlay4_Click()
Label23 = 0

Do
cmdPlay_Click
Label23 = Val(Label23) + 1
Loop Until Label23 = 10000
End Sub

Private Sub Command2_Click()
Do
cmdPlay_Click
Loop Until lblPB5 > 0
End Sub

Private Sub Command3_Click()
frmDailyNumber.Show
End Sub

Private Sub Form_Load()
Label23.Visible = False
End Sub



Private Sub txtPB_Change()
If Val(txtPB) < 1 Then txtPB = ""
If Val(txtPB) > 39 Then txtPB = ""
End Sub

Private Sub txtPB_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdPlay_Click

End Sub


Private Sub txtPlayer_Change(Index As Integer)
If Val(txtPlayer(Index)) < 1 Or Val(txtPlayer(Index)) > 59 Then txtPlayer(Index) = ""


End Sub

Private Sub txtPlayer_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn And Index < 4 Then txtPlayer(Index + 1).SetFocus

If KeyAscii = vbKeyReturn And Index = 4 Then txtPB.SetFocus

If KeyAscii = vbKeyReturn And Index = 1 Then
    If txtPlayer(1) = txtPlayer(0) Then
        txtPlayer(1) = ""
        txtPlayer(1).SetFocus
    Else
        txtPlayer(Index + 1).SetFocus
    End If
End If

If KeyAscii = vbKeyReturn And Index = 2 Then
    For c = 0 To 1
        If txtPlayer(2) = txtPlayer(c) Then
            txtPlayer(2) = ""
            txtPlayer(2).SetFocus
        Else
            txtPlayer(Index + 1).SetFocus
        End If
    Next c
End If

If KeyAscii = vbKeyReturn And Index = 3 Then
    For c = 0 To 2
        If txtPlayer(3) = txtPlayer(c) Then
            txtPlayer(3) = ""
            txtPlayer(3).SetFocus
        Else
            txtPlayer(Index + 1).SetFocus
        End If
    Next c
End If

If KeyAscii = vbKeyReturn And Index = 4 Then
    For c = 0 To 3
        If txtPlayer(4) = txtPlayer(c) Then
            txtPlayer(4) = ""
            txtPlayer(4).SetFocus
        Else
            txtPB.SetFocus
        End If
    Next c
End If
    
End Sub
