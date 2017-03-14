VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSComCt2.ocx"
Begin VB.Form frmCar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Racing Game"
   ClientHeight    =   10320
   ClientLeft      =   4410
   ClientTop       =   525
   ClientWidth     =   6945
   ForeColor       =   &H00FF00FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   688
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   463
   Begin VB.CommandButton Command1 
      Caption         =   "Pie Chart"
      Height          =   375
      Left            =   5400
      TabIndex        =   111
      Top             =   9720
      Width           =   1095
   End
   Begin VB.Frame Frame5 
      Caption         =   "Multiple Runs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4320
      TabIndex        =   107
      Top             =   3360
      Width           =   2535
      Begin VB.CommandButton cmdRun 
         Caption         =   "Run"
         Height          =   375
         Left            =   600
         TabIndex        =   110
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtRuns 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   108
         Text            =   "10"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "Number of runs:"
         Height          =   375
         Left            =   120
         TabIndex        =   109
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Cumulative Stats"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   71
      Top             =   6840
      Width           =   6735
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear Stats"
         Height          =   375
         Left            =   3840
         TabIndex        =   95
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "%"
         Height          =   255
         Index           =   5
         Left            =   5640
         TabIndex        =   101
         Top             =   2520
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label11 
         Caption         =   "%"
         Height          =   255
         Index           =   4
         Left            =   5640
         TabIndex        =   100
         Top             =   2160
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label11 
         Caption         =   "%"
         Height          =   255
         Index           =   3
         Left            =   5640
         TabIndex        =   99
         Top             =   1800
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label11 
         Caption         =   "%"
         Height          =   255
         Index           =   2
         Left            =   5640
         TabIndex        =   98
         Top             =   1440
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label11 
         Caption         =   "%"
         Height          =   255
         Index           =   1
         Left            =   5640
         TabIndex        =   97
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label11 
         Caption         =   "%"
         Height          =   255
         Index           =   0
         Left            =   5640
         TabIndex        =   96
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "Total races:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   94
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2160
         TabIndex        =   93
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   5
         Left            =   4560
         TabIndex        =   92
         Top             =   2520
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   4
         Left            =   4560
         TabIndex        =   91
         Top             =   2160
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   4560
         TabIndex        =   90
         Top             =   1800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   4560
         TabIndex        =   89
         Top             =   1440
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   4560
         TabIndex        =   88
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   4560
         TabIndex        =   87
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblWins 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   5
         Left            =   2160
         TabIndex        =   86
         Top             =   2520
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblWins 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   4
         Left            =   2160
         TabIndex        =   85
         Top             =   2160
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblWins 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   2160
         TabIndex        =   84
         Top             =   1800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblWins 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   83
         Top             =   1440
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblWins 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   82
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblWins 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   2160
         TabIndex        =   81
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Percent of total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   4200
         TabIndex        =   80
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Number of wins:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   79
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Pink Car:"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   78
         Top             =   2520
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Orange Car:"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   77
         Top             =   2160
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Yellow Car:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   76
         Top             =   1800
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Green Car:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   75
         Top             =   1440
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Blue Car:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   74
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Red Car:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   73
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Car:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   72
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "By Percent"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   4320
      TabIndex        =   48
      Top             =   3360
      Visible         =   0   'False
      Width           =   2535
      Begin VB.CommandButton cmdReset1 
         Caption         =   "Reset"
         Height          =   375
         Left            =   1320
         TabIndex        =   103
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "Move"
         Height          =   375
         Left            =   120
         TabIndex        =   102
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox txtProbability 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1320
         TabIndex        =   65
         Top             =   2400
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtProbability 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1320
         TabIndex        =   63
         Top             =   2040
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtProbability 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1320
         TabIndex        =   61
         Top             =   1680
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtProbability 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1320
         TabIndex        =   59
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtProbability 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   57
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtProbability 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   55
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   106
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "%"
         Height          =   255
         Index           =   5
         Left            =   1920
         TabIndex        =   66
         Top             =   2400
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "%"
         Height          =   255
         Index           =   4
         Left            =   1920
         TabIndex        =   64
         Top             =   2040
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "%"
         Height          =   255
         Index           =   3
         Left            =   1920
         TabIndex        =   62
         Top             =   1680
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "%"
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   60
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "%"
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   58
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "%"
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   56
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Red Car:"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   54
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Blue Car:"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   53
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Green Car:"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   52
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Yellow Car:"
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   51
         Top             =   1680
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Orange Car:"
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   50
         Top             =   2040
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Pink Car:"
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   49
         Top             =   2400
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "By Die"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   4095
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         Height          =   375
         Left            =   2880
         TabIndex        =   105
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton cmdRoll 
         Caption         =   "Roll Die"
         Height          =   375
         Left            =   120
         TabIndex        =   104
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CheckBox chk6 
         Caption         =   "6"
         Height          =   255
         Index           =   5
         Left            =   3600
         TabIndex        =   44
         Top             =   2400
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chk5 
         Caption         =   "5"
         Height          =   255
         Index           =   5
         Left            =   3120
         TabIndex        =   43
         Top             =   2400
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chk4 
         Caption         =   "4"
         Height          =   255
         Index           =   5
         Left            =   2640
         TabIndex        =   42
         Top             =   2400
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chk3 
         Caption         =   "3"
         Height          =   255
         Index           =   5
         Left            =   2160
         TabIndex        =   41
         Top             =   2400
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chk2 
         Caption         =   "2"
         Height          =   255
         Index           =   5
         Left            =   1680
         TabIndex        =   40
         Top             =   2400
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chk1 
         Caption         =   "1"
         Height          =   255
         Index           =   5
         Left            =   1200
         TabIndex        =   39
         Top             =   2400
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chk6 
         Caption         =   "6"
         Height          =   255
         Index           =   4
         Left            =   3600
         TabIndex        =   38
         Top             =   2040
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chk5 
         Caption         =   "5"
         Height          =   255
         Index           =   4
         Left            =   3120
         TabIndex        =   37
         Top             =   2040
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chk4 
         Caption         =   "4"
         Height          =   255
         Index           =   4
         Left            =   2640
         TabIndex        =   36
         Top             =   2040
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chk3 
         Caption         =   "3"
         Height          =   255
         Index           =   4
         Left            =   2160
         TabIndex        =   35
         Top             =   2040
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chk2 
         Caption         =   "2"
         Height          =   255
         Index           =   4
         Left            =   1680
         TabIndex        =   34
         Top             =   2040
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chk1 
         Caption         =   "1"
         Height          =   255
         Index           =   4
         Left            =   1200
         TabIndex        =   33
         Top             =   2040
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chk6 
         Caption         =   "6"
         Height          =   255
         Index           =   3
         Left            =   3600
         TabIndex        =   32
         Top             =   1680
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chk5 
         Caption         =   "5"
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   31
         Top             =   1680
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chk4 
         Caption         =   "4"
         Height          =   255
         Index           =   3
         Left            =   2640
         TabIndex        =   30
         Top             =   1680
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chk3 
         Caption         =   "3"
         Height          =   255
         Index           =   3
         Left            =   2160
         TabIndex        =   29
         Top             =   1680
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chk2 
         Caption         =   "2"
         Height          =   255
         Index           =   3
         Left            =   1680
         TabIndex        =   28
         Top             =   1680
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chk1 
         Caption         =   "1"
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   27
         Top             =   1680
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chk6 
         Caption         =   "6"
         Height          =   255
         Index           =   2
         Left            =   3600
         TabIndex        =   26
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chk5 
         Caption         =   "5"
         Height          =   255
         Index           =   2
         Left            =   3120
         TabIndex        =   25
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chk4 
         Caption         =   "4"
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   24
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chk3 
         Caption         =   "3"
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   23
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chk2 
         Caption         =   "2"
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   22
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chk1 
         Caption         =   "1"
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   21
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chk6 
         Caption         =   "6"
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   19
         Top             =   960
         Width           =   375
      End
      Begin VB.CheckBox chk5 
         Caption         =   "5"
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   18
         Top             =   960
         Width           =   375
      End
      Begin VB.CheckBox chk4 
         Caption         =   "4"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   17
         Top             =   960
         Width           =   375
      End
      Begin VB.CheckBox chk3 
         Caption         =   "3"
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   16
         Top             =   960
         Width           =   375
      End
      Begin VB.CheckBox chk2 
         Caption         =   "2"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   15
         Top             =   960
         Width           =   375
      End
      Begin VB.CheckBox chk1 
         Caption         =   "1"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   14
         Top             =   960
         Width           =   375
      End
      Begin VB.CheckBox chk6 
         Caption         =   "6"
         Height          =   255
         Index           =   0
         Left            =   3600
         TabIndex        =   13
         Top             =   600
         Width           =   375
      End
      Begin VB.CheckBox chk5 
         Caption         =   "5"
         Height          =   255
         Index           =   0
         Left            =   3120
         TabIndex        =   12
         Top             =   600
         Width           =   375
      End
      Begin VB.CheckBox chk4 
         Caption         =   "4"
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   11
         Top             =   600
         Width           =   375
      End
      Begin VB.CheckBox chk3 
         Caption         =   "3"
         Height          =   255
         Index           =   0
         Left            =   2160
         TabIndex        =   10
         Top             =   600
         Width           =   375
      End
      Begin VB.CheckBox chk2 
         Caption         =   "2"
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   9
         Top             =   600
         Width           =   375
      End
      Begin VB.CheckBox chk1 
         Caption         =   "1"
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   8
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   70
         Top             =   180
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Pink Car:"
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   47
         Top             =   2400
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Orange Car:"
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   46
         Top             =   2040
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Yellow Car:"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   45
         Top             =   1680
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Green Car:"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Blue Car:"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Red Car:"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove Car"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5520
         TabIndex        =   69
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton optEnter 
         Caption         =   "Percent"
         Height          =   255
         Left            =   3360
         TabIndex        =   68
         Top             =   1440
         Width           =   855
      End
      Begin VB.OptionButton optDie 
         Caption         =   "Die"
         Height          =   255
         Left            =   2520
         TabIndex        =   67
         Top             =   1440
         Value           =   -1  'True
         Width           =   615
      End
      Begin MSComCtl2.UpDown updSegments 
         Height          =   285
         Left            =   1920
         TabIndex        =   7
         Top             =   1440
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtSegments"
         BuddyDispid     =   196647
         OrigLeft        =   2040
         OrigTop         =   2640
         OrigRight       =   2280
         OrigBottom      =   2895
         Max             =   32
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtSegments 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Text            =   "1"
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add Car"
         Height          =   375
         Left            =   4320
         TabIndex        =   1
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   840
         X2              =   840
         Y1              =   750
         Y2              =   1330
      End
      Begin VB.Label Label3 
         Caption         =   "Race Segments:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Image imgCar 
         Height          =   285
         Index           =   5
         Left            =   0
         Picture         =   "frmCar.frx":0000
         Top             =   2190
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Image imgCar 
         Height          =   285
         Index           =   4
         Left            =   0
         Picture         =   "frmCar.frx":0AA6
         Top             =   1900
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Image imgCar 
         Height          =   285
         Index           =   3
         Left            =   0
         Picture         =   "frmCar.frx":154C
         Top             =   1605
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Image imgCar 
         Height          =   285
         Index           =   2
         Left            =   0
         Picture         =   "frmCar.frx":1FF2
         Top             =   1320
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Image imgLane 
         Height          =   300
         Index           =   5
         Left            =   0
         Picture         =   "frmCar.frx":2A98
         Top             =   2190
         Visible         =   0   'False
         Width           =   6705
      End
      Begin VB.Image imgLane 
         Height          =   300
         Index           =   4
         Left            =   0
         Picture         =   "frmCar.frx":93DA
         Top             =   1900
         Visible         =   0   'False
         Width           =   6705
      End
      Begin VB.Image imgLane 
         Height          =   300
         Index           =   3
         Left            =   0
         Picture         =   "frmCar.frx":FD1C
         Top             =   1610
         Visible         =   0   'False
         Width           =   6705
      End
      Begin VB.Image imgLane 
         Height          =   300
         Index           =   2
         Left            =   0
         Picture         =   "frmCar.frx":1665E
         Top             =   1320
         Visible         =   0   'False
         Width           =   6705
      End
      Begin VB.Image imgCar 
         Height          =   240
         Index           =   1
         Left            =   0
         Picture         =   "frmCar.frx":1CFA0
         Top             =   1080
         Width           =   690
      End
      Begin VB.Image imgCar 
         Height          =   270
         Index           =   0
         Left            =   0
         Picture         =   "frmCar.frx":1D8A2
         Top             =   765
         Width           =   720
      End
      Begin VB.Image imgBackground 
         Height          =   1215
         Left            =   0
         Picture         =   "frmCar.frx":1E304
         Top             =   120
         Width           =   6705
      End
   End
End
Attribute VB_Name = "frmCar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'travel distance is 0 - 5880

'Project essentially finished
'All cars run, all statistics accumulate correctly

'Possible things to do:
'--------------------------
'-Clean up the pie chart, if possible (fill it in); got an error with drawing
'circle one time, unsure if true problem (likely fixed now - probably an error with
'terminate > 2pi, which is prevented now)
'
'-Add an animation for the dice roll
'
'

Private Type color
    r As Integer
    g As Integer
    b As Integer
End Type

Dim lane_count As Integer
Dim seg_count As Integer
Dim distance As Integer
Dim car_count As Integer
Dim over As Boolean
Dim moves(5) As Integer
Dim total As Single
'Dim change As Boolean
Dim probability(5) As Integer
Dim multiple As Boolean
Dim clr As color
Dim pi As Single
Dim check As Boolean

Private Sub chk1_Click(Index As Integer)

If chk1(Index).Value = 1 Then
    For x = 0 To car_count - 1
        If x <> Index Then
            chk1(x).Value = 0
        End If
    Next x
End If

End Sub

Private Sub chk2_Click(Index As Integer)

If chk2(Index).Value = 1 Then
    For x = 0 To car_count - 1
        If x <> Index Then
            chk2(x).Value = 0
        End If
    Next x
End If

End Sub

Private Sub chk3_Click(Index As Integer)

If chk3(Index).Value = 1 Then
    For x = 0 To car_count - 1
        If x <> Index Then
            chk3(x).Value = 0
        End If
    Next x
End If

End Sub

Private Sub chk4_Click(Index As Integer)

If chk4(Index).Value = 1 Then
    For x = 0 To car_count - 1
        If x <> Index Then
            chk4(x).Value = 0
        End If
    Next x
End If

End Sub

Private Sub chk5_Click(Index As Integer)

If chk5(Index).Value = 1 Then
    For x = 0 To car_count - 1
        If x <> Index Then
            chk5(x).Value = 0
        End If
    Next x
End If

End Sub

Private Sub chk6_Click(Index As Integer)

If chk6(Index).Value = 1 Then
    For x = 0 To car_count - 1
        If x <> Index Then
            chk6(x).Value = 0
        End If
    Next x
End If

End Sub

Private Sub cmdAdd_Click()

lane_count = lane_count + 1
car_count = car_count + 1


imgLane(lane_count - 1).Visible = True
imgCar(car_count - 1).Visible = True

For x = 0 To seg_count - 1
    Line1(x).Y2 = Line1(x).Y2 + imgLane(2).Height
Next x

Label1(car_count - 1).Visible = True
chk1(car_count - 1).Visible = True
chk2(car_count - 1).Visible = True
chk3(car_count - 1).Visible = True
chk4(car_count - 1).Visible = True
chk5(car_count - 1).Visible = True
chk6(car_count - 1).Visible = True

Label4(car_count - 1).Visible = True
txtProbability(car_count - 1).Visible = True
Label2(car_count - 1).Visible = True

Label7(car_count - 1).Visible = True
lblWins(car_count - 1).Visible = True
lblPercent(car_count - 1).Visible = True
Label11(car_count - 1).Visible = True


cmdAdd.Top = cmdAdd.Top + imgLane(2).Height
Label3.Top = cmdAdd.Top
txtSegments.Top = cmdAdd.Top
updSegments.Top = cmdAdd.Top
optDie.Top = cmdAdd.Top
optEnter.Top = cmdAdd.Top
cmdRemove.Top = cmdAdd.Top

If lane_count = 6 Then
    cmdAdd.Enabled = False
End If

If lane_count >= 2 Then
    cmdRemove.Enabled = True
End If

End Sub

Private Sub cmdClear_Click()

For x = 0 To 5
    lblWins(x) = ""
    lblPercent(x) = ""
    lblTotal = ""
    total = 0
Next x

End Sub

Private Sub cmdMove_Click()

Dim x As Integer
Dim random As Integer

check = False

probability(0) = Val(txtProbability(0))

For x = 1 To car_count - 1
    probability(x) = Val(txtProbability(x)) + probability(x - 1)
Next x

Randomize

random = Int(Rnd * 100) + 1

If random <= probability(0) Then
    If multiple = False Then
        imgCar(0).Left = imgCar(0).Left + distance / seg_count
    End If
    moves(0) = moves(0) + 1
    check = True
Else
    For x = 1 To car_count - 1
        If random > probability(x - 1) And random <= probability(x) Then
            If multiple = False Then
                imgCar(x).Left = imgCar(x).Left + distance / seg_count
            End If
            moves(x) = moves(x) + 1
            check = True
        End If
    Next x
End If
    
For x = 0 To car_count - 1
    If moves(x) = seg_count Then
        cmdMove.Enabled = False
        over = True
        total = total + 1
        lblTotal = total
        lblWins(x) = Val(lblWins(x)) + 1
        
        If x = 0 Then
            Label12 = "Red wins!"
            Label12.ForeColor = &HFF&
        End If
        
        If x = 1 Then
            Label12 = "Blue wins!"
            Label12.ForeColor = &HFF0000
        End If
        
        If x = 2 Then
            Label12 = "Green wins!"
            Label12.ForeColor = &HFF00&
        End If
        
        If x = 3 Then
            Label12 = "Yellow wins!"
            Label12.ForeColor = &HFFFF&
        End If
        
        If x = 4 Then
            Label12 = "Orange wins!"
            Label12.ForeColor = &H80FF&
        End If

        If x = 5 Then
            Label12 = "Pink wins!"
            Label12.ForeColor = &HFF00FF
        End If
    End If
Next x

If over = True Then

    For x = 0 To car_count - 1
        lblPercent(x) = Val(lblWins(x)) / total
        lblPercent(x) = Int(lblPercent(x) * 10000 + 0.5) / 100
    Next x
    
End If

End Sub

Private Sub cmdRemove_Click()

imgLane(lane_count - 1).Visible = False
imgCar(car_count - 1).Visible = False

For x = 0 To seg_count - 1
    Line1(x).Y2 = Line1(x).Y2 - imgLane(2).Height
Next x

Label1(car_count - 1).Visible = False
chk1(car_count - 1).Visible = False
chk2(car_count - 1).Visible = False
chk3(car_count - 1).Visible = False
chk4(car_count - 1).Visible = False
chk5(car_count - 1).Visible = False
chk6(car_count - 1).Visible = False

Label4(car_count - 1).Visible = False
txtProbability(car_count - 1).Visible = False
Label2(car_count - 1).Visible = False

Label7(car_count - 1).Visible = False
lblWins(car_count - 1).Visible = False
lblPercent(car_count - 1).Visible = False
Label11(car_count - 1).Visible = False

cmdAdd.Top = cmdAdd.Top - imgLane(2).Height
Label3.Top = cmdAdd.Top
txtSegments.Top = cmdAdd.Top
updSegments.Top = cmdAdd.Top
optDie.Top = cmdAdd.Top
optEnter.Top = cmdAdd.Top
cmdRemove.Top = cmdAdd.Top

lane_count = lane_count - 1
car_count = car_count - 1

If lane_count <= 5 Then
    cmdAdd.Enabled = True
End If

If lane_count = 2 Then
    cmdRemove.Enabled = False
End If

End Sub

Private Sub cmdReset_Click()

over = False

For x = 0 To car_count - 1
    imgCar(x).Left = 0
    moves(x) = 0
Next x

Label5 = ""
Label12 = ""


cmdRoll.Enabled = True
cmdMove.Enabled = True

End Sub

Private Sub cmdReset1_Click()
cmdReset_Click
End Sub

Private Sub cmdRoll_Click()

Dim random As Integer
Dim x As Integer

check = False

Randomize

random = Int(Rnd * 6) + 1

If random = 1 Then
    For x = 0 To (car_count - 1)
        If chk1(x).Value = 1 Then
            If multiple = False Then
                imgCar(x).Left = imgCar(x).Left + distance / seg_count
            End If
            moves(x) = moves(x) + 1
            check = True
        End If
    Next x
End If

If random = 2 Then
    For x = 0 To (car_count - 1)
        If chk2(x).Value = 1 Then
            If multiple = False Then
                imgCar(x).Left = imgCar(x).Left + distance / seg_count
            End If
            moves(x) = moves(x) + 1
            check = True
        End If
    Next x
End If

If random = 3 Then
    For x = 0 To (car_count - 1)
        If chk3(x).Value = 1 Then
            If multiple = False Then
                imgCar(x).Left = imgCar(x).Left + distance / seg_count
            End If
            moves(x) = moves(x) + 1
            check = True
        End If
    Next x
End If

If random = 4 Then
    For x = 0 To (car_count - 1)
        If chk4(x).Value = 1 Then
            If multiple = False Then
                imgCar(x).Left = imgCar(x).Left + distance / seg_count
            End If
            moves(x) = moves(x) + 1
            check = True
        End If
    Next x
End If

If random = 5 Then
    For x = 0 To (car_count - 1)
        If chk5(x).Value = 1 Then
            If multiple = False Then
                imgCar(x).Left = imgCar(x).Left + distance / seg_count
            End If
            moves(x) = moves(x) + 1
            check = True
        End If
    Next x
End If

If random = 6 Then
    For x = 0 To (car_count - 1)
        If chk6(x).Value = 1 Then
            If multiple = False Then
                imgCar(x).Left = imgCar(x).Left + distance / seg_count
            End If
            moves(x) = moves(x) + 1
            check = True
        End If
    Next x
End If

For x = 0 To car_count - 1
    If moves(x) = seg_count Then
        cmdRoll.Enabled = False
        over = True
        total = total + 1
        lblTotal = total
        lblWins(x) = Val(lblWins(x)) + 1
        
        If x = 0 Then
            Label5 = "Red wins!"
            Label5.ForeColor = &HFF&
        End If
        
        If x = 1 Then
            Label5 = "Blue wins!"
            Label5.ForeColor = &HFF0000
        End If
        
        If x = 2 Then
            Label5 = "Green wins!"
            Label5.ForeColor = &HFF00&
        End If
        
        If x = 3 Then
            Label5 = "Yellow wins!"
            Label5.ForeColor = &HFFFF&
        End If
        
        If x = 4 Then
            Label5 = "Orange wins!"
            Label5.ForeColor = &H80FF&
        End If

        If x = 5 Then
            Label5 = "Pink wins!"
            Label5.ForeColor = &HFF00FF
        End If
        
    End If
Next x

If over = True Then
    For x = 0 To car_count - 1
        lblPercent(x) = Val(lblWins(x)) / total
        lblPercent(x) = Int(lblPercent(x) * 10000 + 0.5) / 100
    Next x
End If


End Sub

Private Sub cmdRun_Click()

Dim x As Integer

multiple = True

If optDie = True Then
    For x = 1 To txtRuns
        Do
            cmdRoll_Click
        Loop Until over = True Or check = False
        cmdReset_Click
        If check = False Then
            frmError.Show
            frmError.Label1 = "Error: unchecked boxes."
        End If
    Next x
End If

If optEnter = True Then
    For x = 1 To txtRuns
        Do
            cmdMove_Click
        Loop Until over = True Or check = False
        cmdReset_Click
        If check = False Then
            frmError.Show
            frmError.Label1 = "Error: unentered percentages."
        End If
    Next x
End If

multiple = False

End Sub

Private Sub Command1_Click()

frmChart.Show

Dim percent(5) As Single
Dim begin, terminate As Single
Dim x As Integer

For x = 0 To car_count - 1
    percent(x) = Val(lblPercent(x))
Next x

begin = 0
terminate = 0

frmChart.Picture1.Cls

For x = 0 To car_count - 1
    If x = 0 Then
        clr.r = 255
        clr.g = 0
        clr.b = 0
    End If

    If x = 1 Then
        clr.r = 0
        clr.g = 0
        clr.b = 255
    End If
    
    If x = 2 Then
        clr.r = 0
        clr.g = 255
        clr.b = 0
    End If
    
    If x = 3 Then
        clr.r = 255
        clr.g = 255
        clr.b = 0
    End If
    
    If x = 4 Then
        clr.r = 241
        clr.g = 103
        clr.b = 1
    End If
    
    If x = 5 Then
        clr.r = 255
        clr.g = 0
        clr.b = 255
    End If
    
    terminate = 2 * pi * (percent(x) / 100) + begin
    
    If terminate > 2 * pi Then
        terminate = 2 * pi
    End If
    
    If percent(x) <> 0 Then
        frmChart.Picture1.Circle (125, 125), 124, RGB(clr.r, clr.g, clr.b), begin, -terminate, 1
    End If
    
    begin = terminate
    
    
Next x

End Sub

Private Sub Form_Load()
lane_count = 2
seg_count = 1
distance = 5880
car_count = 2
'change = True
pi = 3.141592654

End Sub

Private Sub optDie_Click()
Frame2.Visible = True
Frame3.Visible = False
Frame5.Left = Frame3.Left
Frame5.Width = 169
Label13.Left = 120
txtRuns.Left = 1560
cmdRun.Left = 600
End Sub

Private Sub optEnter_Click()
Frame2.Visible = False
Frame3.Visible = True
Frame5.Left = Frame2.Left
Frame5.Width = 273
Label13.Left = 720
txtRuns.Left = 2280
cmdRun.Left = 1320
End Sub


Private Sub txtProbability_Change(Index As Integer)
If car_count = 2 Then
    If Index = 0 Then
        txtProbability(1) = 100 - Val(txtProbability(0))
    End If
    
    If Index = 1 Then
        txtProbability(0) = 100 - Val(txtProbability(1))
    End If
End If

If car_count = 3 Then
    If Index = 0 Or Index = 1 Then
        txtProbability(2) = 100 - Val(txtProbability(0)) - Val(txtProbability(1))
    End If
End If

If car_count = 4 Then
    If Index = 0 Or Index = 1 Or Index = 2 Then
        txtProbability(3) = 100 - Val(txtProbability(0)) - Val(txtProbability(1)) - Val(txtProbability(2))
    End If
End If

If car_count = 5 Then
    If Index = 0 Or Index = 1 Or Index = 2 Or Index = 3 Then
        txtProbability(4) = 100 - Val(txtProbability(0)) - Val(txtProbability(1)) - Val(txtProbability(2)) - Val(txtProbability(3))
    End If
End If

If car_count = 6 Then
    If Index = 0 Or Index = 1 Or Index = 2 Or Index = 3 Or Index = 4 Then
        txtProbability(5) = 100 - Val(txtProbability(0)) - Val(txtProbability(1)) - Val(txtProbability(2)) - Val(txtProbability(3)) - Val(txtProbability(4))
    End If
End If




End Sub

Private Sub txtProbability_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
    If Index = car_count - 1 Then
        txtProbability(0).SetFocus
    Else
        txtProbability(Index + 1).SetFocus
    End If
    
End If
End Sub

Private Sub txtSegments_Change()

If seg_count > 1 Then
    For x = 1 To seg_count - 1
        Unload Line1(x)
    Next x
End If

seg_count = Val(txtSegments)

For x = 1 To seg_count - 1
    Load Line1(x)
    Line1(x).Visible = True
    Line1(x).ZOrder (Front)
    
    Line1(x).X1 = Line1(x - 1).X1 + distance / seg_count
    Line1(x).X2 = Line1(x).X1
Next x

End Sub
