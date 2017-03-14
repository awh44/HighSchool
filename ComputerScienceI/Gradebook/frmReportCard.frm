VERSION 5.00
Begin VB.Form frmReportCard 
   Caption         =   "Report Card"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   2070
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   7860
   ScaleWidth      =   15240
   Begin VB.Frame fraReportCard 
      Caption         =   "Report Card"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14895
      Begin VB.Timer Timer1 
         Left            =   2880
         Top             =   6720
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   840
         TabIndex        =   124
         Top             =   5880
         Width           =   3135
      End
      Begin VB.CommandButton cmdCompute 
         Caption         =   "Compute"
         Height          =   495
         Left            =   7920
         TabIndex        =   123
         Top             =   6840
         Width           =   1455
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Open"
         Height          =   495
         Left            =   6240
         TabIndex        =   122
         Top             =   6840
         Width           =   1455
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   495
         Left            =   4560
         TabIndex        =   121
         Top             =   6840
         Width           =   1455
      End
      Begin VB.TextBox txtCredit 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   13560
         TabIndex        =   120
         Text            =   "1.17"
         Top             =   5040
         Width           =   975
      End
      Begin VB.TextBox txtCredit 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   13560
         TabIndex        =   119
         Text            =   ".33"
         Top             =   4560
         Width           =   975
      End
      Begin VB.TextBox txtCredit 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   13560
         TabIndex        =   118
         Text            =   ".33"
         Top             =   4080
         Width           =   975
      End
      Begin VB.TextBox txtCredit 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   13560
         TabIndex        =   117
         Text            =   "1"
         Top             =   3600
         Width           =   975
      End
      Begin VB.TextBox txtCredit 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   13560
         TabIndex        =   116
         Text            =   "1"
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox txtCredit 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   13560
         TabIndex        =   115
         Text            =   "1"
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox txtCredit 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   13560
         TabIndex        =   114
         Text            =   "1"
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtCredit 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   13560
         TabIndex        =   113
         Text            =   "1"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtCredit 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   13560
         TabIndex        =   112
         Text            =   "0.17"
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtCredit 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   13560
         TabIndex        =   111
         Text            =   "0.5"
         Top             =   720
         Width           =   975
      End
      Begin VB.ComboBox cmbType 
         Height          =   315
         Index           =   7
         ItemData        =   "frmReportCard.frx":0000
         Left            =   12120
         List            =   "frmReportCard.frx":000D
         TabIndex        =   110
         Text            =   "Honors"
         Top             =   5040
         Width           =   1215
      End
      Begin VB.ComboBox cmbType 
         Height          =   315
         Index           =   8
         ItemData        =   "frmReportCard.frx":0026
         Left            =   12120
         List            =   "frmReportCard.frx":0033
         TabIndex        =   109
         Text            =   "Regular"
         Top             =   4560
         Width           =   1215
      End
      Begin VB.ComboBox cmbType 
         Height          =   315
         Index           =   9
         ItemData        =   "frmReportCard.frx":004C
         Left            =   12120
         List            =   "frmReportCard.frx":0059
         TabIndex        =   108
         Text            =   "Regular"
         Top             =   4080
         Width           =   1215
      End
      Begin VB.ComboBox cmbType 
         Height          =   315
         Index           =   6
         ItemData        =   "frmReportCard.frx":0072
         Left            =   12120
         List            =   "frmReportCard.frx":007F
         TabIndex        =   107
         Text            =   "Honors"
         Top             =   3600
         Width           =   1215
      End
      Begin VB.ComboBox cmbType 
         Height          =   315
         Index           =   5
         ItemData        =   "frmReportCard.frx":0098
         Left            =   12120
         List            =   "frmReportCard.frx":00A5
         TabIndex        =   106
         Text            =   "Regular"
         Top             =   3120
         Width           =   1215
      End
      Begin VB.ComboBox cmbType 
         Height          =   315
         Index           =   4
         ItemData        =   "frmReportCard.frx":00BE
         Left            =   12120
         List            =   "frmReportCard.frx":00CB
         TabIndex        =   105
         Text            =   "Honors"
         Top             =   2640
         Width           =   1215
      End
      Begin VB.ComboBox cmbType 
         Height          =   315
         Index           =   3
         ItemData        =   "frmReportCard.frx":00E4
         Left            =   12120
         List            =   "frmReportCard.frx":00F1
         TabIndex        =   104
         Text            =   "Honors"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.ComboBox cmbType 
         Height          =   315
         Index           =   2
         ItemData        =   "frmReportCard.frx":010A
         Left            =   12120
         List            =   "frmReportCard.frx":0117
         TabIndex        =   103
         Text            =   "Regular"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.ComboBox cmbType 
         Height          =   315
         Index           =   1
         ItemData        =   "frmReportCard.frx":0130
         Left            =   12120
         List            =   "frmReportCard.frx":013D
         TabIndex        =   102
         Text            =   "Regular"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.ComboBox cmbType 
         Height          =   315
         Index           =   0
         ItemData        =   "frmReportCard.frx":0156
         Left            =   12120
         List            =   "frmReportCard.frx":0163
         TabIndex        =   101
         Text            =   "Regular"
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkYear 
         Height          =   195
         Index           =   7
         Left            =   10560
         TabIndex        =   100
         Top             =   5040
         Width           =   255
      End
      Begin VB.CheckBox chkYear 
         Height          =   195
         Index           =   8
         Left            =   10560
         TabIndex        =   99
         Top             =   4560
         Width           =   255
      End
      Begin VB.CheckBox chkYear 
         Height          =   195
         Index           =   9
         Left            =   10560
         TabIndex        =   98
         Top             =   4080
         Width           =   255
      End
      Begin VB.CheckBox chkYear 
         Height          =   195
         Index           =   6
         Left            =   10560
         TabIndex        =   97
         Top             =   3600
         Width           =   255
      End
      Begin VB.CheckBox chkYear 
         Height          =   195
         Index           =   5
         Left            =   10560
         TabIndex        =   96
         Top             =   3120
         Width           =   255
      End
      Begin VB.CheckBox chkYear 
         Height          =   195
         Index           =   4
         Left            =   10560
         TabIndex        =   95
         Top             =   2640
         Width           =   255
      End
      Begin VB.CheckBox chkYear 
         Height          =   195
         Index           =   3
         Left            =   10560
         TabIndex        =   94
         Top             =   2160
         Width           =   255
      End
      Begin VB.CheckBox chkYear 
         Height          =   195
         Index           =   2
         Left            =   10560
         TabIndex        =   93
         Top             =   1680
         Width           =   255
      End
      Begin VB.CheckBox chkYear 
         Height          =   195
         Index           =   1
         Left            =   10560
         TabIndex        =   92
         Top             =   1200
         Width           =   255
      End
      Begin VB.CheckBox chkYear 
         Height          =   195
         Index           =   0
         Left            =   10560
         TabIndex        =   91
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox txtFinal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   9480
         TabIndex        =   80
         Top             =   5040
         Width           =   975
      End
      Begin VB.TextBox txtFinal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   9480
         TabIndex        =   79
         Top             =   4560
         Width           =   975
      End
      Begin VB.TextBox txtFinal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   9480
         TabIndex        =   78
         Top             =   4080
         Width           =   975
      End
      Begin VB.TextBox txtFinal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   9480
         TabIndex        =   77
         Top             =   3600
         Width           =   975
      End
      Begin VB.TextBox txtFinal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   9480
         TabIndex        =   76
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox txtFinal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   9480
         TabIndex        =   75
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox txtFinal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   9480
         TabIndex        =   74
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtFinal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   9480
         TabIndex        =   73
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtFinal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   9480
         TabIndex        =   72
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtFinal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   9480
         TabIndex        =   71
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblGPA 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   11640
         TabIndex        =   140
         Top             =   6360
         Width           =   1815
      End
      Begin VB.Label Label34 
         Caption         =   "Quarter Average"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7800
         TabIndex        =   139
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblQuarters 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   7
         Left            =   8040
         TabIndex        =   138
         Top             =   5040
         Width           =   855
      End
      Begin VB.Label lblQuarters 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   8
         Left            =   8040
         TabIndex        =   137
         Top             =   4560
         Width           =   855
      End
      Begin VB.Label lblQuarters 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   9
         Left            =   8040
         TabIndex        =   136
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label lblQuarters 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   6
         Left            =   8040
         TabIndex        =   135
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label lblQuarters 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   5
         Left            =   8040
         TabIndex        =   134
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label lblQuarters 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   4
         Left            =   8040
         TabIndex        =   133
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label lblQuarters 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   8040
         TabIndex        =   132
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lblQuarters 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   8040
         TabIndex        =   131
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblQuarters 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   8040
         TabIndex        =   130
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblQuarters 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   8040
         TabIndex        =   129
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblAverage 
         Height          =   255
         Left            =   10680
         TabIndex        =   128
         Top             =   5880
         Width           =   1455
      End
      Begin VB.Label Label33 
         Caption         =   "GPA:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11160
         TabIndex        =   127
         Top             =   6360
         Width           =   495
      End
      Begin VB.Label Label32 
         Caption         =   "Average:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9840
         TabIndex        =   126
         Top             =   5880
         Width           =   975
      End
      Begin VB.Label Label31 
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   125
         Top             =   5880
         Width           =   975
      End
      Begin VB.Label lblYear 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   7
         Left            =   10920
         TabIndex        =   90
         Top             =   5040
         Width           =   855
      End
      Begin VB.Label lblYear 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   8
         Left            =   10920
         TabIndex        =   89
         Top             =   4560
         Width           =   855
      End
      Begin VB.Label lblYear 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   9
         Left            =   10920
         TabIndex        =   88
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label lblYear 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   6
         Left            =   10920
         TabIndex        =   87
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label lblYear 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   5
         Left            =   10920
         TabIndex        =   86
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label lblYear 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   4
         Left            =   10920
         TabIndex        =   85
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label lblYear 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   10920
         TabIndex        =   84
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lblYear 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   10920
         TabIndex        =   83
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblYear 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   10920
         TabIndex        =   82
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblYear 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   10920
         TabIndex        =   81
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblQuarter4 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   7
         Left            =   6600
         TabIndex        =   70
         Top             =   5040
         Width           =   855
      End
      Begin VB.Label lblQuarter4 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   8
         Left            =   6600
         TabIndex        =   69
         Top             =   4560
         Width           =   855
      End
      Begin VB.Label lblQuarter4 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   9
         Left            =   6600
         TabIndex        =   68
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label lblQuarter4 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   6
         Left            =   6600
         TabIndex        =   67
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label lblQuarter4 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   5
         Left            =   6600
         TabIndex        =   66
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label lblQuarter4 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   4
         Left            =   6600
         TabIndex        =   65
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label lblQuarter4 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   6600
         TabIndex        =   64
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lblQuarter4 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   6600
         TabIndex        =   63
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblQuarter4 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   6600
         TabIndex        =   62
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblQuarter4 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   6600
         TabIndex        =   61
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblQuarter3 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   7
         Left            =   5400
         TabIndex        =   60
         Top             =   5040
         Width           =   855
      End
      Begin VB.Label lblQuarter3 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   8
         Left            =   5400
         TabIndex        =   59
         Top             =   4560
         Width           =   855
      End
      Begin VB.Label lblQuarter3 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   9
         Left            =   5400
         TabIndex        =   58
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label lblQuarter3 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   6
         Left            =   5400
         TabIndex        =   57
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label lblQuarter3 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   5
         Left            =   5400
         TabIndex        =   56
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label lblQuarter3 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   4
         Left            =   5400
         TabIndex        =   55
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label lblQuarter3 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   5400
         TabIndex        =   54
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lblQuarter3 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   5400
         TabIndex        =   53
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblQuarter3 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   5400
         TabIndex        =   52
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblQuarter3 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   5400
         TabIndex        =   51
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblQuarter2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   7
         Left            =   4200
         TabIndex        =   50
         Top             =   5040
         Width           =   855
      End
      Begin VB.Label lblQuarter1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   7
         Left            =   3000
         TabIndex        =   49
         Top             =   5040
         Width           =   855
      End
      Begin VB.Label Label30 
         Caption         =   "Career Education"
         Height          =   255
         Left            =   1200
         TabIndex        =   48
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         Caption         =   "1"
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblQuarter2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   8
         Left            =   4200
         TabIndex        =   46
         Top             =   4560
         Width           =   855
      End
      Begin VB.Label lblQuarter2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   9
         Left            =   4200
         TabIndex        =   45
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label lblQuarter2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   6
         Left            =   4200
         TabIndex        =   44
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label lblQuarter2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   5
         Left            =   4200
         TabIndex        =   43
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label lblQuarter2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   4
         Left            =   4200
         TabIndex        =   42
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label lblQuarter2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   4200
         TabIndex        =   41
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lblQuarter2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   4200
         TabIndex        =   40
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblQuarter2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   4200
         TabIndex        =   39
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblQuarter2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   4200
         TabIndex        =   38
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblQuarter1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   8
         Left            =   3000
         TabIndex        =   37
         Top             =   4560
         Width           =   855
      End
      Begin VB.Label lblQuarter1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   9
         Left            =   3000
         TabIndex        =   36
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label lblQuarter1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   6
         Left            =   3000
         TabIndex        =   35
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label lblQuarter1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   5
         Left            =   3000
         TabIndex        =   34
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label lblQuarter1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   4
         Left            =   3000
         TabIndex        =   33
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label lblQuarter1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   32
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lblQuarter1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   31
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblQuarter1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   30
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblQuarter1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   29
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label28 
         Caption         =   "Chemistry / Lab"
         Height          =   255
         Left            =   1200
         TabIndex        =   28
         Top             =   5040
         Width           =   1695
      End
      Begin VB.Label Label27 
         Caption         =   "Physical Education"
         Height          =   255
         Left            =   1200
         TabIndex        =   27
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label Label26 
         Caption         =   "Traffic Safety"
         Height          =   255
         Left            =   1200
         TabIndex        =   26
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Label Label25 
         Caption         =   "Algebra 2"
         Height          =   255
         Left            =   1200
         TabIndex        =   25
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label Label24 
         Caption         =   "German 3"
         Height          =   255
         Left            =   1200
         TabIndex        =   24
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Label Label23 
         Caption         =   "US History 1"
         Height          =   255
         Left            =   1200
         TabIndex        =   23
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label22 
         Caption         =   "English"
         Height          =   255
         Left            =   1200
         TabIndex        =   22
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label21 
         Caption         =   "Computer Science 1"
         Height          =   255
         Left            =   1200
         TabIndex        =   21
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label20 
         Caption         =   "Band"
         Height          =   255
         Left            =   1200
         TabIndex        =   20
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "9"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   5040
         Width           =   615
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "8"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   4560
         Width           =   615
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "8"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   4080
         Width           =   615
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "7"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "6"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "4"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "3"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "1"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Credit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   13560
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   12120
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Year to Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10680
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Final Exam"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9480
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Quarter 4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Quarter 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Quarter 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Quarter 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Course"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Period"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmReportCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Austin Herring
'frmReportCard
'Gradebook Project

Private Sub cmdCompute_Click()

For c = 0 To 9

If Len(lblQuarter1(c)) > 0 And Len(lblQuarter2(c)) = 0 And Len(lblQuarter3(c)) = 0 And Len(lblQuarter4(c)) = 0 And Len(txtFinal(c)) = 0 Then
    lblYear(c) = lblQuarter1(c)
    lblQuarters(c) = lblYear(c)
End If
    
If Len(lblQuarter1(c)) > 0 And Len(lblQuarter2(c)) > 0 And Len(lblQuarter3(c)) = 0 And Len(lblQuarter4(c)) = 0 And Len(txtFinal(c)) = 0 Then
    lblYear(c) = (Val(lblQuarter1(c)) + Val(lblQuarter2(c))) / 2
    lblQuarters(c) = lblYear(c)
End If

If Len(lblQuarter1(c)) > 0 And Len(lblQuarter2(c)) > 0 And Len(lblQuarter3(c)) > 0 And Len(lblQuarter4(c)) = 0 And Len(txtFinal(c)) = 0 Then
    lblYear(c) = (Val(lblQuarter1(c)) + Val(lblQuarter2(c)) + Val(lblQuarter3(c))) / 3
    lblQuarters(c) = lblYear(c)
End If

If Len(lblQuarter1(c)) > 0 And Len(lblQuarter2(c)) > 0 And Len(lblQuarter3(c)) > 0 And Len(lblQuarter4(c)) > 0 And Len(txtFinal(c)) = 0 Then
    lblYear(c) = (Val(lblQuarter1(c)) + Val(lblQuarter2(c)) + Val(lblQuarter3(c)) + Val(lblQuarter4(c))) / 4
    lblQuarters(c) = lblYear(c)
End If

If Len(lblQuarter1(c)) > 0 And Len(lblQuarter2(c)) > 0 And Len(lblQuarter3(c)) > 0 And Len(lblQuarter4(c)) > 0 And Len(txtFinal(c)) > 0 Then
    Dim quarters As Single
    quarters = (Val(lblQuarter1(c)) + Val(lblQuarter2(c)) + Val(lblQuarter3(c)) + Val(lblQuarter4(c))) / 4
    lblQuarters(c) = quarters
    lblYear(c) = (quarters * (8 / 9)) + (Val(txtFinal(c)) * (1 / 9))
End If

Next c

For c = 0 To 9
    If Len(lblYear(0)) > 0 And Len(lblYear(1)) > 0 And Len(lblYear(2)) > 0 And Len(lblYear(3)) > 0 And Len(lblYear(4)) > 0 And Len(lblYear(5)) > 0 And Len(lblYear(6)) > 0 And Len(lblYear(7)) > 0 And Len(lblYear(8)) > 0 And Len(lblYear(9)) > 0 Then
        Dim n As Single
        n = n + Val(lblYear(c))
        lblAverage = n / 10
        lblAverage = Int(lblAverage + 0.5)
    End If
Next c

If Len(lblYear(2)) > 0 And Len(lblYear(3)) > 0 And Len(lblYear(4)) > 0 And Len(lblYear(5)) > 0 And Len(lblYear(6)) > 0 And Len(lblYear(7)) > 0 Then

lblGPA = 0

For c = 2 To 7
    If cmbType(c).Text = "Regular" Then lblGPA = Val(lblGPA) + Val(lblYear(c))
    If cmbType(c).Text = "Honors" Then lblGPA = Val(lblGPA) + (Val(lblYear(c)) * 1.05)
    If cmbType(c).Text = "AP" Then lblGPA = Val(lblGPA) + (Val(lblYear(c)) * 1.1)
Next c

lblGPA = Val(lblGPA) / 6
lblGPA = Int(lblGPA * 100 + 0.5) / 100


End If
    

    
End Sub






Private Sub lblQuarter1_Change(Index As Integer)
If lblQuarter1(Index) >= 92 Then
    lblQuarter1(Index).BackColor = RGB(0, 100, 255)
End If

If 83 <= lblQuarter1(Index) And lblQuarter1(Index) < 92 Then
    lblQuarter1(Index).BackColor = RGB(0, 255, 0)
End If

If 74 <= lblQuarter1(Index) And lblQuarter1(Index) < 83 Then
    lblQuarter1(Index).BackColor = RGB(243, 237, 1)
End If

If lblQuarter1(Index) >= 65 And lblQuarter1(Index) < 74 Then
    lblQuarter1(Index).BackColor = RGB(242, 116, 2)
End If

If lblQuarter1(Index) < 65 Then
    lblQuarter1(Index).BackColor = RGB(255, 0, 0)
End If
End Sub

Private Sub lblQuarter2_Change(Index As Integer)
If lblQuarter2(Index) >= 92 Then
    lblQuarter2(Index).BackColor = RGB(0, 100, 255)
End If

If 83 <= lblQuarter2(Index) And lblQuarter2(Index) < 92 Then
    lblQuarter2(Index).BackColor = RGB(0, 255, 0)
End If

If 74 <= lblQuarter2(Index) And lblQuarter2(Index) < 83 Then
    lblQuarter2(Index).BackColor = RGB(243, 237, 1)
End If

If lblQuarter2(Index) >= 65 And lblQuarter2(Index) < 74 Then
    lblQuarter2(Index).BackColor = RGB(242, 116, 2)
End If

If lblQuarter2(Index) < 65 Then
    lblQuarter2(Index).BackColor = RGB(255, 0, 0)
End If
End Sub

Private Sub lblQuarter3_Change(Index As Integer)
If lblQuarter3(Index) >= 92 Then
    lblQuarter3(Index).BackColor = RGB(0, 100, 255)
End If

If 83 <= lblQuarter3(Index) And lblQuarter3(Index) < 92 Then
    lblQuarter3(Index).BackColor = RGB(0, 255, 0)
End If

If 74 <= lblQuarter3(Index) And lblQuarter3(Index) < 83 Then
    lblQuarter3(Index).BackColor = RGB(243, 237, 1)
End If

If lblQuarter3(Index) >= 65 And lblQuarter3(Index) < 74 Then
    lblQuarter3(Index).BackColor = RGB(242, 116, 2)
End If

If lblQuarter3(Index) < 65 Then
    lblQuarter3(Index).BackColor = RGB(255, 0, 0)
End If
End Sub

Private Sub lblQuarter4_Change(Index As Integer)
If lblQuarter4(Index) >= 92 Then
    lblQuarter4(Index).BackColor = RGB(0, 100, 255)
End If

If 83 <= lblQuarter4(Index) And lblQuarter4(Index) < 92 Then
    lblQuarter4(Index).BackColor = RGB(0, 255, 0)
End If

If 74 <= lblQuarter4(Index) And lblQuarter4(Index) < 83 Then
    lblQuarter4(Index).BackColor = RGB(243, 237, 1)
End If

If lblQuarter4(Index) >= 65 And lblQuarter4(Index) < 74 Then
    lblQuarter4(Index).BackColor = RGB(242, 116, 2)
End If

If lblQuarter4(Index) < 65 Then
    lblQuarter4(Index).BackColor = RGB(255, 0, 0)
End If
End Sub

Private Sub lblQuarters_Change(Index As Integer)
If lblQuarters(Index) >= 92 Then
    lblQuarters(Index).BackColor = RGB(0, 100, 255)
End If

If 83 <= lblQuarters(Index) And lblQuarters(Index) < 92 Then
    lblQuarters(Index).BackColor = RGB(0, 255, 0)
End If

If 74 <= lblQuarters(Index) And lblQuarters(Index) < 83 Then
    lblQuarters(Index).BackColor = RGB(243, 237, 1)
End If

If lblQuarters(Index) >= 65 And lblQuarters(Index) < 74 Then
    lblQuarters(Index).BackColor = RGB(242, 116, 2)
End If

If lblQuarters(Index) < 65 Then
    lblQuarters(Index).BackColor = RGB(255, 0, 0)
End If

lblQuarters(Index) = Int(lblQuarters(Index) + 0.5)
End Sub

Private Sub lblYear_Change(Index As Integer)
If lblYear(Index) >= 92 Then
    lblYear(Index).BackColor = RGB(0, 100, 255)
End If

If 83 <= lblYear(Index) And lblYear(Index) < 92 Then
    lblYear(Index).BackColor = RGB(0, 255, 0)
End If

If 74 <= lblYear(Index) And lblYear(Index) < 83 Then
    lblYear(Index).BackColor = RGB(243, 237, 1)
End If

If lblYear(Index) >= 65 And lblYear(Index) < 74 Then
    lblYear(Index).BackColor = RGB(242, 116, 2)
End If

If lblYear(Index) < 65 Then
    lblYear(Index).BackColor = RGB(255, 0, 0)
End If

lblYear(Index) = Int(lblYear(Index) + 0.05)
End Sub
