VERSION 5.00
Begin VB.Form frmGrades 
   Caption         =   "Grades"
   ClientHeight    =   10785
   ClientLeft      =   2115
   ClientTop       =   345
   ClientWidth     =   11085
   LinkTopic       =   "Form1"
   ScaleHeight     =   10785
   ScaleWidth      =   11085
   Begin VB.Frame fraSubject 
      Caption         =   "Subject"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5280
      TabIndex        =   73
      Top             =   9000
      Width           =   5295
      Begin VB.OptionButton optClass 
         Caption         =   "Computer Science"
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   83
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optClass 
         Caption         =   "Chemistry"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   82
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton optClass 
         Caption         =   "Traffic Safety"
         Height          =   255
         Index           =   9
         Left            =   1200
         TabIndex        =   81
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton optClass 
         Caption         =   "Phys. Ed."
         Height          =   255
         Index           =   8
         Left            =   4200
         TabIndex        =   80
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton optClass 
         Caption         =   "Algebra"
         Height          =   255
         Index           =   6
         Left            =   3000
         TabIndex        =   79
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton optClass 
         Caption         =   "German"
         Height          =   255
         Index           =   5
         Left            =   1200
         TabIndex        =   78
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton optClass 
         Caption         =   "History"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   77
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton optClass 
         Caption         =   "English"
         Height          =   255
         Index           =   3
         Left            =   4200
         TabIndex        =   76
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optClass 
         Caption         =   "Career Ed"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   75
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optClass 
         Caption         =   "Band"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   74
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame fraQuarter 
      Caption         =   "Quarter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5280
      TabIndex        =   68
      Top             =   7800
      Width           =   3855
      Begin VB.OptionButton optFourth 
         Caption         =   "Fourth"
         Height          =   375
         Left            =   2880
         TabIndex        =   72
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton optThird 
         Caption         =   "Third"
         Height          =   375
         Left            =   2040
         TabIndex        =   71
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton optSecond 
         Caption         =   "Second"
         Height          =   375
         Left            =   960
         TabIndex        =   70
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton optFirst 
         Caption         =   "First"
         Height          =   375
         Left            =   120
         TabIndex        =   69
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton frmReport 
      Caption         =   "Report Card Form"
      Height          =   615
      Left            =   9240
      TabIndex        =   67
      Top             =   8040
      Width           =   1455
   End
   Begin VB.Frame fraAverages 
      Caption         =   "Grades By Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   240
      TabIndex        =   51
      Top             =   8040
      Width           =   4815
      Begin VB.Label lblHomework 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3600
         TabIndex        =   66
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblTest 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3600
         TabIndex        =   65
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblProject 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3600
         TabIndex        =   64
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblPMax 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2520
         TabIndex        =   63
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblPPoints 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   62
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblTeMax 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2520
         TabIndex        =   61
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblTePoints 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   60
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblHMax 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2520
         TabIndex        =   59
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblHPoints 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   58
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Homework"
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
         Left            =   120
         TabIndex        =   57
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Test"
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
         Left            =   120
         TabIndex        =   56
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Project"
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
         Left            =   120
         TabIndex        =   55
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Points"
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
         Left            =   1440
         TabIndex        =   54
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "Max"
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
         Left            =   2520
         TabIndex        =   53
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Percentage"
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
         Left            =   3480
         TabIndex        =   52
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.TextBox txtMax 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   5520
      TabIndex        =   45
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtMax 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   5520
      TabIndex        =   44
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtMax 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   5520
      TabIndex        =   43
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtMax 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   5520
      TabIndex        =   42
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Frame fraAssignments 
      Caption         =   "Assignments"
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
      Width           =   10815
      Begin VB.VScrollBar VScroll1 
         Height          =   6375
         LargeChange     =   50
         Left            =   10560
         Max             =   0
         SmallChange     =   20
         TabIndex        =   50
         Top             =   120
         Width           =   255
      End
      Begin VB.CommandButton cmdCompute 
         Caption         =   "Compute"
         Height          =   255
         Index           =   4
         Left            =   9480
         TabIndex        =   49
         Top             =   2640
         Width           =   855
      End
      Begin VB.CommandButton cmdCompute 
         Caption         =   "Compute"
         Height          =   255
         Index           =   3
         Left            =   9480
         TabIndex        =   48
         Top             =   2160
         Width           =   855
      End
      Begin VB.CommandButton cmdCompute 
         Caption         =   "Compute"
         Height          =   255
         Index           =   2
         Left            =   9480
         TabIndex        =   47
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton cmdCompute 
         Caption         =   "Compute"
         Height          =   255
         Index           =   1
         Left            =   9480
         TabIndex        =   46
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtMax 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   5400
         TabIndex        =   41
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Index           =   3
         Left            =   240
         TabIndex        =   15
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Index           =   4
         Left            =   240
         TabIndex        =   14
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox txtPoints 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   3960
         TabIndex        =   13
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtPoints 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   3960
         TabIndex        =   12
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtPoints 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   3960
         TabIndex        =   11
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtPoints 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   3960
         TabIndex        =   10
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtPoints 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   3960
         TabIndex        =   9
         Top             =   2640
         Width           =   1215
      End
      Begin VB.ComboBox cmbType 
         Height          =   315
         Index           =   0
         ItemData        =   "frmGrades.frx":0000
         Left            =   2520
         List            =   "frmGrades.frx":000D
         TabIndex        =   8
         Text            =   "Project"
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox cmbType 
         Height          =   315
         Index           =   1
         ItemData        =   "frmGrades.frx":002A
         Left            =   2520
         List            =   "frmGrades.frx":0037
         TabIndex        =   7
         Text            =   "Project"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.ComboBox cmbType 
         Height          =   315
         Index           =   2
         ItemData        =   "frmGrades.frx":0054
         Left            =   2520
         List            =   "frmGrades.frx":0061
         TabIndex        =   6
         Text            =   "Project"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.ComboBox cmbType 
         Height          =   315
         Index           =   3
         ItemData        =   "frmGrades.frx":007E
         Left            =   2520
         List            =   "frmGrades.frx":008B
         TabIndex        =   5
         Text            =   "Project"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.ComboBox cmbType 
         Height          =   315
         Index           =   4
         ItemData        =   "frmGrades.frx":00A8
         Left            =   2520
         List            =   "frmGrades.frx":00B5
         TabIndex        =   4
         Text            =   "Project"
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CommandButton cmdCompute 
         Caption         =   "Compute"
         Height          =   255
         Index           =   0
         Left            =   9480
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton cmdCompute2 
         Caption         =   "Compute Totals"
         Height          =   255
         Left            =   2400
         TabIndex        =   2
         Top             =   3240
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add Assignment"
         Height          =   255
         Left            =   2400
         TabIndex        =   1
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
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
         TabIndex        =   40
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
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
         Left            =   2520
         TabIndex        =   39
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Points"
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
         Left            =   3960
         TabIndex        =   38
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Max"
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
         TabIndex        =   37
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Percentage"
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
         Left            =   6840
         TabIndex        =   36
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Letter Grade"
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
         Left            =   8160
         TabIndex        =   35
         Top             =   480
         Width           =   1575
      End
      Begin VB.Line Line1 
         X1              =   2400
         X2              =   2400
         Y1              =   360
         Y2              =   3120
      End
      Begin VB.Line Line2 
         X1              =   3840
         X2              =   3840
         Y1              =   360
         Y2              =   3120
      End
      Begin VB.Line Line3 
         X1              =   5280
         X2              =   5280
         Y1              =   360
         Y2              =   3120
      End
      Begin VB.Line Line4 
         X1              =   6720
         X2              =   6720
         Y1              =   360
         Y2              =   3120
      End
      Begin VB.Line Line5 
         X1              =   8040
         X2              =   8040
         Y1              =   360
         Y2              =   3120
      End
      Begin VB.Line Line6 
         X1              =   10680
         X2              =   0
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Label lblPercentage 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   6960
         TabIndex        =   34
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblPercentage 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   6960
         TabIndex        =   33
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblPercentage 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   6960
         TabIndex        =   32
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblPercentage 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   6960
         TabIndex        =   31
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lblPercentage 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   4
         Left            =   6960
         TabIndex        =   30
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   8400
         TabIndex        =   29
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   8400
         TabIndex        =   28
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   8400
         TabIndex        =   27
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   8400
         TabIndex        =   26
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   4
         Left            =   8400
         TabIndex        =   25
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label lblTPoints 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   24
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label lblTMax 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   23
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label lblTPercentage 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6840
         TabIndex        =   22
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label lblTLetter 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8400
         TabIndex        =   21
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label lblN 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Number of Assignments"
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
         Left            =   120
         TabIndex        =   19
         Top             =   3600
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmGrades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Austin Herring
'Gradebook Project
'Form Name: frmGrades
'Started: October 5, 2009

Private Sub cmbType_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then txtPoints(Index).SetFocus
End Sub

Private Sub cmdAdd_Click()
Dim n As Integer

If lblN = 11 Then
    VScroll1.Visible = True
    VScroll1.Max = 0
End If

If lblN >= 12 Then
    VScroll1.Visible = True
    VScroll1.Max = (lblN - 11) * 480
End If

If lblN <= 11 Then
    Line6.Y1 = Line6.Y1 + 480
    Line6.Y2 = Line6.Y2 + 480
    cmdCompute2.Top = cmdCompute2.Top + 480
    lblTPoints.Top = lblTPoints.Top + 480
    lblTMax.Top = lblTPoints.Top
    lblTPercentage.Top = lblTPercentage.Top + 480
    lblTLetter.Top = lblTLetter.Top + 480
    Label7.Top = Label7.Top + 480
    lblN.Top = lblN.Top + 480
    cmdAdd.Top = cmdAdd.Top + 480
    Line1.Y2 = Line1.Y2 + 480
    Line2.Y2 = Line2.Y2 + 480
    Line3.Y2 = Line3.Y2 + 480
    Line4.Y2 = Line4.Y2 + 480
    Line5.Y2 = Line5.Y2 + 480
'Else
    'For c = 0 To (lblN - 1)
       'txtName(c).Top = txtName(c).Top - 480
     '  cmbType(c).Top = cmbType(c).Top - 480
      ' txtPoints(c).Top = txtPoints(c).Top - 480
'       txtMax(c).Top = txtMax(c).Top - 480
 '      lblPercentage(c).Top = lblPercentage(c).Top - 480
  '     lblLetter(c).Top = lblLetter(c).Top - 480
   '    cmdCompute(c).Top = cmdCompute(c).Top - 480
    '   Label1.Top = txtName(0).Top - 240
     '  Label2.Top = Label1.Top
      ' Label3.Top = Label1.Top
       'Label4.Top = Label1.Top
'       Label5.Top = Label1.Top
 '      Label6.Top = Label1.Top
  '  Next c
End If

n = Val(lblN)



Load txtName(n) ' load the next assignment
txtName(n).Visible = True
txtName(n).Top = txtName(n - 1).Top + 480
txtName(n) = ""

Load cmbType(n)
cmbType(n).Visible = True
cmbType(n).Top = cmbType(n - 1).Top + 480
cmbType(n).AddItem "Project"
cmbType(n).AddItem "Test"
cmbType(n).AddItem "Homework"

Load txtPoints(n)
txtPoints(n).Visible = True
txtPoints(n).Top = txtPoints(n - 1).Top + 480
txtPoints(n) = ""

Load txtMax(n)
txtMax(n).Visible = True
txtMax(n).Top = txtPoints(n).Top
txtMax(n) = ""

Load lblPercentage(n)
lblPercentage(n).Visible = True
lblPercentage(n).Top = lblPercentage(n - 1).Top + 480
lblPercentage(n) = ""

Load lblLetter(n)
lblLetter(n).Visible = True
lblLetter(n).Top = lblLetter(n - 1).Top + 480
lblLetter(n) = ""
lblLetter(n).BackColor = &H8000000F

Load cmdCompute(n)
cmdCompute(n).Visible = True
cmdCompute(n).Top = cmdCompute(n - 1).Top + 480

If txtName(n).Top >= (Line6.Y1 - 285) Then
        txtName(n).Visible = False
        cmbType(n).Visible = False
        txtPoints(n).Visible = False
        txtMax(n).Visible = False
        lblPercentage(n).Visible = False
        lblLetter(n).Visible = False
        cmdCompute(n).Visible = False
    Else
        txtName(n).Visible = True
        cmbType(n).Visible = True
        txtPoints(n).Visible = True
        txtMax(n).Visible = True
        lblPercentage(n).Visible = True
        lblLetter(n).Visible = True
        cmdCompute(n).Visible = True
End If

lblN = lblN + 1
End Sub




Private Sub cmdCompute_Click(Index As Integer)
'compute grade for any assignment
'the assignment number is stored in 'index' starting at 0
Dim grade As Integer ' store the percentage grade

If Val(txtPoints(Index)) = 0 Or Val(txtMax(Index)) = 0 Then
    j = 1
Else

lblPercentage(Index) = txtPoints(Index) / txtMax(Index) * 100
lblPercentage(Index) = Int(lblPercentage(Index) + 0.5)

grade = Val(lblPercentage(Index)) ' store grade in variable

'92 up is A
'83-91 is B
'74-82 C
'65-73 D
'< 65 F



If grade >= 92 Then
    lblLetter(Index) = "A"
    lblLetter(Index).BackColor = RGB(0, 100, 255)
End If

If 83 <= grade And grade < 92 Then
    lblLetter(Index) = "B"
    lblLetter(Index).BackColor = RGB(0, 255, 0)
End If

If 74 <= grade And grade < 83 Then
    lblLetter(Index) = "C"
    lblLetter(Index).BackColor = RGB(243, 237, 1)
End If

If grade >= 65 And grade < 74 Then
    lblLetter(Index) = "D"
    lblLetter(Index).BackColor = RGB(242, 116, 2)
End If

If grade < 65 Then
    lblLetter(Index) = "F"
    lblLetter(Index).BackColor = RGB(255, 0, 0)
End If

lblPercentage(Index) = lblPercentage(Index) + "%"

For c = 0 To (lblN - 1)
    If cmbType(c) = "Project" Then
        lblPPoints = 0: lblPMax = 0
    End If
Next c

For c = 0 To (lblN - 1)
    If cmbType(c) = "Project" Then
        lblPPoints = Val(lblPPoints) + Val(txtPoints(c))
        lblPMax = Val(lblPMax) + Val(txtMax(c))
        lblProject = lblPPoints / lblPMax * 100
        lblProject = Int(lblProject + 0.5)
        lblProject = lblProject + "%"
    End If
Next c

For c = 0 To (lblN - 1)
    If cmbType(c) = "Test" Then
        lblTePoints = 0: lblTeMax = 0
    End If
Next c

For c = 0 To (lblN - 1)
    If cmbType(c) = "Test" Then
    lblTePoints = Val(lblTePoints) + Val(txtPoints(c))
    lblTeMax = Val(lblTeMax) + Val(txtMax(c))
    lblTest = lblTePoints / lblTeMax * 100
    lblTest = Int(lblTest + 0.5)
    lblTest = lblTest + "%"
    End If
Next c

For c = 0 To (lblN - 1)
    If cmbType(c) = "Homework" Then
        lblHPoints = 0: lblHMax = 0
    End If
Next c

For c = 0 To (lblN - 1)
    If cmbType(c) = "Homework" Then
        lblHPoints = Val(lblHPoints) + Val(txtPoints(c))
        lblHMax = Val(lblHMax) + Val(txtMax(c))
        lblHomework = lblHPoints / lblHMax * 100
        lblHomework = Int(lblHomework + 0.5)
        lblHomework = lblHomework + "%"
    End If
Next c
    

End If
End Sub

Private Sub cmdCompute2_Click()
Dim Tmax, Tpoints As Integer
Dim c As Integer ' count number of assignments

Tmax = 0: Tpoints = 0

For c = 0 To (lblN - 1)
    Tpoints = Tpoints + Val(txtPoints(c))
Next c

For c = 0 To (lblN - 1)
    Tmax = Tmax + Val(txtMax(c))
Next c

lblTPoints = Tpoints
lblTMax = Tmax

If lblTMax = "" Or lblTMax = "0" Then
    j = 1
Else

lblTPercentage = lblTPoints / lblTMax * 100
lblTPercentage = Int(lblTPercentage + 0.5)

If lblTPercentage >= 92 Then
    lblTLetter = "A"
    lblTLetter.BackColor = RGB(0, 100, 255)
End If

If 83 <= lblTPercentage And lblTPercentage < 92 Then
    lblTLetter = "B"
    lblTLetter.BackColor = RGB(0, 255, 0)
End If

If 74 <= lblTPercentage And lblTPercentage < 83 Then
    lblTLetter = "C"
    lblTLetter.BackColor = RGB(243, 237, 1)
End If

If lblTPercentage >= 65 And lblTPercentage < 74 Then
    lblTLetter = "D"
    lblTLetter.BackColor = RGB(242, 116, 2)
End If

If lblTPercentage < 65 Then
    lblTLetter = "F"
    lblTLetter.BackColor = RGB(255, 0, 0)
End If

lblTPercentage = lblTPercentage

End If

End Sub


Private Sub Form_Load()
VScroll1.Visible = False
End Sub

Private Sub frmReport_Click()
frmReportCard.Show

For c = 0 To 9

If optClass(c) = True And optFirst = True Then
    frmReportCard.lblQuarter1(c) = lblTPercentage
End If

If optClass(c) = True And optSecond = True Then
    frmReportCard.lblQuarter2(c) = lblTPercentage
End If

If optClass(c) = True And optThird = True Then
    frmReportCard.lblQuarter3(c) = lblTPercentage
End If

If optClass(c) = True And optFourth = True Then
    frmReportCard.lblQuarter4(c) = lblTPercentage
End If

Next c

End Sub

Private Sub lblPercentage_Change(Index As Integer)
cmdCompute2_Click
End Sub
Private Sub txtMax_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn And Index < (lblN - 1) And txtMax(Index).Top < 6000 Then
    cmdCompute_Click (Index)
    txtName(Index + 1).SetFocus
End If

'If KeyAscii = vbKeyReturn And lblN > 12 And txtMax(Index).Top >= 6000 Then
 '   For c = 0 To (lblN - 1)
  '      txtName(c).Top = txtName(c).Top - 480
   '     txtName(c).Visible = True
    '    cmbType(c).Top = txtName(c).Top
     '   cmbType(c).Visible = True
      '  txtPoints(c).Top = txtName(c).Top
       ' txtPoints(c).Visible = True
        'txtMax(c).Top = txtName(c).Top
'        txtMax(c).Visible = True
 '       lblPercentage(c).Top = txtName(c).Top
  '      lblPercentage(c).Visible = True
   '     lblLetter(c).Top = txtName(c).Top
    '    lblLetter(c).Visible = True
     '   cmdCompute(c).Top = txtName(c).Top
      '  cmdCompute(c).Visible = True
        
       ' txtName(c).SetFocus
    'Next c
'End If
        

If KeyAscii = vbKeyReturn And Index = (lblN - 1) And Index < 11 Then
    cmdCompute_Click (Index)
    cmdAdd_Click
    txtName(lblN - 1).SetFocus
End If

If KeyAscii = vbKeyReturn And Index >= 11 Then
    cmdCompute_Click (Index)
End If
End Sub

Private Sub txtName_KeyPress(Index As Integer, KeyAscii As Integer)
' index is the index number of the active text box txtName(0) to txtName (4)
If KeyAscii = vbKeyReturn Then cmbType(Index).SetFocus
End Sub

Private Sub txtPoints_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then txtMax(Index).SetFocus
End Sub

Private Sub VScroll1_Change()
Dim c As Integer

For c = 0 To (lblN - 1)
    txtName(c).Top = (-(-(c * 480 + 720) + VScroll1.Value))
    cmbType(c).Top = txtName(c).Top
    txtPoints(c).Top = txtName(c).Top
    txtMax(c).Top = txtName(c).Top
    lblPercentage(c).Top = txtName(c).Top
    lblLetter(c).Top = txtName(c).Top
    cmdCompute(c).Top = txtName(c).Top
    Label1.Top = txtName(0).Top - 240
    Label2.Top = Label1.Top
    Label3.Top = Label1.Top
    Label4.Top = Label1.Top
    Label5.Top = Label1.Top
    Label6.Top = Label1.Top
    If txtName(c).Top >= (Line6.Y1 - 285) Then
        txtName(c).Visible = False
        cmbType(c).Visible = False
        txtPoints(c).Visible = False
        txtMax(c).Visible = False
        lblPercentage(c).Visible = False
        lblLetter(c).Visible = False
        cmdCompute(c).Visible = False
    Else
        txtName(c).Visible = True
        cmbType(c).Visible = True
        txtPoints(c).Visible = True
        txtMax(c).Visible = True
        lblPercentage(c).Visible = True
        lblLetter(c).Visible = True
        cmdCompute(c).Visible = True
    End If
Next c


End Sub

