VERSION 5.00
Begin VB.Form frmPaycheck 
   Caption         =   "Paycheck"
   ClientHeight    =   4665
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   8370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdComputePay 
      Caption         =   "Compute Pay"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   7200
      TabIndex        =   10
      Top             =   1560
      Width           =   975
   End
   Begin VB.Frame fraPayStub 
      Caption         =   "Pay Stub"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   6975
      Begin VB.Label Label9 
         Caption         =   "Total Taxes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   20
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label lblTotalTax 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3840
         TabIndex        =   19
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Federal Tax"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "State Tax"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   17
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Local Tax"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Social Security"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   15
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label lblSocialSecurity 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1920
         TabIndex        =   14
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label lblLocalTax 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label lblStateTax 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1920
         TabIndex        =   12
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lblFederalTax 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lblNetPay 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1920
         TabIndex        =   9
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Net Pay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblGrossPay 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Gross Pay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame fraInput 
      Caption         =   "Input"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtHours 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Overtime Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   26
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblOvertimeRate 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4440
         TabIndex        =   25
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblOvertimePay 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6240
         TabIndex        =   24
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label11 
         Caption         =   "Overtime Pay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6240
         TabIndex        =   23
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblOvertimeHours 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2640
         TabIndex        =   22
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Overtime Hours"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   21
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Hours"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Menu mnuOperations 
      Caption         =   "&Operations"
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuComputePay 
         Caption         =   "Compute &Pay"
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "frmPaycheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Austin Herring
'Project Name: Paycheck
'Start date: September 2, 2009
'Description: A program to compute pay stubs and pay checks for hourly rate employees.



Private Sub cmdComputePay_Click()

If txtHours <= 40 Then
    lblGrossPay = txtHours * txtRate
    lblOvertimePay = ""
    lblOvertimeRate = ""
    lblOvertimeHours = ""
Else    ' otherwise compute overtime pay
    lblOvertimeHours = txtHours - 40
    lblOvertimeRate = txtRate * 1.5
    lblOvertimePay = lblOvertimeHours * lblOvertimeRate
    lblGrossPay = 40 * txtRate + lblOvertimePay
End If

lblFederalTax = lblGrossPay * 0.08
lblStateTax = lblGrossPay * 0.02
lblLocalTax = lblGrossPay * 0.01
lblSocialSecurity = lblGrossPay * 0.07
lblTotalTax = lblGrossPay * 0.18
lblNetPay = lblGrossPay - lblTotalTax

End Sub

Private Sub mnuClear_Click()
txtRate = ""
txtHours = ""
End Sub

Private Sub mnuComputePay_Click()
cmdComputePay_Click
End Sub

Private Sub txtHours_Change()
If Val(txtHours) > 50 Or Val(txtHours) <= 0 Then txtHours = ""
End Sub

Private Sub txtHours_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then txtRate.SetFocus
End Sub

Private Sub txtRate_Change()
'If Val(txtRate) < 7 Then txtRate = ""
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdComputePay_Click
End Sub



