VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSComCt2.ocx"
Begin VB.Form frmTree 
   Caption         =   "Fractal Trees"
   ClientHeight    =   9165
   ClientLeft      =   2310
   ClientTop       =   1305
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   ScaleHeight     =   9165
   ScaleWidth      =   10545
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Menu"
      Height          =   375
      Left            =   9480
      TabIndex        =   13
      Top             =   600
      Width           =   735
   End
   Begin VB.CheckBox chkThickness 
      Caption         =   "Vary Thickness"
      Height          =   255
      Left            =   2760
      TabIndex        =   12
      Top             =   600
      Width           =   1695
   End
   Begin VB.CheckBox chkColor 
      Caption         =   "Color"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   600
      Value           =   1  'Checked
      Width           =   855
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   420
      Left            =   2280
      TabIndex        =   10
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   741
      _Version        =   393216
      Value           =   4
      BuddyControl    =   "txtIterations"
      BuddyDispid     =   196615
      OrigLeft        =   2520
      OrigTop         =   240
      OrigRight       =   2760
      OrigBottom      =   615
      Max             =   100
      Enabled         =   -1  'True
   End
   Begin VB.CheckBox chkRandom 
      Caption         =   "Random Angle"
      Height          =   195
      Left            =   1200
      TabIndex        =   9
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox txtTrunk 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   7
      Text            =   "100"
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtPercentage 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   5
      Text            =   "60"
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtAngle 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Text            =   "45"
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtIterations 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1560
      TabIndex        =   2
      Text            =   "4"
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      ForeColor       =   &H00000000&
      Height          =   8175
      Left            =   120
      ScaleHeight     =   541
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   677
      TabIndex        =   0
      Top             =   960
      Width           =   10215
   End
   Begin VB.Label Label4 
      Caption         =   "Trunk Length:"
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
      Left            =   7560
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Percentage:"
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
      Left            =   4800
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Angle:"
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
      Left            =   2880
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Iterations:"
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
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Austin Herring
'frmTree

'random turns
'change color
'change thickness


Private Type Color
    r As Integer
    g As Integer
    b As Integer
End Type


Private Sub Command1_Click()
frmSierpinski.Show
End Sub


Private Sub Command2_Click()
frmTriangles.Show
End Sub




Private Sub Command3_Click()
frmSquares.Show
End Sub

Private Sub cmdMenu_Click()
frmStartup.Show
frmTree.Hide
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim iterations, length, length1, angle, angle1 As Integer
Dim percentage, XX, YY As Single

If chkThickness.Value = 1 Then
    Picture1.drawwidth = 13
Else
    Picture1.drawwidth = 1
End If

iterations = txtIterations
length = txtTrunk
angle = txtAngle
angle1 = angle
angle2 = angle
percentage = txtPercentage / 100


Picture1.Cls

If chkColor.Value = 1 Then
    Picture1.Line (X, Y)-(X, Y - length), RGB(139, 69, 19)
Else
    Picture1.Line (X, Y)-(X, Y - length)
End If



XX = X
YY = Y - length
length1 = length


If iterations - 1 > 0 Then
    Tree iterations - 1, length, 90 + angle, angle1, percentage, X, Y - length, txtIterations / 2, 10 * 0.7
    Tree iterations - 1, length1, 90 - angle, angle1, percentage, X, YY, txtIterations / 2, 10 * 0.7
End If

    
    
    

End Sub

Private Sub Tree(iterations, length, angle, angle1 As Integer, percentage, X, Y As Single, half As Integer, drawwidth As Single)

Dim X2, Y2 As Single
Dim radians As Single
Dim length1 As Integer

If drawwidth > 0.5 And chkThickness.Value = 1 Then
    Picture1.drawwidth = drawwidth
Else
    Picture1.drawwidth = 1
End If

radians = angle * (3.14159 / 180)

length = percentage * length
length1 = length
X2 = X + (Cos(radians) * length)
Y2 = Y - (Sin(radians) * length)

If chkColor.Value = 1 Then
    If iterations > half Then
        Picture1.Line (X, Y)-(X2, Y2), RGB(139, 69, 19)
    Else
        Picture1.Line (X, Y)-(X2, Y2), RGB(0, 100, 0)
    End If
Else
    Picture1.Line (X, Y)-(X2, Y2)
End If
    

If iterations - 1 > 0 Then
    If chkRandom.Value <> 1 Then
        Tree iterations - 1, length, angle - angle1, angle1, percentage, X2, Y2, half, drawwidth * 0.7
        Tree iterations - 1, length1, angle + angle1, angle1, percentage, X2, Y2, half, drawwidth * 0.7
    Else
        Randomize
        Tree iterations - 1, length, angle - (Rnd * 90), angle1, percentage, X2, Y2, half, drawwidth * 0.7
        Tree iterations - 1, length1, angle + (Rnd * 90), angle1, percentage, X2, Y2, half, drawwidth * 0.7
    End If

    
End If


End Sub

Private Sub UpDown1_Change()

txtIterations = UpDown1.Value



End Sub
