VERSION 5.00
Begin VB.Form frmWire2D 
   Caption         =   "2D Wireframe"
   ClientHeight    =   9960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   ScaleHeight     =   664
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   806
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtInc 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   10800
      TabIndex        =   35
      Text            =   "10"
      Top             =   9000
      Width           =   855
   End
   Begin VB.TextBox txtAngle 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   10800
      TabIndex        =   34
      Text            =   "0"
      Top             =   8520
      Width           =   855
   End
   Begin VB.CommandButton cmdRotate 
      Caption         =   "Rotate"
      Height          =   615
      Left            =   9480
      TabIndex        =   33
      Top             =   8520
      Width           =   975
   End
   Begin VB.CommandButton cmdDrawSquare 
      Caption         =   "Draw Square"
      Height          =   615
      Left            =   10800
      TabIndex        =   32
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton cmdInitSquare 
      Caption         =   "Initialize Square"
      Height          =   615
      Left            =   9360
      TabIndex        =   31
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CheckBox chkYMinus 
      Caption         =   "Check3"
      Height          =   195
      Left            =   11400
      TabIndex        =   30
      Top             =   2880
      Width           =   255
   End
   Begin VB.CheckBox chkYPlus 
      Caption         =   "Check6"
      Height          =   195
      Left            =   11160
      TabIndex        =   29
      Top             =   2880
      Width           =   255
   End
   Begin VB.CheckBox chkXMinus 
      Caption         =   "Check6"
      Height          =   195
      Left            =   11400
      TabIndex        =   28
      Top             =   2640
      Width           =   255
   End
   Begin VB.CheckBox chkXPlus 
      Caption         =   "Check3"
      Height          =   195
      Left            =   11160
      TabIndex        =   27
      Top             =   2640
      Width           =   255
   End
   Begin VB.CheckBox chkClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   9600
      TabIndex        =   26
      Top             =   5640
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.TextBox txtRadius 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9600
      TabIndex        =   24
      Text            =   "7"
      Top             =   5040
      Width           =   735
   End
   Begin VB.TextBox txtTheta 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10680
      TabIndex        =   22
      Text            =   "10"
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton cmdScale 
      Caption         =   "Scale"
      Height          =   375
      Left            =   9840
      TabIndex        =   21
      Top             =   6720
      Width           =   735
   End
   Begin VB.TextBox txtLength 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10800
      TabIndex        =   20
      Text            =   "200"
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdInitialize 
      Caption         =   "Initialize"
      Height          =   375
      Left            =   9720
      TabIndex        =   19
      Top             =   240
      Width           =   735
   End
   Begin VB.Timer Rotation 
      Interval        =   50
      Left            =   11640
      Top             =   840
   End
   Begin VB.CheckBox chkRotatePlus 
      Caption         =   "Check3"
      Height          =   195
      Left            =   11160
      TabIndex        =   16
      Top             =   2160
      Width           =   255
   End
   Begin VB.CheckBox chkRotateMinus 
      Caption         =   "Check6"
      Height          =   195
      Left            =   11400
      TabIndex        =   15
      Top             =   2160
      Width           =   255
   End
   Begin VB.OptionButton optCircle 
      Caption         =   "Circle"
      Height          =   255
      Left            =   9480
      TabIndex        =   14
      Top             =   4200
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.OptionButton optSquare 
      Caption         =   "Square"
      Height          =   375
      Left            =   9480
      TabIndex        =   13
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox txtTranslateY 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10560
      TabIndex        =   9
      Text            =   "0"
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox txtTranslateX 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9720
      TabIndex        =   8
      Text            =   "0"
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox txtRotate 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10080
      TabIndex        =   7
      Text            =   "0"
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox txtScale 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10080
      TabIndex        =   6
      Text            =   "1."
      Top             =   1320
      Width           =   855
   End
   Begin VB.HScrollBar hscrTranslateY 
      Height          =   255
      Left            =   9960
      Max             =   20
      Min             =   -20
      TabIndex        =   5
      Top             =   2880
      Width           =   1095
   End
   Begin VB.HScrollBar hscrTranslateX 
      Height          =   255
      Left            =   9960
      Max             =   20
      Min             =   -20
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
   Begin VB.HScrollBar hscrRotate 
      Height          =   255
      LargeChange     =   20
      Left            =   9960
      Max             =   360
      Min             =   -360
      SmallChange     =   5
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin VB.HScrollBar hscrScale 
      Height          =   255
      LargeChange     =   100
      Left            =   9960
      Max             =   200
      Min             =   10
      SmallChange     =   10
      TabIndex        =   2
      Top             =   960
      Value           =   100
      Width           =   1095
   End
   Begin VB.CommandButton cmdDrawWire2D 
      Caption         =   "Draw"
      Height          =   375
      Left            =   10680
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.PictureBox pctWire2D 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0080FF80&
      Height          =   9780
      Left            =   -840
      ScaleHeight     =   648
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   648
      TabIndex        =   0
      Top             =   120
      Width           =   9780
   End
   Begin VB.Label Label5 
      Caption         =   "Radius"
      Height          =   255
      Left            =   9840
      TabIndex        =   25
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Theta"
      Height          =   255
      Left            =   10920
      TabIndex        =   23
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label15 
      Caption         =   "+"
      Height          =   255
      Left            =   11160
      TabIndex        =   18
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Label16 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11400
      TabIndex        =   17
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Translate"
      Height          =   255
      Left            =   9120
      TabIndex        =   12
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Rotate"
      Height          =   255
      Left            =   9240
      TabIndex        =   11
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Scale"
      Height          =   255
      Left            =   9360
      TabIndex        =   10
      Top             =   1200
      Width           =   495
   End
End
Attribute VB_Name = "frmWire2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Austin Herring


Dim w As New Wire2D
Dim w2 As New Wire2Dxy



Private Sub cmdDrawSquare_Click()
w2.Draw
End Sub

    'Dimension "w" as an object of the class Wire2D

Private Sub cmdDrawWire2D_Click()

'pctWire2D.Cls

w.Draw frmWire2D.pctWire2D
End Sub


Private Sub cmdInitialize_Click()

txtScale = 1
txtRotate = 0
txtTranslateY = 0
txtTranslateX = 0

If optSquare = True Then
    w.initialize
End If

If optCircle = True Then
    w.init_circle txtTheta, txtRadius
End If

End Sub

Private Sub cmdInitSquare_Click()
w2.init_square

End Sub

Private Sub cmdRotate_Click()
w2.rotate
End Sub

Private Sub cmdScale_Click()

w.re_size

End Sub

Private Sub hscrRotate_Change()

txtRotate = hscrRotate



End Sub

Private Sub hscrScale_Change()

txtScale = hscrScale / 100

End Sub

Private Sub hscrTranslateX_Change()

txtTranslateX = hscrTranslateX.Value

End Sub

Private Sub hscrTranslateY_Change()

txtTranslateY = hscrTranslateY.Value

End Sub

Private Sub Rotation_Timer()

If hscrRotate >= 360 Then
    hscrRotate = -360
End If
If chkRotatePlus.Value = 1 Then
    hscrRotate = hscrRotate + 5
End If


If hscrRotate <= -360 Then
    hscrRotate = 360
End If
If chkRotateMinus.Value = 1 Then
    hscrRotate = hscrRotate - 5
End If



If hscrTranslateX = 20 Then
    hscrTranslateX = -20
End If
If chkXPlus.Value = 1 Then
    hscrTranslateX = hscrTranslateX + 1
End If

If hscrTranslateX = -20 Then
    hscrTranslateX = 20
End If
If chkXMinus.Value = 1 Then
    hscrTranslateX = hscrTranslateX - 1
End If




If hscrTranslateY = 20 Then
    hscrTranslateY = -20
End If
If chkYPlus.Value = 1 Then
    hscrTranslateY = hscrTranslateY + 1
End If

If hscrTranslateY = -20 Then
    hscrTranslateY = 20
End If
If chkYMinus.Value = 1 Then
    hscrTranslateY = hscrTranslateY - 1
End If



End Sub

Private Sub txtRotate_Change()

If optSquare = True Then
    w.initialize
End If

If optCircle = True Then
    w.init_circle txtTheta, txtRadius
End If

End Sub

Private Sub txtScale_Change()


'w.re_size (txtScale)

If optSquare = True Then
    w.initialize
Else
    w.init_circle txtTheta, txtRadius
End If

End Sub

Private Sub txtTranslateX_Change()

If optSquare = True Then
    w.initialize
Else
    w.init_circle txtTheta, txtRadius
End If

End Sub

Private Sub txtTranslateY_Change()

If optSquare = True Then
    w.initialize
Else
    w.init_circle txtTheta, txtRadius
End If


End Sub
