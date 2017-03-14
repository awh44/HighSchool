VERSION 5.00
Begin VB.Form frmWire3D 
   Caption         =   "3D Wireframe"
   ClientHeight    =   9585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11205
   LinkTopic       =   "Form2"
   ScaleHeight     =   9585
   ScaleWidth      =   11205
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSphere 
      Caption         =   "Sphere"
      Height          =   195
      Left            =   10080
      TabIndex        =   44
      Top             =   1080
      Width           =   975
   End
   Begin VB.CheckBox chkClear 
      Caption         =   "Clear"
      Height          =   255
      Left            =   9120
      TabIndex        =   43
      Top             =   720
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   255
      Left            =   9240
      TabIndex        =   42
      Top             =   3960
      Width           =   135
   End
   Begin VB.CheckBox chkCone 
      Caption         =   "Cone"
      Height          =   255
      Left            =   10080
      TabIndex        =   41
      Top             =   720
      Width           =   1095
   End
   Begin VB.CheckBox chkCylinder 
      Caption         =   "Cylinder"
      Height          =   255
      Left            =   10080
      TabIndex        =   40
      Top             =   480
      Width           =   1095
   End
   Begin VB.CheckBox chkCube 
      Caption         =   "Cube"
      Height          =   255
      Left            =   10080
      TabIndex        =   39
      Top             =   240
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "2D Wireframe"
      Height          =   495
      Left            =   10080
      TabIndex        =   37
      Top             =   8520
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Polar Thing"
      Height          =   495
      Left            =   9120
      TabIndex        =   36
      Top             =   8520
      Width           =   855
   End
   Begin VB.CheckBox chkRotateZM 
      Caption         =   "Check6"
      Height          =   195
      Left            =   10920
      TabIndex        =   25
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox chkRotateYM 
      Caption         =   "Check5"
      Height          =   195
      Left            =   10920
      TabIndex        =   24
      Top             =   2880
      Width           =   255
   End
   Begin VB.CheckBox chkRotateXM 
      Caption         =   "Check4"
      Height          =   195
      Left            =   10920
      TabIndex        =   23
      Top             =   2640
      Width           =   255
   End
   Begin VB.CheckBox chkRotateZP 
      Caption         =   "Check3"
      Height          =   195
      Left            =   10680
      TabIndex        =   22
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox chkRotateYP 
      Caption         =   "Check2"
      Height          =   195
      Left            =   10680
      TabIndex        =   21
      Top             =   2880
      Width           =   255
   End
   Begin VB.CheckBox chkRotateXP 
      Caption         =   "Check1"
      Height          =   195
      Left            =   10680
      TabIndex        =   20
      Top             =   2640
      Width           =   255
   End
   Begin VB.Timer Rotation 
      Interval        =   1
      Left            =   10680
      Top             =   3840
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   375
      Left            =   9600
      TabIndex        =   19
      Top             =   7080
      Width           =   855
   End
   Begin VB.TextBox txtRotateZ 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9600
      TabIndex        =   15
      Text            =   "0"
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox txtRotateY 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9600
      TabIndex        =   14
      Text            =   "0"
      Top             =   3840
      Width           =   855
   End
   Begin VB.HScrollBar hscrRotateZ 
      Height          =   255
      LargeChange     =   20
      Left            =   9480
      Max             =   360
      Min             =   -360
      SmallChange     =   5
      TabIndex        =   13
      Top             =   3120
      Width           =   1095
   End
   Begin VB.HScrollBar hscrRotateY 
      Height          =   255
      LargeChange     =   20
      Left            =   9480
      Max             =   360
      Min             =   -360
      SmallChange     =   5
      TabIndex        =   12
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtTranslateZ 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9600
      TabIndex        =   11
      Text            =   "0"
      Top             =   6600
      Width           =   855
   End
   Begin VB.HScrollBar hscrTranslateZ 
      Height          =   255
      Left            =   9480
      Max             =   20
      Min             =   -20
      TabIndex        =   10
      Top             =   5520
      Width           =   1095
   End
   Begin VB.HScrollBar hscrScale 
      Height          =   255
      Left            =   9480
      Max             =   20
      TabIndex        =   9
      Top             =   1560
      Value           =   10
      Width           =   1095
   End
   Begin VB.HScrollBar hscrRotateX 
      Height          =   255
      LargeChange     =   20
      Left            =   9480
      Max             =   360
      Min             =   -360
      SmallChange     =   5
      TabIndex        =   8
      Top             =   2640
      Width           =   1095
   End
   Begin VB.HScrollBar hscrTranslateX 
      Height          =   255
      Left            =   9480
      Max             =   20
      Min             =   -20
      TabIndex        =   7
      Top             =   5040
      Width           =   1095
   End
   Begin VB.HScrollBar hscrTranslateY 
      Height          =   255
      Left            =   9480
      Max             =   20
      Min             =   -20
      TabIndex        =   6
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox txtScale 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9600
      TabIndex        =   5
      Text            =   "1."
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox txtRotateX 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9600
      TabIndex        =   4
      Text            =   "0"
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox txtTranslateX 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9600
      TabIndex        =   3
      Text            =   "0"
      Top             =   5880
      Width           =   855
   End
   Begin VB.TextBox txtTranslateY 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9600
      TabIndex        =   2
      Text            =   "0"
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton cmdDrawCube 
      Caption         =   "Draw"
      Height          =   375
      Left            =   9120
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.PictureBox pctWire3D 
      BackColor       =   &H0080FF80&
      Height          =   9055
      Left            =   0
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   0
      Top             =   0
      Width           =   9055
   End
   Begin VB.Label Label10 
      Height          =   390
      Left            =   120
      TabIndex        =   38
      Top             =   9120
      Width           =   4815
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
      Left            =   10920
      TabIndex        =   35
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label15 
      Caption         =   "+"
      Height          =   255
      Left            =   10680
      TabIndex        =   34
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label14 
      Caption         =   "+"
      Height          =   255
      Left            =   10680
      TabIndex        =   33
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Label13 
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
      Left            =   9360
      TabIndex        =   32
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Z"
      Height          =   255
      Left            =   9120
      TabIndex        =   31
      Top             =   5520
      Width           =   255
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Y"
      Height          =   255
      Left            =   9120
      TabIndex        =   30
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "X"
      Height          =   255
      Left            =   9120
      TabIndex        =   29
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Z"
      Height          =   255
      Left            =   9120
      TabIndex        =   28
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Y"
      Height          =   255
      Left            =   9120
      TabIndex        =   27
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "X"
      Height          =   255
      Left            =   9120
      TabIndex        =   26
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Scale"
      Height          =   255
      Left            =   9480
      TabIndex        =   18
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Rotate"
      Height          =   255
      Left            =   9480
      TabIndex        =   17
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Translate"
      Height          =   255
      Left            =   9600
      TabIndex        =   16
      Top             =   4680
      Width           =   855
   End
End
Attribute VB_Name = "frmWire3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim w As New Wire3D
Dim x_keep, y_keep As Integer
'
'


Private Sub cmdDrawCube_Click()

txtScale = 1

hscrRotateX = 0
hscrRotateY = 0
hscrRotateZ = 0

hscrTranslateX = 0
hscrTranslateY = 0
hscrTranslateZ = 0


If chkCube.Value = 1 Then
    w.initialize
End If

If chkCylinder.Value = 1 Then
    w.init_cylinder
End If

If chkCone.Value = 1 Then
    w.init_cone
End If

If chkSphere.Value = 1 Then
    w.init_sphere
End If

End Sub

Private Sub Command3_Click()
frmWire2D.Show
End Sub

Private Sub hscrRotateX_Change()
txtRotateX = hscrRotateX
End Sub

Private Sub hscrRotateY_Change()
txtRotateY = hscrRotateY
End Sub

Private Sub hscrRotateZ_Change()
txtRotateZ = hscrRotateZ
End Sub

Private Sub hscrScale_Change()
txtScale = hscrScale / 10
End Sub

Private Sub hscrTranslateX_Change()
txtTranslateX = hscrTranslateX
End Sub

Private Sub hscrTranslateY_Change()
txtTranslateY = hscrTranslateY
End Sub

Private Sub hscrTranslateZ_Change()
txtTranslateZ = hscrTranslateZ
End Sub





Private Sub Rotation_Timer()

If hscrRotateX >= 360 Then
    hscrRotateX = -360
End If
If chkRotateXP.Value = 1 Then
    hscrRotateX = hscrRotateX + 5
End If

If hscrRotateX <= -360 Then
    hscrRotateX = 360
End If
If chkRotateXM.Value = 1 Then
    hscrRotateX = hscrRotateX - 5
End If


If hscrRotateY >= 360 Then
    hscrRotateY = -360
End If
If chkRotateYP.Value = 1 Then
    hscrRotateY = hscrRotateY + 5
End If

If hscrRotateY <= -360 Then
    hscrRotateY = 360
End If
If chkRotateYM.Value = 1 Then
    hscrRotateY = hscrRotateY - 5
End If


If hscrRotateZ >= 360 Then
    hscrRotateZ = -360
End If
If chkRotateZP.Value = 1 Then
    hscrRotateZ = hscrRotateZ + 5
End If

If hscrRotateZ <= -360 Then
    hscrRotateZ = 360
End If
If chkRotateZM.Value = 1 Then
    hscrRotateZ = hscrRotateZ - 5
End If

End Sub

Private Sub txtRotateX_Change()
If chkCube.Value = 1 Then w.initialize
If chkCylinder.Value = 1 Then w.init_cylinder

If chkCone.Value = 1 Then
    w.init_cone
End If
If chkSphere.Value = 1 Then w.init_sphere
End Sub

Private Sub txtRotateY_Change()
If chkCube.Value = 1 Then w.initialize
If chkCylinder.Value = 1 Then w.init_cylinder

If chkCone.Value = 1 Then
    w.init_cone
End If
If chkSphere.Value = 1 Then w.init_sphere
End Sub

Private Sub txtRotateZ_Change()
If chkCube.Value = 1 Then w.initialize
If chkCylinder.Value = 1 Then w.init_cylinder

If chkCone.Value = 1 Then
    w.init_cone
End If
If chkSphere.Value = 1 Then w.init_sphere
End Sub


Private Sub txtScale_Change()
If chkCube.Value = 1 Then w.initialize
If chkCylinder.Value = 1 Then w.init_cylinder
If chkCone.Value = 1 Then w.init_cone
If chkSphere.Value = 1 Then w.init_sphere
End Sub

Private Sub txtTranslateX_Change()
If chkCube.Value = 1 Then w.initialize
If chkCylinder.Value = 1 Then w.init_cylinder
If chkCone.Value = 1 Then w.init_cone
If chkSphere.Value = 1 Then w.init_sphere
End Sub

Private Sub txtTranslateY_Change()
If chkCube.Value = 1 Then w.initialize
If chkCylinder.Value = 1 Then w.init_cylinder
If chkCone.Value = 1 Then w.init_cone
If chkSphere.Value = 1 Then w.init_sphere
End Sub

Private Sub txtTranslateZ_Change()
If chkCube.Value = 1 Then w.initialize
If chkCylinder.Value = 1 Then w.init_cylinder
If chkCone.Value = 1 Then w.init_cone
If chkSphere.Value = 1 Then w.init_sphere
End Sub
