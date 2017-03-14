VERSION 5.00
Begin VB.Form frmBoxes 
   Caption         =   "Boxes"
   ClientHeight    =   9165
   ClientLeft      =   2490
   ClientTop       =   1695
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   ScaleHeight     =   611
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   703
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Menu"
      Height          =   375
      Left            =   7800
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtLength 
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
      Text            =   "300"
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      ForeColor       =   &H00000000&
      Height          =   8415
      Left            =   120
      ScaleHeight     =   557
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   557
      TabIndex        =   2
      Top             =   600
      Width           =   8415
   End
   Begin VB.CommandButton cmdDraw 
      Caption         =   "Draw"
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
      Left            =   5280
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox txtLevels 
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
      Left            =   1320
      TabIndex        =   0
      Text            =   "4"
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Length:"
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
      Left            =   2520
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Levels:"
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
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmBoxes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Vertex
    X As Integer
    Y As Integer
End Type

Dim v(4) As Vertex

Private Sub draw_line(p1 As Vertex, p2 As Vertex)

    Picture1.Line (p1.X, p1.Y)-(p2.X, p2.Y), RGB(0, 0, 0)

End Sub


Private Sub cmdDraw_Click()

Picture1.Cls

Dim iterations As Integer
Dim length As Single

length = Val(txtLength)
iterations = Val(txtLevels)

v(1).X = Picture1.ScaleWidth / 2 - length / 2
v(1).Y = Picture1.ScaleHeight / 2 - length / 2

v(2).X = v(1).X
v(2).Y = v(1).Y + length

v(3).X = v(2).X + length
v(3).Y = v(2).Y

v(4).X = v(3).X
v(4).Y = v(1).Y

Picture1.Line (v(1).X, v(1).Y)-Step(length, length), RGB(100, 100, 100), BF

draw_line v(1), v(2)
draw_line v(2), v(3)
draw_line v(3), v(4)
draw_line v(4), v(1)

If iterations - 1 > 0 Then
    draw v(1), length / 2, iterations - 1
    draw v(2), length / 2, iterations - 1
    draw v(3), length / 2, iterations - 1
    draw v(4), length / 2, iterations - 1
End If

End Sub

Private Sub draw(vc As Vertex, length As Single, iterations As Integer)

Dim v1(4) As Vertex

v1(1).X = vc.X - (length / 2): v1(1).Y = vc.Y - (length / 2)

v1(2).X = v1(1).X
v1(2).Y = v1(1).Y + length

v1(3).X = v1(2).X + length
v1(3).Y = v1(2).Y

v1(4).X = v1(3).X
v1(4).Y = v1(1).Y

Picture1.Line (v1(1).X, v1(1).Y)-Step(length, length), RGB(100, 100, 100), BF

draw_line v1(1), v1(2)
draw_line v1(2), v1(3)
draw_line v1(3), v1(4)
draw_line v1(4), v1(1)

If iterations - 1 > 0 Then
    draw v1(1), length / 2, iterations - 1
    draw v1(2), length / 2, iterations - 1
    draw v1(3), length / 2, iterations - 1
    draw v1(4), length / 2, iterations - 1
End If

End Sub

Private Sub Command1_Click()
frmTree.Show
End Sub

Private Sub cmdMenu_Click()
frmStartup.Show
frmBoxes.Hide
End Sub
