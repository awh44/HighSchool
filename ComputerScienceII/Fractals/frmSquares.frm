VERSION 5.00
Begin VB.Form frmSquares 
   Caption         =   "Fractal Squares"
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
      Left            =   7800
      TabIndex        =   4
      Top             =   120
      Width           =   735
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
      TabIndex        =   2
      Text            =   "4"
      Top             =   120
      Width           =   975
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
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      ForeColor       =   &H00000000&
      Height          =   8415
      Left            =   120
      ScaleHeight     =   557
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   557
      TabIndex        =   0
      Top             =   600
      Width           =   8415
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
Attribute VB_Name = "frmSquares"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Austin Herring
'12-21-10

Private Type Vertex
    X As Integer
    Y As Integer
End Type

Private Type Color
    red As Integer
    green As Integer
    blue As Integer
End Type


Dim v(4) As Vertex
    

Dim j As Color
Dim h As Color



Private Sub draw_line(p1 As Vertex, p2 As Vertex, c As Color)

    Picture1.Line (p1.X, p1.Y)-(p2.X, p2.Y), RGB(c.red, c.green, c.blue)

End Sub

Private Sub clear_line(p1 As Vertex, p2 As Vertex)

    Picture1.Line (p1.X, p1.Y)-(p2.X, p2.Y), &H8000000F

End Sub

Private Sub plot_point(v As Vertex, c As Color)

    Picture1.PSet (v.X, v.Y), RGB(c.red, c.green, c.blue)
    
End Sub


Private Sub cmdDraw_Click()

Picture1.Cls

v(1).X = Picture1.ScaleWidth / 2 - 200: v(1).Y = Picture1.ScaleHeight / 2 - 200
v(2).X = v(1).X: v(2).Y = Picture1.ScaleHeight / 2 + 200
v(3).X = Picture1.ScaleWidth / 2 + 200: v(3).Y = v(2).Y
v(4).X = v(3).X: v(4).Y = v(1).Y

draw_line v(1), v(2), j
draw_line v(2), v(3), j
draw_line v(3), v(4), j
draw_line v(4), v(1), j

'plot_point v(1), j
'plot_point v(2), j
'plot_point v(3), j
'plot_point v(4), j



If txtLevels - 1 > 0 Then
    draw_square v(1), v(2), v(3), v(4), txtLevels - 1
End If


End Sub

Private Sub draw_square(v1 As Vertex, v2 As Vertex, v3 As Vertex, v4 As Vertex, iterations As Integer)

'numbering goes counter-clockwise

Dim distance As Single
Dim n(4) As Vertex
Dim s(16) As Vertex

distance = (v2.Y - v1.Y) / 3

n(1).X = v1.X: n(1).Y = v1.Y + distance
n(2).X = n(1).X: n(2).Y = n(1).Y + distance
n(3).X = n(2).X + distance: n(3).Y = n(2).Y
n(4).X = n(3).X: n(4).Y = n(1).Y

s(1) = v1
s(2) = n(1)

s(5) = n(2)
s(8) = n(3)

clear_line n(1), n(2)
draw_line n(2), n(3), j
draw_line n(3), n(4), j
draw_line n(4), n(1), j

'------------------------------------

n(1) = n(3)
n(2).X = n(1).X: n(2).Y = v2.Y
n(3).X = n(2).X + distance: n(3).Y = n(2).Y
n(4).X = n(3).X: n(4).Y = n(1).Y

s(6) = v2
s(7) = n(2)

s(9) = n(4)
s(10) = n(3)

draw_line n(1), n(2), j
clear_line n(2), n(3)
draw_line n(3), n(4), j
draw_line n(4), n(1), j

'------------------------------------

n(2) = n(4)
n(3).X = v3.X: n(3).Y = n(2).Y
n(4).X = n(3).X: n(4).Y = n(3).Y - distance
n(1).X = n(2).X: n(1).Y = n(4).Y

s(11) = v3
s(12) = n(3)

s(14) = n(1)
s(15) = n(4)

draw_line n(1), n(2), j
draw_line n(2), n(3), j
clear_line n(3), n(4)
draw_line n(4), n(1), j

'------------------------------------

n(3) = n(1)
n(4).X = n(3).X: n(4).Y = n(3).Y - distance
n(1).X = n(4).X - distance: n(1).Y = n(4).Y
n(2).X = n(1).X: n(2).Y = n(3).Y

s(16) = v4
s(13) = n(4)

s(3) = n(2)
s(4) = n(1)

draw_line n(1), n(2), j
draw_line n(2), n(3), j
draw_line n(3), n(4), j
clear_line n(4), n(1)

'------------------------------------

If iterations - 1 > 0 Then
    draw_square s(1), s(2), s(3), s(4), iterations - 1
    draw_square s(5), s(6), s(7), s(8), iterations - 1
    draw_square s(9), s(10), s(11), s(12), iterations - 1
    draw_square s(13), s(14), s(15), s(16), iterations - 1
    draw_square s(3), s(8), s(9), s(14), iterations - 1
End If




End Sub





Private Sub Command1_Click()
frmTree.Show
End Sub

Private Sub cmdMenu_Click()
frmStartup.Show
frmSquares.Hide
End Sub
