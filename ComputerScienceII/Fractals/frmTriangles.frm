VERSION 5.00
Begin VB.Form frmTriangles 
   Caption         =   "Triangles"
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
      Left            =   7440
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      ForeColor       =   &H00000000&
      Height          =   8415
      Left            =   120
      ScaleHeight     =   557
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   557
      TabIndex        =   3
      Top             =   720
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
      Left            =   2880
      TabIndex        =   2
      Top             =   240
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
      Top             =   240
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
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmTriangles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Austin Herring
'More fracals: two triangles


Private Type Vertex
    X As Integer
    Y As Integer
End Type

Dim v(3) As Vertex
Dim iteration_count As Integer


Private Function midpoint(p1 As Vertex, p2 As Vertex) As Vertex

midpoint.X = (p1.X + p2.X) / 2
midpoint.Y = (p1.Y + p2.Y) / 2

End Function

Private Sub draw_line(p1 As Vertex, p2 As Vertex)

    Picture1.Line (p1.X, p1.Y)-(p2.X, p2.Y)

End Sub

Private Sub plot_point(v As Vertex)

    Picture1.PSet (v.X, v.Y), RGB(0, 0, 0)
    
End Sub


Private Sub cmdDraw_Click()

Picture1.Cls

v(1).X = Picture1.ScaleWidth / 2: v(1).Y = 50
v(2).X = 50: v(2).Y = Picture1.ScaleHeight / 2
v(3).X = Picture1.ScaleWidth - 50: v(3).Y = v(2).Y

draw_line v(1), v(2)
draw_line v(2), v(3)
draw_line v(3), v(1)

If txtLevels - 1 > 0 Then
    draw v(1), v(2), v(3), txtLevels - 1
End If

v(1).X = Picture1.ScaleWidth / 2: v(1).Y = Picture1.ScaleHeight - 50
v(2).X = 50: v(2).Y = Picture1.ScaleHeight / 2
v(3).X = Picture1.ScaleWidth - 50: v(3).Y = v(2).Y

draw_line v(1), v(2)
draw_line v(1), v(3)

If txtLevels - 1 > 0 Then
    draw v(1), v(2), v(3), txtLevels - 1
End If


End Sub


Private Function draw(v1 As Vertex, v2 As Vertex, v3 As Vertex, iterations As Integer)

'numbering goes counter-counterclockwise

Dim m(3) As Vertex

m(1) = midpoint(v1, v2)
m(2) = midpoint(v2, v3)
m(3) = midpoint(v3, v1)

draw_line m(1), m(2)
draw_line m(2), m(3)
draw_line m(3), m(1)

If iterations - 1 > 0 Then
    draw m(1), v2, m(2), iterations - 1
    draw v1, m(1), m(3), iterations - 1
    draw m(3), m(2), v3, iterations - 1
    draw m(1), m(2), m(3), iterations - 1
End If

End Function


Private Sub Command1_Click()
frmTree.Show
End Sub

Private Sub cmdMenu_Click()
frmStartup.Show
frmTriangles.Hide
End Sub

Private Sub txtLevels_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
    cmdDraw_Click
End If

End Sub

