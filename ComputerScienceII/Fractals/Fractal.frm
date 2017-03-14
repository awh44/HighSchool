VERSION 5.00
Begin VB.Form frmSierpinski 
   Caption         =   "Sierpinksi Triangles"
   ClientHeight    =   9165
   ClientLeft      =   2310
   ClientTop       =   1305
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   ScaleHeight     =   611
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   703
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Menu"
      Height          =   375
      Left            =   8640
      TabIndex        =   5
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmdChaos 
      Caption         =   "Chaos / Random"
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
      TabIndex        =   4
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton cmdDraw 
      Caption         =   "Draw / Recursive"
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
      TabIndex        =   3
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
      Left            =   1200
      TabIndex        =   1
      Text            =   "4"
      Top             =   240
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      ForeColor       =   &H00000000&
      Height          =   8295
      Left            =   120
      ScaleHeight     =   549
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   677
      TabIndex        =   0
      Top             =   720
      Width           =   10215
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
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmSierpinski"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Austin Herring
'frmSierpinksi
'12-6-10

Private Type Vertex
    X As Integer
    Y As Integer
End Type

Private Type Color
    red As Integer
    green As Integer
    blue As Integer
End Type

Dim v(3) As Vertex

Private Function midpoint(p1 As Vertex, p2 As Vertex) As Vertex

midpoint.X = (p1.X + p2.X) / 2
midpoint.Y = (p1.Y + p2.Y) / 2

End Function

Private Sub draw_line(p1 As Vertex, p2 As Vertex, c As Color)

    frmSierpinski.Picture1.Line (p1.X, p1.Y)-(p2.X, p2.Y), RGB(c.red, c.green, c.blue)
    

End Sub

Private Sub plot_point(v As Vertex, c As Color)

    Picture1.PSet (v.X, v.Y), RGB(c.red, c.green, c.blue)
    
End Sub

Private Function distance(p1 As Vertex, p2 As Vertex) As Single

distance = Sqr((p2.X - p1.X) ^ 2 + (p2.Y - p1.Y) ^ 2)


End Function

Private Sub cmdChaos_Click()

Dim number As Integer
Dim X As Single    ' loop counter
Dim point As Vertex
Dim c As Color

v(1).X = Picture1.Width / 2: v(1).Y = 10
v(2).X = 50:    v(2).Y = Picture1.Height - 10
v(3).X = Picture1.Width - 50: v(3).Y = v(2).Y

Picture1.drawwidth = 3

Picture1.PSet (v(1).X, v(1).Y)
Picture1.PSet (v(2).X, v(2).Y)
Picture1.PSet (v(3).X, v(3).Y)


point.X = midpoint(v(2), v(3)).X
point.Y = midpoint(v(1), v(2)).Y

c.red = 0: c.blue = 0: c.green = 0
plot_point point, c


For X = 1 To 10000
    Randomize
    number = Int(Rnd * 3 + 1)
    Label2 = number

    point = midpoint(point, v(number))
    
    If distance(point, v(1)) < distance(point, v(2)) And distance(point, v(1)) < distance(point, v(3)) Then
        c.red = 255
        c.blue = 0
        c.green = 0
    End If
    
    If distance(point, v(2)) < distance(point, v(1)) And distance(point, v(2)) < distance(point, v(3)) Then
        c.red = 0
        c.blue = 255
        c.green = 0
    End If
    
    If distance(point, v(3)) < distance(point, v(1)) And distance(point, v(3)) < distance(point, v(2)) Then
        c.red = 0
        c.blue = 0
        c.green = 255
    End If
    
    plot_point point, c
Next X

End Sub

Private Sub draw_triangle(v1 As Vertex, v2 As Vertex, v3 As Vertex, c As Color, iteration As Integer)

Dim m(3) As Vertex

'numbering goes counter-clockwise

m(1).X = midpoint(v1, v2).X
m(1).Y = midpoint(v1, v2).Y

m(2).X = midpoint(v2, v3).X
m(2).Y = midpoint(v2, v3).Y

m(3).X = midpoint(v1, v3).X
m(3).Y = midpoint(v1, v3).Y

draw_line m(1), m(2), c
draw_line m(2), m(3), c
draw_line m(3), m(1), c

If iteration - 1 > 0 Then
    draw_triangle v1, m(1), m(3), c, iteration - 1
    draw_triangle m(1), v2, m(2), c, iteration - 1
    draw_triangle m(3), m(2), v3, c, iteration - 1
End If

End Sub


Private Sub draw(v1 As Vertex, v2 As Vertex, v3 As Vertex, iterations As Integer)

Dim m(3) As Vertex  ' 3 midpoints
Dim c As Color

c.red = 0 'Int(256 * Rnd)
c.blue = 0 ' Int(256 * Rnd)
c.green = 0 '  Int(256 * Rnd)

m(1) = midpoint(v1, v2)
m(2) = midpoint(v2, v3)
m(3) = midpoint(v1, v3)

draw_line m(1), m(2), c
draw_line m(2), m(3), c
draw_line m(3), m(1), c

If iterations > 0 Then

    draw m(1), v2, m(2), iterations - 1
    draw v1, m(1), m(3), iterations - 1
   
    draw m(3), m(2), v3, iterations - 1
    c.blue = c.blue + 30
    c.green = c.green + 30
    c.red = c.red + 30
    
    
End If


End Sub
Private Sub cmdDraw_Click()

Picture1.Cls

Dim c As Color
Dim levels As Integer

c.red = 0: c.green = 0: c.blue = 0

levels = Val(txtLevels)
' draw the first triangle

v(1).X = Picture1.ScaleWidth / 2: v(1).Y = 0
v(2).X = Picture1.ScaleWidth: v(2).Y = Picture1.ScaleHeight - 5
v(3).X = 0: v(3).Y = Picture1.ScaleHeight - 5

draw_line v(1), v(2), c
draw_line v(2), v(3), c
draw_line v(3), v(1), c


draw v(1), v(2), v(3), levels



End Sub

Private Sub Command1_Click()
frmTree.Show
End Sub

Private Sub cmdMenu_Click()
frmStartup.Show
frmSierpinski.Hide
End Sub
