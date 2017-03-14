VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Wire2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Austin Herring
'A class module to define 2-dimensional wireframe objects
'1-25-11

'Data:
'   Points
'   Lines
'   Number of points
'   Number of lines
'   Color, etc.

'Functions:
'   Draw
'   Rotate
'   Scale
'   Translate
'   "Initialize" - generate all points and lines


'Points:
'   p(0).x = 100
'   p(0).y = 100
'
'   p(0).x = 150
'   p(0).y = 100

'   p(0).x = 150
'   p(0).y = 150

'   p(0).x = 100
'   p(0).y = 150

'Lines:
'   l(0).p1 = 0
'   l(0).p2 = 0
'   l(0).red = 255
'   l(0).green = 255
'   l(0).blue = 255

'Npoints = 4
'Nlines = 6

Private Type point
    X As Single
    Y As Single
End Type

Private Type color
    r As Integer
    g As Integer
    b As Integer
End Type

Private Type line
    p1 As Integer   'point number of first point
    p2 As Integer   'point number of second point
    clr As color
End Type

'Data definition: In VB, properties of an object
Private pivot As point
Private lines(2000) As line
Private points(1000) As point

Private line_count As Integer
Private point_count As Integer

Private x_origin As Integer
Private y_origin As Integer
Private pixelX As Integer
Private pixelY As Integer


Public Sub initialize()

If frmWire2D.optSquare = True Then

    Dim length As Single
    Dim scl As Single
    Dim translateX, translateY As Single
    
    pixelX = frmWire2D.pctWire2D.ScaleWidth / 20
    pixelY = frmWire2D.pctWire2D.ScaleHeight / 20
    
    x_origin = frmWire2D.pctWire2D.ScaleWidth / 2
    y_origin = frmWire2D.pctWire2D.ScaleHeight / 2
    
    'scl = frmWire2D.txtScale
    
    'translateX = (frmWire2D.pctWire2D.Width / 20) * frmWire2D.txtTranslateX
    'translateY = (frmWire2D.pctWire2D.Height / 20) * frmWire2D.txtTranslateY
    
    'length = (frmWire2D.txtLength * scl) / 2
    'length = (frmWire2D.txtLength) / 2
    
    'pivot.X = (frmWire2D.pctWire2D.ScaleWidth / 2 - x_origin) / pixelX
    'pivot.Y = (frmWire2D.pctWire2D.ScaleHeight / 2 - y_origin) / pixelY
    
    pivot.X = 0
    pivot.Y = 0
    
    'points(0).X = frmWire2D.pctWire2D.ScaleWidth / 2 - length + translateX: points(0).Y = frmWire2D.pctWire2D.ScaleHeight / 2 - length + translateY
    'points(1).X = frmWire2D.pctWire2D.ScaleWidth / 2 + length + translateX: points(1).Y = frmWire2D.pctWire2D.ScaleHeight / 2 - length + translateY
    'points(2).X = frmWire2D.pctWire2D.ScaleWidth / 2 + length + translateX: points(2).Y = frmWire2D.pctWire2D.ScaleHeight / 2 + length + translateY
    'points(3).X = frmWire2D.pctWire2D.ScaleWidth / 2 - length + translateX: points(3).Y = frmWire2D.pctWire2D.ScaleHeight / 2 + length + translateY
    
    points(0).X = 5: points(0).Y = 5
    points(1).X = -5: points(1).Y = 5
    points(2).X = -5: points(2).Y = -5
    points(3).X = 5: points(3).Y = -5
    
    point_count = 4
    
    lines(0).p1 = 0: lines(0).p2 = 1
    lines(0).clr.r = 255: lines(0).clr.g = 0: lines(0).clr.b = 0
    
    lines(1).p1 = 1: lines(1).p2 = 2
    lines(1).clr.r = 0: lines(1).clr.g = 255: lines(1).clr.b = 0
    
    lines(2).p1 = 2: lines(2).p2 = 3
    lines(2).clr.r = 0: lines(2).clr.g = 0: lines(2).clr.b = 255
    
    lines(3).p1 = 3: lines(3).p2 = 0
    lines(3).clr.r = 255: lines(3).clr.g = 255: lines(3).clr.b = 0
    
    lines(4).p1 = 0: lines(4).p2 = 2
    lines(4).clr.r = 255: lines(4).clr.g = 0: lines(4).clr.b = 255
    
    lines(5).p1 = 1: lines(5).p2 = 3
    lines(5).clr.r = 0: lines(5).clr.g = 255: lines(5).clr.b = 255
    
    line_count = 6
    
    
    re_size
    rotate
    translate
    
    
    draw frmWire2D.pctWire2D
    
End If  'initialize the square

End Sub

Public Sub init_circle(theta As Single, radius As Integer)

Dim c As Integer
Dim angle, radians As Single
Dim zero As Boolean
Dim c_count As Integer
Dim c_int As Single

line_count = 0
point_count = 0

pixelX = frmWire2D.pctWire2D.ScaleWidth / 20
pixelY = frmWire2D.pctWire2D.ScaleHeight / 20
    
x_origin = frmWire2D.pctWire2D.ScaleWidth / 2
y_origin = frmWire2D.pctWire2D.ScaleHeight / 2

pivot.X = 0
pivot.Y = 0

points(0).X = 0: points(0).Y = 0

point_count = (360 / theta) + 1

radians = theta * (3.141592654 / 180)
angle = 0

c_int = 255 / (point_count - 1)

For c = 1 To point_count

    points(c).X = radius * Cos(angle)
    points(c).Y = radius * Sin(angle)
    
    angle = angle + radians

Next c

For c = 1 To point_count - 1
    lines(c).p1 = c
    lines(c).p2 = c + 1
    
    If zero = False Then
        lines(c).clr.r = 255 - 2 * (c * c_int)
    End If
    
    If zero = True Then
        lines(c).clr.r = (c - c_count) * c_int * 2
    End If
    
    If lines(c).clr.r <= 0 Then
        zero = True
        c_count = c
        If lines(c).clr.r < 0 Then lines(c).clr.r = 0
    End If
    
'    If zero = False Then
'        lines(c).clr.r = 255 - 5 * c
'    End If
'
'    If lines(c).clr.r = 0 Then
'        zero = True
'        c_count = c
'    End If
'
'    If zero = True Then
'        lines(c).clr.r = 5 * (c - c_count)
'    End If
'
'    If lines(c).clr.r = 255 Then
'        zero = False
'    End If
    
    line_count = line_count + 1
    
Next c

zero = False

For c = (point_count) To (2 * point_count)
    lines(c).p1 = 0
    lines(c).p2 = c - point_count
    
    If zero = False Then
        lines(c).clr.r = 255 - 2 * ((c - point_count) * c_int)
    End If
    
    If zero = True Then
        lines(c).clr.r = (c - c_count) * c_int * 2
    End If
    
    If lines(c).clr.r <= 0 Then
        zero = True
        c_count = c
        If lines(c).clr.r < 0 Then lines(c).clr.r = 0
    End If
    
'    If zero = False Then
'        lines(c).clr.r = 255 - 5 * (c - point_count)
'    End If
'
'    If lines(c).clr.r = 0 Then
'        zero = True
'        c_count = c
'    End If
'
'    If zero = True Then
'        lines(c).clr.r = 5 * (c - c_count)
'    End If
'
'    If lines(c).clr.r = 255 Then
'        zero = False
'    End If

    line_count = line_count + 1
Next c

re_size
rotate
translate


draw frmWire2D.pctWire2D



End Sub


Public Sub draw(ByRef pct As PictureBox)

' Line( , ) - ( , ), RGB( , , )

Dim x1, y1, x2, y2 As Single
Dim c As color


'For j = 0 To line_count - 1
'
'
'    c.r = lines(j).clr.r
'    c.g = lines(j).clr.g
'    c.b = lines(j).clr.b
'
'    x1 = points(lines(j).p1).X: y1 = points(lines(j).p1).Y
'    x2 = points(lines(j).p2).X: y2 = points(lines(j).p2).Y
'
'    frmWire2D.pctWire2D.Line (x1, y1)-(x2, y2), RGB(c.r, c.g, c.b)
'
'Next j

If frmWire2D.chkClear.Value = 1 Then
    frmWire2D.pctWire2D.Cls
End If

For j = 0 To line_count - 1

    x1 = x_origin + points(lines(j).p1).X * pixelX
    y1 = y_origin - points(lines(j).p1).Y * pixelY
    
    x2 = x_origin + points(lines(j).p2).X * pixelX
    y2 = y_origin - points(lines(j).p2).Y * pixelY
    
    c.r = lines(j).clr.r
    c.g = lines(j).clr.g
    c.b = lines(j).clr.b
    
    pct.Line (x1, y1)-(x2, y2), RGB(c.r, c.g, c.b)
    
Next j

End Sub


Public Sub re_size()

Dim i As Integer
Dim x_dist, y_dist As Single
Dim scl As Single

scl = frmWire2D.txtScale

'For i = 0 To point_count - 1
'
'    x_dist = points(i).X - pivot.X
'    y_dist = points(i).Y - pivot.Y
'
'    x_dist = x_dist * scl
'    y_dist = y_dist * scl
'
'    points(i).X = pivot.X + x_dist
'    points(i).Y = pivot.Y + y_dist
'
'Next i

'frmWire2D.pctWire2D.Cls

'If frmWire2D.optSquare = True Then
'    initialize 0
'End If

'If frmWire2D.optCircle = True Then
'    init_circle frmWire2D.txtTheta, frmWire2D.txtRadius
'End If

For i = 0 To point_count

    x_dist = points(i).X - pivot.X
    y_dist = points(i).Y - pivot.Y

    x_dist = x_dist * scl
    y_dist = y_dist * scl
    
    points(i).X = pivot.X + x_dist
    points(i).Y = pivot.Y + y_dist
    
Next i

'draw frmWire2D.pctWire2D

End Sub


Public Sub translate()

Dim translateX, translateY As Integer


translateX = frmWire2D.txtTranslateX
translateY = frmWire2D.txtTranslateY

For c = 0 To point_count
    points(c).X = points(c).X + translateX
    points(c).Y = points(c).Y + translateY
Next c
    


End Sub


Public Sub rotate()

Dim x_new, y_new As Single
Dim c As Integer
Dim theta, radians As Single

'If kind = True Then
'    init_circle frmWire2D.txtTheta, frmWire2D.txtRadius
'Else
'    initialize 0
'End If

theta = frmWire2D.txtRotate
radians = theta * (3.141592 / 180)


For c = 0 To point_count
    x_new = points(c).X * Cos(radians) - points(c).Y * Sin(radians)
    y_new = points(c).X * Sin(radians) + points(c).Y * Cos(radians)
    
    points(c).X = x_new
    points(c).Y = y_new

Next c

'draw frmWire2D.pctWire2D






End Sub









