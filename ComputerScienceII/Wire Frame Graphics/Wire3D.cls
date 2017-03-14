VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Wire3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Austin Herring
'2-14-11
'projection: (x, y, z) to (x, y)

Private Type point
    X As Single
    Y As Single
    Z As Single
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
Private lines(5000) As line
Private points(5000) As point

Private line_count As Integer
Private point_count As Integer

Private x_origin As Single
Private y_origin As Single
Private pixelX As Single
Private pixelY As Single

Public Sub initialize()

points(0).X = 5: points(0).Y = 5: points(0).Z = -5
points(1).X = -5: points(1).Y = 5: points(1).Z = -5
points(2).X = -5: points(2).Y = -5: points(2).Z = -5
points(3).X = 5: points(3).Y = -5: points(3).Z = -5

points(4).X = 5: points(4).Y = 5: points(4).Z = 5
points(5).X = -5: points(5).Y = 5: points(5).Z = 5
points(6).X = -5: points(6).Y = -5: points(6).Z = 5
points(7).X = 5: points(7).Y = -5: points(7).Z = 5

point_count = 8


'Front square-----------------
lines(0).p1 = 0: lines(0).p2 = 1
lines(0).clr.r = 0: lines(0).clr.g = 0: lines(0).clr.b = 0
    
lines(1).p1 = 1: lines(1).p2 = 2
lines(1).clr.r = 0: lines(1).clr.g = 0: lines(1).clr.b = 0
    
lines(2).p1 = 2: lines(2).p2 = 3
lines(2).clr.r = 0: lines(2).clr.g = 0: lines(2).clr.b = 0
    
lines(3).p1 = 3: lines(3).p2 = 0
lines(3).clr.r = 0: lines(3).clr.g = 0: lines(3).clr.b = 0
'-----------------------------

'Back square------------------
lines(4).p1 = 4: lines(4).p2 = 5
lines(4).clr.r = 255: lines(0).clr.g = 0: lines(0).clr.b = 0

lines(5).p1 = 5: lines(5).p2 = 6
lines(5).clr.r = 255: lines(0).clr.g = 0: lines(0).clr.b = 0

lines(6).p1 = 6: lines(6).p2 = 7
lines(6).clr.r = 255: lines(0).clr.g = 0: lines(0).clr.b = 0

lines(7).p1 = 7: lines(7).p2 = 4
lines(7).clr.r = 255: lines(0).clr.g = 0: lines(0).clr.b = 0
'-----------------------------

'Left side--------------------
lines(8).p1 = 1: lines(8).p2 = 5
lines(8).clr.b = 255: lines(0).clr.g = 0: lines(0).clr.r = 0

lines(9).p1 = 2: lines(9).p2 = 6
lines(9).clr.b = 255: lines(0).clr.g = 0: lines(0).clr.r = 0
'-----------------------------

'Right side--------------------
lines(10).p1 = 0: lines(10).p2 = 4
lines(10).clr.b = 255: lines(0).clr.g = 0: lines(0).clr.r = 0

lines(11).p1 = 3: lines(11).p2 = 7
lines(11).clr.b = 255: lines(0).clr.g = 0: lines(0).clr.r = 0
'-----------------------------

'"X"--------------------------
lines(12).p1 = 4: lines(12).p2 = 6
lines(12).clr.r = 255: lines(0).clr.g = 0: lines(0).clr.b = 0

lines(13).p1 = 5: lines(13).p2 = 7
lines(13).clr.r = 255: lines(0).clr.g = 0: lines(0).clr.b = 0
'-----------------------------

line_count = 14

re_size
rotate_x
rotate_y
rotate_z
translate

draw

End Sub

Public Sub init_cylinder()

Dim c As Integer
Dim theta, radius, line_c As Integer
Dim radians, angle As Single

line_count = 0

theta = 15
radius = 3
radians = theta * 3.141592654 / 180
point_count = (360 / theta) * 2

angle = 0

For c = 0 To point_count - 1
    points(c).X = radius * Cos(angle)
    points(c).Z = radius * Sin(angle)
    points(c).Y = -5
    
    angle = angle + radians
Next c

angle = 0

For c = 0 To point_count - 1
    points(c + point_count).X = radius * Cos(angle)
    points(c + point_count).Z = radius * Sin(angle)
    points(c + point_count).Y = 5
    
    angle = angle + radians
Next c

For c = 0 To point_count - 2
    lines(c).p1 = c
    lines(c).p2 = c + 1
    
    line_count = line_count + 1
    lines(c).clr.b = 255: lines(c).clr.r = 0: lines(c).clr.g = 0
Next c

For c = 0 To point_count - 2
    lines(c + point_count).p1 = c + point_count
    lines(c + point_count).p2 = c + 1 + point_count
    
    line_count = line_count + 1
    
    lines(c + point_count).clr.b = 255: lines(c + point_count).clr.r = 0: lines(c + point_count).clr.g = 0
Next c

line_c = line_count - 1

For c = 0 To point_count - 1
    lines(c + line_c).p1 = c
    lines(c + line_c).p2 = c + point_count
    
    line_count = line_count + 1
    
    lines(c + line_c).clr.r = 255: lines(c + line_c).clr.b = 255: lines(c + line_c).clr.g = 0
Next c

point_count = point_count * 2

re_size
rotate_x
rotate_y
rotate_z
translate

draw
    
    

End Sub

Public Sub init_cone()

Dim c As Integer
Dim theta, radius As Integer
Dim radians, angle As Single

line_count = 0

theta = 15
radius = 4
radians = theta * 3.141592654 / 180
point_count = (360 / theta) + 2

angle = 0

points(0).X = 0
points(0).Y = 5
points(0).Z = 0

For c = 1 To point_count - 1
    points(c).X = radius * Cos(angle)
    points(c).Z = radius * Sin(angle)
    points(c).Y = -5
    
    angle = angle + radians
    
Next c

For c = 1 To point_count - 1
    lines(c - 1).p1 = 0
    lines(c - 1).p2 = c
    line_count = line_count + 1
    
    lines(c).clr.r = 255: lines(c).clr.b = 255: lines(c).clr.g = 0
Next c

For c = point_count To (point_count - 1) * 2
    lines(c - 1).p1 = c - point_count
    lines(c - 1).p2 = c - point_count + 1
    line_count = line_count + 1
    
    lines(c).clr.b = 255: lines(c).clr.r = 0: lines(c).clr.g = 0
Next c

re_size
rotate_x
rotate_y
rotate_z
translate

draw


End Sub


Public Sub draw()

Dim X(1), Y(1), Z(1) As Single
Dim x1(1), y1(1), x2(1), y2(1) As Single
Dim j As Integer
Dim c As color

x_origin = frmWire3D.pctWire3D.ScaleWidth / 2
y_origin = frmWire3D.pctWire3D.ScaleHeight / 2

pixelX = frmWire3D.pctWire3D.ScaleWidth / 20
pixelY = frmWire3D.pctWire3D.ScaleHeight / 20


For j = 0 To line_count - 1

    

    X(0) = points(lines(j).p1).X
    Y(0) = points(lines(j).p1).Y
    Z(0) = points(lines(j).p1).Z

    X(1) = points(lines(j).p2).X
    Y(1) = points(lines(j).p2).Y
    Z(1) = points(lines(j).p2).Z

    x1(0) = X(0) * (100 / (100 + Z(0)))
    y1(0) = Y(0) * (100 / (100 + Z(0)))

    x2(0) = X(1) * (100 / (100 + Z(1)))
    y2(0) = Y(1) * (100 / (100 + Z(1)))

    x1(1) = x_origin + (x1(0) * pixelX)
    y1(1) = y_origin - (y1(0) * pixelY)

    x2(1) = x_origin + (x2(0) * pixelX)
    y2(1) = y_origin - (y2(0) * pixelY)
    
    c.r = lines(j).clr.r
    c.g = lines(j).clr.g
    c.b = lines(j).clr.b
    
    frmWire3D.pctWire3D.Line (x1(1), y1(1))-(x2(1), y2(1)), RGB(c.r, c.g, c.b)
    
Next j


End Sub

Public Sub rotate_z()

Dim x_new, y_new As Single
Dim c As Integer
Dim theta, radians As Single

If frmWire3D.chkClear.Value = 1 Then frmWire3D.pctWire3D.Cls

theta = frmWire3D.txtRotateZ
radians = theta * (3.141592654 / 180)

For c = 0 To point_count
    x_new = points(c).X * Cos(radians) - points(c).Y * Sin(radians)
    y_new = points(c).X * Sin(radians) + points(c).Y * Cos(radians)
    
    points(c).X = x_new
    points(c).Y = y_new

Next c

End Sub

Public Sub rotate_y()

Dim x_new, z_new As Single
Dim c As Integer
Dim theta, radians As Single

If frmWire3D.chkClear.Value = 1 Then frmWire3D.pctWire3D.Cls

theta = frmWire3D.txtRotateY
radians = theta * (3.141592654 / 180)

For c = 0 To point_count
    x_new = points(c).X * Cos(radians) - points(c).Z * Sin(radians)
    z_new = points(c).X * Sin(radians) + points(c).Z * Cos(radians)
    
    points(c).X = x_new
    points(c).Z = z_new

Next c

End Sub

Public Sub rotate_x()

Dim y_new, z_new As Single
Dim c As Integer
Dim theta, radians As Single

If frmWire3D.chkClear.Value = 1 Then frmWire3D.pctWire3D.Cls

theta = frmWire3D.txtRotateX
radians = theta * (3.141592654 / 180)

For c = 0 To point_count
    y_new = points(c).Y * Cos(radians) - points(c).Z * Sin(radians)
    z_new = points(c).Y * Sin(radians) + points(c).Z * Cos(radians)
    
    points(c).Y = y_new
    points(c).Z = z_new

Next c

End Sub

Public Sub translate()

Dim translateX, translateY, translateZ As Integer


translateX = frmWire3D.txtTranslateX
translateY = frmWire3D.txtTranslateY
translateZ = frmWire3D.txtTranslateZ

For c = 0 To point_count
    points(c).X = points(c).X + translateX
    points(c).Y = points(c).Y + translateY
    points(c).Z = points(c).Z + translateZ
Next c

End Sub

Public Sub re_size()

Dim scl As Single
Dim x_dist, y_dist, z_dist As Single
Dim i As Integer

scl = frmWire3D.txtScale

For i = 0 To point_count

    x_dist = points(i).X
    y_dist = points(i).Y
    z_dist = points(i).Z

    x_dist = x_dist * scl
    y_dist = y_dist * scl
    z_dist = z_dist * scl
    
    points(i).X = x_dist
    points(i).Y = y_dist
    points(i).Z = z_dist
    
Next i



End Sub

Public Sub init_sphere()



Dim c, j As Integer
Dim rho As Integer
Dim theta, phi As Single
Dim t_count, p_count As Single

line_count = 0
point_count = 360 / 10

rho = 6
theta = 10 * (3.141592654 / 180)
phi = 10 * (3.141592654 / 180)


For j = 0 To point_count - 1

    For c = 0 To point_count - 1
        points(c + j * point_count).X = rho * Sin(p_count) * Cos(t_count)
        points(c + j * point_count).Y = rho * Sin(p_count) * Sin(t_count)
        points(c + j * point_count).Z = rho * Cos(p_count)

        p_count = p_count + phi
        
'        If c <> point_count - 1 Then
'            lines(c + j * point_count).p1 = c + j * point_count
'            lines(c + j * point_count).p2 = c + j * point_count + 1
'        End If
    Next c

    p_count = 0
    t_count = t_count + theta

Next j

t_count = 0
p_count = 0

For j = 0 To point_count - 1

    For c = point_count ^ 2 To (point_count ^ 2 * 2) - 1
        points(c + j * point_count).X = rho * Sin(p_count) * Cos(t_count)
        points(c + j * point_count).Y = rho * Sin(p_count) * Sin(t_count)
        points(c + j * point_count).Z = rho * Cos(p_count)

        t_count = t_count + theta
        
'        If c <> (point_count ^ 2 * 2) - 1 Then
'            lines(c + j * point_count).p1 = c + j * point_count
'            lines(c + j * point_count).p2 = c + j * point_count + 1
'        End If
    Next c

    t_count = 0
    p_count = p_count + phi

Next j


For c = 0 To point_count ^ 2 * 2 - 1

    If (c + 1) Mod 36 <> 0 Then
        lines(c).p1 = c
        lines(c).p2 = c + 1

        line_count = line_count + 1
    End If
Next c

'For j = 0 To point_count - 1
'
'    For c = 0 To point_count - 1
'        lines(c + j * point_count).p1 = c + j * point_count
'        lines(c + j * point_count).p2 = c + 1 + j * point_count
'    Next c
'
'Next j

point_count = point_count ^ 2 * 2

re_size
rotate_x
rotate_y
rotate_z
translate


draw


End Sub

Public Sub plot()

Dim X(1), Y(1), Z(1) As Single
Dim x1(1), y1(1), x2(1), y2(1) As Single
Dim c As Integer


x_origin = frmWire3D.pctWire3D.ScaleWidth / 2
y_origin = frmWire3D.pctWire3D.ScaleHeight / 2

pixelX = frmWire3D.pctWire3D.ScaleWidth / 20
pixelY = frmWire3D.pctWire3D.ScaleHeight / 20

For c = 0 To point_count - 1

    X(0) = points(c).X
    Y(0) = points(c).Y
    Z(0) = points(c).Z

    x1(0) = X(0) * (100 / (100 + Z(0)))
    y1(0) = Y(0) * (100 / (100 + Z(0)))

    x1(1) = x_origin + (x1(0) * pixelX)
    y1(1) = y_origin - (y1(0) * pixelY)

    frmWire3D.pctWire3D.PSet (x1(1), y1(1))
Next c

End Sub



