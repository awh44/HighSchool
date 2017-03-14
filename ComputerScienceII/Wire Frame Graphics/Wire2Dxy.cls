VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Wire2Dxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: Wire2Dxy
' Defining wireframe objects using the
' cartesian coordinate system (x,y)

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

Private lines(2000) As line
Private points(1000) As point

Private line_count As Integer
Private point_count As Integer

Private x_origin As Integer
Private y_origin As Integer
Private pixelX As Integer
Private pixelY As Integer


Public Sub init_square()
points(0).X = 5: points(0).Y = 5
points(1).X = -5: points(1).Y = 5
points(2).X = -5: points(2).Y = -5
points(3).X = 5: points(3).Y = -5
point_count = 4

lines(0).p1 = 0: lines(0).p2 = 1
lines(1).p1 = 1: lines(1).p2 = 2
lines(2).p1 = 2: lines(2).p2 = 3
lines(3).p1 = 3: lines(3).p2 = 0


line_count = 4
x_origin = frmWire2D.pctWire2D.ScaleWidth / 2
y_origin = frmWire2D.pctWire2D.ScaleHeight / 2

pixelX = frmWire2D.pctWire2D.ScaleWidth / 20
pixelY = frmWire2D.pctWire2D.ScaleHeight / 20


End Sub

Public Sub Draw()

Dim c As Integer
Dim x1, y1, x2, y2 As Single ' cartesian coordinates
Dim xx1, yy1, xx2, yy2 As Single ' pixel coordinates


For c = 0 To line_count - 1

    x1 = points(lines(c).p1).X
    y1 = points(lines(c).p1).Y
    x2 = points(lines(c).p2).X
    y2 = points(lines(c).p2).Y

    ' Draw the line
    
    xx1 = x_origin + x1 * pixelX
    yy1 = y_origin - y1 * pixelY
    
    xx2 = x_origin + x2 * pixelX
    yy2 = y_origin - y2 * pixelY
    
    frmWire2D.pctWire2D.Line (xx1, yy1)-(xx2, yy2)
    

Next c

End Sub


Public Sub rotate()
Dim angle, radians As Single
Dim c As Integer
Dim tempx, tempy As Single

angle = Val(frmWire2D.txtAngle)

angle = angle + Val(frmWire2D.txtInc)

radians = angle * 3.14159 / 180

For c = 0 To point_count - 1
    tempx = points(c).X: tempy = points(c).Y
    points(c).X = tempx * Cos(radians) - tempy * Sin(radians)
    points(c).Y = tempx * Sin(radians) + tempy * Cos(radians)
    
Next c

Draw

'frmWire2D.txtAngle = angle

End Sub




