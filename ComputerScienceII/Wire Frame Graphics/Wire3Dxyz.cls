VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Wire3Dxyz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Wire3Dxyz
' Started:  Monday, Dec. 5, 2011

' Class Module: Wire2Dxy
' Defining wireframe objects using the
' cartesian coordinate system (x,y,z)

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

Private lines(2000) As line
Private points(1000) As point

Private line_count As Integer
Private point_count As Integer

Private x_origin As Integer
Private y_origin As Integer
Private pixelX As Integer
Private pixelY As Integer


Public Function init_cube()
' front face
points(0).X = 5: points(0).Y = 5: points(0).Z = -5
points(1).X = 5: points(1).Y = -5: points(1).Z = -5
points(2).X = -5: points(2).Y = -5: points(2).Z = -5
points(3).X = -5: points(3).Y = 5: points(3).Z = -5

' back face
points(4).X = 5: points(4).Y = 5: points(4).Z = 5
points(5).X = 5: points(5).Y = -5: points(5).Z = 5
points(6).X = -5: points(6).Y = -5: points(6).Z = 5
points(7).X = -5: points(7).Y = 5: points(7).Z = 5

' front face
lines(0).p1 = 0: lines(0).p2 = 1: lines(0).clr.r = 0: lines(0).clr.g = 0: lines(0).clr.b = 0
lines(1).p1 = 1: lines(1).p2 = 2: lines(1).clr.r = 0: lines(1).clr.g = 0: lines(1).clr.b = 0
lines(2).p1 = 2: lines(2).p2 = 3: lines(2).clr.r = 0: lines(2).clr.g = 0: lines(2).clr.b = 0
lines(3).p1 = 3: lines(3).p2 = 0: lines(3).clr.r = 0: lines(3).clr.g = 0: lines(3).clr.b = 0

' back face
lines(4).p1 = 4: lines(4).p2 = 5: lines(4).clr.r = 0: lines(4).clr.g = 255: lines(4).clr.b = 0
lines(5).p1 = 5: lines(5).p2 = 6: lines(5).clr.r = 0: lines(5).clr.g = 255: lines(5).clr.b = 0
lines(6).p1 = 6: lines(6).p2 = 7: lines(6).clr.r = 0: lines(6).clr.g = 255: lines(6).clr.b = 0
lines(7).p1 = 7: lines(7).p2 = 4: lines(7).clr.r = 0: lines(7).clr.g = 255: lines(7).clr.b = 0

'front to back
lines(8).p1 = 0: lines(8).p2 = 4: lines(8).clr.r = 255: lines(8).clr.g = 0: lines(8).clr.b = 0
lines(9).p1 = 1: lines(9).p2 = 5: lines(9).clr.r = 255: lines(9).clr.g = 0: lines(9).clr.b = 0
lines(10).p1 = 2: lines(10).p2 = 6: lines(10).clr.r = 255: lines(10).clr.g = 0: lines(10).clr.b = 0
lines(11).p1 = 3: lines(11).p2 = 7: lines(11).clr.r = 255: lines(11).clr.g = 0: lines(11).clr.b = 0

' back diagonals
'
lines(12).p1 = 4: lines(12).p2 = 6: lines(12).clr.r = 0: lines(12).clr.g = 255: lines(12).clr.b = 0
lines(13).p1 = 5: lines(12).p2 = 7: lines(13).clr.r = 0: lines(13).clr.g = 255: lines(13).clr.b = 0

line_count = 14: point_count = 8

x_origin = frmWire3Dxyz.pctWire3D.ScaleWidth / 2
y_origin = frmWire3Dzyz.pctWire3D.ScaleHeight / 2

pixelX = frmWire3Dxyz.pctWire3D.ScaleWidth / 20
pixelY = frmWire3Dxyz.pctWire3D.ScaleHeight / 20
End Function

Public Function draw()

' projection (x1,y1,z1) to (x2,y2)

' x2 = x1 * distance/(distance + z1)
' y2 = y1 * distance/ (distance + z1)

' distance = 50

' (5,5,5)  projected point :  x2 = 5 * 50/(50 + 5) =5 * .909090  = 4.54
' (5,5,-5) projected point :  x2 = 5 * 50/(50 + -5) = 5 * 1.11 = 5.555


End Function
