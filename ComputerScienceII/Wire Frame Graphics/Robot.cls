VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Robot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Austin Herring
'Robot module

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
    p1 As Integer
    p2 As Integer
    clr As color
End Type

Private lines(1000) As line
Private points(5000) As point

Private line_count As Integer
Private point_count As Integer

Private x_origin As Single
Private y_origin As Single
Private pixelX As Single
Private pixelY As Single


Public Sub draw()



End Sub

Private Sub draw_base()

Dim theta As Integer
Dim radius As Integer
Dim radians, angle As Single

line_count = 0

theta = 45
radius = 4
radians = theta * 3.141592654 / 180
point_count = (360 / theta) * 2

For c = 0 To point_count - 1
    points(c).X = radius * Cos(angle)
    points(c).Z = radius * Sin(angle)
    points(c).Y = -5
    
    angle = angle + radians
Next c





End Sub
