VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AABB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'axis-aligned bounding box with half dimension and center
Public center As XYPoint
Public halfDimension As Double

Private Sub Class_Initialize()
    Set center = New XYPoint
End Sub

' @description returns true if the bounding box contains the point. Containment is true for all interior points. Points on the bottom and right boundaries (i.e. at the max X and max Y values) are included. Points on the top and left boundaries (i.e. at the min X and min Y values) are not included.
Public Function contains(ByRef P As XYPoint) As Boolean
    Dim xMin As Double, xMax As Double, yMin As Double, yMax As Double
    xMin = center.X - halfDimension
    xMax = center.X + halfDimension
    yMin = center.Y - halfDimension
    yMax = center.Y + halfDimension
    
    If xMin >= P.X Then
        Exit Function
    ElseIf xMax < P.X Then
        Exit Function
    ElseIf yMin >= P.Y Then
        Exit Function
    ElseIf yMax < P.Y Then
        Exit Function
    End If
    
    contains = True
    
End Function

Public Function intersects(ByRef box As AABB) As Boolean
    Dim dist As Double
    dist = halfDimension + boxhalfDimension
    
    If Abs(center.X - box.center.X) >= dist Then
        Exit Function 'x-dimension doesn't overlap
    End If
    
    If Abs(center.Y - box.center.Y) >= dist Then
        Exit Function 'y-dimension doesn't overlap
    End If
    
    'must be < `dist` away in both x- and y-dimensions.. they intersect
    intersects = True
    
End Function
