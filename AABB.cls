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
Public center As XYpoint
Public halfDimension As Double

Private Sub Class_Initialize()
    Set center = New XYpoint
End Sub

' @description returns true if the bounding box contains the point. Containment is true for all interior points. Points on the bottom and right boundaries (i.e. at the max X and max Y values) are included. Points on the top and left boundaries (i.e. at the min X and min Y values) are not included.
Public Function contains(ByRef p As XYpoint) As Boolean
    Dim xMin As Double, xMax As Double, yMin As Double, yMax As Double
    xMin = center.x - halfDimension
    xMax = center.x + halfDimension
    yMin = center.y - halfDimension
    yMax = center.y + halfDimension
    
    If xMin >= p.x Then
        Exit Function
    ElseIf xMax < p.x Then
        Exit Function
    ElseIf yMin >= p.y Then
        Exit Function
    ElseIf yMax < p.y Then
        Exit Function
    End If
    
    contains = True
    
End Function

Public Function intersects(ByRef box As AABB) As Boolean
    Dim dist As Double
    dist = halfDimension + boxhalfDimension
    
    If Abs(center.x - box.center.x) >= dist Then
        Exit Function 'x-dimension doesn't overlap
    End If
    
    If Abs(center.y - box.center.y) >= dist Then
        Exit Function 'y-dimension doesn't overlap
    End If
    
    'must be < `dist` away in both x- and y-dimensions.. they intersect
    intersects = True
    
End Function
