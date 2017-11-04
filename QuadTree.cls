VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QuadTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

'' A quadtree is a tree data structure in which each internal node has exactly four children. Quadtrees are the two-dimensional analog of octrees and are most often used to partition a two-dimensional space by recursively subdividing it into four quadrants or regions.
'' This implementation acts like a set for storage of two-dimensional data as X-Y coordinates.

'arbitrary constant to indicate how many elements can be stored in thsi quad tree node
Private Const QT_NODE_CAPACITY = 4

'AABB that represents the boundary of the quad-trr
Private boundary As AABB

'Points in this quad tree node
Private points As Collection

'Children
Private northWest As QuadTree
Private northEast As QuadTree
Private southWest As QuadTree
Private southEast As QuadTree

'constructor
Private Sub Class_Initialize()
    Set points = New Collection
    Set boundary = New AABB
End Sub

' @description This method should be used immediately after instantiation of the QuadTree if not using the factor method for creating a QuadTree. This method changes the boundary of this QuadTree to the specified coordinates and size.
' @param centerX the center of the bounding box for this QuadTree
' @param centerY the center of the boundnig box for this QuadTree
' @param halfDimension axis-aligned distance from the center of the bounding box to the edge of the bounding box
Public Sub setBoundary(centerX As Double, centerY As Double, halfDimension As Double)
    Debug.Assert halfDimension > 0
    Set boundary = New AABB
    boundary.center.x = centerX
    boundary.center.y = centerY
    boundary.halfDimension = halfDimension
End Sub

' @description returns the axis-aligned bounding box of this QuadTree
' @return returns the axis-aligned bounding box of this QuadTree
Public Function getBoundary() As AABB
    Set getBoundary = boundary
End Function


Public Sub demo()
    Dim qt As QuadTree, i As Long, halfSize As Double, queryResults As Collection, p As XYpoint
    Set qt = New QuadTree
    halfSize = 400
    qt.setBoundary 400, 400, halfSize
    
    For i = 1 To 500
        qt.insertCoordinate WorksheetFunction.NormInv(Rnd(), 400, halfSize), WorksheetFunction.NormInv(Rnd(), 400, halfSize)
    Next i
    
    qt.drawTree
    
    Set p = New XYpoint
    p.x = Rnd() * halfSize
    p.y = Rnd() * halfSize
    
    Debug.Print "Query Target Center: " & p.getTranslation(-halfSize, -halfSize).toString & vbCrLf & "Query Target Size: " & 2 * (halfSize - 1)
    Set queryResults = qt.queryRangeBoundary(p.x, p.y, halfSize - 1)
    
    For Each p In queryResults
        Debug.Print "  " & p.toString()
    Next p
    
End Sub

' @description inserts a point into the QuadTree
' @param x the x-coodinate of the new point
' @param y the y-coordinate of the new point
' @return returns true if the point was successfully inserted; returns false, otherwise.
Public Function insertCoordinate(x As Double, y As Double) As Boolean
    Dim p As XYpoint
    Set p = New XYpoint
    p.x = x
    p.y = y
    insertCoordinate = insertPoint(p)
End Function

' @description inserts a point into the QuadTree
' @param x the x-coodinate of the new point
' @param y the y-coordinate of the new point
' @return returns true if the point was successfully inserted; returns false, otherwise.
Public Function insertPoint(p As XYpoint) As Boolean

    'ignore objects that do not belong to this quad tree node
    If (Not boundary.contains(p)) Then
        insertPoint = False
        Exit Function
    End If
    
    'if this node has not been subdivided
    If northWest Is Nothing Then
        'check if there is still room at this level of the tree
        If points.count < QT_NODE_CAPACITY Then
            points.add p
            insertPoint = True
            Exit Function
        Else
            'otherwise, subdivide add redistribute the contained points among its children
            subdivide
        End If
    End If
    
    If northWest.insertPoint(p) Then
        insertPoint = True
        Exit Function
    ElseIf northEast.insertPoint(p) Then
        insertPoint = True
        Exit Function
    ElseIf southWest.insertPoint(p) Then
        insertPoint = True
        Exit Function
    ElseIf southEast.insertPoint(p) Then
        insertPoint = True
        Exit Function
    End If
    
    'otherwise, the point cannot be inserted for some unknown reason (should never happen)
    insertPoint = False
    
End Function

' @description a factory method for producing a QuadTree
' @param centerLocation the point at which the QuadTree is centered
' @param halfDimension the distance to the dividing edge of the QuadTree's area of coverage
' @return returns a quadtree with the specified center and height/width
Public Function factory(centerLocation As XYpoint, halfDimension As Double) As QuadTree
    Dim qt As QuadTree, range As AABB
    Set qt = New QuadTree
    qt.setBoundary centerLocation.x, centerLocation.y, halfDimension
    Set factory = qt
End Function

'subdivides this quadtree into 4 child quadtrees and then distributes points according to their XY coordinates
Private Sub subdivide()
    Dim newSize As Double, p As XYpoint
    newSize = boundary.halfDimension / 2#
    
    'prepare child quadtrees
    Set northWest = factory(boundary.center.getTranslation(-newSize, -newSize), newSize)
    Set northEast = factory(boundary.center.getTranslation(newSize, -newSize), newSize)
    Set southWest = factory(boundary.center.getTranslation(-newSize, newSize), newSize)
    Set southEast = factory(boundary.center.getTranslation(newSize, newSize), newSize)
    
    'move points in this quadtree node to the appropriate child node
    For Each p In points
        'if the point p is successfully inserted into any one of the children, then the subsequent IfElse statements are not evaluated. And we continue moving the remaining points
        If northWest.insertPoint(p) Then
        ElseIf northEast.insertPoint(p) Then
        ElseIf southWest.insertPoint(p) Then
        ElseIf southEast.insertPoint(p) Then
        Else
            err.Raise 17, Description:="Illegal state: Unable to find appropriate child node for subdividion."
        End If
    Next p
    
    'since all points now reside in the child quadtree nodes, clear the points Collection
    Set points = New Collection
    
End Sub

Public Function queryRangeBoundary(centerX As Double, centerY As Double, boundaryHalfDimension As Double) As Collection
    Dim range As AABB
    Set range = New AABB
    range.center.x = centerX
    range.center.y = centerY
    range.halfDimension = boundaryHalfDimension
    Set queryRangeBoundary = queryRange(range)
End Function

Public Function queryRange(range As AABB) As Collection
    Dim pointsInRange As Collection, p As XYpoint
    Set pointsInRange = New Collection
    
    'automatically abort if range does not intersect this quad
    If Not boundary.intersects(range) Then
        Set queryRange = pointsInRange
        Exit Function
    End If
    
    'check objects at this quad level
    For Each p In points
        If range.contains(p) Then
            pointsInRange.add p
        End If
    Next p
    
    'terminate here if there are no children
    If northWest Is Nothing Then
        Set queryRange = pointsInRange
        Exit Function
    End If
    
    'otherwise, recursively add points from the children
    Collections.addAll pointsInRange, northWest.queryRange(range)
    Collections.addAll pointsInRange, northEast.queryRange(range)
    Collections.addAll pointsInRange, southWest.queryRange(range)
    Collections.addAll pointsInRange, southEast.queryRange(range)
    
    Set queryRange = pointsInRange
    
End Function

Public Function contains(p As XYpoint) As Boolean
    contains = boundary.contains(p)
End Function


Public Sub drawTree()
    Dim wksht As Worksheet, p As XYpoint, r As Single, left As Single, top As Single, width As Single, height As Single, shp As Shape
    Set wksht = ThisWorkbook.Sheets.add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    r = 2
    
    'draw boundary
    left = boundary.center.x - boundary.halfDimension + 1
    top = boundary.center.y - boundary.halfDimension + 1
    width = 2 * boundary.halfDimension
    height = 2 * boundary.halfDimension
    Set shp = wksht.Shapes.AddShape(msoShapeRectangle, left, top, width, height)
    shp.Line.Visible = msoTrue
    shp.Line.Weight = 0.1
    shp.Line.ForeColor.RGB = RGB(0, 0, 0)
    shp.Fill.Visible = msoTrue
    shp.Fill.Transparency = 0.8
    shp.Fill.ForeColor.RGB = RGB(120, 120, 120)
    shp.TextFrame.characters.Text = points.count
    shp.TextFrame.characters.Font.ColorIndex = 2
    shp.TextFrame.VerticalAlignment = xlVAlignCenter
    shp.TextFrame.HorizontalAlignment = xlHAlignCenter
    
    'recurse
    If Not northWest Is Nothing Then
        northWest.drawTree
        northEast.drawTree
        southWest.drawTree
        southEast.drawTree
        
        Set shp = wksht.Shapes.AddLine(boundary.center.x, boundary.center.y, northWest.getBoundary().center.x, northWest.getBoundary().center.y)
        Set shp = wksht.Shapes.AddLine(boundary.center.x, boundary.center.y, northEast.getBoundary().center.x, northEast.getBoundary().center.y)
        Set shp = wksht.Shapes.AddLine(boundary.center.x, boundary.center.y, southWest.getBoundary().center.x, southWest.getBoundary().center.y)
        Set shp = wksht.Shapes.AddLine(boundary.center.x, boundary.center.y, southEast.getBoundary().center.x, southEast.getBoundary().center.y)
        
    End If
    
     'draw points last so they are on top
    For Each p In points
        Set shp = wksht.Shapes.AddShape(msoShapeOval, left:=p.x - r + 1, top:=p.y - r + 1, width:=2 * r, height:=2 * r)
        shp.Line.Visible = msoFalse
        shp.Fill.Transparency = 0.3
        shp.Fill.ForeColor.RGB = RGB(0, 0, 0)
    Next p
    
End Sub
