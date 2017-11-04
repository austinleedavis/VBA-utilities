VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XYPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'simple coordinate object to represent points and vectors
Public x As Double
Public y As Double

Public Function getTranslation(xShift As Double, yShift As Double) As XYpoint
    Dim p As XYpoint
    Set p = New XYpoint
    p.x = x + xShift
    p.y = y + yShift
    Set getTranslation = p
End Function


Public Function toString() As String
    toString = "[Class: XYPoint; x: " & x & ", y: " & y & " ]"
End Function
