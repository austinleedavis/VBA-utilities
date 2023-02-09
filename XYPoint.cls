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
Public X As Double
Public Y As Double

Public Function getTranslation(xShift As Double, yShift As Double) As XYPoint
    Dim P As XYPoint
    Set P = New XYPoint
    P.X = X + xShift
    P.Y = Y + yShift
    Set getTranslation = P
End Function


Public Function toString() As String
    toString = "[Class: XYPoint; x: " & X & ", y: " & Y & " ]"
End Function
