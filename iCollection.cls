VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function add(ByRef e As Variant) As Boolean
    'interface method
End Function

Public Function addAll(ByRef c As iCollection) As Boolean
'interface method
End Function

Public Sub clear()
'interface method
End Sub

Public Function contains(ByRef o As Variant) As Boolean
'interface method
End Function

Public Function equals(ByRef o As Variant) As Boolean
'interface method
End Function

Public Function hashCode() As Long
'interface method
End Function

Public Function isEmpty() As Boolean
'interface method
End Function

'iterator()

'parallelStream()

Public Function remove(ByRef o As Variant) As Boolean
'interface method
End Function

Public Function removeAll(ByRef c As iCollection) As Boolean
'interface method
End Function

Public Function retainAll(ByRef c As iCollection) As Boolean
'interface method
End Function

Public Function size() As Long
'interface method
End Function

'spliterator
'stream
Public Function toArray() As Variant()
'interface method
End Function

Public Property Get NewEnum() As IUnknown
	Attribute NewEnum.VB_UserMemId = -4
    ''specified in iIterable
End Property

