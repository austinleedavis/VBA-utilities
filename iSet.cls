VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

''Methods inherited from the iCollection interface

Public Function add(ByRef e As Variant) As Boolean
    'inherited method from iCollection
End Function

Public Function addAll(ByRef c As iCollection) As Boolean
'inherited method from iCollection
End Function

Public Sub clear()
'inherited method from iCollection
End Sub

Public Function contains(ByRef o As Variant) As Boolean
'inherited method from iCollection
End Function

Public Function equals(ByRef o As Variant) As Boolean
'inherited method from iCollection
End Function

Public Function hashCode() As Long
'inherited method from iCollection
End Function

Public Function isEmpty() As Boolean
'inherited method from iCollection
End Function

'iterator()

'parallelStream()

Public Function remove(o As Variant) As Boolean
'inherited method from iCollection
End Function

Public Function removeAll(c As iCollection) As Boolean
'inherited method from iCollection
End Function

Public Function retainAll(c As iCollection) As Boolean
'inherited method from iCollection
End Function

Public Function size() As Long
'inherited method from iCollection
End Function

'spliterator
'stream
Public Function toArray() As Variant()
'inherited method from iCollection
End Function

Public Property Get NewEnum() As IUnknown
'interface method
End Property

''Methods provided by the iSet interface
' NONE

