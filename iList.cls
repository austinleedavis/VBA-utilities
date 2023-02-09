VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iList"
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

Public Function addAll(ByRef C As iCollection) As Boolean
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

Public Function removeAll(C As iCollection) As Boolean
'inherited method from iCollection
End Function

Public Function retainAll(C As iCollection) As Boolean
'inherited method from iCollection
End Function

Public Function size() As Long
'inherited method from iCollection
End Function

'spliterator
'stream
Public Function ToArray() As Variant()
'inherited method from iCollection
End Function

''Methods provided by the iList interface

Public Sub addAt(index As Long, e As Variant)
'instance method
End Sub

Public Function addAllAt(index As Long, C As iCollection) As Boolean
'instance method
End Function

Public Function getIndex(index As Long) As Variant
'instance method
End Function

Public Function indexOf(o As Variant) As Long
'instance method
End Function

Public Function lastIndexOf(o As Variant) As Long
'instance method
End Function

Public Function removeIndex(index As Long) As Variant
'instance method
End Function

'replaceAll

Public Function setIndex(index As Long, element As Variant) As Variant
'instance method
End Function

'sort

Public Function subList(fromIndex As Long, toIndex As Long) As iList
'instance method
End Function



