VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

'Methods provided by the iMap interface
Public Sub clear()
'interface method
End Sub

'compute
'computeIfAbsent
'computeIfPresent

Public Function containsKey(Key As Variant) As Boolean
'interface method
End Function

Public Function containsValue(value As Variant) As Boolean
'interface method
End Function

Public Function entrySet() As iSet
'interface method
End Function

Public Function equals(o As Variant) As Boolean
'interface method
End Function

'foreach

Public Function getByKey(Key As Variant) As Variant
'interface method
End Function

Public Function getOrDefault(Key As Variant, DefaultValue As Variant) As Variant
'interface method
End Function

Public Function keySet() As iSet
'interface method
End Function

'merge

Public Function putKV(Key As Variant, value As Variant) As Variant
'interface method
End Function

Public Sub putAll(m As iMap)
'interface method
End Sub

Public Function putIfAbsent(Key As Variant, value As Variant) As Variant
'interface method
End Function

Public Function remove(Key As Variant) As Variant
'interface method
End Function

Public Function removeIfValueMatch(Key As Variant, value As Variant) As Boolean
'interface method
End Function

Public Function replace(Key As Variant, value As Variant) As Variant
'interface method
End Function

Public Function replaceIfValueMatch(Key As Variant, oldValue As Variant, newValue As Variant) As Boolean
'interface method
End Function

'replaceAll

Public Function size() As Long
'interface method
End Function

Public Function values() As iCollection
'interface method
End Function
