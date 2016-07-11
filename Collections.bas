Attribute VB_Name = "Collections"
Option Explicit
Option Base 0

'Returns True if the Collection has the key, Key. Otherwise, returns False
Public Function hasKey(Key As Variant, col As Collection) As Boolean
    Dim obj As Variant
    On Error GoTo err
    hasKey = True
    obj = col(Key)
    Exit Function

err:
    hasKey = False
End Function

'Returns True if the Collection contains an element equal to value
Public Function contains(value As Variant, col As Collection) As Boolean
    contains = (indexOf(value, col) >= 0)
End Function


'Returns the first index of an element equal to value. If the Collection
'does not contain such an element, returns -1.
Public Function indexOf(value As Variant, col As Collection) As Long

    Dim index As Long
    
    For index = 1 To col.Count Step 1
        If col(index) = value Then
            indexOf = index
            Exit Function
        End If
    Next index
    indexOf = -1
End Function

'Sorts the given collection using the Arrays.MergeSort algorithm.
' O(n log(n)) time
' O(n) space
Public Sub MergeSort(col As Collection)
    Dim A() As Variant
    Dim B() As Variant
    A = Collections.ToArray(col)
    Arrays.MergeSort A()
    Set col = Collections.FromArray(A())
End Sub

'Returns an array which exactly matches this collection.
' Note: This function is not safe for concurrent modification.
Public Function ToArray(col As Collection) As Variant
    Dim A() As Variant
    ReDim A(0 To col.Count)
    Dim i As Long
    For i = 0 To col.Count - 1
        A(i) = col(i + 1)
    Next i
    ToArray = A()
End Function

'Returns a Collection which exactly matches the given Array
' Note: This function is not safe for concurrent modification.
Public Function FromArray(A() As Variant) As Collection
    Dim col As Collection
    Set col = New Collection
    Dim element As Variant
    For Each element In A
        col.Add element
    Next element
    Set FromArray = col
End Function
