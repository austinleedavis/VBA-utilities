Attribute VB_Name = "Collections"
Option Explicit
Option Base 0

'Returns True if the Collection has the key, Key. Otherwise, returns False
Public Function hasKey(Key As Variant, col As collection) As Boolean
    Dim obj As Variant
    On Error GoTo err
    hasKey = True
    obj = col(Key)
    Exit Function

err:
    hasKey = False
End Function

'Returns True if the Collection contains an element equal to value
Public Function contains(value As Variant, col As collection) As Boolean
    contains = (indexOf(value, col) >= 0)
End Function


'Returns the first index of an element equal to value. If the Collection
'does not contain such an element, returns -1.
Public Function indexOf(value As Variant, col As collection) As Long

    Dim index As Long
    
    For index = 1 To col.count Step 1
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
Public Sub sort(col As collection, Optional ByRef c As IVariantComparator)
    Dim a() As Variant
    Dim b() As Variant
    a = Collections.ToArray(col)
    Arrays.sort a(), c
    Set col = Collections.FromArray(a())
End Sub

'Returns an array which exactly matches this collection.
' Note: This function is not safe for concurrent modification.
Public Function ToArray(col As collection) As Variant
    Dim a() As Variant
    ReDim a(0 To col.count)
    Dim i As Long
    For i = 0 To col.count - 1
        a(i) = col(i + 1)
    Next i
    ToArray = a()
End Function

'Returns a Collection which exactly matches the given Array
' Note: This function is not safe for concurrent modification.
Public Function FromArray(a() As Variant) As collection
    Dim col As collection
    Set col = New collection
    Dim element As Variant
    For Each element In a
        col.Add element
    Next element
    Set FromArray = col
End Function

'Adds all elements from the source collection, src, to the destination collection, dest.
'Returns true if the destination collection changed as a result of this operation; false otherwise.
Public Function AddAllFromCol(ByRef src As collection, ByRef dest As collection) As Boolean
    Dim count As Long
    count = dest.count
    
    For Each element In src
        dest.Add element
    Next element
    
    AddAllFromCol = (dest.count = count)
End Function

'Adds all elements from the source array, src, to the destination collection, dest
'Returns true if the destination collection changed as a result of this operation; false otherwise.
Public Function AddAllFromArray(src() As Variant, ByRef dest As collection) As Boolean
    Dim count As Long
    count = dest.count
    
    For Each element In src
        dest.Add element
    Next element
    
    AddAllFromCol = (dest.count = count)
End Function



