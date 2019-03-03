Attribute VB_Name = "Collections"
Option Explicit
Option Base 1

'Returns True if the Collection contains the specified key. Otherwise, returns False
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
    
    For index = 1 To col.count Step 1
        If col(index) = value Then
            indexOf = index
            Exit Function
        End If
    Next index
    indexOf = -1
End Function

''Sorts the given collection using the Arrays.MergeSort algorithm.
'' O(n log(n)) time
'' O(n) space
'Public Sub mergeSort(col As Collection)
'    Dim A() As Variant
'    Dim B() As Variant
'    A = Collections.ToArray(col)
'    Arrays.mergeSort A()
'    Set col = Collections.FromArray(A())
'End Sub

'Returns an array which exactly matches this collection.
' Note: This function is not safe for concurrent modification.
Public Function toArray(col As Collection) As Variant
    Dim A() As Variant
    ReDim A(0 To col.count)
    Dim i As Long
    For i = 0 To col.count - 1
        A(i) = col(i + 1)
    Next i
    toArray = A()
End Function

'Returns a Collection which exactly matches the given Array
' Note: This function is not safe for concurrent modification.
Public Function FromArray(A() As Variant) As Collection
    Dim col As Collection
    Set col = New Collection
    Dim element As Variant
    For Each element In A
        col.add element
    Next element
    Set FromArray = col
End Function

Public Sub BubbleSort()

    Dim cFruit As Collection
    Dim vItm As Variant
    Dim i As Long, j As Long
    Dim vTemp As Variant

    Set cFruit = New Collection

    'fill the collection
    cFruit.add "Mango", "Mango"
    cFruit.add "Apple", "Apple"
    cFruit.add "Peach", "Peach"
    cFruit.add "Kiwi", "Kiwi"
    cFruit.add "Lime", "Lime"

    'Two loops to bubble sort
   For i = 1 To cFruit.count - 1
        For j = i + 1 To cFruit.count
            If cFruit(i) > cFruit(j) Then
                'store the lesser item
               vTemp = cFruit(j)
                'remove the lesser item
               cFruit.remove j
                're-add the lesser item before the
               'greater Item
               cFruit.add vTemp, vTemp, i
            End If
        Next j
    Next i

    'Test it
   For Each vItm In cFruit
        Debug.Print vItm
    Next vItm

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''
'Sorts the array using the MergeSort algorithm (follows the pseudocode on Wikipedia
'O(n*log(n)) time; O(n) space
Public Sub mergeSort(A As Collection)
    Dim b As Collection
    Set b = New Collection
    Collections.Copy A, 1, b, 1, A.count
    TopDownSplitMerge A, 1, A.count, b
End Sub

'Used by MergeSortAlgorithm
Private Sub TopDownSplitMerge(A As Collection, iBegin As Long, iEnd As Long, b As Collection)
    
    If iEnd - iBegin < 2 Then ' if run size = 1
        Exit Sub ' consider it sorted
    End If
    
    ' recursively split runs into two halves until run size = 1
    ' then merge them and return back up the call chain
    Dim iMiddle As Long
    iMiddle = (iEnd + iBegin) / 2 ' iMiddle = mid point
    TopDownSplitMerge A, iBegin, iMiddle, b 'split-merge left half
    TopDownSplitMerge A, iMiddle, iEnd, b ' split-merge right half
    TopDownMerge A, iBegin, iMiddle, iEnd, b ' merge the two half runs
    Copy b, iBegin, A, iBegin, iEnd - iBegin 'copy the merged runs back to A
End Sub

'Used by MergeSort algirtm
Private Sub TopDownMerge(A As Collection, iBegin As Long, iMiddle As Long, iEnd As Long, b As Collection)
    'left half is A[iBegin:iMiddle-1]
    'right half is A[iMiddle:iEnd-1]
    Dim i As Long
    Dim j As Long
    Dim k As Long
    i = iBegin
    j = iMiddle
    
    'while there are elements in the left or right runs...
    For k = iBegin To iEnd Step 1
        'If left run head exists and is <= existing right run head.
        If i < iMiddle And (j >= iEnd Or A(i) <= A(j)) Then
            b.add A(i)
            i = i + 1
        Else
            b(k) = A(j)
            j = j + 1
        End If
    Next k
End Sub

'Used by MergeSort algorithm
Private Sub CopyRange(source As Collection, iBegin As Long, iEnd As Long, dest As Collection)
    Dim k As Long
    For k = iBegin To iEnd Step 1
        destination(k) = source(k)
    Next k
End Sub

'Copies an array from the specified source array, beginning at the specified position, to the specified position in the destination array
Public Sub Copy(ByRef src As Collection, srcPos As Long, ByRef dst As Collection, dstPos As Long, length As Long)
    
    'Check if all offsets and lengths are non negative
    If srcPos < 1 Then
        err.Raise 9, , "srcPos too small: " & destPos
    End If
    If destPos < 1 Then
        err.Raise 9, , "destPos too small: " & destPos
    End If
    If length < 0 Then
        err.Raise 9, , "negative length provided"
    End If
    
     
    'Check if ranges are valid
    If length + srcPos - 1 > src.count Then
        err.Raise 9, , "Not enough elements to copy, (src+length - 1): " & srcPos + length - 1 & ", src.Count: " & src.count
    End If
    If length + dstPos - 1 > dst.count Then
        err.Raise 9, , "Not enough room in destination array. (dstPos+length - 1): " & dstPos + length - 1 & ", dst.Count: " & dst.count
    End If
    Dim i As Long
    i = 0
    'copy src elements to dst
    Do While length > i
        dst(dstPos + i) = src(srcPos + i)
        i = i + 1
    Loop
    
End Sub

' @description adds all elements of the source Collection to the destination Collection
' @param dest the destination collection to which the elements will be added
' @param source the collection from which the elements originate
Public Sub addAll(dest As Collection, source As Collection)
    Dim v As Variant
    
    For Each v In source
        dest.add v
    Next v
    
End Sub
