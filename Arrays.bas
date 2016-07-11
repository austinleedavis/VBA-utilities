Attribute VB_Name = "Arrays"
Option Compare Text
Option Explicit
Option Base 0

'Sorts the array using the MergeSort algorithm (follows the pseudocode on Wikipedia
'O(n*log(n)) time; O(n) space
Public Sub MergeSort(A() As Variant)
    Dim B() As Variant
    ReDim B(0 To UBound(A))
    TopDownSplitMerge A, 0, UBound(A), B
End Sub

'Used by MergeSortAlgorithm
Private Sub TopDownSplitMerge(A() As Variant, iBegin As Long, iEnd As Long, B() As Variant)
    
    If iEnd - iBegin < 2 Then ' if run size = 1
        Exit Sub ' consider it sorted
    End If
    
    ' recursively split runs into two halves until run size = 1
    ' then merge them and return back up the call chain
    Dim iMiddle As Long
    iMiddle = (iEnd + iBegin) / 2 ' iMiddle = mid point
    TopDownSplitMerge A, iBegin, iMiddle, B 'split-merge left half
    TopDownSplitMerge A, iMiddle, iEnd, B ' split-merge right half
    TopDownMerge A, iBegin, iMiddle, iEnd, B ' merge the two half runs
    Copy B, iBegin, A, iBegin, iEnd - iBegin 'copy the merged runs back to A
End Sub

'Used by MergeSort algirtm
Private Sub TopDownMerge(A() As Variant, iBegin As Long, iMiddle As Long, iEnd As Long, B() As Variant)
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
            B(k) = A(i)
            i = i + 1
        Else
            B(k) = A(j)
            j = j + 1
        End If
    Next k
End Sub

'Used by MergeSort algorithm
Private Sub CopyRange(source() As Variant, iBegin As Long, iEnd As Long, dest() As Variant)
    Dim k As Long
    For k = iBegin To iEnd Step 1
        destination(k) = source(k)
    Next k
End Sub

'Copies an array from the specified source array, beginning at the specified position, to the specified position in the destination array
Public Sub Copy(ByRef src() As Variant, srcPos As Long, ByRef dst() As Variant, dstPos As Long, length As Long)
    
    'Check if all offsets and lengths are non negative
    If srcPos < 0 Or dstPos < 0 Or length < 0 Then
        err.Raise 9, , "negative value supplied"
    End If
     
    'Check if ranges are valid
    If length + srcPos > UBound(src) Then
        err.Raise 9, , "Not enough elements to copy, src+length: " & srcPos + length & ", UBound(src): " & UBound(src)
    End If
    If length + dstPos > UBound(dst) Then
        err.Raise 9, , "Not enough room in destination array. dstPos+length: " & dstPos + length & ", UBound(dst): " & UBound(dst)
    End If
    Dim i As Long
    i = 0
    'copy src elements to dst
    Do While length > i
        dst(dstPos + i) = src(srcPos + i)
        i = i + 1
    Loop
    
End Sub
