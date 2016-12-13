Attribute VB_Name = "Arrays"
Option Compare Text
Option Explicit
Option Base 0

Private Const INSERTIONSORT_THRESHOLD As Long = 7

'Sorts the array using the MergeSort algorithm (follows the Java legacyMergesort algorithm
'O(n*log(n)) time; O(n) space
Public Sub sort(ByRef a() As Variant, Optional ByRef c As IVariantComparator)

    If c Is Nothing Then
        MergeSort copyOf(a), a, 0, length(a), 0, Factory.newNumericComparator
    Else
        MergeSort copyOf(a), a, 0, length(a), 0, c
    End If
End Sub


Private Sub MergeSort(ByRef src() As Variant, ByRef dest() As Variant, low As Long, high As Long, off As Long, ByRef c As IVariantComparator)
    Dim length As Long
    Dim destLow As Long
    Dim destHigh As Long
    Dim mid As Long
    Dim i As Long
    Dim p As Long
    Dim q As Long
    
    length = high - low
    
    ' insertion sort on small arrays
    If length < INSERTIONSORT_THRESHOLD Then
        i = low
        Dim j As Long
        Do While i < high
            j = i
            Do While True
                If (j <= low) Then
                    Exit Do
                End If
                If (c.compare(dest(j - 1), dest(j)) <= 0) Then
                    Exit Do
                End If
                swap dest, j, j - 1
                j = j - 1 'decrement j
            Loop
            i = i + 1 'increment i
        Loop
        Exit Sub
    End If
    
    'recursively sort halves of dest into src
    destLow = low
    destHigh = high
    low = low + off
    high = high + off
    mid = (low + high) / 2
    MergeSort dest, src, low, mid, -off, c
    MergeSort dest, src, mid, high, -off, c
    
    'if list is already sorted, we're done
    If c.compare(src(mid - 1), src(mid)) <= 0 Then
        copy src, low, dest, destLow, length - 1
        Exit Sub
    End If
    
    'merge sorted halves into dest
    i = destLow
    p = low
    q = mid
    Do While i < destHigh
        If (q >= high) Then
           dest(i) = src(p)
           p = p + 1
        Else
            'Otherwise, check if p<mid AND src(p) preceeds scr(q)
            'See description of following idom at: http://stackoverflow.com/a/3245183/3795219
            Select Case True
               Case p >= mid, c.compare(src(p), src(q)) > 0
                   dest(i) = src(q)
                   q = q + 1
               Case Else
                   dest(i) = src(p)
                   p = p + 1
            End Select
        End If
       
        i = i + 1
    Loop
    
End Sub

'Sorts the array using the MergeSort algorithm (follows the Java legacyMergesort algorithm
'O(n*log(n)) time; O(n) space
Public Sub sortObjects(ByRef a() As Variant, ByRef c As IObjectComparator)

    If c Is Nothing Then
        err.Raise 3, "Arrays.sortObjects", "No IObjectComparator Provided to the sortObjects method."
    End If
    MergeSortObjects copyOfObjects(a), a, 0, length(a), 0, c
End Sub

Private Sub MergeSortObjects(ByRef src() As Object, ByRef dest() As Object, low As Long, high As Long, off As Long, ByRef c As IObjectComparator)
    Dim length As Long
    Dim destLow As Long
    Dim destHigh As Long
    Dim mid As Long
    Dim i As Long
    Dim p As Long
    Dim q As Long
    
    length = high - low
    
    ' insertion sort on small arrays
    If length < INSERTIONSORT_THRESHOLD Then
        i = low
        Dim j As Long
        Do While i < high
            j = i
            Do While True
                If (j <= low) Then
                    Exit Do
                End If
                If (c.compare(dest(j - 1), dest(j)) <= 0) Then
                    Exit Do
                End If
                swapObjects dest, j, j - 1
                j = j - 1 'decrement j
            Loop
            i = i + 1 'increment i
        Loop
        Exit Sub
    End If
    
    'recursively sort halves of dest into src
    destLow = low
    destHigh = high
    low = low + off
    high = high + off
    mid = (low + high) / 2
    MergeSortObjects dest, src, low, mid, -off, c
    MergeSortObjects dest, src, mid, high, -off, c
    
    'if list is already sorted, we're done
    If c.compare(src(mid - 1), src(mid)) <= 0 Then
        copy src, low, dest, destLow, length - 1
        Exit Sub
    End If
    
    'merge sorted halves into dest
    i = destLow
    p = low
    q = mid
    Do While i < destHigh
        If (q >= high) Then
           dest(i) = src(p)
           p = p + 1
        Else
            'Otherwise, check if p<mid AND src(p) preceeds scr(q)
            'See description of following idom at: http://stackoverflow.com/a/3245183/3795219
            Select Case True
               Case p >= mid, c.compare(src(p), src(q)) > 0
                   dest(i) = src(q)
                   q = q + 1
               Case Else
                   dest(i) = src(p)
                   p = p + 1
            End Select
        End If
       
        i = i + 1
    Loop
    
End Sub

Private Sub swap(arr() As Variant, a As Long, b As Long)
    Dim t As Variant
    t = arr(a)
    arr(a) = arr(b)
    arr(b) = t
End Sub

Private Sub swapObjects(arr() As Object, a As Long, b As Long)
    Dim t As Object
    t = arr(a)
    arr(a) = arr(b)
    arr(b) = t
End Sub

Public Function copyOf(ByRef original() As Variant) As Variant()
    Dim dest() As Variant
    ReDim dest(LBound(original) To UBound(original))
    CopyRange original, LBound(original), UBound(original), dest
    copyOf = dest
End Function

Private Sub CopyRange(source() As Variant, iBegin As Long, iEnd As Long, dest() As Variant)
    Dim k As Long
    For k = iBegin To iEnd Step 1
        dest(k) = source(k)
    Next k
End Sub

Private Sub CopyRangeObjects(source() As Object, iBegin As Long, iEnd As Long, dest() As Object)
    Dim k As Long
    For k = iBegin To iEnd Step 1
        dest(k) = source(k)
    Next k
End Sub

Public Function copyOfObjects(ByRef original() As Object) As Object()
    Dim dest() As Object
    ReDim dest(LBound(original) To UBound(original))
    CopyRange original, LBound(original), UBound(original), dest
    copyOf = dest
End Function

'Copies an array from the specified source array, beginning at the specified position, to the specified position in the destination array
Public Sub copy(ByRef src() As Variant, srcPos As Long, ByRef dst() As Variant, dstPos As Long, length As Long)
    
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

'Copies an array from the specified source array, beginning at the specified position, to the specified position in the destination array
Public Sub copyObjects(ByRef src() As Object, srcPos As Long, ByRef dst() As Object, dstPos As Long, length As Long)
    
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



'Adds all elements from the source collection, src, to the destination collection, dest.
'Returns true if the destination collection changed as a result of this operation; false otherwise.
Public Function AddAllFromCol(ByRef src As collection, ByRef dest() As Variant) As Boolean
    Dim count, i As Long
    count = dest.count
    i = 1
    ReDim Preserve dest(count + src.count)
    
    For Each element In src
        Set dest(count + i) = element
    Next element
    
    AddAllFromCol = (dest.count = count)
End Function

'Adds all elements from the source array, src, to the destination collection, dest
'Returns true if the destination collection changed as a result of this operation; false otherwise.
Public Function AddAllFromArray(ByRef src() As Variant, ByRef dest As collection) As Boolean
    Dim count, i As Long
    count = dest.count
    i = 1
    ReDim Preserve dest(count + src.count)
    
    For Each element In src
        Set dest(count + i) = element
    Next element
    
    AddAllFromCol = (dest.count = count)
End Function

Public Function length(ByRef a() As Variant) As Long
    length = UBound(a) - LBound(a) + 1
End Function


Public Function toString(ByRef a() As Variant) As String
    If length(a) <= 0 Then
        toString = "[]"
    ElseIf length(a) = 1 Then
        toString = "[ " & a(UBound(a)) & " ]"
    Else
        toString = "[ "
        Dim element As Variant
        For Each element In a
            toString = toString & element & " "
        Next element
        toString = toString & " ]"
    End If
End Function


