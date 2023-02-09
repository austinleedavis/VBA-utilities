VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Base 0 'Index all arrays from 0

Private Const DEFAULT_CAPACITY As Long = 10 ' The default intial capacity of the ArrayList
Private Const MAX_ARRAY_SIZE As Long = 2147483639 ' maximum size of array to allocate
Private elementData() As Variant ' The actual array of Variants in the ArrayList
Private size As Long ' size of the ArrayList
Private modCount As Long ' The number of times this list has been structurally modified


Private Sub Class_Initialize()
    size = 0
    ReDim elementData(0 To DEFAULT_CAPACITY)
End Sub


'Trims the capacity of this ArrayList instance to list's current size. An application can use this operation to minimize the storage of an ArrayList instance
Public Sub trimToSize()
    modCount = modCount + 1
    If size < UBound(elementData) Then
        ReDim Preserve elementData(0 To size)
    End If
End Sub

Public Sub ensureCapacity(minCapacity As Long)
    modCount = modCount + 1
    
    'Overflow-conscious code
    If minCapacity - UBound(elementData) > 0 Then
        grow (minCapacity)
    End If
        
End Sub

Private Sub grow(minCapacity As Long)
    'Overflow-conscious code
    Dim oldCapacity As Long
    oldCapacity = UBound(elementData)
    Dim newCapacity As Long
    newCapacity = oldCapacity + (oldCapacity / 2)
    If newCapacity - minCapacity < 0 Then
        newCapacity = minCapacity
    End If
    If newCapacity - MAX_ARRAY_SIZE > 0 Then
        newCapacity = hugeCapacity(minCapacity)
    End If
    
    'minCapacity is usually close to size, so this is a win
    ReDim Preserve elementData(0 To newCapacity)
    
End Sub

Private Function hugeCapacity(minCapacity As Long) As Long
    If minCapacity < 0 Then
        err.Raise 6 ' overflow
    End If
    hugeCapacity = IIf(minCapacity > MAX_ARRAY_SIZE, _
        2147483647, MAX_ARRAY_SIZE)
End Function

Private Sub ensureCapacityInternal(minCapacity As Long)
    ensureExplicitCapacity (minCapacity)
End Sub

Private Sub ensureExplicitCapacity(minCapacity As Long)
    modCount = modCount + 1
    
    'overflow-conscious code
    If minCapacity - UBound(elementData) > 0 Then
        grow (minCapacity)
    End If
    
End Sub

Private Property Let setSize(value As Long)
    size = value
End Property

' Returns the number of elements in this list
Public Property Get getSize() As Long
    getSize = size
End Property

'Returns true if this list contains no elements
Public Function isEmpty() As Boolean
    isEmpty = (size = 0)
End Function

' Returns true if this list contains the specified element
Public Function contains(o As Variant) As Boolean
    contains = (indexOf(o) >= 0)
End Function

'Returns the index of the first occurrence of the specified element in this list, or -1 if this list does not contain the element
Public Function indexOf(o As Variant) As Long
    Dim i As Long
    If o = Empty Then
        
        For i = 0 To size Step 1
            If elementData(i) Is Nothing Then
                indexOf = i
                Exit Function
            End If
        Next i
    Else
        
        For i = 0 To size Step 1
            If o = elementData(i) Then
                indexOf = i
                Exit Function
            End If
        Next i
    End If
    indexOf = -1
End Function

'Returns the index of the last occurrence of the specified element in this list, or -1 if this list does not contain the element
Public Function lastIndexOf(o As Variant) As Long
    Dim i As Long
    If (o Is Nothing) Then
        For i = size - 1 To 0 Step 1
            If elementData(i) Is Nothing Then
                lastIndexOf = i
                Exit Function
            End If
        Next i
    Else
        For i = size - 1 To 0 Step 1
            If o Is elementData(i) Then
                lastIndexOf = i
                Exit Function
            End If
        Next i
    End If
End Function

Public Function ToArray() As Variant()
    Dim arrCopy() As Variant
    ReDim arrCopy(0 To size)
    Dim i As Long
    For i = 0 To size Step 1
        arrCopy(i) = elementData(i)
    Next i
    ToArray = arrCopy
End Function


Public Function getIndex(index As Long) As Variant
    rangeCheck (index)
    getIndex = elementData(index)
End Function

Public Function setIndex(index As Long, element As Variant) As Variant
    rangeCheck (index)
    
    Dim oldValue As Variant
    oldValue = elementData(index)
    elementData(index) = element
    setIndex = oldValue
End Function

Public Function add(e As Variant) As Boolean
    ensureCapacityInternal (size + 1)
    elementData(size) = e
    size = size + 1
    add = True
End Function

Public Sub addIndex(index As Long, element As Variant)
    rangeCheckForAdd (index)
    ensureCapacity (size + 1)
    
    'shift current elements to the right one
    arrayCopy elementData, index, elementData, index + 1, size - index
    elementData(index) = element
    size = size + 1
End Sub

Public Function removeIndex(index As Long) As Variant
    rangeCheck (index)
    
    modCount = modCount + 1
    
    Dim oldValue As Variant
    oldValue = elementData(index)
    
    
    Dim numMoved As Long
    numMoved = size - index - 1
    If numMoved > 0 Then
        'move the remaining elements left one
        arrayCopy elementData, index + 1, elementData, index, numMoved
    End If
    
    size = size - 1
    elementData(size) = Empty
    removeIndex = oldValue
End Function

Private Sub rangeCheck(index As Long)
    If index >= size Then
        err.Raise 9, , outOfBoundsMsg(index)
    End If
End Sub

Private Sub rangeCheckForAdd(index As Long)
    If index >= size Or index < 0 Then
        err.Raise 9, , outOfBoundsMsg(index)
    End If
End Sub

Private Function outOfBoundsMsg(index As Long) As String
    outOfBoundsMsg = "Index " & index & ", Size: " & size
End Function

Public Function remove(o As Variant) As Boolean
    Dim index As Long
    If o = Empty Then
        For index = 0 To size Step 1
            If elementData(index) Is Nothing Then
                fastRemove (index)
                removeVariant = True
                Exit Function
            End If
        Next index
    Else
        For index = 0 To size Step 1
            If elementData(index) = o Then
                fastRemove (index)
                removeVariant = True
                Exit Function
            End If
        Next index
    End If
    removeVariant = False
End Function

Private Sub fastRemove(index As Long)
    modCount = modCount + 1
    Dim numMoved As Long
    numMoved = size - index - 1
    If numMoved > 0 Then
        arrayCopy elementData, index + 1, elementData, index, numMoved
    End If
    size = size - 1
    elementData(size) = Null
End Sub

Public Sub clear()
    modCount = modCount + 1
    
    Dim i As Long
    For i = 0 To size Step 1
        elementData(i) = Null
    Next i
    
    size = 0
End Sub

Public Function addAll(Arr() As Variant) As Boolean
    Dim numNew As Long
    numNew = UBound(Arr) - LBound(Arr)
    ensureCapacityInternal (size + numNew)
    
    Dim i As Long
    Dim j As Long
    j = size
    For i = LBound(Arr) To UBound(Arr) Step 1
        elementData(j) = Arr(i)
        j = j + 1
    Next i
    size = size + numNew
    addAll = numNew <> 0
End Function

Public Function addAllAtIndex(index As Long, Arr() As Variant) As Boolean
    rangeCheckForAdd (index)
    
    Dim numNew As Long
    numNew = UBound(Arr) - LBound(Arr)
    ensureCapacityInternal (size + numNew)
    
    Dim numMoved As Long
    numMoved = size - index
    If numMoved > 0 Then
        arrayCopy elementData, index, elementData, index + numNew, numMoved
    End If
    
    Dim i As Long
    Dim j As Long
    j = LBound(Arr)
    For i = index To index + numNew Step 1
        elementData(i) = Arr(j)
        j = j + 1
    Next i
    
    size = size + numNew
    addAllAtIndex = numNew <> 0
End Function

Sub removeRange(fromIndex As Long, toIndex As Long)
    modCount = modCount + 1
    Dim numMoved As Long
    numMoved = size - toIndex
    arrayCopy elementData, toIndex, elementData, fromIndex, numMoved
    
    'clear to let GC do its work
    Dim newSize As Long
    newSize = size - (toIndex - fromIndex)
    Dim i As Long
    For i = newSize To size Step 1
        elementData(i) = Null
    Next i
    size = newSize
End Sub

Public Function removeAll(C As ArrayList) As Boolean
    If C Is Nothing Then
        err.Raise 31004, , "Collection c Is Nothing"
    End If
    batchRemove C, False
End Function

Public Function retainAll(C As ArrayList) As Boolean
    If C Is Nothing Then
        err.Raise 31004, , "Collection c Is Nothing"
    End If
    batchRemove C, True
End Function

Private Function batchRemove(C As ArrayList, complement As Boolean) As Boolean
    Dim r As Integer
    Dim w As Integer
    r = 0
    w = 0
    Dim modified As Boolean
    modified = False
    
    On Error GoTo finally
    For r = 0 To size Step 1
        If (C.contains(elementData(r)) = complement) Then
            elementData(w) = elementData(r)
            w = w + 1
        End If
    Next r
    batchRemove = modified
    Exit Function
finally:
    If r <> size Then
        ''''arrayCopy elementData, r, elementData, w, size - r
        w = w + size - r
    End If
    If w <> size Then
        'clear to let GC do its work
        Dim i As Long
        For i = w To size Step 1
            elementData(i) = Null
        Next i
        modCount = modCount + size - w
        size = w
        modified = True
        
    End If
    batchRemove = modified
    Exit Function
    
    
    
End Function


Public Sub arrayCopy(ByRef src() As Variant, srcPos As Long, ByRef dst() As Variant, dstPos As Long, length As Long)
    
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

Public Function subList(fromIndex As Long, toIndex As Long) As ArrayList
    subListRangeCheck fromIndex, toIndex, Me.getSize
    subList = New ArrayList
    subList.addAll elementData

End Function

Private Sub subListRangeCheck(fromIndex As Long, toIndex As Long, curSize As Long)
    If fromIndex < 0 Then
        err.Raise 6, , "fromIndex = " & fromIndex
    End If
    If toIndex > curSize Then
        err.Raise 6, , "toIndex = " & toIndex
    End If
    If fromIndex > toIndex Then
        err.Raise 328, , "fromIndex(" & fromIndex & ") > toIndex(" & toIndex & ")"
    End If
End Sub

Public Property Get NewEnum() As IUnknown
'interface method
End Property



