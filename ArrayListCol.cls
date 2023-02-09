VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayListCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Base 0 'Index all arrays from 0
Implements iCollection
'Implements iList
Implements iIterable

Private Const DEFAULT_CAPACITY As Long = 10 ' The default intial capacity of the ArrayListCol
Private Const MAX_ARRAY_SIZE As Long = 2147483639 ' maximum size of array to allocate
Private elementData As Collection ' The actual array of Variants in the ArrayListCol
Private modCount As Long ' The number of times this list has been structurally modified


Private Sub Class_Initialize()
    Set elementData = New Collection
End Sub


'Trims the capacity of this ArrayListCol instance to list's current size. An application can use this operation to minimize the storage of an ArrayListCol instance
'Public Sub trimToSize()
'    modCount = modCount + 1
'    capacity = elementData.Count - 1
'End Sub

'Public Sub ensureCapacity(minCapacity As Long)
'    modCount = modCount + 1
'
'    'Overflow-conscious code
'    If minCapacity - elementData.Count - 1 > 0 Then
'        grow (minCapacity)
'    End If
'
'End Sub

'Private Sub grow(minCapacity As Long)
'    'Overflow-conscious code
'    Dim oldCapacity As Long
'    oldCapacity = UBound(elementData)
'    Dim newCapacity As Long
'    newCapacity = oldCapacity + (oldCapacity / 2)
'    If newCapacity - minCapacity < 0 Then
'        newCapacity = minCapacity
'    End If
'    If newCapacity - MAX_ARRAY_SIZE > 0 Then
'        newCapacity = hugeCapacity(minCapacity)
'    End If
'
'    'minCapacity is usually close to size, so this is a win
'    ReDim Preserve elementData(0 To newCapacity)
'
'End Sub

'Private Function hugeCapacity(minCapacity As Long) As Long
'    If minCapacity < 0 Then
'        err.Raise 6 ' overflow
'    End If
'    hugeCapacity = IIf(minCapacity > MAX_ARRAY_SIZE, _
'        2147483647, MAX_ARRAY_SIZE)
'End Function

'Private Sub ensureCapacityInternal(minCapacity As Long)
'    ensureExplicitCapacity (minCapacity)
'End Sub
'
'Private Sub ensureExplicitCapacity(minCapacity As Long)
'    modCount = modCount + 1
'
'    'overflow-conscious code
'    If minCapacity - UBound(elementData) > 0 Then
'        grow (minCapacity)
'    End If
'
'End Sub


' Returns the number of elements in this list
Public Function size() As Long
    size = elementData.Count
End Function

'Returns true if this list contains no elements
Public Function isEmpty() As Boolean
    isEmpty = (size = 0)
End Function

' Returns true if this list contains the specified element
Public Function contains(ByRef o As Variant) As Boolean
    contains = (indexOf(o) >= 0)
End Function

'Returns the index of the first occurrence of the specified element in this list, or -1 if this list does not contain the element
Public Function indexOf(ByRef o As Variant) As Long
    Dim i As Long
    If o = Empty Then
        err.Raise 424, , "NullPointerException"
    Else
        If size = 0 Then
            indexOf = -1
            Exit Function
        End If
        For i = 0 To size - 1 Step 1
            If o = getIndexInternal(i) Then
                indexOf = i
                Exit Function
            End If
        Next i
    End If
    indexOf = Longs.valueOf(-1)
End Function

'Returns the index of the last occurrence of the specified element in this list, or -1 if this list does not contain the element
Public Function lastIndexOf(ByRef o As Variant) As Long
        Dim i As Long
    If o = Empty Then
        err.Raise 424, , "NullPointerException"
    Else
        If size = 0 Then
            lastIndexOf = -1
            Exit Function
        End If
        For i = size - 1 To 0 Step -1
            If o = getIndexInternal(i) Then
                lastIndexOf = i
                Exit Function
            End If
        Next i
    End If
    lastIndexOf = -1
End Function

Public Function ToArray() As Variant()
    Dim arrCopy() As Variant
    ReDim arrCopy(0 To size - 1)
    Dim i As Long
    For i = 0 To size - 1 Step 1
        arrCopy(i) = getIndexInternal(i)
    Next i
    ToArray = arrCopy
End Function


Public Function getIndex(index As Long) As Variant
    rangeCheck (index)
    getIndex = elementData.Item(index + 1)
End Function

Private Function getIndexInternal(index As Long) As Variant
    getIndexInternal = elementData.Item(index + 1)
End Function

Public Function setIndex(index As Long, ByRef element As Variant) As Variant
    rangeCheck (index)
    
    Dim oldValue As Variant
    oldValue = getIndexInternal(index)
    elementData.add element, , , index + 2 = element
    removeIndex (index)
    setIndex = oldValue
End Function

Public Function add(ByRef e As Variant) As Boolean
'    ensureCapacityInternal (size + 1)
    elementData.add e
'    size = size + 1
    add = True
End Function

Public Sub addAtIndex(index As Long, ByRef element As Variant)
    rangeCheckForAdd (index)
'    ensureCapacity (size + 1)
    
    'shift current elements to the right one
'    arrayCopy elementData, index, elementData, index + 1, size - index
    If index = elementData.Count Then
        elementData.add element
    Else
        elementData.add element, , index + 1
    End If
End Sub

Public Function removeIndex(index As Long) As Variant
    rangeCheck (index)
    
    modCount = modCount + 1
    
    Dim oldValue As Variant
    oldValue = getIndex(index)
    
    elementData.remove index + 2
'    Dim numMoved As Long
'    numMoved = size - index - 1
'    If numMoved > 0 Then
'        'move the remaining elements left one
'        arrayCopy elementData, index + 1, elementData, index, numMoved
'    End If
'
'    size = size - 1
'    elementData(size) = Empty
    removeIndex = oldValue
End Function

Private Sub rangeCheck(index As Long)
    If index >= size Then
        err.Raise 9, , outOfBoundsMsg(index)
    End If
End Sub

Private Sub rangeCheckForAdd(index As Long)
    If index > size Or index < 0 Then
        err.Raise 9, , outOfBoundsMsg(index)
    End If
End Sub

Private Function outOfBoundsMsg(index As Long) As String
    outOfBoundsMsg = "Index " & index & ", Size: " & size
End Function

Public Function remove(ByRef o As Variant) As Boolean
    Variants.requireNonNull o
    
    Dim index As Long
    For index = 0 To size - 1 Step 1
        If getIndexInternal(index) = o Then
            fastRemove (index)
            remove = True
            Exit Function
        End If
    Next index
    
    remove = False
End Function

Private Sub fastRemove(index As Long)
    modCount = modCount + 1
    
    elementData.remove index + 1
'    Dim numMoved As Long
'    numMoved = size - index - 1
'    If numMoved > 0 Then
'        arrayCopy elementData, index + 1, elementData, index, numMoved
'    End If
'    size = size - 1
'    elementData(size) = Null
End Sub

Public Sub clear()
    modCount = modCount + 1
    
    Set elementData = Nothing
    Set elementData = New Collection
'    Dim i As Long
'    For i = 0 To size Step 1
'        elementData(i) = Null
'    Next i
'
'    size = 0
End Sub

Public Function addAll(ByRef C As iCollection) As Boolean
    Variants.requireNonNull C
        
        Dim A() As Variant
        A = C.ToArray
    
    For Each element In A
        elementData.add element
    Next element
    addAll = C.size > 0
    
'    Dim numNew As Long
'    numNew = UBound(arr) - LBound(arr)
''    ensureCapacityInternal (size + numNew)
'
'    Dim i As Long
'    Dim j As Long
'    j = size
'    For i = LBound(arr) To UBound(arr) Step 1
'        elementData(j) = arr(i)
'        j = j + 1
'    Next i
'    size = size + numNew
'    addAll = numNew <> 0
End Function


Public Function addAllAtIndex(index As Long, ByRef C As iCollection) As Boolean
    rangeCheckForAdd index
    Variants.requireNonNull C
        
    If index < elementData.Count Then
        Dim offset As Long
        offset = 0
        Dim A() As Variant
        A = C.ToArray
                
        For Each element In A
            elementData.add element, , , index + 1 + offset
            offset = offset + 1
        Next element
    Else
        addAll C
    End If
    addAllAtIndex = (C.size > 0)
    
    
'    rangeCheckForAdd (index)
'
'    Dim numNew As Long
'    numNew = UBound(arr) - LBound(arr)
'    ensureCapacityInternal (size + numNew)
'
'    Dim numMoved As Long
'    numMoved = size - index
'    If numMoved > 0 Then
'        arrayCopy elementData, index, elementData, index + numNew, numMoved
'    End If
'
'    Dim i As Long
'    Dim j As Long
'    j = LBound(arr)
'    For i = index To index + numNew Step 1
'        elementData(i) = arr(j)
'        j = j + 1
'    Next i
'
'    size = size + numNew
'    addAllAtIndex = numNew <> 0
End Function

Sub removeRange(fromIndex As Long, toIndex As Long)
    modCount = modCount + 1
    Dim numRemoved As Long
    numRemoved = toIndex - fromIndex
    Dim i As Long
    
    For i = numRemoved To 0 Step -1
        removeIndex (fromIndex)
    Next i
    
    
'    Dim numMoved As Long
'    numMoved = size - toIndex
'    arrayCopy elementData, toIndex, elementData, fromIndex, numMoved
'
'    'clear to let GC do its work
'    Dim newSize As Long
'    newSize = size - (toIndex - fromIndex)
'    Dim i As Long
'    For i = newSize To size Step 1
'        elementData(i) = Null
'    Next i
'    size = newSize
End Sub

Public Function removeAll(ByRef C As iCollection) As Boolean
    Variants.requireNonNull C
    batchRemove C, False
End Function

Public Function retainAll(ByRef C As iCollection) As Boolean
    Variants.requireNonNull C
    batchRemove C, True
End Function

Private Function batchRemove(ByRef C As ArrayListCol, complement As Boolean) As Boolean
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




Public Function subList(fromIndex As Long, toIndex As Long) As ArrayListCol
    subListRangeCheck fromIndex, toIndex, Me.getSize
    subList = New ArrayListCol
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
    Set NewEnum = elementData.[_NewEnum]
End Property

' ------------------ iCollection Interface Methods


Public Function iCollection_add(ByRef e As Variant) As Boolean
    iCollection_add = add(e)
End Function

Public Function iCollection_addAll(ByRef C As iCollection) As Boolean
    iCollection_addAll = addAll(C)
End Function

Public Sub iCollection_clear()
    clear
End Sub

Public Function iCollection_contains(ByRef o As Variant) As Boolean
    iCollection_contains = contains(o)
End Function

Public Function iCollection_equals(ByRef o As Variant) As Boolean
    ''TODO
End Function

Public Function iCollection_hashCode() As Long
    ''TODO
End Function

Public Function iCollection_isEmpty() As Boolean
    iCollection_isEmpty = isEmpty
End Function

'iterator()

'parallelStream()

Public Function iCollection_remove(o As Variant) As Boolean
    iCollection_remove = remove(o)
End Function

Public Function iCollection_removeAll(ByRef C As iCollection) As Boolean
    iCollection_removeAll = removeAll(C)
End Function

Public Function iCollection_retainAll(ByRef C As iCollection) As Boolean
    iCollection_retainAll = retainAll(C)
End Function

Public Function iCollection_size() As Long
    iCollection_size = size
End Function

'spliterator
'stream
Public Function iCollection_toArray() As Variant()
    iCollection_toArray = ToArray
End Function

Public Property Get iCollection_NewEnum() As IUnknown
    ''specified in iIterable
    Set iCollection_NewEnum = NewEnum
End Property

Public Property Get iIterable_NewEnum() As IUnknown
    ''specified in iIterable
    Set iIterable_NewEnum = NewEnum
End Property





