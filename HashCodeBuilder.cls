VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HashCodeBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private constant_ As Long
Private total_ As Long


Public Sub initializeVariables(Optional initialNonZeroOddNumber As Long, Optional multiplierNonZeroOddNumber As Long)

    If (initialNonZeroOddNumber = 0) Then
        constant_ = 37
    ElseIf (initialNonZeroOddNumber Mod 2 = 0) Then
        err.Raise Number:=5, source:="Factory.createHashCodeBuilder", Description:="HashCodeBuilder requires an odd initial value"
    Else
        constant_ = multiplierNonZeroOddNumber
    End If
    
    If (multiplierNonZeroOddNumber = 0) Then
        'variable was not supplied
        total_ = 17
    ElseIf (multiplierNonZeroOddNumber Mod 2 = 0) Then
        err.Raise Number:=5, source:="Factory.createHashCodeBuilder", Description:="HashCodeBuilder requires an odd multiplier"
    Else
        total_ = multiplierNonZeroOddNumber
    End If
End Sub

'' *****************************
'' ****** PRIMITIVE TYPES ******
'' *****************************

Public Function appendLong(value As Long) As HashCodeBuilder
    total_ = total_ * constant_ + value
    Set appendLong = Me
End Function

Public Function appendInteger(value As Integer) As HashCodeBuilder
    total_ = total_ * constant_ + value
    Set appendInteger = Me
End Function

Public Function appendByte(value As Byte) As HashCodeBuilder
    total_ = total_ * constant_ + value
    Set appendByte = Me
End Function

Public Function appendSingle(value As Single) As HashCodeBuilder
    total_ = total_ * constant_ + value
    Set appendSingle = Me
End Function

Public Function appendDouble(value As Double) As HashCodeBuilder
    total_ = total_ * constant_ + value
    Set appendDouble = Me
End Function

Public Function appendBoolean(value As Boolean) As HashCodeBuilder
    If value = True Then
        total_ = total_ * constant_ + value
    End If
    Set appendBoolean = Me
End Function

'' ***************************************
'' ****** ARRAYS OF PRIMITIVE TYPES ******
'' ***************************************

Public Function appendLongArray(Arr() As Long) As HashCodeBuilder
    If UBound(Arr, 1) - LBound(Arr, 1) + 1 = 0 Then
        total_ = total_ * constant_
    Else
        Dim val As Long
        For Each val In Arr
            appendLong val
        Next val
        
    Set appendLongArray = Me
End Function

Public Function appendIntegerArray(Arr() As Integer) As HashCodeBuilder
    If UBound(Arr, 1) - LBound(Arr, 1) + 1 = 0 Then
        total_ = total_ * constant_
    Else
        Dim val As Integer
        For Each val In Arr
            appendInteger val
        Next val
        
    Set appendIntegerArray = Me
End Function

Public Function appendByteArray(Arr() As Byte) As HashCodeBuilder
    If UBound(Arr, 1) - LBound(Arr, 1) + 1 = 0 Then
        total_ = total_ * constant_
    Else
        Dim val As Byte
        For Each val In Arr
            appendByte val
        Next val
        
    Set appendByteArray = Me
End Function

Public Function appendSingleArray(Arr() As Single) As HashCodeBuilder
     If UBound(Arr, 1) - LBound(Arr, 1) + 1 = 0 Then
        total_ = total_ * constant_
    Else
        Dim val As Single
        For Each val In Arr
            appendShort val
        Next val
        
    Set appendSingleArray = Me
End Function

Public Function appendDoubleArray(Arr() As Double) As HashCodeBuilder
     If UBound(Arr, 1) - LBound(Arr, 1) + 1 = 0 Then
        total_ = total_ * constant_
    Else
        Dim val As Double
        For Each val In Arr
            appendDouble val
        Next val
        
    Set appendDoubleArray = Me
End Function

Public Function appendBooleanArray(Arr() As Boolean) As HashCodeBuilder
     If UBound(Arr, 1) - LBound(Arr, 1) + 1 = 0 Then
        total_ = total_ * constant_
    Else
        Dim val As Boolean
        For Each val In Arr
            appendBoolean val
        Next val
        
    Set appendBooleanArray = Me
End Function

Public Function hashCode() As Long
    hashCode = total_
End Function
