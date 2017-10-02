Attribute VB_Name = "Random"
Option Explicit

' This Class extends the functionality of the built-in Rnd function by
' providing functions for generating random strings, integers, Longs, etc.
'
' Warning: If you don't call the Randomize function before calling these
' functions, they may return the same random number value each time. And
' therefore, you may not get a truly random number.

Private Const CHAR_ARR As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
Private Const LEN_CHAR_ARR As Integer = 62
Private LastStringValue As String

' Resets the Rnd seed. If the seed parameter is provided, this function _
 will cause Rnd to produce a consistent sequence of pseudorandom numbers. _
 For example, if 5 is given as the seed, the sequence of random numbers will always be
Public Sub setSeed(Optional seed As Long)
    If (seed Is Nothing) Then
        Randomize
    Else
        Rnd (-1) ' This must be called
        Randomize (seed)
    End
End Sub


' Generates a randomized string of a a specific length
' @param length - The desired length of the resulting string. If no value is provided, the default length of eight (8) is used.
' @param characters - A string from which characters will be selected at random. If not provided, a random string will be generated using all characters 0-9, A-z
Function NextString(Optional length As Long = 8, Optional characters As String = CHAR_ARR) As String

    Dim s As String
    s = Space(length)
    Dim charLen As Long
    charLen = Len(characters) - 1
    Dim n As Long
    Dim nl As Long
    For n = 1 To length 'don't hardcode the length twice
'        Do
'            ch = Rnd() * 127 'This could be more efficient.
'            '48 is '0', 57 is '9', 65 is 'A', 90 is 'Z', 97 is 'a', 122 is 'z'.
'        Loop While ch < 48 Or ch > 57 And ch < 65 Or ch > 90 And ch < 97 Or ch > 122
'        Mid(s, n, 1) = Chr(ch) 'bit more efficient than concatenation
        nl = NextLong(1, charLen)
        Mid(s, n, 1) = Mid(characters, nl, 1)
    Next

    LastStringValue = s
    NextString = s

End Function

Function LastString() As String
    LastString = LastStringValue
End Function

Function NextLong(Optional LowerBound As Long = 0, Optional UpperBound As Long = Longs.MAX_VALUE) As Long
    NextLong = (UpperBound - LowerBound + 1) * Rnd + LowerBound
End Function

Function LastLong(Optional LowerBound As Long = 0, Optional UpperBound As Long = Longs.MAX_VALUE) As Long
    LastLong = (UpperBound - LowerBound + 1) * Rnd(0) + LowerBound
End Function

Function NextBoolean(Optional trueFrequency As Double = 0.5) As Boolean
    NextBoolean = IIf(Rnd() < trueFrequency, True, False)
End Function

Function LastBoolean(Optional trueFrequency As Double = 0.5) As Boolean
    LastBoolean = IIf(Rnd(0) < trueFrequency, True, False)
End Function
















