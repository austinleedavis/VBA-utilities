Attribute VB_Name = "Strings"
Option Base 0

' Computes the Levenshtein distance between two strings. Levenshtein distance (LD) is a measure of the similarity
' between two strings: the source, string1, and the target, string2. The distance is the number of deletions, insertions,
' or substitutions required to transform string1 into string2.
' This implementation is provided by Patrick OBeirne of StackOverflow.com (ref http://stackoverflow.com/a/11584381/3795219)
Function Levenshtein(ByVal string1 As String, ByVal string2 As String) As Long

    'POB: fn with byte array is 17 times faster
    Dim i As Long, j As Long, bs1() As Byte, bs2() As Byte
    Dim string1_length As Long
    Dim string2_length As Long
    Dim distance() As Long
    Dim min1 As Long, min2 As Long, min3 As Long
    
    string1_length = Len(string1)
    string2_length = Len(string2)
    ReDim distance(string1_length, string2_length)
    bs1 = string1
    bs2 = string2
    
    For i = 0 To string1_length
        distance(i, 0) = i
    Next
    
    For j = 0 To string2_length
        distance(0, j) = j
    Next
    
    For i = 1 To string1_length
        For j = 1 To string2_length
            'slow way: If Mid$(string1, i, 1) = Mid$(string2, j, 1) Then
            If bs1((i - 1) * 2) = bs2((j - 1) * 2) Then   ' *2 because Unicode every 2nd byte is 0
                distance(i, j) = distance(i - 1, j - 1)
            Else
                'distance(i, j) = Application.WorksheetFunction.Min _
                (distance(i - 1, j) + 1, _
                 distance(i, j - 1) + 1, _
                 distance(i - 1, j - 1) + 1)
                ' spell it out, 50 times faster than worksheetfunction.min
                min1 = distance(i - 1, j) + 1
                min2 = distance(i, j - 1) + 1
                min3 = distance(i - 1, j - 1) + 1
                If min1 <= min2 And min1 <= min3 Then
                    distance(i, j) = min1
                ElseIf min2 <= min1 And min2 <= min3 Then
                    distance(i, j) = min2
                Else
                    distance(i, j) = min3
                End If
    
            End If
        Next
    Next
    
    Levenshtein = distance(string1_length, string2_length)

End Function

Function FuzzyMatch(ByVal string1 As String, _
                    ByVal string2 As String, _
                    Optional min_percentage As Long = 70) As String

Dim i As Long, j As Long
Dim string1_length As Long
Dim string2_length As Long
Dim distance() As Long, result As Long

string1_length = Len(string1)
string2_length = Len(string2)

' Check if not too long
If string1_length >= string2_length * (min_percentage / 100) Then
    ' Check if not too short
    If string1_length <= string2_length * ((200 - min_percentage) / 100) Then

        ReDim distance(string1_length, string2_length)
        For i = 0 To string1_length: distance(i, 0) = i: Next
        For j = 0 To string2_length: distance(0, j) = j: Next

        For i = 1 To string1_length
            For j = 1 To string2_length
                If Asc(Mid$(string1, i, 1)) = Asc(Mid$(string2, j, 1)) Then
                    distance(i, j) = distance(i - 1, j - 1)
                Else
                    distance(i, j) = Application.WorksheetFunction.Min _
                    (distance(i - 1, j) + 1, _
                     distance(i, j - 1) + 1, _
                     distance(i - 1, j - 1) + 1)
                End If
            Next
        Next
        result = distance(string1_length, string2_length) 'The distance
    End If
End If

If result <> 0 Then
    FuzzyMatch = (CLng((100 - ((result / string1_length) * 100)))) & _
                 "% (" & result & ")" 'Convert to percentage
Else
    FuzzyMatch = "Not a match"
End If

End Function

