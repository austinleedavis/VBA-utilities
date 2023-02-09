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

Public Function JaccardDistance(ByVal A As String, ByVal B As String, Optional k As Long = 5) As Double
    Dim aUb As Scripting.Dictionary, aSet As Scripting.Dictionary, bSet As Scripting.Dictionary
    Dim m As Long 'length of A
    Dim N As Long 'length of B
    Dim ngram As Variant
    Dim aNb_Size As Long
    Dim i As Long
    
    If A = B Then
        JaccardDistance = 1#
        Exit Function
    End If
    
    Set aUb = New Scripting.Dictionary
    Set aSet = getNgramProfile(A, k)
    Set bSet = getNgramProfile(B, k)
    
    'compute the intersection and unions
    For Each ngram In aSet
        aUb(ngram) = ngram
    Next ngram
    For Each ngram In bSet
        aUb(ngram) = ngram
    Next ngram
    
    aNb_Size = aSet.Count + bSet.Count - aUb.Count
    JaccardDistance = aNb_Size / aUb.Count
    
    
'    For Each i In aUb
'    Next i
    
End Function

Private Function getNgramProfile(s As String, Optional k As Long = 3) As Scripting.Dictionary
    Dim i As Long
    Dim old As Long
    Dim ngram As String
    Dim ngrams As Scripting.Dictionary
    Dim string_no_space As String
    string_no_space = normalize(s, " ,./;'[]\!@#$%^&*()_") 'Replace(s, " ", "")
    Set ngrams = New Scripting.Dictionary
    
    
    For i = 1 To (Len(string_no_space) - k + 1)
        ngram = Mid(string_no_space, i, k)
        If ngrams.Exists(ngram) Then
            old = ngrams.Item(ngram)
            ngrams(ngram) = old + 1
        Else
            ngrams(ngram) = 1
        End If
        
    Next i
    
    Set getNgramProfile = ngrams
    
End Function

Private Function normalize(s As String, special_characters As String) As String
    Dim i As Long
    
    For i = 1 To Len(special_characters)
        s = replace(s, Mid(special_characters, i, 1), "")
    Next i
    
    normalize = s
End Function


'FuzzyMatch uses the Levenshtein distance to match strings in the input array to strings in the output array. The results
'are printed to the current worksheet in a 3-column output range defined prior to execution. The first column shows the
'string in the search array that is closest to the input string. The second column shows the Levenshtein distance between
'the closest match. The third column shows the proportional similarity between the Levenshtein distance and the length of
'the longer of the two strings, i.e. the input and the closest match; it is useful for giving a relative similarity among
'a large list of strings
Public Function FuzzyMatch(lookup_value As String, table_array As range) As Variant
    
    Dim C As range, result(1 To 3) As Variant, cell_value As String, min_dist As Long, best_match As String, _
    lev_dist As Long
    
    'normalize the lookup_value by removing extra spaces
    lookup_value = Trim(lookup_value)
    
    For Each C In table_array
        If (Trim(C.value) = lookup_value) Then
            result(1) = C.value
            result(2) = 0 ' zero levenshtein distance
            result(3) = 1 ' perfect match is 100% accurate
            FuzzyMatch = result
            Exit Function
        End If
    Next C
    
    'No exact match found, must compute values pairwise to determine
    min_dist = 2147483647
    best_match = xlErrNA
    
    For Each C In table_array
        cell_value = Trim(C.value)
        lev_dist = Levenshtein(lookup_value, cell_value)
        If lev_dist < min_dist Then
            min_dist = lev_dist
            best_match = C.value 'use unnormalized cell value
        End If
    Next C

    result(1) = best_match
    result(2) = min_dist
    result(3) = 1 - (min_dist) / IIf(Len(best_match) > Len(input_val), Len(best_match), Len(input_val))
    
    FuzzyMatch = result
    
End Function

'FuzzyMatch uses the Levenshtein distance to match strings in the input array to strings in the output array. The results
'are printed to the current worksheet in a 3-column output range defined prior to execution. The first column shows the
'string in the search array that is closest to the input string. The second column shows the Levenshtein distance between
'the closest match. The third column shows the proportional similarity between the Levenshtein distance and the length of
'the longer of the two strings, i.e. the input and the closest match; it is useful for giving a relative similarity among
'a large list of strings
Sub FuzzyMatch_Batch()
    
    Dim _
    input_arr As range, _
    search_arr As range, _
    ouptut_arr As range, _
    search_val As range, _
    input_val As range, _
    min_dist As Long, _
    best_match As String, _
    lev_dist As Long, _
    i As Long, _
    N As Long, _
    m As Long, _
    outputValues() As String
    

    
    Set input_arr = Application.InputBox("Select input values", "Obtain Range Object", Type:=8)
    Set search_arr = Application.InputBox("Select lookup table array", "Obtain Range Object", Type:=8)
    Set output_arr = Application.InputBox("Select Top Left corner of output range", "Obtain Range Object", Type:=8)
    
    N = input_arr.Count
    m = search_arr.Count
    
    
    If m > 500 Then
        If MsgBox("The search array you provided contains " & m & " elements. Processing " & N & " input values against this search space may take a while. Do you wish to continue?", vbYesNo, "Large Selection Detected") = vbNo Then
            Exit Sub
        End If
    End If
    
    ReDim outputValues(1 To 3, 1 To N)
    
    i = 1
    
    For Each input_val In input_arr
    
        If i Mod 10 = 0 Then
            Application.StatusBar = "Fuzzy matching item: " & i & " of " & N
            output_arr.Resize(N, 3).value = Application.transpose(outputValues)
        End If
        
        min_dist = 2147483647
        best_match = xlErrNA
        
        For Each search_val In search_arr
            If input_val.value = search_val.value Then
                min_dist = 0
                best_match = search_val.value
                GoTo ExitFor
            End If
            
            lev_dist = Levenshtein(Trim(input_val.value), Trim(search_val.value))
            If lev_dist < min_dist Then
                min_dist = lev_dist
                best_match = Trim(search_val.value)
            End If
            
        Next search_val
        
ExitFor:
        
        outputValues(1, i) = best_match
        outputValues(2, i) = min_dist
        outputValues(3, i) = 1 - (min_dist) / IIf(Len(best_match) > Len(input_val), Len(best_match), Len(input_val))

        i = i + 1
    Next input_val
    
    'Copy output to worksheet
    output_arr.Resize(N, 3).value = Application.transpose(outputValues)
    
    Application.StatusBar = ""
    
End Sub


