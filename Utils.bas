Attribute VB_Name = "Utils"

Public Function Logistic(x, x_0, k, L)
    'The Logistic family of functions exhibit a common "S" shaped curve behavior.
    ' Inputs:
    ' x = the input value, a real number from -infty to +infty
    ' x_0 = the x-value of the sigmoid's midpoint
    ' L = the curve's maximum value
    ' k = the steepness of the curve

    Logistic = L / (1 + Exp(-k * (x - x_0)))
End Function


Public Function Sigmoid(x)
    'A special case of the logistic curve where x_0=0.5, L=1, and k=1. Sigmoids are commonly used as the activation function of artificial neurons and statistics as the CDFs
    ' Inputs:
    ' x = the input value, a real number from -infty to +infty
    Sigmoid = 1 / (1 + Exp(-x))
End Function

Public Function ReLU(x)
    'Rectifier is an activation function equal to max(0,x). ReLU is used extensively in the context of artificial neural networks.
    ReLU = IIf(x > 0, x, 0)
End Function

Public Function LeakyReLU(x, A)
    'Rectifier that allows a small, non-zero gradient when the unit is not active.
    ' Inputs:
    ' x = the input scalar value, a real number from -infty to +infty
    ' a = the coefficient of leakage
    LeakyReLU = IIf(x > 0, x, A * x)
End Function

Public Function NoisyReLU(x)
    'Rectifier that includes Gaussian noise.
    
    Dim Y As Double
    Y = WorksheetFunction.Norm_Inv(Rnd(), 0, Sigmoid(x)) + x
    NoisyReLU = IIf(Y > 0, Y, 0)
End Function

Public Function Softplus(x)
    'Softplus is a smooth approximation of the Linear rectifier function
    Softplus = Math.Log(1 + Exp(x))
End Function


Public Function CosineSimilarity(ByRef Arr1 As range, ByRef Arr2 As range)
    'Computes the Cosine Similarity Metric between two Ranges. Cosine similarity is a measure of similarity between two non-zero vectors of an inner product space that measures the cosine of the angle between them.
    
    AB = Dot(Arr1, Arr2)
    AA = Dot(Arr1, Arr1)
    BB = Dot(Arr2, Arr2)
    
    CosineSimilarity = AB / (Sqr(AA) * Sqr(BB))
    
End Function

Public Function Hamming(s As String, t As String)
    'Computes the hamming distance between two Strings. Hamming Distance measures the minimum number of substitutions required to change one string into the other, or the minimum number of errors that could have transformed one string into the other
    Dim i As Long, cost As Long
    
    If Len(s) <> Len(t) Then
        err.Raise xlErrValue
    End If
    
    cost = 0
    
    For i = 1 To Len(s)
        If Mid(s, i, 1) <> Mid(t, i, 1) Then
            cost = cost + 1
        End If
    Next i
    
    Hamming = cost
    
End Function


Public Function Levenshtein(s As String, t As String)
    'Computs the Levenshtein distance between two Strings. Levenshtein distance is a metric for measuring the difference between two Strings. Informally, the Levenshtein distance between two words is the minimum number of single-character edits (insertions, deletions or substitutions) required to change one word into the other.
    Dim v0() As Long, v1() As Long, temp() As Long, m As Long, n As Long, i As Long, j As Long, substitutionCost As Long
    m = Len(s)
    n = Len(t)
    
    'create two work vectors
    ReDim v0(0 To n + 1)
    ReDim v1(0 To n + 1)
    
    'initialize v0 (the previous row of distances
    'this row is A[0][i]: edit distance for an empty s
    'the distance is just the number of characters to delete from t
    
    For i = 0 To n
        v0(i) = i
    Next i
    
    For i = 0 To m - 1
        'calculate v1 (current row distances) from the previous row v0
        
        'first element of v1 is A[i+1][0]
        ' edit distance is delete(i+1) chars from s to match empty t
        v1(0) = i + 1
        
        'use formula to fill in the rest of the row
        For j = 0 To n - 1
            If Mid(s, i + 1, 1) = Mid(t, j + 1, 1) Then
                substitutionCost = 0
            Else
                substitutionCost = 1
            End If
            
            v1(j + 1) = WorksheetFunction.min(v1(j) + 1, v0(j + 1) + 1, v0(j) + substitutionCost)
            
        Next j
        
        'copy v1 (current row) to v0 (previous row for each iteration
        temp = v1
        v1 = v0
        v0 = temp
    Next i
    
    Levenshtein = v0(n)

End Function


Public Function Dot(ByRef A As range, ByRef B As range)
    'Computes the dot product between two ranges. Assumes ranges are equally sized
    Dot = Application.Evaluate("SUMPRODUCT(" & A.Address & "," & B.Address & ")")
End Function

