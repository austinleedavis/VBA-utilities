Attribute VB_Name = "tSNE"
Option Explicit

Private Const DBL_MAX As Double = 1.79769313486231E+308
Private Const PI As Double = 3.14159265358979
Private Const DBL_MIN As Double = 1 / DBL_MAX
Private Const FLT_MAX As Double = 3.402823E+38
Private Const FLT_MIN As Double = 3.402823E-38
Private Const RAND_MAX As Long = 5

Option Base 0

' Runs tSNE on X and returns the results as an Array with N rows and out_dims columns
' Inputs:
'   X: a 2D array with N rows
'   out_dims: the desired number of output dimensions
Public Function RunQuick(ByRef X() As Double, out_dims As Long) As Double()
    
    Dim flatX() As Double, flatY() As Double, Y() As Double, i As Long, j As Long, N As Long, D As Long
    flatX = flatten(X)
    N = UBound(X, 1) + 1 ' num rows in X
    D = UBound(X, 2) + 1 ' num columns in X
    
    ReDim Y(0 To N - 1, 0 To out_dims - 1)
    ReDim flatY(0 To (N * out_dims - 1))
    
    RunFull flatX, N, D, flatY, out_dims, (N - 1) / D - 1, 0, -1
    
    simpleRun = unflatten(flatY, out_dims)
    
End Function

' Runs the tSNE algorithm. This function is exposed so users can customize the parameters of tSNE
' Inputs:
'   X: a 2D array with N rows
'   out_dims: the desired number of output dimensions
'   perplexity: (Optional)
'   theta: (Optional)
'   rand_seed: (Optional)
'   skip_random_init: (Optional)
'   max_iter: (Optional)
'   stop_lying_iter: (Optional)
'   mom_switch_iter: (Optional)
Public Function Run(ByRef fullX() As Double, out_dims As Long, _
    Optional perplexity As Double = DBL_MIN, _
    Optional theta As Double = 0, _
    Optional rand_seed As Long = 0, _
    Optional skip_random_init As Boolean = False, _
    Optional max_iter As Long = 500, _
    Optional stop_lying_iter As Long = 2, _
    Optional mom_switch_iter As Long = 180)
    
    Dim N As Long, D As Long, Y() As Double
    N = UBound(X, 1) + 1 ' num rows in X
    D = UBound(X, 2) + 1 ' num columns in X
    
    Dim X() As Double, Y() As Double
    X = flatten(fullX)
    ReDim Y(0 To (N * out_dims - 1))
    
    
    If perplexity = DBL_MIN Then
        perplexity = (N - 1) / D - 1
    End If
    
    Application.StatusBar = "Running t-SNE"
    
    'On Error GoTo whatError
    
    Dim i As Long, j As Long, lastC As Double
    
    'set random seed
    If Not skip_random_init Then
        If rand_seed >= 0 Then
            Debug.Print "Using random seed: " & rand_seed
            Randomize (rand_seed)
        Else
            Debug.Print "Using current time as random seed..."
            Randomize Now()
        End If
    End If
    
    'determine whether we are using an exact algorithm
    
    If N - 1 < 3 * perplexity Then
        Debug.Print "Perplexity too large for the number of data points!"
        Exit Function
    End If
    
    Debug.Print "Using out_dims = " & out_dims & ", perplexity = " & perplexity & ", and theta = " & theta
    Dim exact As Boolean
    exact = (theta = 0#)
    
    'set learning parameters
    Dim total_time As Double
    total_time = 0#
    Dim clock As Timer, endTime As Long
    Set clock = New Timer
    Dim momentum As Double, final_momentum As Double, eta As Double
    momentum = 0.5
    final_momentum = 0.8
    eta = 200#
    
    'allocate arrays
    Dim dY() As Double, uY() As Double, gains() As Double
    ReDim dY(0 To N * out_dims - 1)
    ReDim uY(0 To N * out_dims - 1)
    ReDim gains(0 To N * out_dims - 1)
    
    For i = 0 To N * out_dims - 1
        uY(i) = 0#
        gains(i) = 1#
    Next i
    
    'normalize input data (to prevent numerical problems
    Debug.Print "Computing input similarities..."
    clock.StartCounter
    zeroMean X, N, D
    squash X, N, D
    
    'compute input similarities for exact t-SNE
    Dim P() As Double, row_P() As Long, col_P() As Long, val_P() As Double 'these Longs are supposed to be unsigned -- watch for overflow!
    Dim sum_P As Double
    
    If (exact) Then
        
        'compute simlarities
        Debug.Print "Exact?"
        ReDim P(0 To N * N - 1)
        computeGaussianPerplexity X, N, D, P, perplexity
        
        'Symmetrize input similarities
        Debug.Print "Symmetrizing..."
        Dim nN As Long, mN As Long
        nN = 0
        For i = 0 To N - 1
            mN = (i + 1) * N
            For j = i + 1 To N - 1
                P(nN + j) = P(nN + j) + P(mN + i)
                P(mN + i) = P(nN + j)
                mN = mN + N
            Next j
            nN = nN + N
        Next i
        
        sum_P = 0#
        For i = 0 To N * N - 1
            sum_P = sum_P + P(i)
        Next i
        For i = 0 To N * N - 1
            P(i) = P(i) / sum_P
        Next i
    
    'compute input similarities for approximate t-SNE
    Else
        
        'compute asymmetric pairwise input simlarities
        computeGaussianPerplexityApprox X, N, D, row_P, col_P, val_P, perplexity, Int(3 * perplexity)
        
        'symmetrize input similarities
        symmetrizeMatrix row_P, col_P, val_P, N
        
        sum_P = 0#
        For i = 0 To N * N - 1
            sum_P = sum_P + P(i)
        Next i
        For i = 0 To N * N - 1
            P(i) = P(i) / sum_P
        Next i
        
    End If
    
    endTime = clock.TimeElapsed
    
    'lie about the p-values
    If (exact) Then
        For i = 0 To N * N - 1
            P(i) = P(i) * 12#
        Next i
    Else
        For i = 0 To row_P(N) - 1
            val_P(i) = val_P(i) * 12#
        Next i
    End If
    
    'initialize solution (randomly)
    If Not skip_random_init Then
        For i = 0 To N * out_dims - 1
            Y(i) = randn() * 0.0001
        Next i
    End If
    
    'perform main training loop
    If exact Then
        Debug.Print "Input Similarities computed in " & endTime / 1000 & " seconds!"
        Debug.Print "Learning Embedding..."
    Else
        Debug.Print "Input Similarities computed in " & endTime / 1000 & " seconds!" & "(sparsity = " & row_P(N) / (N * N)
        Debug.Print "Learning Embedding..."
    End If
    
    clock.StartCounter
    
    Dim iter As Long
    For iter = 0 To max_iter - 1
        Application.StatusBar = "Running t-SNE. Iteration " & iter + 1 & " of " & max_iter
    
        'compute (approximate) gradient
        If exact Then
            computeExactGradient P, Y, N, out_dims, dY
        Else
            computeGradient P, row_P, col_P, val_P, Y, N, out_dims, dY, theta
        End If
        
        'update gains
        For i = 0 To N * out_dims - 1
            If (sign(dY(i)) <> sign(uY(i))) Then
                gains(i) = gains(i) + 0.2
            Else
                gains(i) = gains(i) * 0.8
            End If
        Next i
        
        For i = 0 To N * out_dims - 1
            If gains(i) < 0.01 Then
                gains(i) = 0.01
            End If
        Next i
        
        'perform gradient update (with momentum and gains)
        For i = 0 To N * out_dims - 1
            uY(i) = momentum * uY(i) - eta * gains(i) * dY(i)
        Next i
        For i = 0 To N * out_dims - 1
            Y(i) = Y(i) + uY(i)
        Next i
        
        'make solution zero-mean
        zeroMean Y, N, out_dims
        
        'stop lying about the P-values after a while, and switch momentum
        If iter = stop_lying_iter Then
            If exact Then
                For i = 0 To N * N - 1
                    P(i) = P(i) / 12#
                Next i
            Else
                For i = 0 To row_P(N) - 1
                    val_P(i) = val_P(i) / 12#
                Next i
            End If
        End If
        If iter = mom_switch_iter Then
            momentum = final_momentum
        End If
        
        'print out progress
        If (iter > 0) And ((iter Mod 50) = 0 Or (iter = max_iter - 1)) Then
            Dim C As Double
            C = 0#
            If exact Then
                C = evaluateError(P, Y, N, out_dims)
            Else
                C = evaluateErrorApprox(row_P, col_P, val_P, Y, N, out_dims, theta) 'doing approximate computation here!
            End If
            If iter = 0 Then
                Debug.Print "Iteration " & iter + 1 & ": error is " & C
            Else
                total_time = total_time + clock.TimeElapsed
                Debug.Print "Iteration " & iter + 1 & ": error is " & C & " (50 iterations in " & clock.TimeElapsed / 1000 & " seconds)"
            End If
            
            If Abs(lastC - C) < 0.000001 Then
                iter = max_iter + 1
                Debug.Print "No progress"
            End If
            
            lastC = C
           
            clock.StartCounter
        End If
        
    Next iter

    
    Debug.Print "Fitting performed in " & (total_time + clock.TimeElapsed) / 1000 & " seconds"
    
    Run = unflatten(Y, out_dims)
    
End Function


'indicates if D is positive, negative or zero
Private Function sign(D As Double) As Long
    sign = IIf(D > 0, 1, IIf(D < 0, -1, 0))
End Function

'centers data on the mean
Private Sub zeroMean(ByRef X() As Double, N As Long, D As Long)
    
    Dim ni As Long, di As Long
        
    'compute data mean
    Dim mean() As Double
    ReDim mean(0 To D - 1)
    
    Dim nD As Long
    nD = 0
    For ni = 0 To N - 1
        For di = 0 To D - 1
            mean(di) = mean(di) + X(nD + di)
        Next di
        nD = nD + D
    Next ni
    
    For di = 0 To D - 1
        mean(di) = mean(di) / CDbl(N)
    Next di
    
    'subtract data mean
    nD = 0
    For ni = 0 To N - 1
        For di = 0 To D - 1
            X(nD + di) = X(nD + di) - mean(di)
        Next di
        nD = nD + D
    Next ni

End Sub

'normalizes all X values to a range of (-1,1)
Private Sub squash(ByRef X() As Double, N As Long, D As Long)
    Dim Max_X As Double, i As Long
    Max_X = 0#
    For i = 0 To N * D - 1
        If (Abs(X(i)) > Max_X) Then
            Max_X = Abs(X(i))
        End If
    Next i
    
    For i = 0 To N * D - 1
        X(i) = X(i) / Max_X
    Next i
End Sub

' returns NormInv with mean =0, stdev = 1
Private Function randn() As Double
    'randn = WorksheetFunction.NormInv(Rnd(), 0, 1)
    Dim u1 As Double, u2 As Double
    u1 = Rnd()
    u2 = Rnd()
    
    randn = (Sqr(-2 * Log(u1))) * Cos(2 * PI * u2)
End Function


Private Sub computeGaussianPerplexity(ByRef X() As Double, N As Long, D As Long, ByRef P() As Double, perplexity As Double)
    
    'compute the squared Euclidean distance matrix
    Dim DD() As Double
    ReDim DD(N * N - 1)
    
    computeSquaredEuclideanDistance X, N, D, DD
    
    'compute the gaussian kernel row by row
    Dim ni As Long, i As Long, j As Long, nN As Long
    nN = 0
    For i = 0 To N - 1
        
        'initialize some variables
        Dim found As Boolean, beta As Double, min_beta As Double, max_beta As Double, tol As Double, sum_P As Double
        found = False
        beta = 1#
        min_beta = -DBL_MAX
        max_beta = DBL_MAX
        tol = 0.00001
        
        ' Iterate until we found a good perplexity
        Dim iter As Long
        iter = 0
        While (Not found And iter < 200)
            
            'compute gaussian kernel row
            For j = 0 To N - 1
                P(nN + j) = Exp(-beta * DD(nN + j))
            Next j
            P(nN + i) = DBL_MIN
            
            'compute entropy of current row
            sum_P = DBL_MIN
            For j = 0 To N - 1
                sum_P = sum_P + P(nN + j)
            Next j
            
            Dim H As Double
            H = 0#
            For j = 0 To N - 1
                H = H + beta * (DD(nN + j) * P(nN + j))
            Next j
            H = (H / sum_P) + Log(sum_P)
            
            'evaluate whether the entropy is within the tolerance level
            Dim Hdiff As Double
            Hdiff = H - Log(perplexity)
            If (Hdiff < tol And -Hdiff < tol) Then
                found = True
            Else
                If Hdiff > 0 Then
                    min_beta = beta
                    If (max_beta = DBL_MAX) Or (max_beta = -DBL_MAX) Then
                        beta = beta * 2#
                    Else
                        beta = (beta + max_beta) / 2#
                    End If
                Else
                    max_beta = beta
                    If (min_beta = -DBL_MAX) Or (min_beta = DBL_MAX) Then
                        beta = beta / 2#
                    Else
                        beta = (beta + min_beta) / 2#
                    End If
                End If
            End If
               
            'update iteration counter
            iter = iter + 1
        
        Wend
        
        'row normalize P
        For j = 0 To N - 1
            P(nN + j) = P(nN + j) / sum_P
        Next j
        
        nN = nN + N
        
    Next i
    
End Sub

Private Sub computeSquaredEuclideanDistanceVersion1(ByRef X() As Double, N As Long, D As Long, ByRef DD() As Double)
    Dim i As Long, j As Long
    For i = 0 To N - 1
        For j = 0 To N - 1
            DD(i * N + j) = (X(i) - X(j)) * (X(i) - X(j))
        Next j
    Next i
End Sub

'computes distance between two vectors of the 2D array X
Private Function L2(X() As Double, i1 As Long, i2 As Long, D As Long) As Double
    Dim dist As Double, j As Long, x1 As Double, x2 As Double
    dist = 0#
    For j = 0 To D - 1
        x1 = X(i1 * D + j)
        x2 = X(i2 * D + j)
        dist = dist + (x1 - x2) * (x1 - x2)
    Next j
    L2 = dist
End Function

Private Sub computeSquaredEuclideanDistance(ByRef X() As Double, N As Long, D As Long, ByRef DD() As Double)
    Dim i As Long, j As Long, dist As Double
    For i = 0 To N - 1
        For j = i + 1 To N - 1
            dist = L2(X, i, j, D)
            DD(i * N + j) = dist
            DD(j * N + i) = dist
        Next j
    Next i
End Sub

Private Sub computeSquaredEuclideanDistanceVersion2(ByRef X() As Double, N As Long, D As Long, ByRef DD() As Double)
    Dim i As Long, j As Long, k As Long, XnD As Long, XmD As Long
    XnD = 0
    For i = 0 To N - 1
        XmD = XnD + D
        Dim curr_elem As Double
        curr_elem = i * N + i
        DD(curr_elem) = 0#
        Dim curr_elem_sym As Double
        curr_elem_sym = curr_elem + N
        For j = i + 1 To N - 1
            curr_elem = curr_elem + 1
            DD(curr_elem) = 0#
            
            For k = 0 To D - 1
                DD(curr_elem) = DD(curr_elem) + (DD(XnD + D) - DD(XmD + D)) * (DD(XnD + D) - DD(XmD + D))
            Next k
            curr_elem_sym = curr_elem
            
            XmD = XmD + D
            curr_elem_sym = curr_elem_sym + N
        Next j
    Next i
End Sub



'compute the t-SNE cost function (exactly)
Private Function evaluateError(P() As Double, Y() As Double, N As Long, D As Long) As Double
    
    'Compute the squared Eclidean distance matrix
    Dim DD() As Double, Q() As Double
    ReDim DD(0 To N * N - 1)
    ReDim Q(0 To N * N - 1)
    
    computeSquaredEuclideanDistance Y, N, D, DD
    
    'compute Q-matrix and normalization sum
    Dim nN As Long, sum_Q As Double, i As Long, j As Long
    sum_Q = DBL_MIN
    
    For i = 0 To N - 1
        For j = 0 To N - 1
            If (i <> j) Then
                Q(nN + j) = 1 / (1 + DD(nN + j))
                sum_Q = sum_Q + Q(nN + j)
            Else
                Q(nN + j) = DBL_MIN
            End If
        Next j
        nN = nN + N
    Next i
    
    For i = 0 To N * N - 1
        Q(i) = Q(i) / sum_Q
    Next i
    
    'sum t-SNE error
    Dim C As Double
    For i = 0 To N * N - 1
        C = C + P(i) * Log((P(i) + FLT_MIN) / (Q(i) + FLT_MIN))
    Next i
    
    evaluateError = C
    
End Function

Private Sub computeExactGradient(P() As Double, Y() As Double, N As Long, D As Long, dC() As Double)
    
    Dim i As Long, j As Long, k As Long
    
    'make sure the current gradient contains zeros
    For i = 0 To N * D - 1
        dC(i) = 0#
    Next i
    
    'compute the squared Euclidean distance matrix
    Dim DD() As Double
    ReDim DD(0 To N * N - 1)
    computeSquaredEuclideanDistance Y, N, D, DD
    
    'compute the Q-matrix and normalize sum
    Dim Q() As Double
    ReDim Q(0 To N * N - 1)
    Dim sum_Q As Double, nN As Long
    For i = 0 To N - 1
        For j = 0 To N - 1
            If (i <> j) Then
                Q(nN + j) = 1 / (1 + DD(nN + j))
                sum_Q = sum_Q + Q(nN + j)
            End If
        Next j
        nN = nN + N
    Next i
    
    'perform the computation of the gradient
    nN = 0
    Dim nD As Long
    nD = 0
    For i = 0 To N - 1
        Dim mD As Long
        mD = 0
        For j = 0 To N - 1
            If (i <> j) Then
                Dim mult As Double
                mult = (P(nN + j) - (Q(nN + j) / sum_Q)) * Q(nN + j)
                For k = 0 To D - 1
                    dC(nD + k) = dC(nD + k) + (Y(nD + k) - Y(mD + k)) * mult
                Next k
            End If
            mD = mD + D
        Next j
        nN = nN + N
        nD = nD + D
    Next i
    
End Sub

Private Function evaluateErrorApprox(ByRef row_P() As Long, ByRef col_P() As Long, ByRef val_P() As Double, ByRef Y() As Double, N As Long, D As Long, theta As Double) As Double
    'TODO
End Function


'compute input simlarities with a fixed perplexity using ball trees
Private Sub computeGaussianPerplexityApprox(X() As Double, N As Long, D As Long, row_P_() As Long, col_P_() As Long, val_P_() As Double, perplexity As Double, k As Long)
    'TODO
End Sub

Private Sub symmetrizeMatrix(row_P() As Long, col_P() As Long, val_P() As Double, N As Long)
    'TODO
End Sub

Private Sub computeGradient(P() As Double, inp_row_P() As Long, inp_col_P() As Long, inp_val_P() As Double, Y() As Double, N As Long, D As Long, dC() As Double, theta As Double)
    'TODO
End Sub



Private Function unflatten(ByRef flatX() As Double, out_dims As Long) As Double()
    Dim N As Long, i As Long, j As Long, X() As Double
    N = (UBound(flatX) + 1) / out_dims
    
    ReDim X(0 To N - 1, 0 To out_dims - 1)
    
    For i = 0 To N - 1
        For j = 0 To out_dims - 1
            X(i, j) = flatX(i * out_dims + j)
        Next j
    Next i
    
    unflatten = X
    
End Function

Private Function flatten(ByRef X() As Double) As Double()
    Dim N As Long, M As Long, F() As Double, i As Long, j As Long
    
    N = UBound(X, 1) + 1
    M = UBound(X, 2) + 1
    ReDim F(0 To (M * N - 1))
    
    For i = 0 To N - 1
        For j = 0 To M - 1
            F(i * M + j) = X(i, j)
        Next j
    Next i
    
    flatten = F
    
End Function
