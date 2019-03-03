Attribute VB_Name = "Matrix"
Option Explicit
Option Base 1

' This module provides methods for performing various matrix operations. Rather than using a custom User-defined class, this module
' uses 2-dimensional, 1-based arrays to represent a matrix. For performance purposes, these methods assume the input is allocated,
' non-empty, 1-based (vice zero-based) and Double-valued. If these assumptions fail, the results are unpredictable.
'
'

'Private subroutine to test the various methods in this module. Completion without assertion errors implies all tests passed
Private Sub test()

    Dim mat() As Double, i As Long, j As Long
        
    mat = Matrix.diagonal(3, 4, 5)
    Debug.Assert UBound(mat, 1) = 3
    Debug.Assert UBound(mat, 2) = 4
    Debug.Assert UBound(identity(12), 1) = 12
    Debug.Assert UBound(identity(12), 2) = 12
    
    
    For i = 1 To UBound(mat, 1)
        For j = 1 To UBound(mat, 2)
            Debug.Assert mat(i, j) = IIf(i = j, 5, 0)
        Next j
    Next i
    
    Debug.Assert sameSize(identity(3), identity(3))
    Debug.Assert Not sameSize(identity(3), identity(4))
    Debug.Assert isEqual(mat, mat)
    Debug.Assert isEqual(identity(2), identity(2))
    Debug.Assert Not isEqual(identity(2), identity(3))
    Debug.Assert Not isEqual(identity(3), identity(2))
    Debug.Assert Not isEqual(randomUniform(3, 3), randomUniform(3, 3))
    Debug.Assert isEqual(transpose(transpose(mat)), mat)
    Debug.Assert trace(identity(3)) = 3
    Debug.Assert trace(identity(5)) = 5
    Debug.Assert trace(identity(39)) = 39
    Debug.Assert det(identity(12)) = 1
    Debug.Assert det(randomUniform(100, 100, 1, 2)) > 0
    Debug.Print "All tests complete: " & Now()
End Sub

' @description Given an invertable n-by-n matrix A, returns the matrix A' such that (A*A')=(A'*A)=I_n
' @param A an invertable n-by-n
' @return the inverse A' of A
Public Function inverse(A As Variant) As Variant
    Debug.Assert UBound(A, 1) = UBound(A, 2)
    det = WorksheetFunction.MInverse(A)
End Function

' @description Given square matrix A, returns the determinant of the matrix A which can be viewed as the scaling factor of the transofmration described by the matrix, c.f. https://en.wikipedia.org/wiki/Determinant
' @param A an n-by-n (square) matrix
' @return the determiniate of A
Public Function det(A As Variant) As Double
Attribute det.VB_Description = "Given square matrix A, returns the determinant of the matrix A which can be viewed as the scaling factor of the transofmration described by the matrix, c.f. https://en.wikipedia.org/wiki/Determinant"
Attribute det.VB_ProcData.VB_Invoke_Func = " \n18"
    Debug.Assert UBound(A, 1) = UBound(A, 2)
    det = WorksheetFunction.MDeterm(A)
End Function

' @description Computes the sum of the elements on the main diagonal of the matrix A, i.e. tr(A)=a_11+a_22+...+a_nn
' @param A an n-by-n (square) matrix
' @return the trace of A defined as the sum of its diagonal elements
Public Function trace(A As Variant) As Double
    Debug.Assert UBound(A, 1) = UBound(A, 2)
    Dim i As Long
    
    For i = 1 To UBound(A, 1)
        trace = trace + A(i, i)
    Next i
    
End Function
' @description performs scalar multiplication
' @param A an arbitrary sized m-by-n matrix
' @param s a double-valued scalar
' @return returns matrix C with elements c_ij=(s*a_ij)
Public Function timesScalar(A As Variant, s) As Variant
    Dim result As Variant, i As Long, j As Long
    ReDim result(UBound(A, 1), UBound(A, 2))
    
    For i = 1 To UBound(A, 1)
        For j = 1 To UBound(A, 2)
            result(i, j) = A(i, j) * s
        Next j
    Next i
    timesScalar = result
End Function

' @description performs in-place scalar multiplication. Unlike the `timesScalar` method, this method replaces A with the result of scalar multiplication
' @param A an arbitrary sized m-by-n matrix
' @param s a double-valued scalar
' @return returns A after performing scalar multiplication so that (A)_ij now equals s*a_ij.
Public Function timesScalarEquals(A As Variant, s As Double) As Variant
    Dim i As Long, j As Long
    
    For i = 1 To UBound(A, 1)
        For j = 1 To UBound(A, 2)
            A(i, j) = A(i, j) * s
        Next j
    Next i
    
    timesScalarEquals = A
End Function

' @description transposes matrix A
' @Param A an arbitrary sized m-by-n matrix
' @return returns the m-by-n transpose matrix A^t such that (A)_ij = (A^t)_ji
Public Function transpose(A As Variant) As Variant
    Dim result As Variant, i As Long, j As Long
    ReDim result(UBound(A, 1), UBound(A, 2))
    
    For i = 1 To UBound(A, 1)
        For j = 1 To UBound(A, 2)
            result(i, j) = A(i, j)
        Next j
    Next i
    transpose = result
End Function

' @description Returns true if the two matrices are equal, and returns false otherwise. More formally, given m-by-n matrix A and p-by-q matrix B, returns TRUE iff m=p, and n=q, and for all a_ij in A and b_ij in B, a_ij = b_ij. Otherwise, returns FALSE.
' @param A a m-by-n matrix
' @param B a p-by-q matrix
' @returns TRUE iff m=p and n=q and a_ij=b_ij for all i,j. Otherwise, returns FALSE.
Public Function isEqual(A As Variant, b As Variant) As Boolean
    isEqual = False
    
    If Not sameSize(A, b) Then
        Exit Function
    End If
    
    Dim i As Long, j As Long
    
    For i = 1 To UBound(A)
        For j = 1 To UBound(b)
            If (A(i, j) <> b(i, j)) Then
                Exit Function
            End If
        Next j
    Next i
    
    isEqual = True
End Function

' @description Returns the matrix product of A and B--a matrix with the same number of rows as A and same number of columns as B. See https://en.wikipedia.org/wiki/Matrix_multiplication for details of this operation
' @param A an m-by-n matrix
' @param B an n-by-p matrix
' @return the n-by-p matrix product AB
Public Function timesMatrix(A As Variant, b As Variant) As Variant
    Debug.Assert UBound(A, 2) = UBound(b, 1)
    multiply = WorksheetFunction.MMult(A, b)
End Function

' @description performs element-wise addition of two same-sized matrices
' @param A an m-by-n matrix
' @param B an m-by-n matrix
' @return the matrix C whose elements c_ij equal a_ij + b_ij
Public Function plus(A As Variant, b As Variant) As Variant
    Debug.Assert sameSize(A, b)
    
    Dim mat As Variant, i As Long, j As Long
    ReDim mat(UBound(A, 1), UBound(A, 2))
    
    For i = 1 To UBound(A, 1)
        For j = 1 To UBound(A, 2)
            mat(i, j) = A(i, j) + b(i, j)
        Next j
    Next i
    
    add = mat
    
End Function

'Returns true if A and B are the same size. Formally, given m-by-n matrix A and p-by-q matrix B, returns true iff m=p and n=q.
' @description returns TRUE if the matrices are the same size.
' @param A an m-by-n matrix
' @param B a p-by-q matrix
' @return returns TRUE iff m=p and n=q. Otherwise, returns FALSE.
Public Function sameSize(A As Variant, b As Variant) As Boolean
    sameSize = (UBound(A, 1) = UBound(b, 1)) And (UBound(A, 2) = UBound(b, 2))
End Function


' @description creates an identity matrix of size n (denoted I_n) such that for all m-by-n matrices A and n-by-m matrices B, A*I_n = A and I_n*B = B.
' @param size the numberof rows/columns (denoted by n) in the resulting identity matrix.
' @param returns the I_n identity matrix where (I_n)_ij = 1 if i=j and 0 otherwise.
Public Function identity(size As Long) As Variant
    identity = diagonal(size, size, 1#)
End Function

' @description Returns an m-by-n matrix A with the specified size whose diagonal equals the specified value. (Note: This matrix need not be square.)
' @param numberOfRows the number of rows (n) in the resulting matrix
' @param numberOfColumns the number of columns (m) in the resulting matrix
' @param value the value (v) of the elements along the diagonal.
' @return returns an m-by-n matrix A, with elements a_ij = v if i=j, and a_ij = 0 elsewhere.
Public Function diagonal(numberOfRows As Long, numberOfColumns As Long, value As Double) As Variant
    Debug.Assert numberOfRows > 0 And numberOfColumns > 0
    Dim mat() As Double, i As Long, minsize As Long
    
    minsize = IIf(numberOfRows < numberOfColumns, numberOfRows, numberOfColumns)
    ReDim mat(numberOfRows, numberOfColumns)
    
    For i = 1 To minsize
        mat(i, i) = value
    Next i
    
    diagonal = mat

End Function

' @description Returns a m-by-n matrix whose elements are uniformly distributed between the specified range of values. Formally, given a uniform distribution U(minVal,maxVal), returns an m-by-n matrix A with elements a_ij are selected from U(minVal,maxVal) If a minimum values is not specified, 0 is used. If a maximum value is not specified, then 1 is used.
' @param numberOfRows =  the number of rows (denoted m) in the resulting matrix
' @param numberOfColumns = the number of columns (denoted n) in the resulting matrix
' @param minVal = (Optional) the minimum value of the uniform distribution used to fill the resulting matrix. If not specified, 0 is used.
' @param maxVal = (Optional) the maximumvalue of the uniform distribution used to fill the resulting matrix. If not specified, 1 is used.
Public Function randomUniform(numberOfRows As Long, numberOfColumns As Long, Optional minVal As Double = 0#, Optional maxVal As Double = 1#) As Variant
    Debug.Assert numberOfRows > 0 And numberOfColumns > 0
    Debug.Assert maxVal > minVal
    
    Dim mat() As Double, i As Long, j As Long, range As Double
    ReDim mat(numberOfRows, numberOfColumns)
    
    range = maxVal - minVal
    
    For i = 1 To numberOfRows
        For j = 1 To numberOfColumns
            mat(i, j) = range * Rnd() + minVal
        Next j
    Next i
    
    randomUniform = mat
    
End Function

' @description Returns a m-by-n matrix whose elements are normally distributed. Formally, given a normal distribution N(mean,sigma), returns an m-by-n matrix A with elements a_ij are selected from N(mean,sigma)
' @param numberOfRows =  the number of rows (denoted m) in the resulting matrix
' @param numberOfColumns = the number of columns (denoted n) in the resulting matrix
' @param mean = (Optional) the mean of the normal distribution used to fill the resulting matrix. If not specified, 0 is used.
' @param standard_dev = (Optional) the standard deviation of the normal distribution used to fill the resulting matrix. If not specified, 1.0 is used.
Public Function randomNormal(numberOfRows As Long, numberOfColumns As Long, Optional mean As Double = 0#, Optional standard_Dev As Double = 1#) As Variant
    Debug.Assert numberOfRows > 0 And numberOfColumns > 0
    
    Dim mat() As Double, i As Long, j As Long
    ReDim mat(numberOfRows, numberOfColumns)
    
    For i = 1 To numberOfRows
        For j = 1 To numberOfColumns
            mat(i, j) = WorksheetFunction.NormInv(Rnd(), mean, standard_Dev)
        Next j
    Next i
    
    randomNormal = mat
    
End Function

' @description Prints to console the given matrix. This method is especially helpful when debugging code or outputing results of reasonable size (<30 rows/columns)
' @param mat the matrix to print to console
Public Sub printMatrix(mat As Variant)
        Debug.Assert Arrays.NumberOfArrayDimensions(mat) = 2
        Dim result As String, numRows As Long, numCols As Long, i As Long, j As Long, slice As Variant
        numRows = UBound(mat, 1)
        numCols = UBound(mat, 2)
        result = "["
        For i = 1 To numRows
            slice = Application.index(mat, i, 0)
            result = result & vbCrLf & "[" & Join(slice, ", ") & "] "
        Next i
        
        result = result & "]"
        
        Debug.Print result
End Sub



