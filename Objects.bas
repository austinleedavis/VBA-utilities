Attribute VB_Name = "Objects"
Option Explicit

Public Sub requireNonNull(o As Variant)
    If o Is Empty Then
        err.Raise 31004, , "NullPointerException"
    End If
End Sub
