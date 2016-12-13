Attribute VB_Name = "HashCodeBuilderFactory"
Option Explicit

Public Function newHashCodeBuilder(Optional initialNonZeroOddNumber As Long, Optional multiplierNonZeroOddNumber As Long) As HashCodeBuilder
    Set newHashCodeBuilder = New HashCodeBuilder
    newHashCodeBuilder.initializeVariables initialNonZeroOddNumber, multiplierNonZeroOddNumber
End Function

