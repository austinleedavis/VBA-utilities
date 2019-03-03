VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Timer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' A high-precision timer class used to test performance of VBA code.
' This class was copied from the StackOverflow answer at http://stackoverflow.com/a/198702/3795219

Option Explicit

Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long

Private m_CounterStart As LARGE_INTEGER
Private m_CounterEnd As LARGE_INTEGER
Private m_crFrequency As Double

Private Const TWO_32 = 4294967296# ' = 256# * 256# * 256# * 256#

Private Function LI2Double(LI As LARGE_INTEGER) As Double
Dim Low As Double
    Low = LI.lowpart
    If Low < 0 Then
        Low = Low + TWO_32
    End If
    LI2Double = LI.highpart * TWO_32 + Low
End Function

Private Sub Class_Initialize()
Dim PerfFrequency As LARGE_INTEGER
    QueryPerformanceFrequency PerfFrequency
    m_crFrequency = LI2Double(PerfFrequency)
End Sub

Public Sub StartCounter()
    QueryPerformanceCounter m_CounterStart
End Sub

Property Get TimeElapsed() As Double
Dim crStart As Double
Dim crStop As Double
    QueryPerformanceCounter m_CounterEnd
    crStart = LI2Double(m_CounterStart)
    crStop = LI2Double(m_CounterEnd)
    TimeElapsed = 1000# * (crStop - crStart) / m_crFrequency
End Property
