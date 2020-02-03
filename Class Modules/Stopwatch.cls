VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stopwatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Dylan: This is not my code, all I did was turn the seconds elapsed into ms elapsed
'Code from here https://bytecomb.com/accurate-performance-timers-in-vba/

Option Explicit

#If Win64 Then
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As UINT64) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As UINT64) As Long
#Else
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As UINT64) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As UINT64) As Long
#End If

Private pFrequency As Double
Private pStartTS As UINT64
Private pEndTS As UINT64
Private pElapsed As Double
Public msElapsed As Double
Private pRunning As Boolean
   
Private Type UINT64
    LowPart As Long
    HighPart As Long
End Type

Private Const BSHIFT_32 = 4294967296# ' 2 ^ 32

Private Function U64Dbl(U64 As UINT64) As Double
    Dim lDbl As Double, hDbl As Double
    lDbl = U64.LowPart
    hDbl = U64.HighPart
    If lDbl < 0 Then lDbl = lDbl + BSHIFT_32
    If hDbl < 0 Then hDbl = hDbl + BSHIFT_32
    U64Dbl = lDbl + BSHIFT_32 * hDbl
End Function

Private Sub Class_Initialize()
    Dim PerfFrequency As UINT64
    QueryPerformanceFrequency PerfFrequency
    pFrequency = U64Dbl(PerfFrequency)
End Sub

Public Property Get Elapsed() As Double
    If pRunning Then
        Dim pNow As UINT64
        QueryPerformanceCounter pNow
        Elapsed = pElapsed + (U64Dbl(pNow) - U64Dbl(pStartTS)) / pFrequency
        Elapsed = Elapsed * 1000
    Else
        Elapsed = pElapsed
        Elapsed = Elapsed * 1000
    End If
End Property

Public Sub start()
    If Not pRunning Then
        QueryPerformanceCounter pStartTS
        pRunning = True
    End If
End Sub

Public Sub Pause()
    If pRunning Then
        QueryPerformanceCounter pEndTS
        pRunning = False
        pElapsed = pElapsed + (U64Dbl(pEndTS) - U64Dbl(pStartTS)) / pFrequency
    End If
End Sub

Public Sub Reset()
    pElapsed = 0
    pRunning = False
End Sub

Public Sub Restart()
    pElapsed = 0
    QueryPerformanceCounter pStartTS
    pRunning = True
End Sub

Public Property Get Running() As Boolean
   Running = pRunning
End Property






