Attribute VB_Name = "quikDoEvent"
'not used anymore as it made program hang after 3 or so loops

Option Explicit

#If Win64 Then
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Private Declare PtrSafe Function GetQueueStatus Lib "user32" (ByVal qsFlags As Long) As Long
Private Declare PtrSafe Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Private Declare PtrSafe Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Private Declare PtrSafe Function GetThreadPriority Lib "kernel32" (ByVal hThread As Long) As Long
Private Declare PtrSafe Function GetCurrentThread Lib "kernel32" () As Long
Private Declare PtrSafe Function GetCurrentProcess Lib "kernel32" () As Long
#Else
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetQueueStatus Lib "user32" (ByVal qsFlags As Long) As Long
Private Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Private Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Private Declare Function GetThreadPriority Lib "kernel32" (ByVal hThread As Long) As Long
Private Declare Function GetCurrentThread Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
#End If


Private Const Milliseconds_Per_Second As Long = 1000

Private Const QS_HOTKEY As Long = &H80
Private Const QS_KEY As Long = &H1
Private Const QS_MOUSEBUTTON As Long = &H4
Private Const QS_MOUSEMOVE As Long = &H2
Private Const QS_PAINT As Long = &H20
Private Const QS_POSTMESSAGE As Long = &H8
Private Const QS_SENDMESSAGE As Long = &H40
Private Const QS_TIMER As Long = &H10
Private Const QS_ALLINPUT As Long = &HFF
Private Const QS_MOUSE As Long = &H6
Private Const QS_INPUT As Long = &H7
Private Const QS_ALLEVENTS As Long = &HBF

Private Const THREAD_PRIORITY_LOWEST As Long = -2
Private Const THREAD_PRIORITY_HIGHEST As Long = 2
Private Const HIGH_PRIORITY_CLASS As Long = &H80
Private Const IDLE_PRIORITY_CLASS As Long = &H40

Private Get_Time As Long
Private Get_Temperary_Time As Long
Private Milliseconds As Long
Private Get_Frames_Per_Second As Long
Private Frame_Count As Long

Public Sub DoEvents_Fast()

    'This does events only when absolutely
    'necessary and still prevents your
    'program from locking up. The result
    'is a Do loop that is multiple times
    'faster than an ordinary Do/DoEvents
    '/Loop which is needed for realtime
    'loops. I've experimented with
    'multiple methods I've found on Planet
    'Source Code, and here are my results:
    
    'Note - This all has been done on my
    'AMD Athlon 1.2 Ghz Processor. Results
    'may vary.
    '-----------------------------------
    'Highest durations per second
    '--------------------
    'VB - 192136
    'Exe - 296140
    'Slow, slugish, and ugly for realtime.
    
    'DoEvents
    
    '----------------------------------
    'Highest durations per second
    '--------------------
    'VB - 688950
    'Exe - 735468
    'If PeekMessage(Message, 0, 0, 0, PM_NOR
    '     EMOVE) Then
    ' DoEvents
    'End If
    '---------------------------------
    'Highest durations per second
    '--------------------
    'VB - 965230
    'Exe - 1113434
    'Problem with this is that it's only
    'active when an event has occured.
    'With this I just simply held a key
    'down.
    'If GetInputState() Then
    ' DoEvents
    'End If
    '--------------------------------
    'Highest durations per second
    '--------------------
    'VB - 947204
    'Exe - 1101420
    'This is the fastest and most
    'reliable method so far.
    
    If GetQueueStatus(QS_HOTKEY Or QS_KEY Or QS_MOUSEBUTTON Or QS_PAINT) Then
        
        DoEvents
    
    End If

End Sub

'Private Sub Performance_Test()
'
'    'I've been doing a lot of
'    'experimenting on VB and learned
'    'many things on speed after doing
'    'so. Here is what I've found out:
'    ' -Expressions by themselves are
'    '2x to 3x faster than when done
'    'through a function or a sub.
'    'ex.
'    'Expression = Fix(255 / (2 ^ (5)))
'    'is way faster than
'    'Expression = Right_Bit_Shift(255, 5)
'    'even when you used a look up table.
'    'Sometimes it was 3x faster on my
'    'computer.
'    '-Any nurmeral data type is faster
'    'than working with variants.
'    'Only use variants when working with
'    'large numbers that can overflow other
'    'data types.
'    '-If statements cause slowdown.
'    'Minimize how many you use within
'    'your subs and functions.
'    'Optimizations help a lot if you have
'    'too many If statements.
'
'    Dim Expression As Long
'
'    'This loop is a true realtime loop.
'    'I've seen many ways it has been
'    'done and this is by far the fastest
'    'method
'
'    Milliseconds = GetTickCount
'
'    'This will help the ordinary
'    'DoEvents work faster.
'    '-------------------------------
'    SetThreadPriority GetCurrentThread, THREAD_PRIORITY_HIGHEST
'    SetPriorityClass GetCurrentProcess, HIGH_PRIORITY_CLASS
'    '-------------------------------
'    Do
'
'            'Um, no thank you:
'
'        'DoEvents
'
'            'Although calling this seems ok:
'
'        'DoEvents_Fast
'
'            'Inlining it is faster:
'
'        If GetQueueStatus(QS_HOTKEY Or QS_KEY Or QS_MOUSEBUTTON Or QS_PAINT) Then DoEvents
'
'        'Insert your experimental equation/function/sub
'        'etc. here. Or you can leave it
'        'empty to see the performance of
'        'the loop.
'        '-------------------------------
'        'Compare these two:
'
'            'Inline
'        'Expression = Fix(255 / 10 ^ 2)
'
'            'Function Call
'        'Expression = Right_Bit_Shift(255, 10)
'
'        '--------------------------------
'        '
'        Frame_Count = Frame_Count + 1
'
'        'If it has been a whole second...
'
'        If GetTickCount - Milliseconds >= Milliseconds_Per_Second Then
'
'            'This changes whenever it got the
'            'most durations per second,
'            'otherwise the result stays the
'            'same in the output showing it
'            'produced the most durations per
'            'second.
'
'            If Frame_Count > Get_Frames_Per_Second Then
'
'                Get_Frames_Per_Second = Frame_Count
'                Caption = Get_Frames_Per_Second & " durations per second"
'
'            End If
'
'            Frame_Count = 0
'            Milliseconds = GetTickCount
'
'        End If
'
'    Loop
'
'End Sub

Private Function Right_Bit_Shift(ByVal Value As Long, ByVal Bits_To_Shift As Long) As Long
    
    'Just a test function.
    
    Right_Bit_Shift = Fix(Value / Bits_To_Shift ^ 2)

End Function

'Private Sub Form_Activate()
'
'    Performance_Test
'
'End Sub

'Private Sub Form_Unload(Cancel As Integer)
'
'    End
'
'End Sub

