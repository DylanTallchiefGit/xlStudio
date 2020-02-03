Attribute VB_Name = "Seq"
Option Explicit

#If Win64 Then
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Function GetQueueStatus Lib "user32" (ByVal qsFlags As Long) As Long
#Else
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetQueueStatus Lib "user32" (ByVal qsFlags As Long) As Long
#End If




Global stepsDone As Integer
Global loopsLeft As Integer
Global seqOffset As Integer
Dim counter As Integer

Dim nextTick As Long

Dim sleepTime As Variant
Dim swingL As Variant
Dim swingS As Variant
Dim nowsWatch As Variant
Dim slipOver As Variant

Dim currentDrum1 As String
Dim currentDrum2 As String
Dim currentDrum3 As String
Dim currentDrum4 As String
Dim cdc4 As Range
Dim progMode As Integer
Dim flChanCounter As Integer

Dim lastPattern As Integer


Dim kickNo As Integer
Dim snareNo As Integer
Dim CHHNo As Integer
Public sWatch As New Stopwatch
Dim imPlayin As Boolean

Dim currentVel As Integer
Dim currentPitch As Integer
Dim patternList As Integer

Dim lastletter As Variant
Dim lastInteger As Variant
Dim topLeft As Variant
Dim bottomRight As Variant
Dim patternOpen As Variant
Dim SeqOpen As Integer
Dim activeSeqColumn As Variant

Dim pasteSeqLastR As Range
Dim pasteSeqLast As String
Dim pasteSeqFirstR As Range
Dim pasteSeqFirst As String
Dim seqLastR As Range
Dim seqLast As String
Dim SvrColumn As Integer
Dim seqOn As Integer
Dim DMstartPos As Integer


'Sub seqPlay()
'
'
'
'If imPlayin = False Then
'
'startDevice 0, 0, currentPitch, 0
'seqOn = 1
'
'patternList = 1
'patternChange
'
'startLoop
'Else
'stopOkay
'
'End If
'
'End Sub



Sub patternPlay()

stopArrangement
PRstopOkay

If imPlayin = False Then

startDevice 0, 0, currentPitch, 0
seqOn = 0
'patternList = 1
'patternChange
startLoop
Else
stopOkay

End If

End Sub

Sub startLoop()
Attribute startLoop.VB_ProcData.VB_Invoke_Func = " \n14"


Worksheets("Drum Machine").Range("h28:am28").Interior.ColorIndex = 34


sWatch.Restart
nowsWatch = sWatch.Elapsed
'loopsLeft = Worksheets("Drum Machine").Range("C25").Value
slipOver = 0

'not using progMode anymore, 8 channels + 8 programs all time now
'progMode = Range("C26").Value

Dim i As Integer

For i = 0 To 31
    If LCase(Worksheets("Drum Machine").Range("H28").offset(0, i).Value) = "s" Then
    DMstartPos = i
    i = 31
    Else
    DMstartPos = 0
    End If
Next i



TestBPM 'put last


End Sub




Sub TestBPM()


imPlayin = True

sleepTime = 60 / Worksheets("Drum Machine").Range("C23").Value * 1000 / 4   '2 will need to change to 4 when i add more resolution
swingL = sleepTime + ((sleepTime / 5) * Worksheets("Drum Machine").Range("C26"))
swingS = sleepTime - ((sleepTime / 5) * Worksheets("Drum Machine").Range("C26"))

counter = 1
seqOffset = DMstartPos
stepsDone = 0

'these arent being used but i need them because they are arguments of startSeq2
Dim DMpartSelect As Integer
Dim DMvelocity As Integer
Dim DMsemitone As Integer

Dim current16th As Integer
current16th = 1

    Do
    
    
    
    If current16th > 4 Then
    current16th = 1
    End If
    
    
    stopItMIDI
    Call startSeq2(DMpartSelect, DMvelocity, DMsemitone) ' causes about ~9 ms of delay buts its okay because the While will make up for that... thats some lucky genius actually
    
    'Debug.Print "Delay: "; sWatch.Elapsed - nowsWatch '' even better time delay meter
    
    
            ''' in case lag goes over sleepTime, it will deduct from next wait to balance out
    If sWatch.Elapsed - nowsWatch > sleepTime Then
    slipOver = (sWatch.Elapsed - nowsWatch) - sleepTime
    'Debug.Print "Slip Over Activated"
    End If
    
                 'swing
    
    If current16th = 1 Or current16th = 3 Then
    sleepTime = swingL
    Else
    sleepTime = swingS
    End If
    
    
    'this kills the note after being pressed
    'While (sWatch.Elapsed - nowsWatch) - slipOver < (sleepTime / 2) * 1.5  'could change this, just never more than 2
    'Wend
    'stopItMIDI
    
    
    While (sWatch.Elapsed - nowsWatch) - slipOver < sleepTime And imPlayin = True
    'Waits here until next step
    
     DoEvents ' allows u to still click on sheet when running
    
    Wend
    
    slipOver = 0
    
    nowsWatch = sWatch.Elapsed
    
    clearCosmetics
    
    seqOffset = seqOffset + 1
    'counter = counter + 1
    stepsDone = stepsDone + 1
    current16th = current16th + 1
    
    
    If Worksheets("Drum Machine").Range("H28").offset(0, seqOffset).Value = "e" Or _
    Worksheets("Drum Machine").Range("H28").offset(0, seqOffset).Value = "l" Or Worksheets("Drum Machine").Range("H28").offset(0, seqOffset).Address = "$AN$28" Then
        If imPlayin = True Then 'stops its breaking when u perfectly time it to stop on a repeat
        
'            If Worksheets("Piano Roll").Range("H5").offset(0, PRoffset).Value = "e" Then
'            lastOffset = -1
'            End If

        stepsDone = 0
        seqOffset = DMstartPos
        
        End If
    End If
    
    Loop Until stepsDone >= 32
    
    
    
seqOffset = DMstartPos
stepsDone = 0



Looper





End Sub


Sub clearCosmetics()

'Debug.Print seqOffset

Worksheets("Drum Machine").Range("h28").offset(0, seqOffset).Interior.ColorIndex = 34   ' this creates occasional quirks of 1ms - not the end of the world but still

             'clears name boxes                            ''causes a bit over 10 ms delay
'Range("g34").Interior.Color = RGB(255, 242, 204) ' default name box color
'Range("g33").Interior.Color = RGB(255, 242, 204) ' default name box color
'Range("g32").Interior.Color = RGB(255, 242, 204) ' default name box color
'Range("g31").Interior.Color = RGB(255, 242, 204) ' default name box color



End Sub

Sub clearCosmeticsFromAR()

If seqOffset <> 0 Then
Worksheets("Drum Machine").Range("h28").offset(0, seqOffset - 1).Interior.ColorIndex = 34 ' this creates occasional quirks of 1ms - not the end of the world but still
Else
Worksheets("Drum Machine").Range("h28").offset(0, 15).Interior.ColorIndex = 34
Worksheets("Drum Machine").Range("h28").offset(0, 31).Interior.ColorIndex = 34
End If


End Sub


Public Sub stopOkay()

loopsLeft = 0
stepsDone = 32

Worksheets("Drum Machine").Range("h28:am28").Interior.ColorIndex = 34

'counter = 0
stopItMIDIAgain
imPlayin = False



End Sub




Sub Looper()



If loopsLeft > 1 Then

loopsLeft = loopsLeft - 1


If seqOn = 1 Then
patternList = patternList + 1
patternChange
Else
'Worksheets("Drum Machine").Range("B57").offset(0, patternList - 1).Interior.Color = RGB(255, 242, 204)
'patternList = patternList - 1
End If


TestBPM
seqOffset = 0

End If

imPlayin = False
               'clear x's
Worksheets("Drum Machine").Range("f31").Value = ""
Worksheets("Drum Machine").Range("f34").Value = ""
Worksheets("Drum Machine").Range("f37").Value = ""
Worksheets("Drum Machine").Range("f40").Value = ""
Worksheets("Drum Machine").Range("f43").Value = ""
Worksheets("Drum Machine").Range("f46").Value = ""
Worksheets("Drum Machine").Range("f49").Value = ""
Worksheets("Drum Machine").Range("f52").Value = ""


sWatch.Pause
'Range("B57").offset(0, patternList).Interior.ColorIndex = 0
'Worksheets("Drum Machine").Range("B57").offset(0, patternList).Interior.Color = RGB(255, 242, 204)
stopItMIDIAgain

End Sub



Sub patternChange()


'Debug.Print "Pattern List:" & patternList
Worksheets("Drum Machine").Range("B57").offset(0, patternList).Replace 0, "", xlWhole

If Worksheets("Drum Machine").Range("B57").offset(0, patternList).Value = "patternList" Then

Else

If Worksheets("Drum Machine").Range("B57").offset(0, patternList).Value = "loop" Or Worksheets("Drum Machine").Range("B57").offset(0, patternList).Value = "l" Then
Worksheets("Drum Machine").Range("B57").offset(0, patternList - 1).Interior.Color = RGB(255, 242, 204)
patternList = 1
End If

If Not Range("B57").offset(0, patternList).Value = "" Then
Worksheets("Drum Machine").Range("B57").offset(0, patternList).Interior.ColorIndex = 34
Worksheets("Drum Machine").Range("B57").offset(0, patternList - 1).Interior.Color = RGB(255, 242, 204)

patternOpen = Worksheets("Drum Machine").Range("B57").offset(0, patternList).Value


Else
'patternList = patternList - 1


End If



If Worksheets("Drum Machine").Range("B57").offset(0, patternList).Value = "" Then


patternList = patternList - 1 ' 0


Else
saveArray

topLeft = (patternOpen * 24) - 23
bottomRight = patternOpen * 24

Dim pattern1 As Variant
pattern1 = Worksheets("PatternSaver").Range("B" & topLeft & ":AI" & bottomRight).Value
Worksheets("Drum Machine").Range("F31:W54").Value = pattern1
'Worksheets("PatternSaver").Range("A4").Value = Worksheets("Drum Machine").Range("D52").Value

Worksheets("PatternSaver").Range("A4").Value = Worksheets("Drum Machine").Range("B57").offset(0, patternList).Value
End If



End If



   Select Case Worksheets("PatternSaver").Range("A4").Value
      Case 1
    '     Debug.Print "case is" & 1
         Worksheets("Drum Machine").Range("D31:D33").offset(0, 0).Interior.ColorIndex = 40
      Case 2
         Worksheets("Drum Machine").Range("D34:D36").offset(0, 0).Interior.ColorIndex = 40
      Case 3
         Worksheets("Drum Machine").Range("D37:D39").offset(0, 0).Interior.ColorIndex = 40
      Case 4
         Worksheets("Drum Machine").Range("D40:D42").offset(0, 0).Interior.ColorIndex = 40
      Case 5
         Worksheets("Drum Machine").Range("D43:D45").offset(0, 0).Interior.ColorIndex = 40
      Case 6
         Worksheets("Drum Machine").Range("D46:D48").offset(0, 0).Interior.ColorIndex = 40
      Case 7
         Worksheets("Drum Machine").Range("D49:D51").offset(0, 0).Interior.ColorIndex = 40
      Case Else
  '       Debug.Print "case says" & Worksheets("PatternSaver").Range("A4").Value
        Worksheets("Drum Machine").Range("D52:D54").offset(0, 0).Interior.ColorIndex = 40
        Worksheets("Drum Machine").Range("D52").Value = Worksheets("PatternSaver").Range("A4").Value
   End Select



End Sub

Sub oopsie()

'Worksheets("PatternSaver").Range("A4").Value = 1
array1
loopsLeft = 0
stepsDone = 32
'sWatch.Pause
imPlayin = False
stopItMIDI  'this calls the macro in playMidi



End Sub



Sub startSeq2(DMpartSelect, DMvelocity, DMsemitone)

If DMvelocity = 0 Then
DMvelocity = 100
End If

If seqOffset = 0 Then
seqOffset = seqOffset + DMpartSelect
End If

 DoEvents ' allows u to still click on sheet when running


Dim cdc1 As Range
Dim cdc2 As Range
Dim cdc3 As Range
Dim cdc4 As Range
Dim cdc5 As Range
Dim cdc6 As Range
Dim cdc7 As Range
Dim cdc8 As Range


Dim currentDrum1 As String
Dim currentDrum2 As String
Dim currentDrum3 As String
Dim currentDrum4 As String
Dim currentDrum5 As String
Dim currentDrum6 As String
Dim currentDrum7 As String
Dim currentDrum8 As String
                                

    
'    Range("f28").Value = sngWaitEnd - Timer

Worksheets("Drum Machine").Range("h28").offset(0, seqOffset).Interior.ColorIndex = 41 ''this is the current step display bar
 ' Range("h28").offset(0, seqOffset).Borders.LineStyle = xlDouble  ''made them all borded by accident but i liked it


''writing values (like the "x") to cells is what "grey locks" the code from running, changed to color instead



              ' looks for X on currentDrum1 cells
currentDrum1 = CStr(Worksheets("Drum Machine").Range("h28").offset(3, seqOffset))
'Worksheets("Drum Machine").Range("f31").Value = currentDrum1 ' shows X when kick  playing
'If currentDrum1 = "x" Then
''Worksheets("Drum Machine").Range("F31").Interior.Color = RGB(252, 228, 214)
'Worksheets("Drum Machine").Range("F31").Interior.ColorIndex = 40
'Else
'Worksheets("Drum Machine").Range("F31").Interior.Color = RGB(248, 203, 173)
'End If
Set cdc1 = Worksheets("Drum Machine").Range("h28").offset(3, seqOffset)

               ' looks for X on currentDrum2 cells
currentDrum2 = CStr(Worksheets("Drum Machine").Range("h28").offset(6, seqOffset))
Set cdc2 = Worksheets("Drum Machine").Range("h28").offset(6, seqOffset)

                 ' looks for X on currentDrum3
currentDrum3 = CStr(Worksheets("Drum Machine").Range("h28").offset(9, seqOffset))
Set cdc3 = Worksheets("Drum Machine").Range("h28").offset(9, seqOffset)

                 ' looks for X on currentDrum4
currentDrum4 = CStr(Worksheets("Drum Machine").Range("h28").offset(12, seqOffset))
Set cdc4 = Worksheets("Drum Machine").Range("h28").offset(12, seqOffset)



currentDrum5 = CStr(Worksheets("Drum Machine").Range("h28").offset(15, seqOffset))
Set cdc5 = Worksheets("Drum Machine").Range("h28").offset(15, seqOffset)

currentDrum6 = CStr(Worksheets("Drum Machine").Range("h28").offset(18, seqOffset))
Set cdc6 = Worksheets("Drum Machine").Range("h28").offset(18, seqOffset)

currentDrum7 = CStr(Worksheets("Drum Machine").Range("h28").offset(21, seqOffset))
Set cdc7 = Worksheets("Drum Machine").Range("h28").offset(21, seqOffset)

currentDrum8 = CStr(Worksheets("Drum Machine").Range("h28").offset(24, seqOffset))
Set cdc8 = Worksheets("Drum Machine").Range("h28").offset(24, seqOffset)




If Left(currentDrum1, 1) = "x" Then

                      'gets Vel
If IsEmpty(Worksheets("Drum Machine").Range("h28").offset(4, seqOffset)) = False And Left(Worksheets("Drum Machine").Range("h28").offset(4, seqOffset).Value, 1) <> " " Then
currentVel = (Worksheets("Drum Machine").Range("h28").offset(4, seqOffset) / 100) * DMvelocity
Else
currentVel = (Worksheets("Drum Machine").Range("f32").Value / 100) * DMvelocity
End If

                    'gets Pitch
If IsEmpty(Range("h28").offset(5, seqOffset)) = False And Left(Worksheets("Drum Machine").Range("h28").offset(5, seqOffset).Value, 1) <> " " Then
currentPitch = Worksheets("Drum Machine").Range("h28").offset(5, seqOffset) + DMsemitone
Else
currentPitch = Worksheets("Drum Machine").Range("f33").Value + DMsemitone
End If

If Worksheets("Drum Machine").Range("C24").Value = 1 Then
midiNote 0, currentVel, currentPitch, 9
Else
'If progMode > 0 Then
'midiNote 0, currentVel, currentPitch, 0
'Else
'midiNote 0, currentVel, currentPitch, 0
'End If
midiNote 0, currentVel, currentPitch, 0
End If

End If

If Left(currentDrum2, 1) = "x" Then
'Snare (snareNo)
'snareNo = snareNo + 1
' TestSnare
'mciPlaySnare
'Range("g33").Interior.Color = RGB(252, 208, 214) ' light up name boxes

                      'gets Vel
If IsEmpty(Worksheets("Drum Machine").Range("h28").offset(7, seqOffset)) = False And Left(Worksheets("Drum Machine").Range("h28").offset(7, seqOffset).Value, 1) <> " " Then
currentVel = (Worksheets("Drum Machine").Range("h28").offset(7, seqOffset) / 100) * DMvelocity
Else
currentVel = (Worksheets("Drum Machine").Range("f35").Value / 100) * DMvelocity
End If

                    'gets Pitch
If IsEmpty(Worksheets("Drum Machine").Range("h28").offset(8, seqOffset)) = False And Left(Worksheets("Drum Machine").Range("h28").offset(8, seqOffset).Value, 1) <> " " Then
currentPitch = Worksheets("Drum Machine").Range("h28").offset(8, seqOffset) + DMsemitone
Else
currentPitch = Worksheets("Drum Machine").Range("f36").Value + DMsemitone
End If

If Worksheets("Drum Machine").Range("C24").Value = 1 Then
midiNote 0, currentVel, currentPitch, 9
Else
'If progMode > 0 Then
'midiNote 1, currentVel, currentPitch, 0
'Else
'midiNote 0, currentVel, currentPitch, 1
'End If
midiNote 1, currentVel, currentPitch, 1
End If
End If


If Left(currentDrum3, 1) = "x" Then

                      'gets Vel
If IsEmpty(Worksheets("Drum Machine").Range("h28").offset(10, seqOffset)) = False And Left(Worksheets("Drum Machine").Range("h28").offset(10, seqOffset).Value, 1) <> " " Then
currentVel = (Worksheets("Drum Machine").Range("h28").offset(10, seqOffset) / 100) * DMvelocity
Else
currentVel = (Worksheets("Drum Machine").Range("f38").Value / 100) * DMvelocity
End If

                    'gets Pitch
If IsEmpty(Worksheets("Drum Machine").Range("h28").offset(11, seqOffset)) = False And Left(Worksheets("Drum Machine").Range("h28").offset(11, seqOffset).Value, 1) <> " " Then
currentPitch = Worksheets("Drum Machine").Range("h28").offset(11, seqOffset) + DMsemitone
Else
currentPitch = Worksheets("Drum Machine").Range("f39").Value + DMsemitone
End If

If Worksheets("Drum Machine").Range("C24").Value = 1 Then
midiNote 0, currentVel, currentPitch, 9
Else
'If progMode > 0 Then
'midiNote 2, currentVel, currentPitch, 0
'Else
'midiNote 0, currentVel, currentPitch, 2
'End If
midiNote 2, currentVel, currentPitch, 2
End If

End If


If Left(currentDrum4, 1) = "x" Then
'lastLetter = Right("currentDrum4", 1)
'lastInteger = ConvertString(lastLetter)
'cdc4.offset(14, seqOffset).Color = 0

                      'gets Vel
If IsEmpty(cdc4.offset(1, 0)) = False And Left(cdc4.offset(1, 0).Value, 1) <> " " Then
currentVel = (cdc4.offset(1, 0) / 100) * DMvelocity
Else
currentVel = (Worksheets("Drum Machine").Range("f41").Value / 100) * DMvelocity
End If

                    'gets Pitch
'Debug.Print cdc4.offset(2, 0).Value
If IsEmpty(cdc4.offset(2, 0)) = False And Left(cdc4.offset(2, 0).Value, 1) <> " " Then
currentPitch = cdc4.offset(2, 0) + DMsemitone
Else
currentPitch = Worksheets("Drum Machine").Range("f42").Value + DMsemitone
End If



If Worksheets("Drum Machine").Range("C24").Value = 1 Then
midiNote 0, currentVel, currentPitch, 9
Else
'If progMode > 0 Then
'midiNote 3, currentVel, currentPitch, 0
'Else
'midiNote 0, currentVel, currentPitch, 3
'End If
midiNote 3, currentVel, currentPitch, 3
End If

End If

If Left(cdc5.Value, 1) = "x" Then
                      'gets Vel
If IsEmpty(cdc5.offset(1, 0)) = False And Left(cdc5.offset(1, 0).Value, 1) <> " " Then
currentVel = (cdc5.offset(1, 0) / 100) * DMvelocity
Else
currentVel = (Worksheets("Drum Machine").Range("f44").Value / 100) * DMvelocity
End If
                    'gets Pitch
If IsEmpty(cdc5.offset(2, 0)) = False And Left(cdc5.offset(2, 0).Value, 1) <> " " Then
currentPitch = cdc5.offset(2, 0) + DMsemitone
Else
currentPitch = Worksheets("Drum Machine").Range("f45").Value + DMsemitone
End If

If Worksheets("Drum Machine").Range("C24").Value = 1 Then
midiNote 0, currentVel, currentPitch, 9
Else
'If progMode > 0 Then
'midiNote 4, currentVel, currentPitch, 0
'Else
'midiNote 0, currentVel, currentPitch, 4
'End If
midiNote 4, currentVel, currentPitch, 4
End If
End If


If Left(cdc6.Value, 1) = "x" Then
'cdc6.offset(2, 0).Value = "test"
                      'gets Vel
If IsEmpty(cdc6.offset(1, 0)) = False And Left(cdc6.offset(1, 0).Value, 1) <> " " Then
currentVel = (cdc6.offset(1, 0) / 100) * DMvelocity
Else
currentVel = (Worksheets("Drum Machine").Range("f47").Value / 100) * DMvelocity
End If
                    'gets Pitch
If IsEmpty(cdc6.offset(2, 0)) = False And Left(cdc6.offset(2, 0).Value, 1) <> " " Then
currentPitch = cdc6.offset(2, 0) + DMsemitone
Else
currentPitch = Worksheets("Drum Machine").Range("f48").Value + DMsemitone
End If


If Worksheets("Drum Machine").Range("C24").Value = 1 Then
midiNote 0, currentVel, currentPitch, 9
Else
'If progMode > 0 Then
'midiNote 5, currentVel, currentPitch, 0
'Else
'midiNote 0, currentVel, currentPitch, 5
'End If
midiNote 5, currentVel, currentPitch, 5
End If

End If


If Left(cdc7.Value, 1) = "x" Then
                      'gets Vel
If IsEmpty(cdc7.offset(1, 0)) = False And Left(cdc7.offset(1, 0).Value, 1) <> " " Then
currentVel = (cdc7.offset(1, 0) / 100) * DMvelocity
Else
currentVel = (Worksheets("Drum Machine").Range("f50").Value / 100) * DMvelocity
End If
                    'gets Pitch
If IsEmpty(cdc7.offset(2, 0)) = False And Left(cdc7.offset(2, 0).Value, 1) <> " " Then
currentPitch = cdc7.offset(2, 0) + DMsemitone
Else
currentPitch = Worksheets("Drum Machine").Range("f51").Value + DMsemitone
End If

If Worksheets("Drum Machine").Range("C24").Value = 1 Then
midiNote 0, currentVel, currentPitch, 9
Else
'If progMode > 0 Then
'midiNote 6, currentVel, currentPitch, 0
'Else
'midiNote 0, currentVel, currentPitch, 6
'End If
midiNote 6, currentVel, currentPitch, 6
End If

End If

If Left(cdc8.Value, 1) = "x" Then
                      'gets Vel
If IsEmpty(cdc8.offset(25, 0)) = False And Left(cdc8.offset(1, 0).Value, 1) <> " " Then
currentVel = (cdc5.offset(25, 0) / 100) * DMvelocity
Else
currentVel = (Worksheets("Drum Machine").Range("f53").Value / 100) * DMvelocity
End If
                    'gets Pitch
If IsEmpty(cdc8.offset(26, 0)) = False And Left(cdc8.offset(2, 0).Value, 1) <> " " Then
currentPitch = cdc5.offset(26, 0) + DMsemitone
Else
currentPitch = Worksheets("Drum Machine").Range("f54").Value + DMsemitone
End If

If Worksheets("Drum Machine").Range("C24").Value = 1 Then
midiNote 0, currentVel, currentPitch, 9
Else
'If progMode > 0 Then
'midiNote 7, currentVel, currentPitch, 0
'Else
'midiNote 0, currentVel, currentPitch, 7
'End If
midiNote 7, currentVel, currentPitch, 7
End If

End If




'  Sleep sleepTime  '' this needs to change depending on bpm  (1000 is 1 sec)   ''sleep is BAD, even application.wait stays on time but sleep doesnt, try timeGetTime instead


'DoEvents_Fast   ''downloaded from internet, supposed to be way better, its better but not perfect  ' didnt use because would make program freeze after 3 loops



End Sub







Sub clearPattern()
Worksheets("Drum Machine").Range("H31:AM54").Value = ""

End Sub

Sub clearDMcounter()
Worksheets("Drum Machine").Range("H28:AM28").Value = ""

End Sub





Public Sub saveArray()

lastPattern = Worksheets("PatternSaver").Range("A4").Value

topLeft = (lastPattern * 24) - 23
bottomRight = lastPattern * 24

Worksheets("PatternSaver").Range("B" & topLeft & ":AI" & bottomRight).Value = Worksheets("Drum Machine").Range("F31:AM54").Value
Worksheets("Drum Machine").Range("D31:D54").Interior.ColorIndex = 0  ''removes all colour from selected patterns

End Sub


Public Sub array1()
lastletter = Right("array1", 1)
lastInteger = ConvertString(lastletter)

'Dim topLeft As Range
'Set topLeft = Cells(1, 2)
'
'Dim bottomRight As Range
'Set bottomRight = Cells(12, 17)

saveArray

topLeft = lastInteger
bottomRight = lastInteger * 24



Dim pattern1 As Variant
pattern1 = Worksheets("PatternSaver").Range("B" & topLeft & ":AI" & bottomRight).Value
'pattern1 = Worksheets("PatternSaver").Range(topLeft, bottomRight).Value
Worksheets("Drum Machine").Range("F31:AM54").Value = pattern1
Worksheets("PatternSaver").Range("A4").Value = lastInteger  ''stores in pattern saver for next saving

Worksheets("Drum Machine").Range("D31:D33").offset(0, 0).Interior.ColorIndex = 40


End Sub


Public Sub array2()


lastletter = Right("array2", 1)
lastInteger = ConvertString(lastletter)

saveArray

topLeft = (lastInteger * 24) - 23
bottomRight = lastInteger * 24

Dim pattern1 As Variant
pattern1 = Worksheets("PatternSaver").Range("B" & topLeft & ":AI" & bottomRight).Value
Worksheets("Drum Machine").Range("F31:AM54").Value = pattern1

Worksheets("PatternSaver").Range("A4").Value = lastInteger
Range("D34:D36").offset(0, 0).Interior.ColorIndex = 40


End Sub

Public Sub array3()
lastInteger = 3

saveArray

topLeft = (lastInteger * 24) - 23
bottomRight = lastInteger * 24

'saveArray

Dim pattern1 As Variant
pattern1 = Worksheets("PatternSaver").Range("B" & topLeft & ":AI" & bottomRight).Value
Worksheets("Drum Machine").Range("F31:AM54").Value = pattern1
Worksheets("PatternSaver").Range("A4").Value = lastInteger
Range("D37:D39").offset(0, 0).Interior.ColorIndex = 40

End Sub


Public Sub array4()
lastInteger = 4
saveArray

topLeft = (lastInteger * 24) - 23
bottomRight = lastInteger * 24

Dim pattern1 As Variant
pattern1 = Worksheets("PatternSaver").Range("B" & topLeft & ":AI" & bottomRight).Value
Worksheets("Drum Machine").Range("F31:AM54").Value = pattern1
Worksheets("PatternSaver").Range("A4").Value = lastInteger
Range("D40:D42").offset(0, 0).Interior.ColorIndex = 40

End Sub

Public Sub array5()
lastInteger = 5
saveArray

topLeft = (lastInteger * 24) - 23
bottomRight = lastInteger * 24

Dim pattern1 As Variant
pattern1 = Worksheets("PatternSaver").Range("B" & topLeft & ":AI" & bottomRight).Value
Worksheets("Drum Machine").Range("F31:AM54").Value = pattern1
Worksheets("PatternSaver").Range("A4").Value = lastInteger
Range("D43:D45").offset(0, 0).Interior.ColorIndex = 40
End Sub

Public Sub array6()
lastInteger = 6
saveArray

topLeft = (lastInteger * 24) - 23
bottomRight = lastInteger * 24

Dim pattern1 As Variant
pattern1 = Worksheets("PatternSaver").Range("B" & topLeft & ":AI" & bottomRight).Value
Worksheets("Drum Machine").Range("F31:AM54").Value = pattern1
Worksheets("PatternSaver").Range("A4").Value = lastInteger
Range("D46:D48").offset(0, 0).Interior.ColorIndex = 40
End Sub

Public Sub array7()
lastInteger = 7
saveArray

topLeft = (lastInteger * 24) - 23
bottomRight = lastInteger * 24

Dim pattern1 As Variant
pattern1 = Worksheets("PatternSaver").Range("B" & topLeft & ":AI" & bottomRight).Value
Worksheets("Drum Machine").Range("F31:AM54").Value = pattern1
Worksheets("PatternSaver").Range("A4").Value = lastInteger
Range("D49:D51").offset(0, 0).Interior.ColorIndex = 40
End Sub

Public Sub openPattern()

If Worksheets("Drum Machine").Range("D52").Value < 1 Then

Else

lastInteger = Worksheets("Drum Machine").Range("D52").Value
saveArray

topLeft = (lastInteger * 24) - 23
bottomRight = lastInteger * 24

Dim pattern1 As Variant
pattern1 = Worksheets("PatternSaver").Range("B" & topLeft & ":AI" & bottomRight).Value
Worksheets("Drum Machine").Range("F31:AM54").Value = pattern1
Worksheets("PatternSaver").Range("A4").Value = Worksheets("Drum Machine").Range("D52").Value
Worksheets("Drum Machine").Range("D52:D54").offset(0, 0).Interior.ColorIndex = 40
End If

End Sub

Sub seqUpPat()

Worksheets("Drum Machine").Range("D52").Value = Worksheets("Drum Machine").Range("D52").Value + 1

openPattern

End Sub


Sub seqDownPat()

If Worksheets("Drum Machine").Range("D52").Value >= 2 Then

Worksheets("Drum Machine").Range("D52").Value = Worksheets("Drum Machine").Range("D52").Value - 1

openPattern

End If

End Sub


'copied off internet   : https://excel.officetuts.net/en/vba/convert-string-to-integer
Function ConvertString(myString)
    Dim finalNumber As Variant
    If IsNumeric(myString) Then
        If IsEmpty(myString) Then
            finalNumber = "-"
        Else
            finalNumber = CInt(myString)
        End If
    Else
        finalNumber = "-"
    End If
    
    ConvertString = finalNumber
End Function



Sub copyPattern()
Worksheets("PatternSaver").Range("BA1:CH24").Value = Worksheets("Drum Machine").Range("F31:AM54").Value
Worksheets("Drum Machine").Range("Q22:R22").Interior.ColorIndex = 40

'Worksheets("Drum Machine").Range("Q25:R25").Interior.ColorIndex = 0
Worksheets("Drum Machine").Range("Q25:R25").Interior.Color = RGB(226, 239, 218)


End Sub

Sub pastePattern()
Worksheets("Drum Machine").Range("F31:AM54").Value = Worksheets("PatternSaver").Range("BA1:CH24").Value
Worksheets("Drum Machine").Range("Q25:R25").Interior.ColorIndex = 40
'Worksheets("Drum Machine").Range("Q22:R22").Interior.ColorIndex = 0
Worksheets("Drum Machine").Range("Q22:R22").Interior.Color = RGB(226, 239, 218)

End Sub


Sub co2String()

'Debug.Print Range("AM2").Row & ", " & Range("AM2").Column

End Sub

Sub duplicateDM()

Worksheets("Drum Machine").Range("X31:AM54").Value = Worksheets("Drum Machine").Range("H31:W54").Value


End Sub


'Sub saveSeq()
'Dim delSvrR As Range
'Dim delSvr As String
'
'SeqOpen = Worksheets("PatternSaver").Range("T4").Value
'
'Set seqLastR = Cells(57, 2 + activeSeqColumn)
'seqLast = seqLastR.Address(RowAbsolute:=False, ColumnAbsolute:=False)
'
'Set pasteSeqFirstR = Cells(2, SeqOpen + 38)
'pasteSeqFirst = pasteSeqFirstR.Address(RowAbsolute:=False, ColumnAbsolute:=False)
'
'Set pasteSeqLastR = Cells(activeSeqColumn + 1, SeqOpen + 38)
'pasteSeqLast = pasteSeqLastR.Address(RowAbsolute:=False, ColumnAbsolute:=False)
'
'Set delSvrR = Cells(SvrColumn + 1, SeqOpen + 38)
'delSvr = delSvrR.Address(RowAbsolute:=False, ColumnAbsolute:=False)
'
'Worksheets("PatternSaver").Range(pasteSeqFirst, delSvr).Value = ""
'
''MsgBox testRange.Address(RowAbsolute:=False, ColumnAbsolute:=False)   'yesss game changing
'
''Range("J61:J" & 9 + activeSeqColumn).Value = Range("B57:B" & activeSeqColumn + 2).Value
'
'Worksheets("PatternSaver").Range(pasteSeqFirst, pasteSeqLast).Value = Application.WorksheetFunction.Transpose(Worksheets("Drum Machine").Range("C57:" & seqLast).Value)
'Worksheets("PatternSaver").Range(pasteSeqFirst, pasteSeqLast).Replace "#N/A", "", xlWhole
'' Range(Cells(60, 14), Cells(70, 17)).Value ' this actually works
'
'End Sub
'
'
'
'Sub openSeq1()
'Dim delSeqR As Range
'Dim delSeq As String
'
'SeqOpen = Worksheets("PatternSaver").Range("T4").Value
''SeqOpen = SeqOpen + 2
''Debug.Print SeqOpen
''Set delSeqR = C
'Worksheets("Drum Machine").Cells(60, SeqOpen + 2).Interior.ColorIndex = 0
'
'countSeqColumn
'saveSeq
'SeqOpen = 1
'countSvrColumn
'
'Worksheets("PatternSaver").Range("T4").Value = SeqOpen
'
'
'Set pasteSeqFirstR = Cells(2, SeqOpen + 38)
'pasteSeqFirst = pasteSeqFirstR.Address(RowAbsolute:=False, ColumnAbsolute:=False)
'
'Set pasteSeqLastR = Cells(SvrColumn + 1, SeqOpen + 38)
'pasteSeqLast = pasteSeqLastR.Address(RowAbsolute:=False, ColumnAbsolute:=False)
''activeSeqColumn
'Set seqLastR = Cells(57, 2 + SvrColumn)
'seqLast = seqLastR.Address(RowAbsolute:=False, ColumnAbsolute:=False)
'
'Set delSeqR = Cells(57, activeSeqColumn + 3)
'delSeq = delSeqR.Address(RowAbsolute:=False, ColumnAbsolute:=False)
'Worksheets("Drum Machine").Range("C57", delSeq).Value = ""
'
''Debug.Print "seqLast is" & seqLast
'Worksheets("Drum Machine").Range("C57:" & seqLast).Value = Application.WorksheetFunction.Transpose(Worksheets("PatternSaver").Range(pasteSeqFirst, pasteSeqLast).Value)
'Worksheets("Drum Machine").Range("C57:" & seqLast).Replace "#N/A", "", xlWhole
'Worksheets("Drum Machine").Range("C57:" & seqLast).Replace 0, "", xlWhole
'
''If Worksheets("Drum Machine").Range("C57:" & seqLast).Value = "#N/A" Then
''Worksheets("Drum Machine").Range("C57:" & seqLast).Value = ""
''End If
'
'Range("C60").Interior.ColorIndex = 40
'
'
'
'End Sub
'
'Sub openSeq2()
'Dim delSeqR As Range
'Dim delSeq As String
'
'SeqOpen = Worksheets("PatternSaver").Range("T4").Value
'Worksheets("Drum Machine").Cells(60, SeqOpen + 2).Interior.ColorIndex = 0
'
'
'countSeqColumn
'saveSeq
'SeqOpen = 2
'countSvrColumn
'
'
'Worksheets("PatternSaver").Range("T4").Value = SeqOpen
'
'
'Set pasteSeqFirstR = Cells(2, SeqOpen + 38)
'pasteSeqFirst = pasteSeqFirstR.Address(RowAbsolute:=False, ColumnAbsolute:=False)
'
'Set pasteSeqLastR = Cells(SvrColumn + 1, SeqOpen + 38)
'pasteSeqLast = pasteSeqLastR.Address(RowAbsolute:=False, ColumnAbsolute:=False)
'
'Set seqLastR = Cells(57, 2 + SvrColumn)
'seqLast = seqLastR.Address(RowAbsolute:=False, ColumnAbsolute:=False)
'
'Set delSeqR = Cells(57, activeSeqColumn + 3)
'delSeq = delSeqR.Address(RowAbsolute:=False, ColumnAbsolute:=False)
'Worksheets("Drum Machine").Range("C57", delSeq).Value = ""
'
'
'Worksheets("Drum Machine").Range("C57:" & seqLast).Value = Application.WorksheetFunction.Transpose(Worksheets("PatternSaver").Range(pasteSeqFirst, pasteSeqLast).Value)
'Worksheets("Drum Machine").Range("C57:" & seqLast).Replace "#N/A", "", xlWhole
'Worksheets("Drum Machine").Range("C57:" & seqLast).Replace 0, "", xlWhole
'
'
'Range("D60").Interior.ColorIndex = 40
'
'End Sub
'
'Sub openSeq3()
'Dim delSeqR As Range
'Dim delSeq As String
'
'SeqOpen = Worksheets("PatternSaver").Range("T4").Value
'Worksheets("Drum Machine").Cells(60, SeqOpen + 2).Interior.ColorIndex = 0
'
'countSeqColumn
'saveSeq
'SeqOpen = 3
'countSvrColumn
'
'
'Worksheets("PatternSaver").Range("T4").Value = SeqOpen
'
'
'Set pasteSeqFirstR = Cells(2, SeqOpen + 38)
'pasteSeqFirst = pasteSeqFirstR.Address(RowAbsolute:=False, ColumnAbsolute:=False)
'
'Set pasteSeqLastR = Cells(SvrColumn + 1, SeqOpen + 38)
'pasteSeqLast = pasteSeqLastR.Address(RowAbsolute:=False, ColumnAbsolute:=False)
'
'Set seqLastR = Cells(57, 2 + SvrColumn)
'seqLast = seqLastR.Address(RowAbsolute:=False, ColumnAbsolute:=False)
'
'Set delSeqR = Cells(57, activeSeqColumn + 3)
'delSeq = delSeqR.Address(RowAbsolute:=False, ColumnAbsolute:=False)
'Worksheets("Drum Machine").Range("C57", delSeq).Value = ""
'
'
'Worksheets("Drum Machine").Range("C57:" & seqLast).Value = Application.WorksheetFunction.Transpose(Worksheets("PatternSaver").Range(pasteSeqFirst, pasteSeqLast).Value)
'Worksheets("Drum Machine").Range("C57:" & seqLast).Replace "#N/A", "", xlWhole
'Worksheets("Drum Machine").Range("C57:" & seqLast).Replace 0, "", xlWhole
'
'
'
'Worksheets("Drum Machine").Range("E60").Interior.ColorIndex = 40
'
'
'End Sub
'
'Sub openSeq4()
'Dim delSeqR As Range
'Dim delSeq As String
'
'SeqOpen = Worksheets("PatternSaver").Range("T4").Value
'Worksheets("Drum Machine").Cells(60, SeqOpen + 2).Interior.ColorIndex = 0
'
'countSeqColumn
'saveSeq
'SeqOpen = 4
'countSvrColumn
'
'
'Worksheets("PatternSaver").Range("T4").Value = SeqOpen
'
'
'Set pasteSeqFirstR = Cells(2, SeqOpen + 38)
'pasteSeqFirst = pasteSeqFirstR.Address(RowAbsolute:=False, ColumnAbsolute:=False)
'
'Set pasteSeqLastR = Cells(SvrColumn + 1, SeqOpen + 38)
'pasteSeqLast = pasteSeqLastR.Address(RowAbsolute:=False, ColumnAbsolute:=False)
'
'Set seqLastR = Cells(57, 2 + SvrColumn)
'seqLast = seqLastR.Address(RowAbsolute:=False, ColumnAbsolute:=False)
'
'Set delSeqR = Cells(57, activeSeqColumn + 3)
'delSeq = delSeqR.Address(RowAbsolute:=False, ColumnAbsolute:=False)
'Worksheets("Drum Machine").Range("C57", delSeq).Value = ""
'
'
'Worksheets("Drum Machine").Range("C57:" & seqLast).Value = Application.WorksheetFunction.Transpose(Worksheets("PatternSaver").Range(pasteSeqFirst, pasteSeqLast).Value)
'Worksheets("Drum Machine").Range("C57:" & seqLast).Replace "#N/A", "", xlWhole
'Worksheets("Drum Machine").Range("C57:" & seqLast).Replace 0, "", xlWhole
'
'
'Worksheets("Drum Machine").Range("F60").Interior.ColorIndex = 40
'
'End Sub
'
'
'Sub openSeqX()
'Dim delSeqR As Range
'Dim delSeq As String
'
'SeqOpen = Worksheets("PatternSaver").Range("T4").Value
'Worksheets("Drum Machine").Cells(60, SeqOpen + 2).Interior.ColorIndex = 0
'
'countSeqColumn
'saveSeq
'SeqOpen = Range("H61").Value
'countSvrColumn
'
'
'Worksheets("PatternSaver").Range("T4").Value = SeqOpen
'
'
'Set pasteSeqFirstR = Cells(2, SeqOpen + 38)
'pasteSeqFirst = pasteSeqFirstR.Address(RowAbsolute:=False, ColumnAbsolute:=False)
'
'Set pasteSeqLastR = Cells(SvrColumn + 1, SeqOpen + 38)
'pasteSeqLast = pasteSeqLastR.Address(RowAbsolute:=False, ColumnAbsolute:=False)
'
'Set seqLastR = Cells(57, 2 + SvrColumn)
'seqLast = seqLastR.Address(RowAbsolute:=False, ColumnAbsolute:=False)
'
'Set delSeqR = Cells(57, activeSeqColumn + 3)
'delSeq = delSeqR.Address(RowAbsolute:=False, ColumnAbsolute:=False)
'Worksheets("Drum Machine").Range("C57", delSeq).Value = ""
'
'
'Worksheets("Drum Machine").Range("C57:" & seqLast).Value = Application.WorksheetFunction.Transpose(Worksheets("PatternSaver").Range(pasteSeqFirst, pasteSeqLast).Value)
'Worksheets("Drum Machine").Range("C57:" & seqLast).Replace "#N/A", "", xlWhole
'Worksheets("Drum Machine").Range("C57:" & seqLast).Replace 0, "", xlWhole
'
'
'Worksheets("Drum Machine").Range("G60").Interior.ColorIndex = 40
'SeqOpen = 5
'
'End Sub
'
'Sub countSeqColumn()
'
'    With Worksheets("Drum Machine")
'   ' lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
'   activeSeqColumn = .Cells(57, Columns.Count).End(xlToLeft).Column - 2 ' -1 for some random reason, -2 for so it dont count "pattern seq"
'  ' MsgBox lastRow
''Debug.Print "Seq is " & activeSeqColumn
'    End With
'End Sub
'
'Sub countSvrColumn()
'
'
'    With Worksheets("PatternSaver")
'   SvrColumn = .Cells(.Rows.Count, SeqOpen + 38).End(xlUp).Row - 1
'     'SvrColumn = .Cells(57, Columns.Count).End(xlToLeft).Column - 2 ' -1 for some random reason, -2 for so it dont count "pattern seq"
'  ' MsgBox lastRow
'' Debug.Print "Svr is " & SvrColumn
'
' If SvrColumn = 0 Then
' SvrColumn = 1
' End If
'
'    End With
'End Sub






