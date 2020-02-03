Attribute VB_Name = "Arrangement"
Option Explicit

#If Win64 Then
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Function GetQueueStatus Lib "user32" (ByVal qsFlags As Long) As Long
#Else
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetQueueStatus Lib "user32" (ByVal qsFlags As Long) As Long
#End If




Dim stepsDone As Integer
Dim loopsLeft As Integer
Dim ARoffset As Integer
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
Dim sWatch As New Stopwatch
Dim imPlayinAR As Boolean

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
Dim startFinder As Integer

Dim endArrangement As Integer
Dim patternStep As Integer
Dim ColorRepeater
Global lastOffsetArray() As Variant
Dim patternContinues() As Variant



Sub arrangementStart()

PRstopOkay
stopOkay

'Worksheets("Piano Roll").Range("H5").offset(0, PRoffset).Interior.ColorIndex = 34

'Debug.Print imPlayinAR
If imPlayinAR = False Then

    savePianoPatternOnPlay
    Seq.saveArray
    
    'startDevice 0, 0, 0, 0

    Worksheets("Arrangement").Range("h29:w29").Interior.Color = RGB(233, 242, 251)    'Index = 34
    startDevice 0, 0, currentPitch, 0
    
    pianorollProgram = Worksheets("Arrangement").Range("G31").Value - 1
    
    sWatch.Restart
    nowsWatch = sWatch.Elapsed
    loopsLeft = Range("C25").Value
    slipOver = 0
    
    'not using progMode anymore, 8 channels + 8 programs all time now
    'progMode = Range("C26").Value
    
    startFinder = 0
    Do
    startFinder = startFinder + 1
    Loop Until Left(Range("H29").offset(0, startFinder - 1).Value, 1) = "s" Or startFinder = Cells(29, Columns.Count).End(xlToLeft).Column
    
    If Left(Range("H29").offset(0, startFinder - 1).Value, 1) = "s" Then
    arrangementLoop 'put last
    Else
    'MsgBox "Couldn't find start marker"
    'Debug.Print 3555
    startFinder = 1
    arrangementLoop
    End If

Else
'Debug.Print "stoP"
stopArrangement
End If


End Sub




Sub arrangementLoop()


imPlayinAR = True


counter = 1
ARoffset = 0
seqOffset = 0
PRoffset = 0
stepsDone = 0


Dim numberTracks As Integer
numberTracks = howManyARTracks()

'x,0 = last PRoffset  ; x,1 = last Pattern ; x,2 = prSemitone;
ReDim lastOffsetArray(numberTracks, 3) As Variant
'ReDim patternContinues(numberTracks, 1) As Variant

endArrangement = 0
Do
endArrangement = endArrangement + 1

Loop Until LCase(Left(Range("H29").offset(0, (startFinder - 1) + endArrangement - 1).Value, 1)) = "e" Or endArrangement = Cells(29, Columns.Count).End(xlToLeft).Column

Dim i As Integer

If endArrangement = Cells(29, Columns.Count).End(xlToLeft).Column Then 'this means no E was found

    endArrangement = 0
    For i = 0 To howManyARTracks() - 1
    
        If Cells(31 + (i * 3), Columns.Count).End(xlToLeft).Column - 6 - (startFinder - 1) > endArrangement Then
        'Debug.Print Cells(31 + (i * 3), Columns.Count).End(xlToLeft).Column
        
        'Debug.Print "startFinder"; startFinder
        
        endArrangement = Cells(31 + (i * 3), Columns.Count).End(xlToLeft).Column - 6 - (startFinder - 1)
        End If
    Next i
End If

Dim loopPoint As Integer
loopPoint = 0

For i = 0 To Cells(29, Columns.Count).End(xlToLeft).Column

If LCase(Left(Range("H29").offset(0, (startFinder - 1) + i).Value, 1)) = "l" Then
loopPoint = i
End If

Next i

'Debug.Print "loopPoint"; loopPoint
'Debug.Print "enddd"; endArrangement
'Debug.Print "startFinder"; startFinder
stepsDone = 0


' i put this here because the first beat would be a bit delayed on +120bpms so this kinda fixed that
While (sWatch.Elapsed - nowsWatch) < 50 And imPlayinAR = True

    
Wend
nowsWatch = sWatch.Elapsed

'Debug.Print "sleeeep "; ((60 / Worksheets("Arrangement").Range("G22").Value) * 1000) / 4



    Do
    
    
    
    
    'stopItMIDI
    
    sleepTime = ((60 / Worksheets("Arrangement").Range("G22").Value) * 1000) / 4
    'sleepTime = (60 / 120 * 1000) * 4 'temp 120 bpm
    swingL = sleepTime + ((sleepTime / 5) * Worksheets("Arrangement").Range("G25"))
    swingS = sleepTime - ((sleepTime / 5) * Worksheets("Arrangement").Range("G25"))


    
    arrangementSequence
    
    
    
    
    
    
'    While (sWatch.Elapsed - nowsWatch) - slipOver < sleepTime
'    'Waits here until next step
'    Wend
'
'    slipOver = 0
'
'    nowsWatch = sWatch.Elapsed
'
    
    
    
    clearArrangementCosmetics
    
    
    
    ARoffset = ARoffset + 1
    'counter = counter + 1
    stepsDone = stepsDone + 1
    
    
    
    
    'loopMode - i  think i accidently removed my loopPoint finder lol
    
        If stepsDone = loopPoint Or Worksheets("Arrangement").Range("h29").offset(0, startFinder - 1 + ARoffset).Value = "l" And imPlayinAR = True Then
        stepsDone = 0
        ARoffset = 0
        'Debug.Print "ARoffset"; ARoffset
        justStopNote 0, 120, 0, 0
        End If
    
    
    Loop Until stepsDone >= endArrangement - 1

stepsDone = 0


'Looper

ARoffset = 0
seqOffset = 0
PRoffset = 0

stopItMIDI
stopItMIDIAgain
stopArrangement

End Sub


Sub arrangementSequence()

'Debug.Print "new bar"

Dim numberTracks As Integer
numberTracks = howManyARTracks()

Dim soloState As Boolean
soloState = isSoloOn()

 
Worksheets("Arrangement").Range("h29").offset(0, startFinder - 1 + ARoffset).Interior.ColorIndex = 41 ''this is the current step display bar

Dim i As Integer
Dim pianorollProgram As Integer
Dim pianrollPattern As Integer

Dim current16th As Integer
current16th = 1




 'Drum Machine stuff
    
Dim DMpartSelect As Integer
If Len(Worksheets("Arrangement").Range("h28").offset(3, startFinder - 1 + ARoffset).Value2) > 2 And Left(Worksheets("Arrangement").Range("h28").offset(3, startFinder - 1 + ARoffset).Value2, 1) <> " " And IsNumeric(Worksheets("Arrangement").Range("h28").offset(3, startFinder - 1 + ARoffset).Value2) = True Then
    If Mid(Worksheets("Arrangement").Range("h28").offset(3, startFinder - 1 + ARoffset).Value2, Len(Worksheets("Arrangement").Range("h28").offset(3, startFinder - 1 + ARoffset).Value2) - 1, 1) = "." And Right(Worksheets("Arrangement").Range("h28").offset(3, startFinder - 1 + ARoffset).Value2, 1) < 3 Then
    
    DMpartSelect = Right(Worksheets("Arrangement").Range("h28").offset(3, startFinder - 1 + ARoffset).Value2, 1) * 16 - 16
    End If
End If




Dim DMpattern As Integer
    
    If Worksheets("Arrangement").Range("h28").offset(3, startFinder - 1 + ARoffset).Value2 = "." Then
    DMpartSelect = 16
        If IsNumeric(Worksheets("Arrangement").Range("h28").offset(3, startFinder - 1 + ARoffset - 1).Value2) = True Then
        DMpattern = Application.WorksheetFunction.RoundDown(Worksheets("Arrangement").Range("h28").offset(3, startFinder - 1 + ARoffset - 1).Value2, 0)
        End If
    ElseIf IsNumeric(Worksheets("Arrangement").Range("h28").offset(3, startFinder - 1 + ARoffset).Value2) = True Then
    DMpattern = Application.WorksheetFunction.RoundDown(Worksheets("Arrangement").Range("h28").offset(3, startFinder - 1 + ARoffset).Value2, 0)
    End If
    
'turn off if Update is off
If Worksheets("Drum Machine").Range("C25").Value = "On" Then

    If IsNumeric(Worksheets("Arrangement").Range("h28").offset(3, startFinder - 1 + ARoffset - 1).Value2) = True Then
        Worksheets("Drum Machine").Range("D52").Value2 = DMpattern
    End If


    If IsNumeric(Worksheets("Arrangement").Range("h28").offset(3, startFinder - 1 + ARoffset).Value2) = True Then
    Worksheets("Drum Machine").Range("D52").Value2 = Application.WorksheetFunction.RoundDown(Worksheets("Arrangement").Range("h28").offset(3, startFinder - 1 + ARoffset).Value2, 0)
    End If
    
    If Left(Worksheets("Arrangement").Range("h28").offset(3, startFinder - 1 + ARoffset).Value2, 1) <> " " And Worksheets("Arrangement").Range("h28").offset(3, startFinder - 1 + ARoffset).Value2 > 0 Then
    
    
    openPattern
    End If
End If


Dim DMvelocity As Integer
If Worksheets("Arrangement").Range("h28").offset(4, startFinder - 1 + ARoffset).Value2 <> "" And Left(Worksheets("Arrangement").Range("h28").offset(4, startFinder - 1 + ARoffset).Value2, 1) <> " " And IsNumeric(Worksheets("Arrangement").Range("h28").offset(4, startFinder - 1 + ARoffset).Value2) = True Then
    DMvelocity = Worksheets("Arrangement").Range("h28").offset(4, startFinder - 1 + ARoffset).Value2
ElseIf Worksheets("Arrangement").Range("G32").Value2 <> "" And Left(Worksheets("Arrangement").Range("G32").Value2, 1) <> " " And IsNumeric(Worksheets("Arrangement").Range("G32").Value2) = True Then
    DMvelocity = Worksheets("Arrangement").Range("G32").Value2
End If

Dim DMsemitone As Integer
If Worksheets("Arrangement").Range("h28").offset(5, startFinder - 1 + ARoffset).Value2 <> "" And Left(Worksheets("Arrangement").Range("h28").offset(5, startFinder - 1 + ARoffset).Value2, 1) <> " " And IsNumeric(Worksheets("Arrangement").Range("h28").offset(5, startFinder - 1 + ARoffset).Value2) = True Then
    DMsemitone = Worksheets("Arrangement").Range("h28").offset(5, startFinder - 1 + ARoffset).Value2
ElseIf Worksheets("Arrangement").Range("G33").Value2 <> "" And Left(Worksheets("Arrangement").Range("G33").Value2, 1) <> " " And IsNumeric(Worksheets("Arrangement").Range("G33").Value2) = True Then
    DMsemitone = Worksheets("Arrangement").Range("G33").Value2
End If


    
    For patternStep = 1 To 16
        
    DoEvents
    
    If current16th > 4 Then
    current16th = 1
    End If
    
    
    
    If current16th = 1 Or current16th = 3 Then
    'Debug.Print "swingL"; swingL
    sleepTime = swingL
    Else
    sleepTime = swingS
    'Debug.Print "swingS"; swingS
    End If
    current16th = current16th + 1
    
                ''' in case lag goes over sleepTime, it will deduct from next wait to balance out
    If sWatch.Elapsed - nowsWatch > sleepTime Then
    slipOver = (sWatch.Elapsed - nowsWatch) - sleepTime
    'Debug.Print "Slip Over Activated"
    End If
    
    

    'moved wait before notes otherwise the notes might be occasionally delayed (but still in sync)
    'Debug.Print sWatch.Elapsed - nowsWatch
    While (sWatch.Elapsed - nowsWatch) - slipOver < sleepTime And imPlayinAR = True

    
    Wend
    nowsWatch = sWatch.Elapsed
    Seq.clearCosmeticsFromAR
    
    
    If soloState = True And LCase(Worksheets("Arrangement").Range("D31").Value2) = "s" Then
       
        If Left(Worksheets("Arrangement").Range("h28").offset(3, startFinder - 1 + ARoffset).Value2, 1) <> " " And Worksheets("Arrangement").Range("h28").offset(3, startFinder - 1 + ARoffset).Value2 > 0 Then
        
            If Worksheets("Drum Machine").Range("C25").Value = "On" Then
            Call Seq.startSeq2(DMpartSelect, DMvelocity, DMsemitone)
            Else
            Call DMseqFromSaver(DMpartSelect, DMvelocity, DMsemitone, DMpattern)
            End If
        
        End If
    
    ElseIf soloState = False And LCase(Worksheets("Arrangement").Range("D31").Value2) <> "m" Then
    
        If Left(Worksheets("Arrangement").Range("h28").offset(3, startFinder - 1 + ARoffset).Value2, 1) <> " " And Worksheets("Arrangement").Range("h28").offset(3, startFinder - 1 + ARoffset).Value2 > 0 Then
        
            If Worksheets("Drum Machine").Range("C25").Value = "On" Then
            Call Seq.startSeq2(DMpartSelect, DMvelocity, DMsemitone)
            Else
            Call DMseqFromSaver(DMpartSelect, DMvelocity, DMsemitone, DMpattern)
            End If
        
        End If
    
    End If
    
    
    '' unecessary for send2PR to run 16 times, only needs to run once (move outside of loop)
        For i = 0 To numberTracks - 2 'amount of piano roll tracks
        
        
            If soloState = True And LCase(Worksheets("Arrangement").Range("D34").offset(3 * i, 0).Value2) = "s" Then
            
            Call send2PR(i, lastOffsetArray)
            
            ElseIf soloState = False And LCase(Worksheets("Arrangement").Range("D34").offset(3 * i, 0).Value2) <> "m" Then
            
            Call send2PR(i, lastOffsetArray)
            
            End If
            
        
        Next i
    
    
    
    
    slipOver = 0
    
    seqOffset = seqOffset + 1
    'pianroll.PRclearCosmetics
    
    
    PRoffset = PRoffset + 1
    
    

    'DoEvents ' allows u to still click on sheet when running
    
    Next patternStep
    
    
seqOffset = 0
PRoffset = 0

End Sub

Sub DMseqFromSaver(DMpartSelect, DMvelocity, DMsemitone, DMpattern)

'Debug.Print "seqqq"

If DMpattern = 0 Or DMpattern = "" Then
DMpattern = 5709
End If

If DMvelocity = 0 Or DMvelocity = "" Then
DMvelocity = 100
End If

Dim i As Integer

For i = 0 To 7


If Left(Worksheets("PatternSaver").Range("D1").offset((DMpattern * 24 - 24) + (i * 3), seqOffset + DMpartSelect).Value, 1) = "x" Then

If Worksheets("PatternSaver").Range("D1").offset((DMpattern * 24 - 24) + (i * 3) + 1, seqOffset + DMpartSelect).Value <> "" And IsNumeric(Worksheets("PatternSaver").Range("D1").offset((DMpattern * 24 - 24) + (i * 3) + 1, seqOffset + DMpartSelect).Value) = True Then
currentVel = Worksheets("PatternSaver").Range("D1").offset((DMpattern * 24 - 24) + (i * 3) + 1, seqOffset + DMpartSelect).Value
Else
currentVel = Worksheets("PatternSaver").Range("D1").offset((DMpattern * 24 - 24) + (i * 3) + 1, -2).Value
End If

If Worksheets("PatternSaver").Range("D1").offset((DMpattern * 24 - 24) + (i * 3) + 2, seqOffset + DMpartSelect).Value <> "" And IsNumeric(Worksheets("PatternSaver").Range("D1").offset((DMpattern * 24 - 24) + (i * 3) + 2, seqOffset + DMpartSelect).Value) = True Then
currentPitch = Worksheets("PatternSaver").Range("D1").offset((DMpattern * 24 - 24) + (i * 3) + 2, seqOffset + DMpartSelect).Value
Else
currentPitch = Worksheets("PatternSaver").Range("D1").offset((DMpattern * 24 - 24) + (i * 3) + 2, -2).Value
End If


    If Worksheets("Drum Machine").Range("C24").Value = 1 Then
    midiNote 0, (currentVel / 100) * DMvelocity, currentPitch + DMsemitone, 9
    Else
    midiNote i, (currentVel / 100) * DMvelocity, currentPitch + DMsemitone, i
    End If

End If

Next i

End Sub


Sub send2PR(i, ByRef lastOffsetArray As Variant)



If Worksheets("Arrangement").Range("H34").offset(3 * i, startFinder - 1 + ARoffset).Value2 <> "" And IsNumeric(Worksheets("Arrangement").Range("H34").offset(3 * i, startFinder - 1 + ARoffset).Value2) = True Or Worksheets("Arrangement").Range("H34").offset(3 * i, startFinder - 1 + ARoffset).Value2 = "." Then
'pianroll.PRseq
    



    pianorollProgram = Worksheets("Arrangement").Range("E35").offset(3 * i, 0).Value2 - 1
    
    'decimal part goes here
    Dim prPartSelect As Integer
    If Len(Worksheets("Arrangement").Range("H34").offset(3 * i, startFinder - 1 + ARoffset).Value) > 2 Then
     If Mid(Worksheets("Arrangement").Range("H34").offset(3 * i, startFinder - 1 + ARoffset).Value, Len(Worksheets("Arrangement").Range("H34").offset(3 * i, startFinder - 1 + ARoffset).Value) - 1, 1) = "." And Right(Worksheets("Arrangement").Range("H34").offset(3 * i, startFinder - 1 + ARoffset).Value, 1) < 5 Then
     prPartSelect = Right(Worksheets("Arrangement").Range("H34").offset(3 * i, startFinder - 1 + ARoffset).Value, 1) * 16 - 16
     End If
    End If
        
        Dim x As Integer
        Dim pianrollPattern As Integer
        For x = 0 To 3

            If Worksheets("Arrangement").Range("H34").offset(3 * i, startFinder - 1 + ARoffset - x).Value <> "." And IsNumeric(Worksheets("Arrangement").Range("H34").offset(3 * i, startFinder - 1 + ARoffset - x).Value) = True Then
            pianrollPattern = Application.WorksheetFunction.RoundDown(Worksheets("Arrangement").Range("H34").offset(3 * i, startFinder - 1 + ARoffset - x).Value, 0)

            Select Case x
                Case 1
                prPartSelect = 16
                Case 2
                prPartSelect = 32
                Case 3
                prPartSelect = 48
            End Select

            x = 3
            End If
        Next x
    
    
    
    Dim prChannel As Integer
    Dim prSemitone As Integer
    Dim arVelocity As Integer
      
    
    If Worksheets("Arrangement").Range("E36").offset(3 * i, 0).Value <> "" Then
    prChannel = Worksheets("Arrangement").Range("E36").offset(3 * i, 0).Value - 1
    Else
    prChannel = 0
    End If
    
    
    'velocity
    If IsNumeric(Worksheets("Arrangement").Range("H34").offset(3 * i + 1, startFinder - 1 + ARoffset).Value) = True And Left(Worksheets("Arrangement").Range("H34").offset(3 * i + 1, startFinder - 1 + ARoffset).Value, 1) <> " " And Worksheets("Arrangement").Range("H34").offset(3 * i + 1, startFinder - 1 + ARoffset).Value <> "" Then
    arVelocity = Worksheets("Arrangement").Range("H34").offset(3 * i + 1, startFinder - 1 + ARoffset).Value
    ElseIf Worksheets("Arrangement").Range("G35").offset(3 * i, 0).Value <> "" And Left(Worksheets("Arrangement").Range("G35").offset(3 * i, 0).Value, 1) <> " " Then
    arVelocity = Worksheets("Arrangement").Range("G35").offset(3 * i, 0).Value
    Else
    arVelocity = 100
    End If
    
    'semitones
    If IsNumeric(Worksheets("Arrangement").Range("H34").offset(3 * i + 2, startFinder - 1 + ARoffset).Value) = True And Left(Worksheets("Arrangement").Range("H34").offset(3 * i + 2, startFinder - 1 + ARoffset).Value, 1) <> " " And Worksheets("Arrangement").Range("H34").offset(3 * i + 2, startFinder - 1 + ARoffset).Value <> "" Then

    prSemitone = Worksheets("Arrangement").Range("H34").offset(3 * i + 2, startFinder - 1 + ARoffset).Value
    ElseIf Worksheets("Arrangement").Range("G36").offset(3 * i, 0).Value <> "" Then
        prSemitone = Worksheets("Arrangement").Range("G36").offset(3 * i, 0).Value
    Else
    prSemitone = 0
        
    End If
    
    Dim curTr As Integer
    curTr = i
    
    
    'Debug.Print "arVelocity"; arVelocity
        
     If imPlayinAR = True Then
     Call pianroll.PRseqFromSaver(pianrollPattern, pianorollProgram, prChannel, arVelocity, prSemitone, lastOffsetArray, curTr, prPartSelect, sWatch)
     End If
     
     'Call pianroll.PRseqFromSaver(pianrollPattern, pianorollProgram, prChannel, prVelocity, prSemitone, lastOffset, 0)
     'Worksheets("Piano Roll").Range("BB3").Value = Worksheets("Arrangement").Range("h28").offset(4, ARoffset).Value
 
 
 
'if AR cell is empty but previous wasnt (so note off for last cell)
ElseIf Worksheets("Arrangement").Range("H34").offset(3 * i, startFinder - 1 + ARoffset - 1).Value2 <> "" And _
IsNumeric(Worksheets("Arrangement").Range("H34").offset(3 * i, startFinder - 1 + ARoffset - 1).Value2) = True Or _
Worksheets("Arrangement").Range("H34").offset(3 * i, startFinder - 1 + ARoffset - 1).Value2 = "." Then


    If Worksheets("Arrangement").Range("G24").Value = "Off" Then
    

    
        pianorollProgram = Worksheets("Arrangement").Range("E35").offset(3 * i, 0).Value2 - 1
    
        If Worksheets("Arrangement").Range("E36").offset(3 * i, 0).Value <> "" Then
            prChannel = Worksheets("Arrangement").Range("E36").offset(3 * i, 0).Value - 1
        Else
            prChannel = 0
        End If
    
    
        'If IsNumeric(Worksheets("Arrangement").Range("H34").offset(3 * i + 2, startFinder - 1 + ARoffset).Value) = True And Left(Worksheets("Arrangement").Range("H34").offset(3 * i + 2, startFinder - 1 + ARoffset).Value, 1) <> " " Then
        '    prSemitone = Worksheets("Arrangement").Range("H34").offset(3 * i + 2, startFinder - 1 + ARoffset).Value
        '    ElseIf Worksheets("Arrangement").Range("G36").offset(3 * i, 0).Value <> "" Then
        '        prSemitone = Worksheets("Arrangement").Range("G36").offset(3 * i, 0).Value
        '        Else
        '        prSemitone = 0
        '
        'End If
                
        curTr = i
        
        If PRoffset = 0 And imPlayinAR = True Then
        Call pianroll.PRseqFromSaver(5709, pianorollProgram, prChannel, arVelocity, prSemitone, lastOffsetArray, curTr, 0, sWatch)
        End If
    
    
    End If
End If


'patternContinues(i + 1, 0) = patternContinues(i + 1, 0) + 1

End Sub




Sub clearArrangementCosmetics()


'Worksheets("Arrangement").Range("h28").offset(0, startFinder - 1 + ARoffset).Interior.ColorIndex = 34  ' this creates occasional quirks of 1ms - not the end of the world but still
Worksheets("Arrangement").Range("h29").offset(0, startFinder - 1 + ARoffset).Interior.Color = RGB(233, 242, 251)


End Sub


Sub stopArrangement()

loopsLeft = 0
'stepsDone = 32
'counter = 0
Worksheets("Piano Roll").Range("H5").offset(0, PRoffset).Interior.ColorIndex = 34

If Worksheets("Drum Machine").Range("C25") = "On" Then
Worksheets("Drum Machine").Range("H28").offset(0, seqOffset).Interior.ColorIndex = 34
End If

If startFinder - 1 + ARoffset <> -1 Then
Worksheets("Arrangement").Range("h29").offset(0, startFinder - 1 + ARoffset).Interior.Color = RGB(233, 242, 251) 'Index = 34
End If

seqOffset = 0
PRoffset = 0
ARoffset = 0
patternStep = 16

stepsDone = endArrangement - 1


stopItMIDIAgain
imPlayinAR = False

End Sub




Sub changePatternsFromArrangement(ByRef playThesePatterns As Variant)

'make loop
If Worksheets("Arrangement").Range("h28").offset(3, ARoffset).Value > 0 Then

End If

'If Worksheets("Arrangement").Range("h28").offset(3, ARoffset).Value > 0 Then
'Worksheets("Piano Roll").Range("BB3").Value = Worksheets("Arrangement").Range("h28").offset(3, ARoffset).Value
'pianorollProgram = Worksheets("Arrangement").Range("h28").offset(3, -1).Value - 1
'End If


'openPianoPattern




End Sub

Sub addARTrack()

Dim numberTracks As Integer
numberTracks = howManyARTracks()

Dim trackPlacement As Range
Set trackPlacement = Worksheets("Arrangement").Range("D31")

If numberTracks = 0 Then
MsgBox "You weren't supposed to remove the Drums track..."
MsgBox "I'm not angry.. just disappointed"
Else


'group
Rows(32 + (numberTracks * 3) & ":" & 33 + (numberTracks * 3)).Group

'keeps closed if above track is closed too
If Rows(31 + (numberTracks * 3) - 2).Height = 0 Then

Rows(31 + (numberTracks * 3) + 1).ShowDetail = False
Else

End If


'format conditions
With Range(Cells(31 + (numberTracks * 3), 8), Cells(31 + (numberTracks * 3), Columns.Count)).FormatConditions.Add(xlNoBlanksCondition)
    With .Interior
    .ColorIndex = numberTracks - (numberTracks - 30) + numberTracks Mod 10 + 5
    End With
End With

'With Range(Cells(32 + (numberTracks * 3), 8), Cells(33 + (numberTracks * 3), Columns.Count)).FormatConditions.Add(xlNoBlanksCondition)
'    With .Interior
'    .ColorIndex = ColorRepeater
'    .TintAndShade = 0.5
'    End With
'End With


Range(Cells(31 + (numberTracks * 3), 8), Cells(33 + (numberTracks * 3), Columns.Count)).HorizontalAlignment = xlCenter
Range(Cells(32 + (numberTracks * 3), 8), Cells(33 + (numberTracks * 3), Columns.Count)).Interior.Color = RGB(245, 245, 245) 'grey inbetween lanes
Range(Cells(32 + (numberTracks * 3), 8), Cells(33 + (numberTracks * 3), Columns.Count)).BorderAround Color:=RGB(231, 230, 230), Weight:=xlThin
Range(Cells(32 + (numberTracks * 3), 8), Cells(33 + (numberTracks * 3), Columns.Count)).Borders(xlInsideHorizontal).Color = RGB(231, 230, 230) 'put grid lines back
Range(Cells(32 + (numberTracks * 3), 8), Cells(33 + (numberTracks * 3), Columns.Count)).Borders(xlInsideVertical).Color = RGB(231, 230, 230) 'put grid lines back

trackPlacement.offset(numberTracks * 3, 0).Interior.Color = RGB(226, 240, 217)
trackPlacement.offset(numberTracks * 3, 0).HorizontalAlignment = xlCenter

Range(Cells(31 + (numberTracks * 3), 5), Cells(31 + (numberTracks * 3), 7)).Merge


Range("E31:G33").offset(numberTracks * 3, 0).HorizontalAlignment = xlCenter

Range("E31").offset(numberTracks * 3, 0).Interior.Color = RGB(255, 242, 204)
Range("E31").offset(numberTracks * 3, 0).Value = "Track " & numberTracks + 1

Range("D32:D33").offset(numberTracks * 3, 0).Interior.Color = RGB(189, 215, 238)
Range("D32").offset(numberTracks * 3, 0).Value = "P:"
Range("D33").offset(numberTracks * 3, 0).Value = "C:"
Range("D33").offset(numberTracks * 3, 1).Value = numberTracks Mod 10 + 1

Range("E32:E33").offset(numberTracks * 3, 0).Interior.Color = RGB(221, 235, 247)

Range("F32:F33").offset(numberTracks * 3, 0).Interior.Color = RGB(248, 203, 173)
Range("F32").offset(numberTracks * 3, 0).Value = "V:"
Range("F33").offset(numberTracks * 3, 0).Value = "S:"
Range("G32:G33").offset(numberTracks * 3, 0).Interior.Color = RGB(252, 228, 214)

trackPlacement.offset(numberTracks * 3, 0).BorderAround ColorIndex:=1, Weight:=xlThin
Range("D31:G33").offset(numberTracks * 3, 0).BorderAround ColorIndex:=1, Weight:=xlThin
Range("E31").offset(numberTracks * 3, 0).Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous 'had to put separate for whatever reason
Range("F31").offset(numberTracks * 3, 0).Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous 'something to do with merged cell
Range("G31").offset(numberTracks * 3, 0).Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous




End If


End Sub

Sub removeARTrack()

Dim numberTracks As Integer
numberTracks = howManyARTracks()

If numberTracks = 1 Then
MsgBox "Please don't remove the drum track, thank you : )"

Else

'format conditions
Range(Cells(31 + (numberTracks * 3) - 3, 8), Cells(33 + (numberTracks * 3) - 3, Columns.Count)).FormatConditions.Delete
'    With .Interior
'    .Color = -4142
'    End With



Range("D31:G33").offset((numberTracks * 3) - 3, 0).Interior.Color = -4142 'resets grid to Transparent (not white!!! which masks gridlines)
Range("D31:G33").offset((numberTracks * 3) - 3, 0).Value = ""
Range("D31:G33").offset((numberTracks * 3) - 3, 0).UnMerge
Range("D31:G33").offset((numberTracks * 3) - 3, 0).Borders.LineStyle = xlLineStyleNone

Range(Cells(32 + (numberTracks * 3) - 3, 8), Cells(33 + (numberTracks * 3) - 3, Columns.Count)).Interior.Color = -4142
Range(Cells(32 + (numberTracks * 3) - 3, 8), Cells(33 + (numberTracks * 3) - 3, Columns.Count)).Borders.LineStyle = xlLineStyleNone

If Rows(31 + (numberTracks * 3)).ShowDetail <> True Then
Rows(31 + (numberTracks * 3)).ShowDetail = True
End If

Rows(32 + (numberTracks * 3) - 3 & ":" & 33 + (numberTracks * 3) - 3).Ungroup

End If


End Sub


'Private Sub oopsieAAAA()
'
''Worksheets("PatternSaver").Range("A4").Value = 1
'array1
'loopsLeft = 0
'stepsDone = 16
''sWatch.Pause
'imPlayinAR = False
'Seq.stopItMIDI  'this calls the macro in playMidi
'
'
'
'End Sub

Sub clearARMS()

Dim i As Variant

For i = 0 To howManyARTracks()

Worksheets("Arrangement").Range("D31").offset(i * 3, 0).Value = ""

Next i


End Sub

Sub tstCOndd()

With Range("H31", Cells(31 + (0 * 3), Columns.Count)).FormatConditions.Add(xlNoBlanksCondition)
    With .Interior
    .ColorIndex = 35
    End With
End With


End Sub



Function howManyARTracks()
'R: 226 G: 240 B: 217

'it cant detect the text when groups are folded!!
'Debug.Print Worksheets("Arrangement").Range("D" & Rows.Count).End(xlUp).Row

Dim i As Variant

For i = 0 To Rows.Count  'Worksheets("Arrangement").Range("D" & Rows.Count).End(xlUp).Row


If Worksheets("Arrangement").Range("D31").offset(i * 3, 0).Interior.ColorIndex <> 35 Then
'Debug.Print "there are "; i
howManyARTracks = i
i = Rows.Count 'Worksheets("Arrangement").Range("D" & Rows.Count).End(xlUp).Row
End If

Next i

End Function


Function isSoloOn()

Dim i As Variant

For i = 0 To Range("D" & Rows.Count).End(xlUp).Row

If Worksheets("Arrangement").Range("D31").offset(i * 3, 0).Value = "s" Then
'Debug.Print "solo on"
isSoloOn = True
i = Range("D" & Rows.Count).End(xlUp).Row
End If


Next i

End Function


Sub testModulo()
'neat it works
'Debug.Print

Dim i As Integer
Dim x As Integer

For i = 1 To 40
x = 60 + i
x = x - (x - 50) + x Mod 10
'Debug.Print x

Next i


End Sub






Sub tempcolourDrums()

Dim numberTracks As Integer

numberTracks = 0
Range(Cells(32 + (numberTracks * 3), 8), Cells(33 + (numberTracks * 3), Columns.Count)).Interior.Color = RGB(245, 245, 245) 'grey inbetween lanes
Range(Cells(32 + (numberTracks * 3), 8), Cells(33 + (numberTracks * 3), Columns.Count)).BorderAround Color:=RGB(231, 230, 230), Weight:=xlThin
Range(Cells(32 + (numberTracks * 3), 8), Cells(33 + (numberTracks * 3), Columns.Count)).Borders(xlInsideHorizontal).Color = RGB(231, 230, 230) 'put grid lines back
Range(Cells(32 + (numberTracks * 3), 8), Cells(33 + (numberTracks * 3), Columns.Count)).Borders(xlInsideVertical).Color = RGB(231, 230, 230) 'put grid lines back

End Sub


Sub clearArrangementCounter()

Rows(29).Value = ""

End Sub








