Attribute VB_Name = "pianroll"
Option Explicit

Dim sWatch As New Stopwatch
Global PRoffset As Integer

Dim sleepTime As Variant
Dim swingL As Variant
Dim swingS As Variant
Global nowsWatch As Variant
Dim slipOver As Variant
Dim imPlayin As Boolean
Dim lastOffset As Integer
Global pianorollProgram As Integer
Dim pianorollChannel As Integer
Dim PRstartPos As Integer


Sub createAbletonPiano()

Worksheets("Piano Roll").Range("CB1:CF3").Interior.ColorIndex = 40
Worksheets("Piano Roll").Range("CB4:CF6").Interior.ColorIndex = 0
Dim x As Integer
x = 8
createPiano (x)
End Sub

Sub createFLPiano()

Worksheets("Piano Roll").Range("CB1:CF3").Interior.ColorIndex = 0
Worksheets("Piano Roll").Range("CB4:CF6").Interior.ColorIndex = 40
Dim x As Integer
x = 10
createPiano (x)
End Sub

Sub createPiano(x)

Dim i As Integer
For i = 4 To 127


Range("D12").offset(i, 0).Value = "G" & x
i = i + 1
Range("D12").offset(i, 0).Value = "F#" & x
i = i + 1
Range("D12").offset(i, 0).Value = "F" & x
i = i + 1
Range("D12").offset(i, 0).Value = "E" & x
i = i + 1
Range("D12").offset(i, 0).Value = "D#" & x
i = i + 1
Range("D12").offset(i, 0).Value = "D" & x
i = i + 1
Range("D12").offset(i, 0).Value = "C#" & x
i = i + 1
Range("D12").offset(i, 0).Value = "C" & x
i = i + 1

If i < 127 Then
Range("D12").offset(i, 0).Value = "B" & x
i = i + 1
Range("D12").offset(i, 0).Value = "A#" & x
i = i + 1
Range("D12").offset(i, 0).Value = "A" & x
i = i + 1
Range("D12").offset(i, 0).Value = "G#" & x
End If

x = x - 1

Next i

End Sub



Sub PRpatternPlay()

stopArrangement
stopOkay

'Debug.Print "playin "; imPlayin


If imPlayin = False Then


startDevice 0, 0, 0, 0

'pianorollProgram = Worksheets("Piano Roll").Range("AT2").Value - 1
'    If pianorollProgram < 0 Then
'    pianorollProgram = 0
'    End If
'
'pianorollChannel = Worksheets("Piano Roll").Range("AT3").Value - 1
'    If pianorollChannel < 0 Then
'    pianorollChannel = 0
'    End If
'
Dim seqOn As Integer
seqOn = 0
PRstartLoop
'Debug.Print "donePRstartLoop"

Else
PRstopOkay
'Debug.Print "done PRstopOkay"

End If

End Sub

Sub PRstartLoop()



Worksheets("Piano Roll").Range("H5:BS5").Interior.ColorIndex = 34

sWatch.Restart
nowsWatch = sWatch.Elapsed
'loopsLeft = Worksheets("Piano Roll").Range("C4").Value
slipOver = 0

'not using progMode anymore, 8 channels + 8 programs all time now
'progMode = Range("C26").Value
nowsWatch = 0

Dim i As Integer

For i = 0 To 63
    If LCase(Worksheets("Piano Roll").Range("H5").offset(0, i).Value) = "s" Then
    PRstartPos = i
    i = 63
    Else
    PRstartPos = 0
    End If
Next i



PRLoopSeq 'put last


End Sub




Sub PRLoopSeq()


imPlayin = True

Dim counter As Integer
counter = 1
PRoffset = PRstartPos
stepsDone = 0

Dim current16th As Integer
current16th = 1

'lastOffset = -1 'risky??

Do

sleepTime = ((60 / Worksheets("Piano Roll").Range("E2").Value) * 1000) / 4
swingL = sleepTime + ((sleepTime / 5) * Worksheets("Piano Roll").Range("E5"))
swingS = sleepTime - ((sleepTime / 5) * Worksheets("Piano Roll").Range("E5"))

    If current16th > 4 Then
    current16th = 1
    End If
    
    'metronome
    If current16th = 1 And Worksheets("Piano Roll").Range("BB3").Value = "On" Then
    midiNote 0, 90, 70, 9
    End If

    
    pianorollProgram = Worksheets("Piano Roll").Range("AT2").Value - 1
    If pianorollProgram < 0 Then
    pianorollProgram = 0
    End If
    
    pianorollChannel = Worksheets("Piano Roll").Range("AT3").Value - 1
    If pianorollChannel < 0 Then
    pianorollChannel = 0
    End If
    
    
    
    Worksheets("Piano Roll").Range("H5").offset(0, PRoffset).Interior.ColorIndex = 41
    If imPlayin = True Then
    PRseq  'plays note(s)
    End If


        ''' in case lag goes over sleepTime, it will deduct from next wait to balance out
    If sWatch.Elapsed - nowsWatch > sleepTime And imPlayin = True Then
    slipOver = (sWatch.Elapsed - nowsWatch) - sleepTime
    'Debug.Print "slip mood"
    End If

'             'swing
'
    If current16th = 1 Or current16th = 3 Then
    sleepTime = swingL
    Else
    sleepTime = swingS
    End If

    
        While (sWatch.Elapsed - nowsWatch) - slipOver < sleepTime And imPlayin = True
        'Waits here until next step
        
            
        
        Wend
    
    DoEvents ' allows u to still click on sheet when running
    slipOver = 0
    
    nowsWatch = sWatch.Elapsed
    
    PRclearCosmetics
    
    PRoffset = PRoffset + 1
    'counter = counter + 1
    stepsDone = stepsDone + 1
    current16th = current16th + 1
    
    If Worksheets("Piano Roll").Range("H5").offset(0, PRoffset).Value = "e" Or _
    Worksheets("Piano Roll").Range("H5").offset(0, PRoffset).Value = "l" Or Worksheets("Piano Roll").Range("H5").offset(0, PRoffset).Address = "$BT$5" Then
        If imPlayin = True Then 'stops its breaking when u perfectly time it to stop on a repeat
        
'            If Worksheets("Piano Roll").Range("H5").offset(0, PRoffset).Value = "e" Then
'            lastOffset = -1
'            End If

        stepsDone = 0
        PRoffset = PRstartPos
        stepsDone = 0
        
        End If
    End If


Loop Until stepsDone >= 64

PRoffset = PRstartPos
stepsDone = 0


PRLooper



End Sub



Sub PRseq()

Dim prSemitone As Integer
Dim prVelocity As Integer
prVelocity = Worksheets("Piano Roll").Range("H6").offset(0, PRoffset).Value
'Debug.Print prVelocity

Dim x As Integer
Dim i As Integer

x = 0


For i = 0 To 127


    If Worksheets("Piano Roll").Range("E4").Value = "Off" Then  'if legato mode off
            
        
            If Worksheets("Piano Roll").Range("H12").offset(131 - i, PRoffset).Value = "" And Worksheets("Piano Roll").Range("H12").offset(131 - i, lastOffset).Value <> "" _
            Or Worksheets("Piano Roll").Range("H12").offset(131 - i, PRoffset).Value <> Worksheets("Piano Roll").Range("H12").offset(131 - i, lastOffset).Value And Right(Worksheets("Piano Roll").Range("H12").offset(131 - i, PRoffset).Value, 1) = "!" Then
            
            'Debug.Print "stop"; i; "lastOffset"; lastOffset
            
            'justStopNote2 pianorollProgram, 0, i, pianorollChannel
            stopARchordNotes 0, pianorollProgram, pianorollChannel, 0, i, Worksheets("Piano Roll").Range("H12").offset(131 - i, lastOffset).Value
            'stopARchordNotes(pattern, program, prChannel, prVelocity, ByRef prSemitone As Integer, noteStamp)
            End If
    
    End If



    If Worksheets("Piano Roll").Range("H12").offset(131 - i, PRoffset).Value <> "" And _
    Left(Worksheets("Piano Roll").Range("H12").offset(131 - i, PRoffset).Value, 1) <> " " Then 'if PR cell not empty
     
     If Worksheets("Piano Roll").Range("E4").Value = "Off" And Worksheets("Piano Roll").Range("H12").offset(131 - i, PRoffset - 1).Value = "" _
     Or Worksheets("Piano Roll").Range("E4").Value = "Off" And Right(Worksheets("Piano Roll").Range("H12").offset(131 - i, PRoffset).Value, 1) = "!" _
     Or Worksheets("Piano Roll").Range("E4").Value = "On" Then
     
        'Debug.Print "note on"; i
     
        If Worksheets("Piano Roll").Range("E4").Value = "On" Then 'if legato mode on
            If lastOffset <> PRoffset Then
            'Debug.Print sWatch.Elapsed
            justStopNote pianorollProgram, 120, x, pianorollChannel
            'Debug.Print sWatch.Elapsed
            End If
        End If
        
        
        
        
        'Debug.Print Worksheets("Piano Roll").Range("H12").offset(131 - i, PRoffset).Value
        Select Case Worksheets("Piano Roll").Range("H12").offset(131 - i, PRoffset).Value
        
        Case "x", "x!"
            'Debug.Print "befor "; PRoffset; "  "; sWatch.Elapsed - nowsWatch
        midiNote pianorollProgram, prVelocity, i, pianorollChannel
            'Debug.Print "after "; PRoffset; "  "; sWatch.Elapsed - nowsWatch
        

        
        Case "m", "m!"
        midiNote pianorollProgram, (120 * prVelocity) / 100, i + prSemitone, pianorollChannel
        midiNote pianorollProgram, (120 * prVelocity) / 100, i + prSemitone + 3, pianorollChannel
        midiNote pianorollProgram, (120 * prVelocity) / 100, i + prSemitone + 7, pianorollChannel
        
        Case "M", "M!"
        midiNote pianorollProgram, (120 * prVelocity) / 100, i + prSemitone, pianorollChannel
        midiNote pianorollProgram, (120 * prVelocity) / 100, i + prSemitone + 4, pianorollChannel
        midiNote pianorollProgram, (120 * prVelocity) / 100, i + prSemitone + 7, pianorollChannel
        
        Case "m7", "m7!"
        midiNote pianorollProgram, (120 * prVelocity) / 100, i + prSemitone, pianorollChannel
        midiNote pianorollProgram, (120 * prVelocity) / 100, i + prSemitone + 3, pianorollChannel
        midiNote pianorollProgram, (120 * prVelocity) / 100, i + prSemitone + 7, pianorollChannel
        midiNote pianorollProgram, (120 * prVelocity) / 100, i + prSemitone + 10, pianorollChannel
        
        Case "M7", "M7!"
        midiNote pianorollProgram, (120 * prVelocity) / 100, i + prSemitone, pianorollChannel
        midiNote pianorollProgram, (120 * prVelocity) / 100, i + prSemitone + 4, pianorollChannel
        midiNote pianorollProgram, (120 * prVelocity) / 100, i + prSemitone + 7, pianorollChannel
        midiNote pianorollProgram, (120 * prVelocity) / 100, i + prSemitone + 11, pianorollChannel
        
        Case "m9", "m9!"
        midiNote pianorollProgram, (120 * prVelocity) / 100, i + prSemitone, pianorollChannel
        midiNote pianorollProgram, (120 * prVelocity) / 100, i + prSemitone + 3, pianorollChannel
        midiNote pianorollProgram, (120 * prVelocity) / 100, i + prSemitone + 7, pianorollChannel
        midiNote pianorollProgram, (120 * prVelocity) / 100, i + prSemitone + 10, pianorollChannel
        midiNote pianorollProgram, (120 * prVelocity) / 100, i + prSemitone + 14, pianorollChannel
        
        Case "M9", "M9!"
        midiNote pianorollProgram, (120 * prVelocity) / 100, i + prSemitone, pianorollChannel
        midiNote pianorollProgram, (120 * prVelocity) / 100, i + prSemitone + 4, pianorollChannel
        midiNote pianorollProgram, (120 * prVelocity) / 100, i + prSemitone + 7, pianorollChannel
        midiNote pianorollProgram, (120 * prVelocity) / 100, i + prSemitone + 11, pianorollChannel
        midiNote pianorollProgram, (120 * prVelocity) / 100, i + prSemitone + 14, pianorollChannel
        
        Case "d", "d!"
        midiNote pianorollProgram, (120 * prVelocity) / 100, i + prSemitone, pianorollChannel
        midiNote pianorollProgram, (120 * prVelocity) / 100, i + prSemitone + 3, pianorollChannel
        midiNote pianorollProgram, (120 * prVelocity) / 100, i + prSemitone + 6, pianorollChannel
        
        Case "a", "a!"
        midiNote pianorollProgram, (120 * prVelocity) / 100, i + prSemitone, pianorollChannel
        midiNote pianorollProgram, (120 * prVelocity) / 100, i + prSemitone + 4, pianorollChannel
        midiNote pianorollProgram, (120 * prVelocity) / 100, i + prSemitone + 8, pianorollChannel
        
        Case "D", "D!"
        midiNote pianorollProgram, (120 * prVelocity) / 100, i + prSemitone, pianorollChannel
        midiNote pianorollProgram, (120 * prVelocity) / 100, i + prSemitone + 4, pianorollChannel
        midiNote pianorollProgram, (120 * prVelocity) / 100, i + prSemitone + 7, pianorollChannel
        midiNote pianorollProgram, (120 * prVelocity) / 100, i + prSemitone + 10, pianorollChannel
        
        Case "s", "s!"
        midiNote pianorollProgram, 1, i + prSemitone, pianorollChannel
        
        End Select
        
        If Worksheets("Piano Roll").Range("E4").Value = "On" Then
        lastOffset = PRoffset
        End If
        
      End If
    End If
    



x = x + 1

Next i


If Worksheets("Piano Roll").Range("E4").Value = "Off" Then
lastOffset = PRoffset
'Debug.Print "lastOffset"; lastOffset
End If


End Sub


Sub PRseqFromSaver(pattern, program, prChannel, ByRef arVelocity As Integer, prSemitone, ByRef lastOffsetArray As Variant, curTr, prPartSelect, ByRef sWatch As Stopwatch)

'temphold = sWatch.Elapsed
'Debug.Print curTr; "b4"; sWatch.Elapsed - nowsWatch

'Debug.Print "prog"; curTr; pianorollProgram

pianorollProgram = program
pianorollChannel = prChannel

'prevents multiply error on first run, (curTr,0) doesnt seem to care

Dim prVelocity As Integer

'If lastOffsetArray(curTr, 2) = "" Then
'lastOffsetArray(curTr, 2) = 0
'End If

'Debug.Print lastOffsetArray(curTr, 2)


If pattern = 0 Or pattern = "" Then
pattern = 5709 'because that kinda looks like STOP
'this is my silly solution to stop release in legato mode. this a silent note that on pattern 5709 that triggers a
End If

If lastOffsetArray(curTr, 1) = "" Then
lastOffsetArray(curTr, 1) = pattern
End If

If IsNumeric(Worksheets("PianoSaver").Range("B1").offset(((pattern * 132) - 132), PRoffset + prPartSelect).Value) And Worksheets("PianoSaver").Range("B1").offset(((pattern * 132) - 132), PRoffset + prPartSelect).Value <> "" Then
prVelocity = Worksheets("PianoSaver").Range("B1").offset(((pattern * 132) - 132), PRoffset + prPartSelect).Value
Else
prVelocity = 100
End If

'put value into a var so it doesnt need to read it from the worksheet, does it take less time?
Dim isG24onoff As String
isG24onoff = Worksheets("Arrangement").Range("G24").Value2


'optimizing loop idea (unfinished)
Dim totalStack As Integer
totalStack = Application.WorksheetFunction.CountA(Worksheets("PianoSaver").Range("H16:H143").offset(((pattern * 132) - 132), PRoffset + prPartSelect)) + Application.WorksheetFunction.CountA(Worksheets("PianoSaver").Range("H16:H143").offset(((lastOffsetArray(curTr, 1) * 132) - 132), lastOffsetArray(curTr, 0)))
Dim reachedStack As Integer
'Debug.Print curTr; "ts"; totalStack



'run through the notes on PianoSaver
'quite laggy, reducing by 30 loops saves 10ms (optimizing this would be nice)

Dim i As Integer
Dim y As Integer

For i = 15 To 95   'i start the loop 15 notes above C0 and up til note 95 (B7), this is not a "smart" method but saves 20-30ms so worthwhile (and if u must play at C0 you can still lower notes in the arrangement view)


    If isG24onoff = "Off" Then  'if legato mode off
                
'            Debug.Print "PRoffset"; PRoffset
'            Debug.Print "pattern"; pattern
'            Debug.Print "(curTr, 2)"; lastOffsetArray(curTr, 2) + i
'            Debug.Print "(curTr, 1)"; lastOffsetArray(curTr, 1) * 132
'            Debug.Print "(curTr, 0)"; lastOffsetArray(curTr, 0) + 1
'            Debug.Print "prPartSelect"; prPartSelect + PRoffset

            'Debug.Print Worksheets("PianoSaver").Range("B1").offset(((pattern * 132) - i), PRoffset + prPartSelect).Value
            'Debug.Print Worksheets("PianoSaver").Range("B1").offset(((lastOffsetArray(curTr, 1) * 132) - i), lastOffsetArray(curTr, 0)).Value

'            Debug.Print
            
            
            If Worksheets("PianoSaver").Range("B1").offset(((pattern * 132) - i), PRoffset + prPartSelect).Value2 = "" And Worksheets("PianoSaver").Range("B1").offset(((lastOffsetArray(curTr, 1) * 132) - i), lastOffsetArray(curTr, 0)).Value2 <> "" _
            Or Worksheets("PianoSaver").Range("B1").offset(((pattern * 132) - i), PRoffset + prPartSelect).Value2 <> Worksheets("PianoSaver").Range("B1").offset(((lastOffsetArray(curTr, 1) * 132) - i), lastOffsetArray(curTr, 0)).Value And Right(Worksheets("PianoSaver").Range("B1").offset(((pattern * 132) - i), PRoffset + prPartSelect).Value2, 1) = "!" Then
            
            If PRoffset = 0 Then
            'Debug.Print "ARno"; i; "+"; lastOffsetArray(curTr, 2); "lastOffset"; lastOffsetArray(curTr, 0)
            End If
            
            stopARchordNotes lastOffsetArray(curTr, 1), pianorollProgram, pianorollChannel, 0, i + lastOffsetArray(curTr, 2), Worksheets("PianoSaver").Range("B1").offset(((lastOffsetArray(curTr, 1) * 132) - i), lastOffsetArray(curTr, 0)).Value
            'reachedStack = reachedStack + 1
            End If
    
    End If


    If Worksheets("PianoSaver").Range("B1").offset(((pattern * 132) - i), PRoffset + prPartSelect).Value2 <> "" And _
    Left(Worksheets("PianoSaver").Range("B1").offset(((pattern * 132) - i), PRoffset + prPartSelect).Value2, 1) <> " " Then
    
'    Debug.Print "n1"; Worksheets("PianoSaver").Range("B1").offset(((pattern * 132) - i), PRoffset + prPartSelect).Value; i
'    Debug.Print "n2"; Worksheets("PianoSaver").Range("B1").offset(((pattern * 132) - i), PRoffset + prPartSelect - 1).Value; i
'    Debug.Print Worksheets("Arrangement").Range("G24").Value
    
    If isG24onoff = "Off" And Worksheets("PianoSaver").Range("B1").offset(((pattern * 132) - i), PRoffset + prPartSelect - 1).Value2 = "" _
     Or isG24onoff = "Off" And Right(Worksheets("PianoSaver").Range("B1").offset(((pattern * 132) - i), PRoffset + prPartSelect).Value2, 1) = "!" _
     Or isG24onoff = "On" Then
    
    
    
        'If legato mode is on
        If Worksheets("Arrangement").Range("G24").Value = "On" Then
            If lastOffsetArray(curTr, 0) <> PRoffset + prPartSelect Then   'if a note starts on a new "PRoffset + prPartSelect"/grid position so it should kill the last note/chord
            'Debug.Print "should end own "; lastOffsetArray(curTr, 0)
            
            For y = 0 To 131 'looks at the lastPR notes to find which ones to send a NoteOff to (i could have stored the NoteOns in a array instead of searching for them again
                If Worksheets("PianoSaver").Range("B1").offset(((lastOffsetArray(curTr, 1) * 132) - y), lastOffsetArray(curTr, 0)).Value <> "" Then
                'Debug.Print "note off"; y; "   "; Worksheets("PianoSaver").Range("B1").offset(((lastOffsetArray(curTr, 1) * 132) - y), lastOffsetArray(curTr, 0)).Value; " p "; lastOffsetArray(curTr, 1)
                'moved to a different Sub so i dont fill up space here
                stopARchordNotes lastOffsetArray(curTr, 1), pianorollProgram, pianorollChannel, prVelocity, y + prSemitone, Worksheets("PianoSaver").Range("B1").offset(((lastOffsetArray(curTr, 1) * 132) - y), lastOffsetArray(curTr, 0)).Value
                End If
            Next y
            End If
        End If
    
    
    Select Case Worksheets("PianoSaver").Range("B1").offset(((pattern * 132) - i), PRoffset + prPartSelect).Value
    
    Case "x", "x!"
        'Debug.Print "befor "; PRoffset + prPartSelect; "  "; sWatch.Elapsed - nowsWatch
        'Debug.Print "note on"; i; "+"; prSemitone; "="; i + prSemitone
    midiNote pianorollProgram, (prVelocity / 100) * arVelocity, i + prSemitone, pianorollChannel
    
        'Debug.Print "after "; PRoffset + prPartSelect; "  "; sWatch.Elapsed - nowsWatch
    
    Case "m", "m!"
    midiNote pianorollProgram, (prVelocity / 100) * arVelocity, i + prSemitone, pianorollChannel
    midiNote pianorollProgram, (prVelocity / 100) * arVelocity, i + prSemitone + 3, pianorollChannel
    midiNote pianorollProgram, (prVelocity / 100) * arVelocity, i + prSemitone + 7, pianorollChannel

    
    Case "M", "M!"
    midiNote pianorollProgram, (prVelocity / 100) * arVelocity, i + prSemitone, pianorollChannel
    midiNote pianorollProgram, (prVelocity / 100) * arVelocity, i + prSemitone + 4, pianorollChannel
    midiNote pianorollProgram, (prVelocity / 100) * arVelocity, i + prSemitone + 7, pianorollChannel

    
    Case "m7", "m7!"
    midiNote pianorollProgram, (prVelocity / 100) * arVelocity, i + prSemitone, pianorollChannel
    midiNote pianorollProgram, (prVelocity / 100) * arVelocity, i + prSemitone + 3, pianorollChannel
    midiNote pianorollProgram, (prVelocity / 100) * arVelocity, i + prSemitone + 7, pianorollChannel
    midiNote pianorollProgram, (prVelocity / 100) * arVelocity, i + prSemitone + 10, pianorollChannel

    
    Case "M7", "M7!"
    midiNote pianorollProgram, (prVelocity / 100) * arVelocity, i + prSemitone, pianorollChannel
    midiNote pianorollProgram, (prVelocity / 100) * arVelocity, i + prSemitone + 4, pianorollChannel
    midiNote pianorollProgram, (prVelocity / 100) * arVelocity, i + prSemitone + 7, pianorollChannel
    midiNote pianorollProgram, (prVelocity / 100) * arVelocity, i + prSemitone + 11, pianorollChannel

    
    Case "m9", "m9!"
    midiNote pianorollProgram, (prVelocity / 100) * arVelocity, i + prSemitone, pianorollChannel
    midiNote pianorollProgram, (prVelocity / 100) * arVelocity, i + prSemitone + 3, pianorollChannel
    midiNote pianorollProgram, (prVelocity / 100) * arVelocity, i + prSemitone + 7, pianorollChannel
    midiNote pianorollProgram, (prVelocity / 100) * arVelocity, i + prSemitone + 10, pianorollChannel
    midiNote pianorollProgram, (prVelocity / 100) * arVelocity, i + prSemitone + 14, pianorollChannel

    
    Case "M9", "M9!"
    midiNote pianorollProgram, (prVelocity / 100) * arVelocity, i + prSemitone, pianorollChannel
    midiNote pianorollProgram, (prVelocity / 100) * arVelocity, i + prSemitone + 4, pianorollChannel
    midiNote pianorollProgram, (prVelocity / 100) * arVelocity, i + prSemitone + 7, pianorollChannel
    midiNote pianorollProgram, (prVelocity / 100) * arVelocity, i + prSemitone + 11, pianorollChannel
    midiNote pianorollProgram, (prVelocity / 100) * arVelocity, i + prSemitone + 14, pianorollChannel

    
    Case "d", "d!"
    midiNote pianorollProgram, (prVelocity / 100) * arVelocity, i + prSemitone, pianorollChannel
    midiNote pianorollProgram, (prVelocity / 100) * arVelocity, i + prSemitone + 3, pianorollChannel
    midiNote pianorollProgram, (prVelocity / 100) * arVelocity, i + prSemitone + 6, pianorollChannel

    
    Case "a", "a!"
    midiNote pianorollProgram, (prVelocity / 100) * arVelocity, i + prSemitone, pianorollChannel
    midiNote pianorollProgram, (prVelocity / 100) * arVelocity, i + prSemitone + 4, pianorollChannel
    midiNote pianorollProgram, (prVelocity / 100) * arVelocity, i + prSemitone + 8, pianorollChannel

    
    Case "D", "D!"
    midiNote pianorollProgram, (prVelocity / 100) * arVelocity, i + prSemitone, pianorollChannel
    midiNote pianorollProgram, (prVelocity / 100) * arVelocity, i + prSemitone + 4, pianorollChannel
    midiNote pianorollProgram, (prVelocity / 100) * arVelocity, i + prSemitone + 7, pianorollChannel
    midiNote pianorollProgram, (prVelocity / 100) * arVelocity, i + prSemitone + 10, pianorollChannel

    
    Case "s", "s!"
    midiNote pianorollProgram, 1, i + prSemitone, pianorollChannel

    
    End Select
    
    'reachedStack = reachedStack + 1
    
    
    If isG24onoff = "On" Then
        lastOffsetArray(curTr, 0) = PRoffset + prPartSelect
        lastOffsetArray(curTr, 1) = pattern
        lastOffsetArray(curTr, 2) = prSemitone
    End If
    
    

    End If
    
    
    End If

'an idea to optimize the loop so it doesnt need to run as long thus saving time
'If reachedStack - totalStack Then
'i = 127
'End If

Next i

'Debug.Print curTr; "rs"; reachedStack


If isG24onoff = "Off" Then
lastOffsetArray(curTr, 0) = PRoffset + prPartSelect
lastOffsetArray(curTr, 1) = pattern
lastOffsetArray(curTr, 2) = prSemitone
'Debug.Print "lastOffset"; lastOffset
End If

'Debug.Print curTr; "a4"; sWatch.Elapsed - nowsWatch - temphold

End Sub

Sub stopARchordNotes(pattern, program, prChannel, prVelocity, ByRef prSemitone As Integer, noteStamp)

Dim i As Integer
'Debug.Print "noteStamp "; prSemitone; "  "; noteStamp

Select Case noteStamp
    
    Case "x", "x!"
        'Debug.Print "befor "; PRoffset + prPartSelect; "  "; sWatch.Elapsed - nowsWatch
        'Debug.Print "noteStamp "; prSemitone; "  "; noteStamp
    justStopNote2 pianorollProgram, 0, i + prSemitone, pianorollChannel
        'Debug.Print "after "; PRoffset + prPartSelect; "  "; sWatch.Elapsed - nowsWatch

    
    Case "m", "m!"
    justStopNote2 pianorollProgram, 0, i + prSemitone, pianorollChannel
    justStopNote2 pianorollProgram, 0, i + prSemitone + 3, pianorollChannel
    justStopNote2 pianorollProgram, 0, i + prSemitone + 7, pianorollChannel


    Case "M", "M!"
    justStopNote2 pianorollProgram, 0, i + prSemitone, pianorollChannel
    justStopNote2 pianorollProgram, 0, i + prSemitone + 4, pianorollChannel
    justStopNote2 pianorollProgram, 0, i + prSemitone + 7, pianorollChannel

    Case "m7", "m7!"
    justStopNote2 pianorollProgram, 0, i + prSemitone, pianorollChannel
    justStopNote2 pianorollProgram, 0, i + prSemitone + 3, pianorollChannel
    justStopNote2 pianorollProgram, 0, i + prSemitone + 7, pianorollChannel
    justStopNote2 pianorollProgram, 0, i + prSemitone + 10, pianorollChannel

    Case "M7", "M7!"
    justStopNote2 pianorollProgram, 0, i + prSemitone, pianorollChannel
    justStopNote2 pianorollProgram, 0, i + prSemitone + 4, pianorollChannel
    justStopNote2 pianorollProgram, 0, i + prSemitone + 7, pianorollChannel
    justStopNote2 pianorollProgram, 0, i + prSemitone + 11, pianorollChannel

    Case "m9", "m9!"
    justStopNote2 pianorollProgram, 0, i + prSemitone, pianorollChannel
    justStopNote2 pianorollProgram, 0, i + prSemitone + 3, pianorollChannel
    justStopNote2 pianorollProgram, 0, i + prSemitone + 7, pianorollChannel
    justStopNote2 pianorollProgram, 0, i + prSemitone + 10, pianorollChannel
    justStopNote2 pianorollProgram, 0, i + prSemitone + 14, pianorollChannel

    Case "M9", "M9!"
    justStopNote2 pianorollProgram, 0, i + prSemitone, pianorollChannel
    justStopNote2 pianorollProgram, 0, i + prSemitone + 4, pianorollChannel
    justStopNote2 pianorollProgram, 0, i + prSemitone + 7, pianorollChannel
    justStopNote2 pianorollProgram, 0, i + prSemitone + 11, pianorollChannel
    justStopNote2 pianorollProgram, 0, i + prSemitone + 14, pianorollChannel
    
    Case "d", "d!"
    justStopNote2 pianorollProgram, 0, i + prSemitone, pianorollChannel
    justStopNote2 pianorollProgram, 0, i + prSemitone + 3, pianorollChannel
    justStopNote2 pianorollProgram, 0, i + prSemitone + 6, pianorollChannel

    Case "a", "a!"
    justStopNote2 pianorollProgram, 0, i + prSemitone, pianorollChannel
    justStopNote2 pianorollProgram, 0, i + prSemitone + 4, pianorollChannel
    justStopNote2 pianorollProgram, 0, i + prSemitone + 8, pianorollChannel

    Case "D", "D!"
    justStopNote2 pianorollProgram, 0, i + prSemitone, pianorollChannel
    justStopNote2 pianorollProgram, 0, i + prSemitone + 4, pianorollChannel
    justStopNote2 pianorollProgram, 0, i + prSemitone + 7, pianorollChannel
    justStopNote2 pianorollProgram, 0, i + prSemitone + 10, pianorollChannel

    Case "s", "s!"
    justStopNote2 pianorollProgram, 1, i + prSemitone, pianorollChannel
    
    End Select



End Sub



Sub PRclearCosmetics()

Worksheets("Piano Roll").Range("H5").offset(0, PRoffset).Interior.ColorIndex = 34

End Sub



Sub PRstopOkay()

If imPlayin = True Then

loopsLeft = 0
stepsDone = 64 '32 '16

'counter = 0
stopItMIDIAgain
imPlayin = False

End If

End Sub

Sub PRLooper()


If loopsLeft > 1 Then

loopsLeft = loopsLeft - 1

PRLoopSeq
PRoffset = PRoffset

End If

imPlayin = False


sWatch.Pause
stopItMIDIAgain


End Sub

Sub octaveUp()

Dim x As Integer
Dim i As Integer

If Application.Selection.Count > 1 And ActiveSheet.Name = "Piano Roll" And ActiveSheet.Name = "Piano Roll" And Application.Selection.Row > 10 And Application.Selection.Column > 5 Then

For x = 0 To Application.Selection.Count

    'Debug.Print Application.Selection(x + 1).Value
    
    If Application.Selection(x + 1).Value <> "" Then
        
        
        For i = 0 To Application.Selection.Rows.Count - 1
        
        'Application.Selection.offset(-1, 0).Value = Application.Selection.Value
        Application.Selection.offset(-12, 0).Rows(i + 1).Value = Application.Selection.Rows(i + 1).Value
        
        If i < 12 Then
        Application.Selection.Rows(i + 1).Value = ""
        
        End If
        
        Next i
        
        'End If
        
        x = Application.Selection.Count
        ActiveWindow.ScrollRow = ActiveWindow.ScrollRow - 12
        
    End If
        
Next x
    
Else

Range("H12:BS131").Value = Range("H24:BS143").Value
Range("H132:BS143").Value = ""
ActiveWindow.ScrollRow = ActiveWindow.ScrollRow - 12

End If



End Sub

Sub octaveDown()

Dim x As Integer
Dim i As Integer

If Application.Selection.Count > 1 And ActiveSheet.Name = "Piano Roll" And Application.Selection.Row > 10 And Application.Selection.Column > 5 Then

For x = 0 To Application.Selection.Count

    'Debug.Print Application.Selection(x + 1).Value
    
    If Application.Selection(x).Value <> "" Then
        
        'Debug.Print Application.Selection.Rows.Count
        For i = 0 To Application.Selection.Rows.Count - 1
        
        'Application.Selection.offset(-1, 0).Value = Application.Selection.Value
        Application.Selection.offset(12, 0).Rows((Application.Selection.Rows.Count) - i).Value = Application.Selection.Rows((Application.Selection.Rows.Count) - i).Value
        
        If Application.Selection.Rows.Count - i < 12 Then
        Application.Selection.Rows(Application.Selection.Rows.Count - i).Value = ""
        
        End If
        
        

        Next i
        
        
        

        
        x = Application.Selection.Count
        ActiveWindow.ScrollRow = ActiveWindow.ScrollRow + 12
        
    End If
        
Next x
    
Else

Range("H24:BS143").Value = Range("H12:BS131").Value
Range("H12:BS23").Value = ""
ActiveWindow.ScrollRow = ActiveWindow.ScrollRow + 12

End If




End Sub

Sub transposeUp()

Dim x As Integer
Dim i As Integer


If Application.Selection.Count > 1 And ActiveSheet.Name = "Piano Roll" And ActiveSheet.Name = "Piano Roll" And Application.Selection.Row > 10 And Application.Selection.Column > 5 Then

For x = 0 To Application.Selection.Count

    'Debug.Print Application.Selection(x + 1).Value
    
    If Application.Selection(x + 1).Value <> "" Then
        
        
        For i = 0 To Application.Selection.Rows.Count - 1
        
        'Application.Selection.offset(-1, 0).Value = Application.Selection.Value
        Application.Selection.offset(-1, 0).Rows(i + 1).Value = Application.Selection.Rows(i + 1).Value
        

        Next i
        
        Application.Selection.Rows(Application.Selection.Rows.Count).Value = ""
        
        'End If
        
        x = Application.Selection.Count
        
    End If
        
Next x
    
Else

Range("H12:BS142").Value = Range("H13:BS143").Value
Range("H143:BS143").Value = ""

End If



'Range("H11:BS11").Value = ""



End Sub

Sub transposeDown()

Dim x As Integer
Dim i As Integer


'Debug.Print "top selection"; Application.Selection.Row

If Application.Selection.Count > 1 And ActiveSheet.Name = "Piano Roll" And Application.Selection.Row > 10 And Application.Selection.Column > 5 Then

For x = 0 To Application.Selection.Count

    'Debug.Print Application.Selection(x + 1).Value
    
    If Application.Selection(x).Value <> "" Then
        
        'Debug.Print Application.Selection.Rows.Count
        For i = 0 To Application.Selection.Rows.Count - 1
        
        'Application.Selection.offset(-1, 0).Value = Application.Selection.Value
        Application.Selection.offset(1, 0).Rows((Application.Selection.Rows.Count) - i).Value = Application.Selection.Rows((Application.Selection.Rows.Count) - i).Value
        
            'If i = Application.Selection.Rows.Count Then
            
            'End If
        Next i
        
        Application.Selection.Rows(1).Value = ""
        

        
        x = Application.Selection.Count
        
    End If
        
Next x
    
Else

Range("H13:BS143").Value = Range("H12:BS142").Value
Range("H12:BS12").Value = ""

End If



End Sub


Sub openPianoPattern()

If Worksheets("Piano Roll").Range("BB2").Value >= 1 Then


savePianoPattern

Worksheets("Piano Roll").Range("H16:BS143").Value = Worksheets("PianoSaver").Range("B1:BM132").offset((Worksheets("Piano Roll").Range("BB2").Value * 132) - 127, 0).Value
Worksheets("Piano Roll").Range("H6:BS6").Value = Worksheets("PianoSaver").Range("B1:BM1").offset((Worksheets("Piano Roll").Range("BB2").Value * 132) - 132, 0).Value

Worksheets("PianoSaver").Range("A1048576").Value = Worksheets("Piano Roll").Range("BB2").Value
'Worksheets("Piano Roll").Range("D52:D54").offset(0, 0).Interior.ColorIndex = 40
End If



End Sub

Sub savePianoPattern()

Dim lastPianoPattern As Integer

lastPianoPattern = Worksheets("PianoSaver").Range("A1048576").Value

Worksheets("PianoSaver").Range("B5:BM132").offset((lastPianoPattern * 132) - 131, 0).Value = Worksheets("Piano Roll").Range("H16:BS143").Value
Worksheets("PianoSaver").Range("B1:BM1").offset((lastPianoPattern * 132) - 132, 0).Value = Worksheets("Piano Roll").Range("H6:BS6").Value


End Sub

Sub savePianoPatternOnPlay()

Dim lastPianoPattern As Integer

lastPianoPattern = Worksheets("PianoSaver").Range("A1048576").Value

Worksheets("PianoSaver").Range("B5:BM132").offset((lastPianoPattern * 132) - 131, 0).Value = Worksheets("Piano Roll").Range("H16:BS143").Value
Worksheets("PianoSaver").Range("B1:BM1").offset((lastPianoPattern * 132) - 132, 0).Value = Worksheets("Piano Roll").Range("H6:BS6").Value

End Sub


Sub copyPiano()

Worksheets("PianoSaver").Range("BR1:ED132").Value = Worksheets("Piano Roll").Range("H12:BS143").Value

Worksheets("Piano Roll").Range("BU1:BY3").Interior.ColorIndex = 40
Worksheets("Piano Roll").Range("BU4:BY6").Interior.ColorIndex = 0

End Sub


Sub pastePiano()

Worksheets("Piano Roll").Range("H12:BS143").Value = Worksheets("PianoSaver").Range("BR1:ED132").Value

Worksheets("Piano Roll").Range("BU4:BY6").Interior.ColorIndex = 40
Worksheets("Piano Roll").Range("BU1:BY3").Interior.ColorIndex = 0


''change view to notes
Dim x As Integer
Dim y As Integer
For x = 0 To 15
    For y = 0 To 131
    
        If Worksheets("Piano Roll").Range("H12").offset(y, x).Value <> "" And Left(Worksheets("Piano Roll").Range("H12").offset(y, x).Value, 1) <> " " Then
        
        'Debug.Print "first coordinate "; y; "  "; x
        
        If y <> 0 Then
        ActiveWindow.ScrollRow = y
        Else
        ActiveWindow.ScrollRow = 1
        End If
        
        y = 131
        x = 15
        
        End If
    
    Next y
Next x


End Sub


Sub duplicate4barsPiano()

Worksheets("Piano Roll").Range("X12:AM143").Value = Worksheets("Piano Roll").Range("H12:W143").Value

End Sub

Sub duplicate8barsPiano()

Worksheets("Piano Roll").Range("AN12:BS143").Value = Worksheets("Piano Roll").Range("H12:AM143").Value

End Sub


Sub PRpatUP()

Worksheets("Piano Roll").Range("BB2").Value = Worksheets("Piano Roll").Range("BB2").Value + 1

End Sub

Sub PRpatDown()

Worksheets("Piano Roll").Range("BB2").Value = Worksheets("Piano Roll").Range("BB2").Value - 1

End Sub


Sub learninghow2midi()

'program,vel,pitch,channel

midiNote 0, 100, 60, 1
midiNote 10, 100, 50, 0
midiNote 30, 100, 40, 0


justStopNote 20, 100, 40, 0
justStopNote 0, 100, 50, 1
justStopNote 0, 100, 40, 0


End Sub

Sub tempselecttt()



End Sub


Sub prColumnCount()

'Debug.Print Application.WorksheetFunction.CountA(Range("H16:H143").offset(0, 16))


End Sub





