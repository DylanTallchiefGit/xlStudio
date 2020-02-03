Attribute VB_Name = "xl2xml"
Option Explicit

Sub mshbox()

MsgBox "Once you press OK and Ableton project file will be created in the same folder as this excel document. It may take a while so please be patient!" & vbCrLf _
& vbCrLf & "Note: If you have 7zip installed, an als file will be created from from the xml file (""TesterB"" which can be ignored/deleted). If not you will have to ""gzip"" the xml file yourself and rename the extension to "".als"""

MsgBox "Once you press OK, an Ableton project file will be created in the same folder as this excel document. It may take a while so please be patient!"

End Sub

Sub alsGen()

MsgBox "Once you press OK, an Ableton project file will be created in the same folder as this excel document. It may take a while so please be patient!"
Worksheets("Arrangement").Range("AC22").Value = "0%"

Dim arrHeader() As String, Max As Integer, Col As Integer, arrContent() As String, Row As Integer
Dim filesys, testfile, XTab As String

Dim cellRange As String
Dim STab As String
Dim write2File As Boolean
cellRange = "E30" ' "A21"

STab = "    "
Max = 1
write2File = True

Dim endArrangement As Integer
Dim startFinder As Integer
startFinder = 0

endArrangement = 0
Do
endArrangement = endArrangement + 1

Loop Until LCase(Left(Worksheets("Arrangement").Range("H29").offset(0, (startFinder - 1) + endArrangement - 1).Value, 1)) = "e" Or endArrangement = Worksheets("Arrangement").Cells(29, Columns.Count).End(xlToLeft).Column


If endArrangement = Worksheets("Arrangement").Cells(29, Columns.Count).End(xlToLeft).Column Then 'this means no E was found
'Debug.Print "no e found"

    endArrangement = 0
    Dim i As Variant
    For i = 0 To howManyARTracks() - 1
    
        If Worksheets("Arrangement").Cells(31 + (i * 3), Columns.Count).End(xlToLeft).Column - 6 - (startFinder - 1) > endArrangement Then
        'Debug.Print Worksheets("Arrangement").Cells(31 + (i * 3), Columns.Count).End(xlToLeft).Column
        
        'Debug.Print "startFinder"; startFinder
        
        endArrangement = Worksheets("Arrangement").Cells(31 + (i * 3), Columns.Count).End(xlToLeft).Column - 6 - (startFinder - 1)
        End If
    Next i
End If

'Debug.Print "endArrangement"; endArrangement
Dim Id As Integer

Dim p1 As String
Dim p2 As String

p1 = ActiveWorkbook.Path & "\testerB.xml"
'Debug.Print p1
'p2 = CreateObject("WScript.Shell").specialfolders("Desktop") & "\testerB.xml"
'Debug.Print CreateObject("WScript.Shell").specialfolders("Desktop")


Set filesys = CreateObject("Scripting.FileSystemObject")
'Set testfile = filesys.CreateTextFile("C:\Users\Dylan\Desktop\testerB.xml", True) 'this needs to change lol
Set testfile = filesys.CreateTextFile(p1, True)
'Set testfile = filesys.CreateTextFile(ActiveWorkbook.Path, True)
'Set testfile = filesys.CreateTextFile("\Desktop\testerB.xml", True)

testfile.WriteLine "<?xml version=""1.0"" encoding=""UTF-8""?>"
'testfile.WriteLine "<Ableton MajorVersion=""5"" MinorVersion=""10.0_377"" SchemaChangeCount=""3"" Creator=""Ableton Live 10.1.3"" Revision=""3794fc29d53937b0dbd06319470833698817d635"">"
testfile.WriteLine "<Ableton MajorVersion=""5"" MinorVersion=""10.0_370"" SchemaChangeCount=""3"" Creator=""Ableton Live 10.0.4"" Revision=""1922243a5a00c4566607e21183ba7acad4632272"">"


testfile.WriteLine "    <LiveSet>"

testfile.WriteLine "        <NextPointeeId Value=""20334"" />"
testfile.WriteLine "        <OverwriteProtectionNumber Value=""2561"" />"
testfile.WriteLine "        <LomId Value=""0"" />"
testfile.WriteLine "        <LomIdView Value=""0"" />"



testfile.WriteLine "        <Tracks>"

     Dim nameTracks(100) As Variant  'i wish i didnt have to define the length of the array
     Dim amountTracks As Integer
     amountTracks = Arrangement.howManyARTracks()
     
'    Do
'        Debug.Print Worksheets("Arrangement").Range(cellRange).offset(amountTracks + 1, 0).Value
'        nameTracks(amountTracks) = Range(cellRange).offset(amountTracks + 1, 0).Value 'where it starts
'        amountTracks = amountTracks + 1
'    Loop Until Range(cellRange).offset(amountTracks, 0).Value = ""
    
'    amountTracks = amountTracks - 1
    'Debug.Print amountTracks



    Randomize
    For i = 1 To amountTracks 'Int((16 - 1 + 1) * Rnd + 1)
    
        Randomize
        Dim midiOrAudio As Integer
        midiOrAudio = 2 'Int((2 - 1 + 1) * Rnd + 1)
            
            If midiOrAudio = 1 Then
            testfile.WriteLine "            <AudioTrack Id=""" & i & """>"
            Else
            testfile.WriteLine "            <MidiTrack Id=""" & i & """>"
            End If
            
        Dim y As Integer
        For y = 1 To 9
        testfile.WriteLine Worksheets("xmlpasta").Cells(y, 12).Value
        Next y
          'Debug.Print Worksheets("Arrangement").Cells(31 + ((i - 1) * 3), 5).Value
            
        testfile.WriteLine "                    <EffectiveName Value=""" & Worksheets("Arrangement").Cells(31 + ((i - 1) * 3), 5).Value & """ />"
        testfile.WriteLine "                    <UserName Value=""" & Worksheets("Arrangement").Cells(31 + ((i - 1) * 3), 5).Value & """ />"
            
            If midiOrAudio = 1 Then
                'audio track
                For y = 12 To 434
                    If y = 431 Then
                    testfile.WriteLine "track plugins (devices) go here"
                    Else
                    testfile.WriteLine Worksheets("xmlpasta").Cells(y, 12).Value
                    End If
                Next y
            Else
                'midi track
                For y = 1 To 262
                    If y = 4 Then
                    testfile.WriteLine "                <ColorIndex Value=""" & i + 90 & """ />"
                    Else
                    testfile.WriteLine Worksheets("xmlpasta").Cells(y, 29).Value
                    End If
                Next y
                
                
                    'create midi clips here
                    testfile.WriteLine "                                <Events>"
                    
'                    Randomize
'                    length = (Int((16 - 1 + 1) * Rnd + 1)) * 8
'                    Randomize
'                    clipLegnth = (Int((16 - 1 + 1) * Rnd + 1)) * 8
                    
                    Dim clipFinder As Integer
                    clipFinder = 0
                    'amountTracks = 0
                    
                    
                    Do
                            
                        Do
                        
                    Dim a As Integer
                    Dim b As Integer
                    Dim c As Integer
                    Dim d As Integer
                    Dim clipLength As Integer
                    
                    a = 0
                    If Worksheets("Arrangement").Range("H31").offset((i - 1) * 3, clipFinder).Value <> "" Then
                    clipLength = 1
                        For a = 1 To 3
                        If Worksheets("Arrangement").Range("H31").offset((i - 1) * 3, clipFinder + a).Value = "." Then
                        clipLength = clipLength + 1
                        ElseIf Worksheets("Arrangement").Range("H31").offset((i - 1) * 3, clipFinder + a).Value = "" Or IsNumeric(Worksheets("Arrangement").Range("H31").offset((i - 1) * 3, clipFinder + a).Value) = True Then
                        a = 3
                        End If
                        Next a
                    End If
                        

                        If Worksheets("Arrangement").Range("H31").offset((i - 1) * 3, clipFinder).Value <> "" Then
                        'Debug.Print Range(cellRange).offset(i, 0).Value & " " & Range(cellRange).offset(i, clipFinder + 7).Value
                        'Debug.Print clipFinder
                        
                            testfile.WriteLine _
                            "                                    <MidiClip Id=""" & clipFinder & """ Time=""" & clipFinder * 4 & """>"

                                For a = 2 To 79
                                    If a = 15 Then
                                    testfile.WriteLine _
                                    "                                        <Name Value=""" & Worksheets("Arrangement").Range("H31").offset((i - 1) * 3, clipFinder).Value & """ />"
                                    ElseIf a = 4 Then
                                    testfile.WriteLine _
                                    "                                        <CurrentStart Value=""" & clipFinder * 4 & """ />"
                                    ElseIf a = 5 Then
                                    testfile.WriteLine _
                                    "                                        <CurrentEnd Value=""" & (clipFinder * 4) + clipLength * 4 & """ />"
                                    ElseIf a = 8 Then
                                    testfile.WriteLine _
                                    "                                            <LoopEnd Value=""" & clipLength * 4 & """ />"
                                    ElseIf a = 17 Then 'colours
                                            If TypeName(Range(cellRange).offset(i, clipFinder + 7).Value) <> "String" Then
                                        'extraColour = CInt(Range(cellRange).offset(i, clipFinder + 7).Value)
                                        'Debug.Print "Extra COLOURRRR: "; extraColour
                                        testfile.WriteLine "                                        <ColorIndex Value=""" & i & """ />"
                                        Else
                                        testfile.WriteLine "                                        <ColorIndex Value=""" & i & """ />"
                                        End If
                                    ElseIf a = 64 Then
                                    testfile.WriteLine "                                            <KeyTracks>"
                                    
                                    
                                    
                                    Dim pattern As Integer
                                    Dim partSelect As Integer
                                    Dim curVel As Integer
                                    Dim keyTrackID As Integer
                                    
                                    Select Case Worksheets("Arrangement").Cells(31 + ((i - 1) * 3), 5).Value
                                        
                                    Case "Drums", "drums", "drum", "Drum" 'drum patterns
                                    
                                       
                                 
                                        If Worksheets("Arrangement").Range("H31").offset((i - 1) * 3, clipFinder).Value = "." Then
                                        'partSelect = 16
                                        'pattern = Worksheets("Arrangement").Range("H31").offset((i - 1) * 3, clipFinder - 1).Value
                                        Else
                                        pattern = Worksheets("Arrangement").Range("H31").offset((i - 1) * 3, clipFinder).Value
                                        End If
                                        
                                        Id = 0
                                        keyTrackID = 0
                                        Dim notePitch As Integer
                                        notePitch = 60
                                        
                                        'only 8 because drums
                                        
                                        For b = 0 To 7
                                        
                                        'Debug.Print Worksheets("Arrangement").Range("H31").offset((i - 1) * 3, clipFinder).Value
                                    If IsNumeric(pattern) = True Then
                                        
                                        Dim fillTrack As Boolean
                                        fillTrack = False
                                        For c = 0 To (clipLength * 16) - 1
                                            If Worksheets("PatternSaver").Cells((24 * pattern) - 23 + (b * 3), 4 + c).Value = "x" Then
                                            fillTrack = True
                                            c = (clipLength * 16) - 1
                                            End If
                                        Next c
                                        
                                        If fillTrack = True Then
                                        'this below method has way too many issues so moved to fillTrack = true method instead
'                                    'If Application.WorksheetFunction.CountA(Worksheets("PatternSaver").Range("D1:S1").offset(((24 * pattern) - 24) + (b * 3), 0)) > 0 Then
'
                                        testfile.WriteLine "                                                <KeyTrack Id=""" & keyTrackID & """>"
                                        testfile.WriteLine "                                                    <Notes>"
                                      
                                        For c = 0 To (clipLength * 16) - 1

                                        If Worksheets("PatternSaver").Cells((24 * pattern) - 23 + (b * 3), 4 + c).Value = "x" Then


                                            If Worksheets("PatternSaver").Cells((24 * pattern) - 23 + (b * 3) + 1, 4 + c).Value <> "" _
                                            And IsNumeric(Worksheets("PatternSaver").Cells((24 * pattern) - 23 + (b * 3) + 1, 4 + c).Value) = True Then
                                            curVel = Worksheets("PatternSaver").Cells((24 * pattern) - 23 + (b * 3) + 1, 4 + c).Value

                                            ElseIf Worksheets("PatternSaver").Cells((24 * pattern) - 23 + (b * 3) + 1, 2).Value <> "" _
                                            And IsNumeric(Worksheets("PatternSaver").Cells((24 * pattern) - 23 + (b * 3) + 1, 2).Value) = True Then
                                            curVel = Worksheets("PatternSaver").Cells((24 * pattern) - 23 + (b * 3) + 1, 2).Value
                                            Else
                                            curVel = 100
                                            End If

                                        testfile.WriteLine "                                                        <MidiNoteEvent Time=""" & Replace(c / 4, ",", ".") & _
                                        """ Duration=""0.25"" Velocity=""" & curVel & """ OffVelocity=""64"" IsEnabled=""true"" NoteId=""" & Id & """ />"
                                        Id = Id + 1
                                        End If
''
                                        Next c
'
                                        testfile.WriteLine "                                                    </Notes>"
'
'
'                                        If Worksheets("Arrangement").Range("G23").Value > 1 Then
                                        testfile.WriteLine "                                                    <MidiKey Value=""" & 60 + b & """ />"
'                                        ElseIf Worksheets("Arrangement").Range("G23").Value = 1 Then
'
'
'                                            If Worksheets("PatternSaver").Cells((24 * pattern) - 23 + (b * 3) + 2, 2).Value <> "" _
'                                            And IsNumeric(Worksheets("PatternSaver").Cells((24 * pattern) - 23 + (b * 3) + 2, 2).Value) = True Then
'                                            notePitch = Worksheets("PatternSaver").Cells((24 * pattern) - 23 + (b * 3) + 2, 2).Value
'                                            Else
'                                            notePitch = notePitch + b
'                                            End If
'                                        'Debug.Print notePitch
'                                        testfile.WriteLine "                                                    <MidiKey Value=""" & notePitch & """ />"
'                                        End If
'
                                        testfile.WriteLine "                                                </KeyTrack>"
                                        keyTrackID = keyTrackID + 1
                                    
                                    
                                    End If
                                    
                                    End If
                                        
                                        
                                        Next b
                                        
'                                        For b = 65 To 73
'                                        testfile.WriteLine Worksheets("xmlpasta").Cells(b, 70).Value
'                                        Next b
                                            

                                    Case Else   'PR patterns
                                        
                                    keyTrackID = 0
                                    Id = 0
                                    
                                    partSelect = 0
                                        
                                        
                                        'Debug.Print Worksheets("Arrangement").Range("H31").offset((i - 1) * 3, clipFinder).Value

                                        If Worksheets("Arrangement").Range("H31").offset((i - 1) * 3, clipFinder).Value = "." Then
                                        
                                        For b = 1 To 3
                                        If IsNumeric(Worksheets("Arrangement").Range("H31").offset((i - 1) * 3, clipFinder - b).Value) = True And Worksheets("Arrangement").Range("H31").offset((i - 1) * 3, clipFinder - b).Value <> "" Then
                                        pattern = Worksheets("Arrangement").Range("H31").offset((i - 1) * 3, clipFinder - b).Value
                                        partSelect = b * 16
                                        b = 3
                                        End If
                                        Next b
                                        
                                        Else

                                        pattern = Worksheets("Arrangement").Range("H31").offset((i - 1) * 3, clipFinder).Value
                                        End If

                                        'Debug.Print pattern
                                        'left to right, bottom to top order best?
                                        For b = 0 To 127
                                        
                                        For c = 0 To (clipLength * 16) - 1
                                        If Left(Worksheets("PianoSaver").Range("B1").offset(((pattern * 132) - 132) + 5 + (127 - b), c).Value, 1) = "x" Then
                                        fillTrack = True
                                        c = (clipLength * 16) - 1
                                        End If
                                        Next c

                                    If fillTrack = True Then
                                    'If Application.WorksheetFunction.CountA(Worksheets("PianoSaver").Range("B1:Q1").offset(((132 * pattern) - 132) + 5 + (127 - b), 0 + partSelect)) > 0 Then
                                        'Debug.Print pattern
                                        
                                        testfile.WriteLine "                                                <KeyTrack Id=""" & keyTrackID & """>"
                                        testfile.WriteLine "                                                    <Notes>"

                                        For c = 0 To (clipLength * 16) - 1 '15
                                        Dim noteLength As Integer
                                       
                                        If Left(Worksheets("PianoSaver").Range("B1").offset(((pattern * 132) - 132) + 5 + (127 - b), c + partSelect).Value, 1) = "x" Then
                                        
                                        If Worksheets("PianoSaver").Range("B1").offset(((pattern * 132) - 132), c + partSelect).Value <> "" Then
                                        curVel = Worksheets("PianoSaver").Range("B1").offset(((pattern * 132) - 132), c + partSelect).Value
                                        Else
                                        curVel = 100
                                        End If
                                        
                                        For d = 1 To (clipLength * 16) - c
                                        
                                        If Worksheets("PianoSaver").Range("B1").offset(((pattern * 132) - 132) + 5 + (127 - b), c + d + partSelect).Value = "" Or Right(Worksheets("PianoSaver").Range("B1").offset(((pattern * 132) - 132) + 5 + (127 - b), c + d + partSelect).Value, 1) = "!" Then
                                        noteLength = d
                                        d = (clipLength * 16) - c
                                        Else 'assume full bar
                                        noteLength = (clipLength * 16) - c
                                        End If
                                        Next d
                                        
                                        'Debug.Print noteLength
                                        
                                        testfile.WriteLine "                                                        <MidiNoteEvent Time=""" & Replace(c / 4, ",", ".") & _
                                        """ Duration=""" & Replace(noteLength / 4, ",", ".") & """ Velocity=""" & curVel & """ OffVelocity=""64"" IsEnabled=""true"" NoteId=""" & Id & """ />"
                                        
                                        c = c + noteLength - 1
                                        
                                        Id = Id + 1
                                        End If

                                        Next c

                                        testfile.WriteLine "                                                    </Notes>"
                                        testfile.WriteLine "                                                    <MidiKey Value=""" & b & """ />"
                                        testfile.WriteLine "                                                </KeyTrack>"
                                        keyTrackID = keyTrackID + 1
                                    End If

                                        Next b
                                        
                                        
                                    End Select

                                    testfile.WriteLine "                                            </KeyTracks>"
                                    
                                    Else
                                    testfile.WriteLine Worksheets("xmlpasta").Cells(a, 41).Value
                                    End If
                                Next a

                        
                        
                        End If
                        clipFinder = clipFinder + clipLength
                        
                        Loop Until Worksheets("Arrangement").Range("H29").offset(0, clipFinder).Value = "e" Or clipFinder > endArrangement
                        
                        clipFinder = 0
                        y = y + 1
                            
                    Loop Until Worksheets("Arrangement").Range("E31").offset((y - 1) * 3, 0).Value = ""
                    

                    
                    
                    
                    testfile.WriteLine "                                </Events>"
                
                'rest of midi track pasta
                For y = 264 To 802
                    If y = 798 Then
                    
                    'add midi pitch device
                    For a = 1 To 206
                    testfile.WriteLine Worksheets("xmlpasta").Cells(a, 65).Value
                    Next a
                    Else
                    testfile.WriteLine Worksheets("xmlpasta").Cells(y, 29).Value
                    End If
                Next y
            End If
    
    Worksheets("Arrangement").Range("AC22").Value = round(((i / amountTracks) * 100), 1) & "%"
    Next i
'end tracks

testfile.WriteLine "        </Tracks>"


For i = 1 To 596

If i = 39 Then
testfile.WriteLine "                                <FloatEvent Id=""0"" Time=""-63072000"" Value=""" & Worksheets("Arrangement").Range("G22").Value & """ />"

ElseIf i = 201 Then
testfile.WriteLine "                        <Manual Value=""" & Replace(Worksheets("Arrangement").Range("G22").Value, ",", ".") & """ />"

Else
testfile.WriteLine Worksheets("xmlpasta").Cells(i, 1).Value

End If

Next i


                      ''locator stuff
    testfile.WriteLine "           <Locators>"
    testfile.Close
    
                'i use different file write method from here on
    'Open "C:\Users\Dylan\Desktop\testerB.xml" For Append As #1
    Open p1 For Append As #1
    
  '  randomHouse
  
'    sectionFinder = 0
'    Dim nameLocators(999) As Variant
'    Do
'
'    If Range(cellRange).offset(0, sectionFinder + 7).Value <> "" Then
'    Debug.Print Range(cellRange).offset(0, sectionFinder + 7).Value
'    nameLocators((sectionFinder - 1) / 8) = Range(cellRange).offset(0, sectionFinder + 7).Value
'    'Debug.Print "Array " & nameLocators((sectionFinder - 1) / 8)
'    End If
'    sectionFinder = sectionFinder + 1
'
'    Loop Until Range(cellRange).offset(0, sectionFinder).Value = "end"
'
'
          'For i = 0 To nameLocators(i) = ""   'would be rad if u could do it this way
          Dim LocatorName As String
          
          For i = 0 To 3
          
          
          Select Case i
          Case 0
          LocatorName = "SUBSCRIBE"
          Case 1
          LocatorName = "TO"
          Case 2
          LocatorName = "DYLAN"
          Case 3
          LocatorName = "TALLCHIEF"
          End Select
          
          
            Print #1, "                <Locator Id=""" & i & """>"
            Print #1, "                    <LomId Value=""0"" />"
            Print #1, "                    <Time Value=""" & (i * 8) * 4 & """ />"
            Print #1, "                    <Name Value=""" & LocatorName & """ />"
            Print #1, "                    <Annotation Value="""" />"
            Print #1, "                    <IsSongStart Value=""false"" />"
            Print #1, "                </Locator>"
          Next i
    
    Print #1, "           </Locators>"



For i = 598 To 648



If i = 613 Then
    For y = 1 To 151
    Print #1, Worksheets("xmlpasta").Cells(y, 20).Value
    Next y
Else
Print #1, Worksheets("xmlpasta").Cells(i, 1).Value
End If


Next i


Print #1, "</LiveSet>"
Print #1, "</Ableton>"



Close #1

xl2xml.zipQuik

Worksheets("Arrangement").Range("AC22:AI22").Value = ""

End Sub

Sub zipQuik()

Dim pathInput As String
Dim pathOutput As String

pathInput = ActiveWorkbook.Path & "\testerB.xml"
pathOutput = ActiveWorkbook.Path & "\zipped.als"

Call Shell(Chr(34) & "C:\Program Files\7-Zip\7z.exe" & Chr(34) _
& " a -tgzip " & Chr(34) & pathOutput & Chr(34) & " " & Chr(34) _
& pathInput & Chr(34), vbNormalNoFocus)

'7z a -tgzip "C:\Users\Dylan\Desktop\zipped.als" "C:\Users\Dylan\Desktop\testerB.xml"


End Sub


Sub testpath()

Dim p1 As String
Dim p2 As String

p1 = ActiveWorkbook.Path
'Debug.Print p1
p2 = CreateObject("WScript.Shell").specialfolders("Desktop")
'Debug.Print CreateObject("WScript.Shell").specialfolders("Desktop")

End Sub

Sub testrangs()

'Debug.Print IsEmpty(Worksheets("PatternSaver").Range("D1").offset(((24 * 1) - 24) + (1 * 3), 0))


'Debug.Print Application.WorksheetFunction.CountA(Worksheets("Notes").Range("A1:G1"))

'Debug.Print Application.WorksheetFunction.CountA(Worksheets("Notes").Range(Cells(1, 1), Cells(1, 19)))
'Debug.Print Application.WorksheetFunction.CountA(Worksheets("PatternSaver").Range(Cells(1, 4), Cells(1, 19)))

End Sub



Sub decimalll()

Dim deci As Integer
deci = 1
Dim deciV As Variant
deciV = 1

MsgBox deci / 2 & "      " & Replace(deci / 2, ",", ".") & "      " & deciV / 2



End Sub

