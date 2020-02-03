Attribute VB_Name = "PlayMIDI"
' a lot of the midi functions were ripped from Excel MIDI:  https://sourceforge.net/projects/excel-midi/
' check it out for a more in-depth midi sequencer
' a lot of these variables dont do anything anymore

Option Explicit

Private Const MAXPNAMELEN               As Integer = 32
Private Const MMSYSERR_BASE             As Integer = 0
Private Const MMSYSERR_BADDEVICEID      As Integer = (MMSYSERR_BASE + 2)
Private Const MMSYSERR_INVALPARAM       As Integer = (MMSYSERR_BASE + 11)
Private Const MMSYSERR_NODRIVER         As Integer = (MMSYSERR_BASE + 6)
Private Const MMSYSERR_NOMEM            As Integer = (MMSYSERR_BASE + 7)
Private Const MMSYSERR_INVALHANDLE      As Integer = (MMSYSERR_BASE + 5)
Private Const MIDIERR_BASE              As Integer = 64
Private Const MIDIERR_STILLPLAYING      As Integer = (MIDIERR_BASE + 1)
Private Const MIDIERR_NOTREADY          As Integer = (MIDIERR_BASE + 3)
Private Const MIDIERR_BADOPENMODE       As Integer = (MIDIERR_BASE + 6)

Private Type MIDIOUTCAPS
   wMid             As Integer
   wPid             As Integer
   wTechnology      As Integer
   wVoices          As Integer
   wMessages        As Integer
   wChannelMask     As Integer
   vDriverVersion   As Long
   dwSupport        As Long
   szPname          As String * MAXPNAMELEN
End Type

#If Win64 Then
    Private Declare PtrSafe Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As LongPtr) As Long
    Private Declare PtrSafe Function midiOutOpen Lib "winmm.dll" (lphMidiOut As LongPtr, ByVal uDeviceID As LongPtr, ByVal dwCallback As LongPtr, ByVal dwInstance As LongPtr, ByVal dwFlags As LongPtr) As Long
    Private Declare PtrSafe Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As LongPtr, ByVal dwMsg As LongPtr) As Long
    Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long
#Else
    Private Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
    Private Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
    Private Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long
    Declare Function timeGetTime Lib "winmm.dll" () As Long
#End If

#If Win64 Then
    Private mlngCurDevice      As LongPtr
    Private mlngHmidi          As LongPtr
    Private mlngRc             As LongPtr
    Private mlngMidiMsg        As LongPtr
    Private mlngMidiMsgLast    As LongPtr
#Else
    Private mlngCurDevice      As Long
    Private mlngHmidi          As Long
    Private mlngRc             As Long
    Private mlngMidiMsg        As Long
    Private mlngMidiMsgLast    As LongPtr
#End If


Private mlngTickCount          As Long
Private mintChannel            As Integer
Private mintVelocity           As Integer
Private mintController         As Integer
Private mintMessageLength      As Long
Private mintMessageNumber      As Integer
Private mintPressure           As Integer
Private mintValue              As Integer
Private mintProgram            As Integer
Private mintLSB                As Integer
Private mintMSB                As Integer
Private mstrDeviceName         As String
Private mblnIsDeviceOpen       As Boolean
Private i As Integer

Private Const INT_DEFAULT_CHANNEL            As Integer = 0
Private Const INT_DEFAULT_VELOCITY           As Integer = 160
Private Const INT_DEFAULT_MESSAGE_LENGTH     As Integer = 1000
Private Const INT_DEFAULT_CUR_DEVICE         As Integer = 0     'Define Device HERE

Private Sub Class_Initialize()
    mintChannel = INT_DEFAULT_CHANNEL
    mlngCurDevice = INT_DEFAULT_CUR_DEVICE
    mintVelocity = INT_DEFAULT_VELOCITY
    mintMessageLength = INT_DEFAULT_MESSAGE_LENGTH
    mblnIsDeviceOpen = False
    'Call OpenDevice
End Sub

'Private Sub Class_Terminate()
'    Call CloseDevice
'End Sub

'Public Sub InitiateDevice(ByVal Device As Integer)
'
'    mintChannel = INT_DEFAULT_CHANNEL
'    mlngCurDevice = Device   'if I use Device 4 this will send 3
'    mintVelocity = INT_DEFAULT_VELOCITY
'    mintMessageLength = INT_DEFAULT_MESSAGE_LENGTH
'    Call OpenDevice
'
'End Sub

Sub midiNote(program As Integer, dVel As Integer, dPitch As Integer, dChan As Integer)


'Dim sWatch As New Stopwatch

'Dim Message          As New csMessage
'Dim Devices(15)        As New csMidi

'Message.MessageChannel = 0   ''channel
'Message.MessageNumber = 60  ''pitch
'Message.MessageProgram = 0   ''program

''inititate device stuff ' Devices(0).InitiateDevice 3
    mintChannel = INT_DEFAULT_CHANNEL
    mlngCurDevice = Worksheets("Drum Machine").Range("c24").Value - 1 ' 3   'if I use Device 4 this will send 3
    'mintVelocity = 120 'INT_DEFAULT_VELOCITY
    mintMessageLength = INT_DEFAULT_MESSAGE_LENGTH
   ' mintMessageNumber = 60 'default pitch
    
    

    
'    Call OpenDevice
    If Not mblnIsDeviceOpen Then
        mlngRc = midiOutClose(mlngHmidi)
        mlngRc = midiOutOpen(mlngHmidi, mlngCurDevice, 0, 0, 0)

        If (mlngRc <> 0) Then
            MsgBox "Couldn't open midi out " & mlngRc & ". Try another Device or Restart Excel. Make sure you have MIDI devices connected and running."
            stopOkay
            mblnIsDeviceOpen = False
        End If
        mblnIsDeviceOpen = True
    End If
    


   
    '' set program stuff 'Devices(0).SetProgram Message
        mintChannel = dChan ' Message.MessageChannel
        mintProgram = program ' 0 ' Message.MessageProgram
'    Call ProgramChange
'    If mblnIsDeviceOpen = True Then
'        mlngMidiMsg = (mintProgram * 256) + &HC0 + mintChannel + (0 * 256) * 256
'        midiOutShortMsg mlngHmidi, mlngMidiMsg
'    End If

'    Debug.Print mlngCurDevice
    'play note stuff 'Devices(0).PlayNote Message
      '  mintChannel = 0 ' Message.MessageChannel
'    mintMessageNumber = 60 ' this is C5, the default sample pitch in FL  (ableton is C3 so that will need to change)
     mintMessageNumber = dPitch
'    mintVelocity = 120 ' Message.MessageVelocity
     mintVelocity = dVel
     
     'if no pitch or vel set
    If mintVelocity = 0 Then
     mintVelocity = 120
     End If
     
         If mintMessageNumber = 0 Then
     mintMessageNumber = 60
     End If
     
         If mintVelocity > 127 Then
     mintVelocity = 120
     End If
     
         If mintMessageNumber > 127 Then
     mintMessageNumber = 127
     End If


         If mintProgram > 127 Then
    mintProgram = 127
    End If


      'this message is for the channel + program, send seperately
        mlngMidiMsg = (mintProgram * 256) + &HC0 + mintChannel + (0 * 256) * 256
        midiOutShortMsg mlngHmidi, mlngMidiMsg
  '   Call StartNote
          mlngMidiMsg = &H90 + (mintMessageNumber * &H100) + (mintVelocity * &H10000) + mintChannel   ''mlngMidiMsg is a long combination integer in hexadecimal that contains all the note information (pitch,vel, channel) into one unqiue number, see maths up top
          midiOutShortMsg mlngHmidi, mlngMidiMsg

'While sWatch.Elapsed < sWatch.Elapsed + 4
'
'Wend


'
' '   Call StopNote
'     mlngMidiMsg = &H80 + (mintMessageNumber * &H100) + mintChannel 'without veloctity so vel = 0
'    midiOutShortMsg mlngHmidi, mlngMidiMsg
'
'
''    Call CloseDevice
'    If mblnIsDeviceOpen Then
'        mlngRc = midiOutClose(mlngHmidi)
'        mblnIsDeviceOpen = False
'    End If


 'DoEvents ' allows u to still click on sheet when running
 
 
 

End Sub


Sub midiNoteAR(program As Integer, dVel As Integer, dPitch As Integer, dChan As Integer)


    mlngCurDevice = Worksheets("Drum Machine").Range("c24").Value - 1 ' 3   'if I use Device 4 this will send 3
    'mintVelocity = 120 'INT_DEFAULT_VELOCITY
    mintMessageLength = INT_DEFAULT_MESSAGE_LENGTH
   ' mintMessageNumber = 60 'default pitch
    
    

    
'    Call OpenDevice
    If Not mblnIsDeviceOpen Then
        mlngRc = midiOutClose(mlngHmidi)
        mlngRc = midiOutOpen(mlngHmidi, mlngCurDevice, 0, 0, 0)

        If (mlngRc <> 0) Then
            MsgBox "Couldn't open midi out " & mlngRc & ". Try another Device or Restart Excel. Make sure you have MIDI devices connected and running."
            stopOkay
            mblnIsDeviceOpen = False
        End If
        mblnIsDeviceOpen = True
    End If
    


   
    '' set program stuff 'Devices(0).SetProgram Message
        mintChannel = dChan ' Message.MessageChannel
        mintProgram = program ' 0 ' Message.MessageProgram


'    Debug.Print mlngCurDevice
    'play note stuff 'Devices(0).PlayNote Message
      '  mintChannel = 0 ' Message.MessageChannel
'    mintMessageNumber = 60 ' this is C5, the default sample pitch in FL  (ableton is C3 so that will need to change)
     mintMessageNumber = dPitch
'    mintVelocity = 120 ' Message.MessageVelocity
     mintVelocity = dVel
     
     
         If mintMessageNumber = 0 Then
     mintMessageNumber = 60
     End If
     
         If mintVelocity > 127 Then
     mintVelocity = 120
     End If
     
         If mintMessageNumber > 127 Then
     mintMessageNumber = 60
     End If


         If mintProgram > 127 Then
    mintProgram = 127
    End If


      'this message is for the channel + program, send seperately
        mlngMidiMsg = (mintProgram * 256) + &HC0 + mintChannel + (0 * 256) * 256
        midiOutShortMsg mlngHmidi, mlngMidiMsg
  '   Call StartNote
          mlngMidiMsg = &H90 + (mintMessageNumber * &H100) + (mintVelocity * &H10000) + mintChannel   ''mlngMidiMsg is a long combination integer in hexadecimal that contains all the note information (pitch,vel, channel) into one unqiue number, see maths up top
          midiOutShortMsg mlngHmidi, mlngMidiMsg


 
 
 

End Sub



Sub stopItMIDI()

If Not Worksheets("Drum Machine").Range("C24").Value = 1 Then


 '   Call StopNote
     mlngMidiMsg = &H80 + (mintMessageNumber * &H100) + mintChannel 'without veloctity so vel = 0
    midiOutShortMsg mlngHmidi, mlngMidiMsg


'    Call CloseDevice
    If mblnIsDeviceOpen Then
       mlngRc = midiOutClose(mlngHmidi)
       mblnIsDeviceOpen = False
    End If

End If

End Sub


Sub justStopNote(program As Integer, dVel As Integer, dPitch As Integer, dChan As Integer)


 '   Call StopNote
  '   midiOutShortMsg mlngHmidi, mlngMidiMsgLast
  '   mlngMidiMsgLast = &H80 + (dPitch * &H100) + dChan 'without veloctity so vel = 0
'
'works but slower (8ms)
'  For i = 1 To 127
'  mlngMidiMsgLast = &H80 + (i * &H100) + dChan 'without veloctity so vel = 0
'  midiOutShortMsg mlngHmidi, mlngMidiMsgLast
'  Next i

'less than 1ms



mlngMidiMsgLast = &HB0 + (123 * &H100) + (0 * &H10000) + dChan
midiOutShortMsg mlngHmidi, mlngMidiMsgLast
    

End Sub

Sub justStopNote2(program As Integer, dVel As Integer, dPitch As Integer, dChan As Integer)

If dPitch > 127 Then
dPitch = 127
End If


'
'works but slower (8ms)
'  For i = 1 To 127
'  mlngMidiMsgLast = &H80 + (i * &H100) + dChan 'without veloctity so vel = 0
'  midiOutShortMsg mlngHmidi, mlngMidiMsgLast
'  Next i

'less than 1ms



'        mlngMidiMsg = (program * 256) + &HC0 + dChan + (0 * 256) * 256
'        midiOutShortMsg mlngHmidi, mlngMidiMsg

 ' StopNote
     mlngMidiMsg = &H80 + (dPitch * &H100) + (0 * &H10000) + dChan 'without veloctity so vel = 0
    midiOutShortMsg mlngHmidi, mlngMidiMsg

'mlngMidiMsgLast = &HB0 + (123 * &H100) + (0 * &H10000) + dChan
'midiOutShortMsg mlngHmidi, mlngMidiMsgLast

    

End Sub

Sub stopItMIDIAgain()

'Debug.Print "stop midid"

 '   Call StopNote
     mlngMidiMsg = &H80 + (mintMessageNumber * &H100) + mintChannel 'without veloctity so vel = 0
    midiOutShortMsg mlngHmidi, mlngMidiMsg
    
    Dim y As Integer
    
    'For i = 0 To 15
    For y = 0 To 127
            mlngMidiMsg = (y * 256) + &HC0 + 1 + (0 * 256) * 256
        midiOutShortMsg mlngHmidi, mlngMidiMsg
    Next y
    'Next i

For i = 0 To 15
mlngMidiMsgLast = &HB0 + (123 * &H100) + (0 * &H10000) + i
midiOutShortMsg mlngHmidi, mlngMidiMsgLast
Next i

'    Call CloseDevice
    If mblnIsDeviceOpen Then
       mlngRc = midiOutClose(mlngHmidi)
       mblnIsDeviceOpen = False
    End If


End Sub

Sub startDevice(program As Integer, dVel As Integer, dPitch As Integer, dChan As Integer)

Dim sWatch As New Stopwatch

''inititate device stuff ' Devices(0).InitiateDevice 3
    mintChannel = INT_DEFAULT_CHANNEL
    mlngCurDevice = Worksheets("Drum Machine").Range("c24").Value - 1 ' 3   'if I use Device 4 this will send 3
    mintMessageLength = INT_DEFAULT_MESSAGE_LENGTH
    
        
'    Call OpenDevice
    If Not mblnIsDeviceOpen Then
        mlngRc = midiOutClose(mlngHmidi)
        mlngRc = midiOutOpen(mlngHmidi, mlngCurDevice, 0, 0, 0)

        If (mlngRc <> 0) Then
            MsgBox "Couldn't open midi out " & mlngRc & ". Try another Device or Restart Excel. Make sure you have MIDI devices connected and running."
            stopOkay
            mblnIsDeviceOpen = False
        End If
        mblnIsDeviceOpen = True
    End If

'        mintChannel = dChan ' Message.MessageChannel
'        mintProgram = Program ' 0 ' Message.MessageProgram
'
'     mintMessageNumber = dPitch
'     mintVelocity = dVel
'
'    If mintVelocity = 0 Then
'     mintVelocity = 120
'     End If
'
'         If mintMessageNumber = 0 Then
'     mintMessageNumber = 60
'     End If
'
'         If mintVelocity > 127 Then
'     mintVelocity = 120
'     End If
'
'         If mintMessageNumber > 127 Then
'     mintMessageNumber = 60
'     End If
'
'      'this message is for the channel + program, send seperately
'        mlngMidiMsg = (mintProgram * 256) + &HC0 + mintChannel + (0 * 256) * 256
'        midiOutShortMsg mlngHmidi, mlngMidiMsg
'  '   Call StartNote
'          mlngMidiMsg = &H90 + (mintMessageNumber * &H100) + (mintVelocity * &H10000) + mintChannel   ''mlngMidiMsg is a long combination integer in hexadecimal that contains all the note information (pitch,vel, channel) into one unqiue number, see maths up top
'          midiOutShortMsg mlngHmidi, mlngMidiMsg
'
'
'
' DoEvents ' allows u to still click on sheet when running
 
 
 

End Sub





