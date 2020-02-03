VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
 
 
 
If Target.Address = "$E$2" And ActiveSheet.Name = "Piano Roll" Then
        Worksheets("Drum Machine").Range("C23").offset(0, 0).Value = Worksheets("Piano Roll").Range("E2").offset(0, 0).Value
        Worksheets("Arrangement").Range("G22").offset(0, 0).Value = Worksheets("Piano Roll").Range("E2").offset(0, 0).Value
 
End If

If Target.Address = "$E$3" And ActiveSheet.Name = "Piano Roll" Then
        Worksheets("Drum Machine").Range("C23").offset(1, 0).Value = Worksheets("Piano Roll").Range("E2").offset(1, 0).Value
        Worksheets("Arrangement").Range("G22").offset(1, 0).Value = Worksheets("Piano Roll").Range("E2").offset(1, 0).Value
 
End If

'swing
If Target.Address = "$E$5" And ActiveSheet.Name = "Piano Roll" Then
        Worksheets("Drum Machine").Range("C23").offset(3, 0).Value = Worksheets("Piano Roll").Range("E2").offset(3, 0).Value
        Worksheets("Arrangement").Range("G22").offset(3, 0).Value = Worksheets("Piano Roll").Range("E2").offset(3, 0).Value
 
End If

'legato
If Target.Address = "$E$4" And ActiveSheet.Name = "Piano Roll" Then
        Worksheets("Arrangement").Range("G22").offset(2, 0).Value = Worksheets("Piano Roll").Range("E4").Value
 
End If




If Target.Address = "$BB$2" And ActiveSheet.Name = "Piano Roll" Then

If Worksheets("Piano Roll").Range("BB2").Value <> 5709 Then
pianroll.openPianoPattern

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

Else
MsgBox "This is the only one you can't use :) "
End If
 
End If
 
 
 
End Sub



Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    If Target.Address = "$BB$3:$BC$3" Then  'Worksheets("Piano Roll").Range("AT5").Address Then
      'Range("B5").Value = 10
      If Worksheets("Piano Roll").Range("BB3").Value = "Off" Then
      Worksheets("Piano Roll").Range("BB3").Value = "On"
      Else
      Worksheets("Piano Roll").Range("BB3").Value = "Off"
      End If

    End If
    
    
    If Target.Address = "$E$4:$F$4" Then  'Worksheets("Piano Roll").Range("AT5").Address Then
      'Range("B5").Value = 10
      If Worksheets("Piano Roll").Range("E4").Value = "Off" Then
      Worksheets("Piano Roll").Range("E4").Value = "On"
      Else
      Worksheets("Piano Roll").Range("E4").Value = "Off"
      End If

    End If
    
End Sub







