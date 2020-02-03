VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Private Sub Patt1_Click()

Dim lastletter As Variant
Dim lastInteger As Variant

lastletter = Right("array1", 1)
lastInteger = ConvertString(lastletter)

'Dim topLeft As Range
'Set topLeft = Cells(1, 2)
'
'Dim bottomRight As Range
'Set bottomRight = Cells(12, 17)

saveArray

Dim topLeft As Variant
Dim bottomRight As Variant

topLeft = lastInteger
bottomRight = lastInteger * 24



Dim pattern1 As Variant
pattern1 = Worksheets("PatternSaver").Range("B" & topLeft & ":Q" & bottomRight).Value
'pattern1 = Worksheets("PatternSaver").Range(topLeft, bottomRight).Value
Worksheets("Drum Machine").Range("H31:W54").Value = pattern1
Worksheets("PatternSaver").Range("A4").Value = lastInteger  ''stores in pattern saver for next saving

ActiveSheet.Patt1.BackColor = RGB(220, 105, 0)


End Sub


Private Sub Worksheet_Change(ByVal Target As Range)

    If Target.Address = "$C$23" And ActiveSheet.Name = "Drum Machine" Then
        Worksheets("Arrangement").Range("G22").offset(0, 0).Value = Worksheets("Drum Machine").Range("C23").offset(0, 0).Value
        Worksheets("Piano Roll").Range("E2").offset(0, 0).Value = Worksheets("Drum Machine").Range("C23").offset(0, 0).Value
     
    End If
    
    If Target.Address = "$C$24" And ActiveSheet.Name = "Drum Machine" Then
        Worksheets("Arrangement").Range("G22").offset(1, 0).Value = Worksheets("Drum Machine").Range("C23").offset(1, 0).Value
        Worksheets("Piano Roll").Range("E2").offset(1, 0).Value = Worksheets("Drum Machine").Range("C23").offset(1, 0).Value
     
    End If
    
    If Target.Address = "$C$26" And ActiveSheet.Name = "Drum Machine" Then
        Worksheets("Arrangement").Range("G22").offset(3, 0).Value = Worksheets("Drum Machine").Range("C23").offset(3, 0).Value
        Worksheets("Piano Roll").Range("E2").offset(3, 0).Value = Worksheets("Drum Machine").Range("C23").offset(3, 0).Value
     
    End If

End Sub


Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    If Target.Address = "$C$25" Then  'Worksheets("Piano Roll").Range("AT5").Address Then
      'Range("B5").Value = 10
      If Worksheets("Drum Machine").Range("C25").Value = "Off" Then
      Worksheets("Drum Machine").Range("C25").Value = "On"
      Else
      Worksheets("Drum Machine").Range("C25").Value = "Off"
      End If

    End If
    
    
    
End Sub


