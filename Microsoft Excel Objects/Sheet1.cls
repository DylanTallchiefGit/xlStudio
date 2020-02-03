VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Private Sub Worksheet_Change(ByVal Target As Range)
 
 
 
If Target.Address = "$G$22" And ActiveSheet.Name = "Arrangement" Then
        Worksheets("Drum Machine").Range("C23").offset(0, 0).Value = Worksheets("Arrangement").Range("G22").offset(0, 0).Value
        Worksheets("Piano Roll").Range("E2").offset(0, 0).Value = Worksheets("Arrangement").Range("G22").offset(0, 0).Value
 
End If

If Target.Address = "$G$23" And ActiveSheet.Name = "Arrangement" Then
        Worksheets("Drum Machine").Range("C23").offset(1, 0).Value = Worksheets("Arrangement").Range("G22").offset(1, 0).Value
        Worksheets("Piano Roll").Range("E2").offset(1, 0).Value = Worksheets("Arrangement").Range("G22").offset(1, 0).Value
 
End If

If Target.Address = "$G$24" And ActiveSheet.Name = "Arrangement" Then
        Worksheets("Piano Roll").Range("E2").offset(2, 0).Value = Worksheets("Arrangement").Range("G22").offset(2, 0).Value
 
End If


If Target.Address = "$G$25" And ActiveSheet.Name = "Arrangement" Then
        Worksheets("Drum Machine").Range("C23").offset(3, 0).Value = Worksheets("Arrangement").Range("G22").offset(3, 0).Value
        Worksheets("Piano Roll").Range("E2").offset(3, 0).Value = Worksheets("Arrangement").Range("G22").offset(3, 0).Value
 
End If
 
 
End Sub



Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    
    
    If Target.Address = "$G$24" Then  'Worksheets("Piano Roll").Range("AT5").Address Then
      'Range("B5").Value = 10
      If Worksheets("Arrangement").Range("G24").Value = "Off" Then
      Worksheets("Arrangement").Range("G24").Value = "On"
      Else
      Worksheets("Arrangement").Range("G24").Value = "Off"
      End If

    End If
    
    
    If Target.Address = "$J$14:$Y$14" Then
      'Range("B5").Value = 10
      If Worksheets("Arrangement").Rows(16).RowHeight = 0 Then
      Worksheets("Arrangement").Rows(16).RowHeight = 260
      Else
      Worksheets("Arrangement").Rows(16).RowHeight = 0
      End If

    End If
    
End Sub
