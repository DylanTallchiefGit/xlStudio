Attribute VB_Name = "openLink"
''code from:  https://stackoverflow.com/questions/3166265/open-an-html-page-in-default-browser-with-vba
Option Explicit

  #If Win64 Then
Private Declare PtrSafe Function ShellExecute _
  Lib "shell32.dll" Alias "ShellExecuteA" ( _
  ByVal hwnd As Long, _
  ByVal Operation As String, _
  ByVal Filename As String, _
  Optional ByVal Parameters As String, _
  Optional ByVal Directory As String, _
  Optional ByVal WindowStyle As Long = vbMinimizedFocus _
  ) As Long
  #Else
 Private Declare Function ShellExecute _
  Lib "shell32.dll" Alias "ShellExecuteA" ( _
  ByVal hWnd As Long, _
  ByVal Operation As String, _
  ByVal Filename As String, _
  Optional ByVal Parameters As String, _
  Optional ByVal Directory As String, _
  Optional ByVal WindowStyle As Long = vbMinimizedFocus _
  ) As Long
  #End If

Public Sub OpenUrl()

    Dim lSuccess As Long
    lSuccess = ShellExecute(0, "Open", "http://www.youtube.com/channel/UCIu2Fj4x_VMn2dgSB1bFyQA?sub_confirmation=1")

End Sub

