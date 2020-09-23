<div align="center">

## Create links from labels\!


</div>

### Description

Turns labels into links which when clicked, load the contents of the caption as the URL into a webbrowser.
 
### More Info
 
Private Sub Form_Load()

MakeLink Label1, Startup

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

MakeLink Label1, FormMove

End Sub

Private Sub Label1_Click()

MakeLink Label1, Click, Me

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

MakeLink Label1, LinkMove

End Sub


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Paul Spiteri](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/paul-spiteri.md)
**Level**          |Beginner
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/paul-spiteri-create-links-from-labels__1-9316/archive/master.zip)

### API Declarations

```
Public Enum OpType
  Startup = 1
  Click = 2
  FormMove = 3
  LinkMove = 4
End Enum
Dim Clicked As Boolean
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
```


### Source Code

```
Public Sub MakeLink(LabelName As Label, Operation As OpType, Optional FormName As Form)
  Dim Openpage As Integer
  Select Case Operation
  Case LinkMove
    LabelName.ForeColor = 255
    LabelName.FontUnderline = True
  Case Click
    Openpage = ShellExecute(FormName.hwnd, "Open", LabelName.Caption, "", App.Path, 1)
    LabelName.ForeColor = 8388736
    Clicked = True
  Case FormMove
    LabelName.FontUnderline = False
    If Not Clicked Then
      LabelName.ForeColor = 16711680
    Else
      LabelName.ForeColor = 8388736
    End If
  Case Startup
    LabelName.ForeColor = 16711680
  End Select
End Sub
```

