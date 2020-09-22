<div align="center">

## Drag Away


</div>

### Description

To Drag A Form Without Using The Annoying Blue Bar
 
### More Info
 
No need To Know Anything Sept This Will Save You Alot Of Time

No Side Effects


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Compudyne 2o0\`1](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/compudyne-2o0-1.md)
**Level**          |Beginner
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/compudyne-2o0-1-drag-away__1-29261/archive/master.zip)

### API Declarations

```
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "User32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
```


### Source Code

```
'Make A New Form
'Copy And Past this In the Main Code Area For
'Form
Private Sub Form_Load()
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim lngReturnValue As Long
 If Button = 1 Then
  Call ReleaseCapture
  lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
 End If
End Sub
'And Werk Like You would Always Werk Then
'Press
'Play And Dag It From The
'Main body Of the Form
'Instead Of the Blue Bar
```

