<div align="center">

## ComboBox on Toolbar


</div>

### Description

This subroutine shows how to Really put a ComboBox (or any control with a hWnd) onto a ToolBar

(or any other control/window with a hWnd).
 
### More Info
 
There are no input paramaters

Add a ComboBox, CheckBox and Toolbar to Form1.

Keep the default names of the above mentioned controls.

Don't worry about control placement or size.

Click the form after you run the app.

There are no function returns

No side effects, completely safe.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Andrew Mitchell Barfield](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/andrew-mitchell-barfield.md)
**Level**          |Unknown
**User Rating**    |4.2 (164 globes from 39 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/andrew-mitchell-barfield-combobox-on-toolbar__1-820/archive/master.zip)





### Source Code

```
Option Explicit
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Sub Form_Load()
'Set Toolbar1 as Combo1's parent, then move Combo1 where we want it.
   SetParent Combo1.hwnd, Toolbar1.hwnd
   MoveWindow Combo1.hwnd, 100, 1, 50, 50, True 'Note: units are pixels
'Set Toolbar1 as Check1's parent, then move Check1 where we want it.
   SetParent Check1.hwnd, Toolbar1.hwnd
   MoveWindow Check1.hwnd, 175, 5, 150, 15, True
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Demonstrate that Combo1 and Check1 are really "on" Toolbar1 by moving Toolbar1 when
'the form is clicked.
   Toolbar1.Move X, Y
End Sub
```

