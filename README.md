<div align="center">

## Cool screen wipes


</div>

### Description

You can achieve some cool form wipes with judicious use of the Move method. For example, to draw a curtain from right to left use this routine. It is also possible to wipe a form from bottom to top, and from both sides to the middle, using similar routines
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[VB Pro](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vb-pro.md)
**Level**          |Unknown
**User Rating**    |5.9 (630 globes from 107 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vb-pro-cool-screen-wipes__1-126/archive/master.zip)





### Source Code

```
Sub WipeRight (Lt%, Tp%, frm As Form)
Dim s, Wx, Hx, i
s = 90 'number of steps to use in the wipe
Wx = frm.Width / s 'size of vertical steps
Hx = frm.Height / s 'size of horizontal steps
' top and left are static
' while the width gradually shrinks
For i = 1 To s - 1
frm.Move Lt, Tp, frm.Width - Wx
Next
End Sub
Call the routine from a command button by using this code:
L = Me.Left
T = Me.Top
WipeRight L, T, Me
```

