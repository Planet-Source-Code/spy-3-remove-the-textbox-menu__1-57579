<div align="center">

## Remove The Textbox Menu


</div>

### Description

This article will show you how you can remove that annoying textbox menu and replace it with your own or just have it gone.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[SPY\-3](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/spy-3.md)
**Level**          |Intermediate
**User Rating**    |3.5 (14 globes from 4 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/spy-3-remove-the-textbox-menu__1-57579/archive/master.zip)





### Source Code

<b>Non-API<br>
To do this simply add this code to the MouseDown part of the textbox<br>
This way will work with probably all versions of VB<br>
<font color=blue>If</font><font color=black> Button = 2</font> <font color=blue>Then</font><br>
<font color=green>YourTextboxName</font><font color=black>.Enabled =</font><font color=blue> False</font><br>
<font color=green>YourTextboxName</font><font color=black>.Enabled =</font><font color=blue> True</font><br>
<font color=green>YourTextboxName</font><font color=black>.SetFocus</font><br>
PopupMenu <font color=green>YourMenuName</font><br>
<font color=blue>End If</font><br>
Replace all the <font color=green>Green</font> text with what your control names are.<br>
Hope this helped.<br><br>
This way will only work in VB 5.0 and VB 6.0 as far as i know<br>
API<br>
<font color="blue">Option Explicit</font><br>
<font color="green">'Parts of this were orginally made by<br>
' Written by Matt Hart<br>
'Altered by SPY-3<br>
'This was originally written for a webbrowser see<br>
'http://blackbeltvb.com/index.htm?free/webbmenu.htm<br><br>
</font>
<font color="blue">
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long<br>
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)<br>
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long<br>
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long<br>
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long<br><br>
Public Const GWL_WNDPROC = (-4)<br><br>
Public Const GW_HWNDNEXT = 2<br>
Public Const GW_CHILD = 5<br><br>
Public Const WM_MOUSEACTIVATE = &H21<br>
Public Const WM_CONTEXTMENU = &H7B<br>
Public Const WM_RBUTTONDOWN = &H204<br><br>
Public origWndProc As Long<br><br>
Public Function AppWndProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long<br>
Select Case Msg<br>
Case WM_MOUSEACTIVATE<br>
Dim C As Integer<br>
Call CopyMemory(C, ByVal VarPtr(lParam) + 2, 2)<br>
If C = WM_RBUTTONDOWN Then<br>
<font color="green">YourForm</font>.PopupMenu <font color="green">YourForm</font>.<font color="green">YourMenu</font><br>
SendKeys "{ESC}"<br>
End If<br>
Case WM_CONTEXTMENU<br>
<font color="green">YourForm</font>.PopupMenu <font color="green">YourForm</font>.<font color="green">YourMenu</font><br>
SendKeys "{ESC}"<br>
End Select<br>
AppWndProc = CallWindowProc(origWndProc, hwnd, Msg, wParam, lParam)<br>
End Function<br></font>Then under Form_Load() put this<br><font color="blue">origWndProc = SetWindowLong(<font color="green">YourTextBox</font>.hwnd, GWL_WNDPROC, AddressOf AppWndProc)</font>
<br><br>
<font color="black">http://Tiamat-Studios.vze.com</font>

