Attribute VB_Name = "Module1"
''''''''''''''''''''''''''''''''''''
''Once again, thanks to AllAPI.net''
''''''''''''''''''''''''''''''''''''
Private OriginalWindowProc As Long
Public Const MF_STRING = &H0&
Public Const MF_ENABLED = &H0&
Public Const IDM_MYMENUITEM = 2003
Public Const WM_SYSCOMMAND = &H112
Public Const GWL_WNDPROC = (-4)

Public Declare Function GetSystemMenu Lib "user32" _
  (ByVal hWnd As Long, ByVal bRevert As Long) As Long

Public Declare Function AppendMenu Lib "user32" _
  Alias "AppendMenuA" (ByVal hMenu As Long, _
  ByVal wflags As Long, ByVal wIDNewItem As Long, _
  ByVal lpNewItem As String) As Long

Public Declare Function SetWindowLong Lib "user32" _
  Alias "SetWindowLongA" (ByVal hWnd As Long, _
  ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function CallWindowProc Lib "user32" _
  Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
  ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, _
  ByVal lParam As Long) As Long


Public Sub AddToSystemMenu(ByVal hWnd As Long)

Dim hSystemMenu As Long

' Get the system menu's handle.
hSystemMenu = GetSystemMenu(hWnd, False)

' Append a custom command to the menu.
AppendMenu hSystemMenu, MF_STRING + MF_ENABLED, _
IDM_MYMENUITEM, "My Menu Item"

' Tell Windows to call MyMenuProc when a system
' menu command is selected.
OriginalWindowProc = SetWindowLong(hWnd, GWL_WNDPROC, _
AddressOf MyMenuProc)

End Sub


Public Function MyMenuProc(ByVal hWnd As Long, ByVal msg As Long, _
  ByVal wParam As Long, ByVal lParam As Long) As Long

' If the custom menu item was selected display a message.
If msg = WM_SYSCOMMAND And wParam = IDM_MYMENUITEM Then
Form1.Label1.Caption = "New menu item clicked!"
Exit Function
End If

' Otherwise pass the command on for normal processing.
MyMenuProc = CallWindowProc(OriginalWindowProc, hWnd, msg, _
  wParam, lParam)

End Function
 

