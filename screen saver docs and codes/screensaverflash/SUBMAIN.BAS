Attribute VB_Name = "submain"
'#########################################################
'Author: Somdutt Ganguly
'Email: gangulysomdutt@yahoo.com
'Current status: Master of computer application - IInd
'                semister, North Gujarat University
'                 from SV institute of computer
'                studies, KADI, Gujarat, INDIA
'Resident: no 6, chandrodaya apt, bhaikaka nagar, thaltej, ahmedabad, gujarat, india. 380059
'Date: FEB - 2003
'Code: Screen Saver
'###########################################################
Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOZORDER = &H4
Public Const SWP_SHOWWINDOW = &H40

Public Const HWND_TOP = 0

Public Const WS_CHILD = &H40000000
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_STYLE = (-16)

Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Sub Main()
'sub main module which initiates first
'see while building a screen saver you have to specificly
'keep the below format...always!!
'REMEMBER...
'app.previnstance ==> checks to see if the screen saver's previous instance..
'command keyword ==> returns the argument for the windows command
'                    windows forwards /s argument if it wants to start a screensaver
'                    windows forwards /c argument if the setup button is pressed in the screen saver's properties
'                    windows forwards /p (numbers) if it wants to start the preview window
'got it!!
'########################################################
Dim preview_hwnd As Long
Dim preview_rect As RECT
Dim window_style As Long
Dim args As String


'If App.PrevInstance Then End

If InStr(Command, "/s") > 0 Then
'MsgBox "showing screensaver", vbInformation
    Call CursorOff
    frmmain.Show
  
ElseIf InStr(Command, "/c") > 0 Then
'MsgBox "showing configuration", vbInformation
    MsgBox "There is no setup", _
    vbInformation, " Screen Saver"
'preview logic
ElseIf InStr(Command, "/p") > 0 Then
args = UCase$(Trim$(Command$))
' Get the preview area hWnd.
            preview_hwnd = GetHwndno(args)
            ' Get the dimensions of the preview area.
            GetClientRect preview_hwnd, preview_rect
            Load frmmain
            ' Get the current window style.
            window_style = GetWindowLong(frmmain.hwnd, GWL_STYLE)
            ' Add WS_CHILD to make this a child window.
            window_style = (window_style Or WS_CHILD)
            ' Set the window's new style.
            SetWindowLong frmmain.hwnd, _
                GWL_STYLE, window_style
            ' Set the window's parent so it appears
            ' inside the preview area.
            SetParent frmmain.hwnd, preview_hwnd
            ' Save the preview area's hWnd in
            ' the form's window structure.
            SetWindowLong frmmain.hwnd, _
                GWL_HWNDPARENT, preview_hwnd

            ' Show the preview.
            SetWindowPos frmmain.hwnd, _
                HWND_TOP, 0&, 0&, _
                preview_rect.Right, _
                preview_rect.Bottom, _
                SWP_NOZORDER Or SWP_NOACTIVATE Or _
                    SWP_SHOWWINDOW
Else
Call CursorOff
 frmmain.Show
End If
'#######################################################

End Sub


Public Sub CursorOff()
ShowCursor False
End Sub

Public Sub CursorOn()
ShowCursor True
End Sub


' Get the hWnd for the preview window from the
' command line arguments.
Private Function GetHwndno(ByVal args As String) As Long
Dim argslen As Integer
Dim i As Integer
Dim ch As String

    ' Take the rightmost numeric characters.
    args = Trim$(args)
    argslen = Len(args)
    For i = argslen To 1 Step -1
        ch = Mid$(args, i, 1)
        If ch < "0" Or ch > "9" Then Exit For
    Next i

    GetHwndFromCommand = CLng(Mid$(args, i + 1))
End Function


