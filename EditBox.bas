Attribute VB_Name = "EditBox"
Private Const WM_SETFONT As Long = &H30
Private Const WM_CUT As Long = &H300
Private Const WM_COPY As Long = &H301
Private Const WM_PASTE As Long = &H302
Private Const WM_UNDO As Long = &H304
Private Const EM_SETSEL As Long = &HB1
Private Const EM_SETMARGINS = &HD3
Private Const EC_LEFTMARGIN = &H1

Enum EditMenu
    M_UNDO = 0
    M_CUT
    M_COPY
    M_PASTE
    M_DEL
    M_SELECT_ALL
End Enum

Function SetMargin(hwnd As Long, Optional MarginSize As Long = 8)
    ' This just sets a margin on to the edit box
    SendMessageBynum hwnd, EM_SETMARGINS, EC_LEFTMARGIN, MarginSize
End Function

Function SetEditDefaultFont(hwnd As Long, hFont As Long) As Long
    ' function used to setup the default font for the editbox
    SendMessage hwnd, WM_SETFONT, hFont, 1
End Function

Function EditBoxSetFoucs(hwnd As Long)
    ' This function is used to set focus on to the edit box
    SetFocus hwnd
End Function

Function GetEditText(hwnd As Long) As String
Dim iLen As Long, lpRet As Long
Dim StrB As String
    ' Get the text from the edit box
    iLen = GetWindowTextLength(hwnd) + 1
    If iLen = 0 Then GetEditText = "": Exit Function
    StrB = Space(iLen) 'Create a buffer to hold the text
    lpRet = GetWindowText(hwnd, StrB, iLen)
    GetEditText = Left(StrB, InStr(1, StrB, Chr(0)) - 1)
    
    StrB = ""
    iLen = 0
    
End Function

Function PutEditText(hwnd As Long, sText As String) As Long
    ' Sets text on to the edit box
    If hwnd = 0 Then PutEditText = 0: Exit Function ' There not much point carrying on if no hangle is found
    PutEditText = SetWindowText(hwnd, sText) ' Place the text on to the edit box
    EditBoxSetFoucs hwnd
End Function

Function EditBoxLen(hwnd As Long) As Long
    ' Returns the length of the text box
    EditBoxLen = GetWindowTextLength(hwnd)
End Function

Function EditBoxEdit(hwnd As Long, mOp As EditMenu)
    Select Case mOp
        Case M_UNDO ' undo edit command does not seem to work tho not sure why
            SendMessage hwnd, WM_UNDO, 0, 0
            Exit Function
        Case M_CUT ' Cut edit command
            SendMessage hwnd, WM_CUT, 0, 0
            Exit Function
        Case M_COPY ' Copy edit command
            SendMessage hwnd, WM_COPY, 0, 0
            Exit Function
        Case M_PASTE ' paste edit command
            SendMessage hwnd, WM_PASTE, 0, 0
            Exit Function
        Case M_SELECT_ALL ' select all edit command
            SendMessage hwnd, EM_SETSEL, 0, 0
            Exit Function
    End Select
    
End Function
