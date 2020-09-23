Attribute VB_Name = "modWin"

Function AddEditBox(hwnd As Long) As Long
    ' This function adds a new edit box to the window
    WndEditBox = CreateWindowEx(WS_EX_CLIENTEDGE, "EDIT", "", TextBoxStyle, 0, 0 _
    , 100, 100, hwnd, 0, GetModuleHandle(0), ByVal 0&)
    
    SetEditDefaultFont WndEditBox, GetStockObject(DEFAULT_GUI_FONT) ' Set the deafult font for the edit box
    
    AddEditBox = WndEditBox
    SetMargin WndEditBox, 6 ' set the margin size to six
    SetFocus WndEditBox ' Place focus on the edit box
    
End Function

Function Addmenu(hwnd As Long) As Long
Dim hMenu As Long, hSubMenu1 As Long, hSubMenu2 As Long
    ' This function adds a basic menu to our window
    hMenu = CreateMenu() ' Create a new menu for our window
    hSubMenu1 = CreatePopupMenu() ' Create the file popup menu
    hSubMenu2 = CreatePopupMenu()
    AppendMenu hMenu, MF_STRING Or MF_POPUP, hSubMenu1, "&File" ' Top Level File Menu
    AppendMenu hMenu, MF_STRING Or MF_POPUP, hSubMenu2, "&Edit" ' Top Level Edit Menu
    
    'File Menu
    AppendMenu hSubMenu1, MF_STRING, DM_MENU_NEW, "&New" ' Sub item
    AppendMenu hSubMenu1, MF_STRING, DM_MENU_OPEN, "&Open..." ' Sub item
    AppendMenu hSubMenu1, MF_STRING, DM_MENU_SAVE, "&Save" ' Sub item
    AppendMenu hSubMenu1, MF_SEPARATOR, -1, 0&
    AppendMenu hSubMenu1, MF_STRING, DM_MENU_ABOUT, "About" ' Sub Item
    AppendMenu hSubMenu1, MF_SEPARATOR, -1, 0&
    AppendMenu hSubMenu1, MF_STRING, DM_MENU_EXIT, "E&xit" ' Sub item
    'Edit menu
    AppendMenu hSubMenu2, MF_STRING, DM_MNU_UNDO, "&Undo" ' Sub item
    AppendMenu hSubMenu2, MF_SEPARATOR, -1, ByVal 0&
    AppendMenu hSubMenu2, MF_STRING, DM_MNU_CUT, "&Cut" ' Sub item
    AppendMenu hSubMenu2, MF_STRING, DM_MNU_COPY, "&Copy" ' Sub item
    AppendMenu hSubMenu2, MF_STRING, DM_MNU_PASTE, "&Paste" ' Sub item
    AppendMenu hSubMenu2, MF_SEPARATOR, -1, ByVal 0&
    AppendMenu hSubMenu2, MF_STRING, DM_MNU_SELECT_ALL, "Select &All" ' Sub item
    Addmenu = SetMenu(hwnd, hMenu) ' update the window with our menu
End Function

Function WinProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    ' This function is used to hangle all the messages that the window will recive
    Dim sTmp As String
    Dim ans As Integer
    Dim MyRect As RECT

    Select Case wMsg
        Case WM_CREATE
            ' Create a menu for our app
            If Addmenu(hwnd) <> 1 Then
                MessageBox WinHwnd, "Unable to set menu", "WM_CREATE", MB_OK Or MB_ICONEXCLAMATION
                SendMessage hwnd, WM_CLOSE, ByVal 0&, ByVal 0&
                Exit Function
            End If
            'Next we create a resiable editbox / textbox
            If AddEditBox(hwnd) = 0 Then
                MessageBox WinHwnd, "Unable to create edit box", "WM_CREATE", MB_OK Or MB_ICONEXCLAMATION
                Exit Function
            Else
               EditBoxSetFoucs WndEditBox
               Exit Function
            End If
            
        Case WM_COMMAND ' Message Commands
            Select Case wParam
                'START OF FILE MENU MESSAGES
                Case DM_MENU_NEW
                    If EditBoxLen(WndEditBox) <> 0 Then
                        ans = MessageBox(hwnd, "You have unsaved work " & vbCrLf & "Do you want to save your work now?", "Open", MB_ICONQUESTION Or MB_YESNO)
                        If ans = vbNo Then
                            PutEditText WndEditBox, "" 'Clear the edit box
                            Exit Function
                        Else
                            ClsDialog.DlgHwnd = hwnd
                            ClsDialog.Filter = "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "Diz Files (*.diz)" + Chr$(0) + "*.diz" + Chr$(0)
                            ClsDialog.Flags = 0
                            ClsDialog.FilterIndex = 1
                            ClsDialog.hInst = App.hInstance
                            ClsDialog.InitialDir = FixPath(App.Path)
                            ClsDialog.ShowSave
                            If ClsDialog.CancelError = False Then
                                Exit Function
                            Else
                                SaveToFile ClsDialog.FileName + AddFileExt(ClsDialog.FilterIndex), GetEditText(WndEditBox)
                                PutEditText WndEditBox, ""
                            End If
                        End If
                    End If
                    
                    Exit Function
                Case DM_MENU_OPEN ' Show the OpenFile Dialog
                    ' Fill in the dialog Struc information
                    ClsDialog.DlgHwnd = hwnd
                    ClsDialog.Filter = "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "Diz Files (*.diz)" + Chr$(0) + "*.diz" + Chr$(0)
                    ClsDialog.Flags = 0
                    ClsDialog.FilterIndex = 1
                    ClsDialog.hInst = App.hInstance
                    ClsDialog.InitialDir = FixPath(App.Path)

                    If EditBoxLen(WndEditBox) <> 0 Then
                        ans = MessageBox(hwnd, "You have unsaved work " & vbCrLf & "Do you want to save your work now?", "Open", MB_ICONQUESTION Or MB_YESNO)
                        If ans = vbYes Then
                            ' do save here
                            ClsDialog.DialogTitle = "Save As" ' Change the dialogs title
                            ClsDialog.ShowSave ' show save dialog
                            If Not ClsDialog.CancelError Then Exit Function
                            SaveToFile ClsDialog.FileName + AddFileExt(ClsDialog.FilterIndex), GetEditText(WndEditBox)
                        End If
                    End If
                    ' Open Dialog
                    ClsDialog.DialogTitle = "Open Text Files"
                    ClsDialog.ShowOpen
                    If ClsDialog.CancelError = False Then Exit Function ' Cancel was pressed
                    ' Update the editor with the text
                    PutEditText WndEditBox, OpenFile(ClsDialog.FileName)
                    Exit Function
                    
                Case DM_MENU_EXIT ' menu exit
                    ClsDialog.DlgHwnd = hwnd
                    ClsDialog.Filter = "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "Diz Files (*.diz)" + Chr$(0) + "*.diz" + Chr$(0)
                    ClsDialog.Flags = 0
                    ClsDialog.FilterIndex = 1
                    ClsDialog.hInst = App.hInstance
                    ClsDialog.InitialDir = FixPath(App.Path)
                    
                    If EditBoxLen(WndEditBox) <> 0 Then
                        ans = MessageBox(hwnd, "You have unsaved work " & vbCrLf & "Do you want to save your work now?", "Open", MB_ICONQUESTION Or MB_YESNO)
                        If ans = vbYes Then
                            ' do save here
                            ClsDialog.DialogTitle = "Save As" ' Change the dialogs title
                            ClsDialog.ShowSave ' show save dialog
                            If Not ClsDialog.CancelError Then Exit Function
                            SaveToFile ClsDialog.FileName + AddFileExt(ClsDialog.FilterIndex), GetEditText(WndEditBox)
                            SendMessage hwnd, WM_CLOSE, 99, ByVal 0&
                            Exit Function
                        Else
                            SendMessage hwnd, WM_CLOSE, 99, ByVal 0&
                            Exit Function
                        End If
                    Else
                        SendMessage hwnd, WM_CLOSE, 99, ByVal 0&
                    End If
                Case DM_MENU_ABOUT
                    ' Display an about box
                    ShellAbout hwnd, "DM All API Notepad", "Written By Dreamvb", 0
                    Exit Function
               'END OF FILE MENU MESSAGES
               
                'START OF EDIT MENU MESSAGES
                Case DM_MNU_UNDO ' Undo
                    EditBoxEdit WndEditBox, M_UNDO
                    Exit Function
                Case DM_MNU_CUT 'Cut
                    EditBoxEdit WndEditBox, M_CUT
                    Exit Function
                Case DM_MNU_COPY 'Copy text to clipboard
                    EditBoxEdit WndEditBox, M_COPY
                    Exit Function
                Case DM_MNU_PASTE ' Paste text from clipboard
                    EditBoxEdit WndEditBox, M_PASTE
                    Exit Function
                Case DM_MNU_SELECT_ALL ' select all text in the editbox
                    EditBoxEdit WndEditBox, M_SELECT_ALL
                    Exit Function
                'END OF EDIT MENU MESSAGES
            End Select
            
        Case WM_DROPFILES ' Look for any droped file on our window
            If GetFileDropCount(wParam) <> 1 Then
                MessageBox hwnd, "Only one file may be droped at a time.", "WM_DROPFILES", MB_ICONEXCLAMATION Or MB_OK
                DragFinish wParam ' Finish Drag
                Exit Function
            Else
                DropFileName = GetFileDrop(wParam) ' Get the droped filename
                If DropFileName = vbNullChar Then ' If no name was found exit
                    MessageBox hwnd, "There was an error getting the droped file.", "WM_DROPFILES", MB_ICONEXCLAMATION Or MB_OK
                    Exit Function
                Else
                    PutEditText WndEditBox, OpenFile(DropFileName)
                    DropFileName = "" 'Clear up
                    DragFinish wParam ' Finish Drag
                End If
            End If
            
            Exit Function
            
        Case WM_SETFOCUS
            SetFocus hwnd
            
        Case WM_CLOSE ' User has clicked the X on the form so we need to destroy the window
            If wParam = 99 Then DestroyWindow WinHwnd: Exit Function
            ans = MessageBox(WinHwnd, "Do you want to quite this program now?", "Quit...", MB_ICONQUESTION Or MB_YESNO)
            If ans = vbNo Then Exit Function ' If the users answer was yes then we can then destroy the window
            DestroyWindow WinHwnd
        'Case WM_DESTROY ' Using this seems to close down the VB IDE so make sure you save any work first
            'PostQuitMessage 0
        Case WM_MOUSEMOVE
            ' Add your mouse move code here
        Case WM_SIZE ' Window is resizeing
            GetClientRect hwnd, MyRect ' Get the width and height information of the widnow
            SetWindowPos WndEditBox, 0, 0, 0, MyRect.Right, MyRect.Bottom, SWP_NOZORDER
            ' the Line above resizes the edit box to the size of the window
            Exit Function
        Case Else
            WinProc = DefWindowProc(hwnd, wMsg, wParam, lParam)  ' Keep fireing the messages back to the window
    End Select
    
End Function

Function GetAddress(Address) As Long
    GetAddress = Address
End Function
