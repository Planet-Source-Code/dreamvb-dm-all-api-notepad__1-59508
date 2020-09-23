Attribute VB_Name = "ModMain"
Private Declare Function LoadIcon Lib "user32.dll" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long

' Hi all this is my new update of my ALL API form it now sort of turned into a notepad app
' anyway what are the features:
' Open and Save Text Files
' Edit menu support
' Cut, Copy Paste, Select ALL
' New About box added the API one
' The editor also suppots s margin
' and also supports the draging and droping of files.
' well I hope you like this update and just remmber that is is Pure API not just VB Controls as you can tell from the fact there is no form :)

' well anyway hope you like the code please feel free to do as you please with it
' I try and make a new update tommrow for it. try and add some more features.
' Also note that this editor does have a size limit of 29.2kb I am not sure
' but if anyone has any inside this this let me know.

Sub Main()
    ' This is the part that loads our window first
    WindowCaption = "DM API-NotePad" 'Caption for our new window
    WinClassName = "MyWinClass" 'Class name for our window
    
    'Fill the class struc with the information needed for the new window
    With wc
        .lpfnwndproc = GetAddress(AddressOf WinProc)
        .cbClsextra = 0
        .cbWndExtra2 = 0
        .hInstance = App.hInstance
        .lpszMenuName = vbNullString
        .style = 0
        .hbrBackground = 16
        .lpszClassName = WinClassName
    End With
    
    If RegisterClass(wc) = 0 Then ' Check if the windows class was registered
        MessageBox 0, "RegisterClass Faild.", "Error", MB_ICONEXCLAMATION Or MB_OK
        End
    Else
        ' Create the window
        WinHwnd = CreateWindowEx(0&, WinClassName, WindowCaption, _
        WindowStyle, CW_USEDEFAULT, CW_USEDEFAULT, 300, 300, 0, 0, App.hInstance, ByVal 0&)
    
        If WinHwnd = 0 Then ' Check if the window was created
            MessageBox 0, "CreateWindowEx Faild.", "Error", MB_ICONEXCLAMATION Or MB_OK
            Exit Sub
            End
        Else
            WndDC = GetDC(WinHwnd) ' Get the Windows DC
            ShowWindow WinHwnd, SW_NORMAL ' Show the window in normal mode
            UpdateWindow WinHwnd ' Update the new window
            
            DragAcceptFiles WinHwnd, True ' Allow our windows to Accept droped files
            
            'Do the Message Loop
            Do While GetMessage(WinMsg, WinHwnd, 0, 0) > 0
                TranslateMessage WinMsg
                DispatchMessage WinMsg
                DoEvents
            Loop
        End If
    End If
    
End Sub

