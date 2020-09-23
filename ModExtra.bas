Attribute VB_Name = "ModExtra"
Public Function IsFileHere(lzFilename As String) As Boolean
    ' Checks if a given filename is found
    If Dir(lzFilename) = "" Then IsFileHere = False: Exit Function Else IsFileHere = True
End Function

Public Function SaveToFile(lzFile As String, sData As String)
Dim nFile As Long
    nFile = FreeFile ' pointer to free file
    Open lzFile For Binary As #nFile
        Put #nFile, , sData ' save sdata contents to the file
    Close #nFile ' close file
End Function

Public Function GetFileExt(lzFile As String) As String
Dim ipos As Integer
    ' used to get the files ext eg GetFileExt "hello.txt" returns txt
    ipos = InStrRev(lzFile, ".", Len(lzFile), vbTextCompare)
    GetFileExt = Mid(lzFile, ipos + 1, Len(lzFile))
End Function

Public Function OpenFile(lzFile As String) As String
Dim s_Data As String
Dim nFile As Long
    nFile = FreeFile
    ' Open a File
    Open lzFile For Binary As #nFile
        s_Data = Space(LOF(1))
        Get #nFile, , s_Data
    Close #nFile
    
    OpenFile = s_Data
    s_Data = ""
    
End Function

Function FixPath(lzpath As String) As String
    If Right(FixPath, 1) = "\" Then FixPath = lzpath: Exit Function Else FixPath = lzpath & "\"
End Function

Function AddFileExt(index As Integer) As String
    If index = 1 Then
        AddFileExt = ".txt"
    ElseIf index = 2 Then
        AddFileExt = ".diz"
    ' add others
    End If
    
End Function

Function GetFileDropCount(hDrop As Long) As Long
'Get the number of files been droped
    GetFileDropCount = DragQueryFile(hDrop, &HFFFFFFFF, 0, 0)
End Function

Function GetFileDrop(hDrop As Long) As String
Dim lzFileDrop As String
Dim lRet As Long
    lzFileDrop = Space(215) ' Create some room to store the droped filename
    lRet = DragQueryFile(hDrop, 0, lzFileDrop, Len(lzFileDrop))
    If lRet = 0 Then
        GetFileDrop = vbNullChar
        lzFileDrop = ""
        Exit Function
    Else
        GetFileDrop = Left(lzFileDrop, InStr(1, lzFileDrop, Chr(0), vbBinaryCompare) - 1)
    End If
    
    
End Function
