Attribute VB_Name = "modSubs"
Option Explicit
'''''This sub saves strings to the registry
Public Sub SaveString(hKey As Long, strpath As String, strValue As String, strdata As String)
   Dim keyhand As Long
   Dim r As Long
   X = RegCreateKey(hKey, strpath, keyhand)
   X = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
   X = RegCloseKey(keyhand)
End Sub
'''''

'''''This sub does what it says...It reads the attributes of the image file.
Public Sub ReadImageInfo(sFileName As String)
    
    Dim i As Long
    Dim Size As Integer
    
    Image_Width = 0
    Image_Height = 0
    Image_FileSize = 0
    Image_Type = itUNKNOWN
    Size = FreeFile
    Open sFileName For Binary As Size
    Image_FileSize = LOF(Size)
    ReDim bBuf(Image_FileSize)
    Get #Size, 1, bBuf()
    Close Size
'Check For PNG
    If bBuf(0) = 137 And bBuf(1) = 80 And bBuf(2) = 78 Then
        Image_Type = itPNG
        If Image_Type Then
            Image_Width = BEWord(18)
            Image_Height = BEWord(22)
        End If
    End If
' Check For GIF
    If bBuf(0) = 71 And bBuf(1) = 73 And bBuf(2) = 70 Then
        Image_Type = itGIF
        Image_Width = LEWord(6)
        Image_Height = LEWord(8)
    End If
' Check For BMP
    If bBuf(0) = 66 And bBuf(1) = 77 Then
        Image_Type = itBMP
        Image_Width = LEWord(18)
        Image_Height = LEWord(22)
    End If
' Check For JPEG
    If Image_Type = itUNKNOWN Then
        Dim lPos As Long
        Do
            If (bBuf(lPos) = &HFF And bBuf(lPos + 1) = &HD8 And bBuf(lPos + 2) = &HFF) _
            Or (lPos >= Image_FileSize - 10) Then Exit Do
            lPos = lPos + 1
        Loop
        lPos = lPos + 2
        If lPos >= Image_FileSize - 10 Then Exit Sub
        Do
            Do
                If bBuf(lPos) = &HFF And bBuf(lPos + 1) <> &HFF Then Exit Do
                lPos = lPos + 1
                If lPos >= Image_FileSize - 10 Then Exit Sub
            Loop
            lPos = lPos + 1
            If (bBuf(lPos) >= &HC0) And (bBuf(lPos) <= &HC3) Then Exit Do
            lPos = lPos + BEWord(lPos + 1)
            If lPos >= Image_FileSize - 10 Then Exit Sub
        Loop
        Image_Type = itJPEG
        Image_Height = BEWord(lPos + 4)
        Image_Width = BEWord(lPos + 6)
    End If
    ReDim bBuf(0)
End Sub
'''''

'''''These 2 subs get and set the BG Color of the desktop
Public Sub GetBGColor()
    OLD_BGCOLOR = GetSysColor(COLOR_BACKGROUND)
End Sub

Public Sub SetBGColor(NEW_BGCOLOR)
    SetSysColors 1, COLOR_BACKGROUND, NEW_BGCOLOR
End Sub
'''''

