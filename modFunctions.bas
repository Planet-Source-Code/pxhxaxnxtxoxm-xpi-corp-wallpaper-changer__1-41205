Attribute VB_Name = "modFunctions"
Option Explicit

'''''This function retrieves the Windows path.
Public Function WinPath() As String
    If m_WinPath = "" Then
        m_WinPath = String(1024, 0)
        GetWindowsDirectory m_WinPath, Len(m_WinPath)
        m_WinPath = Left(m_WinPath, InStr(m_WinPath, Chr(0)) - 1)
        If Right(m_WinPath, 1) <> "\" Then m_WinPath = m_WinPath & "\"
    End If
    WinPath = m_WinPath
End Function
'''''

''''' These 2 functions are used in getting the images attributes.
Public Function LEWord(position As Long) As Long
    Dim x1 As WordBytes
    Dim x2 As WordWrapper
    x1.byte1 = bBuf(position)
    x1.byte2 = bBuf(position + 1)
    LSet x2 = x1
    LEWord = x2.Value
End Function

Public Function BEWord(position As Long) As Long
    Dim x1 As WordBytes
    Dim x2 As WordWrapper
    x1.byte1 = bBuf(position + 1)
    x1.byte2 = bBuf(position)
    LSet x2 = x1
    BEWord = x2.Value
End Function
'''''
