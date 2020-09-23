Attribute VB_Name = "modLoad"
Option Explicit

'''''This module contains the Subs that change the wallpaper on a command line arg.
Public Sub comChange()
    Dim subT As Long
    Dim retval

        ReadImageInfo (FILE)
        Image_Width = Image_Width / 2
        Image_Height = Image_Height / 2
        If Image_Height = Image_Width Then
            If Image_Height <= 150 Then
                If Image_Height >= 50 Then
                    Form1.optTile.Value = True
                    GoTo Complete
                End If
            End If
        End If
        subT = Image_Width - Image_Height
        If subT = 32 Then
            Form1.optStretch = True
            GoTo Complete
        End If
        If subT = 64 Then
            Form1.optStretch = True
            GoTo Complete
        End If
        If subT = 128 Then
            Form1.optStretch = True
            GoTo Complete
        End If
        Form1.optCenter.Value = True
        
Complete:

    If Form1.optTile = True Then
        'Tile Wallpaper
        Call SaveString(HKEY_CURRENT_USER, "Control Panel\Desktop", "WallpaperStyle", "0")
        Call SaveString(HKEY_CURRENT_USER, "Control Panel\Desktop", "TileWallpaper", "1")
    End If
    If Form1.optCenter = True Then
        
        retval = MsgBox("Would you like to change the system background color?", vbYesNo + vbQuestion, "System Background Color")
        If retval = vbYes Then
            Form1.CDL1.ShowColor
            Call SetBGColor(Form1.CDL1.Color)
        End If
        'Center wallpaper
        Call SaveString(HKEY_CURRENT_USER, "Control Panel\Desktop", "WallpaperStyle", "0")
        Call SaveString(HKEY_CURRENT_USER, "Control Panel\Desktop", "TileWallpaper", "0")
    End If
    If Form1.optStretch = True Then
        'Stretch wallpaper
        Call SaveString(HKEY_CURRENT_USER, "Control Panel\Desktop", "WallpaperStyle", "2")
        Call SaveString(HKEY_CURRENT_USER, "Control Panel\Desktop", "TileWallpaper", "0")
    End If
        'Set Wallpaper to the loaded image
        
    SystemParametersInfo SPI_SETDESKWALLPAPER, 2&, ByVal xFile, SPIF_SENDWININICHANGE Or SPIF_UPDATEINIFILE

End Sub

Sub CheckCmdLineArgs(cmdLine As String)
Dim currChr
Dim j As Integer
Dim fileTitle As String
Dim filePath As String

If cmdLine <> "" Then
    Dim numCmdLine As Integer, cmdLen As Long, i As Integer
    ReDim cmdArgs(0)
    cmdLen = Len(cmdLine)
    numCmdLine = 0
    
    For i = 1 To cmdLen
        currChr = Mid(cmdLine, i, 1)
        If currChr = Chr(34) Then
            Do
                i = i + 1
                currChr = Mid(cmdLine, i, 1)
                If currChr = Chr(34) Then Exit Do
                cmdArgs(numCmdLine) = cmdArgs(numCmdLine) & currChr
                
            Loop
            i = i + 1: numCmdLine = numCmdLine + 1
            ReDim Preserve cmdArgs(numCmdLine)
        
        ElseIf currChr = " " Then
            i = i + 1
            Do
                currChr = Mid(cmdLine, i, 1)
                If currChr = " " Or currChr = "" Then Exit Do
                cmdArgs(numCmdLine) = cmdArgs(numCmdLine) & currChr
                i = i + 1
            Loop
            i = i - 1
            numCmdLine = numCmdLine + 1
            ReDim Preserve cmdArgs(numCmdLine)
        
        ElseIf i = 1 Then
            currChr = Mid(cmdLine, i, 1)
            Do
                
                currChr = Mid(cmdLine, i, 1)
                If currChr = " " Or currChr = "" Then Exit Do
                cmdArgs(numCmdLine) = cmdArgs(numCmdLine) & currChr
                i = i + 1
            Loop
            numCmdLine = numCmdLine + 1
            ReDim Preserve cmdArgs(numCmdLine)
            i = i - 1
        
        End If
    Next i
        
    For i = 1 To numCmdLine
        
        For j = Len(cmdArgs(i - 1)) To 1 Step -1
            If Mid(cmdArgs(i - 1), j - 1, 1) = "\" Then Exit For
        Next j
        fileTitle = Mid(cmdArgs(i - 1), j)
        filePath = Mid(cmdArgs(i - 1), 1, j - 1)
        
        FILE = filePath & fileTitle
    Next i
End If
End Sub

