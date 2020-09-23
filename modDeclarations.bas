Attribute VB_Name = "modDeclarations"
Option Explicit

'''API Declarations
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As String, ByVal fuWinIni As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function SetSysColors Lib "user32" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

'Constant Declarations
Public Const WM_STYLECHANGED = &H7D
Public Const GWL_WNDPROC = (-4)
Public Const WM_HSCROLL = &H114
Public Const WM_VSCROLL = &H115
Public Const SPI_SETDESKWALLPAPER = 20
Public Const SPIF_UPDATEINIFILE = &H1
Public Const SPIF_SENDWININICHANGE = &H2
Public Const HKEY_CURRENT_USER = &H80000001
Public Const ERROR_SUCCESS = 0&
Public Const REG_SZ = 1
Public Const AppName = "XPI Corp. Wallpaper Changer"
Public Const COLOR_BACKGROUND = 1

'Variable Declarations
Public AppPath As String
Public ThumbPath As String
Public AddedToSTartup As Boolean
Public gPrevWndProc As Long
Public gHW As Long
Public m_WinPath As String
Public bBuf() As Byte
Public Image_Width As Long
Public Image_Height As Long
Public Image_Type As eImageType
Public Image_FileSize As Long
Public xFile As String
Public cmdArgs() As String
Public FILE As String
Public X As Long
Public OLD_BGCOLOR

'Type Declarations
Public Type POINTAPI
   X  As Long
   Y  As Long
End Type

Public Type WordBytes
    byte1 As Byte
    byte2 As Byte
End Type

Public Type DWordBytes
    byte1 As Byte
    byte2 As Byte
End Type

Public Type WordWrapper
    Value As Integer
End Type

'Enumeration Declarations
Public Enum eImageType
    itUNKNOWN = 0
    itGIF = 1
    itJPEG = 2
    itPNG = 3
    itBMP = 4
End Enum
