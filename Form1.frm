VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "XPI Corp. Wallpaper Changer"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8085
   Icon            =   "Form1.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   8085
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picSplash 
      Height          =   465
      Left            =   6675
      Picture         =   "Form1.frx":2892
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   12
      Top             =   7575
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   202
      Picture         =   "Form1.frx":B806
      ScaleHeight     =   465
      ScaleWidth      =   7680
      TabIndex        =   11
      Top             =   5475
      Width           =   7710
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2025
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   6075
      Width           =   3990
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   6600
      Top             =   6975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Style"
      Height          =   1515
      Left            =   2047
      TabIndex        =   6
      Top             =   6525
      Width           =   1890
      Begin VB.OptionButton optTile 
         Caption         =   "Tile"
         Height          =   240
         Left            =   375
         TabIndex        =   9
         Top             =   1125
         Width           =   1215
      End
      Begin VB.OptionButton optCenter 
         Caption         =   "Center"
         Height          =   240
         Left            =   375
         TabIndex        =   8
         Top             =   712
         Width           =   1215
      End
      Begin VB.OptionButton optStretch 
         Caption         =   "Stretch"
         Height          =   240
         Left            =   375
         TabIndex        =   7
         Top             =   300
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controls"
      Height          =   1515
      Left            =   4147
      TabIndex        =   3
      Top             =   6525
      Width           =   1890
      Begin VB.CommandButton Command2 
         Caption         =   "Change"
         Height          =   465
         Left            =   300
         TabIndex        =   5
         Top             =   900
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Open"
         Height          =   465
         Left            =   315
         TabIndex        =   4
         Top             =   300
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog CDL1 
      Left            =   150
      Top             =   6075
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   5825
      Left            =   175
      ScaleHeight     =   384
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   512
      TabIndex        =   2
      Top             =   150
      Width           =   7735
   End
   Begin VB.PictureBox PicPreview 
      Height          =   315
      Left            =   75
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   75
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox PicScreen 
      Height          =   315
      Left            =   90
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   90
      Visible         =   0   'False
      Width           =   315
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim xFile As String
Dim subT As Long
CDL1.DialogTitle = "Open Image"
CDL1.Filter = "All Image Files (*.jpg, *.bmp, *.gif)|*.jpg;*.bmp;*.gif"
CDL1.ShowOpen
xFile = WinPath & "XPI Wallpaper.bmp"
    If CDL1.Filename = "" Then
        Exit Sub
    Else
        Call GetBGColor
        Picture1.BackColor = OLD_BGCOLOR
        Text1.Text = CDL1.fileTitle
        FILE = CDL1.Filename
        Picture1.Cls
        ReadImageInfo (CDL1.Filename)
        Image_Width = Image_Width / 2
        Image_Height = Image_Height / 2
        If Image_Height = Image_Width Then
            If Image_Height <= 150 Then
                If Image_Height >= 50 Then
                    Call optTile_Click
                    optTile.Value = True
                    Exit Sub
                End If
            End If
        End If
        subT = Image_Width - Image_Height
        If subT = 32 Then
            Call optStretch_Click
            optStretch = True
            Exit Sub
        End If
        If subT = 64 Then
            Call optStretch_Click
            optStretch = True
            Exit Sub
        End If
        If subT = 128 Then
            Call optStretch_Click
            optStretch = True
            Exit Sub
        End If
        
        
        Call optCenter_Click
        optCenter.Value = True
    End If
    

 
 End Sub

Private Sub Command2_Click()

On Error GoTo XPI
Dim xFile As String
Dim retval
CheckCmdLineArgs (Command)
If FILE = "" Then
    GoTo XPI
End If
       xFile = WinPath & "XPI Wallpaper.bmp"
    If 1 = 2 Then
        PicScreen.Cls
        PicScreen.PaintPicture LoadPicture(FILE), 0, 0, PicScreen.ScaleWidth, PicScreen.ScaleHeight
    Else
        Set PicScreen.Picture = LoadPicture(FILE)
    End If
 
    'Load Picture in Picbox & Save to "C:\WINDOWS\XPI Wallpaper.bmp"
    PicPreview.PaintPicture PicScreen.Picture, 0, 0, PicPreview.ScaleWidth, PicPreview.ScaleHeight
    PicPreview.Refresh
    SavePicture PicScreen.Picture, xFile
   
   If optTile = True Then
        'Tile Wallpaper
        Call SaveString(HKEY_CURRENT_USER, "Control Panel\Desktop", "WallpaperStyle", "0")
        Call SaveString(HKEY_CURRENT_USER, "Control Panel\Desktop", "TileWallpaper", "1")
    End If
    If optCenter = True Then
        
        retval = MsgBox("Would you like to change the system background color?", vbYesNo + vbQuestion, "System Background Color")
        If retval = vbYes Then
            CDL1.ShowColor
            Call SetBGColor(CDL1.Color)
            Picture1.BackColor = CDL1.Color
            Call optCenter_Click
        End If
        'Center wallpaper
        Call SaveString(HKEY_CURRENT_USER, "Control Panel\Desktop", "WallpaperStyle", "0")
        Call SaveString(HKEY_CURRENT_USER, "Control Panel\Desktop", "TileWallpaper", "0")
    End If
    If optStretch = True Then
        'Stretch wallpaper
        Call SaveString(HKEY_CURRENT_USER, "Control Panel\Desktop", "WallpaperStyle", "2")
        Call SaveString(HKEY_CURRENT_USER, "Control Panel\Desktop", "TileWallpaper", "0")
    End If
        'Set Wallpaper to the loaded image
    SystemParametersInfo SPI_SETDESKWALLPAPER, 2&, ByVal xFile, SPIF_SENDWININICHANGE Or SPIF_UPDATEINIFILE
XPI:

    
End Sub

Private Sub Form_DblClick()
    optStretch = False
    optTile = False
    optCenter = False
End Sub

Private Sub Form_Load()
On Error GoTo Erroneous

Picture1.PaintPicture picSplash.Picture, 0, 0, 512, 384
Picture1.Refresh

CheckCmdLineArgs (Command)
If FILE = "" Then
    GoTo XPI
End If
       xFile = WinPath & "XPI Wallpaper.bmp"
    If 1 = 2 Then
     
        PicScreen.Cls
        PicScreen.PaintPicture LoadPicture(FILE), 0, 0, PicScreen.ScaleWidth, PicScreen.ScaleHeight
    Else
        Set PicScreen.Picture = LoadPicture(FILE)
    End If
 
    'Picture zu Image
    PicPreview.PaintPicture PicScreen.Picture, 0, 0, PicPreview.ScaleWidth, PicPreview.ScaleHeight
    PicPreview.Refresh
    SavePicture PicScreen.Picture, xFile
   
    Call comChange
   
Erroneous:
End
XPI:
End Sub

Private Sub optCenter_Click()
If CDL1.Filename = "" Then
    Exit Sub
End If

    Picture1.Cls
    Picture1.PaintPicture LoadPicture(CDL1.Filename), Picture1.ScaleWidth / 2 - Image_Width / 2, Picture1.ScaleHeight / 2 - Image_Height / 2, Image_Width, Image_Height
End Sub

Private Sub optStretch_Click()
If CDL1.Filename = "" Then
    Exit Sub
End If

Picture1.PaintPicture LoadPicture(CDL1.Filename), 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, 0, 0
End Sub

Private Sub optTile_Click()
If CDL1.Filename = "" Then
    Exit Sub
End If

Dim i, j
 For i = 0 To Picture1.ScaleWidth Step Image_Width
  For j = 0 To Picture1.ScaleHeight Step Image_Height
   Picture1.PaintPicture LoadPicture(CDL1.Filename), i, j, Image_Width, Image_Height, 0, 0
  Next
 Next

End Sub
