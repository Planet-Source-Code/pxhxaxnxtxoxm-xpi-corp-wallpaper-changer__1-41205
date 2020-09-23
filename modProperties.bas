Attribute VB_Name = "modProperties"
Option Explicit

Public Property Get ImageWidth() As Long
    ImageWidth = Image_Width
End Property

Public Property Get ImageHeight() As Long
    ImageHeight = Image_Height
End Property

Public Property Get FileSize() As Long
    FileSize = Image_FileSize
End Property
