VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileIcons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type uwpIconInfo
    ExePath As String
    Icon As Images
End Type
Dim uwpIcons() As uwpIconInfo
Dim hImageList As Long
Public Sub DrawFileIcon(DC As Long, Graphics As Long, ByVal x As Long, ByVal y As Long, ByVal Path As String)
    Dim hImage As Long, udtFI As SHFILEINFOA
    
    If Path Like "*\WindowsApps\*_*_*\*.exe" Then 'UWP
    
        For i = 1 To UBound(uwpIcons)
            If uwpIcons(i).ExePath = Path Then
                uwpIcons(i).Icon.Draw DC, x, y: Exit Sub
            End If
        Next
        
        'not found
        ReDim Preserve uwpIcons(UBound(uwpIcons) + 1)
        With uwpIcons(UBound(uwpIcons))
            .ExePath = Path
            Set .Icon = New Images
            .Icon.Create TargetDC, 32, 32, False
            Dim Brush As Long, uwpXML As String, CurrentPath As String, uwpIcon As String, uwpPath As String
            CurrentPath = GetCurrentPath(Path)                                                                                           '获得UWP EXE所在文件夹
            uwpXML = ReadXML(CurrentPath & "AppxManifest.xml", "Package//Properties", "Logo")              '获取UWP 图标文件名称
            uwpXML = Replace(uwpXML, ".png", "*.png")                                                                             '添加通配符
            uwpIcon = Dir(CurrentPath & uwpXML)                                                                                      '获取图标
            uwpPath = GetCurrentPath(uwpXML)                                                                                         '获取图标所在目录
            If uwpIcon = "" Then '如果图标不存在，则位于特殊文件夹中（看起来是为了亮色主题准备的）
                If Dir(CurrentPath & uwpPath & "contrast-black", vbDirectory) <> "" Then                                '存在于暗色主题文件夹
                    uwpXML = Replace(uwpXML, uwpPath, uwpPath & "contrast-black\")
                ElseIf Dir(CurrentPath & uwpPath & "contrast-white", vbDirectory) <> "" Then                          '存在于亮色主题文件夹
                    uwpXML = Replace(uwpXML, uwpPath, uwpPath & "contrast-white\")
                End If
                uwpIcon = Dir(CurrentPath & uwpXML)                                                                                 '重新获取图标
            End If
            '绘制
            GdipCreateSolidFill argb(255, 106, 215, 235), Brush
            GdipFillEllipse .Icon.Graphics, Brush, 0, 0, 31, 31
            GdipCreateBitmapFromFile StrPtr(CurrentPath & GetCurrentPath(uwpXML) & uwpIcon), hImage
            GdipDrawImageRect .Icon.Graphics, hImage, 0, 0, 32, 32
            GdipDisposeImage hImage
            GdipDeleteBrush Brush
            .Icon.Draw DC, x, y: Exit Sub
        End With
    Else
        hImage = SHGetFileInfo(Path, ByVal 0&, udtFI, Len(udtFI), SHGFI_LARGEICON Or SHGFI_SYSICONINDEX)
        ImageList_Draw hImage, udtFI.iIcon, DC, x, y, ILD_TRANSPARENT
        ImageList_Destroy hImage
    End If
End Sub

Sub Dispose()
    'ImageList_Destroy hImageList
    For i = 1 To UBound(uwpIcons)
        uwpIcons(i).Icon.Dispose
    Next
End Sub

Private Sub Class_Initialize()
    Dim udtFI As SHFILEINFOA
    ReDim uwpIcons(0)
    AddMember Me
    'hImageList = SHGetFileInfo(vbNullString, ByVal 0&, udtFI, Len(udtFI), SHGFI_ICON Or BASIC_SHGFI_FLAGS Or SHGFI_SYSICONINDEX Or SHGFI_USEFILEATTRIBUTES)
End Sub
