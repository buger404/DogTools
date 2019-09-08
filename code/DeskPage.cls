VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DeskPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type Desk_Files
    Path As String
    Icon As Images
    filename As String
    Type As Integer
End Type

Public HC As Long, VC As Long, UpdateWay As Integer
Dim Draw As New Images, Back As New Images, Back2 As New Images, Back3 As New Images
Dim Mode As Integer, Page As Integer, Mode2 As Boolean
Dim LnkList() As Desk_Files, DirList() As Desk_Files
Dim DeskPath As String, LastCheckTime As Long

Dim FR As Long, FG As Long, FB As Long

Public Style As Long

Dim MouseIndex As Integer, SelectedIndex As Integer, LastClickTime As Long, MouseDowned As Boolean

Public Property Get Visible() As Boolean
    Visible = True
End Property
Sub Refresh()
    SelectedIndex = 0: MouseIndex = 0: MouseDowned = False

    Dim Path As String, file As String
    Path = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    
    ReDim LnkList(0)
    
    Dim Brush As Long, Color As Long, ColorA(3) As Byte
    GdipCreateSolidFill argb(255, 255, 255, 255), Brush
    
    file = Dir(Path & "\")
    Do While file <> ""
        If file Like "*.lnk" Then
            ReDim Preserve LnkList(UBound(LnkList) + 1)
            With LnkList(UBound(LnkList))
                .filename = Replace(file, ".lnk", "")
                .Path = Path & "\" & file
                .Type = 2
                Set .Icon = New Images
                
                .Icon.Create Draw.CompatibleDC, 84, 84, False
                
                GdipSetSolidFillColor Brush, IIf(Style = 0, argb(255, 255, 255, 255), argb(255, 64, 64, 64))
                GdipFillEllipse .Icon.Graphics, Brush, 0, 0, 83, 83
                
                FIcon.DrawFileIcon .Icon.CompatibleDC, .Icon.Graphics, 26, 26, Path & "\" & file

            End With
        End If
        DoEvents
        file = Dir()
    Loop
    
    ReDim DirList(0)
    
    file = Dir(Path & "\", vbDirectory)
    Do While file <> ""
        If (Not (file = "." Or file = "..")) And CreateObject("Scripting.FileSystemObject").FolderExists(Path & "\" & file) Then
            ReDim Preserve DirList(UBound(DirList) + 1)
            With DirList(UBound(DirList))
                .filename = file
                .Path = Path & "\" & file
                .Type = 1
                Set .Icon = New Images
                
                .Icon.Create Draw.CompatibleDC, 84, 84, False
                
                UIImg.Image("folder3.png").Draw .Icon.CompatibleDC, 10, 10
                
            End With
        End If
        DoEvents
        file = Dir(, vbDirectory)
    Loop
    
    file = Dir(Path & "\")
    Do While file <> ""
        If Not (file Like "*.lnk") Then
            ReDim Preserve DirList(UBound(DirList) + 1)
            With DirList(UBound(DirList))
                .filename = file
                .Path = Path & "\" & file
                .Type = 0
                Set .Icon = New Images
                
                .Icon.Create Draw.CompatibleDC, 84, 84, False
                
                GdipSetSolidFillColor Brush, IIf(Style = 0, argb(255, 255, 255, 255), argb(255, 64, 64, 64))
                GdipFillEllipse .Icon.Graphics, Brush, 0, 0, 83, 83
                
                FIcon.DrawFileIcon .Icon.CompatibleDC, .Icon.Graphics, 26, 26, Path & "\" & file
            End With
        End If
        DoEvents
        file = Dir()
    Loop
    
    GdipDeleteBrush Brush
End Sub
Sub FileDrop(ByVal FileList As String)
    Dim Files() As String
    Files = Split(FileList, vbCrLf)
    
    For i = 0 To UBound(Files) - 1
        
    Next
End Sub
Sub KeyUp(ByVal KeyCode As Long)
    If KeyCode = 13 Then
        If SelectedIndex <> 0 Then
            If Mode = 0 Then
                ShellExecuteA 0, "open", LnkList(SelectedIndex).Path, "", "", SW_SHOW
            Else
                ShellExecuteA 0, "open", DirList(SelectedIndex).Path, "", "", SW_SHOW
            End If
        End If
    End If
End Sub
Sub Display()
    
    Select Case UpdateWay
        Case 0
            Dim p As POINT
            GetCursorPos p
            If WindowFromPoint(p.X, p.Y) <> DeskWindow.Hwnd Then Exit Sub
        Case 1
            If GetActiveWindow <> DeskWindow.Hwnd Then Exit Sub
    End Select
    
    If Style = 2 Then
        Back3.Draw Draw.CompatibleDC, 0, 0
    Else
        Back.Draw Draw.CompatibleDC, 0, 0
    End If
    
    If Style <> 2 Then BackImg.ImageByIndex(4).Draw Draw.CompatibleDC, 0, 0, 50
    
    If Style = 2 Then GoTo last
    
    Dim Alpha As Long
    Alpha = 0
    
    With UIImg.Image("app.png")
        SetClickArea2 GWW - 72, 0, 72, 72
        Alpha = IIf(IsMouseIn, 40, 0)
        If Mode = 0 Then Alpha = Alpha + 30
        If IsClick Then
            Mode = 0
            Page = 0
            SelectedIndex = 0: MouseIndex = 0: MouseDowned = False
            Mode2 = False
            ProCore.FadePage mNowShow
        End If
        BackImg.ImageByIndex(1).DrawClip Draw.CompatibleDC, CtrlX, CtrlY, 0, 0, CtrlW, CtrlH, Alpha
        .Draw Draw.CompatibleDC, GWW - 32 - 20, 20
    End With
    
    With UIImg.Image("folder2.png")
        SetClickArea2 GWW - 72, 72, 72, 72
        Alpha = IIf(IsMouseIn, 40, 0)
        If Mode = 1 Then Alpha = Alpha + 30
        If IsClick Then
            Mode = 1
            Page = 0
            SelectedIndex = 0: MouseIndex = 0: MouseDowned = False
            Mode2 = False
            ProCore.FadePage mNowShow
        End If
        BackImg.ImageByIndex(1).DrawClip Draw.CompatibleDC, CtrlX, CtrlY, 0, 0, CtrlW, CtrlH, Alpha
        .Draw Draw.CompatibleDC, GWW - 32 - 20, 20 + 72
    End With
    
    With UIImg.Image("computer.png")
        SetClickArea2 GWW - 72, 72 * 2, 72, 72
        Alpha = IIf(IsMouseIn, 40, 0)
        If Mode = 1 Then Alpha = Alpha + 30
        If IsClick Then
            ShellExecuteA 0, "open", "explorer.exe", "", "", SW_SHOW
        End If
        BackImg.ImageByIndex(1).DrawClip Draw.CompatibleDC, CtrlX, CtrlY, 0, 0, CtrlW, CtrlH, Alpha
        .Draw Draw.CompatibleDC, GWW - 32 - 20, 20 + 72 * 2
    End With
    
    With UIImg.Image("refresh.png")
        SetClickArea2 GWW - 72, GWH - 72 * 2 - 40, 72, 72
        Alpha = IIf(IsMouseIn, 40, 0)
        If IsClick Then
            Call Dispose
            Call Refresh
            ProCore.FadePage mNowShow
        End If
        BackImg.ImageByIndex(1).DrawClip Draw.CompatibleDC, CtrlX, CtrlY, 0, 0, CtrlW, CtrlH, Alpha
        .Draw Draw.CompatibleDC, GWW - 32 - 20, GWH - 72 * 2 + 20 - 40
    End With
    
    Dim Y As Long
    Y = GWH / 2 - ((1 + VC) * 170) / 2 + 170
    Y = Y / 2 - 60 / 2
    
    ProFont.DrawText Draw.Graphics, 0, Y, GWW - 60 / 0.7 / 2 + 1, 70, format(Hour(Now), "00") & " : " & format(Minute(Now), "00"), argb(255, 255, 255, 255), center, 60
    
    If Mode = 0 Then
        DrawFileList LnkList
    Else
        DrawFileList DirList
    End If
    
    With UIImg.Image("wallspaper.png")
        SetClickArea2 GWW - 72, GWH - 72 - 40, 72, 72
        Alpha = IIf(IsMouseIn, 40, 0)
        If IsClick Then
            Call ResetClick: LockX = -1: MouseState = 0
            Dim Wall As New Wallpaperer, Old As String
            Old = DeskPath
            Wall.DownloadWallpaper App.Path & "\desktop.jpg"
            SystemParametersInfoA SPI_SETDESKWALLPAPER, 0, App.Path & "\desktop.jpg", SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE
            DeskPath = App.Path & "\desktop.jpg"
            LoadWallPaper
            ProCore.FadePage mNowShow
            Set Wall = Nothing
            Call Display
            If MsgBox("Love this wallpaper ?", 32 + vbYesNo, "Dog desktop") = vbNo Then
                SystemParametersInfoA SPI_SETDESKWALLPAPER, 0, Old, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE
                DeskPath = Old
                LoadWallPaper
                ProCore.FadePage mNowShow
            End If
        End If
        BackImg.ImageByIndex(1).DrawClip Draw.CompatibleDC, CtrlX, CtrlY, 0, 0, CtrlW, CtrlH, Alpha
        .Draw Draw.CompatibleDC, GWW - 32 - 20, GWH - 72 + 20 - 40
    End With
    
last:
    
    If Moused = False And MouseType = 2 And MouseState = 2 Then
        MenuWindow.Show
        MenuWindow.Move MouseX * 15, MouseY * 15
    End If
    
    Draw.Draw ProDraw.CompatibleDC, 0, 0
    
    If GetTickCount - LastCheckTime >= 1000 Then
        Dim temp As String
        temp = CreateObject("WScript.Shell").RegRead("HKCU\Control Panel\Desktop\Wallpaper")
        If temp <> DeskPath Then
            DeskPath = temp
            LoadWallPaper
        End If
        LastCheckTime = GetTickCount
    End If
End Sub
Private Sub DrawFileList(FileList() As Desk_Files)
    Dim X As Long, Y As Long, Brush As Long, maxPage As Long
    Dim Color As Long, ColorA(3) As Byte
    
    GdipCreatePen1 argb(150, FR, FG, FB), 2, UnitPixel, Brush
    
    Y = GWH / 2 - ((1 + VC) * 170) / 2
    
    For s = 0 To VC - 1
        X = GWW / 2 - (122 * HC) / 2
        Y = Y + 170
        For i = Page * (HC * VC) To Page * (HC * VC) + HC - 1
            If i + 1 + s * HC > UBound(FileList) Then Exit For
            SetClickArea2 X, Y, 84, 84
            With FileList(i + 1 + s * HC)
                .Icon.Draw Draw.CompatibleDC, X, Y, 255 - IIf(Mode2, 100, 0) - IIf(IsMouseIn And MouseState = 1, 100, 0) - IIf(.Type = 1 And (Not IsMouseIn) And (Not (SelectedIndex = i + 1 + s * HC)), 60, 0)
                If IsMouseIn Then
                    If .Type <> 1 Then GdipDrawEllipse Draw.Graphics, Brush, X + 1, Y + 1, 81, 81
                    If MouseIndex <> i + 1 + s * HC Then
                        MouseIndex = i + 1 + s * HC
                        ProCore.FadePage mNowShow
                    End If
                    If MouseState = 1 Then
                        If MouseDowned = False Then MouseDowned = True: ProCore.FadePage mNowShow
                    ElseIf MouseState = 2 Then
                        MouseDowned = False
                    End If
                End If
                If SelectedIndex = i + 1 + s * HC Then
                    If .Type <> 1 Then GdipDrawEllipse Draw.Graphics, Brush, X + 1, Y + 1, 81, 81
                End If
                If IsClick Then
                    If MouseType = 1 Then
                        If Mode2 = False Then
                            If SelectedIndex <> i + 1 + s * HC Then
                                SelectedIndex = i + 1 + s * HC
                            Else
                                If GetTickCount - LastClickTime <= GetDoubleClickTime Then
                                    ShellExecuteA 0, "open", .Path, "", "", SW_SHOW
                                End If
                            End If
                        End If
                        LastClickTime = GetTickCount
                    Else
                        If GetTickCount - ProCore.TransTime >= 400 Then
                            Mode2 = Not Mode2
                            ProCore.FadePage mNowShow
                        Else
                            ShowFileMenu .Path, DeskWindow.Hwnd
                        End If
                    End If
                End If
                
                ProFont.DrawText Draw.Graphics, X - 20, Y + 90, 122, 60, .filename, argb(200, 255, 255, 255), center, 14
                SetClickArea2 Int(X - 20), Int(Y + 90), 122, 60
                If IsClick Then
                    Dim tRetstr As String
                    tRetstr = InputBox("Input its new name .", "Desktop rename", .filename)
                    If tRetstr = "" Then Exit Sub
                    If InStr(tRetstr, "*") Or InStr(tRetstr, "\") Or InStr(tRetstr, "/") Or InStr(tRetstr, """") Or InStr(tRetstr, "?") Or InStr(tRetstr, "<") Or InStr(tRetstr, ">") Or InStr(tRetstr, "|") Or InStr(tRetstr, ":") Or tRetstr = "" Then
                        MsgBox "Wrong filename .", 16, "Desktop"
                        Exit Sub
                    End If
                    If .Type = 2 Then tRetstr = tRetstr & ".lnk"
                    .filename = tRetstr
                    Name (.Path) As (GetCurrentPath(.Path) & .filename)
                    .Path = GetCurrentPath(.Path) & .filename
                    .filename = Replace(tRetstr, ".lnk", "")
                End If
            End With
            
            If Mode2 Then
                With BackImg.ImageByIndex(6)
                    .SetClickArea X + 84, Y - 22
                    .Draw Draw.CompatibleDC, X + 84, Y - 22, IIf(IsMouseIn, 255, 200)
                    If IsClick Then
                        On Error Resume Next
                        Kill FileList(i + 1 + s * HC).Path
                        For p = i + 1 + s * HC To UBound(FileList) - 1
                            FileList(p) = FileList(p + 1)
                        Next
                        ReDim Preserve FileList(UBound(FileList) - 1)
                        ProCore.FadePage mNowShow
                    End If
                End With
            End If
            
            X = X + 122
        Next
    Next
    
    
    maxPage = Int(UBound(FileList) / ((HC * VC) + 1)) + 1
    X = (GWW / 2) - (maxPage * 45) / 2
    For i = 1 To maxPage
        With BackImg.ImageByIndex(5)
            .Draw Draw.CompatibleDC, X, GWH - 90, IIf(Page + 1 = i, 255, 100)
            .SetClickArea X, GWH - 90
            If IsClick Then
                Page = i - 1
                SelectedIndex = 0: MouseIndex = 0: MouseDowned = False
                ProCore.FadePage mNowShow
            End If
        End With
        X = X + 45
    Next
    
    GdipDeletePen Brush
End Sub
Sub Dispose()
    On Error Resume Next
    For i = 1 To UBound(LnkList)
        LnkList(i).Icon.Dispose
    Next
    For i = 1 To UBound(DirList)
        DirList(i).Icon.Dispose
    Next
End Sub
Sub LoadWallPaper()
    Dim Color As Long, ColorA(3) As Byte
    
    Back.LoadImage2 DeskPath, GWW, GWH
    Back2.LoadImage2 DeskPath, 1, 1
    Back3.LoadImage2 DeskPath, GWW, GWH
    Color = GetPixel(Back2.CompatibleDC, 1, 1)
    CopyMemory ColorA(0), Color, 4
    FR = ColorA(0): FG = ColorA(1): FB = ColorA(2)
    If FR < 0 Then FR = 0
    If FG < 0 Then FG = 0
    If FB < 0 Then FB = 0
    
    Back2.LoadImage2 DeskPath, GWW, GWH
    BlurTo2 Back2.CompatibleDC, Back2.CompatibleDC
    Back2.ApplyFade 0, 0, GWH, 130, 255
    Back2.Draw Back.CompatibleDC, 0, 0
End Sub
Private Sub Class_Initialize()
    HC = Val(GetSetting("Dog Tools", "Settings", "HC"))
    If HC = 0 Then HC = 8
    VC = Val(GetSetting("Dog Tools", "Settings", "VC"))
    If VC = 0 Then VC = 2
    UpdateWay = Val(GetSetting("Dog Tools", "Settings", "UpdateWay"))
    Style = Val(GetSetting("Dog Tools", "Settings", "Style"))

    Draw.Create TargetDC, GWW, GWH
    DeskPath = CreateObject("WScript.Shell").RegRead("HKCU\Control Panel\Desktop\Wallpaper")
    
    Back.CreateFromPath2 TargetDC, CreateObject("WScript.Shell").RegRead("HKCU\Control Panel\Desktop\Wallpaper"), GWW, GWH
    Back2.CreateFromPath2 TargetDC, CreateObject("WScript.Shell").RegRead("HKCU\Control Panel\Desktop\Wallpaper"), GWW, GWH
    Back3.CreateFromPath2 TargetDC, CreateObject("WScript.Shell").RegRead("HKCU\Control Panel\Desktop\Wallpaper"), GWW, GWH
    
    LoadWallPaper
    
    AddMember Me
    
    Call Refresh
End Sub




