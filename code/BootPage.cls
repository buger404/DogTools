VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BootPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Draw As New Images
Dim Page As Integer
Private Type BootInfo
    name As String
    Des As String
    Cmd As String
    Location As String
End Type
Private Type BootList
    List() As BootInfo
End Type
Dim Boots As BootList
Dim NewPath As String

Public Property Get Visible() As Boolean
    Visible = (mNowShow = "BootPage")
End Property

Sub Display()
    BackImg.ImageByIndex(1).Draw Draw.CompatibleDC, 0, 0
    
    ProFont.DrawText Draw.Graphics, 0, 110, GWW, 30, "Boot items", argb(255, 56, 60, 107), center, 25, Regular
    ProFont.DrawText Draw.Graphics, 0, 150, GWW, 30, "Manager your boot items .", argb(125, 56, 60, 107), center, 18, Regular
    
    SetClickArea2 20, 65, 20, 20
    ProFont.DrawText Draw.Graphics, 12, 65, 30, 30, "<", IIf(IsMouseIn, argb(255, 0, 176, 240), argb(255, 129, 129, 129)), center, 18, Regular
    If IsClick Then
        ProCore.FadePage "MainPage"
    End If
    
    Dim Start As Long, IcoPath As String, Username As String
    Username = CreateObject("Wscript.Network").Username
    
    Start = Page * 5 + 1
    For i = Start To Start + 4
        If i <= UBound(Boots.List) Then
            With Boots.List(i)
                ProFont.DrawText Draw.Graphics, 102, 200 + (i - Start) * 60, 700, 30, .name, argb(255, 129, 129, 129), near, 18, Regular
                ProFont.DrawText Draw.Graphics, 104, 200 + (i - Start) * 60 + 24, 900, 30, .Cmd, argb(155, 129, 129, 129), near, 14, Regular
                If .Location = "Startup" Then
                    IcoPath = "C:\Users\" & Username & "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\" & Boots.List(i).Cmd
                Else
                    IcoPath = .Cmd
                    If Right(.Cmd, 4) <> ".exe" Then
                        IcoPath = Replace(IcoPath, """", "")
                        IcoPath = Left(IcoPath, InStr(IcoPath, ".exe") + 3)
                    End If
                End If
                FIcon.DrawFileIcon Draw.CompatibleDC, Draw.Graphics, 45, 200 + (i - Start) * 60 + 7, IcoPath
                SetClickArea2 50 + 750, 200 + (i - Start) * 60, 100, 30
                ProFont.DrawText Draw.Graphics, 50 + 750, 200 + (i - Start) * 60, 100, 30, "Delete", IIf(IsMouseIn, argb(255, 56, 60, 107), argb(255, 129, 129, 129)), near, 18, Regular
            End With
            If IsClick Then
                On Error Resume Next
                If Boots.List(i).Location = "Startup" Then
                    Kill "C:\Users\" & CreateObject("Wscript.Network").Username & "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\" & Boots.List(i).Cmd
                Else
                    CreateObject("Wscript.Shell").RegDelete Boots.List(i).Location & "\" & Boots.List(i).name
                End If
                GetBoots
                If Err.Number <> 0 Then
                    MsgBox "Failed to delete : " & Err.Number & vbCrLf & Err.Description, 48, "Boot Items"
                    Err.Clear
                End If
            End If
            With UIImg.Image("folder.png")
                .SetClickArea 50 + 710, 200 + (i - Start) * 60
                .Draw Draw.CompatibleDC, 50 + 710, 200 + (i - Start) * 60, IIf(IsMouseIn, 255, 200)
                If IsClick Then
                    ShellExecuteA 0, "open", "explorer.exe", "/select," & IcoPath, "", SW_SHOW
                End If
            End With
        End If
    Next
    
    If Page < Int(UBound(Boots.List) / 5) Then
        SetClickArea2 GWW - 170, 220 + 9 * 30, 120, 30
        ProFont.DrawText Draw.Graphics, GWW - 170, 220 + 9 * 30, 120, 30, "Next Page >", IIf(IsMouseIn, argb(255, 56, 60, 107), argb(255, 129, 129, 129)), center, 18, Regular
        If IsClick Then Page = Page + 1
    End If
    
    ProFont.DrawText Draw.Graphics, 0, 220 + 9 * 30, GWW, 30, Page + 1 & " / " & Int(UBound(Boots.List) / 5) + 1, argb(255, 129, 129, 129), center, 18, Regular
    
    If Page > 0 Then
        SetClickArea2 50, 220 + 9 * 30, 120, 30
        ProFont.DrawText Draw.Graphics, 50, 220 + 9 * 30, 120, 30, "< Last Page", IIf(IsMouseIn, argb(255, 56, 60, 107), argb(255, 129, 129, 129)), center, 18, Regular
        If IsClick Then Page = Page - 1
    End If
    
    BackImg.ImageByIndex(2).DrawClip Draw.CompatibleDC, 0, GWH - 60, 0, 0, GWW, 60
    
    ProFont.DrawText Draw.Graphics, 30, GWH - 46, 200, 30, "Path", argb(255, 129, 129, 129), near, 18, Regular
    ProFont.DrawText Draw.Graphics, 100, GWH - 46, 500, 30, NewPath, argb(255, 169, 169, 169), near, 18, Regular
    SetClickArea2 100, GWH - 46, 500, 30
    If IsShowEdit(NewPath) Then
        If IsRetEdit Then
            NewPath = GetRetEdit
        End If
    End If
    
    Dim RealList As Boolean
    RealList = "" = ""
    CtrlImg.ImageByIndex(IIf(RealList, 2, 1)).DrawCtrl Draw.CompatibleDC, Draw.Graphics, GWW - 115, GWH - 45, "Add", IIf(RealList, argb(255, 255, 255, 255), argb(255, 129, 129, 129)), 16, Regular, Button
    If IsClick Then
        On Error Resume Next
        Dim filename() As String
        filename = Split(NewPath, "\")
        CreateObject("Wscript.Shell").Regwrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run\" & filename(UBound(filename)), NewPath
        GetBoots
        If Err.Number <> 0 Then
            MsgBox "Failed to create : " & Err.Number & vbCrLf & Err.Description, 48, "Boot Items"
            Err.Clear
        Else
            MsgBox "Success", 64, "Boot Items"
        End If
    End If
    
    Draw.Draw ProDraw.CompatibleDC, 0, 0
End Sub
Sub GetBoots()
    ReDim Boots.List(0)
    Dim objWMIService As Object, colStartupCommands As Object, objStartupCommand As Object
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colStartupCommands = objWMIService.ExecQuery("SELECT * FROM Win32_StartupCommand")
    For Each objStartupCommand In colStartupCommands
        ReDim Preserve Boots.List(UBound(Boots.List) + 1)
        With Boots.List(UBound(Boots.List))
            .Cmd = objStartupCommand.Command
            .Des = objStartupCommand.Description
            .Location = objStartupCommand.Location
            .Location = Replace(.Location, "HKU", "HKEY_USERS")
            .Location = Replace(.Location, "HKCU", "HKEY_CURRENT_USER")
            .Location = Replace(.Location, "HKLM", "HKEY_LOCAL_MACHINE")
            .name = objStartupCommand.name
        End With
    Next
End Sub
Private Sub Class_Initialize()
    NewPath = "..."

    Draw.Create TargetDC, GWW, GWH
    ReDim Boots.List(0)
End Sub


