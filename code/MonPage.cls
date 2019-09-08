VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MonPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Draw As New Images
Dim Page As Integer, Mode As Integer
Private Type RecordInfo
    ImName As String
    Path As String
    Time As Single
    Title As String
    data(4) As String
End Type
Private Type RecordList
    List() As RecordInfo
End Type
Dim Records As RecordList
Dim AddL As RecordInfo
Dim LastPID As Long, StartTime As Long, LastHwnd As Long, LastText As String * 255
Dim filename As String, FileName3 As String
Dim TotalTime As Single, Report As RecordList

Public Property Get Visible() As Boolean
    Visible = (mNowShow = "MonPage")
End Property
Sub CarryKeyBoard(Str As String)
    If AddL.ImName = "" Then Exit Sub
    Open DataPath & "\Logs\Keyboard\" & AddL.ImName & " - " & year(Now) & "年" & Month(Now) & "月" & Day(Now) & "日.txt" For Append As #1
    Print #1, Now & "    " & Str
    Close #1
End Sub
Sub Carry2()
    Dim FileName2 As String
    FileName2 = DataPath & "\Monitor\" & year(Now) & "年" & Month(Now) & "月" & Day(Now) & "日.db"
    If filename <> FileName2 Then
        ReDim Records.List(0)
        filename = FileName2
    End If
End Sub
Sub Carry()
    Dim Title As String * 255, Hwnd As Long, PID As Long
    Dim Title2 As String
    Dim ImN() As String, ImName2 As String, ImName As String
    Hwnd = GetForegroundWindow
    GetWindowThreadProcessId Hwnd, PID
    
    If LastHwnd = Hwnd Then
        GetWindowTextA Hwnd, Title, 255
        If LastText <> Title Then
            AddL.Title = Title
            Log "Monitor", "窗口标题发生变化： " & UnSpace(AddL.Title) & " (" & AddL.ImName & ")。"
            LastText = Title
        End If
        Exit Sub
    End If
    
    If LastHwnd <> 0 Then
        ReDim Preserve Records.List(UBound(Records.List) + 1)
        Records.List(UBound(Records.List)) = AddL
        Open filename For Binary As #1
        Put #1, , Records
        Close #1
        AddL.Time = GetTickCount - StartTime
        Log "Monitor", "移开了窗口 " & AddL.Title & " (" & AddL.ImName & ")，距离打开时间 " & Int(AddL.Time / 100) / 10 & "s。"
    End If
    
    LastPID = PID: LastHwnd = Hwnd
    StartTime = GetTickCount
    
    GetWindowTextA Hwnd, Title, 255
    LastText = Title
    
    ImName2 = GetProcessPath(Hwnd)
    ImN = Split(ImName2, "\")
    ImName = LCase(ImN(UBound(ImN)))

    Title2 = LCase(UnSpace(Title))

    With AddL
        .ImName = ImName
        .Title = Title2
        .Path = ImName2
    End With
    
    Log "Monitor", "切换到了窗口 " & AddL.Title & " (" & AddL.ImName & ")。"
End Sub
Sub Display2()
    BackImg.ImageByIndex(1).Draw Draw.CompatibleDC, 0, 0
    
    Dim File2() As String
    File2 = Split(FileName3, "\")
    
    ProFont.DrawText Draw.Graphics, 0, 100, GWW, 28, Replace(Replace(Replace(Replace(File2(UBound(File2)), ".db", ""), "年", "."), "月", "."), "日", ""), argb(125, 42, 52, 0), center, 24, Regular
    
    SetClickArea2 GWW - 140, 65, 120, 30
    ProFont.DrawText Draw.Graphics, GWW - 140, 65, 120, 30, "Next day >", IIf(IsMouseIn, argb(255, 0, 176, 240), argb(255, 129, 129, 129)), center, 18, Regular
    If IsClick Then
        Dim test As String
        test = Dir(DataPath & "\Monitor\")
        Do While (test <> "" And DataPath & "\Monitor\" & test <> FileName3)
            test = Dir()
        Loop
        test = Dir()
        If test = "" Then test = Dir(DataPath & "\Monitor\")
        FileName3 = DataPath & "\Monitor\" & test
        If FileName3 = "" Then FileName3 = filename
        LoadReport FileName3
    End If
    
    SetClickArea2 20, 65, 20, 20
    ProFont.DrawText Draw.Graphics, 12, 65, 30, 30, "<", IIf(IsMouseIn, argb(255, 0, 176, 240), argb(255, 129, 129, 129)), center, 18, Regular
    If IsClick Then
        Mode = 0
        ProCore.FadePage mNowShow
    End If
    
    Dim Start As Long
    
    Start = Page * 5 + 1
    For i = Start To Start + 4
        If i <= UBound(Report.List) Then
            With Report.List(i)
                BackImg.ImageByIndex(2).DrawClip Draw.CompatibleDC, 80, 160 + (i - Start) * 70, 0, 0, GWW - 160, 60
                BackImg.ImageByIndex(3).DrawClip Draw.CompatibleDC, 80, 160 + (i - Start) * 70, 0, 0, .Time / TotalTime * (GWW - 160), 60, 80 + .Time / TotalTime * 170
                FIcon.DrawFileIcon Draw.CompatibleDC, Draw.Graphics, 80 + 14, 160 + (i - Start) * 70 + 14, .Path
                ProFont.DrawText Draw.Graphics, 140, 160 + (i - Start) * 70 + 4, 400, 30, .ImName, argb(255, 129, 129, 129), near, 18, Regular
                ProFont.DrawText Draw.Graphics, 140, 160 + (i - Start) * 70 + 29, 400, 30, format(Int(.Time * 10) / 10, "0.0") & "s", argb(255, 129, 129, 129), near, 18, Regular
                ProFont.DrawText Draw.Graphics, 410, 160 + (i - Start) * 70 + 4, 400, 50, format(Int(.Time / TotalTime * 1000) / 10, "0.0") & "%", argb(255, 109, 109, 109), far, 36, Regular
            End With
            With UIImg.Image("folder.png")
                .SetClickArea 50 + 610, 160 + (i - Start) * 70 + 17
                .Draw Draw.CompatibleDC, 50 + 610, 160 + (i - Start) * 70 + 17, IIf(IsMouseIn, 255, 200)
                If IsClick Then
                    ShellExecuteA 0, "open", "explorer.exe", "/select," & Report.List(i).Path, "", SW_SHOW
                End If
            End With
            With UIImg.Image("keyboard.png")
                .SetClickArea 50 + 560, 160 + (i - Start) * 70 + 17
                .Draw Draw.CompatibleDC, 50 + 560, 160 + (i - Start) * 70 + 17, IIf(IsMouseIn, 200, 160)
                If IsClick Then
                    Dim logFile As String, test2() As String
                    test2 = Split(filename, "\")
                    logFile = Replace(test2(UBound(test2)), ".db", ".txt")
                    If Dir(DataPath & "\Logs\Keyboard\" & Report.List(i).ImName & " - " & logFile) <> "" Then
                        ShellExecuteA 0, "open", "notepad.exe", DataPath & "\Logs\Keyboard\" & Report.List(i).ImName & " - " & logFile, "", SW_SHOW
                    End If
                End If
            End With
        End If
    Next
    
    If Page < Int(UBound(Report.List) / 5) Then
        SetClickArea2 GWW - 170, 220 + 10 * 30 + 15, 120, 30
        ProFont.DrawText Draw.Graphics, GWW - 170, 220 + 10 * 30 + 15, 120, 30, "Next Page >", IIf(IsMouseIn, argb(255, 42, 52, 0), argb(255, 129, 129, 129)), center, 18, Regular
        If IsClick Then Page = Page + 1
    End If
    
    ProFont.DrawText Draw.Graphics, 0, 220 + 10 * 30 + 15, GWW, 30, Page + 1 & " / " & Int(UBound(Report.List) / 5) + 1, argb(255, 129, 129, 129), center, 18, Regular
    
    If Page > 0 Then
        SetClickArea2 50, 220 + 10 * 30 + 15, 120, 30
        ProFont.DrawText Draw.Graphics, 50, 220 + 10 * 30 + 15, 120, 30, "< Last Page", IIf(IsMouseIn, argb(255, 42, 52, 0), argb(255, 129, 129, 129)), center, 18, Regular
        If IsClick Then Page = Page - 1
    End If
    
    Draw.Draw ProDraw.CompatibleDC, 0, 0
End Sub
Sub Display()
    If Mode = 1 Then Call Display2: Exit Sub

    BackImg.ImageByIndex(1).Draw Draw.CompatibleDC, 0, 0
    
    ProFont.DrawText Draw.Graphics, 0, 110, GWW, 30, "Monitor", argb(255, 42, 52, 0), center, 25, Regular
    ProFont.DrawText Draw.Graphics, 0, 150, GWW, 30, "Record all your actions .", argb(125, 42, 52, 0), center, 18, Regular
    
    SetClickArea2 20, 65, 20, 20
    ProFont.DrawText Draw.Graphics, 12, 65, 30, 30, "<", IIf(IsMouseIn, argb(255, 0, 176, 240), argb(255, 129, 129, 129)), center, 18, Regular
    If IsClick Then
        ProCore.FadePage "MainPage"
    End If
    
    SetClickArea2 30, 230, GWW - 60, 30
    BackImg.ImageByIndex(IIf(IsMouseIn, 2, 1)).DrawClip Draw.CompatibleDC, CtrlX, CtrlY, 0, 0, CtrlW, CtrlH
    ProFont.DrawText Draw.Graphics, 30, 230, GWW - 60, 30, "Breaker Logs", argb(255, 129, 129, 129), near, 18, Regular
    If IsClick Then
        ShellExecuteA 0, "open", "explorer.exe", DataPath & "\Logs\Breaker\", "", SW_SHOW
    End If
    
    SetClickArea2 30, 270, GWW - 60, 30
    BackImg.ImageByIndex(IIf(IsMouseIn, 2, 1)).DrawClip Draw.CompatibleDC, CtrlX, CtrlY, 0, 0, CtrlW, CtrlH
    ProFont.DrawText Draw.Graphics, 30, 270, GWW - 60, 30, "Monitor Logs", argb(255, 129, 129, 129), near, 18, Regular
    If IsClick Then
        ShellExecuteA 0, "open", "explorer.exe", DataPath & "\Logs\Monitor\", "", SW_SHOW
    End If
    
    SetClickArea2 30, 310, GWW - 60, 30
    BackImg.ImageByIndex(IIf(IsMouseIn, 2, 1)).DrawClip Draw.CompatibleDC, CtrlX, CtrlY, 0, 0, CtrlW, CtrlH
    ProFont.DrawText Draw.Graphics, 30, 310, GWW - 60, 30, "Tool Logs", argb(255, 129, 129, 129), near, 18, Regular
    If IsClick Then
        ShellExecuteA 0, "open", "explorer.exe", DataPath & "\Logs\Tools\", "", SW_SHOW
    End If
    
    SetClickArea2 30, 350, GWW - 60, 30
    BackImg.ImageByIndex(IIf(IsMouseIn, 2, 1)).DrawClip Draw.CompatibleDC, CtrlX, CtrlY, 0, 0, CtrlW, CtrlH
    ProFont.DrawText Draw.Graphics, 30, 350, GWW - 60, 30, "Visual Monitor Report", argb(255, 129, 129, 129), near, 18, Regular
    If IsClick Then
        LoadReport filename
        Mode = 1
        ProCore.FadePage mNowShow
    End If
    
    SetClickArea2 30, 460, GWW - 60, 30
    BackImg.ImageByIndex(IIf(IsMouseIn, 2, 1)).DrawClip Draw.CompatibleDC, CtrlX, CtrlY, 0, 0, CtrlW, CtrlH
    ProFont.DrawText Draw.Graphics, 30, 460, GWW - 60, 30, "Delete all logs", argb(255, 241, 77, 31), near, 18, Regular
    
    Dim DelCount As Long
    If IsClick Then
        If MsgBox("Really?", 48 + vbYesNo, "Cleanup") = vbNo Then Exit Sub
        DelCount = DelCount + Cleanup(0, DataPath & "\Logs\")
        DelCount = DelCount + Cleanup(0, DataPath & "\Logs\Breaker\")
        DelCount = DelCount + Cleanup(0, DataPath & "\Logs\Monitor\")
        DelCount = DelCount + Cleanup(0, DataPath & "\Logs\Tools\")
        DelCount = DelCount + Cleanup(0, DataPath & "\Monitor\")
        DelCount = DelCount + Cleanup(0, DataPath & "\Logs\Keyboard\")
        MsgBox "Success." & vbCrLf & vbCrLf & DelCount & " logs deleted .", 64, "Cleanup"
    End If
    SetClickArea2 30, 500, GWW - 60, 30
    BackImg.ImageByIndex(IIf(IsMouseIn, 2, 1)).DrawClip Draw.CompatibleDC, CtrlX, CtrlY, 0, 0, CtrlW, CtrlH
    ProFont.DrawText Draw.Graphics, 30, 500, GWW - 60, 30, "Reserve recent 7 days' logs", argb(255, 241, 77, 31), near, 18, Regular
    If IsClick Then
        If MsgBox("Really?", 48 + vbYesNo, "Cleanup") = vbNo Then Exit Sub
        DelCount = DelCount + Cleanup(7, DataPath & "\Logs\")
        DelCount = DelCount + Cleanup(7, DataPath & "\Logs\Breaker\")
        DelCount = DelCount + Cleanup(7, DataPath & "\Logs\Monitor\")
        DelCount = DelCount + Cleanup(7, DataPath & "\Logs\Tools\")
        DelCount = DelCount + Cleanup(7, DataPath & "\Monitor\")
        DelCount = DelCount + Cleanup(7, DataPath & "\Logs\Keyboard\")
        MsgBox "Success." & vbCrLf & vbCrLf & DelCount & " logs deleted .", 64, "Cleanup"
    End If
    SetClickArea2 30, 540, GWW - 60, 30
    BackImg.ImageByIndex(IIf(IsMouseIn, 2, 1)).DrawClip Draw.CompatibleDC, CtrlX, CtrlY, 0, 0, CtrlW, CtrlH
    ProFont.DrawText Draw.Graphics, 30, 540, GWW - 60, 30, "Reserve recent 1 month's logs", argb(255, 241, 77, 31), near, 18, Regular
    If IsClick Then
        If MsgBox("Really?", 48 + vbYesNo, "Cleanup") = vbNo Then Exit Sub
        DelCount = DelCount + Cleanup(30, DataPath & "\Logs\")
        DelCount = DelCount + Cleanup(30, DataPath & "\Logs\Breaker\")
        DelCount = DelCount + Cleanup(30, DataPath & "\Logs\Monitor\")
        DelCount = DelCount + Cleanup(30, DataPath & "\Logs\Tools\")
        DelCount = DelCount + Cleanup(30, DataPath & "\Monitor\")
        DelCount = DelCount + Cleanup(30, DataPath & "\Logs\Keyboard\")
        MsgBox "Success." & vbCrLf & vbCrLf & DelCount & " logs deleted .", 64, "Cleanup"
    End If
    
    Draw.Draw ProDraw.CompatibleDC, 0, 0
End Sub
Public Function Cleanup(ByVal Reserve As Single, ByVal Folder As String) As Long
    Dim FSO As Object, file As String
    If Right(Folder, 1) <> "\" Then Folder = Folder & "\"
    file = Dir(Folder)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Do While file <> ""
        If (Now - FSO.GetFile(Folder & file).DateLastModified) > Reserve Then
            Kill Folder & file
            Cleanup = Cleanup + 1
        End If
        file = Dir()
        DoEvents
    Loop
End Function
Private Sub Class_Initialize()
    Draw.Create TargetDC, GWW, GWH
    filename = DataPath & "\Monitor\" & year(Now) & "年" & Month(Now) & "月" & Day(Now) & "日.db"
    If Dir(filename) = "" Then
        ReDim Records.List(0)
        Open filename For Binary As #1
        Put #1, , Records
        Close #1
    Else
        Open filename For Binary As #1
        Get #1, , Records
        Close #1
    End If
    
    FileName3 = filename
End Sub

Sub LoadReport(ByVal file As String)
    Dim List As RecordList, List2 As RecordList, Find As Boolean
    Dim SortMap() As Long, Wait As Long
    
    Open file For Binary As #1
    Get #1, , List
    Close #1
    
    ReDim List2.List(0)
    
    For i = 1 To UBound(List.List)
        If List.List(i).ImName = "applicationframehost.exe" Then List.List(i).ImName = "UWP Application": List.List(i).Path = "C:\Windows\Explorer.exe"
        Find = False
        For s = 1 To UBound(List2.List)
            If List2.List(s).Path = List.List(i).Path Then
                List2.List(s).Time = List2.List(s).Time + List.List(i).Time / 1000
                Find = True
                Exit For
            End If
        Next
        If Find = False Then
            ReDim Preserve List2.List(UBound(List2.List) + 1)
            List2.List(UBound(List2.List)) = List.List(i)
            List2.List(UBound(List2.List)).Time = List2.List(UBound(List2.List)).Time / 1000
        End If
    Next
    
    TotalTime = 0
    
    ReDim SortMap(UBound(List2.List))
    For i = 1 To UBound(List2.List)
        SortMap(i) = i
        TotalTime = TotalTime + List2.List(i).Time
    Next
    
    For i = 1 To UBound(List2.List)
        For s = 1 To UBound(List2.List)
            If s <> i Then
                If List2.List(SortMap(s)).Time < List2.List(SortMap(i)).Time Then
                    Wait = SortMap(i)
                    SortMap(i) = SortMap(s)
                    SortMap(s) = Wait
                End If
            End If
        Next
    Next
    
    ReDim List.List(UBound(List2.List))
    For i = 1 To UBound(List2.List)
        List.List(i) = List2.List(SortMap(i))
    Next
    
    Report = List
End Sub

