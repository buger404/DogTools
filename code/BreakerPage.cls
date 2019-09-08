VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BreakerPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Draw As New Images
Dim Page As Integer
Private Type BreakInfo
    ClassName As String
    Title As String
    ImName As String
    data(4) As String
End Type
Private Type BreakList
    List() As BreakInfo
End Type
Dim Breaks As BreakList
Dim AddL As BreakInfo
Dim LastHwnd As Long, LastTitle As String * 255
Public Property Get Visible() As Boolean
    Visible = (mNowShow = "BreakerPage")
End Property
Sub Carry()
    If Logined Then Exit Sub

    Dim Title As String * 255, Class As String * 255, Hwnd As Long, PID As Long
    Dim Title2 As String, Class2 As String
    Dim ImN() As String, ImName2 As String, ImName As String
    Hwnd = GetForegroundWindow
    
    If LastHwnd = Hwnd Then
        GetWindowTextA Hwnd, Title, 255
        If LastTitle <> Title Then
            LastTitle = Title
            GoTo Continue
        End If
        Exit Sub
    End If
    
Continue:
    LastHwnd = Hwnd
    
    GetWindowTextA Hwnd, Title, 255
    GetClassNameA Hwnd, Class, 255
    LastTitle = Title
    
    ImName2 = GetProcessPath(Hwnd)
    ImN = Split(ImName2, "\")
    ImName = LCase(ImN(UBound(ImN)))

    Title2 = LCase(UnSpace(Title))
    Class2 = LCase(UnSpace(Class))

    For i = 1 To UBound(Breaks.List)

        With Breaks.List(i)
            If (Title2 Like LCase(.Title)) And (Class2 Like LCase(.ClassName)) And (ImName Like LCase(.ImName)) Then
                CloseWindow Hwnd '最小化！
                ShowWindow Hwnd, SW_HIDE '藏起来！
                DestroyWindow Hwnd '关掉！
                SendMessageA Hwnd, WM_CLOSE, 0, 0 '必杀技！
                SaveSetting "Dog Tools", "Data", "Breaker", Val(GetSetting("Dog Tools", "Data", "Breaker")) + 1
                Log "Breaker", "成功阻止运行：" & .Title & " | " & .ClassName & " | " & .ImName
                LastHwnd = 0: LastTitle = ""
            End If
        End With
        
    Next
End Sub
Sub Display()
    BackImg.ImageByIndex(1).Draw Draw.CompatibleDC, 0, 0
    
    ProFont.DrawText Draw.Graphics, 0, 110, GWW, 30, "Breaker", argb(255, 241, 77, 31), center, 25, Regular
    ProFont.DrawText Draw.Graphics, 0, 150, GWW, 30, "Total number of blocked : " & Val(GetSetting("Dog Tools", "Data", "Breaker")), argb(125, 241, 77, 31), center, 18, Regular
    
    SetClickArea2 20, 65, 20, 20
    ProFont.DrawText Draw.Graphics, 12, 65, 30, 30, "<", IIf(IsMouseIn, argb(255, 0, 176, 240), argb(255, 129, 129, 129)), center, 18, Regular
    If IsClick Then
        ProCore.FadePage "MainPage"
    End If
    
    BackImg.ImageByIndex(2).DrawClip Draw.CompatibleDC, 0, GWH - 120, 0, 0, GWW, 120
    
    ProFont.DrawText Draw.Graphics, 30, GWH - 100, 200, 30, "Title", argb(255, 129, 129, 129), near, 18, Regular
    ProFont.DrawText Draw.Graphics, 30, GWH - 70, 200, 30, AddL.Title, argb(255, 169, 169, 169), near, 18, Regular
    SetClickArea2 30, GWH - 70, 200, 30
    If IsShowEdit(AddL.Title) Then
        If IsRetEdit Then
            AddL.Title = GetRetEdit
        End If
    End If
    
    ProFont.DrawText Draw.Graphics, 30 + 230, GWH - 100, 200, 30, "ClassName", argb(255, 129, 129, 129), near, 18, Regular
    ProFont.DrawText Draw.Graphics, 30 + 230, GWH - 70, 200, 30, AddL.ClassName, argb(255, 169, 169, 169), near, 18, Regular
    SetClickArea2 30 + 230, GWH - 70, 200, 30
    If IsShowEdit(AddL.ClassName) Then
        If IsRetEdit Then
            AddL.ClassName = GetRetEdit
        End If
    End If
    
    ProFont.DrawText Draw.Graphics, 30 + 230 * 2, GWH - 100, 200, 30, "ImageName", argb(255, 129, 129, 129), near, 18, Regular
    ProFont.DrawText Draw.Graphics, 30 + 230 * 2, GWH - 70, 200, 30, AddL.ImName, argb(255, 169, 169, 169), near, 18, Regular
    SetClickArea2 30 + 230 * 2, GWH - 70, 200, 30
    If IsShowEdit(AddL.ImName) Then
        If IsRetEdit Then
            AddL.ImName = GetRetEdit
        End If
    End If
    
    Dim RealList As Boolean
    RealList = AddL.Title <> "" And AddL.ClassName <> "" And AddL.ImName <> ""
    CtrlImg.ImageByIndex(IIf(RealList, 2, 1)).DrawCtrl Draw.CompatibleDC, Draw.Graphics, GWW - 130, GWH - 55, "Limit", IIf(RealList, argb(255, 255, 255, 255), argb(255, 129, 129, 129)), 16, Regular, Button
    If IsClick Then
        ReDim Preserve Breaks.List(UBound(Breaks.List) + 1)
        Breaks.List(UBound(Breaks.List)) = AddL
        Open DataPath & "\break.db" For Binary As #1
        Put #1, , Breaks
        Close #1
    End If
    
    Dim Start As Long
    Start = Page * 7 + 1
    For i = Start To Start + 6
        If i <= UBound(Breaks.List) Then
            With Breaks.List(i)
                ProFont.DrawText Draw.Graphics, 50, 220 + (i - Start) * 30, 700, 30, "Limit:        " & .Title & " | " & .ClassName & " | " & .ImName, argb(255, 129, 129, 129), near, 18, Regular
                SetClickArea2 50 + 750, 220 + (i - Start) * 30, 100, 30
                ProFont.DrawText Draw.Graphics, 50 + 750, 220 + (i - Start) * 30, 100, 30, "Delete", IIf(IsMouseIn, argb(255, 241, 77, 31), argb(255, 129, 129, 129)), near, 18, Regular
            End With
            If IsClick Then
                For s = i To UBound(Breaks.List) - 1
                    Breaks.List(s) = Breaks.List(s + 1)
                Next
                ReDim Preserve Breaks.List(UBound(Breaks.List) - 1)
                Open DataPath & "\break.db" For Binary As #1
                Put #1, , Breaks
                Close #1
            End If
        End If
    Next
    
    If Page < Int(UBound(Breaks.List) / 7) Then
        SetClickArea2 GWW - 170, 220 + 7 * 30, 120, 30
        ProFont.DrawText Draw.Graphics, GWW - 170, 220 + 7 * 30, 120, 30, "Next Page >", IIf(IsMouseIn, argb(255, 241, 77, 31), argb(255, 129, 129, 129)), center, 18, Regular
        If IsClick Then Page = Page + 1
    End If
    
    ProFont.DrawText Draw.Graphics, 0, 220 + 7 * 30, GWW, 30, Page + 1 & " / " & Int(UBound(Breaks.List) / 7) + 1, argb(255, 129, 129, 129), center, 18, Regular
    
    If Page > 0 Then
        SetClickArea2 50, 220 + 7 * 30, 120, 30
        ProFont.DrawText Draw.Graphics, 50, 220 + 7 * 30, 120, 30, "< Last Page", IIf(IsMouseIn, argb(255, 241, 77, 31), argb(255, 129, 129, 129)), center, 18, Regular
        If IsClick Then Page = Page - 1
    End If
        
    With UIImg.Image("target.png")
        .SetClickArea GWW - .Width - 20, GWH - .Height - 130
        .Draw Draw.CompatibleDC, GWW - .Width - 20, GWH - .Height - 130, IIf(IsMouseIn, 255, 200)
        If IsMouseDown Then
            Screen.MousePointer = 2
        End If
        If IsMouseUp Then
            Dim p As Point
            GetCursorPos p
            HandlePage.GetWindowInfo WindowFromPoint(p.X, p.Y)
            With AddL
                .Title = UnSpace(HandlePage.hContent)
                .ImName = UnSpace(HandlePage.hImname)
                .ClassName = UnSpace(HandlePage.hClass)
            End With
            Screen.MousePointer = 0
        End If
    End With
        
    Draw.Draw ProDraw.CompatibleDC, 0, 0
End Sub
Private Sub Class_Initialize()
    Draw.Create TargetDC, GWW, GWH
    AddL.ImName = "..."
    AddL.Title = "..."
    AddL.ClassName = "..."
    If Dir(DataPath & "\break.db") = "" Then
        ReDim Breaks.List(0)
        Open DataPath & "\break.db" For Binary As #1
        Put #1, , Breaks
        Close #1
    Else
        Open DataPath & "\break.db" For Binary As #1
        Get #1, , Breaks
        Close #1
    End If
End Sub


