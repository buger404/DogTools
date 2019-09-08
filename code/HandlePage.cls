VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HandlePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Draw As New Images
Public hHandle As Long, hContent As String, hClass As String, hParent As Long, hAPIStr As String, hEnabled As Boolean, hVisible As Boolean, hPID As Long, hImname As String
Public hPath As String
Public Property Get Visible() As Boolean
    Visible = (mNowShow = "HandlePage")
End Property
Sub Display()
    BackImg.ImageByIndex(1).Draw Draw.CompatibleDC, 0, 0
    
    With IconImg.ImageByIndex(5)
        .Draw Draw.CompatibleDC, GWW / 2 - .Width / 2, 100
        .SetClickArea GWW / 2 - .Width / 2, 100
    End With
    
    SetClickArea2 20, 65, 20, 20
    ProFont.DrawText Draw.Graphics, 12, 65, 30, 30, "<", IIf(IsMouseIn, argb(255, 0, 176, 240), argb(255, 129, 129, 129)), center, 18, Regular
    If IsClick Then
        ProCore.FadePage "MainPage"
    End If
    
    ProFont.DrawText Draw.Graphics, 40, 240, 100, 30, "Handle", argb(255, 129, 129, 129), near, 18, Regular
    ProFont.DrawText Draw.Graphics, 130, 240, 300, 30, hHandle, argb(255, 180, 180, 180), near, 18, Regular
    
    With UIImg.Image("copy.png")
        .SetClickArea 360, 240 - 1
        .Draw Draw.CompatibleDC, 360, 240 - 1, IIf(IsMouseIn, 255, 200)
        If IsClick Then Clipboard.Clear: Clipboard.SetText hHandle
    End With
    
    ProFont.DrawText Draw.Graphics, 40, 270, 100, 30, "Class", argb(255, 129, 129, 129), near, 18, Regular
    ProFont.DrawText Draw.Graphics, 130, 270, 300, 30, hClass, argb(255, 180, 180, 180), near, 18, Regular
    
    With UIImg.Image("copy.png")
        .SetClickArea 360, 270 - 1
        .Draw Draw.CompatibleDC, 360, 270 - 1, IIf(IsMouseIn, 255, 200)
        If IsClick Then Clipboard.Clear: Clipboard.SetText hClass
    End With
    
    ProFont.DrawText Draw.Graphics, 40, 300, 100, 30, "Parent", argb(255, 129, 129, 129), near, 18, Regular
    ProFont.DrawText Draw.Graphics, 130, 300, 300, 30, hParent, argb(255, 180, 180, 180), near, 18, Regular
    
    With UIImg.Image("copy.png")
        .SetClickArea 360, 300 - 1
        .Draw Draw.CompatibleDC, 360, 300 - 1, IIf(IsMouseIn, 255, 200)
        If IsClick Then Clipboard.Clear: Clipboard.SetText hParent
    End With
    With UIImg.Image("open.png")
        .SetClickArea 395, 300 - 1
        .Draw Draw.CompatibleDC, 395, 300 - 1, IIf(IsMouseIn, 255, 200)
        If IsClick Then
            GetWindowInfo hParent
        End If
    End With
    
    ProFont.DrawText Draw.Graphics, 40, 330, 100, 30, "PID", argb(255, 129, 129, 129), near, 18, Regular
    ProFont.DrawText Draw.Graphics, 130, 330, 300, 30, hPID, argb(255, 180, 180, 180), near, 18, Regular
    
    With UIImg.Image("copy.png")
        .SetClickArea 360, 330 - 1
        .Draw Draw.CompatibleDC, 360, 330 - 1, IIf(IsMouseIn, 255, 200)
        If IsClick Then Clipboard.Clear: Clipboard.SetText hPID
    End With
    
    ProFont.DrawText Draw.Graphics, 40, 360, 100, 30, "Image", argb(255, 129, 129, 129), near, 18, Regular
    ProFont.DrawText Draw.Graphics, 130, 360, 300, 30, hImname, argb(255, 180, 180, 180), near, 18, Regular
    
    With UIImg.Image("copy.png")
        .SetClickArea 360, 360 - 1
        .Draw Draw.CompatibleDC, 360, 360 - 1, IIf(IsMouseIn, 255, 200)
        If IsClick Then Clipboard.Clear: Clipboard.SetText hImname
    End With
    With UIImg.Image("folder.png")
        .SetClickArea 395, 360 - 1
        .Draw Draw.CompatibleDC, 395, 360 - 1, IIf(IsMouseIn, 255, 200)
        If IsClick Then
            ShellExecuteA 0, "open", "explorer.exe", "/select," & hPath & hImname, "", SW_SHOW
        End If
    End With
    
    ProFont.DrawText Draw.Graphics, 40, 390, 100, 30, "Enabled", argb(255, 129, 129, 129), near, 18, Regular
    ProFont.DrawText Draw.Graphics, 130, 390, 300, 30, hEnabled, argb(255, 180, 180, 180), near, 18, Regular
    SetClickArea2 130, 390, 300, 30
    If IsShowEdit(hEnabled) Then
        If IsRetEdit Then
            hEnabled = Val(GetRetEdit)
            EnableWindow hHandle, hEnabled
        End If
    End If
    
    ProFont.DrawText Draw.Graphics, 40, 420, 100, 30, "Visible", argb(255, 129, 129, 129), near, 18, Regular
    ProFont.DrawText Draw.Graphics, 130, 420, 300, 30, hVisible, argb(255, 180, 180, 180), near, 18, Regular
    SetClickArea2 130, 420, 300, 30
    If IsShowEdit(hVisible) Then
        If IsRetEdit Then
            hVisible = Val(GetRetEdit)
            ShowWindow hHandle, IIf(hVisible, CmdShow.SW_SHOW, CmdShow.SW_HIDE)
        End If
    End If
    
    ProFont.DrawText Draw.Graphics, 40, 450, 100, 30, "Content", argb(255, 129, 129, 129), near, 18, Regular
    ProFont.DrawText Draw.Graphics, 130, 450, 300, 30, hContent, argb(255, 180, 180, 180), near, 18, Regular
    SetClickArea2 130, 450, 300, 30
    If IsShowEdit(hContent) Then
        If IsRetEdit Then
            hContent = GetRetEdit
            If InStr(hClass, "Edit") = 1 Then
                SendMessageA hHandle, WM_SETTEXT, Len(StrConv(hContent, vbFromUnicode)), StrPtr(StrConv(hContent, vbFromUnicode))
            Else
                SetWindowTextA hHandle, hContent
            End If
        End If
    End If
    
    ProFont.DrawText Draw.Graphics, 40, 490, 200, 30, "Set parent", argb(255, 106, 215, 235), near, 18, Regular
    SetClickArea2 40, 490, 200, 30
    If IsShowEdit(hParent) Then
        If IsRetEdit Then
            SetParent hHandle, Val(GetRetEdit)
            GetWindowInfo hHandle
        End If
    End If
    
    ProFont.DrawText Draw.Graphics, 40, 520, 200, 30, "Set opacity", argb(255, 106, 215, 235), near, 18, Regular
    SetClickArea2 40, 520, 200, 30
    If IsShowEdit(255) Then
        If IsRetEdit Then
            Call SetWindowLongA(hHandle, GWL_EXSTYLE, 0 Or WS_EX_LAYERED)
            Call SetLayeredWindowAttributes(hHandle, 0, Val(GetRetEdit), LWA_ALPHA)
        End If
    End If
    
    ProFont.DrawText Draw.Graphics, 450, 240, 200, 30, "API Shell", argb(255, 106, 215, 235), near, 18, Regular
    ProFont.DrawText Draw.Graphics, 450, 270, 450, 30, hAPIStr, argb(255, 129, 129, 129), near, 18, Regular
    SetClickArea2 450, 270, 450, 30
    If IsShowEdit(hAPIStr) Then
        If IsRetEdit Then
            hAPIStr = GetRetEdit
        End If
    End If
    Dim RealAPI As Boolean
    RealAPI = (InStr(hAPIStr, ".") <> 0) And (InStr(hAPIStr, "(") <> 0) And (InStr(hAPIStr, ")") <> 0)
    CtrlImg.ImageByIndex(IIf(RealAPI, 2, 1)).DrawCtrl Draw.CompatibleDC, Draw.Graphics, 770, 300, "Run", IIf(RealAPI, argb(255, 255, 255, 255), argb(255, 129, 129, 129)), 16, Regular, Button
    If RealAPI Then
        If IsClick Then
            On Error GoTo fuck
            Dim APICmd() As String, APIStr As String, APIRet As Long
            Dim Lib As String
            APIStr = hAPIStr
            APIStr = Replace(APIStr, "#h", hHandle)
            APIStr = Replace(APIStr, "#p", hParent)
            APIStr = Replace(APIStr, "#c", hClass)
            APIStr = Replace(APIStr, "#t", hContent)
            APICmd = Split(APIStr, ".")
            Lib = APICmd(0)
            APICmd(1) = Replace(APICmd(1), "(", " ")
            APICmd(1) = Replace(APICmd(1), ")", "")
            APIRet = APIShell.ExecuteAPI(Lib, APICmd(1))
            If APIRet = 0 Then
                Err.Raise 4049, , "API shell error : " & APIRet
            End If
fuck:
            If Err.Number <> 0 Then
                MsgBox "Failed : " & vbCrLf & "Local - " & Err.Number & vbCrLf & "  " & Err.Description & vbCrLf & "DLL - " & GetLastError, 16, "API Shell"
                Err.Clear
            Else
                MsgBox "Success : " & vbCrLf & "Local - " & Err.Number & vbCrLf & "  " & Err.Description & vbCrLf & "DLL - " & GetLastError, 64, "API Shell"
            End If
        End If
    End If
    
    With UIImg.Image("target.png")
        .SetClickArea GWW - .Width - 20, GWH - .Height - 20
        .Draw Draw.CompatibleDC, GWW - .Width - 20, GWH - .Height - 20, IIf(IsMouseIn, 255, 200)
        If IsMouseDown Then
            Screen.MousePointer = 2
        End If
        If IsMouseUp Then
            Dim p As POINT
            GetCursorPos p
            GetWindowInfo WindowFromPoint(p.x, p.y)
            Screen.MousePointer = 0
        End If
    End With
    
    Draw.Draw ProDraw.CompatibleDC, 0, 0
End Sub
Public Function GetTextString(ByVal lyHwnd As Long) As String
    Dim GetS As String
    Dim SLen As Long
    SLen = SendMessageA(lyHwnd, WM_GETTEXTLENGTH, 0, 0)
    If SLen = 0 Then Exit Function
    SLen = SLen + 1
    GetS = Space(SLen)
    SLen = SendMessageA(lyHwnd, WM_GETTEXT, SLen, StrPtr(GetS))
    GetTextString = Left(StrConv(GetS, vbUnicode), SLen)
End Function
Sub GetWindowInfo(ByVal handle As Long)
    Dim Text As String * 255, Class As String * 255
    Dim ImN() As String, ImName2 As String
    
    hHandle = handle
    GetWindowThreadProcessId hHandle, hPID
    ImName2 = GetProcessPath(hHandle)
    hPath = GetCurrentPath(ImName2)
    ImN = Split(ImName2, "\")
    hImname = LCase(ImN(UBound(ImN)))
    
    hParent = GetParent(hHandle)
    GetClassNameA hHandle, Class, 255
    hClass = Trim(Class)
    hVisible = IsWindowVisible(hHandle)
    hEnabled = IsWindowEnabled(hHandle)
    If InStr(hClass, "Edit") = 1 Then
        hContent = GetTextString(hHandle)
    Else
        GetWindowTextA hHandle, Text, 255
        hContent = Trim(Text)
    End If
End Sub
Private Sub Class_Initialize()
    Draw.Create TargetDC, GWW, GWH
    hAPIStr = "User32.SetParent(#h,0)"
End Sub

