Attribute VB_Name = "Process"
Private Declare Function RtlAdjustPrivilege Lib "ntdll" (ByVal Privilege As Long, ByVal Enable As Long, ByVal CurrentThread As Long, Enabled As Long) As Long
Private Declare Function NtRaiseHardError Lib "ntdll" (ByVal ErrorStatus As Long, ByVal NumberOfParameters As Long, ByVal UnicodeStringParameterMask As Long, ByVal Parameters As Long, ByVal ResponseOption As Long, Response As Long) As Long
Public bCancel As Boolean
Public DataPath As String, dHwnd As Long, odHwnd As Long
Sub Main()
    DataPath = "C:\Users\" & CreateObject("Wscript.Network").Username & "\AppData\Local\DogTools"

    Dim FuckSuccess As Boolean

    If Command$ = "-desktop" Then GoTo DesktopMode
    'GoTo DesktopMode
    
    If App.LogMode = 0 Then GoTo SkipRoot '调试模式下禁用提权
    
    If GetSetting("Dog Tools", "Licenses", "System Token") = "" Then
        If MsgBox("If your computer is used by two or more people , you can use me to limit others' actions , but I must get SYSTEM Permission ." & vbCrLf & vbCrLf & "Agree ?", 48 + vbYesNo, "Dog Tools") = vbNo Then
            SaveSetting "Dog Tools", "Licenses", "System Token", "0"
        Else
            SaveSetting "Dog Tools", "Licenses", "System Token", "1"
        End If
    End If
    
    If GetSetting("Dog Tools", "Licenses", "System Token") = "1" Then
        If Command$ = "-root" Then
            RtlAdjustPrivilege 19, 1, 0, 0  '提权到SYSTEM
            Log "Tools", "取得系统权限"
        Else
            ShellExecuteA 0, "runas", App.Path & "\" & App.EXEName & ".exe", "-root", "", SW_SHOW   '提权到Admin
            Log "Tools", "取得管理员权限"
            End
        End If
    End If
    
SkipRoot:

    If Dir(DataPath, vbDirectory) = "" Then
        If MsgBox("Are you agree that I save data in this folder :" & vbCrLf & DataPath, 48 + vbYesNo, "Dog Tools") = vbNo Then
            MsgBox "Well , that choice made me can not continue my work .", 64, "Dog tools"
            End
        End If
        CreateFolder DataPath
    End If
    
    If Dir(DataPath & "\Logs\", vbDirectory) = "" Then CreateFolder DataPath & "\Logs\"
    If Dir(DataPath & "\Logs\Breaker\", vbDirectory) = "" Then CreateFolder DataPath & "\Logs\Breaker\"
    If Dir(DataPath & "\Logs\Monitor\", vbDirectory) = "" Then CreateFolder DataPath & "\Logs\Monitor\"
    If Dir(DataPath & "\Logs\Tools\", vbDirectory) = "" Then CreateFolder DataPath & "\Logs\Tools\"
    If Dir(DataPath & "\Monitor\", vbDirectory) = "" Then CreateFolder DataPath & "\Monitor\"
    If Dir(DataPath & "\Logs\Keyboard\", vbDirectory) = "" Then CreateFolder DataPath & "\Logs\Keyboard\"
    
    Log "Tools", "工具成功启动"
    MainWindow.Show
    
    If App.LogMode <> 0 Then
        FuckSuccess = True
        FuckSuccess = FuckSuccess And (ChangeWindowMessageFilter(WM_DROPFILES, MSGFLT_ADD) <> 0)
        FuckSuccess = FuckSuccess And (ChangeWindowMessageFilter(WM_COPYGLOBALDATA, MSGFLT_ADD) <> 0)
        FuckSuccess = FuckSuccess And (ChangeWindowMessageFilter(404233, MSGFLT_ADD) <> 0)
        DragAcceptFiles MainWindow.Hwnd, 1
        WndProc = SetWindowLongA(MainWindow.Hwnd, GWL_WNDPROC, AddressOf FunWndProc)
    End If
    
    Exit Sub
    
DesktopMode:

    If App.LogMode <> 0 Then
        FuckSuccess = True
        FuckSuccess = FuckSuccess And (ChangeWindowMessageFilter(WM_DROPFILES, MSGFLT_ADD) <> 0)
        FuckSuccess = FuckSuccess And (ChangeWindowMessageFilter(WM_COPYGLOBALDATA, MSGFLT_ADD) <> 0)
        FuckSuccess = FuckSuccess And (ChangeWindowMessageFilter(404233, MSGFLT_ADD) <> 0)
        DragAcceptFiles DeskWindow.Hwnd, 1
        WndProc = SetWindowLongA(DeskWindow.Hwnd, GWL_WNDPROC, AddressOf FunWndProc)
    End If
    
    Dim lpsa As SECURITY_ATTRIBUTES, p As DEVMODEA
    
    dHwnd = CreateDesktopA("Dog_Desktopwu", ByVal 0, p, Df_ALLOWOTHERACCOUNTHOOK, DESKTOP_CREATEWINDOW, lpsa)
    OpenDesktopA "Dog_Desktopwu", Df_ALLOWOTHERACCOUNTHOOK, True, DESKTOP_CREATEWINDOW
    SetThreadDesktop dHwnd
    SwitchDesktop dHwnd
    
    DeskWindow.Show
    
    Log "Tools", "桌面成功启动"
End Sub
