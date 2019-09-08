Attribute VB_Name = "Core"
Option Explicit
Public Declare Function ChangeWindowMessageFilterEx Lib "user32" (ByVal Hwnd As Long, ByVal Message As Long, ByVal action As Long, PCHANGEFILTERSTRUCT As Any) As Long
Public Declare Function ChangeWindowMessageFilter Lib "user32" (ByVal Message As Long, ByVal dwFlag As Long) As Long
Public Const MSGFLT_ALLOW = 1
Public Const MSGFLT_DISALLOW = 2
Public Const MSGFLT_RESET = 0
Public Const MSGFLT_ADD = 1
Public Const MSGFLT_REMOVE = 2
Public Const WM_COPYGLOBALDATA = 73
Public Const WM_DROPFILES = &H233

Public BackImg As New ImageCollection, IconImg As New ImageCollection, CtrlImg As New ImageCollection, UIImg As New ImageCollection

Public MainPage As MainPage, PassPage As PassPage, HandlePage As HandlePage, BreakerPage As BreakerPage, MonPage As MonPage, BootPage As BootPage, RegPage As RegPage, DeskPage As DeskPage, BatchPage As BatchPage
Public Password As String, Logined As Boolean

Public APIShell As New APIShell

Public FIcon As New FileIcons

Sub Sort(a() As String)
    Dim i As Integer, j As Integer, s As Integer, c As String
    Dim t As Long, t2 As Long
    
    For i = UBound(a) To 0 Step -1
        For j = UBound(a) To i Step -1
            For s = 1 To Len(a(i))
                If s <= Len(t2) Then
                    t = Asc(Mid(a(i), s, 1)): t2 = Asc(Mid(a(j), s, 1))
                    If t > t2 Then
                        c = a(i): a(i) = a(j): a(j) = c
                    End If
                Else: Exit For
                End If
            Next
        Next
    Next
End Sub
Public Function UnSpace(ByVal Str As String) As String
    If InStr(Str, Chr(0)) <> 0 Then
        UnSpace = Left(Str, InStr(Str, Chr(0)) - 1)
    Else
        UnSpace = Str
    End If
End Function
Sub Log(ByVal Func As String, ByVal Text As String)
    On Error Resume Next
    
    Open DataPath & "\Logs\" & Func & "\" & year(Now) & "年" & Month(Now) & "月" & Day(Now) & "日   " & Hour(Now) & "时.txt" For Append As #1
    Print #1, Now & "    " & Text
    Close #1
    
    If Err.Number <> 0 Then Err.Clear
End Sub
Public Function GetProcessPath(Hwnd As Long) As String
    On Error GoTo z
    
recheck:
    
    Dim PID As Long, Class As String * 255
    Dim cbNeeded As Long, szBuf(1 To 250) As Long, ret As Long, szPathName As String, nSize As Long, hProcess As Long
    
    Class = "": PID = 0
    
    GetWindowThreadProcessId Hwnd, PID
    GetClassNameA Hwnd, Class, 255
    
    If UnSpace(Class) = "ApplicationFrameWindow" And Hwnd <> 0 Then 'UWP
        Hwnd = uwpFind(Hwnd)
        GoTo recheck
    End If
    
    hProcess = OpenProcess(&H400 Or &H10, 0, PID)
    If hProcess <> 0 Then
        szPathName = Space(260): nSize = 500
        ret = GetModuleFileNameExA(hProcess, szBuf(1), szPathName, nSize)
        GetProcessPath = Left(szPathName, ret)
    End If
    
    ret = CloseHandle(hProcess)
    If GetProcessPath = "" Then
        GetProcessPath = "System"
    End If
    
    Exit Function
z:
End Function
Sub LoadAssets()
    Dim ProScale As Single
    ProScale = 0.8
    BackImg.LoadDir ProPath & "\assets\Back\", ProScale
    IconImg.LoadDir ProPath & "\assets\Icons\", ProScale
    CtrlImg.LoadDir ProPath & "\assets\Controls\", ProScale
    UIImg.LoadDir ProPath & "\assets\UI\", ProScale
    
    Set ProFont = New Fonts
    ProFont.RegFont ProPath & "\font.ttf"
    ProFont.Create "Selawik"
    ProFont.SetFont
    
    Set MainPage = New MainPage
    Set PassPage = New PassPage
    Set HandlePage = New HandlePage
    Set BreakerPage = New BreakerPage
    Set MonPage = New MonPage
    Set BootPage = New BootPage
    Set RegPage = New RegPage
    Set BatchPage = New BatchPage
    
    ProCore.AddScreen MainPage
    ProCore.AddScreen PassPage
    ProCore.AddScreen HandlePage
    ProCore.AddScreen BreakerPage
    ProCore.AddScreen MonPage
    ProCore.AddScreen BootPage
    ProCore.AddScreen RegPage
    ProCore.AddScreen BatchPage
    
    mNowShow = "MainPage"
    
    DrawCloseButton = True
End Sub
Sub LoadDeskAssets()
    Dim ProScale As Single
    ProScale = 1
    BackImg.LoadDir ProPath & "\assets\Back\", ProScale
    IconImg.LoadDir ProPath & "\assets\Icons\", ProScale
    CtrlImg.LoadDir ProPath & "\assets\Controls\", ProScale
    UIImg.LoadDir ProPath & "\assets\UI\", ProScale
    
    Set ProFont = New Fonts
    'ProFont.RegFont ProPath & "\font.ttf"
    ProFont.Create "微软雅黑"
    ProFont.SetFont
    
    Set DeskPage = New DeskPage
    
    ProCore.AddScreen DeskPage
    
    mNowShow = "DeskPage"
    
    DrawCloseButton = False
End Sub
