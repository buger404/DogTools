Attribute VB_Name = "WinSockFun"
Public WebPath As String

Sub SendText(Obj As Winsock, ByVal SData As String)
    Dim data() As Byte
    data = SData
    Obj.SendData "HTTP/1.1 200 OK" & vbCrLf & _
                                  "Date: " & Weekday(Now) & ", " & Day(Now) & ", " & MonthName(Month(Now)) & " " & year(Now) & " " & _
                                                  format(Hour(Now), "00") & ":" & format(Minute(Now), "00") & ":" & format(Second(Now), "00") & " GMT" & vbCrLf & _
                                    "Connection: keep-alive" & vbCrLf & _
                                    "Content-Type: text/html" & vbCrLf & _
                                    "Content-Length: " & (UBound(data) + 1) & vbCrLf & vbCrLf & SData
End Sub
Sub SendFile(Obj As Winsock, ByVal Path As String)
    Dim data() As Byte, ret As String, FileFormat As String, File As String, temp() As String
    If FileLen(Path) <> 0 Then
        ReDim data(FileLen(Path) - 1)
        Open Path For Binary As #1
        Get #1, , data
        Close #1
    Else
        If Dir(WebPath & "\404.htm") <> "" Then SendFile Obj, WebPath & "\404.htm"
        Exit Sub
    End If
    
    File = LCase(Split(Path, "\")(UBound(Split(Path, "\"))))
    temp = Split(File, ".")
    
    On Error Resume Next
    FileFormat = CreateObject("wscript.shell").regread("HKEY_CLASSES_ROOT\." & temp(UBound(temp)) & "\Content Type")
    
    If FileFormat = "" Then FileFormat = "*/*"
    
    ret = "HTTP/1.1 200 OK" & vbCrLf & _
            "Date: " & Weekday(Now) & ", " & Day(Now) & ", " & MonthName(Month(Now)) & " " & year(Now) & " " & _
                            format(Hour(Now), "00") & ":" & format(Minute(Now), "00") & ":" & format(Second(Now), "00") & " GMT" & vbCrLf & _
            "Content-Type: " & FileFormat & vbCrLf & _
            "Content-Length: " & (UBound(data) + 1) & vbCrLf & _
            "Connection: keep-alive" & vbCrLf & _
            "Accept-Ranges: bytes" & vbCrLf & vbCrLf

    Obj.SendData ret
    Obj.SendData data

End Sub

Sub SendFileWithReplace(Obj As Winsock, ByVal Path As String, ParamArray rep())
    Dim reps As String, ret As String
    If FileLen(Path) <> 0 Then
        Open Path For Input As #1
        Do While Not EOF(1)
            Line Input #1, ret
            reps = reps & ret & vbCrLf
        Loop
        Close #1
        For I = 0 To UBound(rep)
            reps = Replace(reps, "*[var" & I & "]*", rep(I))
        Next
        Open WebPath & "\temp.htm" For Output As #1
        Print #1, reps
        Close #1
        SendFile Obj, WebPath & "\temp.htm"
    Else
        If Dir(WebPath & "\404.htm") <> "" Then SendFile Obj, WebPath & "\404.htm"
        Exit Sub
    End If
    
End Sub

Public Function URLDecode(ByRef strURL As String) As String
    Dim I As Long
 
    If InStr(strURL, "%") = 0 Then URLDecode = strURL: Exit Function
 
    For I = 1 To Len(strURL)
        If Mid(strURL, I, 1) = "%" Then
            If Val("&H" & Mid(strURL, I + 1, 2)) > 127 Then
                URLDecode = URLDecode & Chr(Val("&H" & Mid(strURL, I + 1, 2) & Mid(strURL, I + 4, 2)))
                I = I + 5
            Else
                URLDecode = URLDecode & Chr(Val("&H" & Mid(strURL, I + 1, 2)))
                I = I + 2
            End If
        Else
            URLDecode = URLDecode & Mid(strURL, I, 1)
        End If
    Next
End Function
