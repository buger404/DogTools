VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RegPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Draw As New Images
Dim Mode As Integer

Public Property Get Visible() As Boolean
    Visible = (mNowShow = "RegPage")
End Property
Sub FileDrop(ByVal FileList As String)
    Dim Files() As String
    Files = Split(FileList, vbCrLf)
    
    For i = 0 To UBound(Files) - 1
        If Mode = 0 Then
            ShellExecuteA Hwnd, "runas", "C:\Windows\System32\regsvr32.exe ", """" & Files(i) & """", 0, SW_SHOW
            Log "Tools", "×¢²áÁË " & Files(i)
        Else
            ShellExecuteA Hwnd, "runas", "C:\Windows\System32\regsvr32.exe ", "-u """ & Files(i) & """", 0, SW_SHOW
            Log "Tools", "·´×¢²áÁË " & Files(i)
        End If
    Next
End Sub
Sub Display()
    BackImg.ImageByIndex(1).Draw Draw.CompatibleDC, 0, 0
    
    ProFont.DrawText Draw.Graphics, 0, 110, GWW, 40, "Register", argb(255, 255, 172, 0), center, 25, Regular
    ProFont.DrawText Draw.Graphics, 0, 150, GWW, 30, "Drop your files here .", argb(125, 255, 172, 0), center, 18, Regular
    
    SetClickArea2 20, 65, 20, 20
    ProFont.DrawText Draw.Graphics, 12, 65, 30, 30, "<", IIf(IsMouseIn, argb(255, 0, 176, 240), argb(255, 129, 129, 129)), center, 18, Regular
    If IsClick Then
        ProCore.FadePage "MainPage"
    End If
    
    SetClickArea2 150, 230, GWW - 300, 30
    BackImg.ImageByIndex(IIf(IsMouseIn, 2, 1)).DrawClip Draw.CompatibleDC, CtrlX, CtrlY, 0, 0, CtrlW, CtrlH
    ProFont.DrawText Draw.Graphics, 150, 230, GWW - 300, 30, "Register", IIf(Mode = 0, argb(125, 255, 172, 0), argb(255, 129, 129, 129)), near, 18, Regular
    If IsClick Then Mode = 0
    
    SetClickArea2 150, 270, GWW - 300, 30
    BackImg.ImageByIndex(IIf(IsMouseIn, 2, 1)).DrawClip Draw.CompatibleDC, CtrlX, CtrlY, 0, 0, CtrlW, CtrlH
    ProFont.DrawText Draw.Graphics, 150, 270, GWW - 300, 30, "Unregister", IIf(Mode = 1, argb(125, 255, 172, 0), argb(255, 129, 129, 129)), near, 18, Regular
    If IsClick Then Mode = 1
    
    Draw.Draw ProDraw.CompatibleDC, 0, 0
End Sub
Private Sub Class_Initialize()
    Draw.Create TargetDC, GWW, GWH
End Sub



