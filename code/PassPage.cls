VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PassPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Draw As New Images
Public Password2 As String, Types As Integer
Public Property Get Visible() As Boolean
    Visible = (mNowShow = "PassPage")
End Property
Sub Display()
    BackImg.ImageByIndex(1).Draw Draw.CompatibleDC, 0, 0
    
    If Password = "" Then mNowShow = "MainPage"
    
    ProFont.DrawText Draw.Graphics, 20, 90, GWW - 20, 30, IIf(Types = 0, "Enter", "Set") & " your password .", argb(255, 129, 129, 129), center, 18, Regular
    
    Dim x As Long, IconText(7) As String, Pen As Long, Index As Integer, Index2 As Integer, Passed(8) As Boolean
    Dim DrawX(8) As Long, DrawY(8) As Long
    GdipCreatePen1 argb(255, 241, 77, 31), 3, UnitPixel, Pen
    
    For s = 0 To 2
        x = GWW / 2 - (IconImg.ImageByIndex(IconImg.Count).Width * 3 + 50 * 2) / 2
        For i = 0 To 2
            With IconImg.ImageByIndex(IconImg.Count)
                DrawX(s * 3 + i) = x + .Width / 2: DrawY(s * 3 + i) = 160 + s * 120 + .Height / 2
                x = x + .Width + 50
            End With
        Next
    Next
    
    For i = 1 To Len(Password2)
        Passed(Val(Mid(Password2, i, 1))) = True
        If i > 1 Then
            Index = Val(Mid(Password2, i - 1, 1))
            Index2 = Val(Mid(Password2, i, 1))
            GdipDrawLine Draw.Graphics, Pen, DrawX(Index), DrawY(Index), DrawX(Index2), DrawY(Index2)
        End If
    Next
    
    For i = 0 To 8
        With IconImg.ImageByIndex(IIf(Passed(i), IconImg.Count, IconImg.Count - 1))
            x = DrawX(i) - .Width / 2
            .SetClickArea x, DrawY(i) - .Height / 2
            If IsClick And (Not Passed(i)) Then Password2 = Password2 & i
            .Draw Draw.CompatibleDC, x, DrawY(i) - .Height / 2, IIf(Passed(i), 255, IIf(IsMouseIn, 255, 210))
        End With
    Next
    
    For i = 0 To 8
        ProFont.DrawText Draw.Graphics, DrawX(i), DrawY(i) - 26, 0, 0, i + 1, argb(200, 255, 255, 255), center, 36, Regular
    Next
    
    GdipDeletePen Pen
    
    If Types = 1 Then
        CtrlImg.ImageByIndex(1).DrawCtrl Draw.CompatibleDC, Draw.Graphics, 30, GWH - 60, "Cancel", argb(255, 129, 129, 129), 16, Regular, Button
        If IsClick Then ProCore.FadePage "MainPage"
    End If
    
    CtrlImg.ImageByIndex(1).DrawCtrl Draw.CompatibleDC, Draw.Graphics, GWW - 250, GWH - 60, "Clear", argb(255, 129, 129, 129), 16, Regular, Button
    If IsClick Then Password2 = ""
    CtrlImg.ImageByIndex(2).DrawCtrl Draw.CompatibleDC, Draw.Graphics, GWW - 130, GWH - 60, "Finish", argb(255, 255, 255, 255), 16, Regular, Button
    If IsClick Then
        If Types = 1 Then
            If Len(Password2) < 4 Then MsgBox "Please connect at least four points.", 48: Exit Sub
            Password = BMEA(Password2)
            SaveSetting "Dog Tools", "Gen", "Password", Password
            Logined = True
            ProCore.FadePage "MainPage"
            Log "Tools", "密码设置成功"
        Else
            If BMEA(Password2) = Password Then
                Logined = True
                ProCore.FadePage "MainPage"
                Log "Tools", "密码验证通过"
            Else
                MsgBox "Password error .", 48: Exit Sub
            End If
        End If
    End If
    Draw.Draw ProDraw.CompatibleDC, 0, 0
End Sub
Private Sub Class_Initialize()
    Draw.Create TargetDC, GWW, GWH
End Sub

