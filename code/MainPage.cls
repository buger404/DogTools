VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MainPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Draw As New Images
Public Property Get Visible() As Boolean
    Visible = (mNowShow = "MainPage")
End Property
Sub Display()
    If Logined = False Then
        ProCore.FadePage "PassPage"
        PassPage.Password2 = ""
        PassPage.Types = 0
    End If
    
    BackImg.ImageByIndex(1).Draw Draw.CompatibleDC, 0, 0
    
    If Password = "" Then
        ProFont.DrawText Draw.Graphics, 20, 90, GWW - 20, 30, "You have never set the password .", argb(255, 241, 77, 31), center, 18, Regular
        SetClickArea2 20, 90, GWW - 20, 30
        If IsClick Then
            ProCore.FadePage "PassPage"
            PassPage.Password2 = ""
            PassPage.Types = 1
        End If
    End If
    
    Dim x As Long, IconText(8) As String
    x = GWW / 2 - (IconImg.ImageByIndex(1).Width * 5 + 50 * 4) / 2
    IconText(1) = "Breaker"
    IconText(2) = "Monitor"
    IconText(3) = "Batch naming"
    IconText(4) = "Register"
    IconText(5) = "Handle"
    IconText(6) = "Settings"
    IconText(7) = "Boot items"
    IconText(8) = "Desktop"
    
    For i = 0 To 4
        With IconImg.ImageByIndex(i + 1)
            .SetClickArea x, 170
            .Draw Draw.CompatibleDC, x, 170, IIf(IsMouseIn, 255, 220)
            ProFont.DrawText Draw.Graphics, x - 20, 170 + .Height + 10, .Width + 40, 30, IconText(i + 1), argb(255, 129, 129, 129), center, 18, Regular
            If IsClick Then
                Select Case i
                    Case 0
                        ProCore.FadePage "BreakerPage"
                    Case 1
                        ProCore.FadePage "MonPage"
                    Case 2
                        ProCore.FadePage "BatchPage"
                    Case 3
                        ProCore.FadePage "RegPage"
                    Case 4
                        ProCore.FadePage "HandlePage"
                End Select
            End If
            x = x + .Width + 50
        End With
    Next
    
    x = GWW / 2 - (IconImg.ImageByIndex(1).Width * 5 + 50 * 4) / 2
    
    For i = 0 To 2
        With IconImg.ImageByIndex(i + 1 + 5)
            .SetClickArea x, 320
            .Draw Draw.CompatibleDC, x, 320, IIf(IsMouseIn, 255, 220)
            ProFont.DrawText Draw.Graphics, x - 20, 320 + .Height + 10, .Width + 40, 30, IconText(i + 1 + 5), argb(255, 129, 129, 129), center, 18, Regular
            If IsClick Then
                Select Case i
                    Case 1
                        BootPage.GetBoots
                        ProCore.FadePage "BootPage"
                    Case 2
                        ShellExecuteA 0, "open", App.Path & "\" & App.EXEName & ".exe", "-desktop", "", SW_SHOW
                End Select
            End If
            x = x + .Width + 50
        End With
    Next
    
    If Password <> "" Then
        With IconImg.ImageByIndex(IconImg.Count - 2)
            .SetClickArea 20, GWH - .Height - 20
            .Draw Draw.CompatibleDC, 20, GWH - .Height - 20, IIf(IsMouseIn, 255, 220)
            If IsClick Then
                Logined = False
                Log "Tools", "Ëø¶¨¹¤¾ß"
            End If
        End With
    End If
    
    Draw.Draw ProDraw.CompatibleDC, 0, 0
End Sub
Private Sub Class_Initialize()
    Draw.Create TargetDC, GWW, GWH
End Sub
