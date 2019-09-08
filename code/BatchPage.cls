VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BatchPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Draw As New Images
Dim Rename As String

Public Property Get Visible() As Boolean
    Visible = (mNowShow = "BatchPage")
End Property
Sub FileDrop(ByVal FileList As String)
    Dim Files() As String
    Dim temp() As String, Path As String, s As Integer
    Dim temp2 As String, temp3() As String
    
    Files = Split(FileList, vbCrLf)
    
    s = 1
    temp = Split(Files(0), "\")
    For i = 0 To UBound(temp) - 1
        Path = Path & temp(i) & "\"
    Next
    
    If Dir(Path & Replace(Rename, "*", "1")) <> "" Then
        If MsgBox("There are other filenames like this , next these files ?", 48 + vbYesNo, "Hello ?") = vbYes Then
            s = Val(InputBox("Next Number ?", "Hey ."))
        End If
    End If
    
    For i = 0 To UBound(Files) - 1
        Name Files(i) As Path & Replace(Rename, "*", format(i, "00"))
    Next
End Sub
Sub Display()
    BackImg.ImageByIndex(1).Draw Draw.CompatibleDC, 0, 0
    
    ProFont.DrawText Draw.Graphics, 0, 110, GWW, 40, "Batch naming", argb(255, 106, 56, 251), center, 25, Regular
    ProFont.DrawText Draw.Graphics, 0, 150, GWW, 30, "Set your name then drop your files here .", argb(125, 106, 56, 251), center, 18, Regular
    
    SetClickArea2 20, 65, 20, 20
    ProFont.DrawText Draw.Graphics, 12, 65, 30, 30, "<", IIf(IsMouseIn, argb(255, 0, 176, 240), argb(255, 129, 129, 129)), center, 18, Regular
    If IsClick Then
        ProCore.FadePage "MainPage"
    End If
    
    SetClickArea2 150, 230, GWW - 300, 30
    ProFont.DrawText Draw.Graphics, 150, 230, GWW - 300, 30, Rename, argb(255, 129, 129, 129), near, 18, Regular
    If IsShowEdit(Rename) Then
        If IsRetEdit Then
            Rename = GetRetEdit
        End If
    End If
    
    Draw.Draw ProDraw.CompatibleDC, 0, 0
End Sub
Private Sub Class_Initialize()
    Draw.Create TargetDC, GWW, GWH
    Rename = "*.png"
End Sub




