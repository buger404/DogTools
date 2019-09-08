VERSION 5.00
Begin VB.Form DeskWindow 
   BorderStyle     =   0  'None
   Caption         =   "Desktop"
   ClientHeight    =   7575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12060
   Icon            =   "DeskWindow.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   505
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   804
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer DrawTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   11250
      Top             =   750
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
      Left            =   3150
      TabIndex        =   0
      Text            =   "123"
      Top             =   5100
      Visible         =   0   'False
      Width           =   6315
   End
End
Attribute VB_Name = "DeskWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DrawTimer_Timer()
    ProCore.Display
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    DeskPage.KeyUp KeyCode
End Sub

Private Sub Form_Load()
    Me.Move 0, 0, Screen.Width, Screen.Height
    
    Dim Hwnd As Long
    
    Hwnd = FindWindowA("Progman", "Program Manager")
    SetParent Me.Hwnd, Hwnd

    Set PublicTextBox = EditText
    
    StartProgram Me, App.Path, True
    
    Call LoadDeskAssets
    
    DrawTimer.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    DrawTimer.Enabled = False

    StartProgram Me, App.Path, False
    
    CloseDesktop dHwnd
    
    On Error Resume Next
    Unload MenuWindow
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If EditText.Visible Then Exit Sub
    UpdateClickTest X, Y, 1
    MouseType = Button
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If EditText.Visible Then Exit Sub
    UpdateClickTest X, Y, IIf(Button = 1, 1, 0)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If EditText.Visible Then Exit Sub
    UpdateClickTest X, Y, 2
    MouseType = Button
End Sub

Private Sub EditText_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then tRet = True
End Sub
