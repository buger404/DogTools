VERSION 5.00
Begin VB.Form TryWindow 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "TryWindow"
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2625
   Icon            =   "TryWindow.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Line Line2 
      BorderColor     =   &H00E8E8E8&
      X1              =   10
      X2              =   160
      Y1              =   200
      Y2              =   200
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BAB539&
      Height          =   285
      Left            =   300
      TabIndex        =   7
      Top             =   2550
      Width           =   2040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E8E8E8&
      X1              =   10
      X2              =   160
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "About us"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00919191&
      Height          =   285
      Left            =   300
      TabIndex        =   6
      Top             =   2250
      Width           =   2040
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Check version"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00919191&
      Height          =   285
      Left            =   300
      TabIndex        =   5
      Top             =   1950
      Width           =   2040
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00919191&
      Height          =   285
      Left            =   300
      TabIndex        =   4
      Top             =   1350
      Width           =   2040
   End
   Begin VB.Label LockButton 
      BackStyle       =   0  'Transparent
      Caption         =   "Lock my tools"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BAB539&
      Height          =   285
      Left            =   300
      TabIndex        =   3
      Top             =   3150
      Width           =   2040
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "locked"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   300
      TabIndex        =   2
      Top             =   600
      Width           =   600
   End
   Begin VB.Label TitleText 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dog Tools"
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
      Height          =   315
      Left            =   300
      TabIndex        =   1
      Top             =   300
      Width           =   1155
   End
   Begin VB.Label Background 
      BackColor       =   &H00F2F2F2&
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2715
   End
End
Attribute VB_Name = "TryWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oShadow As New aShadow

Private Sub Form_Load()
    With oShadow
        If .Shadow(Me) Then
            .Color = RGB(0, 0, 0)
            .Depth = 10
            .Transparency = 20
        End If
    End With
    
    SetWindowPos Me.Hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oShadow = Nothing
End Sub

Private Sub Label2_Click()
    MainWindow.Show
    MainWindow.SetFocus
    Me.Hide
End Sub

Private Sub Label4_Click()
    AboutWindow.Show
End Sub

Private Sub Label5_Click()
    If MsgBox("Continue to quit ?", 32 + vbYesNo, "Dog tools") = vbNo Then Exit Sub
    Unload MainWindow
End Sub

Private Sub LockButton_Click()
    Log "Tools", "Ëø¶¨¹¤¾ß"
    Logined = False
    Label1.Caption = "Locked"
    LockButton.Visible = False
    Label5.Visible = False
End Sub
