VERSION 5.00
Begin VB.Form MenuWindow 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "MenuWindow"
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2880
   Icon            =   "MenuWindow.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   192
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer CheckTimer 
      Interval        =   100
      Left            =   2250
      Top             =   150
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Open desktop folder"
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
      Left            =   450
      TabIndex        =   8
      Top             =   1950
      Width           =   2040
   End
   Begin VB.Label PerButton 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Personalization"
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
      Left            =   450
      TabIndex        =   3
      Top             =   1650
      Width           =   2040
   End
   Begin VB.Label CtrlButton 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Control Panel"
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
      Left            =   450
      TabIndex        =   7
      Top             =   1350
      Width           =   2040
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E8E8E8&
      X1              =   10
      X2              =   180
      Y1              =   80
      Y2              =   80
   End
   Begin VB.Label SetButton 
      BackStyle       =   0  'Transparent
      Caption         =   "Settings"
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
      Top             =   2550
      Width           =   2040
   End
   Begin VB.Label CreateButton 
      BackStyle       =   0  'Transparent
      Caption         =   "Create ..."
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
      Top             =   450
      Width           =   2040
   End
   Begin VB.Label RefreshButton 
      BackStyle       =   0  'Transparent
      Caption         =   "Refresh"
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
      Top             =   750
      Width           =   2040
   End
   Begin VB.Label FolderButton 
      BackStyle       =   0  'Transparent
      Caption         =   "New Folder"
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
      TabIndex        =   2
      Top             =   150
      Width           =   2040
   End
   Begin VB.Label AboutButton 
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
      TabIndex        =   1
      Top             =   2850
      Width           =   2040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E8E8E8&
      X1              =   10
      X2              =   180
      Y1              =   160
      Y2              =   160
   End
   Begin VB.Label ExitButton 
      BackStyle       =   0  'Transparent
      Caption         =   "Exit Dog Desktop"
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
      TabIndex        =   0
      Top             =   3150
      Width           =   2040
   End
   Begin VB.Image CtrlIcon 
      Height          =   885
      Left            =   150
      Picture         =   "MenuWindow.frx":000C
      Top             =   1200
      Width           =   915
   End
End
Attribute VB_Name = "MenuWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oShadow As New aShadow

Private Sub AboutButton_Click()
    AboutWindow.Show
    Me.Hide
End Sub

Private Sub CheckTimer_Timer()
    If GetActiveWindow <> Me.Hwnd Then Me.Hide
    If Me.top / 15 + Me.ScaleHeight >= Screen.Height / 15 - 50 Then
        Me.top = (Screen.Height / 15 - 50 - Me.ScaleHeight) * 15
    End If
End Sub

Private Sub CtrlButton_Click()
    ShellExecuteA 0, "open", "control.exe", "", "", SW_SHOW
    Me.Hide
End Sub

Private Sub ExitButton_Click()
    Unload DeskWindow
End Sub

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

Private Sub Label1_Click()
    ShellExecuteA 0, "open", "explorer.exe", CreateObject("Wscript.Shell").SpecialFolders("Desktop"), "", SW_SHOW
    Me.Hide
End Sub

Private Sub PerButton_Click()
    ShellExecuteA 0, "open", "control.exe", "/name Microsoft.Personalization", "", SW_SHOW
    Me.Hide
End Sub

Private Sub RefreshButton_Click()
    DeskPage.Refresh
    ProCore.FadePage mNowShow
    Me.Hide
End Sub

Private Sub SetButton_Click()
    SetWindow.Show
    Me.Hide
End Sub
