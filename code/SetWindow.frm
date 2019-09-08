VERSION 5.00
Begin VB.Form SetWindow 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Settings"
   ClientHeight    =   6105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7740
   Icon            =   "SetWindow.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   407
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   516
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.OptionButton StyleCheck 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Compulsion"
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
      Height          =   315
      Index           =   2
      Left            =   2700
      TabIndex        =   17
      Top             =   4650
      Width           =   1665
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   300
      TabIndex        =   13
      Top             =   2100
      Width           =   6315
      Begin VB.OptionButton UpdateOption 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Only when I got the focus"
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
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   16
         Top             =   300
         Width           =   3465
      End
      Begin VB.OptionButton UpdateOption 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Only when the mouse in me"
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
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Value           =   -1  'True
         Width           =   4365
      End
      Begin VB.OptionButton UpdateOption 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Forever"
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
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   14
         Top             =   600
         Width           =   3465
      End
   End
   Begin VB.OptionButton StyleCheck 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dark"
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
      Height          =   315
      Index           =   1
      Left            =   1500
      TabIndex        =   12
      Top             =   4650
      Width           =   1065
   End
   Begin VB.OptionButton StyleCheck 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Light"
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
      Height          =   315
      Index           =   0
      Left            =   300
      TabIndex        =   11
      Top             =   4650
      Value           =   -1  'True
      Width           =   1065
   End
   Begin VB.TextBox VText 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Left            =   4800
      TabIndex        =   6
      Text            =   "2"
      Top             =   1200
      Width           =   1365
   End
   Begin VB.TextBox HText 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Left            =   1650
      TabIndex        =   4
      Text            =   "8"
      Top             =   1200
      Width           =   1365
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Style"
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
      TabIndex        =   10
      Top             =   4350
      Width           =   450
   End
   Begin VB.Label BootButton 
      Alignment       =   2  'Center
      BackColor       =   &H00E8E8E8&
      Caption         =   "Add to boot items"
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
      TabIndex        =   9
      Top             =   3750
      Width           =   2235
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Other"
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
      TabIndex        =   8
      Top             =   3300
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UI update time"
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
      TabIndex        =   7
      Top             =   1800
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vertical"
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
      Left            =   3750
      TabIndex        =   5
      Top             =   1200
      Width           =   690
   End
   Begin VB.Label OKButton 
      Alignment       =   2  'Center
      BackColor       =   &H00E8E8E8&
      Caption         =   "OK"
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
      Left            =   6450
      TabIndex        =   3
      Top             =   5550
      Width           =   1035
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Horizontal"
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
      Top             =   1200
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Layout"
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
      Top             =   900
      Width           =   630
   End
   Begin VB.Label TitleText 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BAB539&
      Height          =   315
      Left            =   300
      TabIndex        =   0
      Top             =   300
      Width           =   915
   End
End
Attribute VB_Name = "SetWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oShadow As New aShadow

Private Sub BootButton_Click()
    On Error Resume Next
    CreateObject("Wscript.Shell").Regwrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run\Dog Desktop", """" & App.Path & "\" & App.EXEName & """" & " -desktop"
    If Err.Number <> 0 Then
        MsgBox "Failed to create : " & Err.Number & vbCrLf & Err.Description, 48, "Boot Items"
        Err.Clear
    Else
        MsgBox "Success", 64, "Boot Items"
    End If
End Sub

Private Sub Form_Load()

    HText.Text = DeskPage.HC: VText.Text = DeskPage.VC
    UpdateOption(DeskPage.UpdateWay).value = True
    StyleCheck(DeskPage.Style).value = True
    
    With oShadow
        If .Shadow(Me) Then
            .Color = RGB(0, 0, 0)
            .Depth = 10
            .Transparency = 20
        End If
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oShadow = Nothing
End Sub

Private Sub Label11_Click()
    
End Sub

Private Sub OKButton_Click()
    If Val(HText.Text) = 0 Then MsgBox "Horizontal count can not be zero .", 16, "Desktop": Exit Sub
    If Val(VText.Text) = 0 Then MsgBox "Vertical count can not be zero .", 16, "Desktop": Exit Sub
    
    DeskPage.HC = Val(HText.Text)
    DeskPage.VC = Val(VText.Text)
    
    For i = 0 To UpdateOption.UBound
        If UpdateOption(i).value Then DeskPage.UpdateWay = i
    Next
    For i = 0 To StyleCheck.UBound
        If StyleCheck(i).value Then DeskPage.Style = i
    Next
    
    SaveSetting "Dog Tools", "Settings", "HC", DeskPage.HC
    SaveSetting "Dog Tools", "Settings", "VC", DeskPage.VC
    SaveSetting "Dog Tools", "Settings", "UpdateWay", DeskPage.UpdateWay
    SaveSetting "Dog Tools", "Settings", "Style", DeskPage.Style
    
    Unload Me
End Sub
