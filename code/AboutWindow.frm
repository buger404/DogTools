VERSION 5.00
Begin VB.Form AboutWindow 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "About us"
   ClientHeight    =   5610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7740
   Icon            =   "AboutWindow.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   374
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   516
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Image Image1 
      Height          =   720
      Left            =   6750
      Picture         =   "AboutWindow.frx":1BCC2
      Top             =   1500
      Width           =   720
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E8E8E8&
      X1              =   20
      X2              =   490
      Y1              =   250
      Y2              =   250
   End
   Begin VB.Label Label11 
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
      TabIndex        =   12
      Top             =   5100
      Width           =   1035
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You are using this software , it means you accept all these licenses ."
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
      TabIndex        =   11
      Top             =   4200
      Width           =   6210
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Don't use this software to do something bad , conceited at your own risk ."
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
      TabIndex        =   10
      Top             =   4500
      Width           =   6840
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This software is free ."
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
      TabIndex        =   9
      Top             =   3900
      Width           =   1935
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ris_vb@126.com"
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
      Left            =   1350
      TabIndex        =   8
      Top             =   2100
      Width           =   1530
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail"
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
      Top             =   2100
      Width           =   570
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1361778219"
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
      Left            =   1350
      TabIndex        =   6
      Top             =   1800
      Width           =   1200
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QQ"
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
      Top             =   1800
      Width           =   330
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Error 404"
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
      Left            =   1350
      TabIndex        =   4
      Top             =   1500
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Maker"
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
      TabIndex        =   3
      Top             =   1500
      Width           =   585
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 0.1.0"
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
      TabIndex        =   0
      Top             =   600
      Width           =   1200
   End
   Begin VB.Label Background 
      BackColor       =   &H00F2F2F2&
      Height          =   1215
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7815
   End
End
Attribute VB_Name = "AboutWindow"
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

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oShadow = Nothing
End Sub

Private Sub Label11_Click()
    Unload Me
End Sub
