VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ristiseiska"
   ClientHeight    =   9915
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   10455
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   661
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   697
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1440
      Index           =   17
      Left            =   2760
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   10
      Top             =   0
      Width           =   1065
      Visible         =   0   'False
   End
   Begin MSComctlLib.StatusBar sbrStatus 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   9
      Top             =   9615
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   529
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1680
      Index           =   2
      Left            =   6495
      ScaleHeight     =   112
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   47
      TabIndex        =   8
      Top             =   120
      Width           =   705
   End
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1545
      Index           =   4
      Left            =   7680
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   77
      TabIndex        =   3
      Top             =   4050
      Width           =   1155
   End
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   3
      Left            =   8970
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   2
      Top             =   3840
      Width           =   1335
   End
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   720
      Index           =   1
      Left            =   120
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   1
      Top             =   4080
      Width           =   1335
   End
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1680
      Index           =   0
      Left            =   6135
      ScaleHeight     =   112
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   0
      Top             =   7800
      Width           =   1065
   End
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1890
      Index           =   5
      Left            =   2850
      ScaleHeight     =   126
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   11
      Top             =   2040
      Width           =   1065
   End
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1890
      Index           =   6
      Left            =   4050
      ScaleHeight     =   126
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   12
      Top             =   2040
      Width           =   1065
   End
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1890
      Index           =   7
      Left            =   5250
      ScaleHeight     =   126
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   13
      Top             =   2040
      Width           =   1065
   End
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1890
      Index           =   8
      Left            =   6450
      ScaleHeight     =   126
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   14
      Top             =   2040
      Width           =   1065
   End
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1440
      Index           =   9
      Left            =   2850
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   15
      Top             =   4050
      Width           =   1065
   End
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1440
      Index           =   10
      Left            =   4050
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   16
      Top             =   4050
      Width           =   1065
   End
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1440
      Index           =   11
      Left            =   5250
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   17
      Top             =   4050
      Width           =   1065
   End
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1440
      Index           =   12
      Left            =   6450
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   18
      Top             =   4050
      Width           =   1065
   End
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1890
      Index           =   16
      Left            =   6450
      ScaleHeight     =   126
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   19
      Top             =   5610
      Width           =   1065
   End
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1890
      Index           =   15
      Left            =   5250
      ScaleHeight     =   126
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   20
      Top             =   5610
      Width           =   1065
   End
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1890
      Index           =   14
      Left            =   4050
      ScaleHeight     =   126
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   21
      Top             =   5610
      Width           =   1065
   End
   Begin VB.PictureBox picDeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1890
      Index           =   13
      Left            =   2850
      ScaleHeight     =   126
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   22
      Top             =   5610
      Width           =   1065
   End
   Begin VB.Label lblPlayer 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pelaaja 4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   9525
      TabIndex        =   7
      Top             =   5040
      Width           =   780
   End
   Begin VB.Label lblPlayer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pelaaja 3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   7320
      TabIndex        =   6
      Top             =   120
      Width           =   780
   End
   Begin VB.Label lblPlayer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pelaaja 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   780
   End
   Begin VB.Label lblPlayer 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pelaaja 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   3900
      TabIndex        =   4
      Top             =   9240
      Width           =   885
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Peli"
      Begin VB.Menu mnuGameNew 
         Caption         =   "&Uusi peli"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuGameNetwork 
         Caption         =   "&Verkkopeli..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGame0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameSettings 
         Caption         =   "&Asetukset..."
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuGameScore 
         Caption         =   "&Pisteet..."
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuGameSound 
         Caption         =   "&Äänet"
      End
      Begin VB.Menu mnuGame1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameDemo 
         Caption         =   "&Demo"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuGame2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameExit 
         Caption         =   "&Lopeta"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Ohje"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Ohjeen aiheet"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelp0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&Tietoja..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Me.Refresh
    GameInit
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DeckClick -1
End Sub


Private Sub Form_Resize()
    FormMainResize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    GameUninit
End Sub


Private Sub lblPlayer_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    DeckClick -1
End Sub


Private Sub mnuGameDemo_Click()
    GameDemo
End Sub
Private Sub mnuGameExit_Click()
    Form_Unload False
End Sub

Private Sub mnuGameNetwork_Click()
    GameNetwork
End Sub

Private Sub mnuGameNew_Click()
    GameNew
End Sub
Private Sub mnuGameScore_Click()
    GameScore True
End Sub

Private Sub mnuGameSettings_Click()
    GameSettings
End Sub

Private Sub mnuGameSound_Click()
    mnuGameSound.Checked = Not mnuGameSound.Checked
    Game.Sound = mnuGameSound.Checked
    SaveSettings
End Sub

Private Sub mnuHelpAbout_Click()
    HelpAbout
End Sub

Private Sub mnuHelpContents_Click()
    HelpContents
End Sub

Private Sub picDeck_KeyPress(Index As Integer, KeyAscii As Integer)
    DeckKeyPress Index, KeyAscii
End Sub

Private Sub picDeck_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then
        Exit Sub
    End If
    DeckClick Index, GetCardIndex(Deck(Index), X, Y)
End Sub

