VERSION 5.00
Begin VB.Form setup 
   Caption         =   "IRCPre2 Setup"
   ClientHeight    =   1980
   ClientLeft      =   1575
   ClientTop       =   1425
   ClientWidth     =   3015
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1980
   ScaleWidth      =   3015
   Begin VB.CommandButton OK 
      Caption         =   "&OK"
      Height          =   400
      Left            =   704
      TabIndex        =   7
      Top             =   1472
      Width           =   1040
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   1856
      TabIndex        =   6
      Top             =   1472
      Width           =   1040
   End
   Begin VB.TextBox NickText 
      Height          =   304
      Left            =   1088
      MaxLength       =   9
      TabIndex        =   5
      Text            =   "IRCPre2"
      Top             =   1024
      Width           =   1808
   End
   Begin VB.TextBox PortText 
      Height          =   304
      Left            =   1088
      TabIndex        =   3
      Text            =   "6667"
      Top             =   576
      Width           =   1808
   End
   Begin VB.ComboBox ServerCombo 
      Height          =   336
      ItemData        =   "setup.frx":0000
      Left            =   1088
      List            =   "setup.frx":0013
      TabIndex        =   1
      Text            =   "irc.neosoft.com"
      Top             =   128
      Width           =   1808
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nickname:"
      Height          =   208
      Index           =   2
      Left            =   128
      TabIndex        =   4
      Top             =   1088
      Width           =   816
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Port:"
      Height          =   208
      Index           =   1
      Left            =   128
      TabIndex        =   2
      Top             =   640
      Width           =   352
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Server:"
      Height          =   208
      Index           =   0
      Left            =   128
      TabIndex        =   0
      Top             =   192
      Width           =   544
   End
End
Attribute VB_Name = "setup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cancel_Click()

  ' Close setup
  Unload Me
  
End Sub


Private Sub Form_Load()

  ' Center myself in the middle of the screen
  Me.Left = (Screen.Width \ 2) - (Me.Width \ 2)
  Me.Top = (Screen.Height \ 2) - (Me.Height \ 2)
  
End Sub

Private Sub OK_Click()

  ' Make sure all fields have data
  If ServerCombo.Text = "" Then
    Beep
    ServerCombo.SetFocus: Exit Sub
  End If
  If PortText.Text = "" Then
    Beep
    PortText.SetFocus: Exit Sub
  End If
  If NickText.Text = "" Then
    Beep
    NickText.SetFocus: Exit Sub
  End If

  ' Set the global variables
  Server = ServerCombo.Text
  Port = PortText.Text
  Nickname = NickText.Text
  ' Close setup
  Unload Me

End Sub


