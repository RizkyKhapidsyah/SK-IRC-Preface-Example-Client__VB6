VERSION 5.00
Begin VB.Form about 
   Caption         =   "About IRCPre2"
   ClientHeight    =   4500
   ClientLeft      =   1605
   ClientTop       =   1440
   ClientWidth     =   5640
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   376
   Begin VB.CommandButton OK 
      Caption         =   "&OK"
      Height          =   400
      Left            =   4544
      TabIndex        =   8
      Top             =   4032
      Width           =   976
   End
   Begin VB.Frame Frame1 
      Height          =   2896
      Left            =   128
      TabIndex        =   2
      Top             =   1024
      Width           =   5392
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Dasmius can be reached by email at danny@telecomputer.com"
         Height          =   208
         Left            =   128
         TabIndex        =   7
         Top             =   2496
         Width           =   4768
      End
      Begin VB.Label Label6 
         Caption         =   "http://www.telecomputer.com/ircpre"
         ForeColor       =   &H00FF0000&
         Height          =   208
         Left            =   256
         TabIndex        =   6
         Top             =   2112
         Width           =   2768
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "The official IRCPre home page is:"
         Height          =   208
         Left            =   128
         TabIndex        =   5
         Top             =   1856
         Width           =   2528
      End
      Begin VB.Label Label4 
         Caption         =   $"about.frx":0000
         ForeColor       =   &H00000000&
         Height          =   848
         Left            =   128
         TabIndex        =   4
         Top             =   832
         Width           =   5136
      End
      Begin VB.Label Label3 
         Caption         =   "IRCPre2 was written by Dasmius. He can be found on EFNet IRC, in the channel #visualbasic."
         Height          =   464
         Left            =   128
         TabIndex        =   3
         Top             =   256
         Width           =   5136
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "IRCPre2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   47.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1024
      Left            =   128
      TabIndex        =   1
      Top             =   0
      Width           =   3152
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "IRCPre2 Copyright (C) 1996 by Dann M. Daggett II"
      Height          =   208
      Left            =   576
      TabIndex        =   0
      Top             =   4112
      Width           =   3808
   End
End
Attribute VB_Name = "about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

  ' Center myself in the middle of the screen
  Me.Left = (Screen.Width \ 2) - (Me.Width \ 2)
  Me.Top = (Screen.Height \ 2) - (Me.Height \ 2)

End Sub


Private Sub OK_Click()

  ' Close About box
  Unload Me
  
End Sub


