VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "IRC Preface Example Client - Revised"
   ClientHeight    =   4665
   ClientLeft      =   1545
   ClientTop       =   1665
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   311
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   459
   Begin MSWinsockLib.Winsock TCP1 
      Left            =   4800
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox NameList 
      Height          =   3765
      Left            =   5296
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   384
      Width           =   1488
   End
   Begin VB.TextBox Topic 
      Height          =   304
      Left            =   64
      TabIndex        =   2
      Top             =   64
      Width           =   6720
   End
   Begin VB.TextBox Outgoing 
      Height          =   300
      Left            =   64
      TabIndex        =   1
      Top             =   3888
      Width           =   5232
   End
   Begin VB.TextBox Incoming 
      Height          =   3504
      Left            =   64
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   384
      Width           =   5232
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu FileConnect 
         Caption         =   "&Connect"
      End
      Begin VB.Menu FileSetup 
         Caption         =   "&Setup"
      End
      Begin VB.Menu dash 
         Caption         =   "-"
      End
      Begin VB.Menu FileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu HelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This example and included document are
' Copyright (C) 1996 by Dann Daggett II

' Please read the document that comes with this
' program.

Dim CRLF As String ' Cairrage return/Line feed
Dim OldText As String ' Holds any text still
                      ' needing processed
Dim channel As String ' Holds the channel name
Dim CMode ' CurrentMode of client
          ' 0 is logged in
          ' 1 is joining channel
          ' 2 is in channel

Sub AddText(textmsg As String)

  ' Add the data in textmsg to the Incoming
  ' text box and force the text down
  Incoming.Text = Incoming.Text & textmsg & CRLF

End Sub


Sub SendData(textmsg As String)

  ' Send the data in textmsg to the server, and
  ' add a CRLF
  TCP1.SendData textmsg & CRLF
  
End Sub



Private Sub FileConnect_Click()

  If FileConnect.Caption = "&Connect" Then
    ' Set the RemoteHost to the IRC Server Host
    TCP1.RemoteHost = Server
    ' Set the Port to connect to
    TCP1.RemotePort = Port
    ' Connect
    TCP1.Connect
    ' Clear textbox, topic and listbox
    Incoming.Text = ""
    NameList.Clear
    Topic.Text = ""
    AddText "*** Attempting to connect to " & Server & "..."
    FileConnect.Caption = "&Disconnect"
  Else
    FileConnect.Caption = "&Connect"
    AddText "*** Disconnected"
    ' Close the socket
    TCP1.Close
  End If

End Sub

Private Sub FileExit_Click()

  ' Close the program
  Unload Me

End Sub

Private Sub FileSetup_Click()

  ' Show the setup form
  setup.Show 1

End Sub

Private Sub Form_Activate()

  ' Scroll the textbox down again
  Incoming_Change
  
End Sub

Private Sub Form_Load()

  ' Set CRLF to be Cairrage Return + Line Feed,
  ' ALL IRC messages end with this
  CRLF = Chr$(13) & Chr$(10)
  ' Set the current mode to 0
  CMode = 0
  
  'Set the default values
  Server = "irc.neosoft.com"
  Port = 6667
  Nickname = "IRCPre2"
  
End Sub

Private Sub HelpAbout_Click()

  about.Show 1

End Sub

Private Sub Incoming_Change()

' We want this box to scroll down automatically.

  Incoming.SelStart = Len(Incoming.Text)

' What this does is says, make the start of my
' selected text the end of the entire text,
' which effectively scrolls down the textbox,
' but does not select anything. The len()
' command returns the length of characters of
' the text, in a number.

End Sub


Private Sub Incoming_GotFocus()

' We don't want the client to be able to edit
' the Incoming textbox.

  Outgoing.SetFocus

' This make it so the user cannot click inside
' the Incoming text box, but can still scroll it.
' It does this by giving another object the
' focus.

End Sub


Private Sub Outgoing_KeyPress(KeyAscii As Integer)

  Dim msg As String
  
  ' Exit unless its a return, then process
  If KeyAscii <> 13 Then Exit Sub
  KeyAscii = 0 ' Stop that stupid beep!
  msg = Outgoing.Text
  If Left$(msg, 1) <> "/" Then
    ' they want to send a msg, send it if we're
    ' in a channel
    If NameList.ListCount > 0 Then
      SendData "PRIVMSG " & channel & " :" & msg
      AddText "> " & msg
    End If
  Else
    Outgoing.Text = Mid$(Outgoing.Text, 2)
    msg = Mid$(Outgoing.Text, InStr(Outgoing.Text, " ") + 1)
    Select Case UCase$(Left$(Outgoing.Text, InStr(Outgoing.Text, " ") - 1)) ' see what kind of action to do
      Case "JOIN"
        SendData "JOIN " & msg: CMode = 1 ' join the channel, set the mode
        channel = msg
      Case "ME"
        ' if we're in a channel, then do an action
        If NameList.ListCount > 0 Then SendData "PRIVMSG " & channel & " :" & Chr$(1) & "ACTION " & msg & Chr$(1)
        AddText "* " & Nickname & " " & msg
      Case "MSG"
        ' send a priv msg
        SendData "PRIVMSG " & Left$(msg, InStr(msg, " ") - 1) & " :" & Mid$(msg, InStr(msg, " ") + 1)
        AddText "=->" & Left$(msg, InStr(msg, " ") - 1) & "<-= " & Mid$(msg, InStr(msg, " ") + 1)
    End Select
  End If
  ' clear the textbox
  Outgoing.Text = ""

End Sub


Private Sub TCP1_Close()

  FileConnect.Caption = "&Connect"
  AddText "*** Disconnected"
  ' Close the socket
  TCP1.Close
  
End Sub

Private Sub TCP1_Connect()

  ' Physical connect
  AddText "*** Connection established."
  AddText "*** Sending login information..."
  
  ' Send the server my nickname
  SendData "NICK " & Nickname
  ' Send the server the user information
  SendData "USER email " & TCP1.LocalIP & " " & Server & " :username"

End Sub

Private Sub TCP1_DataArrival(ByVal bytesTotal As Long)

  Dim inData As String
  Dim sline As String
  Dim msg As String
  Dim msg2 As String
  Dim x
  
  ' Get the incoming data into a string
  TCP1.GetData inData, vbString
  ' Add any unprocessed text on first
  inData = OldText & inData
  
  ' Some IRC servers are only using a Cairrage
  ' Retrun, or a LineFeed, instead of both, so
  ' we need to be prepared for that
  x = 0
  If Right$(inData, 2) = CRLF Then x = 1
  If Right$(inData, 1) = Chr$(10) Then x = 1
  If Right$(inData, 1) = Chr$(13) Then x = 1
  If x = 1 Then
    OldText = "" ' its a full send, process
  Else
    OldText = inData: Exit Sub ' incomplete send
                               ' save and exit
  End If
  
again:
  GoSub parsemsg ' get next msg fragment
  If Left$(sline, 6) = "PING :" Then ' we need to pong to stay alive
    AddText "PING? PONG!"
    SendData "PONG " & Server
    GoTo again ' get next msg
  End If
  If Left$(sline, 5) = "ERROR" Then ' some error
    AddText "*** ERROR " & Mid$(sline, InStr(sline, "("))
  End If
  If Left$(sline, Len(Nickname) + 1) = ":" & Nickname Then
    ' a command for the client only
    sline = Mid$(sline, InStr(sline, " ") + 1)
    Select Case Left$(sline, InStr(sline, " ") - 1)
      Case "MODE"
        AddText "*** Your mode is now " & Mid$(sline, InStr(sline, ":") + 1)
    End Select
  End If
  If Mid$(sline, InStr(sline, " ") + 1, 7) = "PRIVMSG" Then
    'someone /msged us
    msg = Mid$(sline, InStr(sline, " ") + 9)
    If LCase$(Left$(msg, InStr(msg, " ") - 1)) = LCase$(Nickname) Then ' private msg
      ' add so its: --nick-- msg here
      AddText "--" & Mid$(sline, 2, InStr(sline, "!") - 2) & "-- " & Mid$(msg, InStr(msg, ":") + 1)
    End If
  End If
  Select Case CMode
    Case 0 ' not in channel
      If Mid$(sline, InStr(1, sline, " ") + 1, 3) = "001" Then
        Server = Mid$(sline, 2, InStr(sline, " ") - 2)
      End If
      If Left$(sline, Len(Server) + 1) = ":" & Server Then
        ' its a server msg, add the important part
        sline = Mid$(sline, InStr(2, sline, ":") + 1)
        ':washington.dc.us.undernet.org 001 Das2 :Welcome to the Internet Relay Network Das2
        AddText sline
      End If
    Case 1 ' joining channel
      If Left$(sline, Len(Server) + 1) = ":" & Server Then
        msg = Mid$(sline, InStr(sline, " ") + 1)
        Select Case Left$(msg, InStr(msg, " ") - 1)
          Case "332" ' Topic
            Topic.Text = Mid$(msg, InStr(msg, ":") + 1)
          Case "353" ' Name list
            msg = Mid$(msg, InStr(msg, ":") + 1)
            Do Until msg = "" ' break apart names and add them seperatly
              x = InStr(msg, " ")
              If x <> 0 Then
                NameList.AddItem Left$(msg, x - 1)
                msg = Mid$(msg, x + 1)
              Else
                NameList.AddItem msg
                msg = ""
              End If
            Loop
          Case "366" ' End of Name List
            CMode = 2 ' change mode to joined channel
        End Select
      Else
        ' someone joined the channel, us!
        If Left$(sline, InStr(sline, " ") - 1) = "JOIN" Then
          AddText "*** " & Nickname & " has joined " & channel
        End If
      End If
    Case 2 ' in a channel
      If Mid$(sline, InStr(sline, " ") + 1, 7) = "PRIVMSG" Then
        msg = Mid$(sline, InStr(sline, " ") + 9)
        If LCase$(Left$(msg, InStr(msg, " ") - 1)) = LCase$(Nickname) Then ' private msg
          AddText "--" & Mid$(sline, 2, InStr(sline, "!") - 2) & "-- " & Mid$(msg, InStr(msg, ":") + 1)
        Else ' channel msg
          If Left$(Mid$(msg, InStr(msg, ":") + 1), 1) = Chr$(1) Then ' action
            msg2 = Mid$(msg, InStr(msg, ":") + 9)
            AddText "* " & Mid$(sline, 2, InStr(sline, "!") - 2) & " " & Left$(msg2, Len(msg2) - 1)
          Else ' msg
            AddText "<" & Mid$(sline, 2, InStr(sline, "!") - 2) & "> " & Mid$(msg, InStr(msg, ":") + 1)
          End If
        End If
      Else
        ' command not yet supported, just display it
        AddText sline
      End If
  End Select
  ' Did I say "Good programming practice?"
  ' Sometimes its easier to do this
  GoTo again
Exit Sub

parsemsg:
  ' irc may send more than one msg at a time,
  ' so parse them first
  If inData = "" Then Exit Sub
  x = InStr(inData, CRLF) ' find the break
  If x <> 0 Then
    sline = Left$(inData, x - 1)
    ' strip off the text
    If Len(inData) > x + 2 Then
      inData = Mid$(inData, x + 2)
    Else
      inData = ""
    End If
  Else
    x = InStr(inData, Chr$(13)) ' find the break
    If x = 0 Then
      x = InStr(inData, Chr$(10)) ' find the break
    End If
    If x <> 0 Then
      sline = Left$(inData, x - 1)
    Else
      sline = inData
    End If
    ' strip off the text
    If Len(inData) > x + 1 Then
      inData = Mid$(inData, x + 1)
    Else
      inData = ""
    End If
  End If
Return

End Sub

Private Sub Topic_GotFocus()

  ' We don't want the client to be able to edit
  ' the topic
  Outgoing.SetFocus

End Sub


