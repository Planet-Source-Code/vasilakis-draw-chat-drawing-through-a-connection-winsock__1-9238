VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Draw Chat v1.0"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMain.frx":0442
   ScaleHeight     =   6000
   ScaleWidth      =   7500
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5700
      Width           =   855
   End
   Begin VB.TextBox txtChat 
      Height          =   1455
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   3360
      Width           =   6135
   End
   Begin VB.ComboBox txtMe 
      Height          =   330
      Left            =   1200
      TabIndex        =   9
      Top             =   4800
      Width           =   6135
   End
   Begin VB.CommandButton cmdServer 
      Caption         =   "Server"
      Height          =   350
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5400
      Width           =   855
   End
   Begin MSWinsockLib.Winsock wsockSERVER 
      Left            =   6960
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.TextBox txtNick 
      Height          =   315
      Left            =   5880
      TabIndex        =   7
      Text            =   "NickName"
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   350
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5400
      Width           =   855
   End
   Begin VB.TextBox txtHOST 
      Height          =   315
      Left            =   1200
      TabIndex        =   5
      Text            =   "localhost"
      Top             =   5400
      Width           =   2655
   End
   Begin VB.TextBox txtWidth 
      Height          =   315
      Left            =   600
      TabIndex        =   4
      Text            =   "1"
      Top             =   5280
      Width           =   375
   End
   Begin VB.PictureBox picOther 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3035
      Left            =   3840
      ScaleHeight     =   3000
      ScaleWidth      =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   3510
   End
   Begin VB.PictureBox picME 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3035
      Left            =   120
      ScaleHeight     =   3000
      ScaleWidth      =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   3510
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   19
      ToolTipText     =   "Black Line"
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   18
      ToolTipText     =   "Red Line"
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H000080FF&
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   17
      ToolTipText     =   "Orange Line"
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H0000FFFF&
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   16
      ToolTipText     =   "Yellow line"
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H0000FF00&
      Height          =   375
      Index           =   4
      Left            =   600
      TabIndex        =   15
      ToolTipText     =   "Green Line"
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00FFFF00&
      Height          =   375
      Index           =   5
      Left            =   600
      TabIndex        =   14
      ToolTipText     =   "Cyan Line"
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00FF0000&
      Height          =   375
      Index           =   6
      Left            =   600
      TabIndex        =   13
      ToolTipText     =   "Blue Line"
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00FF00FF&
      Height          =   375
      Index           =   7
      Left            =   120
      TabIndex        =   12
      ToolTipText     =   "Mauve Line"
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   8
      Left            =   600
      TabIndex        =   11
      ToolTipText     =   "White Line"
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label lblST 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1200
      TabIndex        =   3
      Top             =   5750
      Width           =   465
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public oldX As Long
Public oldY As Long
Public mDown

Private Sub cmdClear_Click()
On Error Resume Next
picME.Cls
wsockSERVER.SendData "clear;" & vbCr
End Sub

Private Sub cmdConnect_Click()
Tag = ""
If cmdConnect.Caption = "Close" Then
    wsockSERVER.Close
    wsockSERVER_Close
    wsockSERVER.RemotePort = 0
    wsockSERVER.LocalPort = 0
Else
    cmdServer.Enabled = False
    lblColor(0).BorderStyle = 1
    txtWidth.Text = "1"
    picME.Cls
    picOther.Cls
    wsockSERVER.Close
    wsockSERVER.RemotePort = 0
    wsockSERVER.LocalPort = 0
    lblST.Caption = "Connecting [" & txtHOST.Text & "]"
    wsockSERVER.Connect txtHOST.Text, 7676
    Do
        DoEvents
    Loop Until wsockSERVER.State = sckConnected Or wsockSERVER.State = sckError
    If wsockSERVER.State = sckConnected Then
        lblST.Caption = "Connected with " & wsockSERVER.RemoteHostIP & "..."
        Tag = txtNick.Text
        wsockSERVER.SendData "nick;" & txtNick.Text & vbCr
        DoEvents
        cmdConnect.Caption = "Close"
    Else
        cmdServer.Enabled = True
        lblST.Caption = "Unable to establish a connection with " & txtHOST.Text & "..."
        cmdConnect.Caption = "Connect"
    End If
End If

End Sub

Private Sub cmdServer_Click()
Tag = ""
On Error Resume Next
If cmdServer.Caption = "Server" Then
 wsockSERVER.Close
 wsockSERVER.RemotePort = 0
 wsockSERVER.LocalPort = 7676
 Err = 0
 wsockSERVER.Listen
 If Err <> 0 Then Exit Sub
 lblST.Caption = "Listening for Connections..."
 cmdServer.Caption = "Close"
 cmdConnect.Enabled = False
Else
 cmdConnect.Enabled = True
 wsockSERVER.Close
 cmdServer.Caption = "Server"
 lblST.Caption = "Idle"
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
wsockSERVER.Close
wsockSERVER.SendData "draw;" & oldX & ":" & oldY & ":" & X & ":" & Y & vbCr: DoEvents
Caption = "Draw Chat! v1.0 - Running on " & wsockSERVER.LocalHostName
txtNick.Text = GetSetting("Draw Chat!", "1.0\Settings", "Nickname", "Nick")
Top = Val(GetSetting("Draw Chat!", "1.0\Settings", "Top"))
Left = Val(GetSetting("Draw Chat!", "1.0\Settings", "left"))
txtNick.Tag = txtNick.Text
lblST.Caption = "Idle"
Unload frmSup
End Sub



Private Sub Form_Unload(Cancel As Integer)
SaveSetting "Draw Chat!", "1.0\Settings", "Nickname", txtNick.Text
SaveSetting "Draw Chat!", "1.0\Settings", "Top", Top
SaveSetting "Draw Chat!", "1.0\Settings", "Left", Left
End
End Sub


Private Sub lblColor_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
For I = 0 To lblColor.Count - 1
 If I <> Index Then
    If lblColor(I).BorderStyle <> 0 Then lblColor(I).BorderStyle = 0
 End If
Next
lblColor(Index).BorderStyle = 1
On Error Resume Next
picME.ForeColor = lblColor(Index).BackColor
wsockSERVER.SendData "color;" & lblColor(Index).BackColor & vbCr
End Sub


Private Sub picME_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
mDown = True
oldX = X
oldY = Y
DoEvents
On Error Resume Next
DoEvents
wsockSERVER.SendData "draw;" & oldX & ":" & oldY & ":" & X & ":" & Y & vbCr: DoEvents
DoEvents
picME.Line (oldX, oldY)-(X, Y)
DoEvents
End Sub


Private Sub picME_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If mDown = False Then Exit Sub
 On Error Resume Next
 wsockSERVER.SendData "draw;" & oldX & ":" & oldY & ":" & X & ":" & Y & vbCr
 DoEvents
 picME.Line (oldX, oldY)-(X, Y)
 oldX = X
 oldY = Y
End Sub


Private Sub picME_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
mDown = False
End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtChat_Change()
txtChat.SelStart = Len(txtChat.Text)
txtChat.SelLength = Len(txtChat.Text)
End Sub

Private Sub txtChat_GotFocus()
txtMe.SetFocus
End Sub


Private Sub txtMe_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If txtMe.Text = "/clear" Then txtChat.Text = "": txtMe.Text = "": Exit Sub
 On Error Resume Next
 wsockSERVER.SendData "chat;" & txtMe.Text & vbCr
 txtChat.Text = txtChat.Text & txtNick.Text & " - " & txtMe.Text & vbCr & Chr$(10)
 If txtMe.List(txtMe.ListCount - 1) <> txtMe.Text Then txtMe.AddItem txtMe.Text
 txtMe.Text = ""
End If
End Sub


Private Sub txtNick_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 On Error Resume Next
 txtNick.Tag = txtNick.Text
 wsockSERVER.SendData "chnick;" & txtNick.Text & vbCr
End If
End Sub


Private Sub txtWidth_Change()
picME.DrawWidth = Val(txtWidth.Text)
On Error Resume Next
wsockSERVER.SendData "width;" & txtWidth.Text & vbCr
End Sub


Private Sub wserver_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

End Sub

Private Sub wsockSERVER_Close()
lblST.Caption = "Idle"
txtChat.Text = txtChat.Text & "The other Side Closed." & vbCr & Chr$(10)
lblST.Caption = "Idle"
wsockSERVER.Close
cmdConnect.Caption = "Connect"
cmdServer.Caption = "Server"
cmdConnect.Enabled = True
cmdServer.Enabled = True
End Sub

Private Sub wsockSERVER_ConnectionRequest(ByVal requestID As Long)
lblColor(0).BorderStyle = 1
txtWidth.Text = "1"
picME.Cls
picOther.Cls
wsockSERVER.Close
wsockSERVER.Accept requestID
lblST.Caption = "Connected with " & wsockSERVER.RemoteHostIP & "..."
End Sub

Private Sub wsockSERVER_DataArrival(ByVal bytesTotal As Long)
Dim vtData As String
Dim curPOS As Single
Dim MESSAGE As String
Dim COMM As String
wsockSERVER.GetData vtData
Debug.Print vtData
curPOS = 0
Do
   COMM = ""
   MESSAGE = ""
   Do
    curPOS = curPOS + 1
    rTMP = Left(vtData, curPOS)
    rTMP = Right(rTMP, 1)
    If rTMP = ";" Then Exit Do
    COMM = COMM & rTMP
   Loop Until curPOS >= bytesTotal
   COMM = LCase$(COMM)
   Do
    curPOS = curPOS + 1
    rTMP = Left(vtData, curPOS)
    rTMP = Right(rTMP, 1)
    If rTMP = Chr$(13) Then Exit Do
    MESSAGE = MESSAGE & rTMP
   Loop Until curPOS >= bytesTotal
   
   If Tag = "" And Tag <> txtNick.Text And COMM <> "nick" Then
    wsockSERVER.Close
    cmdConnect.Caption = "Connect"
    cmdServer.Caption = "Server"
    cmdServer.Enabled = True
    cmdConnect.Enabled = True
    lblST.Caption = "User sent commands before his Identification."
    Exit Sub
   End If

  If COMM = "nick" Then
   If LCase$(MESSAGE) = LCase$(txtNick.Text) Or MESSAGE = "" Then
      txtChat.Text = txtChat.Text & "Nick Colission Detected." & vbCr & Chr$(10)
      wsockSERVER.SendData "nonick;" & vbCr
      Exit Sub
   End If
   txtChat.Text = txtChat.Text & MESSAGE & " has joined DrawChat." & vbCr & Chr$(10)
   If Tag = "" Then wsockSERVER.SendData "nick;" & txtNick.Text & vbCr
   Tag = MESSAGE
  
  ElseIf COMM = "nonick" Then
    cmdConnect.Caption = "Connect"
    cmdServer.Caption = "Server"
    cmdConnect.Enabled = True
    cmdServer.Enabled = True
    If Tag = "" Or Tag = txtNick.Text Then
        txtChat.Text = txtChat.Text & "Nick Colission Detected." & vbCr & Chr$(10)
        wsockSERVER.Close
        lblST.Caption = "Your nickname is invalid. Please choose onother one..."
        txtNick.SetFocus
        Exit Sub
    End If
    txtNick.Text = txtNick.Tag
    lblST.Caption = "Your nickname is invalid. Please choose onother one..."
    txtNick.SetFocus
  
  ElseIf COMM = "chnick" Then
   If LCase$(MESSAGE) = LCase$(txtNick.Text) Or MESSAGE = "" Then
      txtChat.Text = txtChat.Text & "Nick Colission Detected." & vbCr & Chr$(10)
      wsockSERVER.SendData "nonick;" & vbCr
    Else
      txtChat.Text = txtChat.Text & Tag & " is now known as " & MESSAGE & "..." & vbCr & Chr$(10)
      Tag = MESSAGE
  End If
  
  ElseIf COMM = "chat" Then
   txtChat.Text = txtChat.Text & Tag & " - " & MESSAGE & vbCr & Chr$(10)
  
  ElseIf COMM = "draw" Then
    xS = Val(GetPiece(MESSAGE, ":", 1))
    yS = Val(GetPiece(MESSAGE, ":", 2))
    xE = Val(GetPiece(MESSAGE, ":", 3))
    yE = Val(GetPiece(MESSAGE, ":", 4))
    picOther.Line (xS, yS)-(xE, yE)
    DoEvents
   
   ElseIf COMM = "stmsg" Then
    lblcaption = MESSAGE
   
   ElseIf COMM = "width" Then
    picOther.DrawWidth = Val(MESSAGE)
   
   ElseIf COMM = "clear" Then
    picOther.Cls
    
   ElseIf COMM = "color" Then
    picOther.ForeColor = Val(MESSAGE)
   
   End If
Loop Until curPOS >= bytesTotal

End Sub
Function GetPiece(from As String, delim As String, Index) As String
    Dim temp$
    Dim Count
    Dim Where
    '
    temp$ = from & delim
    Where = InStr(temp$, delim)
    Count = 0
    Do While (Where > 0)
        Count = Count + 1
        If (Count = Index) Then
            GetPiece = Left$(temp$, Where - 1)
            Exit Function
        End If
        temp$ = Right$(temp$, Len(temp$) - Where)
        Where = InStr(temp$, delim)
    DoEvents
    Loop
    If (Count = 0) Then
        GetPiece = from
    Else
        GetPiece = ""
    End If
End Function




