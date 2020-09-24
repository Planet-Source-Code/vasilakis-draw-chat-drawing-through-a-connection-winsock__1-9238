VERSION 5.00
Begin VB.Form frmSup 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4935
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrShow 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   4200
      Top             =   1560
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   3060
      Left            =   120
      Picture         =   "frmSup.frx":0000
      ScaleHeight     =   3000
      ScaleWidth      =   1500
      TabIndex        =   1
      Top             =   120
      Width           =   1560
   End
   Begin VB.Label lblBy 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "By Vasilis Sagonas"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   2640
      TabIndex        =   2
      Top             =   2640
      Width           =   1410
   End
   Begin VB.Label lblSup 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chat v1.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   2325
   End
   Begin VB.Shape Shape2 
      DrawMode        =   9  'Not Mask Pen
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   2160
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Shape Shape3 
      DrawMode        =   9  'Not Mask Pen
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   960
      Shape           =   2  'Oval
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label llbText 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSup.frx":288A
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   1455
      Left            =   2040
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape1 
      DrawMode        =   14  'Copy Pen
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   2535
      Left            =   1800
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "frmSup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
CF Me
iWidth = Width
Width = 1
Show
For I = 1 To iWidth Step 10
    Width = I
    DoEvents
Next I
tmrShow.Enabled = True
End Sub

Private Sub tmrShow_Timer()
Load frmMain
frmMain.Show
Unload Me
End Sub


