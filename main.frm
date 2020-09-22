VERSION 5.00
Object = "{EEE78583-FE22-11D0-8BEF-0060081841DE}#1.0#0"; "XVoice.dll"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmmain 
   Caption         =   "Text-To-Speech For Windows 2000"
   ClientHeight    =   4380
   ClientLeft      =   2220
   ClientTop       =   1995
   ClientWidth     =   9300
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   9300
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox rtb1 
      Height          =   1695
      Left            =   840
      TabIndex        =   18
      Top             =   1680
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2990
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"main.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txttone 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Height          =   285
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "100"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtspeed 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Height          =   285
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "150"
      Top             =   120
      Width           =   495
   End
   Begin VB.CheckBox chkclear 
      Caption         =   "Auto Clear Text"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   11
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CheckBox chkontop 
      Caption         =   "Always On Top"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      TabIndex        =   10
      Top             =   2040
      Width           =   1215
   End
   Begin VB.HScrollBar tone 
      Height          =   255
      Left            =   6000
      Max             =   200
      Min             =   50
      TabIndex        =   7
      Top             =   600
      Value           =   100
      Width           =   3255
   End
   Begin VB.HScrollBar speed 
      Height          =   255
      Left            =   6000
      Max             =   450
      Min             =   30
      TabIndex        =   6
      Top             =   120
      Value           =   150
      Width           =   3255
   End
   Begin VB.CommandButton cmdresume 
      Caption         =   "R&esume"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdpause 
      Caption         =   "&Pause"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdstop 
      Caption         =   "&Reset"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "&Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdspeak 
      Caption         =   "&Speak"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
   Begin ACTIVEVOICEPROJECTLibCtl.DirectSS mouth 
      Height          =   1935
      Left            =   2280
      OleObjectBlob   =   "main.frx":04B9
      TabIndex        =   0
      Top             =   0
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controls"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   6480
      TabIndex        =   17
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Brianlinq1@cs.com"
      Height          =   255
      Left            =   7440
      TabIndex        =   16
      Top             =   4080
      Width           =   1425
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Brian Lindquist"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   15
      Top             =   3840
      Width           =   1425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Created By:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   14
      Top             =   3600
      Width           =   1425
   End
   Begin VB.Label lbltone 
      AutoSize        =   -1  'True
      Caption         =   "Voice Tone"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7200
      TabIndex        =   9
      Top             =   840
      Width           =   1200
   End
   Begin VB.Label lblspeed 
      AutoSize        =   -1  'True
      Caption         =   "Talking Speed"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7200
      TabIndex        =   8
      Top             =   360
      Width           =   1560
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuload 
         Caption         =   "&Load File"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub chkontop_Click()
On Error Resume Next
    If chkontop.Value = vbChecked Then
        Call putMeOnTop(Me)
    Else
        Call takeMeDown(Me)
    End If
End Sub
Private Sub cmdclear_Click()
On Error Resume Next
rtb1.Text = ""
End Sub
Private Sub cmdpause_Click()
On Error Resume Next
mouth.AudioPause
End Sub
Private Sub cmdresume_Click()
On Error Resume Next
mouth.AudioResume
End Sub
Private Sub cmdspeak_Click()
On Error Resume Next
If chkclear.Value = 1 Then
mouth.speed = txtspeed.Text
mouth.Pitch = txttone.Text
mouth.Speak (rtb1.Text)
rtb1.Text = ""
Else
mouth.speed = txtspeed.Text
mouth.Pitch = txttone.Text
mouth.Speak (rtb1.Text)
End If
End Sub
Private Sub cmdstop_Click()
On Error Resume Next
mouth.AudioReset
End Sub

Private Sub Form_Load()
'created by brian lindquist
'brianlinq1@cs.com
'please leave this credit here
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
Private Sub mnuexit_Click()
End
End Sub

Private Sub mnuload_Click()
frmload.Visible = True

End Sub

Private Sub speed_Change()
On Error Resume Next
txtspeed.Text = ""
txtspeed.Text = speed.Value
End Sub
Private Sub speed_Scroll()
On Error Resume Next
speed_Change
End Sub
Private Sub tone_Change()
On Error Resume Next
txttone.Text = ""
txttone.Text = tone.Value
End Sub
Private Sub tone_Scroll()
tone_Change
End Sub
