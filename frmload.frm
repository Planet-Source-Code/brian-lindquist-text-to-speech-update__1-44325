VERSION 5.00
Begin VB.Form frmload 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Load File"
   ClientHeight    =   3225
   ClientLeft      =   3405
   ClientTop       =   1065
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   8745
   Begin VB.TextBox txtfile 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2640
      Width           =   2655
   End
   Begin VB.CommandButton cmdload 
      Caption         =   "&Load File"
      Default         =   -1  'True
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
      Left            =   7080
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Left            =   4320
      Pattern         =   "*.txt*"
      TabIndex        =   2
      Top             =   1080
      Width           =   4095
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   4095
   End
   Begin VB.DriveListBox drv1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "3. Select File  (Must be a .txt file)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5280
      TabIndex        =   5
      Top             =   840
      Width           =   2910
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "2. Select Directory"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   4
      Top             =   840
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "1. Select Drive"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1260
   End
End
Attribute VB_Name = "frmload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdload_Click()
On Error Resume Next
frmmain.rtb1.LoadFile (Dir1.Path) + (File1.FileName)
End Sub

Private Sub Dir1_Change()
On Error Resume Next
File1.Path = Dir1.Path
End Sub

Private Sub drv1_Change()
On Error Resume Next
Dir1.Path = drv1.Drive
End Sub

Private Sub File1_Click()
On Error Resume Next
txtfile.Text = (Dir1.Path) + (File1.FileName)
End Sub

Private Sub Form_Load()
On Error Resume Next
Dir1.Path = drv1.Drive
File1.Path = Dir1.Path
End Sub
