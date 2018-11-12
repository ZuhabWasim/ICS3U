VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   4050
   ClientLeft      =   3810
   ClientTop       =   2445
   ClientWidth     =   7200
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Bingo - About Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   7200
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   6120
      TabIndex        =   6
      Top             =   3480
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Clicks"
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "43"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFC0&
      Caption         =   "© Zuhab Wasim 2016       Ver. 2.0.00"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3720
      Width           =   4815
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   360
      Picture         =   "Bingo - About Form.frx":030A
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   600
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Your goal is to get BINGO in the lowest amount of clicks. Good Luck!"
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   2760
      Width           =   5415
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   $"Bingo - About Form.frx":0614
      Height          =   975
      Left            =   1560
      TabIndex        =   2
      Top             =   1560
      Width           =   5415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "How To Play"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "To start the game, click Reveal The Cards! in the File menu."
      Height          =   615
      Left            =   1680
      TabIndex        =   0
      Top             =   720
      Width           =   3855
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()

    'Unloads the form
    Unload frmAbout
    
End Sub

