VERSION 5.00
Begin VB.Form frmScores 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HighScores"
   ClientHeight    =   5640
   ClientLeft      =   3600
   ClientTop       =   2865
   ClientWidth     =   7530
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Bingo - HighScores Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   7530
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   3
      Top             =   5040
      Width           =   1455
   End
   Begin VB.PictureBox picScores 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      FillColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3855
      Left            =   120
      ScaleHeight     =   3795
      ScaleWidth      =   7155
      TabIndex        =   2
      Top             =   1080
      Width           =   7215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "HighScores"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "These are the highscores achieved from players."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   5775
   End
End
Attribute VB_Name = "frmScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdOK_Click()
    
    'Unloads the form
    Unload frmScores
    
End Sub

Private Sub Form_Load()
    
    'Declares local variable K
    Dim K As Integer
    
    'Clears the the picture box
    picScores.Cls
    'Prints the heading
    picScores.Print Spc(6); "Name:"; Tab(40); "Score:"
    
    'Obtains the updated highscores
    GetHighScores
    
    'Displays all highscore entries
    For K = 1 To HIGHSCORE_MAX
        picScores.Print K; ". "; HighScoreNames(K); Tab(40); Format$(HighScores(K), "@@@@")
    Next K
    
End Sub
