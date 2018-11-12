VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Tic-Tac-Toe"
   ClientHeight    =   3345
   ClientLeft      =   1890
   ClientTop       =   1905
   ClientWidth     =   6495
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   6495
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   735
      Left            =   5160
      TabIndex        =   20
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Totals"
      Height          =   1935
      Left            =   3360
      TabIndex        =   1
      Top             =   0
      Width           =   3015
      Begin VB.Label lblDraws 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   2040
         TabIndex        =   9
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblOWins 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   2040
         TabIndex        =   8
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblXWins 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   2160
         TabIndex        =   7
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblPlayedGames 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   2040
         TabIndex        =   6
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Draws:"
         Height          =   375
         Left            =   1080
         TabIndex        =   5
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Player O:"
         Height          =   255
         Left            =   720
         TabIndex        =   4
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Player X:"
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label label1 
         Caption         =   "Games Played:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3270
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3180
      Begin VB.Label lblSquare 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   39
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   8
         Left            =   2040
         TabIndex        =   19
         Top             =   2160
         Width           =   1005
      End
      Begin VB.Label lblSquare 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   39
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   7
         Left            =   1080
         TabIndex        =   18
         Top             =   2160
         Width           =   1005
      End
      Begin VB.Label lblSquare 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   39
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   6
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   1005
      End
      Begin VB.Label lblSquare 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   39
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   5
         Left            =   2040
         TabIndex        =   16
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Label lblSquare 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   39
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   4
         Left            =   1080
         TabIndex        =   15
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Label lblSquare 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   39
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Label lblSquare 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   39
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   2
         Left            =   2040
         TabIndex        =   13
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lblSquare 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   39
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   1
         Left            =   1080
         TabIndex        =   12
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lblSquare 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   39
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1000
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "RCI Inc.© 2016"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   2040
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Constant values are declared to allow ease of alteration of values in the future
Const MAX = 8

'Variables used throughout the form are declared in General Declarations
Dim Games As Integer
Dim XWins As Integer
Dim OWins As Integer
Dim Draws As Integer
Dim ClickNum As Integer
Dim Combination(0 To MAX) As String
Dim IsWin As String
Dim IsDraw As Boolean

Private Sub cmdExit_Click()

'Exit procedure variables and values declared
    Dim ExitMsgType
    Dim ExitMsg
    Dim ExitTitle
    Dim ExitResponse
    
    ExitMsgType = vbYesNo + vbExclamation
    ExitMsg = "Are you sure you want to exit?"
    ExitTitle = "Exit Now?"
    
'User's response of exiting the program is recorded
    ExitResponse = MsgBox(ExitMsg, ExitMsgType, ExitTitle)
    
'Exits the program if the user has confirmed their press of cmdExit
    If ExitResponse = vbYes Then
        End
    End If
    
End Sub

Private Sub Form_Load()
    
    Dim K As Integer

'Variables that require initialization during the program are initialized
    XWins = 0
    OWins = 0
    Draws = 0
    ClickNum = 0
    
    For K = 0 To MAX
        Combination(K) = " "
        lblSquare(K).Caption = " "
    Next K
    
End Sub

Private Sub lblSquare_Click(Index As Integer)
    
    'Increments click value by 1 when a square is clicked
    ClickNum = ClickNum + 1
    
    'Determines which player is playing; odd numbers are X and even numbers are O
    'Places an X or O in the given control array depending on the player
    If ClickNum Mod 2 = 1 Then
        lblSquare(Index).Caption = "X"
    Else
        lblSquare(Index).Caption = "O"
    End If
    
    'Disables the clicked square to not allow alterations of the same square
    lblSquare(Index).Enabled = False
    
    'Checks to see if the game is on turn 5
    'Only at turns 5 or greater is there a need to check if there is a win
    If ClickNum >= 5 Then
        'Obtains every combination horizontall
        GetCombination
        CheckWin
        If ClickNum >= 9 Then
            If IsWin = " " Then
                IsDraw = True
            End If
            EndGame
        Else
            If IsWin <> " " Then
                If IsWin = "X" Then
                    XWins = XWins + 1
                Else
                    OWins = OWins + 1
                End If
                EndGame
            End If
        End If
    End If
        
End Sub

Private Sub EndGame()
    
    Dim MsgType As Integer
    Dim MsgTitle As String
    Dim Msg As String
    Dim Response As Integer
    Dim K As Integer
    
    MsgType = vbYesNo + vbExclamation
    If IsDraw Then
        MsgTitle = "It's a draw!"
        Draws = Draws + 1
        Msg = "The game has ended in a draw. Would you like to play again?"
        IsDraw = False
    Else
        MsgTitle = "Congratulations Player " & IsWin & "!"
        Msg = "Player " & IsWin & " has won. Would you like to play again?"
    End If
    
    Response = MsgBox(Msg, MsgType, MsgTitle)
    
    If Response = vbYes Then
        For K = 0 To MAX
            Combination(K) = " "
            lblSquare(K).Caption = " "
            lblSquare(K).Enabled = True
        Next K
        Games = Games + 1
        lblPlayedGames.Caption = Str(Games)
        lblXWins.Caption = Str(XWins)
        lblOWins.Caption = Str(OWins)
        lblDraws.Caption = Str(Draws)
        ClickNum = 0
    Else
        End
    End If
    
End Sub


Private Sub GetCombination()
    
    Dim K As Integer
    Dim Count As Integer
    
    Count = 0
    
    For K = 0 To 6 Step 3
        Count = Count + 1
        Combination(Count) = lblSquare(K).Caption & lblSquare(K + 1).Caption & lblSquare(K + 2).Caption
    Next K
    
    For K = 0 To 2
        Count = Count + 1
        Combination(Count) = lblSquare(K).Caption & lblSquare(K + 3).Caption & lblSquare(K + 6).Caption
    Next K
    
    For K = 0 To 2 Step 2
        Count = Count + 1
        If K = 0 Then
            Combination(Count) = lblSquare(K).Caption & lblSquare(K + 4).Caption & lblSquare(K + 8).Caption
        Else
            Combination(Count) = lblSquare(K).Caption & lblSquare(K + 2).Caption & lblSquare(K + 4).Caption
        End If
    Next K
    
End Sub

Private Sub CheckWin()
    
    Dim K As Integer
    Dim OCount As Integer
    Dim XCount As Integer
    
    OCount = 0
    XCount = 0
    
    For K = 0 To 8
        If Combination(K) = "OOO" Then
            OCount = OCount + 1
        ElseIf Combination(K) = "XXX" Then
            XCount = XCount + 1
        End If
    Next K
    'Note: Once a combination of "OOO" or "XXX" is found, put end game after it be
    If OCount > 0 Then
        IsWin = "O"
    ElseIf XCount > 0 Then
        IsWin = "X"
    Else
        IsWin = " "
    End If
    
End Sub
