VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hockey Player Information"
   ClientHeight    =   3510
   ClientLeft      =   1815
   ClientTop       =   1830
   ClientWidth     =   8295
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
   MaxButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   8295
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   495
      Left            =   3840
      TabIndex        =   11
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "&Previous"
      Height          =   495
      Left            =   2520
      TabIndex        =   10
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open"
      Height          =   495
      Left            =   1320
      TabIndex        =   8
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label lblTPoints 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Left            =   6960
      TabIndex        =   18
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label lblPlayerAssists 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Left            =   6960
      TabIndex        =   17
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblPlayerGoals 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Left            =   6960
      TabIndex        =   16
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblPlayedGames 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Left            =   6960
      TabIndex        =   15
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblTeamName 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   14
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label lblPlayerName 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   13
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "NHL Player Stats"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   12
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label8 
      Caption         =   "Player:"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Total Points:"
      Height          =   255
      Left            =   5280
      TabIndex        =   6
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Assists:"
      Height          =   255
      Left            =   5280
      TabIndex        =   5
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Goals:"
      Height          =   255
      Left            =   5280
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Games Played:"
      Height          =   255
      Left            =   5280
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Team Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Full Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblCurrentPlayer 
      Height          =   255
      Left            =   4200
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Programmer: Zuhab Wasim 11G 322968165
'Date: 08/12/2015
'Purpose: To read a file containing the player names and their corresponding stats in their league
'         and displays them with the ability to scroll through each player out of the total amount of players


'------------------General Declarations used by the entire program------------------

'Finalizes the constants used within the program to ensure easy alterations in the future
Const INPUTFILE = "NHLStats.txt"
Const MAX = 100

'Declares the variables/arrays used within the program
'Declares the variables without any use but output
Dim PlayerName(1 To MAX) As String
Dim TeamName(1 To MAX) As String
Dim PlayedGames(1 To MAX) As Integer

'Declares the variables that are used within calculations
Dim PlayerGoals(1 To MAX) As Integer
Dim PlayerAssists(1 To MAX) As Integer
Dim PlayerTPoints(1 To MAX) As Integer

'Declares the variables used by the program, unkown to the user
Dim PlayerCount As Integer
Dim PlayerAmount As Integer
Dim PlayerNum As Integer

Private Sub cmdExit_Click()

'------------------The end of the execution of the program------------------
    
    'Asks the user to confirm their click of the exit button
    If MsgBox("Are you sure you want to exit?", vbYesNo, "Exit") = vbYes Then
        End
    End If
    
End Sub

Private Sub cmdNext_Click()

'------------------The Next Button to scroll to the next player and their stats------------------

    'Increments PlayerNum by 1 to have the index of the next player in the list
    PlayerNum = PlayerNum + 1
    
    'Checks to see if the Player currently displayed is the last,
    'Disables the next button to prevent the user to go past the amount of players
    If PlayerNum = PlayerAmount Then
        cmdNext.Enabled = False
    End If
    
    cmdPrevious.Enabled = True
    
    'Calls the Display() procedure to display the player at the index of PlayerNum
    Display
    
End Sub

Private Sub Display()

'------------------Displays the information of the player with the index of PlayerNum------------------

    'Displays the stats of the player
    lblPlayerName.Caption = " " & PlayerName(PlayerNum)
    lblTeamName.Caption = " " & TeamName(PlayerNum)
    lblPlayedGames.Caption = Str$(PlayedGames(PlayerNum))
    
    lblPlayerGoals.Caption = Str$(PlayerGoals(PlayerNum))
    lblPlayerAssists.Caption = Str$(PlayerAssists(PlayerNum))
    lblTPoints.Caption = Str$(PlayerTPoints(PlayerNum))
    
    'Displays the the current player being displayed and the total amount of players read by the file
    lblCurrentPlayer.Caption = Str$(PlayerNum) & " of" & Str$(PlayerAmount)
    
End Sub

Private Sub cmdOpen_Click()

'------------------Reads the data on the input file------------------

    'Declares the local variables used within the procedure
    Dim PathLocation As String
    Dim K As Integer
    Dim X As Integer
    
    'Checks to see if the file location string contains a "\" at the end, adds one otherwise
    If Right$(App.Path, 1) <> "\" Then
        PathLocation = App.Path & "\" & INPUTFILE
    Else
        PathLocation = App.Path & INPUTFILE
    End If
   
    'Initialized the variables used that are incremented later on to avoid errors
    K = 0
    'Opens the file used as labeled as 1
    Open PathLocation For Input As #1
    
    'Adds input from the file to arrays as long as it is not the end of file
    Do While Not EOF(1)
        K = K + 1
        Input #1, PlayerName(K)
        Input #1, TeamName(K)
        Input #1, PlayedGames(K)
        Input #1, PlayerGoals(K)
        Input #1, PlayerAssists(K)
        PlayerTPoints(K) = PlayerGoals(K) + PlayerAssists(K)
    Loop
    
    'Closes the file that has been read
    Close #1
    
    'Assigns the value of the local variable K to the form variable PlayerAmount for usage in other procedures
    PlayerAmount = K
    
    'Calls the Display() procedure to display the player at the index of PlayerNum
    Display
        
    cmdOpen.Enabled = False
    'Predisables the Previous Button to ensure the User does not go beyond the first player
    cmdPrevious.Enabled = False
    'Enables the Next Button for the use of scrolling
    cmdNext.Enabled = True
    
    Exit Sub
    
End Sub

Private Sub cmdPrevious_Click()

'------------------The Previous Button to display the previous player and their stats------------------

    'Decreases PlayerNum by 1 to display the previous player's stats
    PlayerNum = PlayerNum - 1
    
    'Checks to see if the current player displayed is the second to first player
    'Disables the Previous Button to ensure the user does not go beyond the first player given
    If PlayerNum = 1 Then
        cmdPrevious.Enabled = False
    End If
    
    cmdNext.Enabled = True
    
    'Calls the Display() procedure to display the player at the index of PlayerNum
    Display
    
End Sub

Private Sub Form_Load()

'------------------Initialization of the variables/arrays used within the program during start up------------------

'Declares the local variables used within the procedure
    Dim K As Integer

'Assigns all the form variables to 0, using a loop for arrays
'Initialized the variable in order to display the information of PlayerNum 1 for the begining of the program
    PlayerNum = 1
    PlayerAmount = 0
    For K = 1 To MAX
        PlayerName(K) = "No Data"
        TeamName(K) = "No Data"
        PlayedGames(K) = 0
        PlayerGoals(K) = 0
        PlayerAssists(K) = 0
        PlayerTPoints(K) = 0
    Next K
    
    cmdNext.Enabled = False
    cmdPrevious.Enabled = False
    
End Sub

