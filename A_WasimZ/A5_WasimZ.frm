VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Riverdale Lottery 6/49"
   ClientHeight    =   3405
   ClientLeft      =   1890
   ClientTop       =   1905
   ClientWidth     =   5250
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
   ScaleHeight     =   3405
   ScaleWidth      =   5250
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Play"
      Height          =   495
      Left            =   1800
      TabIndex        =   9
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton cmdWinNum 
      Caption         =   "&Winning Numbers"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Winning Numbers:"
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblBank 
      Caption         =   "$20.00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "Bank:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label lblWinNums 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   1440
      Width           =   3735
   End
   Begin VB.Label lblWinLoss 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   4815
   End
   Begin VB.Label lblPlayNums 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Numbers:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programmer: Zuhab Wasim 11G
'Date: 25/01/16
'Purpose: To create a valid program, capable of displaying 6 unique random integers and matching them with 6 other unique random integers

Option Explicit

'Declares the constants that do not change within the program but will have a chance to change in the future
Const MAX = 6
Const LOW = 1
Const HIGH = 49

'Declares the variables needed within the entire program (form variables)
Dim PlayNum(1 To MAX) As Integer
Dim PlayUnique As Boolean

Dim WinNum(1 To MAX) As Integer
Dim WinUnique As Boolean

Dim Bank As Long
Dim LoopCount As Integer

'cmdExit: Procedure used for exiting the program
Private Sub cmdExit_Click()
    
    'Asks the user to confirm their click of cmdExit, exiting the program if they say yes
    If MsgBox("Are you sure you want to exit?", vbYesNo, "Exit") = vbYes Then
        End
    End If
    
End Sub

'cmdPlay: Procedure used to create and display 6 unique matching integers that act as the user's lottery number
Private Sub cmdPlay_Click()

    'Local variables needed for the procedure are declared
    Dim K As Integer
    Dim StrPlayNum As String
    
    'Values that will not have proper assignment and have their value manipulated are initiliazed
    StrPlayNum = ""
    
    'Checks to see if Bank is greater than 2, denying the runthrough of this procedure if Bank is less than or equal to 2
    If Bank >= 2 Then
        Bank = Bank - 2
        For K = 1 To MAX
            'Assigns the value of the local variable K to the general variable LoopCount, used in the CheckUnique procedure
            LoopCount = K
            Do
                PlayNum(K) = Int(Rnd() * (HIGH - LOW + 1)) + LOW
                CheckUnique
            'Continues the loop until CheckUnique returns True (Number does not match any other number in the Play Number)
            Loop Until PlayUnique
            'Adds an extra space to numbers below 10 to make them 2 characters long (i.e. "9" becomes " 9")
            If PlayNum(K) < 10 Then
                StrPlayNum = StrPlayNum & " " & Str$(PlayNum(K)) & " "
            Else
                StrPlayNum = StrPlayNum & Str$(PlayNum(K)) & " "
            End If
        Next K
        
        'Displays Play Number and the new value of Bank, as well as clearing the previous WinNum
        lblWinNums.Caption = ""
        lblPlayNums.Caption = StrPlayNum
        lblBank.Caption = Format$(Bank, "$00.00")
        
        'Disables cmdPlay and enables cmdWinNum to not allow the repitition of the cmdPlay procedure
        cmdPlay.Enabled = False
        cmdWinNum.Enabled = True
    Else
        'States the user does not have enough money to proceed with purchasing a play
        If MsgBox("You do not have sufficient funds to continue playing the game. Do you want to exit?", vbYesNo, "Game Over") = vbYes Then
            End
        End If
    End If
    
End Sub

'FindMatch: Procedure used to find the amount of matches between the Winning Number WinNum and the Play Number Play Num
Private Sub FindMatch()
    
    'Local variables needed for the procedure are declared
    Dim K As Integer
    Dim X As Integer
    Dim MatchCount As Integer
    Dim WinningMessage As String
    
    'Values that will not have proper assignment and have their value manipulated are initiliazed
    MatchCount = 0
    
    'Loops through each number with every number using two loops, incrementing MatchCount by 1 once a match is found
    For K = 1 To MAX
        For X = 1 To MAX
            If PlayNum(K) = WinNum(X) Then
                MatchCount = MatchCount + 1
            End If
        Next X
    Next K
    
    'Displays a unique and appropriate message for all winnings (1 To 6) and incrememts Bank appropriately
    If MatchCount = 6 Then
        Bank = Bank + 20000000
        WinningMessage = "Congratulations! You've won $20,000,000! (6 matches)"
    ElseIf MatchCount = 5 Then
        Bank = Bank + 1000000
        WinningMessage = "Amazing! You've won $1,000,000! (5 matches)"
    ElseIf MatchCount = 4 Then
        Bank = Bank + 100
        WinningMessage = "Awesome! You've won $100! (4 matches)"
    ElseIf MatchCount = 3 Then
        Bank = Bank + 5
        WinningMessage = "Cool! You've won $5! (3 matches)"
    ElseIf MatchCount = 2 Then
        WinningMessage = "Better Luck Next Time! (2 matches)"
    ElseIf MatchCount = 1 Then
        WinningMessage = "Better Luck Next Time! (1 match)"
    Else
        WinningMessage = "Better Luck Next Time! (0 matches)"
    End If
    
    'Displays the message previously derived
    lblWinLoss.Caption = WinningMessage
    lblBank.Caption = Format$(Bank, "$00.00")
    
End Sub

'cmdWinNum: Procedure used to create and display 6 unique matching integers that act as the winning lottery number
Private Sub cmdWinNum_Click()
    
    'Local variables needed for the procedure are declared
    Dim K As Integer
    Dim StrWinNum As String
    
    'Values that will not have proper assignment and have their value manipulated are initiliazed
    StrWinNum = ""
    
    For K = 1 To MAX
        'Assigns the local variable K to the form variable LoopCount used for the CheckUnique procedure
        LoopCount = K
        Do
            WinNum(K) = Int(Rnd() * (HIGH - LOW + 1)) + LOW
            CheckUnique
        'Repeats the loop untill WinUnique is true
        Loop Until WinUnique
        'Adds an extra space to numbers below 10 to make them 2 characters long (i.e. "9" becomes " 9")
        If WinNum(K) < 10 Then
            StrWinNum = StrWinNum & " " & Str$(WinNum(K)) & " "
        Else
            StrWinNum = StrWinNum & Str$(WinNum(K)) & " "
        End If
    Next K
     
    'Calls the FindMatch procedure to determine the amount of matches between WinNum and PlayNum
    FindMatch
    
    'Displays the winning number
    lblWinNums.Caption = StrWinNum
    
    'Disables cmdWinNum and enables cmdPlayNum to not allow consecutive repeats of the WinNum
    cmdPlay.Enabled = True
    cmdWinNum.Enabled = False
    
End Sub
Private Sub CheckUnique()
    
    'Local variables needed for the procedure are declared
    Dim K As Integer
    Dim WNumCount As Integer
    Dim PNumCount As Integer
    
    'Values that will not have proper assignment and have their value manipulated are initiliazed
    WNumCount = 0
    PNumCount = 0
    
    'Checks to see if WinNum or PlayNumat index K, increments the counter by 1 if so
    For K = 1 To MAX
        If WinNum(LoopCount) = WinNum(K) Then
            WNumCount = WNumCount + 1
        End If
        If PlayNum(LoopCount) = PlayNum(K) Then
            PNumCount = PNumCount + 1
        End If
    Next K
    
    'If the number of times WinNum or PlayNum repeated is more than 1, it is no unique, therefore Unique becomes false
    'Otherwise, Unique is true
    If WNumCount > 1 Then
        WinUnique = False
    Else
        WinUnique = True
    End If

    If PNumCount > 1 Then
        PlayUnique = False
    Else
        PlayUnique = True
    End If
    
End Sub
Private Sub Initiliaze()
    
    'Declared local variables needed for this procedure
    Dim K As Integer
    
    'Initiliazes all variables and arrays as either 0 or 20 in the case of bank
    For K = 1 To MAX
        PlayNum(K) = 0
        WinNum(K) = 0
    Next K
    
    Bank = 20
    
    'Displays bank in lblBank.Caption at start up of the program
    lblBank.Caption = Format$(Bank, "$00.00")
    
End Sub

Private Sub Form_Load()
    
    'Randmoizes the values generated for the Rnd() function during the execution of the program
    Randomize
    
    'Calls the initiliaze procedure to intiliaze all variables
    Initiliaze
    
    'Initially sets cmdWinNum as false untill cmdPlay is clicked
    cmdWinNum.Enabled = False
    
End Sub
