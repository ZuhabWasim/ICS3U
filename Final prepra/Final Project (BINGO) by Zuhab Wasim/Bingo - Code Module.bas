Attribute VB_Name = "Module1"
Option Explicit

'Declares global constants used in both the main form and the module, the maximum amount of numbers and the maximum amount of bingo numbers per board
Global Const NUM_MAX = 75
Global Const BINGO_MAX = 50

'Declares the file name and the maximum highscore entries globally
Global Const FNAME = "HighScores.txt"
Global Const HIGHSCORE_MAX = 5

'Declares the first and last values of each board globally
Global Const BOARD1_FIRST = 1
Global Const BOARD1_LAST = 5
Global Const BOARD2_FIRST = 26
Global Const BOARD2_LAST = 30

'Declares constants used in the module only, a temporary max value and a low constant for loops
Const TEMP_MAX = 15
Const LOW = 1

'Declares highscore names and values globally
Global HighScores(1 To HIGHSCORE_MAX) As Integer
Global HighScoreNames(1 To HIGHSCORE_MAX) As String

Public Sub GetHighScores()
    
    'Declares local variable K
    Dim K As Integer
    
    'Opens the file for reading
    Open App.Path & "\" & FNAME For Input As #1
    
    'Assigns the value of the sequential file into the variables
    For K = 1 To HIGHSCORE_MAX
        Input #1, HighScoreNames(K), HighScores(K)
    Next K
    
    'Closes the file
    Close #1
    
End Sub
Public Sub CreateFile()
    
    'Declares local variable K
    Dim K As Integer
    
    'Opens file for reading
    Open App.Path & "\" & FNAME For Output As #1
    
    'Creates five entries of default highscores to be beaten by the user
    For K = 1 To HIGHSCORE_MAX
        Write #1, "Anonymous", 999
    Next K
    
    'Closes the file
    Close #1
    
End Sub

Public Sub ChangeHighScores(ByVal S As Integer, ByVal SN As String)
    
    'Declares local variable K
    Dim K As Integer
    
    'Swaps the names and values of highschool entries with the user's highscore of the highscore entry at K
    For K = 1 To HIGHSCORE_MAX
        If S < HighScores(K) Then
            SwapName SN, HighScoreNames(K)
            SwapNum S, HighScores(K)
        End If
    Next K
    
    'Opens file for reading
    Open App.Path & "\" & FNAME For Output As #1
    
    'Writes down the updated highscore array into the text file
    For K = 1 To HIGHSCORE_MAX
        Write #1, HighScoreNames(K), HighScores(K)
    Next K
    
    'Closes the file
    Close #1
    
End Sub

Public Sub SwapName(N1 As String, N2 As String)
    
    'Declares local variable
    Dim Temp As String
    
    'Swaps the string values of N1 and N2
    Temp = N1
    N1 = N2
    N2 = Temp
    
End Sub

Public Sub SwapNum(N1 As Integer, N2 As Integer)
    
    'Declares local variable
    Dim Temp As Integer
    
    'Swaps the integer values of N1 and N2
    Temp = N1
    N1 = N2
    N2 = Temp
    
End Sub

Public Sub WriteFile()
    
    'Declares local variable
    Dim K As Integer
    
    'Opens file for reading
    Open App.Path & "\" & FNAME For Output As #1
    
    'Writes the updated entries
    For K = 1 To HIGHSCORE_MAX
        Write #1, HighScoreNames(K), HighScores(K)
    Next K
    
    'Closes the file
    Close #1
    
End Sub
Public Sub EndGame(ByVal N As Integer)
    
    'Declares variables used to display EndGame message
    Dim WinMsg As String
    Dim WinType As Integer
    Dim WinTitle As String
    
    
    frmMain.tmrTimer.Enabled = False
    
    'Assigns the appropriate messages and values to the EndGame display variables
    If frmMain.tmrTimer.Enabled Then
        WinMsg = "Congratulations! You have got BINGO in " & Str$(N) & " turns manually! Thanks for playing!"
    Else
        WinMsg = "Congratulations! You have got BINGO in " & Str$(N) & " turns by auto clicking! Thanks for playing!"
    End If
    WinType = vbOKOnly + vbInformation
    WinTitle = "WINNER!"
    
    'Displays the message box with the End Game variables
    MsgBox WinMsg, WinType, WinTitle
    
    'Disables the frame of winning numbers to avoid clicking after win
    frmMain.fraWinNum.Enabled = False
    frmMain.mnuNewGame.Enabled = True
    frmMain.mnuAuto.Enabled = False
    
End Sub
Public Sub RevealCards()
    
    'Local variable K declared for for/next loop
    Dim K As Integer
    
    'Turns images in imageboxes to invisible
    For K = 1 To BINGO_MAX
        'Checks to see if K is at the location of the center, assigns the Free image if so
        If K = 13 Or K = 38 Then
            frmMain.imgBoard(K).Picture = frmMain.imgFree.Picture
        Else
            frmMain.imgBoard(K).Visible = False
        End If
        frmMain.lblBoard(K).Visible = True
    Next K
    
    'Enables the frame for winning nums to allow the clicking of the winning numbers
    frmMain.fraWinNum.Enabled = True
    
End Sub
Public Sub ResetColours()
    
    'Local variable K declared for for/next loop
    Dim K As Integer
    
    'Loops to the amount of Bingo Numbers for each board and resets the colour of each cell
    For K = 1 To BINGO_MAX
        'Checks to see if K is at the location of the center, gives the cell the colour of already obtained because of it being Free
        If K = 13 Or K = 38 Then
            frmMain.lblBoard(K).BackColor = QBColor(10)
        Else
            frmMain.lblBoard(K).BackColor = &HC0FFFF
        End If
    Next K
    
    'Resets all colors of the winning numbers back to default
    For K = 1 To NUM_MAX
        frmMain.lblWinNum(K).BackColor = &HC0FFFF
    Next K
    
End Sub
Public Sub HideCards()
    
    'Local variable K declared for for/next loop
    Dim K As Integer
    
    'Loops to the amount of bingo numbers for each board and turns the image board visible
    For K = 1 To BINGO_MAX
        'Checks to see if K is at the location of the center, reassigns the image of bingoback to the image
        If K = 13 Or K = 38 Then
            frmMain.imgBoard(K).Picture = frmMain.imgBingoBack.Picture
        Else
            frmMain.imgBoard(K).Visible = True
            frmMain.lblBoard(K).Visible = False
        End If
    Next K
    
    'Turns all winning num card backs to visible
    For K = 1 To NUM_MAX
        frmMain.imgWinNum(K).Visible = True
    Next K
    
    'Disables the option of clicking the winning numbers
    frmMain.fraWinNum.Enabled = False
    
End Sub
Public Sub Initialize(ByRef CV As Integer, ByRef MCl As Integer, ByRef WN() As Integer, ByRef MC() As Integer)
    
    'Declares local variables used in procedure
    Dim K As Integer
    
    'Initializes all variables passed by reference as 0 or 1
    CV = 0
    MCl = 0
    For K = 1 To NUM_MAX
        WN(K) = K
    Next K
    For K = 1 To BINGO_MAX
        If K = 13 Or K = 38 Then
            MC(K) = 1
            frmMain.lblBoard(K).BackColor = QBColor(10)
        Else
            MC(K) = 0
        End If
    Next K
    'Resets the click count on the form
    frmMain.lblWinCount = "0"
    
End Sub
Public Sub GenWinNums(WN() As Integer)
    
    'Declares local variables needed for procedure
    Dim Num1, Num2, Temp As Integer
    Dim X As Integer
    
    'Loops to 100 and randomizes array by swapping two numbers at random indices
    For X = 1 To 100
        Num1 = Int(Rnd() * (NUM_MAX - LOW + 1)) + LOW
        Num2 = Int(Rnd() * (NUM_MAX - LOW + 1)) + LOW
        Temp = WN(Num1)
        WN(Num1) = WN(Num2)
        WN(Num2) = Temp
    Next X
    
    'Displays the newly randomized array in the lblWinNum(X) control arrays
    For X = 1 To NUM_MAX
        frmMain.lblWinNum(X).Caption = Str$(WN(X))
    Next X
    
End Sub

Public Sub GenBingoNums(Board As Variant, BNums() As Integer, ByVal L As Integer, ByVal H As Integer)
    
    'Declares local variables needed in procedure
    Dim W, X As Integer
    Dim Num1, Num2, Temp As Integer
    Dim Hi, Lo As Integer
    Dim Count As Integer
    Dim TempNum(1 To TEMP_MAX) As Integer
    
    'Initializes all alternating variables to certain values
    Count = 0
    
    Lo = 1
    Hi = 15

    'Loop goes by colums
    For W = L To H
        'Assigns 15 numbers from low (+ offset) and high (+ offset)
        For X = Lo To Hi
            Count = Count + 1
            TempNum(Count) = X
        Next X
        'Resets counter to zero
        Count = 0
        'Randomizes the 15 numbers within the array to ensure each number is unique
        'Using swapping method
        For X = 1 To 100
            Num1 = Int(Rnd() * (TEMP_MAX - LOW + 1)) + LOW
            Num2 = Int(Rnd() * (TEMP_MAX - LOW + 1)) + LOW
            Temp = TempNum(Num1)
            TempNum(Num1) = TempNum(Num2)
            TempNum(Num2) = Temp
        Next X
        'Displays the numbers in the appropriate cell and
        'assigns the first 5 values of TempNum() into the permanent array BingoNum1
        For X = 0 To 20 Step 5
            Count = Count + 1
            If (W + X) <> 13 And (W + X) <> 38 Then
                'Displays values of the BingoNum array by columns
                Board(W + X).Caption = Str$(TempNum(Count))
                BNums(W + X) = TempNum(Count)
            End If
        Next X
        Count = 0
        'Offsets the High and Low values by 15 for each column used in random number generation
        Lo = Lo + 15
        Hi = Hi + 15
    Next W

End Sub

Public Function CheckWin(Board As Variant, MC() As Integer, ByVal L As Integer, ByVal H As Integer) As Boolean
    
    'Declares local variables needed in procedure
    Dim X, Y As Integer
    Dim Sum As Integer
    
    'Goes through each column and checks win via a sum count, also checks diagonals for wins
    'Changes the background colour of the bingo cells that obtained the win
    For X = L To H
        Sum = MC(X) + MC(X + 5) + MC(X + 10) + MC(X + 15) + MC(X + 20)
        If Sum = 5 Then
            For Y = 0 To 20 Step 5
                Board(X + Y).BackColor = QBColor(13)
            Next Y
            CheckWin = True
        End If
        'Checks to see if X is either 1 or 5 to execute the diagonal check
        If X = 1 Or X = 26 Then
            Sum = MC(X) + MC(X + 6) + MC(X + 12) + MC(X + 18) + MC(X + 24)
            If Sum = 5 Then
                For Y = 0 To 24 Step 6
                    Board(X + Y).BackColor = QBColor(13)
                Next Y
                CheckWin = True
            End If
        ElseIf X = 5 Or X = 30 Then
            Sum = MC(X) + MC(X + 4) + MC(X + 8) + MC(X + 12) + MC(X + 16)
            If Sum = 5 Then
                For Y = 0 To 17 Step 4
                    Board(X + Y).BackColor = QBColor(13)
                Next Y
                CheckWin = True
            End If
        End If
    Next X
    
    'Goes through every row to check for win and
    'Changes the background colour of the bingo cells that obtained the win
    For X = L To (H + 16) Step 5
        Sum = MC(X) + MC(X + 1) + MC(X + 2) + MC(X + 3) + MC(X + 4)
        If Sum = 5 Then
            For Y = 0 To 4
                Board(X + Y).BackColor = QBColor(13)
            Next Y
            CheckWin = True
        End If
    Next X

End Function

Public Sub AssignImages()
    
    'Declares local variables needed
    Dim K As Integer
    
    'Assigns the images of each Bingo Cell
    For K = 1 To BINGO_MAX
        frmMain.imgBoard(K).Picture = frmMain.imgBingoBack.Picture
    Next K
    
    'Assigns the images of each Winning Number cell
    For K = 1 To NUM_MAX
        frmMain.imgWinNum(K).Picture = frmMain.imgWinBack.Picture
    Next K
    
End Sub
