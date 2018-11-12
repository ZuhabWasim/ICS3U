VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6135
   ClientLeft      =   1470
   ClientTop       =   1800
   ClientWidth     =   11640
   FillStyle       =   0  'Solid
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
   ScaleHeight     =   6135
   ScaleWidth      =   11640
   Begin VB.Frame Frame1 
      Caption         =   "6. Craps"
      Height          =   2295
      Left            =   9000
      TabIndex        =   7
      Top             =   3720
      Width           =   2535
      Begin VB.CommandButton cmdRoll 
         Caption         =   "Roll!"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtDie 
         Height          =   360
         Left            =   1440
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblWinLoss 
         Height          =   495
         Left            =   1320
         TabIndex        =   15
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblDie2 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   14
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lblDie1 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Die 2"
         Height          =   375
         Left            =   1560
         TabIndex        =   12
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Die 1"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Enter your number:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdQ5 
      Caption         =   "5. Mix Names"
      Height          =   495
      Left            =   9000
      TabIndex        =   6
      Top             =   3120
      Width           =   2535
   End
   Begin VB.CommandButton cmdQ4 
      Caption         =   "4. Make Words"
      Height          =   495
      Left            =   9000
      TabIndex        =   5
      Top             =   2520
      Width           =   2535
   End
   Begin VB.CommandButton cmdQ3 
      Caption         =   "3. Arrays"
      Height          =   495
      Left            =   9000
      TabIndex        =   4
      Top             =   1920
      Width           =   2535
   End
   Begin VB.CommandButton cmdQ2 
      Caption         =   "2. A - Z"
      Height          =   495
      Left            =   9000
      TabIndex        =   3
      Top             =   1320
      Width           =   2535
   End
   Begin VB.CommandButton cmdQ1 
      Caption         =   "1. 0 - 100"
      Height          =   495
      Left            =   9000
      TabIndex        =   2
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton cmdGetNum 
      Caption         =   "P. Get Num"
      Height          =   495
      Left            =   9000
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.PictureBox picData 
      Height          =   5895
      Left            =   120
      ScaleHeight     =   5835
      ScaleWidth      =   8715
      TabIndex        =   0
      Top             =   120
      Width           =   8775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Const MAX = 20
    
Dim Names(1 To MAX) As String
Dim RandNames(1 To MAX) As String
Dim NamesAmount As Integer

Private Sub cmdGetNum_Click()
        
    picData.Cls
    
    picData.Print Rnd()
    
End Sub

Private Sub cmdQ1_Click()
    
    Dim Num(1 To 50) As Single
    Dim K As Integer
    
    picData.Cls
    
    For K = 1 To MAX
        Num(K) = Int(Rnd() * (100 + 1))
        picData.Print Num(K)
    Next K
    
End Sub

Private Sub cmdQ2_Click()
    
    Dim Letter(1 To 25) As String
    Dim K As Integer
    
    picData.Cls
    
    For K = 1 To 25
        Letter(K) = Chr$(Int(Rnd() * (90 - 65 + 1)) + 65)
        picData.Print Letter(K)
    Next K
    
End Sub

Private Sub cmdQ3_Click()
    
    Const MAX = 1000
    
    Dim K As Integer
    Dim Count As Integer
    Dim Num(1 To 1000) As Integer
    
    picData.Cls
    
    Count = 0
    
    For K = 1 To MAX
        Num(K) = Int(Rnd() * (600 - 100 + 1)) + 100
        If Count <> 8 Then
            picData.Print Num(K);
            Count = Count + 1
        Else
            picData.Print Num(K)
            Count = 0
        End If
    Next K
    
        
End Sub

Private Sub cmdQ4_Click()
    
    Const MAX = 100
    
    Dim Word(1 To MAX) As String
    Dim Char As String
    Dim Count As Integer
    Dim K As Integer
    Dim X As Integer
    Dim Num As Integer
    
    picData.Cls
    
    Count = 0
    
    For K = 1 To MAX
        Word(K) = ""
        Num = Int(Rnd() * (13 - 4 + 1)) + 4
        For X = 1 To Num
            Char = Chr$(Int(Rnd() * (122 - 97 + 1)) + 97)
            Word(K) = Word(K) & Char
        Next X
        If Count <> 4 Then
            picData.Print Word(K),
            Count = Count + 1
        Else
            picData.Print Word(K)
            Count = 0
        End If
        Char = ""
        Num = 0
    Next K
        
End Sub


Private Sub cmdQ5_Click()
    
    Intialize
    
    Read
    
    Mix
    
    Display
    
End Sub
Private Sub Display()
     
    Dim Output As String
    Dim K As Integer
    
    picData.Cls
    
    For K = 1 To NamesAmount
        picData.Print RandNames(K)
    Next K
    Output = "There were " & Str$(NamesAmount) & " names given."
    
End Sub
Private Sub Intialize()
    
    Dim K As Integer
    
    For K = 1 To MAX
        RandNames(K) = ""
        Names(K) = ""
    Next K
    
End Sub
Private Sub Mix()

    Dim K As Integer
    Dim RandNum As Integer
    Dim RandNameLocation As Integer
    
    For K = 1 To NamesAmount
        RandNum = Int(Rnd() * (NamesAmount - 1 + 1)) + 1
        Do While RandNames(RandNum) <> ""
            RandNum = Int(Rnd() * (NamesAmount - 1 + 1)) + 1
        Loop
        RandNames(RandNum) = Names(K)
        RandNum = 0
    Next K
    
End Sub
    
Private Sub Read()
    
    Dim K As Integer
    Dim Count As Integer
    
    Count = 0
    
    Open App.Path & "\Randomize Practice.txt" For Input As #1
    
    Do While Not EOF(1)
        Count = Count + 1
        Input #1, Names(Count)
    Loop
    
    Close #1
    
    NamesAmount = Count

End Sub

Private Sub cmdRoll_Click()
    
    Dim Die1 As Integer
    Dim Die2 As Integer
    Dim Num As Integer
    Dim Outcome As String
    
    If IsNumeric(txtDie.Text) Then
        Num = txtDie.Text
        If (Num >= 2 And Num <= 12) And (Num Mod 1 = 0) Then
            Num = txtDie.Text
        Else
            MsgBox "Please enter a number between 2 and 12", vbOKOnly, "Error 2: Invlaid Number"
        End If
    Else
        MsgBox "Please enter a number.", vbOKOnly, "Error 1: Not a number"
    End If
    
    Die1 = Int(Rnd() * (6 - 1 + 1)) + 1
    Die2 = Int(Rnd() * (6 - 1 + 1)) + 1
    
    If Num = Die1 + Die2 Then
        lblWinLoss.Caption = "You Won!"
    Else
        lblWinLoss.Caption = "You Lost!"
    End If
    lblDie1.Caption = Str$(Die1)
    lblDie2.Caption = Str$(Die2)
    
End Sub

Private Sub Form_Load()
    
    Randomize
    
End Sub
