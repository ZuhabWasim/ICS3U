VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Text Files Questions"
   ClientHeight    =   7605
   ClientLeft      =   1890
   ClientTop       =   1905
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   6585
   Begin VB.PictureBox picData 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      ScaleHeight     =   4035
      ScaleWidth      =   5355
      TabIndex        =   5
      Top             =   120
      Width           =   5415
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   4
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   3
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton cmdQ3 
      Caption         =   "Q&3"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton cmdQ2 
      Caption         =   "Q&2"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   1
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdQ1 
      Caption         =   "Q&1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
    
    picData.Cls
    
End Sub

Private Sub cmdExit_Click()

    If MsgBox("Are you sure you want to exit", vbYesNo, "Exit") = vbYes Then
        End
    End If
    
End Sub

Private Sub cmdQ1_Click()
    
    Dim PriceCount As Integer
    Dim PriceSum As Single
    Dim PriceAverage As Single
    Dim PriceName As String
    Dim PriceAmount As Single
    
    picData.Cls
    
    PriceCount = 0
    
    Open "H:\ICS\ICS3U1\prices.txt" For Input As #1
    
    Do While Not EOF(1)
        Input #1, PriceName, PriceAmount
        picData.Print PriceName; Tab(25); Format$(Format$(PriceAmount, "currency"), "@@@@@@")
        If PriceAmount > 5 Then
            PriceCount = PriceCount + 1
            PriceSum = PriceSum + PriceAmount
        End If
    Loop
    
    Close #1
    
    If PriceCount <> 0 Then
        PriceAverage = PriceSum / PriceCount
    Else
        PriceAverage = 0
    End If
    
    picData.Print
    picData.Print "The average price of all the items above $5.00 is:"
    picData.Print Format$(PriceAverage, "currency")
    
End Sub

Private Sub cmdQ2_Click()
    
    Dim HF As String
    Dim HFBCount As Integer
    Dim HFLetter As String
    
    picData.Cls
    
    HFBCount = 0
    
    Open "H:\ICS\ICS3U1\hf.txt" For Input As #1
    
    Do While Not EOF(1)
        Input #1, HF
        HFLetter = UCase$(Right$(HF, 1))
        If HFLetter = "B" Then
            HFBCount = HFBCount + 1
        End If
        picData.Print HF
    Loop
    
    Close #1
    
    picData.Print
    picData.Print "The number of Home Forms that have a B is: "; HFBCount

End Sub

Private Sub cmdQ3_Click()
    
    Dim Word As String
    Dim FLetter As String
    Dim NString As String
    Dim Count As Integer
    Dim HalfWord As String
    Dim Char As String
    Dim K As Integer
    
    picData.Cls
    picData.Print "Word:"; Tab(14); "New Word:"
    Count = 0
    NString = ""
    HalfWord = ""
    
    Open App.Path & "\words.txt" For Input As #1
    
    Do While Not EOF(1)
        Input #1, Word
        FLetter = Left$(Word, 1)
        HalfWord = LCase$(Right$(Word, Len(Word) - 1))
        If Len(Word) > 4 Then
            FLetter = UCase$(FLetter)
            NString = FLetter & HalfWord
        End If
        picData.Print Word; Tab(14); NString
        HalfWord = ""
        NString = ""
    Loop
    
    Close #1
    
    
End Sub
