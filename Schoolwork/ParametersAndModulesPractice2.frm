VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Practice"
   ClientHeight    =   7605
   ClientLeft      =   1890
   ClientTop       =   2205
   ClientWidth     =   6585
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
   ScaleHeight     =   7605
   ScaleWidth      =   6585
   Begin VB.Menu mnuQ1 
      Caption         =   "Q1"
   End
   Begin VB.Menu mnuQ2 
      Caption         =   "Q2"
   End
   Begin VB.Menu mnuQ3 
      Caption         =   "Q3"
   End
   Begin VB.Menu mnuQ4 
      Caption         =   "Q4"
   End
   Begin VB.Menu mnuQ5 
      Caption         =   "Q5"
   End
   Begin VB.Menu mnuQ6 
      Caption         =   "Q6"
   End
   Begin VB.Menu mnuFactorial 
      Caption         =   "Factorial"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuFactorial_Click()
    
    Dim X As Integer
    
    X = InputBox$("")
    
    Print Str$(Factorial(X))
    
End Sub

Private Sub mnuQ1_Click()
    
    Dim RealNum As Single
    Dim DecNum As Integer
    
    RealNum = InputBox$("Enter real num.")
    DecNum = InputBox$("Enter dec num.")
    
    If IsNumeric(RealNum) And IsNumeric(DecNum) Then
        Round RealNum, DecNum
    Else
        MsgBox "Value not valid", vbOKOnly, "Error"
    End If
    
End Sub

Private Sub mnuQ2_Click()
    
    Dim Char As String
    
    Char = InputBox$("Please enter a character")
    
    CheckChar Char
    
End Sub

Private Sub mnuQ3_Click()
    
    Dim Num As Integer
    
    Num = InputBox$("Enter Num.")
    
    If IsNumeric(Num) Then
        SumOf (Num)
    Else
        MsgBox "hello?"
    End If
    
End Sub

Private Sub mnuQ4_Click()
    
    Dim Str As String
    
    Str = InputBox$("Enter str")
    
    ChangeStr Str
    
End Sub

Private Sub mnuQ5_Click()
    
    Dim UserDate As String
    
    UserDate = InputBox$("Enter date")
    
    ChangeDate UserDate
    
End Sub
