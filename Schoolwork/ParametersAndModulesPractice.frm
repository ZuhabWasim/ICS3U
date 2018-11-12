VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   7605
   ClientLeft      =   1575
   ClientTop       =   2550
   ClientWidth     =   9615
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
   ScaleWidth      =   9615
   Begin VB.Label lblMessage 
      Height          =   375
      Left            =   7800
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Menu mnuPractice 
      Caption         =   "Practice"
      Begin VB.Menu mnuWelcome 
         Caption         =   "Welcome"
      End
      Begin VB.Menu mnuPattern 
         Caption         =   "Pattern"
      End
   End
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

Private Sub mnuPattern_Click()

    Dim Char As String
    Dim Num As Integer
    
    Char = InputBox$("Enter Char")
    Num = Val(InputBox$("Enter Num"))
    
    PrintPattern Char, Num
    
End Sub

Private Sub mnuQ1_Click()
    
    Dim UserStr As String
    
    UserStr = InputBox$("Enter a string")
    
    Print ReverseStr(UserStr)
    
End Sub

Private Sub mnuQ2_Click()
    
    Dim UserChar As String
    Dim UserNum As Integer
    
    UserNum = InputBox$("Enter Num")
    UserChar = InputBox$("")
    
    If IsNumeric(UserNum) Then
        Print MultString(UserChar, UserNum)
    End If
    
End Sub

Private Sub mnuWelcome_Click()

    Welcome
    
End Sub
