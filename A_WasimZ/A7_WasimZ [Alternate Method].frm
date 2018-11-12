VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Tic-Tac-Toe Game"
   ClientHeight    =   7605
   ClientLeft      =   1890
   ClientTop       =   1905
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   6585
   Begin VB.Image imgSquare 
      Height          =   1410
      Index           =   8
      Left            =   2940
      Top             =   2940
      Width           =   1410
   End
   Begin VB.Image imgSquare 
      Height          =   1410
      Index           =   7
      Left            =   1530
      Top             =   2940
      Width           =   1410
   End
   Begin VB.Image imgSquare 
      Height          =   1410
      Index           =   6
      Left            =   120
      Top             =   2940
      Width           =   1410
   End
   Begin VB.Image imgSquare 
      Height          =   1410
      Index           =   5
      Left            =   2940
      Top             =   1530
      Width           =   1410
   End
   Begin VB.Image imgSquare 
      Height          =   1410
      Index           =   4
      Left            =   1530
      Top             =   1530
      Width           =   1410
   End
   Begin VB.Image imgSquare 
      Height          =   1410
      Index           =   3
      Left            =   120
      Top             =   1530
      Width           =   1410
   End
   Begin VB.Image imgSquare 
      Height          =   1410
      Index           =   2
      Left            =   2940
      Top             =   120
      Width           =   1410
   End
   Begin VB.Image imgSquare 
      Height          =   1410
      Index           =   1
      Left            =   1530
      Top             =   120
      Width           =   1410
   End
   Begin VB.Image imgSquare 
      Height          =   1410
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   1410
   End
   Begin VB.Image Image1 
      Height          =   1410
      Left            =   3600
      Picture         =   "A7_WasimZ [Alternate Method].frx":0000
      Top             =   5280
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Image imgO 
      Height          =   1410
      Left            =   1920
      Picture         =   "A7_WasimZ [Alternate Method].frx":075B
      Top             =   5520
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Image imgX 
      Height          =   1410
      Left            =   360
      Picture         =   "A7_WasimZ [Alternate Method].frx":19D0
      Top             =   5520
      Visible         =   0   'False
      Width           =   1410
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub imgSquare_Click(Index As Integer)

End Sub

Function CheckWin(Index As Integer, SymType As String) As Boolean
    
    Dim K As Integer
    Dim CheckSymbol As String
    Dim Row As String
    
    If SymType = "X" Then
        CheckSymbol = "X"
    Else
        CheckSymbol = "O"
    End If
    
    
        
    
End Function
