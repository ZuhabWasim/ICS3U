VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Parameters 3"
   ClientHeight    =   7605
   ClientLeft      =   1890
   ClientTop       =   2205
   ClientWidth     =   6585
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
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MAX = 5
Const FORMNAME = "frmMain"

Private Sub mnuQ1_Click()
    
    Dim St(1 To MAX) As String
    Dim K As Integer
    
    For K = 1 To MAX
        St(K) = Str$(K)
    Next K
    
    Initialize St(), MAX
    
    For K = 1 To MAX
        Print St(K)
    Next K
    
End Sub

Private Sub mnuQ2_Click()
    
    Dim Ints(1 To MAX) As Integer
    Dim Num As Integer
    Dim K As Integer
    
    Num = InputBox$("Enter num.")
    
    For K = 1 To MAX
        Ints(K) = K
    Next K
    
    SetNums Ints(), MAX, Num
    
    For K = 1 To MAX
        Print Ints(K)
    Next K
    
End Sub

Private Sub mnuQ3_Click()
    
    Dim Names(1 To MAX) As String
    Dim K As Integer
    Dim Longest As String
    
    For K = 1 To MAX
        If K <> 1 Then
            Names(K) = Names(K - 1) & Str$(K)
        Else
            Names(K) = Str$(K)
        End If
    Next K
    
    GetLongest Names(), MAX, Longest
    
    Print Longest
    
End Sub

Private Sub mnuQ4_Click()
    
    Dim DecNums(1 To MAX) As Single
    Dim Sum As Single
    Dim K As Integer
    
    For K = 1 To MAX
        DecNums(K) = K / 10
    Next K
    
    SumOf DecNums(), MAX, Sum
    
    Print Sum
    
End Sub

Private Sub mnuQ5_Click()
    
    Dim PersonNames(1 To MAX) As String
    Dim PersonAge(1 To MAX) As Integer
    
    Dim K As Integer
    
    For K = 1 To MAX
            PersonNames(K) = Chr$(K + 64)
            PersonAge(K) = K
    Next K
        
    DisplayPersons PersonNames(), PersonAge(), MAX
            
End Sub
