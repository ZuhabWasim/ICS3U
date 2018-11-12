VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7605
   ClientLeft      =   1890
   ClientTop       =   1905
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   6585
   Begin VB.TextBox txtNum 
      Height          =   495
      Left            =   4440
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub txtNum_KeyPress(KeyAscii As Integer)
    
    Dim K, X As Integer
    Dim Num As Integer
    Dim Remainder(1 To 100) As Integer
    Dim Quotient(1 To 100) As Integer
    
    Cls
    
    Num = Val(txtNum.Text & Chr$(KeyAscii))
    K = 0
    
    Do
        K = K + 1
        If K = 1 Then
            Quotient(K) = Int(Num / 2)
        Else
            Quotient(K) = Int(Quotient(K - 1) / 2)
        End If
        Remainder(K) = Quotient(K) Mod 2
    Loop Until Quotient(K) <= 0
    
    For X = 1 To K
        Print Remainder(X) & " ";
    Next X
    
End Sub
