VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7605
   ClientLeft      =   1890
   ClientTop       =   1905
   ClientWidth     =   6585
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   6585
   Begin VB.CommandButton Command1 
      Caption         =   "J3 - Rovarspraket"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   960
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    
    Dim Word As String
    Dim Char As String
    Dim Letter1 As String
    Dim Letter2 As String
    Dim Letter3 As String
    Dim AmountDown As Integer
    Dim AmountUp As Integer
    Dim K As Integer
    Dim X As Integer
    
    AmountDown = 0
    AmountUp = 0
    Word = InputBox$("Enter a word to translate")
    
    For K = 1 To Len(Word)
        Char = Mid$(Word, K)
        If Not (Char = "a" Or Char = "e" Or Char = "i" Or Char = "o" Or Char = "u") Then
            Letter1 = Char
            If Letter >
    
End Sub
