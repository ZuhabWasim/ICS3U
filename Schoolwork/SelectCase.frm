VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Select Case"
   ClientHeight    =   7605
   ClientLeft      =   1890
   ClientTop       =   2205
   ClientWidth     =   6585
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   6585
   Begin VB.TextBox txtData 
      Height          =   495
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   2295
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
End
Attribute VB_Name = "frmmAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuQ1_Click()
    
    Dim Char As String
    
    Char = VBA.UCase$(txtData.Text)
    
    Cls
    
    Select Case Char
        Case "A", "E", "I", "O", "U"
            Print "You have entered a digit"
        Case "A" To "Z"
            Print "You have entered a consanant"
        Case "0" To "9"
            Print "You have entered a digit"
        Case ""
            Print "You have entered a null string"
        Case Else
            Print "You have entered 'other'"
    End Select
    
End Sub

Private Sub mnuQ2_Click()
    
    Dim Mark As Integer
    
    If IsNumeric(txtData.Text) Then
        Mark = Val(Int(txtData.Text))
    Else
        Print "Error"
    End If
    
    Cls
    
    Select Case Mark
        Case 0 To 49
            Print "0-49"
        Case 50 To 59
            Print "50-59"
        Case 60 To 69
            Print "60-69"
        Case 70 To 79
            Print "70-79"
        Case 80 To 90
            Print "80-90"
        Case 90 To 100
            Print "90-100"
        Case Else
            Print "Invalid Number"
    End Select
        
End Sub

Private Sub mnuQ3_Click()
    
    Dim Sentence As String
    Dim Char As String
    Dim K As Integer
    Dim SpaceCount As Integer
    Dim CharCount As Integer
    Dim NumCount As Integer
    Dim NextSpace As Integer
    Dim WordCount As Integer
    
    Sentence = txtData.Text
    
    NumCount = 0
    SpaceCount = 0
    CharCount = 0
    WordCount = 0
    
    For K = 1 To Len(Sentence)
        Char = VBA.Mid$(Sentence, K, 1)
        Select Case VBA.UCase$(Char)
            Case " "
                SpaceCount = SpaceCount + 1
            Case "A" To "Z"
                CharCount = CharCount + 1
            Case "0" To "9"
                NumCount = NumCount + 1
            Case " "
                If VBA.Mid$(Sentence, K - 1, 1) <> " " Then
                    WordCount = WordCount + 1
                End If
        End Select
    Next K
    
    Print "Letters: "; CharCount
    Print "Digits:  "; NumCount
    Print "Spaces:  "; SpaceCount
    Print "Words:   "; WordCount
    
End Sub

