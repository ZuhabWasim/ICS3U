VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "String Functions"
   ClientHeight    =   2655
   ClientLeft      =   1890
   ClientTop       =   1905
   ClientWidth     =   6150
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   6150
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   495
      Left            =   3720
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   4920
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdAnalyze 
      Caption         =   "&Analyze"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame frmTransPhrase 
      Caption         =   "Analysis Results:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   5895
      Begin VB.Label lblTransPhrase 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   5415
      End
      Begin VB.Label Label2 
         Caption         =   "Transmitted Phrase:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.TextBox txtPhrase 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   5895
   End
   Begin VB.Label Label1 
      Caption         =   "Please enter a phrase:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'By: Zuhab Wasim
'Date: 07/10/15
'Purpose: To censor four letter words when filtering sentences inputed in the program.
Option Explicit

Private Sub cmdAnalyze_Click()
    
    'Declares the variables used in the procedures
    Dim Phrase As String
    Dim TransPhrase  As String
    Dim K As Integer
    Dim Word As String
    Dim Counter As Integer
    Dim Char As String
    
    'Allows input from the user for the sentence they would like to analyze
    Phrase = txtPhrase.Text & " "
    'Assigns the initial values to variables that will not get assignment later on in the procedure
    Counter = 0
    TransPhrase = ""
    Word = ""
    
    For K = 1 To Len(Phrase)
        'Assigns the first character from Phrase to the variable Char for analysis
        Char = Mid$(Phrase, K, 1)
        If Char <> " " And Char <> "." Then
            'If Char is not a space or a period, then proceed to building the word.
            Word = Word & Char
        Else
            'If Char is a space, or a period, this determines the length of the word if it is not a space or period
            Counter = Counter + 1
            If Len(Word) = 4 Then
            'If the word is four characters long, than replace the word with asterisks (censoring it)
                Word = "****"
            End If
            'Since Counter will be 1 only once throughout the entire loop, this If/Else statement will only be true once.
            If Counter = 1 Then
                'This code is for correctly inputing the first word of Phrase in TransPhrase
                If Char = "." Then
                    TransPhrase = Word & Char
                Else
                    TransPhrase = Word
                End If
            Else
                'Determines if Char is the last word of Phrase or not
                If Char = "." Then
                    TransPhrase = TransPhrase & " " & Word & Char
                Else
                    TransPhrase = TransPhrase & " " & Word
                End If
            End If
            'Resets the variable Word to build and analyze the next word in Phrase
            Word = ""
        End If
    Next K
    
    'Outputs the string value in Transphrase
    lblTransPhrase.Caption = TransPhrase
    
End Sub

Private Sub cmdClear_Click()
    
    txtPhrase.Text = ""
    lblTransPhrase.Caption = ""
    
End Sub

Private Sub cmdExit_Click()
    
    'Determines to see if the user wants to confirm their click for exit
    If MsgBox("Are you sure you want to exit?", vbYesNo, "Exit") = vbYes Then
        End
    End If
    
End Sub
