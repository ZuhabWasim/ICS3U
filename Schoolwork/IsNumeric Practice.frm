VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   6750
   ClientLeft      =   1890
   ClientTop       =   1905
   ClientWidth     =   5235
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
   ScaleHeight     =   6750
   ScaleWidth      =   5235
   Begin VB.Frame Frame4 
      Caption         =   "Question 3"
      Height          =   1575
      Left            =   120
      TabIndex        =   12
      Top             =   5040
      Width           =   4935
      Begin VB.PictureBox picValid 
         Height          =   975
         Left            =   1440
         ScaleHeight     =   915
         ScaleWidth      =   3195
         TabIndex        =   14
         Top             =   360
         Width           =   3255
      End
      Begin VB.CommandButton cmdSNumber 
         Caption         =   "Student Number"
         Height          =   615
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Question 2"
      Height          =   2535
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   4935
      Begin VB.PictureBox picSequence 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   1320
         ScaleHeight     =   1875
         ScaleWidth      =   3315
         TabIndex        =   11
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox txtSequence 
         Height          =   360
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdSequence 
         Caption         =   "Do Sequence"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Question 1"
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   4935
      Begin VB.TextBox txtMark 
         Height          =   480
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
      Begin VB.PictureBox picMark 
         Height          =   495
         Left            =   2400
         ScaleHeight     =   435
         ScaleWidth      =   2235
         TabIndex        =   6
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton cmdMark 
         Caption         =   "Mark"
         Height          =   495
         Left            =   1080
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Trial Program"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.TextBox txtNumber 
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdNumber 
         Caption         =   "Number ?"
         Height          =   495
         Left            =   1440
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblNumber 
         Height          =   495
         Left            =   3000
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdMark_Click()
    
    Dim Mark As Single
    
    Mark = txtMark.Text
    
    picMark.Cls
    
    If IsNumeric(Mark) Then
        If Mark >= 50 And Mark <= 100 Then
            picMark.Print Mark
            picMark.Print "You have passed."
        ElseIf Mark >= 0 And Mark < 50 Then
            picMark.Print Mark
            picMark.Print "You have failed."
        Else
            MsgBox "Mark number entered is not between 0 - 100", vbOKOnly, "Error"
        End If
    Else
        MsgBox "Please enter a number", vbOKOnly, "Error: Value not recognized"
    End If
    
End Sub

Private Sub cmdNumber_Click()
    
    If IsNumeric(txtNumber.Text) Then
        lblNumber.Caption = " It's a number! "
    Else
        lblNumber.Caption = " It's NOT a number! "
    End If
    
End Sub

Private Sub cmdSequence_Click()
    
    Dim X As Integer
    Dim Number As Single
    
    picSequence.Cls

    If IsNumeric(txtSequence.Text) Then
        Number = txtSequence.Text
        If (Number >= 0 And Number < 10) And Number Mod 1 = 0 Then
            Do While Number <> 0
                For K = 1 To Number
                    picSequence.Print K;
                Next K
                picSequence.Print
                Number = Number - 1
            Loop
        Else
            MsgBox "Number is not a single-digit number or number is not an integer", vbOKOnly, "Error: Number not valid"
        End If
    Else
        MsgBox "Value given is not a number.", vbOKOnly, "Error: Number not valid"
    End If
    
            
End Sub

Private Sub cmdSNumber_Click()
    
    Dim SNumber As String
    
    SNumber = InputBox$("Please enter a student number:", "Student Number")
    
    picValid.Cls
    
    If IsNumeric(SNumber) Then
        If Len(SNumber) = 9 Then
            picValid.Print "Valid"
        Else
            picValid.Print "Invalid"
        End If
    Else
        MsgBox "Please enter a number"
    End If
        
End Sub
