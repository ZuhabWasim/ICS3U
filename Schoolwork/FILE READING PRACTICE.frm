VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5925
   ClientLeft      =   1890
   ClientTop       =   1905
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   4740
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
      Left            =   3240
      TabIndex        =   2
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdReadInfo 
      Caption         =   "&Read Information"
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
      Left            =   360
      TabIndex        =   1
      Top             =   5280
      Width           =   2775
   End
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
      Height          =   4935
      Left            =   240
      ScaleHeight     =   4875
      ScaleWidth      =   4155
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    
    If MsgBox("Are you sure you want to Exit?", vbYesNo, "Exit?") = vbYes Then
        End
    End If
    
End Sub

Private Sub cmdReadInfo_Click()
    
    Dim X As Integer
    Dim StudentName As String
    Dim Mark As Integer
    Dim HF As String
    Dim RMark As String * 4
    Dim RHF As String * 4
    
    X = 0
    Open "H:\ICS\ICS3U1\MARKS.txt" For Input As #1
    
    picData.Print "Student Name"; Tab(27); "HF"; Tab(32); " Mark %"
    picData.Print
    
    Do While Not EOF(1)
        X = X + 1
        Input #1, StudentName, HF, Mark
        RSet RHF = HF
        RSet RMark = Format$(Mark, "####")
        picData.Print StudentName; Tab(25); RHF; Tab(35); RMark
    Loop
    
    Close #1
    
    picData.Print
    picData.Print "Total number of students: "; X
    
End Sub


