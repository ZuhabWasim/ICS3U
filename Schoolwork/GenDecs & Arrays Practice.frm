VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "GenDecs & Arrays"
   ClientHeight    =   5865
   ClientLeft      =   1890
   ClientTop       =   1905
   ClientWidth     =   7515
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
   ScaleHeight     =   5865
   ScaleWidth      =   7515
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   3
      Top             =   5040
      Width           =   2055
   End
   Begin VB.PictureBox picData 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   120
      ScaleHeight     =   4515
      ScaleWidth      =   6315
      TabIndex        =   2
      Top             =   360
      Width           =   6375
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   120
      TabIndex        =   1
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   2280
      TabIndex        =   0
      Top             =   5040
      Width           =   2040
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Programmer: Zuhab Wasim
'Date: 26/11/15
'Purpose: To read and display the contents of a data file of students, homeforms and marks
'Sequential.vbp. Sequential.frm, Sequential.frx

Const MAXSTUDENTS = 50

Dim Student(1 To MAXSTUDENTS) As String
Dim HomeForm(1 To MAXSTUDENTS) As String
Dim Mark(1 To MAXSTUDENTS) As Integer
Dim NumStudents As Integer

Private Sub cmdExit_Click()
    
    If MsgBox("Are you sure you want to exit", vbYesNo, "Exit") = vbYes Then
        End
    End If
    
End Sub

Private Sub Form_Load()
    
    Dim K As Integer
    
    For K = 1 To MAXSTUDENTS
        Student(K) = ""
        HomeForm(K) = ""
        Mark(K) = 0
    Next K
    
    NumStudents = 0
    
End Sub

Private Sub cmdOpen_Click()
    
    Dim K As Integer
    
    K = 0
    
    Open App.Path & "\marks.txt" For Input As #1
        
    Do While Not EOF(1)
        K = K + 1
        Input #1, Student(K), HomeForm(K), Mark(K)
    Loop
    
    Close #1
    
    NumStudents = K
    
    MsgBox "File opened and data received.", vbOKOnly, "File Opened"
    
End Sub

Private Sub cmdDisplay_Click()
    
    Dim K As Integer
    Dim X As Single
    Dim L As Single
    
    For K = 1 To NumStudents
        picData.Print Spc(1); Student(K); Tab(20); HomeForm(K); Tab(25); Mark(K)
    Next K
    
    For K = picData.Top To 360 Step 0.01

        
    Next K
    
    
End Sub
