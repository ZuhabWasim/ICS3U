VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Array Practice"
   ClientHeight    =   7605
   ClientLeft      =   1890
   ClientTop       =   1905
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   7365
   Begin VB.CommandButton cmdAlt 
      Caption         =   "Alternate"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2160
      TabIndex        =   3
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdReverse 
      Caption         =   "Reverse"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   2
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Top             =   3600
      Width           =   1335
   End
   Begin VB.PictureBox picData 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   360
      ScaleHeight     =   2115
      ScaleWidth      =   5235
      TabIndex        =   0
      Top             =   360
      Width           =   5295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Const NAMEMAX = 100
    
Dim FName(1 To NAMEMAX) As String
Dim NCount As Integer

Private Sub cmdAlt_Click()
    
    Dim K As Integer
    
    For K = 1 To NAMEMAX
        FName(K) = InputBox$("Enter a name:")
        picData.Print FName(K)
        If FName(K) = "" Then
            K = NAMEMAX
        End If
    Next K
    
End Sub

Private Sub cmdRead_Click()

    NCount = 0
    Open App.Path & "/Names.txt" For Input As #1
    
    picData.Cls
    
    Do While Not EOF(1)
        NCount = NCount + 1
        Input #1, FName(NCount)
        picData.Print FName(NCount)
    Loop
    
    Close #1
    
End Sub

Private Sub cmdReverse_Click()
    
    Dim K As Integer
    
    picData.Cls
    
    For K = NCount To 1 Step -1
        picData.Print FName(K)
    Next K
    
End Sub


