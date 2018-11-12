VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Right Justification of Numbers (RSet)"
   ClientHeight    =   2790
   ClientLeft      =   1890
   ClientTop       =   1905
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   ScaleHeight     =   2790
   ScaleWidth      =   8160
   Begin VB.PictureBox picData 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   3960
      ScaleHeight     =   2475
      ScaleWidth      =   4035
      TabIndex        =   7
      Top             =   120
      Width           =   4095
   End
   Begin VB.CommandButton cmdAddInfo 
      Caption         =   "&Add Information"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   6
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox txtHF 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtAge 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1680
      TabIndex        =   3
      Top             =   165
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Home Form:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Age:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   615
   End
   Begin VB.Label label1 
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
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
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddInfo_Click()

    Dim Name As String
    Dim Age As Integer
    Dim AgeStr As String * 3
    Dim HF As String
    Dim HFStr As String * 3
    Dim LastName As String
    Dim FirstName As String
    Dim K As Integer
    Dim Char As String
    Dim Count As Integer
    
    Name = txtName.Text
    Age = Val(txtAge.Text)
    HF = txtHF.Text
    FirstName = ""
    LastName = ""
    Count = 1
    
    For K = Len(Name) To 1 Step -1
        Char = Mid$(Name, K, 1)
        If Count = 1 Then
            If Not (Char = " ") Then
                LastName = Char & LastName
            Else
                Count = 2
            End If
        Else
            FirstName = Char & FirstName
        End If
    Next K
    
    RSet AgeStr = Format$(Age, "###")
    RSet HFStr = HF
    
    picData.Print FirstName; Tab(15); LastName; Tab(30); AgeStr; Tab(36); HFStr
    
End Sub

Private Sub Form_Load()

    picData.Print "FirstName"; Tab(15); "LastName"; Tab(30); "Age"; Tab(36); " HF"
    
End Sub
