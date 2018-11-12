VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Practice"
   ClientHeight    =   2550
   ClientLeft      =   1890
   ClientTop       =   1905
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   ScaleHeight     =   2550
   ScaleWidth      =   7095
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
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
      Left            =   5760
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display"
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
      Left            =   5760
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
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
      Left            =   5760
      TabIndex        =   2
      Top             =   720
      Width           =   1215
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
      Height          =   495
      Left            =   5760
      TabIndex        =   1
      Top             =   120
      Width           =   1215
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
      Height          =   2295
      Left            =   120
      ScaleHeight     =   2235
      ScaleWidth      =   5475
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MAX = 100
Const INPUTFILE = "\PRICESA.txt"

Dim ItemAmount As Integer
Dim Item(1 To MAX) As String
Dim Quantity(1 To MAX) As Integer
Dim UPrice(1 To MAX) As Single
Dim ItemCost(1 To MAX) As Single

Private Sub cmdCalculate_Click()
    
    Dim K As Integer
    
    For K = 1 To MAX
        ItemCost(K) = Quantity(K) * UPrice(K)
    Next K
    
End Sub

Private Sub cmdDisplay_Click()
    
    Dim K As Integer
    Dim CostSum As Single
    
    picData.Cls
    
    For K = 1 To MAX
        If ItemCost(K) > 66.5 Then
            picData.Print K; ". "; _
            Tab(8); Item(K); _
            Tab(24); Format$(Quantity(K), "@@@@@"); _
            Tab(34); Format$(Format$(UPrice(K), "$ #,##0.00"), "@@@@@@@")
        End If
        CostSum = CostSum + ItemCost(K)
    Next K
    
    picData.Print
    picData.Print "The total cost of all items are: "; CostSum
    
    
End Sub

Private Sub cmdRead_Click()
    
    Dim K As Integer
    
    K = 0
    Open App.Path & INPUTFILE For Input As #1
    
    Do While Not EOF(1)
        K = K + 1
        If K > 100 Then
            MsgBox "You have reached the max number of items the program can read.", vbOKOnly, "Error: 1"
        Else
            Input #1, Item(K), Quantity(K), UPrice(K)
        End If
    Loop
    
    ItemAmount = K
    
    If MsgBox("The file has been read, would you like to calculate the cost of each item and display them?", vbYesNo, "File Read") = vbYes Then
        cmdCalculate_Click
        cmdDisplay_Click
    End If
        
End Sub

Private Sub Form_Load()
    
    Dim K As Integer
    
    ItemAmount = 0
    For K = 1 To MAX
        Item(K) = 0
        Quantity(K) = 0
        UPrice(K) = 0
    Next K
    
End Sub
