VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7605
   ClientLeft      =   4515
   ClientTop       =   1110
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   6585
   Begin VB.CommandButton cmdQ9 
      Caption         =   "Q9 - Rootz"
      Height          =   495
      Left            =   3600
      TabIndex        =   23
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdQ7 
      Caption         =   "Q7 - Wolverine"
      Height          =   615
      Left            =   1800
      TabIndex        =   22
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton cmdQ5 
      Caption         =   "Q5 - Contribution"
      Height          =   495
      Left            =   2400
      TabIndex        =   17
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton cmdQ6 
      Caption         =   "Q6 - Greatest"
      Height          =   615
      Left            =   240
      TabIndex        =   16
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdQn4 
      Caption         =   "Q4 - Carpet"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton cmdAverage 
      Caption         =   "Q3 - Average"
      Height          =   495
      Left            =   1680
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdAge 
      Caption         =   "Q2 - Age"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdTime 
      Caption         =   "Q1 - Hours"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtSeconds 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox txtMinutes 
      Height          =   405
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox txtHours 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin VB.Label lblCredit 
      Height          =   255
      Left            =   3480
      TabIndex        =   21
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label lblContribution 
      Height          =   255
      Left            =   3480
      TabIndex        =   20
      Top             =   2040
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   2280
      X2              =   2280
      Y1              =   1320
      Y2              =   3480
   End
   Begin VB.Label label6 
      Caption         =   "Federal Political Tax Credit"
      Height          =   855
      Left            =   2400
      TabIndex        =   19
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label label5 
      Caption         =   "Contribution Amount:"
      Height          =   495
      Left            =   2400
      TabIndex        =   18
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lblTotal 
      Height          =   255
      Left            =   1200
      TabIndex        =   15
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label lblHST 
      Height          =   255
      Left            =   1200
      TabIndex        =   14
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label lblDelivery 
      Height          =   255
      Left            =   1200
      TabIndex        =   13
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label lblSubtotal 
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Total:"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "HST:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Delivery:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Subtotal:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label lblTotalSeconds 
      Height          =   255
      Left            =   4320
      TabIndex        =   4
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAge_Click()
    
    Dim Age As Integer
    Dim Message As String
    
    Age = Val(InputBox$("Please enter your age.", "Enter your age"))
    
        If Age < 16 Then
            Message = "You can't drive or drink!"
        ElseIf Age < 19 Then
            Message = "You are old enough to drive but not drink!"
        Else
            Message = "You are old enough to drive or drink!"
        End If
    
    MsgBox Message
    
End Sub

Private Sub cmdAverage_Click()
    
    Dim K As Integer
    Dim NAmount As Integer
    Dim Num As Single
    Dim Average As Single
    Dim Count As Integer
    
    NAmount = Val(InputBox$("Please enter the amount of number you will put in", "Enter amount"))
    Num = 0
    Count = 0
    
    If NAmount < 1 Then
        MsgBox "Please enter a number that is greater than 0"
    Else
        For K = NAmount To 1 Step -1
            Num = Val(InputBox$("Enter a number.", "Number"))
            Count = Count + Num
        Next K
        
        Average = Count / NAmount
        
        MsgBox "The average of the" & Str$(NAmount) & " numbers you inputed was" & Format$(Average, "#,###")
    End If
    
End Sub

Private Sub cmdQ6_Click()
    
    Dim Number1 As Single
    Dim Number2 As Single
    Dim Number3 As Single
    Dim Greatest As Single
    Dim Least As Single
    Dim Middle As Single
    
    Number1 = Val(InputBox$("Please enter a number.", "Number 1"))
    Number2 = Val(InputBox$("Please enter a number.", "Number 2"))
    Number3 = Val(InputBox$("Please enter a number.", "Number 3"))
    
   
    If Number1 > Number2 Then
        If Number1 > Number3 Then
            If Number2 > Number3 Then
                Greatest = Number1
                Middle = Number2
                Least = Number3
            Else
                Greatest = Number1
                Middle = Number3
                Least = Number2
            End If
        Else
            Greatest = Number3
            Middle = Number1
            Least = Number2
        End If
    Else
      
            
                
            
    End If
    
End Sub

Private Sub cmdQ7_Click()
    
    Dim WolvNum As Long
    Dim K As Long
    Dim Counter As Integer
    
    WolvNum = Val(InputBox$("How many times do you want WOLVERINE to be displayed?", "WOLVERINE"))
    Counter = 0
    
    If WolvNum < 0 Then
        MsgBox "Invalid Amount"
    Else
        For K = 1 To WolvNum
            If Counter = 20 Then
                Print "WOLVERINE"
                Counter = 0
            Else
                Print "WOLVERINE",
                Counter = Counter + 1
            End If
        Next K
    End If
    
End Sub

Private Sub cmdQ9_Click()
    
    Dim Num As Integer
    Dim Root As Single
    Dim K As Integer
    Dim Counter As Integer
    
    Num = Val(InputBox$("Please enter a postive number", "Enter a number"))
    
    For K = 1 To Num
        If Sqr(Counter) Mod 1 = 0 And Sqr(Counter) < Num Then
            Print Sqr(Counter)
        Else
End Sub

Private Sub cmdQn4_Click()
    
    Const Tax = 0.13
    
    Dim Delivery As Integer
    Dim Price As Single
    Dim SubTotal As Single
    Dim Total As Single
    Dim Carpet As Single
    Dim HST As Single
    
    Carpet = Val(InputBox$("How much carpet do you want to order? (In squared meters)", "Carpet"))
    
    If Carpet <= 0 Then
        MsgBox "Please enter an amount greater than 0"
    Else
        If Carpet < 8 Then
            Price = 25
        ElseIf Carpet <= 24 Then
            Price = 21
        Else
            Price = 18
        End If
        
        SubTotal = Carpet * Price
        HST = SubTotal * Tax
        
        If MsgBox("Do you want it delivered at the low cost of 75$?", vbYesNo, "Deliver?") = vbYes Then
            Delivery = 75
        Else
            Delivery = 0
        End If
        
        Total = HST + SubTotal + Delivery
        
        lblSubtotal.Caption = Format$(SubTotal, "$0.00")
        lblHST.Caption = Format$(HST, "$0.00")
        lblDelivery.Caption = Format$(Delivery, "$0.00")
        lblTotal.Caption = Format$(Total, "$0.00")
    
    End If
    
End Sub

Private Sub cmdTime_Click()

    Dim Hours As Long
    Dim Minutes As Integer
    Dim Seconds As Integer
    Dim TotalSeconds As Long
    
    Hours = Val(txtHours.Text)
    Minutes = Val(txtMinutes.Text)
    Seconds = Val(txtSeconds.Text)
    
    If Hours < 0 Then
        Hours = 0
    ElseIf Minutes < 0 Then
        Minutes = 0
    ElseIf Seconds < 0 Then
        Seconds = 0
    End If
    
    Hours = Hours * 3600
    Minutes = Minutes * 60
    
    TotalSeconds = Hours + Minutes + Seconds
    
    If TotalSeconds > 86399 Then
        MsgBox "Error: Please enter time that is less than 24 hours", vbOKOnly, "ERROR"
    ElseIf TotalSeconds < 0 Then
        MsgBox "Error: Time is less than 0", vbOKOnly, "ERROR"
    Else
        lblTotalSeconds.Caption = Str$(TotalSeconds)
    End If
    
End Sub

Private Sub cmdQ5_Click()
    
    Dim ContributionAmount As Single
    Dim Credit As Single
    
    ContributionAmount = Val(InputBox$("Please enter the amount of contribution:", "Contribution Amount", "25"))
    
        If ContributionAmount < 1 Then
            Credit = 0
            MsgBox "Invalid Amount.", vbOKOnly, "Invalid Amount"
        Else
            If ContributionAmount <= 100 Then
                Credit = 0.75 * ContributionAmount
            ElseIf ContributionAmount <= 550 Then
                Credit = 75 + 0.5 * (ContributionAmount - 100)
            ElseIf ContributionAmount <= 750 Then
                Credit = 300 + (0.33 * (ContributionAmount - 550))
            Else
                MsgBox "The maximum you can donate is $750.00", vbOKOnly, "Error2: Donation Over $750"
            End If
        End If
        
    lblContribution.Caption = Format$(ContributionAmount, "currency")
    lblCredit.Caption = Format$(Credit, "currency")
    
End Sub
