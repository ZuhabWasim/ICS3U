VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7620
   ClientLeft      =   1890
   ClientTop       =   1890
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   6585
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton cmdQ6 
      Caption         =   "Q6"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton cmdQ5 
      Caption         =   "Q5"
      Height          =   435
      Left            =   4440
      TabIndex        =   1
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton cmdQ4 
      Caption         =   "Q4"
      Height          =   495
      Left            =   4440
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdQ4_Click()

    Dim Mark As String
    Dim Count As Integer
    Dim OtherCount As Integer
    
    Count = 0
    
    Mark = InputBox$("Enter a mark, Input a '-1' if you want to stop")
    
    Do While Mark <> -1
        If Mark >= 50 Then
            Count = Count + 1
        End If
        Mark = InputBox$("Enter a mark")
    Loop
        
    If Mark <> -1 Then
        Print Count
    End If
        
    
    
End Sub


Private Sub cmdQ5_Click()
    
    Dim Height As Single
    Dim Count As Integer
    Dim HeightAverage As Single
    Dim HeightSum As Single
    
    Cls
    
    Count = 0
    HeightSum = 0
    Height = Val(InputBox$("Enter a height"))
    
    Do While Height <> 0
        Count = Count + 1
        HeightSum = HeightSum + Height
        Height = Val(InputBox$("Enter a height"))
    Loop
    
    If Height = 0 Then
        HeightAverage = 0
    Else
        HeightAverage = HeightSum / Count
    End If
    
    Print HeightAverage
        
End Sub

Private Sub cmdQ6_Click()
    
    Dim Name As String
    
    Name = InputBox$("Enter your name fam.")
    
    Do While Name <> "zzz"
        If UCase$(Right$(Name, 2)) = "NG" Then
            Print Name
        End If
        Name = InputBox$("Enter your name fam.")
    Loop
    
        
End Sub

Private Sub cmdTest_Click()
    
    Dim Name As String
    Dim DirectoryCount As Integer
    Dim DirectoryName As String
    Dim Swag As Integer
    
    DirectoryName = ""
    DirectoryCount = 0
    
    If MsgBox("Do you want to add more messages", vbYesNo, "Input Messages") = vbYes Then
        Do
            Name = InputBox$("Input a number")
            DirectoryCount = DirectoryCount + 1
            DirectoryName = DirectoryName & Name
            If MsgBox("Do you want to stop", vbYesNo, "Titile") = vbYes Then
                Swag = 1
            End If
        Loop While Swag <> 1
    End If
    
    Print DirectoryName
    Print DirectoryCount
End Sub
