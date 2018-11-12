VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control Arrays"
   ClientHeight    =   7050
   ClientLeft      =   1890
   ClientTop       =   1905
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Question 3"
      Height          =   1455
      Left            =   240
      TabIndex        =   33
      Top             =   5160
      Width           =   3735
      Begin VB.PictureBox picShapeColour 
         BackColor       =   &H0000FF00&
         Height          =   375
         Index           =   3
         Left            =   2520
         ScaleHeight     =   315
         ScaleWidth      =   435
         TabIndex        =   37
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox picShapeColour 
         BackColor       =   &H0000FFFF&
         Height          =   375
         Index           =   2
         Left            =   1800
         ScaleHeight     =   315
         ScaleWidth      =   435
         TabIndex        =   36
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox picShapeColour 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   1
         Left            =   1080
         ScaleHeight     =   315
         ScaleWidth      =   435
         TabIndex        =   35
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox picShapeColour 
         BackColor       =   &H000000FF&
         Height          =   375
         Index           =   0
         Left            =   360
         ScaleHeight     =   315
         ScaleWidth      =   435
         TabIndex        =   34
         Top             =   960
         Width           =   495
      End
      Begin VB.Shape shpOval 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   840
         Shape           =   2  'Oval
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Question 2"
      Height          =   1455
      Left            =   120
      TabIndex        =   25
      Top             =   3480
      Width           =   3855
      Begin VB.CommandButton cmdBack 
         Caption         =   "<"
         Height          =   375
         Left            =   2640
         TabIndex        =   31
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdLetter 
         Caption         =   "U"
         Height          =   375
         Index           =   4
         Left            =   2160
         TabIndex        =   30
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdLetter 
         Caption         =   "O"
         Height          =   375
         Index           =   3
         Left            =   1680
         TabIndex        =   29
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdLetter 
         Caption         =   "I"
         Height          =   375
         Index           =   2
         Left            =   1200
         TabIndex        =   28
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdLetter 
         Caption         =   "E"
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   27
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdLetter 
         Caption         =   "A"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblLetter 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   720
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Question 1"
      Height          =   1575
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   3015
      Begin VB.OptionButton optMonth 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Dec"
         Height          =   375
         Index           =   12
         Left            =   2280
         TabIndex        =   24
         Top             =   720
         Width           =   615
      End
      Begin VB.OptionButton optMonth 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nov"
         Height          =   375
         Index           =   11
         Left            =   2280
         TabIndex        =   23
         Top             =   480
         Width           =   615
      End
      Begin VB.OptionButton optMonth 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Oct"
         Height          =   375
         Index           =   10
         Left            =   2280
         TabIndex        =   22
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optMonth 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sep"
         Height          =   375
         Index           =   9
         Left            =   1560
         TabIndex        =   21
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton optMonth 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Aug"
         Height          =   375
         Index           =   8
         Left            =   1560
         TabIndex        =   20
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton optMonth 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Jul"
         Height          =   375
         Index           =   7
         Left            =   1560
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optMonth 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Jun"
         Height          =   375
         Index           =   6
         Left            =   840
         TabIndex        =   18
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton optMonth 
         BackColor       =   &H00C0C0C0&
         Caption         =   "May"
         Height          =   375
         Index           =   5
         Left            =   840
         TabIndex        =   17
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton optMonth 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Apr"
         Height          =   375
         Index           =   4
         Left            =   840
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optMonth 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mar"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton optMonth 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Feb"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton optMonth 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Jan"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optMonth 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Option1"
         Height          =   375
         Index           =   0
         Left            =   3600
         TabIndex        =   12
         Top             =   1680
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblMonth 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   2775
      End
   End
   Begin VB.PictureBox picMain 
      Height          =   1335
      Left            =   2760
      ScaleHeight     =   1275
      ScaleWidth      =   2715
      TabIndex        =   9
      Top             =   240
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin VB.PictureBox picColour 
         BackColor       =   &H0000FF00&
         Height          =   505
         Index           =   7
         Left            =   1920
         ScaleHeight     =   450
         ScaleMode       =   0  'User
         ScaleWidth      =   505
         TabIndex        =   8
         Top             =   840
         Width           =   505
      End
      Begin VB.PictureBox picColour 
         BackColor       =   &H00FF00FF&
         Height          =   505
         Index           =   6
         Left            =   1920
         ScaleHeight     =   450
         ScaleMode       =   0  'User
         ScaleWidth      =   505
         TabIndex        =   7
         Top             =   240
         Width           =   505
      End
      Begin VB.PictureBox picColour 
         BackColor       =   &H00FF0000&
         Height          =   505
         Index           =   5
         Left            =   720
         ScaleHeight     =   450
         ScaleMode       =   0  'User
         ScaleWidth      =   505
         TabIndex        =   6
         Top             =   840
         Width           =   505
      End
      Begin VB.PictureBox picColour 
         BackColor       =   &H0000FFFF&
         Height          =   505
         Index           =   4
         Left            =   120
         ScaleHeight     =   450
         ScaleMode       =   0  'User
         ScaleWidth      =   505
         TabIndex        =   5
         Top             =   840
         Width           =   505
      End
      Begin VB.PictureBox picColour 
         BackColor       =   &H000000FF&
         Height          =   505
         Index           =   3
         Left            =   1320
         ScaleHeight     =   450
         ScaleMode       =   0  'User
         ScaleWidth      =   505
         TabIndex        =   4
         Top             =   840
         Width           =   505
      End
      Begin VB.PictureBox picColour 
         BackColor       =   &H00000000&
         Height          =   505
         Index           =   2
         Left            =   1320
         ScaleHeight     =   450
         ScaleMode       =   0  'User
         ScaleWidth      =   505
         TabIndex        =   3
         Top             =   240
         Width           =   505
      End
      Begin VB.PictureBox picColour 
         BackColor       =   &H00C0C0C0&
         Height          =   505
         Index           =   1
         Left            =   720
         ScaleHeight     =   450
         ScaleMode       =   0  'User
         ScaleWidth      =   505
         TabIndex        =   2
         Top             =   240
         Width           =   505
      End
      Begin VB.PictureBox picColour 
         BackColor       =   &H00FFFFFF&
         Height          =   500
         Index           =   0
         Left            =   120
         ScaleHeight     =   435
         ScaleMode       =   0  'User
         ScaleWidth      =   488.167
         TabIndex        =   1
         Top             =   240
         Width           =   500
      End
      Begin VB.Shape shpBorder 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   4
         Height          =   510
         Left            =   120
         Top             =   240
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
    
    If Not (Len(lblLetter.Caption)) = 0 Then
        lblLetter.Caption = Left$(lblLetter.Caption, Len(lblLetter.Caption) - 1)
    End If
    
End Sub

Private Sub cmdLetter_Click(Index As Integer)
    
    lblLetter.Caption = lblLetter.Caption & cmdLetter(Index).Caption
    
End Sub

Private Sub optMonth_Click(Index As Integer)
    
    lblMonth.Caption = optMonth(Index).Caption & " is month # " & Str$(Index)
    
End Sub

Private Sub picColour_Click(Index As Integer)
    
    picMain.BackColor = picColour(Index).BackColor
    shpBorder.Left = picColour(Index).Left - 5
    shpBorder.Top = picColour(Index).Top - 5
    
End Sub

Private Sub picShapeColour_Click(Index As Integer)
    
    shpOval.FillColor = picShapeColour(Index).BackColor
    
End Sub
