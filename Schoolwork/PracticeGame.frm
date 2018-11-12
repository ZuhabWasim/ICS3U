VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Game"
   ClientHeight    =   7080
   ClientLeft      =   1890
   ClientTop       =   1905
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   10035
   Begin VB.Image imgCar1 
      Height          =   1095
      Left            =   7080
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Image imgUser 
      Height          =   1320
      Left            =   3360
      Picture         =   "PracticeGame.frx":0000
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   705
   End
   Begin VB.Image imgRoad 
      Height          =   7095
      Left            =   2640
      Picture         =   "PracticeGame.frx":1DFE4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4380
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyRight And imgUser.LEFT <> RIGHT Then
        MoveRight
    End If
        
End Sub

Private Sub Form_Load()
    
    imgUser.Move frmMain.ScaleHeight, MID
    
End Sub
