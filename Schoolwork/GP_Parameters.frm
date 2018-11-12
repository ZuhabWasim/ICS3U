VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "General Procedures Examples"
   ClientHeight    =   3555
   ClientLeft      =   2325
   ClientTop       =   2475
   ClientWidth     =   6405
   Icon            =   "GP_Parameters.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6405
   Begin VB.PictureBox picData 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   120
      ScaleHeight     =   3225
      ScaleWidth      =   3825
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   3855
   End
   Begin VB.CommandButton cmdPattern 
      Caption         =   "Pattern - GP1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPattern_Click()
    Dim Char As String
    Dim Lines As Integer
    
    Char = InputBox$("Enter a character to display:", , "*")
    Lines = Val(InputBox$("Enter the number of lines to display:", , "5"))
    
    PrintPattern Char, Lines
End Sub

