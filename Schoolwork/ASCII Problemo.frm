VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ASCIIIIITHIHNSn"
   ClientHeight    =   7605
   ClientLeft      =   1890
   ClientTop       =   1905
   ClientWidth     =   7005
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   7005
   Begin VB.CommandButton cmdASCII 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   5160
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdASCII_Click()

        Dim K As Integer

        For K = 32 To 127
            If K < 100 Then
                Print " ";
            End If
            Print Str$(K); " "; Chr(K),
            If K Mod 5 = 1 Then
                Print
            End If
        Next K
        
End Sub

