VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "More Events"
   ClientHeight    =   7605
   ClientLeft      =   1890
   ClientTop       =   1905
   ClientWidth     =   6585
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
   ScaleHeight     =   7605
   ScaleWidth      =   6585
   Begin VB.TextBox txtNumber 
      Height          =   615
      Left            =   3120
      TabIndex        =   2
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Enter a number:"
      Height          =   855
      Left            =   480
      TabIndex        =   1
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblKey 
      Height          =   855
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Pass As String

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    lblKey.Caption = "You pressed ASCII value " & Str$(KeyAscii) & " or " & Chr$(KeyAscii)
    
End Sub

Private Sub txtNumber_KeyPress(KeyAscii As Integer)
    
    Dim Ch As String
    
    Ch = Chr$(KeyAscii)
    
'           Question 1
'   If KeyAscii <> 8 And _
    (Not(Ch >= "a" Or Ch <= "z")) Then
'        KeyAscii = 0
'   End If

'           Question 2
'   If KeyAscii <> 8 And _
    Not((Ch >= "a" Or Ch <= "z") Or (Ch >= "A" Or Ch <= "Z")) Then
'        KeyAscii = 0
'   End If

'           Question 3
'   If KeyAscii <> 8 And _
    (Ch < "0" Or Ch > "9") Then
'        KeyAscii = 0
'   End If

'           Question 4
'   KeyAscii = Asc(UCase$(Chr$(KeyAscii))

'           Question 5
'   If KeyAscii <> 8 And _
    Not((Ch >= "a" Or Ch <= "z") Or (Ch >= "A" Or Ch <= "Z") Or (Ch >= "0" Or Ch <= "9")) Then
'        KeyAscii = 0
'   End If
            
'           Question 6

    If KeyAscii = 32 Or KeyAscii = 13 Then
        Print Pass
        Pass = ""
    Else
        If KeyAscii <> 8 And (Ch >= "0" And Ch <= "9") Then
            Pass = Pass & Chr$(KeyAscii)
            KeyAscii = Asc("*")
        Else
            txtNumber.Text = ""
        End If
    End If
    
End Sub
