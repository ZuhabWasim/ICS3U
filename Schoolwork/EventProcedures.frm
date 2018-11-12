VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "More Events"
   ClientHeight    =   5430
   ClientLeft      =   3330
   ClientTop       =   1965
   ClientWidth     =   5760
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "EventProcedures.frx":0000
   ScaleHeight     =   6000
   ScaleMode       =   0  'User
   ScaleWidth      =   6000
   Begin VB.TextBox txtQ6 
      Height          =   615
      Left            =   3000
      TabIndex        =   7
      Text            =   "Q6"
      Top             =   4560
      Width           =   1935
   End
   Begin VB.TextBox txtQ5 
      Height          =   615
      Left            =   960
      TabIndex        =   6
      Text            =   "Q5"
      Top             =   4560
      Width           =   1935
   End
   Begin VB.TextBox txtQ4 
      Height          =   360
      Left            =   840
      TabIndex        =   5
      Text            =   "Q4"
      Top             =   4080
      Width           =   1815
   End
   Begin VB.TextBox txtQ3 
      Height          =   615
      Left            =   2640
      TabIndex        =   4
      Text            =   "Q3"
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox txtQ2 
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Text            =   "Q2"
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox txtQ1 
      Height          =   360
      Left            =   480
      TabIndex        =   2
      Text            =   "Q1"
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   4680
      TabIndex        =   0
      Top             =   1320
      Width           =   970
   End
   Begin VB.Image imgUserImage 
      Height          =   480
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblMessage 
      Height          =   975
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const IMGLOCATION = "\\Tdsbshares\schclass$\1392\1392-PickUp\CS\IMAGES\"

Dim ImageLocation As String
Dim ImageFolder As String
Dim ImageName As String


Private Sub cmdExit_Click()

    End
    
End Sub

Private Sub cmdExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lblMessage.Caption = "Click here to exit the program."
    
End Sub

Private Sub Form_Load()
    
    ImageName = UCase$(InputBox$("Enter the name of your image", "Image Name", "Bluejay") & ".BMP")
    ImageFolder = UCase$(InputBox$("Enter the folder it is contained in", "ImageFolder", "bmp") & "\")
    ImageLocation = IMGLOCATION & ImageFolder & ImageName
    
    If ImageName = ".BMP" Or ImageFolder = "\" Then
        MsgBox "No image was loaded", vbOKOnly, "Error: 1"
    Else
        imgUserImage.Picture = LoadPicture(ImageLocation)
    End If
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lblMessage.Caption = ""
    
End Sub


Private Sub imgUserImage_DblClick()
    
    Dim K As Single
    
    If imgUserImage.Left = 0 And imgUserImage.Top = 0 Then
        For K = 1 To frmMain.ScaleHeight - imgUserImage.Height Step 0.01
            imgUserImage.Move K, K
        Next K
    Else
        For K = frmMain.ScaleHeight - imgUserImage.Height To 0 Step -0.01
            imgUserImage.Move K, K
        Next K
    End If
    
End Sub

Private Sub imgUserImage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If ImageName = ".BMP" Then
        lblMessage.Caption = "No Image Loaded"
    Else
        lblMessage.Caption = ImageName
    End If
    
End Sub

Private Sub txtQ1_Click()
    
    txtQ1.Text = ""
    
End Sub

Private Sub txtQ1_KeyPress(KeyAscii As Integer)
    
    Dim Ch As String
    
    Ch = Chr$(KeyAscii)
    
    If Not (Ch >= "a" And Ch <= "z") Or KeyAscii <> 8 Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtQ2_Change()
    
    txtQ2.Text = ""
    
End Sub

Private Sub txtQ2_KeyPress(KeyAscii As Integer)
    
    Dim Ch As String
    
    Ch = Chr$(KeyAscii)
    
    If Not ((Ch >= "a" And Ch <= "z") Or (Ch >= "A" And Ch <= "Z")) Or KeyAscii <> 8 Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtQ3_Change()
    
    txtQ3.Text = ""
    
End Sub

Private Sub txtQ3_KeyPress(KeyAscii As Integer)
    
    Dim Ch As String
    
    Ch = Chr$(KeyAscii)
    
    If Not (Ch >= "0" And Ch <= "9") Or KeyAscii <> 8 Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtQ4_Change()
    
    txtQ4.Text = ""
    
End Sub

Private Sub txtQ4_KeyPress(KeyAscii As Integer)
    
    Dim Ch As String
    
    Ch = Chr$(KeyAscii)
    
    If Ch <> 8 Then
        Ch = UCase$(Ch)
    End If
    
End Sub
