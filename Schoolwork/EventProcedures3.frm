VERSION 5.00
Begin VB.Form frmPacGround 
   Caption         =   "Key Down Practice"
   ClientHeight    =   3255
   ClientLeft      =   1890
   ClientTop       =   1905
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   ScaleHeight     =   3596.685
   ScaleMode       =   0  'User
   ScaleWidth      =   6687.5
   Begin VB.Image imgPacman2 
      Height          =   615
      Left            =   1680
      Top             =   1680
      Width           =   735
   End
   Begin VB.Image imgPacRight 
      Height          =   480
      Left            =   2280
      Picture         =   "EventProcedures3.frx":0000
      Top             =   2760
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPacLeft 
      Height          =   480
      Left            =   1680
      Picture         =   "EventProcedures3.frx":0C42
      Top             =   2760
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPacUp 
      Height          =   480
      Left            =   1080
      Picture         =   "EventProcedures3.frx":1884
      Top             =   2760
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPacDown 
      Height          =   480
      Left            =   480
      Picture         =   "EventProcedures3.frx":24C6
      Top             =   2760
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPacman 
      Height          =   480
      Left            =   0
      Picture         =   "EventProcedures3.frx":3108
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "frmPacGround"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CurrX As Integer
Dim CurrY As Integer
Dim CurrX2 As Integer
Dim CurrY2 As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyRight Then
        imgPacman.Picture = imgPacRight.Picture
        If imgPacman.Left >= frmPacGround.ScaleWidth - imgPacman.Left Then
            imgPacman2.Visible = True
            imgPacman2.Picture = imgPacman.Picture
            CurrX2 = imgPacman.Left - frmPacGround.ScaleWidth
            CurrY2 = imgPacman.Top
        ElseIf imgPacman.Left >= frmPacGround.ScaleWidth Then
            imgPacman2.Visible = False
            CurrX = imgPacman2.Left
            CurrY = imgPacman2.Top
        End If
            'CurrX = 0 - imgPacman.Width
        CurrX = CurrX + 100
    ElseIf KeyCode = vbKeyLeft Then
        If imgPacman.Left <= 0 - imgPacman.Width Then
            CurrX = frmPacGround.ScaleWidth
        End If
        CurrX = CurrX - 100
        imgPacman.Picture = imgPacLeft.Picture
    ElseIf KeyCode = vbKeyUp Then
        If imgPacman.Top <= 0 - imgPacman.Height Then
            CurrY = frmPacGround.ScaleHeight
        End If
        CurrY = CurrY - 100
        imgPacman.Picture = imgPacUp.Picture
    ElseIf KeyCode = vbKeyDown Then
        If imgPacman.Top >= frmPacGround.ScaleHeight Then
            CurrY = 0 - imgPacman.Height
        End If
        CurrY = CurrY + 100
        imgPacman.Picture = imgPacDown.Picture
    Else
        MsgBox "Please use arrow keys", vbOKOnly, "Error"
    End If
    
    imgPacman2.Move CurrX2, CurrY2
    imgPacman.Move CurrX, CurrY
    
End Sub

Private Sub Form_Load()
    
    CurrX = 0
    CurrY = 0
    CurrX2 = 0
    CurrY2 = 0
    imgPacman.Move CurrX, CurrY
    imgPacman2.Visible = False
    
    Randomize
End Sub
