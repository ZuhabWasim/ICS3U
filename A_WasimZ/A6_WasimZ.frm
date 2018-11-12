VERSION 5.00
Begin VB.Form frmPacGround 
   Caption         =   "Animaion - Displaying Images (Key Down)"
   ClientHeight    =   2430
   ClientLeft      =   1890
   ClientTop       =   1905
   ClientWidth     =   5760
   Icon            =   "A6_WasimZ.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleMode       =   0  'User
   ScaleWidth      =   6000
   Begin VB.Image imgPacman 
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   615
   End
   Begin VB.Image imgPacDown 
      Height          =   480
      Left            =   3360
      Picture         =   "A6_WasimZ.frx":08CA
      Top             =   960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPacUp 
      Height          =   480
      Left            =   2760
      Picture         =   "A6_WasimZ.frx":150C
      Top             =   960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPacRight 
      Height          =   480
      Left            =   2160
      Picture         =   "A6_WasimZ.frx":214E
      Top             =   960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPacLeft 
      Height          =   480
      Left            =   1560
      Picture         =   "A6_WasimZ.frx":2D90
      Top             =   960
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmPacGround"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Name: Zuhab Wasim 11G
'Date: 22/02/16
'Purpose: To demonstrate the use of the KeyDown procedure in Visual Basic 6.0 in which an image of _
          a pacman is moved according to the appropriate key press of Up, Down, Left, Right. The _
          image will then wrap around to the appropriate other side of the form once the image goes _
          off screen.
          
Option Explicit

'Declares variabes used within multiple procedures as form variables
'Declraes X and Y values of Pacman
Dim CurrX As Integer
Dim CurrY As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '----------KeyDown Procedure of Program: To move imgPacman----------
        
    'Determines which key the user is pressing (Up, Down, Right, Left)
    If KeyCode = vbKeyRight Then
        'Determines if the imgPacman is completely outside the form, and moves imgPacman to the other side of the form if so
        If imgPacman.Left >= frmPacGround.ScaleWidth Then
            CurrX = 0 - imgPacman.Width
        End If
        'Increases X-Value of imgPacman by 100, giving the location of imgPacman 100 units to the right
        CurrX = CurrX + 100
        'Replaces imgPacman's picture value to the appropriate picture relating to KeyDown
        imgPacman.Picture = imgPacRight.Picture
    ElseIf KeyCode = vbKeyLeft Then
        If imgPacman.Left <= 0 - imgPacman.Width Then
            CurrX = frmPacGround.ScaleWidth
        End If
        'Decreases X-Value of imgPacman by 100, giving the location of imgPacman 100 units to the left
        CurrX = CurrX - 100
        'Replaces imgPacman's picture value to the appropriate picture relating to KeyDown
        imgPacman.Picture = imgPacLeft.Picture
    ElseIf KeyCode = vbKeyUp Then
        If imgPacman.Top <= 0 - imgPacman.Height Then
            CurrY = frmPacGround.ScaleHeight
        End If
        'Decreases Y-Value of imgPacman by 100, giving the location of imgPacman 100 units upwards
        CurrY = CurrY - 100
        'Replaces imgPacman's picture value to the appropriate picture relating to KeyDown
        imgPacman.Picture = imgPacUp.Picture
    ElseIf KeyCode = vbKeyDown Then
        If imgPacman.Top >= frmPacGround.ScaleHeight Then
            CurrY = 0 - imgPacman.Height
        End If
        'Increases Y-Value of imgPacman by 100, giving the location of imgPacman 100 units downwards
        CurrY = CurrY + 100
        'Replaces imgPacman's picture value to the appropriate picture relating to KeyDown
        imgPacman.Picture = imgPacDown.Picture
    Else
        'If the keydown is not valid (not Up, Down, Right, Left), an error message is displayed
        MsgBox "Key Invalid: Please use arrow keys", vbOKOnly, "Error 1: Invalid Key Press"
    End If
    
    'Moves imgPacman to the updated X and Y values determined
    imgPacman.Move CurrX, CurrY
    
End Sub

Private Sub Form_Load()
    
    'Initiliazes the X and Y values of imgPacman's location to initiliaze the starting position of imgPacman
    CurrX = 0
    CurrY = 0
    'Sets imgPacman's appropriate image and moves it to the values of CurrX and CurrY
    imgPacman.Picture = imgPacRight.Picture
    imgPacman.Move CurrX, CurrY
    
End Sub
