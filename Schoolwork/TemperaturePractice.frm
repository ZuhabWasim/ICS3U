VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Temperatures - Menus, Common Dialog and GPs"
   ClientHeight    =   7605
   ClientLeft      =   1890
   ClientTop       =   2205
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   6585
   Begin MSComDlg.CommonDialog cdlDialog 
      Left            =   5520
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picData 
      Height          =   2535
      Left            =   120
      ScaleHeight     =   2475
      ScaleWidth      =   6075
      TabIndex        =   0
      Top             =   240
      Width           =   6135
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuCal 
         Caption         =   "Calculate"
      End
      Begin VB.Menu mnuSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MAX = 50

Dim FileName As String
Dim CountryName(1 To MAX) As String
Dim CountryTemp(1 To MAX) As Single
Dim CountryCount As Integer
Dim IsFileName As Boolean

Private Sub Form_Load()
    
    Dim K As Integer
    
    CountryCount = 0
    IsFileName = True
    
    For K = 1 To MAX
        CountryName(K) = ""
        CountryTemp(K) = 0
    Next K
    
End Sub

Private Sub mnuOpen_Click()
    
    GetFile FileName, IsFileName
    
    If IsFileName = True Then
        ReadFile FileName, CountryName(), CountryTemp(), CountryCount
    End If
End Sub

