VERSION 5.00
Begin VB.Form frmMenu 
   AutoRedraw      =   -1  'True
   Caption         =   "The Menu Program"
   ClientHeight    =   5190
   ClientLeft      =   1890
   ClientTop       =   2205
   ClientWidth     =   8910
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
   ScaleHeight     =   5190
   ScaleWidth      =   8910
   Begin VB.PictureBox picCust 
      Height          =   3735
      Left            =   120
      ScaleHeight     =   3675
      ScaleWidth      =   8595
      TabIndex        =   0
      Top             =   120
      Width           =   8655
   End
   Begin VB.Label lblDis 
      Height          =   495
      Left            =   7440
      TabIndex        =   1
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open   Ctrl+O"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save    Ctrl+S"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDisplay 
         Caption         =   "Display   Ctrl+D"
      End
      Begin VB.Menu mnuSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuForm 
      Caption         =   "&Form"
      Begin VB.Menu mnuSetColour 
         Caption         =   "&Background Colour"
         Begin VB.Menu mnuWhite 
            Caption         =   "&White"
         End
         Begin VB.Menu mnuRed 
            Caption         =   "&Red"
         End
         Begin VB.Menu mnuBlue 
            Caption         =   "&Blue"
         End
      End
      Begin VB.Menu mnuSetSize 
         Caption         =   "Set Si&ze"
         Begin VB.Menu mnuSmall 
            Caption         =   "&Small"
         End
         Begin VB.Menu mnuLarge 
            Caption         =   "&Large"
         End
      End
      Begin VB.Menu mnuFontSize 
         Caption         =   "F&ont Size"
      End
      Begin VB.Menu mnuForegroundColour 
         Caption         =   "Fore&gound  Colour"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MAX = 50

Dim CustName(1 To MAX) As String
Dim CustBDate(1 To MAX) As String
Dim CustStatus(1 To MAX) As String
Dim CustCount As Integer
Dim PathLocation As String

Private Sub cmdExit_Click()
    
    If MsgBox("Are you sure you want to exit?", vbYesNo, "Exit") = vbYes Then
        End
    End If
    
End Sub

Private Sub cmdTest_Click()

    frmMenu.Print "Menu"
    
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    
    Dim Ch As String
    
    picCust.Cls
    
    Ch = Chr$(KeyAscii)
    
    lblDis.Caption = Ch
    'If KeyAscii = Ascii(Ctrl) Then
        
    
End Sub

Private Sub Form_Load()
    
    Dim K As Integer
    
    For K = 1 To MAX
        CustName(K) = ""
        CustBDate(K) = ""
        CustStatus(K) = ""
    Next K
    
    frmMenu.BackColor = QBColor(15)
    mnuWhite.Enabled = False
    mnuSmall.Enabled = False
    
End Sub

Private Sub mnuAbout_Click()
        
    Dim Msg As String
    
    Msg = "Copyright 2015 W. Zuhab"
    MsgBox Msg, vbInformation, "About"
    
End Sub

Private Sub mnuBlue_Click()
    
    frmMenu.BackColor = QBColor(1)
    mnuBlue.Enabled = False
    mnuRed.Enabled = True
    mnuWhite.Enabled = True
    
End Sub

Private Sub mnuDisplay_Click()
    
    Dim K As Integer
    
    picCust.Cls
    Cls
    picCust.Print Spc(1); "Customer Name:"; Tab(30); "BirthDate:"; Tab(60); "Status:"
    
    For K = 1 To CustCount
        picCust.Print Spc(2); CustName(K); Tab(31); CustBDate(K); Tab(61); CustStatus(K)
    Next K
    
    For K = 1 To 17
        Print
    Next K
    
    Print "There were "; Str$(CustCount); " records read from:"
    Print PathLocation
    
End Sub

Private Sub mnuExit_Click()
    
    If MsgBox("Are you sure want to exit?", vbYesNo, "Exit") = vbYes Then
        End
    End If
    
End Sub

Private Sub mnuFontSize_Click()
    
    Dim Font As Single
    
    Font = Val(InputBox$("Please enter a font size. [8,10,12,14]", "FontSize", 10))
    
    If Font > 24 Then
        MsgBox "Font is too big!", vbCritical Or vbOKOnly, "Error"
    ElseIf Font < 6 Then
        MsgBox "Font is too small!", vbCritical Or vbOKOnly, "Error"
    End If
    
    frmMenu.FontSize = Font
        
End Sub

Private Sub mnuLarge_Click()

    frmMenu.WindowState = 2
    mnuSmall.Enabled = True
    mnuLarge.Enabled = False
    
End Sub

Private Sub mnuOpen_Click()
    
    Dim FileName As String
    Dim TxtFinder As String

    FileName = InputBox$("Enter the name of the file", "File Open", ".txt")
    TxtFinder = Right$(FileName, 4)
    
    If Not (TxtFinder = ".txt") Then
        FileName = FileName & ".txt"
    End If
    
    If Right$(App.Path, 1) <> "\" Then
        PathLocation = App.Path & "\" & FileName
    Else
        PathLocation = App.Path & FileName
    End If
    
    Read
    
    Close #1
    
End Sub
Private Sub Read()
    
    Dim K As Integer
        
    K = 0
        
    Open PathLocation For Input As #1
    
    Do While Not EOF(1)
        K = K + 1
        Input #1, CustName(K)
        Input #1, CustBDate(K)
        Input #1, CustStatus(K)
    Loop
    
    CustCount = K
    
End Sub
Private Sub mnuRed_Click()
    
    frmMenu.BackColor = QBColor(4)
    mnuRed.Enabled = False
    mnuBlue.Enabled = True
    mnuWhite.Enabled = True
    
End Sub


Private Sub mnuSmall_Click()
    
    frmMenu.WindowState = 0
    mnuSmall.Enabled = False
    mnuLarge.Enabled = True
    
End Sub

Private Sub mnuWhite_Click()
    
    frmMenu.BackColor = QBColor(15)
    mnuWhite.Enabled = False
    mnuBlue.Enabled = True
    mnuRed.Enabled = True
    
End Sub
