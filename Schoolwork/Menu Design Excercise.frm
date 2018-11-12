VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menu"
   ClientHeight    =   7605
   ClientLeft      =   1890
   ClientTop       =   2205
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   6585
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuDisplay 
      Caption         =   "&Display"
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAll 
         Caption         =   "&All"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuAbout_Click()
    
    MsgBox "Menu Design and File Access", vbOKOnly, "About"
    
End Sub

Private Sub mnuExit_Click()
    
    If MsgBox("Are you sure you want to exit?", vbYesNo, "End") = vbYes Then
        End
    End If
    
End Sub
