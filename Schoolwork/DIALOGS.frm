VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7605
   ClientLeft      =   1890
   ClientTop       =   2205
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   6585
   Begin VB.Menu mnuDialogs 
      Caption         =   "Dialogs"
      Begin VB.Menu mnuOkCancel 
         Caption         =   "OK-Cancel Dialog"
      End
      Begin VB.Menu mnuAbortRetryIgnore 
         Caption         =   "Abort-Retry-Ignore Dialog"
      End
      Begin VB.Menu mnuYesNoCancel 
         Caption         =   "Yes-No-Cancel Dialog"
      End
      Begin VB.Menu mnuYesNo 
         Caption         =   "Yes-No Dialog"
      End
      Begin VB.Menu mnuRetryCancel 
         Caption         =   "Retry-Cancel Dialog"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuDialogs2 
      Caption         =   "Dialogs 2"
      Begin VB.Menu mnuA 
         Caption         =   "Question 1. a)"
      End
      Begin VB.Menu mnuB 
         Caption         =   "Question 1. b)"
      End
      Begin VB.Menu mnuA2 
         Caption         =   "Question 2. a)"
      End
      Begin VB.Menu mnuB2 
         Caption         =   "Question 2. b)"
      End
      Begin VB.Menu mnu3 
         Caption         =   "Question 3"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub mnu3_Click()

    Dim DType As Integer
    Dim DTitle As String
    Dim DMsg As String
    Dim Response As Integer
    
    DType = vbInformation + vbYesNo
    DTitle = "Example"
    DMsg = "Display a Message"
    
    Response = MsgBox(DMsg, DType, DTitle)
    
    If Response = vbYes Then
        MsgBox "You can proceed", vbOKOnly, DTitle
    Else
        MsgBox "Stop!", vbOKOnly, DTitle
    End If
        
End Sub

Private Sub mnuA_Click()
    
    Dim DType As Integer
    Dim DTitle As String
    Dim DMsg As String
    Dim Response As Integer
    
    DType = vbOKCancel + vbExclamation + vbDefaultButton1
    DTitle = "Exit Program"
    DMsg = "Do you want to exit ?"
    
    Response = MsgBox(DMsg, DType, DTitle)
    
End Sub

Private Sub mnuA2_Click()

    Dim DType As Integer
    Dim DTitle As String
    Dim DMsg As String
    Dim Response As Integer
    
    DType = vbQuestion + vbYesNoCancel
    DTitle = "Excel"
    DMsg = "Do you wish to save ?"
    
    Response = MsgBox(DMsg, DType, DTitle)
    
End Sub

Private Sub mnuAbortRetryIgnore_Click()
    
    Dim DType As Integer
    Dim DTitle As String
    Dim DMsg As String
    Dim Response As Integer
    
    DType = vbAbortRetryIgnore + vbExclamation
    DTitle = "Demonstration"
    DMsg = "Select a button."
    
    Response = MsgBox(DMsg, DType, DTitle)
    DType = vbInformation
    
    If Response = vbAbort Then
        MsgBox "You selected ABORT.", DType, DTitle
    ElseIf Response = vbRetry Then
        MsgBox "You selected RETRY.", DType, DTitle
    ElseIf Response = vbIgnore Then
        MsgBox "You selected IGNORE.", DType, DTitle
    End If
    
End Sub

Private Sub mnuB_Click()
    
    Dim DType As Integer
    Dim DTitle As String
    Dim DMsg As String
    Dim Response As Integer
    
    DType = vbAbortRetryIgnore + vbDefaultButton2
    DTitle = ""
    DMsg = "Do you wish to continue ?"
    
    Response = MsgBox(DMsg, DType, DTitle)
    
End Sub

Private Sub mnuB2_Click()
    
    Dim DType As Integer
    Dim DTitle As String
    Dim DMsg As String
    Dim Response As Integer
    
    DType = vbExclamation + vbRetryCancel
    DTitle = "Claris Works"
    DMsg = "Operation failed"
    
    Response = MsgBox(DMsg, DType, DTitle)
    
End Sub

Private Sub mnuExit_Click()
    
    Dim DType As Integer
    Dim DTitle As String
    Dim DMsg As String
    Dim Response As Integer
    
    DType = vbYesNo + vbQuestion
    DTitle = "Termination"
    DMsg = "Are you sure you want to exit?"
    
    Response = MsgBox(DMsg, DType, DTitle)
    
    If Response = vbYes Then
        End
    End If
    
End Sub

Private Sub mnuOkCancel_Click()
    
    Dim DType As Integer
    Dim DTitle As String
    Dim DMsg As String
    Dim Response As Integer
    
    DType = vbOKCancel + vbInformation
    DTitle = "Demonstration"
    DMsg = "Select a button."
    
    Response = MsgBox(DMsg, DType, DTitle)
    DType = vbInformation
    
    If Response = vbOK Then
        MsgBox "You selected OK.", DType, DTitle
    Else
        MsgBox "You selected Cancel.", DType, DTitle
    End If
    
End Sub

Private Sub mnuRetryCancel_Click()
    
    Dim DType As Integer
    Dim DTitle As String
    Dim DMsg As String
    Dim Response As Integer
    
    DType = vbRetryCancel + vbInformation
    DTitle = "Demonstration"
    DMsg = "Select a button."
    
    Response = MsgBox(DMsg, DType, DTitle)
    DType = vbInformation
    
    If Response = vbRetry Then
        MsgBox "You selected Retry", DType, DTitle
    Else
        MsgBox "You selected Cancel.", DType, DTitle
    End If
    
End Sub

Private Sub mnuYesNo_Click()
    
    Dim DType As Integer
    Dim DTitle As String
    Dim DMsg As String
    Dim Response As Integer
    
    DType = vbYesNoCancel + vbInformation
    DTitle = "Demonstration"
    DMsg = "Select a button."
    
    Response = MsgBox(DMsg, DType, DTitle)
    DType = vbInformation
    
    If Response = vbYes Then
        MsgBox "You selected Yes.", DType, DTitle
    ElseIf Response = vbNo Then
        MsgBox "You selected No.", DType, DTitle
    Else
        MsgBox "You selected Cancel.", DType, DTitle
    End If
    
End Sub

Private Sub mnuYesNoCancel_Click()

    Dim DType As Integer
    Dim DTitle As String
    Dim DMsg As String
    Dim Response As Integer
    
    DType = vbYesNo + vbInformation
    DTitle = "Demonstration"
    DMsg = "Select a button."
    
    Response = MsgBox(DMsg, DType, DTitle)
    DType = vbInformation
    
    If Response = vbYes Then
        MsgBox "You selected Yes.", DType, DTitle
    Else
        MsgBox "You selected No.", DType, DTitle
    End If
    
End Sub
