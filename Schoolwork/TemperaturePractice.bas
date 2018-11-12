Attribute VB_Name = "Module1"
Option Explicit

Public Sub GetFile(FName As String, B As Boolean)
    
    With frmMain.cdlDialog
        .FileName = ""
        .Filter = "Text Files|*.txt|All Files|*.*"
        .InitDir = App.Path
        .ShowOpen
    End With
    
    If frmMain.cdlDialog.FileName <> "" Then
        FName = frmMain.cdlDialog.FileName
        B = True
    Else
        B = False
    End If
    
End Sub

Public Sub ReadFile(ByVal FName As String, ByRef C() As String, ByRef T() As Single, Count As Integer)

    Dim K As Integer
    
    K = 0
    
    Open App.Path & FName For Input As #1
    
    Do While Not EOF(1)
        K = K + 1
        Input #1, C(K), T(K)
    Loop
    
    Count = K
    
End Sub
