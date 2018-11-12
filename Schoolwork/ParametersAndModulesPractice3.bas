Attribute VB_Name = "Parameters3"
Option Explicit

Public Sub Initialize(ByRef R() As String, ByVal M As Integer)

    Dim K As Integer
    
    For K = 1 To M
        R(K) = ""
    Next K
    
End Sub

Public Sub SetNums(ByRef R() As Integer, ByVal M As Integer, ByVal N As Integer)
    
    Dim K As Integer
    
    For K = 1 To M
        R(K) = N
    Next K
    
End Sub

Public Sub GetLongest(ByRef R() As String, ByVal M As Integer, ByRef L As String)

    Dim K As Integer
    
    L = ""
    
    For K = 1 To M
        If Len(R(K)) > Len(L) Then
            L = R(K)
        End If
    Next K
    
End Sub

Public Sub SumOf(ByRef R() As Single, ByVal M As Integer, ByRef Sum As Single)

    Dim K As Integer
    
    Sum = 0
        
    For K = 1 To M
        Sum = Sum + R(K)
    Next K
    
End Sub

Public Sub DisplayPersons(ByRef PN() As String, ByRef PA() As Integer, ByVal M As Integer)

    Dim K As Integer
    
    frmMain.Print Spc(1); "#"; Tab(10); "Name"; Tab(20); "Age"
    
    For K = 1 To M
        frmMain.Print Format$(K, "@@@"); ". "; Tab(10); PN(K); Tab(20); PA(K)
    Next K
    
End Sub
