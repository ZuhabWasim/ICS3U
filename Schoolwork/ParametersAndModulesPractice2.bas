Attribute VB_Name = "Module1"
Option Explicit

Public Sub Round(ByVal Num As Single, ByVal RNum As Integer)

    Dim X As Integer
    Dim DecStr As String
    Dim NewStr As String
    
    DecStr = "0."
    
    If RNum = 0 Then
        DecStr = "0"
    Else
        For X = 1 To Abs(RNum)
            DecStr = DecStr & "0"
        Next X
    End If
    
    frmMain.Print Format$(Num, DecStr)
    
End Sub

Public Sub CheckChar(ByVal Ch As String)
    
    Dim AscVal As String

    AscVal = Asc(Ch)
    
    If ((AscVal >= Asc("a")) And (AscVal <= Asc("z"))) Or _
       ((AscVal >= Asc("A")) And (AscVal <= Asc("Z"))) Then
        frmMain.Print "."
    ElseIf (AscVal >= Asc("0") And AscVal <= Asc("9")) Then
        frmMain.Print "-"
    Else
        frmMain.Print "?"
    End If
        
    
End Sub

Public Sub SumOf(ByVal N As Integer)

    Dim Sum As Long
    Dim K As Integer
    
    For K = 1 To N
        Sum = Sum + K
    Next K
    
    frmMain.Print Str$(Sum)
    
End Sub

Public Sub ChangeStr(ByVal St As String)

    Dim K As Integer
    Dim NewSt As String
    Dim Char As String
    
    NewSt = ""
    
    For K = 1 To Len(St)
        Char = Mid$(St, K, 1)
        If ((Asc(Char) >= Asc("a")) And (Asc(Char) <= Asc("z"))) Or _
       ((Asc(Char) >= Asc("A")) And (Asc(Char) <= Asc("Z"))) Then
            NewSt = NewSt & "."
        ElseIf (Asc(Char) >= Asc("0") And Asc(Char) <= Asc("9")) Then
            NewSt = NewSt & "-"
        End If
    Next K
    
    frmMain.Print NewSt
    
End Sub

Public Sub ChangeDate(ByVal D As String)

    Dim St As String
    
    St = UCase$(Format$(D, "dd MMM yyyy"))
    
    frmMain.Print St
    
End Sub

Public Function Factorial(ByVal N As Integer) As Long
    
    If N <= 1 Then
        Factorial = 1
    Else
        Factorial = Factorial(N - 1) * N
    End If
    
End Function

