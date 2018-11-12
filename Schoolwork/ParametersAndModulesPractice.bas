Attribute VB_Name = "Module1"
Option Explicit

Const TAXRATE = 0.8                   'Local to the module
Global Const PI = 3.14159265358979    'Applies to the entire project
Global Const MAX = 50

Global Student(1 To MAX) As String    'Visible in the entire project

Public Sub Welcome()
    
    frmMain.lblMessage.Caption = "Welcome to multiple forms"
    
End Sub

Public Sub PrintPattern(ByVal Ch As String, ByVal N As Integer)

    Dim X As Integer
    Dim Y As Integer

    frmMain.Cls
    For X = 1 To N
        For Y = 1 To X
            frmMain.Print Ch;
        Next Y
        frmMain.Print
    Next X

End Sub


Public Function ReverseStr(ByVal Str As String)

    Dim K As Integer
    Dim RevStr As String
    Dim Char As String
    
    frmMain.Cls
    
    RevStr = ""
    
    For K = Len(Str) To 1 Step -1
        Char = Mid$(Str, K, 1)
        RevStr = RevStr & Char
    Next K
    
    ReverseStr = RevStr
    
End Function

Public Function Factorial(ByVal ParaNum As Integer)
    
    Dim K As Integer
    
    K = 0
    
    Do While K > 1
        K = K - 1
        ParaNum = ParaNum * Factorial(ParaNum)
    Loop
        
End Function

Public Function MultString(ByVal Ch As String, ByVal N As Integer)

    Dim K As Integer
    Dim NewStr As String
    
    NewStr = ""
    
    For K = 1 To N
        NewStr = NewStr & Ch
    Next K
    
    MultString = NewStr
    
End Function
