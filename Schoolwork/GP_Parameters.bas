Attribute VB_Name = "Module1"
Option Explicit

' General procedure that prints a pattern based on a
' character Ch and a number Num (that are passed in).

' Notice that the parameters Ch and Num are independent of
' the variables (Char and Lines) used in the main program.
Public Sub PrintPattern(ByVal Ch As String, ByVal Num As Integer)
    Dim X As Integer
    Dim Y As Integer
    
    frmMain.picData.Cls
    For X = 1 To Num
        For Y = 1 To X
            frmMain.picData.Print Ch;
        Next Y
        frmMain.picData.Print
    Next X
End Sub

