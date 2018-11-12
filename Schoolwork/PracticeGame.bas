Attribute VB_Name = "Module1"
Option Explicit

Global Const MID = 4440
Global Const RIGHT = 5520
Global Const LEFT = 3360

Public Sub MoveRight()
    
    frmMain.imgUser.Move 5520, (frmMain.imgUser.LEFT + 1080)
    
End Sub

Public Sub MoveLeft()

    frmMain.imgUser.Move 5520, (frmMain.imgUser.LEFT - 1080)
    
End Sub
