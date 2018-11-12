VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Wages Project"
   ClientHeight    =   7050
   ClientLeft      =   210
   ClientTop       =   2220
   ClientWidth     =   12405
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   12405
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   6360
      Width           =   2295
   End
   Begin VB.PictureBox picData 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   120
      ScaleHeight     =   6075
      ScaleWidth      =   12075
      TabIndex        =   2
      Top             =   120
      Width           =   12135
   End
   Begin VB.CommandButton cmdReadData 
      Caption         =   "&Read Data"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   1
      Top             =   6360
      Width           =   2295
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      TabIndex        =   0
      Top             =   6360
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Programmer: Zuhab Wasim
'Data: November 15, 2015
'Purpose: To retrieve the information in a text file of FullName, Wage, and Hours Worked and display the information given aswell as the GrossPay and NetPay of each entry

Private Sub cmdClear_Click()
    
    picData.Cls
    
End Sub

Private Sub cmdExit_Click()
    
    If MsgBox("Are you sure you want to exit?", vbYesNo, "Exit") = vbYes Then
        End
    End If
    
End Sub

Private Sub cmdReadData_Click()

'Declares certain values as CONSTANTS due to the possibility of the value changing in the future
Const DEDUCTION = 0.25
Const OVERTIMEBONUS = 1.5
Const NONOVERTIMEHOURS = 40

'Declares the variables used within the program
Dim FullName As String
Dim Path As String

'Each variable with "R" at the beginning is the variable that will have the right aligned value of the given variable
Dim Wage As Single
Dim RWage As String * 6

Dim Hours As Single
Dim RHours As String * 6
Dim OverTimeHours As Single

Dim GrossPay As Single
Dim RGrossPay As String * 13

Dim NetPay As Single
Dim RNetPay As String * 13

Dim EmployeeCount As Integer
Dim REmployeeCount As String * 4

Dim DeductionAmount As Single
Dim NetPayAverage As Single
Dim NetPaySum As Single

'Clears the picture box of anything that was displayed on it previously
picData.Cls

'Creates the headers for the columns of the values soon to be listed
picData.Print Tab(6); "EMPLOYEE NAME"; Tab(45); "WAGE"; Tab(63); "HOURS"; Tab(84); "GROSS PAY"; Tab(106); "NET PAY"
picData.Print

'Initializes the given values to a number to avoid issues with incremented values such as crashes
EmployeeCount = 0
OverTimeHours = 0
NetPaySum = 0
Path = App.Path

'Checks to see if the path information is usable (having a "\" at the end)
'Places a "\" to allow the path information to be usable for file opening
If Right$(Path, 1) <> "\" Then
    Path = Path & "\"
End If

'Opens file to be used (#1)
Open Path & "wages.txt" For Input As #1

'Continues the loop as long as the file has still information to be read in it
Do While Not EOF(1)
    EmployeeCount = EmployeeCount + 1
    Input #1, FullName, Wage, Hours
    If Hours > NONOVERTIMEHOURS Then
        OverTimeHours = Hours - NONOVERTIMEHOURS
        'If the Hours are greater than what the business deems non-overtime
        'The GrossPay is calculated from the non-overtime with the wage, added with the product of the overtime hours and wage
        GrossPay = (NONOVERTIMEHOURS * Wage) + ((OverTimeHours * Wage) * OVERTIMEBONUS)
    Else
        'If the Hours if less
        'The GrossPay is calculated from the hours multiplied the wage
        GrossPay = Hours * Wage
    End If
    'Calculates the amount deducted from the deduction, the netpay, and the netsum used to find the average of all netpays later
    DeductionAmount = GrossPay * DEDUCTION
    NetPay = GrossPay - DeductionAmount
    NetPaySum = NetPaySum + NetPay
    'Right aligns the values calculated previously into new variables that have the corresponding "R" name
    'Formats the values listed before into a currency standard
    RSet RWage = Format$(Wage, "##0.00")
    RSet RHours = Format$(Hours, "###0.0")
    RSet RGrossPay = Format$(GrossPay, "##,###,##0.00")
    RSet RNetPay = Format$(NetPay, "##,###,##0.00")
    RSet REmployeeCount = Format$(EmployeeCount, "##. ")
    'Displays the values into the picture box
    picData.Print Tab(2); REmployeeCount; FullName; _
                  Tab(40); "$"; Tab(43); RWage; _
                  Tab(62); RHours; _
                  Tab(79); "$"; Tab(80); RGrossPay; _
                  Tab(99); "$"; Tab(100); RNetPay
Loop

'Closes the file that was used (which was #1)
Close #1

'Checks to see if the value of EmployeeCount is not zero to avoid crashing the program
If EmployeeCount = 0 Then
    NetPayAverage = 0
Else
    NetPayAverage = NetPaySum / EmployeeCount
End If

'Displays the employee count and average net pay
picData.Print
picData.Print " Number of employees is: "; EmployeeCount
picData.Print " The average net pay is: "; Format$(NetPayAverage, "currency")

End Sub


