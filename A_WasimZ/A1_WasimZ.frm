VERSION 5.00
Begin VB.Form frmFood 
   BackColor       =   &H00C0C000&
   Caption         =   "A1_WasimZ"
   ClientHeight    =   3570
   ClientLeft      =   1890
   ClientTop       =   1905
   ClientWidth     =   4890
   FillColor       =   &H000000FF&
   FillStyle       =   6  'Cross
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   4890
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C000&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   12
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00C0C000&
      Caption         =   "C&lear"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   11
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmdFood 
      BackColor       =   &H00C0C000&
      Caption         =   "&Calculate"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Frame fraFood 
      BackColor       =   &H00FFFF00&
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   2415
      Begin VB.CheckBox chkDrink 
         BackColor       =   &H00FFFF00&
         Caption         =   "Soft Drink"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CheckBox chkFries 
         BackColor       =   &H00FFFF00&
         Caption         =   "French Fries"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1455
      End
      Begin VB.CheckBox chkSandwich 
         BackColor       =   &H00FFFF00&
         Caption         =   "Sandwich"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFF00&
         Caption         =   "Select all three for a combo price of $5.29!"
         Height          =   495
         Left            =   240
         TabIndex        =   16
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFF00&
         Caption         =   "$0.99"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   15
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFF00&
         Caption         =   "$1.59"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   14
         Top             =   780
         Width           =   495
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFF00&
         Caption         =   "$3.69"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "Select what you would like to order in the menu box and click calculate to see the price!"
      Height          =   495
      Left            =   360
      TabIndex        =   18
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "Welcome to [Insert Restaurant]!"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   9
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label lblTax 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   8
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblSubTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   7
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C000&
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Caption         =   "Tax:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "SubTotal:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   1800
      Width           =   855
   End
End
Attribute VB_Name = "frmFood"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Programmer: Zuhab Wasim
'Date: 23/09/15
'Purpose: To provide a program to easily calculate the price of food orders given.

Private Sub cmdClear_Click()
    
'Clears the values inputed by the user in check boxes and clears the labels
    chkSandwich.Value = 0
    chkFries.Value = 0
    chkDrink.Value = 0
    
    lblSubTotal.Caption = ""
    lblTotal.Caption = ""
    lblTax.Caption = ""
    
End Sub

Private Sub cmdExit_Click()

'Exits the program if the User confirms, otherwise it will not exit
    If MsgBox("Are you sure you want to exit?", vbYesNo, "Exit") = vbYes Then
        End
    End If
        
End Sub

Private Sub cmdFood_Click()

'Declares the price of food items on the menu as well as tax as constants to be able to be easily changed in the future if needed
    Const TAX = 0.13
    Const SandWich_Price = 3.69
    Const Fries_Price = 1.59
    Const Drink_Price = 0.99
    Const Combo_Price = 5.29

'Declares the variables that do not initially have inputed values
    Dim SubTotal As Single
    Dim Total As Single
    Dim TaxAmount As Single

'Initially assigns the variables without values to 0
    SubTotal = 0
    
'Checks to see if the User has selected any item from the Food Menu
    If chkSandwich = 0 And chkFries = 0 And chkDrink = 0 Then
        MsgBox "Please select a product from the food menu.", vbInformation, "ERROR: No item selected"
    Else
'Checks to see if the User has selected all the items from the Food Menu, if so then it the SubTotal will be assigned the Combo Price
        If chkSandwich = 1 And chkFries = 1 And chkDrink = 1 Then
            SubTotal = Combo_Price
        Else
            If chkSandwich = 1 Then
                SubTotal = SubTotal + SandWich_Price
            End If
            If chkFries = 1 Then
                SubTotal = SubTotal + Fries_Price
            End If
            If chkDrink = 1 Then
                SubTotal = SubTotal + Drink_Price
            End If
        End If
'Checks to see if the SubTotal is above 4 dollars, if so then tax is applied
        If SubTotal >= 4 Then
            TaxAmount = SubTotal * TAX
        Else
            TaxAmount = 0
        End If
    
        Total = SubTotal + TaxAmount

'Displays the Total, SubTotal, and Tax amount to the user
        lblSubTotal.Caption = Format$(SubTotal, "currency")
        lblTax.Caption = Format$(TaxAmount, "currency")
        lblTotal.Caption = Format$(Total, "currency")
    End If

End Sub


