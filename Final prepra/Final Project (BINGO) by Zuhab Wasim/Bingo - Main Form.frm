VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bingo"
   ClientHeight    =   7875
   ClientLeft      =   1905
   ClientTop       =   2280
   ClientWidth     =   9690
   FillColor       =   &H80000000&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Bingo - Main Form.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   9690
   Begin VB.Timer tmrTimer 
      Interval        =   1000
      Left            =   4920
      Top             =   1080
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Board 2"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Index           =   1
      Left            =   5520
      TabIndex        =   105
      Top             =   120
      Width           =   3975
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   38
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   50
         Left            =   3120
         TabIndex        =   130
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   49
         Left            =   2400
         TabIndex        =   129
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   48
         Left            =   1680
         TabIndex        =   128
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   47
         Left            =   960
         TabIndex        =   127
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   46
         Left            =   240
         TabIndex        =   126
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   45
         Left            =   3120
         TabIndex        =   125
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   44
         Left            =   2400
         TabIndex        =   124
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   43
         Left            =   1680
         TabIndex        =   123
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   42
         Left            =   960
         TabIndex        =   122
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   41
         Left            =   240
         TabIndex        =   121
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   40
         Left            =   3120
         TabIndex        =   120
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   39
         Left            =   2400
         TabIndex        =   119
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   38
         Left            =   1680
         TabIndex        =   118
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   37
         Left            =   960
         TabIndex        =   117
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   36
         Left            =   240
         TabIndex        =   116
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   35
         Left            =   3120
         TabIndex        =   115
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   34
         Left            =   2400
         TabIndex        =   114
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   33
         Left            =   1680
         TabIndex        =   113
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   32
         Left            =   960
         TabIndex        =   112
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   31
         Left            =   240
         TabIndex        =   111
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   30
         Left            =   3120
         TabIndex        =   110
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   29
         Left            =   2400
         TabIndex        =   109
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   28
         Left            =   1680
         TabIndex        =   108
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   27
         Left            =   960
         TabIndex        =   107
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   26
         Left            =   240
         TabIndex        =   106
         Top             =   360
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   50
         Left            =   240
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   49
         Left            =   960
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   48
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   47
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   46
         Left            =   3120
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   45
         Left            =   240
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   44
         Left            =   960
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   43
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   42
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   41
         Left            =   3120
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   40
         Left            =   240
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   39
         Left            =   960
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   37
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   36
         Left            =   3120
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   35
         Left            =   240
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   34
         Left            =   960
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   33
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   32
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   31
         Left            =   3120
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   30
         Left            =   240
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   29
         Left            =   960
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   28
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   27
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   26
         Left            =   3120
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   615
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Clicks"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   4200
      TabIndex        =   102
      Top             =   120
      Width           =   1215
      Begin VB.Label lblWinCount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   103
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Board 1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3975
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   25
         Left            =   3120
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   24
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   23
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   22
         Left            =   960
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   21
         Left            =   240
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   20
         Left            =   3120
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   19
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   18
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   17
         Left            =   960
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   16
         Left            =   240
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   15
         Left            =   3120
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   14
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   13
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   12
         Left            =   960
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   11
         Left            =   240
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   10
         Left            =   3120
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   9
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   8
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   7
         Left            =   960
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   6
         Left            =   240
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   5
         Left            =   3120
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   4
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   3
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   2
         Left            =   960
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Image imgBoard 
         Height          =   615
         Index           =   1
         Left            =   240
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   25
         Left            =   3120
         TabIndex        =   26
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   24
         Left            =   2400
         TabIndex        =   25
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   23
         Left            =   1680
         TabIndex        =   24
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   22
         Left            =   960
         TabIndex        =   23
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   21
         Left            =   240
         TabIndex        =   22
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   20
         Left            =   3120
         TabIndex        =   21
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   19
         Left            =   2400
         TabIndex        =   20
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   18
         Left            =   1680
         TabIndex        =   19
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   17
         Left            =   960
         TabIndex        =   18
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   16
         Left            =   240
         TabIndex        =   17
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   15
         Left            =   3120
         TabIndex        =   16
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   14
         Left            =   2400
         TabIndex        =   15
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   13
         Left            =   1680
         TabIndex        =   14
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   12
         Left            =   960
         TabIndex        =   13
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   11
         Left            =   240
         TabIndex        =   12
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   10
         Left            =   3120
         TabIndex        =   11
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   9
         Left            =   2400
         TabIndex        =   10
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   8
         Left            =   1680
         TabIndex        =   9
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   7
         Left            =   960
         TabIndex        =   8
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   6
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   5
         Left            =   3120
         TabIndex        =   6
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   4
         Left            =   2400
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   1680
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   960
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblBoard 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame fraWinNum 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Winning Numbers"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   4200
      Width           =   9495
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   75
         Left            =   8640
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   74
         Left            =   8040
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   73
         Left            =   7440
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   72
         Left            =   6840
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   71
         Left            =   6240
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   70
         Left            =   5640
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   69
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   68
         Left            =   4440
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   67
         Left            =   3840
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   66
         Left            =   3240
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   65
         Left            =   2640
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   64
         Left            =   2040
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   63
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   62
         Left            =   840
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   61
         Left            =   240
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   60
         Left            =   8640
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   59
         Left            =   8040
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   58
         Left            =   7440
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   57
         Left            =   6840
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   56
         Left            =   6240
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   55
         Left            =   5640
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   54
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   53
         Left            =   4440
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   52
         Left            =   3840
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   51
         Left            =   3240
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   50
         Left            =   2640
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   49
         Left            =   2040
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   48
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   47
         Left            =   840
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   46
         Left            =   240
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   45
         Left            =   8640
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   44
         Left            =   8040
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   43
         Left            =   7440
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   42
         Left            =   6840
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   41
         Left            =   6240
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   40
         Left            =   5640
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   39
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   38
         Left            =   4440
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   37
         Left            =   3840
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   36
         Left            =   3240
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   35
         Left            =   2640
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   34
         Left            =   2040
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   33
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   32
         Left            =   840
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   31
         Left            =   240
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   30
         Left            =   8640
         Stretch         =   -1  'True
         Top             =   960
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   29
         Left            =   8040
         Stretch         =   -1  'True
         Top             =   960
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   28
         Left            =   7440
         Stretch         =   -1  'True
         Top             =   960
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   27
         Left            =   6840
         Stretch         =   -1  'True
         Top             =   960
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   26
         Left            =   6240
         Stretch         =   -1  'True
         Top             =   960
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   25
         Left            =   5640
         Stretch         =   -1  'True
         Top             =   960
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   24
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   960
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   23
         Left            =   4440
         Stretch         =   -1  'True
         Top             =   960
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   22
         Left            =   3840
         Stretch         =   -1  'True
         Top             =   960
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   21
         Left            =   3240
         Stretch         =   -1  'True
         Top             =   960
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   20
         Left            =   2640
         Stretch         =   -1  'True
         Top             =   960
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   19
         Left            =   2040
         Stretch         =   -1  'True
         Top             =   960
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   18
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   960
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   17
         Left            =   840
         Stretch         =   -1  'True
         Top             =   960
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   16
         Left            =   240
         Stretch         =   -1  'True
         Top             =   960
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   15
         Left            =   8640
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   14
         Left            =   8040
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   13
         Left            =   7440
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   12
         Left            =   6840
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   11
         Left            =   6240
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   10
         Left            =   5640
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   9
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   8
         Left            =   4440
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   7
         Left            =   3840
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   6
         Left            =   3240
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   5
         Left            =   2640
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   4
         Left            =   2040
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   3
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   2
         Left            =   840
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Image imgWinNum 
         Height          =   615
         Index           =   1
         Left            =   240
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   75
         Left            =   8640
         TabIndex        =   101
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   74
         Left            =   8040
         TabIndex        =   100
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   73
         Left            =   7440
         TabIndex        =   99
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   72
         Left            =   6840
         TabIndex        =   98
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   71
         Left            =   6240
         TabIndex        =   97
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   70
         Left            =   5640
         TabIndex        =   96
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   69
         Left            =   5040
         TabIndex        =   95
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   68
         Left            =   4440
         TabIndex        =   94
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   67
         Left            =   3840
         TabIndex        =   93
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   66
         Left            =   3240
         TabIndex        =   92
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   65
         Left            =   2640
         TabIndex        =   91
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   64
         Left            =   2040
         TabIndex        =   90
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   63
         Left            =   1440
         TabIndex        =   89
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   62
         Left            =   840
         TabIndex        =   88
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   61
         Left            =   240
         TabIndex        =   87
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   60
         Left            =   8640
         TabIndex        =   86
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   59
         Left            =   8040
         TabIndex        =   85
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   58
         Left            =   7440
         TabIndex        =   84
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   57
         Left            =   6840
         TabIndex        =   83
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   56
         Left            =   6240
         TabIndex        =   82
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   55
         Left            =   5640
         TabIndex        =   81
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   54
         Left            =   5040
         TabIndex        =   80
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   53
         Left            =   4440
         TabIndex        =   79
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   52
         Left            =   3840
         TabIndex        =   78
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   51
         Left            =   3240
         TabIndex        =   77
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   50
         Left            =   2640
         TabIndex        =   76
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   49
         Left            =   2040
         TabIndex        =   75
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   48
         Left            =   1440
         TabIndex        =   74
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   47
         Left            =   840
         TabIndex        =   73
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   46
         Left            =   240
         TabIndex        =   72
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   45
         Left            =   8640
         TabIndex        =   71
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   44
         Left            =   8040
         TabIndex        =   70
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   43
         Left            =   7440
         TabIndex        =   69
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   42
         Left            =   6840
         TabIndex        =   68
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   41
         Left            =   6240
         TabIndex        =   67
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   40
         Left            =   5640
         TabIndex        =   66
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   39
         Left            =   5040
         TabIndex        =   65
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   38
         Left            =   4440
         TabIndex        =   64
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   37
         Left            =   3840
         TabIndex        =   63
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   36
         Left            =   3240
         TabIndex        =   62
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   35
         Left            =   2640
         TabIndex        =   61
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   34
         Left            =   2040
         TabIndex        =   60
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   33
         Left            =   1440
         TabIndex        =   59
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   32
         Left            =   840
         TabIndex        =   58
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   31
         Left            =   240
         TabIndex        =   57
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   30
         Left            =   8640
         TabIndex        =   56
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   29
         Left            =   8040
         TabIndex        =   55
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   28
         Left            =   7440
         TabIndex        =   54
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   27
         Left            =   6840
         TabIndex        =   53
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   26
         Left            =   6240
         TabIndex        =   52
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   25
         Left            =   5640
         TabIndex        =   51
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   24
         Left            =   5040
         TabIndex        =   50
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   23
         Left            =   4440
         TabIndex        =   49
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   22
         Left            =   3840
         TabIndex        =   48
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   21
         Left            =   3240
         TabIndex        =   47
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   20
         Left            =   2640
         TabIndex        =   46
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   19
         Left            =   2040
         TabIndex        =   45
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   18
         Left            =   1440
         TabIndex        =   44
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   17
         Left            =   840
         TabIndex        =   43
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   16
         Left            =   240
         TabIndex        =   42
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   15
         Left            =   8640
         TabIndex        =   41
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   14
         Left            =   8040
         TabIndex        =   40
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   13
         Left            =   7440
         TabIndex        =   39
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   12
         Left            =   6840
         TabIndex        =   38
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   11
         Left            =   6240
         TabIndex        =   37
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   10
         Left            =   5640
         TabIndex        =   36
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   9
         Left            =   5040
         TabIndex        =   35
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   8
         Left            =   4440
         TabIndex        =   34
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   7
         Left            =   3840
         TabIndex        =   33
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   6
         Left            =   3240
         TabIndex        =   32
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   5
         Left            =   2640
         TabIndex        =   31
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   4
         Left            =   2040
         TabIndex        =   30
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   1440
         TabIndex        =   29
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   840
         TabIndex        =   28
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Image imgBoard 
      Height          =   615
      Index           =   0
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblBoard 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   4200
      TabIndex        =   104
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgFree 
      Height          =   615
      Left            =   4800
      Picture         =   "Bingo - Main Form.frx":030A
      Stretch         =   -1  'True
      Top             =   2160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgBingoBack 
      Height          =   600
      Left            =   4800
      Picture         =   "Bingo - Main Form.frx":0614
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgWinBack 
      Height          =   600
      Left            =   4200
      Picture         =   "Bingo - Main Form.frx":091E
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNewGame 
         Caption         =   "New Game"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuReveal 
         Caption         =   "Reveal The Cards!"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuAuto 
         Caption         =   "Auto Draw"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuHighScores 
         Caption         =   "Show HighScores"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
         Shortcut        =   ^H
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programmer: Zuhab Wasim 11G
'Date: 19/05/16
'Purpose: To recreate the game BINGO for practicing purposes.

'Note: Some code that will be used for future modifications of this program are commented.
'      Used code will be indented as if the commented code does not exist

'Forces the declarations and assignments of variables
Option Explicit

'Declares the constant used to detect if a highscore name is greater than 30 characters
Const NAME_LEN = 30

'Declares variables, and arrays that are general between both forms
Dim WinClickNum As Integer
Dim MatchClick As Integer
Dim WinNum(1 To NUM_MAX) As Integer
'Declares the time tracker to see which winning number needs to be clicked
Dim TimerCounter As Integer
'Declares arrays that are used for checking the win of board1 and board2
Dim MatchCount(1 To BINGO_MAX) As Integer

'Declares arrays that are used to display the numbers of board1 and board2
Dim BingoNum(1 To BINGO_MAX) As Integer

'Light Blue: QBColor(9)
'Light Green: QBColor(10)
'Light Cyan: QBColor(11)
'Light Magenta: QBColor(13)
'Light Yellow: QBColor(14)

Private Sub Form_Load()
    
    'Checks to see if file exists, creates the file if not
    If Dir$(FNAME) = "" Then
        CreateFile
    End If
    'Initializes the timer and auto checker
    TimerCounter = 0
    
    GetHighScores
    'Randomizes the seed the program uses to generate random numbers
    Randomize
    'Calls on procedures to start the game
    'Initializes the arrays needed to be reset or zero'd
    Initialize WinClickNum, MatchClick, WinNum(), MatchCount()
    'Assigns images to the imageboxes of Board1, Board2, and WinNum
    AssignImages
    'Generates, randomizes, and displays the winning numbers
    GenWinNums WinNum()
    'Generates, randomizes, and displays the bingo numbers Board1 and Board2
    GenBingoNums lblBoard(), BingoNum(), BOARD1_FIRST, BOARD1_LAST
    GenBingoNums lblBoard(), BingoNum(), BOARD2_FIRST, BOARD2_LAST
    'Hides numbers by setting the visible property of all images to true
    HideCards
    
    'Initially sets NewGame and Auto procedures to false
    'Also sets the timer to false
    tmrTimer.Enabled = False
    mnuNewGame.Enabled = False
    mnuAuto.Enabled = False
    
End Sub

Private Sub imgWinNum_Click(Index As Integer)
    
    'Declares local variables needed for procedure
    Dim K As Integer
    Dim HName As String
    
    'Disables the auto menu option so the user cannot click on it after clicking on the winning numbers
    If mnuAuto.Enabled = True Then
        mnuAuto.Enabled = False
    End If
    
    'Checks to see if WinNum cell has already been clicked, will continue with the code if it is not
    If imgWinNum(Index).Visible = True Then
        'Sets the image clicked to invisible and increments the amount of clicks WinClickNum and displays it in a label
        imgWinNum(Index).Visible = False
        WinClickNum = WinClickNum + 1
        lblWinCount.Caption = Str$(WinClickNum)
        'Checks to see if the clicked Win Cell has the same number as one of the Bingo Cells in each Board
        For K = 1 To BINGO_MAX
            'If the number in the Win Cell exists checks to see which board it exists in
            'and changes the colour of each cell in the board and the win num cell as light green
            If BingoNum(K) = WinNum(Index) Then
                lblBoard(K).BackColor = QBColor(10)
                lblWinNum(Index).BackColor = QBColor(10)
                MatchCount(K) = 1
                'Increments the variable measuring how many matches have been made by 1
                MatchClick = MatchClick + 1
            End If
        Next K
    End If
    
    'Checks to see if atleast 4 matches have been made before using
    If MatchClick >= 4 Then
        'If the CheckWin function returns True, proceed to end the game using the EndGame procedure
        If CheckWin(lblBoard(), MatchCount(), BOARD1_FIRST, BOARD1_LAST) Or _
           CheckWin(lblBoard(), MatchCount(), BOARD2_FIRST, BOARD2_LAST) Then
            EndGame WinClickNum
            'If the score achieved by the user is least than the lowest highscore, then update the scores
            If WinClickNum < HighScores(HIGHSCORE_MAX) Then
                'Asks the user for the name they want to have associated with their score
                HName = InputBox$("You have achieved a High Score! Enter your name:", "Highscore")
                'If the user has entered no name, enter "No Name"
                'If the user enters a name longer than NAME_LEN, then cut the name off
                If Trim$(HName) = "" Then
                    HName = "No Name"
                ElseIf Len(HName) > NAME_LEN Then
                    HName = Mid$(HName, 1, NAME_LEN) & "..."
                End If
                'Changes the high score
                ChangeHighScores WinClickNum, HName
            End If
        End If
    End If

End Sub

Private Sub mnuAbout_Click()
    
    'Loads and shows the About form, and ensure the user cannot click off the form
    Load frmAbout
    frmAbout.Show vbModal
    
End Sub

Private Sub mnuAuto_Click()
    
    'Assigns the interval in which each winning number is clicked
    tmrTimer.Enabled = True
    
    'Disables the auto option when the program is clicking automatically
    mnuAuto.Enabled = False
    
    'Disables the ability to click the winning numbers
    fraWinNum.Enabled = False
    
End Sub

Private Sub mnuExit_Click()
    
    'Declares variables used for displaying the exit message
    Dim ExitMsg As String
    Dim ExitType As Integer
    Dim ExitTitle As String
    Dim ExitResponse As Integer
    
    'Assigns the appropriate values into the given variables
    ExitMsg = "Are you sure you want to exit?"
    ExitType = vbYesNo + vbQuestion
    ExitTitle = "Exit"
    
    'Displays the msgbox function for the response
    ExitResponse = MsgBox(ExitMsg, ExitType, ExitTitle)
    
    'If the user clicks the Yes button, it will end the game
    If ExitResponse = vbYes Then
        'The program saves the highscores before exiting
        WriteFile
        End
    End If
    
End Sub

Private Sub mnuHighScores_Click()
    
    'Loads the highscore form
    Load frmScores
    
    'Displays the highscore form
    frmScores.Show vbModal
    
End Sub

Private Sub mnuNewGame_Click()
    
    'Recalls the procedures used in form load that restarts the game
    
    'Hides numbers by setting the visible property of all images to true
    HideCards
    'Resets the colours of each clicked cell that has different backcolor as the default
    ResetColours
    'Initializes the arrays needed to be reset or zero'd
    Initialize WinClickNum, MatchClick, WinNum(), MatchCount()
    'Generates, randomizes, and displays the winning numbers
    GenWinNums WinNum()
    'Generates, randomizes, and displays the bingo numbers
    GenBingoNums lblBoard(), BingoNum(), BOARD1_FIRST, BOARD1_LAST
    GenBingoNums lblBoard(), BingoNum(), BOARD2_FIRST, BOARD2_LAST
    
    'Reinitializes time counter to zero once a new game is made
    TimerCounter = 0
    
    'Disables the NewGame menu option until the Reveal option is clicked
    mnuNewGame.Enabled = False
    'Enables the reveal option
    mnuReveal.Enabled = True
    mnuAuto.Enabled = False
    
End Sub

Private Sub mnuReveal_Click()
    
    'Calls the RevealCards procedure that makes the
    RevealCards
    
    'Enables the NewGame menu option and disables the Reveal option until NewGame option is clicked
    'mnuNewGame.Enabled = True
    mnuReveal.Enabled = False
    mnuAuto.Enabled = True
    
End Sub

Private Sub tmrTimer_Timer()
    
    'Increments timer counter by 1
    TimerCounter = TimerCounter + 1
    
    'Calls the even procedure of imgWinNum_Click with the arguement of timer counter
    imgWinNum_Click (TimerCounter)
    
End Sub
