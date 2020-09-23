VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "SDI test form"
   ClientHeight    =   2985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4260
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2985
   ScaleWidth      =   4260
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "Vis"
      Height          =   255
      Left            =   3480
      TabIndex        =   16
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Enabled"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2640
      Width           =   1815
   End
   Begin DATEPICKER.Duncan_DatePicker Duncan_DatePicker1 
      Height          =   300
      Left            =   2040
      TabIndex        =   13
      Top             =   2280
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   529
      FirstDayOfWeek  =   1
      DescriptionFormat=   "d mmm yyyy"
      UseHandCursor   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Show Non Month Days"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Change Format"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Short Day Names"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "FirstDayOfWeek"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Hand Mouse Pointer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   3720
      Top             =   720
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set date to 1980"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label lblEnabled 
      Caption         =   "Enabled?"
      Height          =   255
      Left            =   2040
      TabIndex        =   15
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label lblNMD 
      Caption         =   "Show non month days"
      Height          =   255
      Left            =   2160
      TabIndex        =   12
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblShortdayName 
      Caption         =   "Short Day Names"
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblFDOW 
      Caption         =   "First Day Of Week"
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblCursor 
      Caption         =   "Cursor on?"
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Current Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblDate 
      Caption         =   "Timer will set this to be the current date"
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Duncan_DatePicker1.DateSelected = "1/1/1980"
End Sub

Private Sub Command2_Click()
    Duncan_DatePicker1.UseHandCursor = Not Duncan_DatePicker1.UseHandCursor
End Sub

Private Sub Command3_Click()
     Duncan_DatePicker1.FirstDayOfWeek = Duncan_DatePicker1.FirstDayOfWeek + 1
End Sub

Private Sub Command4_Click()
    Duncan_DatePicker1.ShortDayNames = Not Duncan_DatePicker1.ShortDayNames
End Sub

Private Sub Command5_Click()
    Duncan_DatePicker1.DescriptionFormat = Text1
End Sub

Private Sub Command6_Click()
    Duncan_DatePicker1.ShowNonMonthDays = Not Duncan_DatePicker1.ShowNonMonthDays
End Sub

Private Sub Command7_Click()
    Duncan_DatePicker1.Enabled = Not Duncan_DatePicker1.Enabled
End Sub

Private Sub Command8_Click()
    Duncan_DatePicker1.Visible = Not Duncan_DatePicker1.Visible
End Sub



Private Sub Duncan_DatePicker1_DateChanged(ByVal FromDate As Date, ByVal ToDate As Date)
    Debug.Print "Changed to " & ToDate
End Sub

Private Sub Form_Load()
    Text1 = Duncan_DatePicker1.DescriptionFormat
End Sub

Private Sub Timer1_Timer()
    lblDate = Duncan_DatePicker1.DateSelected
    lblCursor = Duncan_DatePicker1.UseHandCursor
    lblFDOW = Duncan_DatePicker1.FirstDayOfWeek
    lblShortdayName = Duncan_DatePicker1.ShortDayNames
    lblNMD = Duncan_DatePicker1.ShowNonMonthDays
    lblEnabled = Duncan_DatePicker1.Enabled
End Sub
