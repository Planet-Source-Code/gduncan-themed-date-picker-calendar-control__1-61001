VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "MDI test form"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7530
   BeginProperty Font 
      Name            =   "Script MT Bold"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   7530
   Begin DATEPICKER.Duncan_DatePicker Duncan_DatePicker3 
      Height          =   300
      Left            =   1680
      TabIndex        =   4
      Top             =   1920
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   529
      FirstDayOfWeek  =   1
      ShortDayNames   =   -1  'True
      DescriptionFormat=   "d mmmm yyyy"
      DateSelected    =   38505
      ShowNonMonthDays=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Script MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "B Button"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "A Button"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   1080
      Width           =   855
   End
   Begin DATEPICKER.Duncan_DatePicker Duncan_DatePicker2 
      Height          =   300
      Left            =   3240
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      FirstDayOfWeek  =   1
      DescriptionFormat=   "d mmm yyyy"
      UseHandCursor   =   0   'False
      ShowNonMonthDays=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DATEPICKER.Duncan_DatePicker Duncan_DatePicker1 
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   529
      FirstDayOfWeek  =   1
      DescriptionFormat=   "d mmm yyyy"
      UseHandCursor   =   0   'False
      DateSelected    =   38504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

