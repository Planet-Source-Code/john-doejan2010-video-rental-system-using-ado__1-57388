VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Splash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "spash screen"
   ClientHeight    =   3825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "splash.frx":0000
   ScaleHeight     =   3825
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   480
      Top             =   2760
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   3000
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      MouseIcon       =   "splash.frx":0D2F
      Scrolling       =   1
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Video Rental System"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   3600
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ver 2.1"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   2
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Caption         =   "Licence: FREEWARE"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   1
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      Height          =   255
      Left            =   0
      Top             =   3600
      Width           =   5775
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   5760
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00004000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00004000&
      BorderStyle     =   0  'Transparent
      Height          =   3855
      Left            =   5400
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "il Duce"
      BeginProperty Font 
         Name            =   "Rage Italic"
         Size            =   50.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2280
      TabIndex        =   0
      Top             =   1080
      Width           =   2775
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Label2.Caption = "0 %"
End Sub



Private Sub Timer2_Timer()
ProgressBar1.Value = ProgressBar1.Value + 1
Label2.Caption = ProgressBar1.Value & "%"
 If ProgressBar1.Value = 100 Then
 Label2.Caption = "Please Wait.."
  Unload Me
 Login.Show
End If
End Sub
