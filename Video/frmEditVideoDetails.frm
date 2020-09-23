VERSION 5.00
Begin VB.Form frmEditVideoDetails 
   Caption         =   "Video Details"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   9105
   Icon            =   "frmEditVideoDetails.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6660
   ScaleWidth      =   9105
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H00FFFFFF&
      Caption         =   "New"
      Height          =   375
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H8000000E&
      Caption         =   "Update"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H8000000E&
      Caption         =   "Exit"
      Height          =   375
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H8000000E&
      Caption         =   "Edit"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Edit Current Record"
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdDel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Delete"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Delete Current Record"
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdLast 
      BackColor       =   &H0000FF00&
      Caption         =   ">>|"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Last Record"
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton cmdFirst 
      BackColor       =   &H0000FF00&
      Caption         =   "|<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "First Record"
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdFoward 
      BackColor       =   &H0000FF00&
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Next Record"
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdPrev 
      BackColor       =   &H0000FF00&
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "PreviousRecord"
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Movie Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   4080
      TabIndex        =   4
      Top             =   720
      Width           =   3495
      Begin VB.TextBox Text9 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "mmm-yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   16
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox Text8 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   15
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox Text7 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   14
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   13
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Title"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Release Year"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Actress"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Actor"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Video Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3495
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "V###-##"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         DataSource      =   "vid"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Condition"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Date In"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Video Id"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Label Label11 
      Caption         =   "Label11"
      Height          =   255
      Left            =   4320
      TabIndex        =   26
      Top             =   6120
      Width           =   2655
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Caption         =   "Video Details"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   615
      Left            =   1080
      TabIndex        =   9
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "frmEditVideoDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private vin As video
Dim msg As String

Private Sub cmdDel_Click()
vin.funcDel
End Sub


Private Sub Text1_LostFocus()
If Not IsNumeric(Text1.Text) Then
msg = MsgBox("The Account Number Can Only Be a Numeric Vealue", vbCritical, "Data Entry Error")
Text1.SetFocus
End If

End Sub

Private Sub Text3_LostFocus()
If Not IsDate(Text3.Text) Then
msg = MsgBox("Enter a Valid Date", vbCritical, "Data Entry Error")
Text3.SetFocus
End If
End Sub

Private Sub Text8_LostFocus()
If Not IsDate(Text3.Text) Then
msg = MsgBox("Enter a Valid Date", vbCritical, "Data Entry Error")
Text3.SetFocus
End If
End Sub

Private Sub Text9_LostFocus()
If Not IsDate(Text9.Text) Then
msg = MsgBox("Enter a Valid Date", vbCritical, "Data Entry Error")
Text9.SetFocus
End If
End Sub





Private Sub cmdEdit_Click()
Text1.Enabled = True

Text3.Enabled = True

Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
cmdUpdate.Enabled = True
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFirst_Click()
vin.funcFirst
End Sub

Private Sub cmdFoward_Click()
vin.funcNext
End Sub

Private Sub cmdLast_Click()
vin.funcLast
End Sub

Private Sub cmdNew_Click()
Text1.Enabled = True



Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
vin.funcNew
Label11.Caption = vin.total2
End Sub

Private Sub cmdPrev_Click()
vin.funcPrev
End Sub

Private Sub cmdUpdate_Click()
vin.funcUpdate
Label11.Caption = vin.total2
MsgBox ("The Record has been saved!"), vbInformation
Text1.Enabled = False

Text3.Enabled = False

Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
End Sub


Private Sub Form_Load()

Set vin = New video
Set Text1.DataSource = vin
Text1.DataField = "VideoID"

Set Text3.DataSource = vin
Text3.DataField = "DateIN"

Set Text5.DataSource = vin
Text5.DataField = "Condition"
Set Text6.DataSource = vin
Text6.DataField = "Title"
Set Text7.DataSource = vin
Text7.DataField = "Actor"
Set Text8.DataSource = vin
Text8.DataField = "Actress"
Set Text9.DataSource = vin
Text9.DataField = "ReleaseYear"

vin.funcCount
cmdNew.Enabled = True
End Sub


