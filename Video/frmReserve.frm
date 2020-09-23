VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmReserve 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reserve Tape"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   Icon            =   "frmReserve.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      BackColor       =   &H0000FF00&
      Caption         =   "Print"
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "New"
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1920
      TabIndex        =   14
      Top             =   3720
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   13
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   12
      Top             =   2400
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FF00&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "Reserve IT"
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Details"
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   4455
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1800
         TabIndex        =   5
         Text            =   "Select Customer"
         Top             =   630
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Text            =   "Select Video"
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "CustomerID :"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Video Id      :"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   255
      Left            =   1920
      TabIndex        =   0
      Top             =   3360
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   450
      _Version        =   393216
      Format          =   22675457
      CurrentDate     =   38281
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Reserve"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   600
      TabIndex        =   16
      Top             =   240
      Width           =   4575
   End
   Begin VB.Label Label7 
      Caption         =   "Video Title"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Customer Name"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Reservation Id"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Reservation Date"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   3360
      Width           =   1575
   End
End
Attribute VB_Name = "frmReserve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag As Boolean
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim conn2 As ADODB.Connection
Dim rs2 As ADODB.Recordset
Private cn As ADODB.Connection 'for adodb connection
Private ro As ADODB.Recordset 'for adodb recordset
Dim sql As String 'for sql queries
Dim msg As String
Dim ras As Boolean
Dim seanPaul As String


Private Sub Command2_Click()
On Error GoTo robert

 If (Text4.Text <> "") And (Text1.Text <> "") And (Text2.Text <> "") Then
  If IsNumeric(Text4.Text) = True Then
   rs.Move 0
   rs.Update
     msg = MsgBox("The Reservation has been made", vbInformation, "Database Report")
   Exit Sub
  Else
   msg = MsgBox("The reservation Id can only contain Numeric Characters", vbCritical, "Data entry Error")
   Exit Sub
  End If
 Else
  msg = MsgBox("One or more records have a null value" & vbCrLf & " Pease re-enter the necessary Data", vbCritical, "Data entry Error")
  Exit Sub
 End If
robert:
msg = MsgBox("The Following error was encountered" & vbCrLf & Err.Description, vbCritical, "To err is Human")
End Sub

Private Sub Combo2_Click()
Text2.Text = Combo2.Text
End Sub

Private Sub Command1_Click()

Call connectDB
rs.Open "Reservation", conn, adOpenStatic, adLockOptimistic, adCmdTable
Set Text4.DataSource = rs
Text4.DataField = "ReservationID"
Set Text1.DataSource = rs
Text1.DataField = "VideoID"
Set Text2.DataSource = rs
Text2.DataField = "CustomerID"
rs.AddNew

End Sub



Private Sub Command4_Click()

seanPaul = InputBox("Enter The TransactionID", 400, 400)
If StrPtr(seanPaul) = 0 Then
Exit Sub
Else
Call connectDB2
sql = "SELECT distinct * FROM Reservation WHERE ReservationId='" & seanPaul & "'"
rs2.Open sql, conn2, adOpenStatic, adLockOptimistic
Set Reserve.DataSource = rs2
Reserve.Show
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Function connectDB2()
Set conn2 = New ADODB.Connection
Set rs2 = New ADODB.Recordset
conn2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source= " & App.Path & "\db.mdb"
End Function


Function connectDB()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source= " & App.Path & "\db.mdb"
End Function


Private Sub Combo1_Click()
Text1.Text = Combo1.Text
End Sub


Private Sub Form_Load()
Set cn = New ADODB.Connection
cn.Provider = "Microsoft.Jet.OLEDB.4.0"
cn.Open App.Path & "./db.mdb"

Set ro = New ADODB.Recordset
sql = "SELECT Distinct title FROM Video"
  
  ro.Open sql, cn, adOpenStatic, adLockOptimistic, adCmdText
  
  While ro.EOF = False
      Combo1.AddItem ro!Title
      ro.MoveNext
  Wend
  ro.Close
  DTPicker1.Value = Now
sql = "SELECT CustomerID FROM Customer"
  
  ro.Open sql, cn, adOpenStatic, adLockOptimistic, adCmdText
  
  While ro.EOF = False
      Combo2.AddItem ro!CustomerID
      ro.MoveNext
  Wend
  ro.Close
Call Command1_Click

'for updating to Reserve Table


End Sub


