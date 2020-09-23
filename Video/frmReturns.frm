VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmReturns 
   BackColor       =   &H000000FF&
   Caption         =   "Returns"
   ClientHeight    =   2730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5865
   Icon            =   "frmReturns.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2730
   ScaleWidth      =   5865
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   4440
      TabIndex        =   22
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   4320
      TabIndex        =   21
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "Calculate Fine"
      Height          =   375
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   6840
      TabIndex        =   16
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   6840
      TabIndex        =   15
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   6840
      TabIndex        =   14
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0000FF00&
      Caption         =   "Search"
      Height          =   375
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   360
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H000000FF&
      Height          =   735
      Left            =   600
      TabIndex        =   10
      Top             =   120
      Width           =   3615
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1440
         TabIndex        =   12
         Text            =   "Select Transaction ID"
         Top             =   330
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H000000FF&
         Caption         =   "Transaction ID"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      Format          =   22609921
      CurrentDate     =   38281
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FF00&
      Caption         =   "View Details   >>"
      Height          =   375
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   840
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000000FF&
      Caption         =   "Transaction Status"
      Height          =   1695
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   3615
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   3480
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   -720
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         CalendarBackColor=   255
         CustomFormat    =   "dd/mm/yy"
         Format          =   22609921
         CurrentDate     =   38281
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000FF&
         Caption         =   "Toady's Date :"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H000000FF&
         Caption         =   "Status :"
         Height          =   375
         Left            =   600
         TabIndex        =   5
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H000000FF&
         Caption         =   "Select Transaction Id"
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Top             =   1080
         Width           =   2295
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "OK"
      Height          =   375
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Date"
      Height          =   255
      Left            =   5880
      TabIndex        =   19
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Video"
      Height          =   255
      Left            =   5880
      TabIndex        =   18
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Customer"
      Height          =   255
      Left            =   5880
      TabIndex        =   17
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "frmReturns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag As Boolean
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Private cn As ADODB.Connection 'for adodb connection
Private ro As ADODB.Recordset 'for adodb recordset
Dim sql As String 'for sql queries
Dim msg As String
Dim ras As Boolean
Dim overdue As Integer


Private Sub Command1_Click()
Dim fine As String
fine = InputBox("Enter The Current Fine", "Fine")

Text6.Text = DTPicker1.Value
If StrPtr(fine) = 0 Then
 Exit Sub
Else
 If overdue < 30 Then
  msg = MsgBox("Fine is $" & Val(fine) * overdue, vbInformation, "Fine")
Else
  msg = MsgBox("Over Due period has exceeded Limit", vbInformation, "Fine")
 End If
End If
End Sub

Private Sub Command2_Click()
Call connectDB
rs.Open "video", cn, adOpenStatic, adLockOptimistic, adCmdTable
conn.Execute "Update Video set Status ='1' where VideoID='" & Text4.Text & "'"

End Sub

Private Sub Command3_Click()
Call connectDB
If Combo1.Text <> "Select Transaction ID" Then
sql = "SELECT DISTINCT * FROM LoanDetails WHERE TransactionID='" & Combo1.Text & "'"
rs.Open sql, cn, adOpenStatic, adLockOptimistic, adCmdText

Set Text3.DataSource = rs
Text3.DataField = "CustomerID"
Set Text4.DataSource = rs
Text4.DataField = "VideoID"
Set Text5.DataSource = rs
Text5.DataField = "DueDate"

Select Case Command3.Caption
Case "View Details   >>"
frmReturns.Width = 8595
Command3.Caption = "Hide Details <<"
Case "Hide Details <<"
frmReturns.Width = 5820
Command3.Caption = "View Details   >>"
End Select
End If
End Sub

Private Sub Command4_Click()
frmSearch.Show
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub


Function connectDB()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source= " & App.Path & "\db.mdb"
End Function


Private Sub Combo1_Click()
Call connectDB
rs.Open "SELECT DISTINCT * FROM LoanDetails WHERE TransactionID='" & Combo1.Text & "'", conn, adOpenStatic, adLockOptimistic, adCmdText
Set Text2.DataSource = rs
Text2.DataField = "DueDate"
DTPicker2.Value = Text2.Text
Text6.Text = DTPicker1.Value
If DTPicker2.Value < DTPicker1.Value Then
overdue = CDate(Text6.Text) - CDate(Text2.Text)
Label3.ForeColor = vbWhite
Label3.Caption = "Over due by :  " & overdue & "  days"
Else

Label3.ForeColor = vbWhite
Label3.Caption = "Due on " & Text2.Text
End If
'Select Case Check1.Value
 'Case 1
  'msg = MsgBox("Video Available....PRoceed?", vbYesNo, "Check Video Status")
   'Select Case msg
    'Case vbYes
     'frmRental.Show

     'conn.Execute "Update Video set Status ='0' where VideoID='" & Combo1.Text & "'"
     'rs.Requery
   'End Select
 'Case 0
'  msg = MsgBox("Video is Currently rented out", vbOKOnly, "Check Video Status")
  'Check1.Enabled = True
 'Case Else
  'MsgBox "Database error"
'End Select
'rs.Close
End Sub


Private Sub Form_Load()
Set cn = New ADODB.Connection
cn.Provider = "Microsoft.Jet.OLEDB.4.0"
cn.Open App.Path & "./db.mdb"

Set ro = New ADODB.Recordset
sql = "SELECT TransactionID FROM LoanDetails"
  
  ro.Open sql, cn, adOpenStatic, adLockOptimistic, adCmdText
  
  While ro.EOF = False
      Combo1.AddItem ro!TransactionID
      ro.MoveNext
  Wend
  ro.Close
  DTPicker1.Value = Date
frmReturns.Width = 5820
Command3.Caption = "View Details   >>"
End Sub

