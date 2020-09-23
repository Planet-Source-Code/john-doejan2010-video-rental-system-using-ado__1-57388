VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRental 
   Caption         =   "Rental"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11595
   Icon            =   "frmRental.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6840
   ScaleWidth      =   11595
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H0000FF00&
      Caption         =   "Print"
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5640
      Width           =   975
   End
   Begin VB.Frame Frame4 
      Caption         =   "Customer"
      Height          =   855
      Left            =   5640
      TabIndex        =   16
      Top             =   1920
      Width           =   2535
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   240
         TabIndex        =   17
         Text            =   "Select CustomerID"
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Video"
      Height          =   855
      Left            =   2760
      TabIndex        =   14
      Top             =   1920
      Width           =   2655
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   360
         TabIndex        =   15
         Text            =   "Select VideoID"
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmNew 
      BackColor       =   &H0000FF00&
      Caption         =   "New"
      Height          =   375
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H0000FF00&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H0000FF00&
      Caption         =   "Update"
      Height          =   375
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5640
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dates"
      Height          =   1815
      Left            =   5640
      TabIndex        =   5
      Top             =   3240
      Width           =   3495
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   1200
         TabIndex        =   8
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   22675457
         CurrentDate     =   38266
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Rental Date"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Due Date"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Transaction Details"
      Height          =   1815
      Left            =   1680
      TabIndex        =   0
      Top             =   3240
      Width           =   3615
      Begin MSMask.MaskEdBox Text1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """L""-00#"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   18
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   393216
         ClipMode        =   1
         PromptChar      =   "_"
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Customer Id"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Video ID"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Transaction ID"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Rental"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   19
      Top             =   600
      Width           =   4695
   End
End
Attribute VB_Name = "frmRental"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag As Boolean
Dim conny As ADODB.Connection
Dim ras As ADODB.Recordset
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim conn2 As ADODB.Connection
Dim rs2 As ADODB.Recordset
Dim sql As String
Dim msg As String
Dim seanPaul As String

Private Sub cmdCancel_Click()
Unload Me
'exit
End Sub

Private Sub cmdPrint_Click()


seanPaul = InputBox("Enter TransID", "Print", "enter valid transaction id", 400, 400)
If StrPtr(seanPaul) = 0 Then
Exit Sub
Else
Call connectDB
sql = "SELECT DISTINCT * FROM Loandetails WHERE TransactionId='" & seanPaul & "'"
ras.Open sql, conny, adOpenStatic, adLockOptimistic
Set TransByCus.DataSource = ras
TransByCus.Show
End If
End Sub

Private Sub cmdUpdate_Click()
Call connectDB2
rs2.Open "Select * from Video where VideoID='" & Text2.Text & "'", conn2

On Error GoTo ras
'check for date,customer,video and trans b4 updating
If (Text1.Text <> "" And Text2.Text <> "") And (Text3.Text <> "" And Text5.Text <> "") Then
 If IsNumeric(Text1.Text) = True Then
  conn2.Execute "Update Video set Status ='0' where VideoID='" & Text2.Text & "'"
  rs.Move 0
  rs.Update
  
  msg = MsgBox("Record Updated...Do you want to Print the receipt?", vbYesNo, "Record Updated")
   Select Case msg
    Case vbYes
     Call cmdPrint_Click
     rs.AddNew
     Text5.Text = DTPicker1.Value + 3
     Exit Sub
    Case vbNo
     
     Unload Me
     Exit Sub
   End Select
   Else
   MsgBox ("Enter numeric values for the Transaction Id in the format 0001"), vbCritical, "Incorrect Format"
   Exit Sub
   End If
Else
MsgBox ("You did not enter all the required fields. Please fill in all the necessary Details"), vbCritical, "Incomplete Record"
Exit Sub
End If

ras:
 MsgBox (Err.Description), vbOKOnly, "Data Entry Error"



End Sub

Private Sub cmNew_Click()
rs.AddNew
Text5.Text = DTPicker1.Value + 3
End Sub

Private Sub Combo1_Click()
'load textbox with selected videoId
If Combo1.Text <> "" Then
 Text2.Text = Combo1.Text
Else
 MsgBox "No Videos In Database"
End If
End Sub

Private Sub Combo2_Click()
'load textbox with selected customerId
If Combo2.Text <> "" Then
 Text3.Text = Combo2.Text
Else
 MsgBox "No Videos In Database"
End If
End Sub

Private Sub Form_Load()
'update date pickre
DTPicker1.Value = Date
'connect
Call connectsDB

' populate video list
Set rs = New ADODB.Recordset
sql = "SELECT VideoID FROM Video"
  
  rs.Open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
  
  While rs.EOF = False
      Combo1.AddItem rs!VideoID
      rs.MoveNext
  Wend
  rs.Close
' populate customre luist
Set rs = New ADODB.Recordset
sql = "SELECT CustomerID FROM Customer"
  
  rs.Open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
  
  While rs.EOF = False
      Combo2.AddItem rs!CustomerID
      rs.MoveNext
  Wend
 rs.Close
rs.Open "LoanDetails", conn, adOpenStatic, adLockOptimistic, adCmdTable
'set the datasourcwes
Set Text1.DataSource = rs
Text1.DataField = "TransactionId"
Set Text2.DataSource = rs
Text2.DataField = "VideoID"
Set Text3.DataSource = rs
Text3.DataField = "CustomerID"
Set Text5.DataSource = rs
Text5.DataField = "DueDate"
rs.AddNew
' load due date
Text5.Text = DTPicker1.Value + 3


End Sub

Function connectDB()
Set conny = New ADODB.Connection
Set ras = New ADODB.Recordset
conny.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source= " & App.Path & "\db.mdb"
End Function

Function connectsDB()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source= " & App.Path & "\db.mdb"
End Function


Function connectDB2()
Set conn2 = New ADODB.Connection
Set rs2 = New ADODB.Recordset
conn2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source= " & App.Path & "\db.mdb"
End Function


