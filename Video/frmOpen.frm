VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOpen 
   Caption         =   " Open..."
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4830
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5318
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tables"
      TabPicture(0)   =   "frmOpen.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "List1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Reports"
      TabPicture(1)   =   "frmOpen.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Command4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Command3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "List2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label4"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label3"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   -73680
         TabIndex        =   12
         Text            =   "Video List"
         Top             =   2355
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   -71280
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "OK"
         Height          =   375
         Left            =   -71280
         TabIndex        =   9
         Top             =   480
         Width           =   975
      End
      Begin VB.ListBox List2 
         Height          =   1425
         ItemData        =   "frmOpen.frx":0038
         Left            =   -74880
         List            =   "frmOpen.frx":0045
         TabIndex        =   8
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Text            =   "Video"
         Top             =   2400
         Width           =   2295
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   3720
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Height          =   375
         Left            =   3720
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
      Begin VB.ListBox List1 
         Height          =   1425
         ItemData        =   "frmOpen.frx":0072
         Left            =   120
         List            =   "frmOpen.frx":007F
         TabIndex        =   2
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label Label4 
         Caption         =   "Report Name"
         Height          =   255
         Left            =   -74880
         TabIndex        =   11
         Top             =   2370
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Select Report"
         Height          =   255
         Left            =   -74760
         TabIndex        =   7
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Table Name"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Select Table"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   510
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag As Boolean
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset

Function connectDB()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source= " & App.Path & "\db.mdb"
End Function


Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Select Case Text2.Text
 Case "Video List"
  Unload Me
  Call connectDB
  rs.Open "SELECT * FROM Video", conn, adOpenStatic, adLockOptimistic
  Set VideoList.DataSource = rs
  VideoList.Show
 Case "Customer List"
  Unload Me
  Call connectDB
  rs.Open "SELECT * FROM Customer", conn, adOpenStatic, adLockOptimistic
  Set CustomerList.DataSource = rs
  CustomerList.Show
 Case "Rentals List"
  Unload Me
  Call connectDB
  rs.Open "SELECT * FROM LoanDetails", conn, adOpenStatic, adLockOptimistic
  Set RentalsList.DataSource = rs
  RentalsList.Show

 Case Else
  MsgBox ("Invalid rEPORT name"), vbInformation, "File not found"
End Select
End Sub
Private Sub Command1_Click()
Select Case Text1.Text
 Case "Video"
  Unload Me
  Call connectDB
  rs.Open "SELECT * FROM Video", conn, adOpenStatic, adLockOptimistic
  frmListView.Show
 Case "Customers"
  Unload Me
  Call connectDB
  rs.Open "SELECT * FROM Customer", conn, adOpenStatic, adLockOptimistic
  frmListCustomers.Show
 Case "Rentals"
  Unload Me
  Call connectDB
  rs.Open "SELECT * FROM LoanDetails", conn, adOpenStatic, adLockOptimistic
  frmListRentals.Show

 Case Else
  MsgBox ("Invalid Table name"), vbInformation, "File not found"
End Select
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub List1_Click()
 Text1.Text = List1.Text
End Sub

Private Sub List1_DblClick()
 Text1.Text = List1.Text
 Call Command1_Click
End Sub

Private Sub List2_Click()

Text2.Text = List2.Text

End Sub

Private Sub List2_DblClick()
Text2.Text = List2.Text
Call Command3_Click
End Sub
