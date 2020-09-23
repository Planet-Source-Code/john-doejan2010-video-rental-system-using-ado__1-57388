VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6930
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   6930
   Begin TabDlg.SSTab SSTab1 
      Height          =   3975
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7011
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Video"
      TabPicture(0)   =   "frmSearch.frx":014A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "List1"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(2)=   "Combo1"
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(4)=   "Label1"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Customer"
      TabPicture(1)   =   "frmSearch.frx":0166
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "List2"
      Tab(1).Control(2)=   "Combo2"
      Tab(1).Control(3)=   "Label13"
      Tab(1).Control(4)=   "Label12"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Transaction"
      TabPicture(2)   =   "frmSearch.frx":0182
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label14"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label15"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "List3"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Combo3"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   240
         TabIndex        =   38
         Text            =   "Select Search Criteria"
         Top             =   840
         Width           =   1695
      End
      Begin VB.ListBox List3 
         Height          =   2205
         ItemData        =   "frmSearch.frx":019E
         Left            =   240
         List            =   "frmSearch.frx":01A0
         TabIndex        =   37
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Frame Frame3 
         Caption         =   "Details"
         Height          =   3015
         Left            =   2040
         TabIndex        =   27
         Top             =   720
         Width           =   4095
         Begin VB.CommandButton Command6 
            Caption         =   ">>"
            Height          =   375
            Left            =   2880
            TabIndex        =   33
            Top             =   2160
            Width           =   735
         End
         Begin VB.CommandButton Command5 
            Caption         =   "<<"
            Height          =   375
            Left            =   1800
            TabIndex        =   32
            Top             =   2160
            Width           =   855
         End
         Begin VB.TextBox Text11 
            Height          =   285
            Left            =   1680
            TabIndex        =   31
            Top             =   1080
            Width           =   1935
         End
         Begin VB.TextBox Text10 
            Height          =   285
            Left            =   1680
            TabIndex        =   30
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox Text9 
            Height          =   285
            Left            =   1680
            TabIndex        =   29
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   720
            TabIndex        =   28
            Text            =   "Text7"
            Top             =   1440
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label11 
            Caption         =   "CustomerID"
            Height          =   255
            Left            =   360
            TabIndex        =   36
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label10 
            Caption         =   "Video ID"
            Height          =   375
            Left            =   360
            TabIndex        =   35
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label9 
            Caption         =   "Transaction ID"
            Height          =   375
            Left            =   360
            TabIndex        =   34
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Details"
         Height          =   3015
         Left            =   -72960
         TabIndex        =   17
         Top             =   600
         Width           =   4095
         Begin VB.TextBox Text7 
            Height          =   375
            Left            =   720
            TabIndex        =   26
            Text            =   "Text7"
            Top             =   1440
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   1680
            TabIndex        =   22
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   1680
            TabIndex        =   21
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   1680
            TabIndex        =   20
            Top             =   1080
            Width           =   1935
         End
         Begin VB.CommandButton Command4 
            Caption         =   "<<"
            Height          =   375
            Left            =   1800
            TabIndex        =   19
            Top             =   2160
            Width           =   855
         End
         Begin VB.CommandButton Command3 
            Caption         =   ">>"
            Height          =   375
            Left            =   2880
            TabIndex        =   18
            Top             =   2160
            Width           =   735
         End
         Begin VB.Label Label8 
            Caption         =   "Customer ID"
            Height          =   375
            Left            =   360
            TabIndex        =   25
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label7 
            Caption         =   "Name"
            Height          =   375
            Left            =   360
            TabIndex        =   24
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Surname"
            Height          =   255
            Left            =   360
            TabIndex        =   23
            Top             =   1080
            Width           =   1095
         End
      End
      Begin VB.ListBox List2 
         Height          =   2205
         ItemData        =   "frmSearch.frx":01A2
         Left            =   -74760
         List            =   "frmSearch.frx":01A4
         TabIndex        =   16
         Top             =   1440
         Width           =   1695
      End
      Begin VB.ListBox List1 
         Height          =   2205
         ItemData        =   "frmSearch.frx":01A6
         Left            =   -74760
         List            =   "frmSearch.frx":01A8
         TabIndex        =   15
         Top             =   1440
         Width           =   1695
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   -74760
         TabIndex        =   14
         Text            =   "Select Search Criteria"
         Top             =   720
         Width           =   1695
      End
      Begin VB.Frame Frame1 
         Caption         =   "Details"
         Height          =   3015
         Left            =   -72960
         TabIndex        =   5
         Top             =   600
         Width           =   4095
         Begin VB.CommandButton Command2 
            Caption         =   ">>"
            Height          =   375
            Left            =   2760
            TabIndex        =   13
            Top             =   2280
            Width           =   735
         End
         Begin VB.CommandButton Command1 
            Caption         =   "<<"
            Height          =   375
            Left            =   1680
            TabIndex        =   12
            Top             =   2280
            Width           =   855
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   1560
            TabIndex        =   11
            Top             =   1200
            Width           =   1935
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   1560
            TabIndex        =   9
            Top             =   840
            Width           =   1935
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   1560
            TabIndex        =   7
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label5 
            Caption         =   "Actress"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Actor"
            Height          =   375
            Left            =   240
            TabIndex        =   8
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Title"
            Height          =   375
            Left            =   240
            TabIndex        =   6
            Top             =   480
            Width           =   1455
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   -74760
         TabIndex        =   2
         Text            =   "Select Criteria"
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label15 
         Caption         =   "Entries Available :"
         Height          =   375
         Left            =   240
         TabIndex        =   42
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "Search by :"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label13 
         Caption         =   "Entries Available :"
         Height          =   375
         Left            =   -74880
         TabIndex        =   40
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   "Search by :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   39
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Search by :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   4
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Entries Available :"
         Height          =   375
         Left            =   -74760
         TabIndex        =   3
         Top             =   1200
         Width           =   1575
      End
   End
   Begin VB.TextBox txtA 
      Height          =   375
      Left            =   6960
      TabIndex        =   0
      Text            =   "Text4"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cn As ADODB.Connection 'for adodb connection
Private rs As ADODB.Recordset 'for adodb recordset
Dim sql As String 'for sql queries



Private Sub Combo1_Click()
On Error Resume Next

rs.Close

If Combo1.Text = "VideoID" Then
  List1.Clear
  sql = "SELECT VideoID FROM Video"
  
  rs.Open sql, cn, adOpenStatic, adLockOptimistic, adCmdText
  
  While rs.EOF = False
      List1.AddItem rs!VideoID
      rs.MoveNext
  Wend
  rs.Close

ElseIf Combo1.Text = "Actor" Then
  List1.Clear
  sql = "SELECT DISTINCT Actor from Video"
  
  rs.Open sql, cn, adOpenStatic, adLockOptimistic, adCmdText
   
   While rs.EOF = False
    List1.AddItem rs!Actor
    rs.MoveNext
    
   Wend
   rs.Close
ElseIf Combo1.Text = "Actress" Then
  List1.Clear
  sql = "SELECT DISTINCT Actress from Video"
  
  rs.Open sql, cn, adOpenStatic, adLockOptimistic, adCmdText
   
   While rs.EOF = False
    List1.AddItem rs!Actress
    rs.MoveNext
    
   Wend
   rs.Close
End If
  
List1.Enabled = True
End Sub

Private Sub Command1_Click()
On Error Resume Next
rs.MoveNext
If rs.EOF = True Then
 rs.MoveLast
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
rs.MovePrevious
If rs.BOF = True Then
 rs.MoveFirst
End If
End Sub

Private Sub Form_Load()

Combo1.AddItem "VideoID"
Combo1.AddItem "Actor"
Combo1.AddItem "Actress"

Combo2.AddItem "CustomerID"
Combo2.AddItem "Surname"
Combo2.AddItem "Name"

Combo3.AddItem "TransactionID"
Combo3.AddItem "VideoID"
Combo3.AddItem "CustomerID"

Set cn = New ADODB.Connection
cn.Provider = "Microsoft.Jet.OLEDB.4.0"
cn.Open App.Path & "./db.mdb"

Set rs = New ADODB.Recordset

End Sub


Private Sub List1_Click()
txtA.Text = List1.List(List1.ListIndex)
sql = "SELECT DISTINCT * FROM Video WHERE " & Combo1.Text & "='" & txtA.Text & "'"
rs.Open sql, cn, adOpenStatic, adLockOptimistic, adCmdText

Set Text1.DataSource = rs
Text1.DataField = "Title"
Set Text2.DataSource = rs
Text2.DataField = "Actor"
Set Text3.DataSource = rs
Text3.DataField = "Actress"
List1.Enabled = False
End Sub
' from here down is for Customer Search
Private Sub Combo2_Click()
On Error Resume Next

rs.Close

If Combo2.Text = "CustomerID" Then
  List2.Clear
  sql = "SELECT CustomerID FROM Customer"
  
  rs.Open sql, cn, adOpenStatic, adLockOptimistic, adCmdText
  
  While rs.EOF = False
      List2.AddItem rs!CustomerID
      rs.MoveNext
  Wend
  rs.Close

ElseIf Combo2.Text = "Name" Then
  List2.Clear
  sql = "SELECT DISTINCT Name from Customer"
  
  rs.Open sql, cn, adOpenStatic, adLockOptimistic, adCmdText
   
   While rs.EOF = False
    List2.AddItem rs!Name
    rs.MoveNext
    
   Wend
   rs.Close
ElseIf Combo2.Text = "Surname" Then
  List2.Clear
  sql = "SELECT DISTINCT Surname from Customer"
  
  rs.Open sql, cn, adOpenStatic, adLockOptimistic, adCmdText
   
   While rs.EOF = False
    List2.AddItem rs!Surname
    rs.MoveNext
    
   Wend
   rs.Close
End If
  
List2.Enabled = True
End Sub

Private Sub Command4_Click()
On Error Resume Next
rs.MoveNext
If rs.EOF = True Then
 rs.MoveLast
End If
End Sub

Private Sub Command3_Click()
On Error Resume Next
rs.MovePrevious
If rs.BOF = True Then
 rs.MoveFirst
End If
End Sub

Private Sub List2_Click()

Text7.Text = List2.List(List2.ListIndex)
sql = "SELECT DISTINCT * FROM Customer WHERE " & Combo2.Text & "='" & Text7.Text & "'"
rs.Open sql, cn, adOpenStatic, adLockOptimistic, adCmdText

Set Text6.DataSource = rs
Text6.DataField = "CustomerID"
Set Text5.DataSource = rs
Text5.DataField = "Name"
Set Text4.DataSource = rs
Text4.DataField = "Surname"
List2.Enabled = False

End Sub


' from here down is for Transaction Search
Private Sub Combo3_Click()
On Error Resume Next

rs.Close

If Combo3.Text = "TransactionID" Then
  List3.Clear
  sql = "SELECT TransactionID FROM LoanDetails"
  
  rs.Open sql, cn, adOpenStatic, adLockOptimistic, adCmdText
  
  While rs.EOF = False
      List3.AddItem rs!TransactionID
      rs.MoveNext
  Wend
  rs.Close

ElseIf Combo3.Text = "VideoID" Then
  List3.Clear
  sql = "SELECT DISTINCT VideoID from LoanDetails"
  
  rs.Open sql, cn, adOpenStatic, adLockOptimistic, adCmdText
   
   While rs.EOF = False
    List3.AddItem rs!VideoID
    rs.MoveNext
    
   Wend
   rs.Close
ElseIf Combo3.Text = "CustomerID" Then
  List3.Clear
  sql = "SELECT DISTINCT CustomerID from LoanDetails"
  
  rs.Open sql, cn, adOpenStatic, adLockOptimistic, adCmdText
   
   While rs.EOF = False
    List3.AddItem rs!CustomerID
    rs.MoveNext
    
   Wend
   rs.Close
End If
  
List3.Enabled = True
End Sub

Private Sub Command5_Click()
On Error Resume Next
rs.MoveNext
If rs.EOF = True Then
 rs.MoveLast
End If
End Sub

Private Sub Command6_Click()
On Error Resume Next
rs.MovePrevious
If rs.BOF = True Then
 rs.MoveFirst
End If
End Sub

Private Sub List3_Click()

Text8.Text = List3.List(List3.ListIndex)
sql = "SELECT DISTINCT * FROM LoanDetails WHERE " & Combo3.Text & "='" & Text8.Text & "'"
rs.Open sql, cn, adOpenStatic, adLockOptimistic, adCmdText

Set Text9.DataSource = rs
Text9.DataField = "TransactionID"
Set Text10.DataSource = rs
Text10.DataField = "VideoID"
Set Text11.DataSource = rs
Text11.DataField = "CustomerID"
List3.Enabled = False

End Sub



