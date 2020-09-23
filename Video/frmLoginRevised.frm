VERSION 5.00
Begin VB.Form Login 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2340
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   6810
   LinkTopic       =   "Form2"
   ScaleHeight     =   2340
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   2520
      TabIndex        =   13
      Text            =   "Text6"
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   2640
      TabIndex        =   12
      Text            =   "Text6"
      Top             =   2880
      Width           =   1935
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   2640
      TabIndex        =   11
      Text            =   "Text6"
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "Login"
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   330
      Left            =   2160
      TabIndex        =   9
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   480
      TabIndex        =   8
      Text            =   "Text5"
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   480
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   3000
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   480
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "Admin Login"
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   ":"
      TabIndex        =   0
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   0
      Width           =   6375
   End
   Begin VB.Image Image1 
      Height          =   555
      Left            =   120
      Picture         =   "frmLoginRevised.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   675
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "Password"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "User Name"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn As ADODB.Connection
Dim conn As ADODB.Connection
Dim rs1 As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim sql As String
Dim msg As String

Function connectDB()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source= " & App.Path & "\db.mdb"
End Function

Function connectsDB()
Set cnn = New ADODB.Connection
Set rs1 = New ADODB.Recordset
cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source= " & App.Path & "\db.mdb"
End Function


Private Sub Command1_Click()
Call connectDB
sql = "SELECT DISTINCT * FROM Admin WHERE Username='" & Text1.Text & "'"
rs.Open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
Set Text3.DataSource = rs
Text3.DataField = "AdminID"
Set Text4.DataSource = rs
Text4.DataField = "UserName"
Set Text5.DataSource = rs
Text5.DataField = "UserPassword"
rs.Close
If Text5.Text = "" Then
msg = MsgBox("You are not an Admin User", vbCritical, "Access Denied")
Exit Sub
End If
 
If Text2.Text = Text5.Text Then
 aMAIN.Show
 aMAIN.mnuUserSetup.Enabled = True

 Unload Me
Else
 msg = MsgBox("Enter the correct Username and Password", vbCritical, "Access Denied")
End If
End Sub


Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
Call connectsDB
sql = "SELECT DISTINCT * FROM users WHERE Username='" & Text1.Text & "'"
rs1.Open sql, cnn, adOpenStatic, adLockOptimistic, adCmdText

Set Text7.DataSource = rs1
Text7.DataField = "UserName"
Set Text8.DataSource = rs1
Text8.DataField = "Password"
rs1.Close
If Text8.Text = "" Then
msg = MsgBox("You are not a registerd User", vbCritical, "Access Denied")
Exit Sub
End If
 
If Text2.Text = Text8.Text Then
 aMAIN.Show
 aMAIN.mnuUserSetup.Enabled = False

 Unload Me
Else
 msg = MsgBox("Enter the correct Username and Password", vbCritical, "Access Denied")
End If
End Sub
