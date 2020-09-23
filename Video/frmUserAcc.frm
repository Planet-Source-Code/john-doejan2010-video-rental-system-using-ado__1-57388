VERSION 5.00
Begin VB.Form frmUserAcc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Users"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6780
   Icon            =   "frmUserAcc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   6780
   Begin VB.Frame Frame6 
      Caption         =   "Password"
      Height          =   735
      Left            =   3360
      TabIndex        =   11
      Top             =   1560
      Width           =   3015
      Begin VB.TextBox txtPassword 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "User Name"
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   3015
      Begin VB.TextBox txtUserName 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Users"
      Height          =   735
      Left            =   1800
      TabIndex        =   7
      Top             =   840
      Width           =   3015
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   1680
      TabIndex        =   2
      Top             =   2520
      Width           =   3135
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Add"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   375
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame5 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   6255
      Begin VB.CommandButton Command5 
         Caption         =   "Exit"
         Height          =   495
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Admin Area"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   13
      Top             =   240
      Width           =   4935
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      BorderWidth     =   10
      Height          =   4575
      Left            =   0
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmUserAcc"
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

Set rs = New ADODB.Recordset
Set cn = New ADODB.Connection
cn.Provider = "Microsoft.Jet.OLEDB.4.0"
cn.Open App.Path & "./db.mdb"
End Function

Private Sub cmdCancel_Click()
Call connectDB

rs.Open "Users", conn, adOpenStatic, adLockOptimistic, adCmdTable

Set txtUserName.DataSource = rs
txtUserName.DataField = "Username"
Set txtPassword.DataSource = rs
txtPassword.DataField = "password"
rs.AddNew

End Sub



Private Sub cmdDelete_Click()

msg = MsgBox("Are you sure you want to delete the user '" & txtUserName & "'?", vbYesNo, "Delete User")
Call connectDB
sql = "Select from Users where username='" & txtUserName & "'"

Select Case msg
 Case vbYes
  conn.Execute ("delete from users where username='" & txtUserName & "'")

  txtPassword.Text = ""
  txtUserName.Text = ""
  msg = MsgBox("'" & txtUserName & "' Deleted", vbInformation, "Delete User")
  Case vbNo
  Exit Sub
End Select
End Sub



Private Sub cmdSave_Click()
msg = MsgBox("Are you sure you want to add the user '" & txtUserName & "'?", vbYesNo, "Add new User")

rs.Update
msg = MsgBox("New User '" & txtUserName & "'created", vbInformation, "Add new User")
End Sub


Private Sub Combo1_Click()
Call connectDB
sql = "SELECT DISTINCT * FROM Users WHERE Username='" & Combo1.Text & "'"
rs.Open sql, conn, adOpenStatic, adLockOptimistic, adCmdText

Set txtUserName.DataSource = rs
txtUserName.DataField = "Username"
Set txtPassword.DataSource = rs
txtPassword.DataField = "password"
End Sub




Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set cn = New ADODB.Connection
Set ro = New ADODB.Recordset
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source= " & App.Path & "\db.mdb"
Set ro = New ADODB.Recordset
sql = "SELECT Username FROM Users"
    ro.Open sql, cn, adOpenStatic, adLockOptimistic, adCmdText
    While ro.EOF = False
      Combo1.AddItem ro!UserName
      ro.MoveNext
  Wend
  ro.Close
End Sub
