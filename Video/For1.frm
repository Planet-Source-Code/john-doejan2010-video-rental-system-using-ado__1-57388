VERSION 5.00
Begin VB.Form frmCheck 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check if Tape is Available"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3345
   Icon            =   "For1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   3345
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Check"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Text            =   "---------Select Video ID----------"
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "frmcheck"
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


Function connectDB()
Set rs = New ADODB.Recordset
Set conn = New ADODB.Connection
conn.Provider = "Microsoft.Jet.OLEDB.4.0"
conn.Open App.Path & "./db.mdb"
End Function
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Call connectDB

rs.Open "SELECT DISTINCT * FROM Video WHERE VideoID ='" & Combo1.Text & "'", conn, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = rs
Text1.DataField = "Status"

Select Case Text1.Text
 Case 1
  msg = MsgBox("Tape Available" & vbCrLf & "Proceed to make a rental?", vbYesNo, "Availability")
   Select Case msg
     Case vbYes
      frmRental.Show
      Unload Me
      Exit Sub
     Case vbNo
      Unload Me
      Exit Sub
   End Select
 Case 0
 msg = MsgBox("Tape is not Available" & vbCrLf & "Do you want to Resserve?", vbYesNo, "Availability")
   Select Case msg
     Case vbYes
      frmReserve.Show
      Unload Me
      Exit Sub
     Case vbNo
      Unload Me
      Exit Sub
   End Select
 Case Else
 MsgBox "Select Video"
 End Select
End Sub


Private Sub Form_Load()
Set cn = New ADODB.Connection
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source= " & App.Path & "\db.mdb"
Set ro = New ADODB.Recordset
sql = "SELECT VideoID FROM Video"
    ro.Open sql, cn, adOpenStatic, adLockOptimistic, adCmdText
    While ro.EOF = False
      Combo1.AddItem ro!VideoID
      ro.MoveNext
  Wend
  ro.Close
End Sub


