VERSION 5.00
Begin VB.Form frmEditCustomer 
   Caption         =   "Edit Customer Details"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmEditCustomer.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6630
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "  Data Remote"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   9360
      TabIndex        =   15
      Top             =   1560
      Width           =   2295
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrev 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "Broadway"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1200
         Width           =   615
      End
      Begin VB.CommandButton cmdForward 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Broadway"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1200
         Width           =   615
      End
      Begin VB.CommandButton cmdFirst 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         Caption         =   "|<<"
         BeginProperty Font 
            Name            =   "Broadway"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton cmdLast 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         Caption         =   ">>|"
         BeginProperty Font 
            Name            =   "Broadway"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton cmdEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton cmdDel 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2280
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Account Details"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   4920
      TabIndex        =   11
      Top             =   840
      Width           =   4335
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1800
         TabIndex        =   29
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   1800
         TabIndex        =   27
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   1800
         TabIndex        =   26
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "Customer ID"
         Height          =   255
         Left            =   480
         TabIndex        =   28
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Join Date"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "Expiry Date"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Contacts"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   4095
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1560
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   2520
         Width           =   2295
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   3240
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "FirstName"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Surname"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Home Phone"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "e-mail address"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   3240
         Width           =   1575
      End
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Caption         =   "Customers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   25
      Top             =   360
      Width           =   6135
   End
   Begin VB.Label Label18 
      Caption         =   "Tip : Enter Details and click Update to save your changes"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   24
      Top             =   6240
      Width           =   7215
   End
End
Attribute VB_Name = "frmEditCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cus As customer
Dim msg As String

Private Sub cmdDel_Click()
cus.funcDel
End Sub

Private Sub cmdEdit_Click()
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True

Text8.Enabled = True
Text9.Enabled = True

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFirst_Click()
cus.funcFirst
End Sub



Private Sub cmdForward_Click()
cus.funcNext
End Sub

Private Sub cmdLast_Click()
cus.funcLast
End Sub

Private Sub cmdNew_Click()

cus.funcNew
Text1.SetFocus
End Sub

Private Sub cmdPrev_Click()
cus.funcPrev
End Sub

Private Sub cmdUpdate_Click()
cus.funcUpdate

MsgBox ("The Record has been saved!"), vbInformation

End Sub



Private Sub Form_Load()

Set cus = New customer
Set Text1.DataSource = cus
Text1.DataField = "Name"
Set Text2.DataSource = cus
Text2.DataField = "Surname"
Set Text3.DataSource = cus
Text3.DataField = "Phone"
Set Text4.DataSource = cus
Text4.DataField = "Address"
Set Text5.DataSource = cus
Text5.DataField = "CustomerID"
Set Text6.DataSource = cus
Text6.DataField = "Email"
Set Text8.DataSource = cus
Text8.DataField = "JoinDate"
Set Text9.DataSource = cus
Text9.DataField = "ExpiryDate"
cus.funcCount

End Sub





Private Sub Text5_LostFocus()
If Not IsNumeric(Text5.Text) Then
msg = MsgBox("The Account Number Can Only Be a Numeric Vealue", vbCritical, "Data Entry Error")
Text5.SetFocus
End If

End Sub

Private Sub Text8_LostFocus()
If Not IsDate(Text8.Text) Then
msg = MsgBox("Enter a Valid Date", vbCritical, "Data Entry Error")
Text8.SetFocus
End If
End Sub

Private Sub Text9_LostFocus()
If Not IsDate(Text9.Text) Then
msg = MsgBox("Enter a Valid Date", vbCritical, "Data Entry Error")
Text9.SetFocus
End If
End Sub
