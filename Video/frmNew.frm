VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmNew 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New..."
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4890
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   6165
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Transaction"
      TabPicture(0)   =   "frmNew.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Command1"
      Tab(0).Control(1)=   "Command2"
      Tab(0).Control(2)=   "Frame3"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Customer"
      TabPicture(1)   =   "frmNew.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "cmdCreate"
      Tab(1).Control(2)=   "cmdCancel"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Video"
      TabPicture(2)   =   "frmNew.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmdCreate2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Command6"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin VB.Frame Frame3 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2415
         Left            =   -74880
         TabIndex        =   13
         Top             =   360
         Width           =   4695
         Begin VB.CheckBox Check3 
            BackColor       =   &H80000009&
            Caption         =   "Transaction Details"
            Height          =   375
            Left            =   720
            TabIndex        =   14
            Top             =   480
            Width           =   2535
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000009&
            Caption         =   "Check the box above to enter details of a new Customer"
            Height          =   2415
            Left            =   960
            TabIndex        =   15
            Top             =   960
            Width           =   3015
         End
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   3600
         TabIndex        =   10
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton cmdCreate2 
         Caption         =   "Create"
         Height          =   375
         Left            =   2280
         TabIndex        =   9
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   2415
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   4695
         Begin VB.CheckBox Check2 
            BackColor       =   &H80000009&
            Caption         =   "New Video"
            Height          =   375
            Left            =   720
            TabIndex        =   11
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label4 
            BackColor       =   &H80000009&
            Caption         =   "Enter Details of a new Video in stock"
            Height          =   735
            Left            =   960
            TabIndex        =   12
            Top             =   960
            Width           =   3135
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   -71400
         TabIndex        =   7
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton cmdCreate 
         Caption         =   "Create"
         Height          =   375
         Left            =   -72720
         TabIndex        =   6
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2415
         Left            =   -74880
         TabIndex        =   3
         Top             =   480
         Width           =   4695
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000009&
            Caption         =   "New Customer "
            Height          =   375
            Left            =   720
            TabIndex        =   4
            Top             =   480
            Width           =   2535
         End
         Begin VB.Label Label3 
            BackColor       =   &H80000009&
            Caption         =   "Check the box above to enter details of a new Customer"
            Height          =   2415
            Left            =   960
            TabIndex        =   5
            Top             =   960
            Width           =   3015
         End
      End
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Left            =   -71400
         TabIndex        =   2
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Create"
         Default         =   -1  'True
         Height          =   375
         Left            =   -72720
         TabIndex        =   1
         Top             =   3000
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private vin As video
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdCreate_Click()
If Check1.Value = 1 Then
   Unload Me

   frmEditCustomer.Show
   
Else
   MsgBox ("Select a New Record To Create"), vbOKOnly, "Error"
End If
End Sub

Private Sub cmdCreate2_Click()

If Check2.Value = 1 Then
 
 With frmEditVideoDetails
      .Text1.Enabled = True

      .Text3.Enabled = True

      .Text5.Enabled = True
      .Text6.Enabled = True
      .Text7.Enabled = True
      .Text8.Enabled = True
      .Text9.Enabled = True
      .cmdEdit.Enabled = False
      .cmdNew.Enabled = False
      .cmdDel.Enabled = False
      .cmdUpdate.Enabled = True
      .Text1.Text = ""

      .Text3.Text = ""

      .Text5.Text = ""
      .Text6.Text = ""
      .Text7.Text = ""
      .Text8.Text = ""
      .Text9.Text = ""
 End With

 Unload Me
 frmEditVideoDetails.Show
Else
 MsgBox ("Select a new record to create"), vbOKOnly, "Error"
End If
End Sub

Private Sub Command5_Click()

End Sub

Private Sub Command1_Click()
If Check3.Value = 1 Then
 Unload Me
 frmRental.Show
Else
 MsgBox ("Select anew record to create"), vbOKOnly
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command6_Click()
Unload Me
End Sub
