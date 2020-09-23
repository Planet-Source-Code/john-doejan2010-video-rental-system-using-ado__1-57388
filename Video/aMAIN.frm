VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm aMAIN 
   BackColor       =   &H8000000C&
   Caption         =   "IL Duce Rental System"
   ClientHeight    =   5865
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7410
   Icon            =   "aMAIN.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   1320
      Top             =   3000
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   5610
      Width           =   7410
      _ExtentX        =   13070
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   21167
            MinWidth        =   21167
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2400
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1800
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483644
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "aMAIN.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "aMAIN.frx":0464
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "aMAIN.frx":05BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "aMAIN.frx":0718
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "aMAIN.frx":0872
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "aMAIN.frx":0CC4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7410
      _ExtentX        =   13070
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Search"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Users"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Change Password"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Home Page"
            ImageIndex      =   6
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      HelpContextID   =   1
      Begin VB.Menu mnuNew 
         Caption         =   "&New.."
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReports 
         Caption         =   "Reports"
         Begin VB.Menu mnuVids 
            Caption         =   "Video List"
         End
         Begin VB.Menu mnuCus 
            Caption         =   "Customer List"
         End
         Begin VB.Menu mnuTransit 
            Caption         =   "Transactions"
            Begin VB.Menu mnuTransCustomer 
               Caption         =   "By Customer"
            End
            Begin VB.Menu mnuAllTransactions 
               Caption         =   "All Transactions"
            End
         End
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEditCustomer 
         Caption         =   "Customer Details"
      End
      Begin VB.Menu mnuEditVideoDetails 
         Caption         =   "Video Details"
      End
   End
   Begin VB.Menu mnuTrans 
      Caption         =   "Transactions"
      Begin VB.Menu mnuRental 
         Caption         =   "Rental"
      End
      Begin VB.Menu mnuRetuns 
         Caption         =   "Return"
      End
      Begin VB.Menu mnuCheck 
         Caption         =   "Check Availability"
      End
      Begin VB.Menu sepChakuti 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQreserve 
         Caption         =   "Quick Reservation"
      End
      Begin VB.Menu mnuQrental 
         Caption         =   "Quick Rental"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuToolbar 
         Caption         =   "Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuStatus 
         Caption         =   "Status Bar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuUserSetup 
         Caption         =   "User Account Setup"
         Enabled         =   0   'False
         Shortcut        =   {F9}
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "Search.."
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuCascade 
         Caption         =   "Cascade Windows"
      End
      Begin VB.Menu mnuTileH 
         Caption         =   "Tile Horizontaly"
      End
      Begin VB.Menu mnuTileV 
         Caption         =   "Tile Verticaly"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About..."
      End
      Begin VB.Menu mnuVisit 
         Caption         =   "Visit Us"
      End
      Begin VB.Menu sep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "Help"
      End
   End
End
Attribute VB_Name = "aMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag As Boolean
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Dim seanPaul As String
Dim msg As Integer

Private Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)
Public Enum StartWindowState
START_HIDDEN = 0
START_NORMAL = 4
START_MINIMIZED = 2
START_MAXIMIZED = 3
End Enum

Public Function ShellDocument(sDocName As String, _
Optional ByVal Action As String = "Open", _
Optional ByVal Parameters As String = vbNullString, _
Optional ByVal Directory As String = vbNullString, _
Optional ByVal WindowState As StartWindowState) As Boolean
Dim Response
Response = ShellExecute(&O0, Action, sDocName, Parameters, Directory, WindowState)
Select Case Response
Case Is < 33
ShellDocument = False
Case Else
ShellDocument = True
End Select
End Function


Function connectDB()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source= " & App.Path & "\db.mdb"
End Function




Private Sub about_Click()
frmAbout.Show , Me
End Sub

Private Sub newacc_Click()
frmNew.Show
End Sub

Private Sub open_Click()
With CommonDialog1
     .DialogTitle = "Open"
     .ShowOpen
     .CancelError = True
     
End With
End Sub

Private Sub pgsetup_Click()
    On Error Resume Next
    With CommonDialog1
        .DialogTitle = "Page Setup"
        .CancelError = True
        .ShowPrinter
    End With

End Sub

Private Sub search_Click()
frmSearch.Show
End Sub







Private Sub MDIForm_Load()

    App.HelpFile = App.Path & "\help\Robertvideoclubproject.chm"
End Sub


Private Sub mnuAllTransactions_Click()
Call connectDB
rs.Open "SELECT * FROM LoanDetails", conn, adOpenStatic, adLockOptimistic
Set RentalsList.DataSource = rs
RentalsList.Show
End Sub

Private Sub mnuCascade_Click()
aMAIN.Arrange vbCascade
End Sub





Private Sub mnuCheck_Click()
frmcheck.Show
End Sub

Private Sub mnuCus_Click()
Call connectDB
rs.Open "SELECT * FROM Customer", conn, adOpenStatic, adLockOptimistic
Set CustomerList.DataSource = rs
CustomerList.Show

End Sub

Private Sub mnuEditCustomer_Click()
frmEditCustomer.Show
End Sub

Private Sub mnuNew_Click()
frmNew.Show
End Sub

Private Sub mnuOpen_Click()
frmOpen.Show
End Sub

Private Sub mnuQrental_Click()
frmRental.Show
End Sub

Private Sub mnuQreserve_Click()
frmReserve.Show
End Sub

Private Sub mnuRental_Click()
frmcheck.Show
End Sub

Private Sub mnuRetuns_Click()
frmReturns.Show
End Sub

Private Sub mnuSearch_Click()
frmSearch.Show
End Sub

Private Sub mnuStatus_Click()
If StatusBar1.Visible = True Then
   StatusBar1.Visible = False
   mnuStatus.Checked = False
Else
   StatusBar1.Visible = True
   mnuStatus.Checked = True
End If
End Sub

Private Sub mnuTileH_Click()
aMAIN.Arrange vbTileHorizontal
End Sub

Private Sub mnuTileV_Click()
aMAIN.Arrange vbTileVertical
End Sub

Private Sub mnuToolbar_Click()
If Toolbar1.Visible = True Then
   Toolbar1.Visible = False
   mnuToolbar.Checked = False
Else
   Toolbar1.Visible = True
   mnuToolbar.Checked = True
End If
End Sub



Private Sub mnuTransCustomer_Click()
Call connectDB
seanPaul = InputBox("Enter the Customer's Account Number", "Print Report", "enter valid customer id", 400, 400)
If StrPtr(seanPaul) = 0 Then
Exit Sub
Else
rs.Open "select *from loandetails where customerId='" & seanPaul & "'", conn
Set TransByCus.DataSource = rs
TransByCus.Show
End If
End Sub

Private Sub mnuUserSetup_Click()
frmUserAcc.Show
End Sub

Private Sub mnuVids_Click()
Call connectDB
rs.Open "SELECT * FROM Video", conn, adOpenStatic, adLockOptimistic
Set VideoList.DataSource = rs
VideoList.Show

End Sub

Private Sub Timer1_Timer()
    Call randomise
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
    Call mnuNew_Click
    Case 2
    Call mnuOpen_Click
    Case 4
    Call mnuSearch_Click
    Case 6
    If mnuUserSetup.Enabled = True Then
     mnuUserSetup_Click
    Else
     Button.Enabled = False
    End If
    
    Case 10
     mnuvisit_Click
   
End Select
End Sub

Private Sub mnueditvideodetails_Click()
frmEditVideoDetails.Show
With frmEditVideoDetails
      .Text1.Enabled = False

      .Text3.Enabled = False
 
      .Text5.Enabled = False
      .Text6.Enabled = False
      .Text7.Enabled = False
      .Text8.Enabled = False
      .Text9.Enabled = False
      .cmdEdit.Enabled = True
      .cmdNew.Enabled = False
      .cmdDel.Enabled = False
      .cmdUpdate.Enabled = False
End With
End Sub

Private Sub mnuvisit_Click()
ShellDocument "http://uk.geocities.com/ilducesystems/ilducehome.html"
End Sub

Private Function randomise()
msg = Int(Rnd * 5)
Select Case msg
 Case 1
     StatusBar1.Panels(1).Text = "Press F1 For Help Topics"
 Case 2
    StatusBar1.Panels(1).Text = "Update all records before printing"
 Case 3
      StatusBar1.Panels(1).Text = "Make sure database exists before Updating"
Case 4
    StatusBar1.Panels(1).Text = "Update records before Exiting Program"
Case Else
    StatusBar1.Panels(1).Text = "Go to File--> Reports to Print Reports"
End Select
End Function

