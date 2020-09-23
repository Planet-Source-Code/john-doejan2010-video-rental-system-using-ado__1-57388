VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmListRentals 
   Caption         =   "Rentals"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6450
   Icon            =   "frmListRentals.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3990
   ScaleWidth      =   6450
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4560
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   10
      ImageHeight     =   10
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListRentals.frx":030A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5106
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      ForeColor       =   255
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Transaction ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Video ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Customer ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Due Date"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmListRentals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim db_file As String
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim column_header As ColumnHeader
Dim list_item As ListItem

   
    ' Get the data.
    db_file = App.Path
    If Right$(db_file, 1) <> "\" Then db_file = db_file & "\"
    db_file = db_file & "db.mdb"

    ' Open a connection.
    Set conn = New ADODB.Connection
    conn.ConnectionString = _
        "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & db_file & ";" & _
        "Persist Security Info=False"
    conn.Open
    Set rs = conn.Execute("SELECT * FROM LoanDetails ORDER BY TransactionID", , adCmdText)

    ' Load the data.
    Do While Not rs.EOF
        Set list_item = ListView1.ListItems.Add(, , rs!TransactionID)
        list_item.SubItems(1) = rs!VideoID
        list_item.SubItems(2) = rs!CustomerID
        list_item.SubItems(3) = rs!DueDate

    

        ' Get the next record.
        rs.MoveNext
    Loop

    ' Close the recordset and connection.
    rs.Close
    conn.Close
End Sub
Private Sub Form_Resize()
    ListView1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub






