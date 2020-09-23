VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmListView 
   Caption         =   "Video Details"
   ClientHeight    =   5520
   ClientLeft      =   165
   ClientTop       =   435
   ClientWidth     =   7440
   Icon            =   "frmListView.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5520
   ScaleWidth      =   7440
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imgLarge 
      Left            =   4560
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListView.frx":030A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgSmall 
      Left            =   4440
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListView.frx":0624
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   6800
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   255
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmListView"
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

    ' Create the column headers.
    Set column_header = ListView1. _
        ColumnHeaders.Add(, , "Video ID", _
        TextWidth("V-1234-23 plus remainder"))
    Set column_header = ListView1. _
        ColumnHeaders.Add(, , "Title", _
        TextWidth("James Bond : Reserve"))
    Set column_header = ListView1. _
        ColumnHeaders.Add(, , "Actor", _
        TextWidth("Jimmy Rowlings"))
    Set column_header = ListView1. _
        ColumnHeaders.Add(, , "Actress", _
        TextWidth("Janet Jackson"))
    Set column_header = ListView1. _
        ColumnHeaders.Add(, , "Release Year", _
        TextWidth("Release Year"))
    Set column_header = ListView1. _
        ColumnHeaders.Add(, , "Condition", _
        TextWidth("Very Good"))


    ' Associate the ImageLists with the
    ' ListView's Icons and SmallIcons properties.
    ListView1.Icons = imgLarge
    ListView1.SmallIcons = imgSmall

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
    Set rs = conn.Execute("SELECT * FROM Video ORDER BY VideoID", , adCmdText)

    ' Load the data.
    Do While Not rs.EOF
        Set list_item = ListView1.ListItems.Add(, , rs!VideoID)
        list_item.SubItems(1) = rs!Title
        list_item.SubItems(2) = rs!Actor
        list_item.SubItems(3) = rs!Actress
        list_item.SubItems(4) = rs!ReleaseYear

        list_item.SubItems(5) = rs!Condition

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





