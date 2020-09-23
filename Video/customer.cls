VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 1  'vbSimpleBound
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private cn As ADODB.Connection
Private rs As ADODB.Recordset

Public total2
Dim temp, mov
Dim Vsearch As String


Private Sub Class_Initialize()

Set cn = New ADODB.Connection
cn.Provider = "Microsoft.Jet.OLEDB.4.0"
cn.Open App.Path & "./db.mdb"

Set rs = New ADODB.Recordset
rs.Open "Customer", cn, adOpenStatic, adLockOptimistic, adCmdTable

End Sub
'4 number of videos
Private Sub Class_GetDataMember(DataMember As String, Data As Object)

Set Data = rs
temp = rs.RecordCount

If temp < 2 Then
 mov = temp & " Customer"
Else
 mov = temp & " Customers"
End If

total2 = "Total " & mov & " in the Database"
frmEditCustomer.Label18.Caption = total2
End Sub
'4 new record

Public Function funcNew() As Boolean
With frmEditCustomer
.Text1.Enabled = True
.Text2.Enabled = True
.Text3.Enabled = True
.Text4.Enabled = True
.Text5.Enabled = True
.Text6.Enabled = True

.Text8.Enabled = True
.Text9.Enabled = True
End With
On Error Resume Next
rs.AddNew
temp = rs.RecordCount

If temp < 2 Then
 mov = temp & " Customer"
Else
 mov = temp & " Customers"
End If

total2 = "Total " & mov & " in the Database"

End Function


Public Function funcUpdate() As Boolean
On Error Resume Next
rs.Update
temp = rs.RecordCount
With frmEditCustomer
.Text1.Enabled = False
.Text2.Enabled = False
.Text3.Enabled = False
.Text4.Enabled = False
.Text5.Enabled = False
.Text6.Enabled = False

.Text8.Enabled = False
.Text9.Enabled = False
End With

If temp < 2 Then
 mov = temp & " Customer"
Else
 mov = temp & " Customers"
End If

total2 = "Total " & mov & " in stock"
End Function

Public Function funcDel() As Boolean
If temp > 0 Then
Dim answer%
answer% = MsgBox("Are you sure you want to delete " & _
"this record from your database?", 4 + 32, _
"Confirm Delete")

If answer% = vbNo Then
 funcCancel
 Exit Function
End If

'Delete a record and move to the first record of
'the database
rs.Delete
temp = rs.RecordCount

If temp < 2 Then
 mov = temp & " Customer"
Else
 mov = temp & " Customers"
End If

total2 = "Total " & mov & " in the Database"

rs.MoveFirst

If temp < 1 Then


End If
Else
MsgBox ("There are no records to delete"), vbCritical
End If
End Function

Public Function funcCount() As Boolean
temp = rs.RecordCount

If temp < 2 Then
 mov = temp & " Customer"
Else
 mov = temp & " Customer"
End If

total2 = "Total " & mov & " in the Database"
frmEditCustomer.Label18.Caption = total2
End Function


Public Function funcCancel() As Boolean
rs.CancelUpdate
End Function

Public Function funcNext() As Boolean
On Error Resume Next
 
  rs.MoveNext
   
If rs.EOF = True Then
 rs.MoveLast
End If
End Function


Public Function funcPrev() As Boolean
On Error Resume Next
  rs.MovePrevious
If rs.BOF = True Then
  rs.MoveFirst
End If
End Function

Public Function funcFirst() As Boolean
 On Error Resume Next
  rs.MoveFirst
End Function

Public Function funcLast() As Boolean
 On Error Resume Next
  rs.MoveLast
End Function




