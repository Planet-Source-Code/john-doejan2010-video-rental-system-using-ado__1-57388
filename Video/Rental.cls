VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 1  'vbSimpleBound
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Rental"
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
rs.Open "LoanDetails", cn, adOpenStatic, adLockOptimistic, adCmdTable

End Sub
'4 number of videos
Private Sub Class_GetDataMember(DataMember As String, Data As Object)

Set Data = rs
temp = rs.RecordCount

If temp < 2 Then
 mov = temp
Else
 mov = temp
End If

total2 = mov

End Sub
'4 new record

Public Function funcNew() As Boolean


rs.AddNew

End Function


Public Function funcUpdate() As Boolean

rs.Update
temp = rs.RecordCount


If temp < 2 Then
 mov = temp
Else
 mov = temp
End If

total2 = mov
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
 mov = temp
Else
 mov = temp
End If

total2 = mov

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
 mov = temp
Else
 mov = temp
End If

total2 = mov

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






