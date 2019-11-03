Attribute VB_Name = "Module1"
Public con As Connection
Public res As Recordset
Public blnRecordSetOpen As Boolean
Public total As Integer
Public Sub dbconnection()
Set con = New ADODB.Connection
Set res = New ADODB.Recordset
With con
    .Provider = "Microsoft.Jet.OLEDB.4.0"
    .ConnectionString = "nacl.mdb"
    .Open
End With
End Sub
Public Sub CustomRecordSetOpen(str As String)
    If blnRecordSetOpen = True Then CloseRecordSet
    res.Open str, con, adOpenDynamic, adLockBatchOptimistic
    blnRecordSetOpen = True
End Sub

Public Sub CloseRecordSet()
    If blnRecordSetOpen = True Then res.Close
    blnRecordSetOpen = False
End Sub

Public Sub ParticularOpen(str As String)
    If blnRecordSetOpen = True Then CloseRecordSet
    res.Open str, con, adOpenKeyset, adLockBatchOptimistic
    total = res.RecordCount
    blnRecordSetOpen = True
End Sub

