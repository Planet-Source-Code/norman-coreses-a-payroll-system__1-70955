Attribute VB_Name = "Module1"



' comments or suggestion please email @: cell_nor@yahoo.com
' if you want full code of this system just contact @: 639212733741



Option Explicit

Global username As String
Global password As String

Public Sub connection(ByRef dConnection As ADODB.connection, ByVal dLocation As String, ByVal spass As String)
    dConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dLocation & ";Persist Security Info=False; Jet OLEDB:Database password=" & spass
End Sub

Public Sub recordset(ByRef sRecordset As ADODB.recordset, ByRef sConnection As ADODB.connection, ByVal sSQL As String)
With sRecordset
.CursorLocation = adUseClient
.Open sSQL, sConnection, adOpenKeyset, adLockOptimistic
End With
End Sub

Public Function sempty(ByRef stext As Variant) As Boolean
If stext = Empty Then
sempty = True
MsgBox "please filled the fields", vbExclamation, "Engcore Merchandising"
stext.SetFocus

Else

sempty = False

End If
End Function

Public Function isempty(ByRef stext As Variant) As Boolean
If stext = "" Then
isempty = True
MsgBox "Enter zero when empty in the fields.", vbExclamation, "Engcore Merchandising"
stext.SetFocus

Else

isempty = False

End If
End Function

Public Function snumber(ByRef stext As Variant) As Boolean
If IsNumeric(stext) = False Then
snumber = True

MsgBox "Cannot accept non-numeric.", vbExclamation, "Engcore Merchandising"

stext.SetFocus
Else

snumber = False

End If
End Function

Public Function recexist(ByVal Table As String, ByVal Field As String, ByRef Entry As Variant, ByRef cn As ADODB.connection) As Boolean
Dim RS As New ADODB.recordset
recexist = True

Call recordset(RS, cn, "SELECT * FROM " & Table & " WHERE " & Field & " ='" & Entry & "'")

If RS.RecordCount > 0 Then
MsgBox "Cannot saved because (" & Entry & ") is exist in the record.", vbExclamation, "Engcore Merchandising"

Entry.SetFocus

Else

recexist = False

End If

Set cn = Nothing
Set RS = Nothing

End Function

Public Function recexistlv(ByVal Table As String, ByVal Field As String, ByRef Entry As Variant, ByRef cn As ADODB.connection) As Boolean
Dim RS As New ADODB.recordset
recexistlv = True

Call recordset(RS, cn, "SELECT * FROM " & Table & " WHERE " & Field & " ='" & Entry & "'")

If RS.RecordCount > 0 Then
MsgBox "Cannot Delete because (" & Entry & ") is exist in the record.", vbExclamation, "Engcore Merchandising"

Else

recexistlv = False

End If

Set cn = Nothing
Set RS = Nothing

End Function

Public Function recexistdel(ByVal Table As String, ByVal Field As String, ByRef Entry As Variant, ByRef cn As ADODB.connection) As Boolean
Dim RS As New ADODB.recordset
recexistdel = True

Call recordset(RS, cn, "SELECT * FROM " & Table & " WHERE " & Field & " ='" & Entry & "'")

If RS.RecordCount > 0 Then
MsgBox "Cannot saved because (" & Entry & ") is exist account in the debt record. click print button to print", vbExclamation, "Engcore Merchandising"

Entry.SetFocus

Else

recexistdel = False

End If

Set cn = Nothing
Set RS = Nothing

End Function

Public Function existdel(ByVal Table As String, ByVal Field As String, ByRef Entry As Variant, ByRef cn As ADODB.connection) As Boolean
Dim RS As New ADODB.recordset
existdel = True

Call recordset(RS, cn, "SELECT * FROM " & Table & " WHERE " & Field & " ='" & Entry & "'")

If RS.RecordCount > 0 Then
MsgBox "Cannot delete because (" & Entry & ") is exist account in the debt record.", vbExclamation, "Engcore Merchandising"

Entry.SetFocus

Else

existdel = False

End If

Set cn = Nothing
Set RS = Nothing

End Function

Public Sub hlfocus(ByRef stext As TextBox)
With stext
.SelStart = 0
.SelLength = Len(stext)
End With
End Sub

Public Function recfound(ByRef sRecordset As ADODB.recordset, ByVal sField As String, ByVal sfindtext As String) As Boolean
    sRecordset.Requery
    sRecordset.Find sField & "='" & sfindtext & "'"
               
If sRecordset.EOF Then
    recfound = False
Else
    recfound = True
    username = sRecordset.Fields!username
    password = sRecordset.Fields!password
    End If
End Function

Sub main()
Form20.Show 1
End Sub
