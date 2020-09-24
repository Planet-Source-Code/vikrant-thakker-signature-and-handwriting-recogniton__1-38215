Attribute VB_Name = "Connection"
Public conn As ADODB.Connection
Public rsChar As ADODB.Recordset
Public Sub Main()
On Error GoTo merr
Dim str1 As String

Set conn = New ADODB.Connection

On Error GoTo errOff97

str1 = "provider=microsoft.jet.oledb.4.0;data source="
str1 = str1 & App.Path & "\data.mdb"


errOff97:
str1 = "provider=microsoft.jet.oledb.3.51;data source="
str1 = str1 & App.Path & "\data.mdb"

conn.Open str1


Set rsChar = New ADODB.Recordset
rsChar.Open "select * from MastChar", conn, adOpenStatic, adLockOptimistic


Load frmMain
frmMain.Show

Exit Sub
merr:
MsgBox Err.Description, vbOKOnly, "Anveshak"
End Sub


