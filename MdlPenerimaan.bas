Attribute VB_Name = "MdlPenerimaan"

Public Conn As New ADODB.Connection
Public RSPelamar As ADODB.Recordset
Public RSKasir As ADODB.Recordset
Public RSJadwal As ADODB.Recordset
Public RSNilai As ADODB.Recordset
Public RSNilai1 As ADODB.Recordset
Public RSDetail As ADODB.Recordset
Public RSHasil As ADODB.Recordset

Public Sub BukaDB()
Set Conn = New ADODB.Connection
Set RSPelamar = New ADODB.Recordset
Set RSKasir = New ADODB.Recordset
Set RSJadwal = New ADODB.Recordset
Set RSNilai = New ADODB.Recordset
Set RSNilai1 = New ADODB.Recordset
Set RSDetail = New ADODB.Recordset
Set RSHasil = New ADODB.Recordset
Conn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\ADOSeleksi.mdb"
End Sub

