Attribute VB_Name = "Module1"

Public Conn As New ADODB.Connection
Public RSFormulir As ADODB.Recordset
Public RSPendaftaran As ADODB.Recordset
Public RSMHS As ADODB.Recordset
Public RSMTKL As ADODB.Recordset
Public RSDosen As ADODB.Recordset
Public RSDetailDOsen As ADODB.Recordset
Public RSTransaksi As ADODB.Recordset
Public RSOperator As ADODB.Recordset
Public RSNilai As ADODB.Recordset
Public RSPesertaHer As ADODB.Recordset
Public RSNilaiHer As ADODB.Recordset

Public pathdata As String

Public Sub BukaDB()
Set Conn = New ADODB.Connection
Set RSFormulir = New ADODB.Recordset
Set RSPendaftaran = New ADODB.Recordset
Set RSMHS = New ADODB.Recordset
Set RSMTKL = New ADODB.Recordset
Set RSDosen = New ADODB.Recordset
Set RSDetailDOsen = New ADODB.Recordset
Set RSTransaksi = New ADODB.Recordset
Set RSOperator = New ADODB.Recordset
Set RSNilai = New ADODB.Recordset
Set RSPesertaHer = New ADODB.Recordset
Set RSNilaiHer = New ADODB.Recordset
pathdata = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBNILAI.mdb"
Conn.Open pathdata
End Sub


