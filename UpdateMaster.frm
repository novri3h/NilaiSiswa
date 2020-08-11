VERSION 5.00
Begin VB.Form UpdateMaster 
   Caption         =   "Updating Tabel Master"
   ClientHeight    =   1095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3690
   BeginProperty Font 
      Name            =   "Century"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   1095
   ScaleWidth      =   3690
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Update Tabel Master"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3500
   End
End
Attribute VB_Name = "UpdateMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call BukaDB
End Sub

Private Sub Command1_Click()
Dim SQLHapus As String
SQLHapus = "Delete From Master"
Conn.Execute (SQLHapus)

'entri NIM dan KodeMK khusus jurusan MI
Dim RSMI As New ADODB.Recordset
RSMI.Open "select distinct nim,kodemk from mahasiswa,matakuliah where mid(nim,4,1)='1' and left(kodemk,1)='1'", Conn
RSMI.MoveFirst
Do While Not RSMI.EOF
    Dim UPMI As String
    UPMI = "Insert Into Master(NIM,KodeMK) values ('" & RSMI!nim & "','" & RSMI!kodemk & "')"
    Conn.Execute (UPMI)
RSMI.MoveNext
Loop

'entri NIM dan KodeMK khusus jurusan KA
Dim RSKA As New ADODB.Recordset
RSKA.Open "select distinct nim,kodemk from mahasiswa,matakuliah where mid(nim,4,1)='2' and left(kodemk,1)='2'", Conn
RSKA.MoveFirst
Do While Not RSKA.EOF
    Dim UPKA As String
    UPKA = "Insert Into Master(NIM,KodeMK) values ('" & RSKA!nim & "','" & RSKA!kodemk & "')"
    Conn.Execute (UPKA)
RSKA.MoveNext
Loop

'entri NIM dan KodeMK khusus jurusan TK
Dim RSTK As New ADODB.Recordset
RSTK.Open "select distinct nim,kodemk from mahasiswa,matakuliah where mid(nim,4,1)='3' and left(kodemk,1)='3'", Conn
RSTK.MoveFirst
Do While Not RSTK.EOF
    Dim UPTK As String
    UPTK = "Insert Into Master(NIM,KodeMK) values ('" & RSTK!nim & "','" & RSTK!kodemk & "')"
    Conn.Execute (UPTK)
RSTK.MoveNext
Loop

MsgBox "Updating Berhasil"
Unload Me
End Sub
