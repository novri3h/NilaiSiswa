VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form CetakMataKuliah 
   Caption         =   "Cetak Mata Kuliah"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   5370
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Text            =   "SEMUA JURUSAN"
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton CmdPerjurusan 
      Caption         =   "Cetak Per Jurusan"
      Height          =   500
      Left            =   3240
      TabIndex        =   6
      Top             =   360
      Width           =   2000
   End
   Begin VB.ListBox List2 
      Height          =   645
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   3015
   End
   Begin VB.CommandButton CmdPersemester 
      Caption         =   "Cetak Per Semester"
      Height          =   500
      Left            =   3240
      TabIndex        =   4
      Top             =   840
      Width           =   2000
   End
   Begin VB.CommandButton CmdPerjurusandansemester 
      Caption         =   "Cetak Per Jurusan Dan Semester"
      Height          =   500
      Left            =   3240
      TabIndex        =   3
      Top             =   1320
      Width           =   2000
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   2520
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton CmdSemuadata 
      Caption         =   "Cetak Semua Data"
      Height          =   500
      Left            =   3240
      TabIndex        =   0
      Top             =   1800
      Width           =   2000
   End
   Begin Crystal.CrystalReport CR 
      Left            =   2640
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Jurusan"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Semester"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   660
   End
End
Attribute VB_Name = "CetakMataKuliah"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Call BukaDB
Dim Tabel As New ADODB.Recordset
Tabel.Open "select Distinct Jurusan from Mahasiswa", Conn
Tabel.Requery
Do While Not Tabel.EOF
    List1.AddItem Tabel!Jurusan
    Tabel.MoveNext
Loop

RSMTKL.Open "select distinct smt from matakuliah", Conn
List2.Clear
Do While Not RSMTKL.EOF
    List2.AddItem RSMTKL!SMT
    RSMTKL.MoveNext
Loop
End Sub

Private Sub List1_Click()
If List1 = "MANAJEMEN INFORMATIKA" Then
    Text1 = "1"
ElseIf List1 = "KOMPUTER AKUNTANSI" Then
    Text1 = "2"
ElseIf List1 = "TEKNIK KOMPUTER" Then
    Text1 = "3"
End If
End Sub

Private Sub CmdPerjurusan_Click()
If Text1 = "" Then
    MsgBox "Pilih Jurusannya...!"
    Exit Sub
End If
'menyaring data kode mata kuliah yang satu digit
'pertamanya tergantung jurusan
CR.SelectionFormula = "{matakuliah.kodemk}[1]='" & Text1 & "'"
CR.Formulas(0) = "jurusan='" & List1 & "'"
CR.ReportFileName = App.Path & "\Mata Kuliah.rpt"
CR.WindowState = crptMaximized
CR.RetrieveDataFiles
CR.Action = 0
CR.Reset

End Sub

Private Sub CmdPerjurusandansemester_Click()
If Text1 = "" Or Text2 = "" Then
    MsgBox "Pilih Jurusan dan Semesternya..!"
    Exit Sub
End If
'pilih data mata kuliah yang satu digit pertamanya
'tergantung jurusan dan semesternya dipilih di list2
CR.SelectionFormula = "{matakuliah.kodemk}[1]='" & Text1 & "' and {matakuliah.SMT}='" & List2 & "'"
CR.ReportFileName = App.Path & "\Mata Kuliah.rpt"
CR.Formulas(1) = "jurusan='" & List1 & "'"
CR.WindowState = crptMaximized
CR.RetrieveDataFiles
CR.Action = 1
CR.Reset
End Sub

Private Sub CmdPersemester_Click()
If Text2 = "" Then
    MsgBox "Pilih Semesternya...!"
    List2.SetFocus
    Exit Sub
End If
'pilih mata kuliah yg nilai semesternya
'sama dengan pilihan di list2
CR.SelectionFormula = "{matakuliah.SMT}='" & List2 & "'"
CR.ReportFileName = App.Path & "\Mata Kuliah.rpt"
CR.Formulas(1) = "jurusan='" & Text3 & "'"
CR.WindowState = crptMaximized
CR.RetrieveDataFiles
CR.Action = 1
CR.Reset
End Sub

Private Sub CmdSemuadata_Click()
'plih semua maka kuliah tanpa kriteria
CR.SelectionFormula = "{matakuliah.kodemk}[1]='1' or {matakuliah.kodemk}[1]='2' or {matakuliah.kodemk}[1]='3' and {matakuliah.SMT}='1' or {matakuliah.SMT}='2'"
CR.ReportFileName = App.Path & "\Mata Kuliah.rpt"
CR.Formulas(1) = "jurusan='" & Text3 & "'"
CR.WindowState = crptMaximized
CR.RetrieveDataFiles
CR.Action = 1
CR.Reset
End Sub

Private Sub List2_Click()
Text2 = List2.Text
End Sub
