VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form CetakBerkas 
   Caption         =   "Cetak Berkas"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5910
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
   ScaleHeight     =   4065
   ScaleWidth      =   5910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Cetak Form Daftar Nilai"
      Height          =   495
      Left            =   240
      TabIndex        =   16
      Top             =   3480
      Width           =   5415
   End
   Begin Crystal.CrystalReport CR 
      Left            =   2520
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.ComboBox CboAbsenUjian 
      Height          =   345
      Left            =   3120
      TabIndex        =   0
      Text            =   "Jurusan"
      Top             =   1800
      Width           =   2500
   End
   Begin VB.CommandButton CmdCetakAbsenUjian 
      Caption         =   "Cetak Absen Ujian"
      Height          =   600
      Left            =   3120
      TabIndex        =   1
      Top             =   2520
      Width           =   2500
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   350
      Left            =   1320
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton CmdCetakKRS 
      Caption         =   "Cetak Kartu Ujian"
      Height          =   500
      Left            =   240
      TabIndex        =   3
      Top             =   2640
      Width           =   2500
   End
   Begin VB.ComboBox CboKRS 
      Height          =   345
      Left            =   240
      TabIndex        =   4
      Text            =   "Jurusan"
      Top             =   1800
      Width           =   2500
   End
   Begin VB.CommandButton CmdCetakKTM 
      Caption         =   "Cetak KTM"
      Height          =   500
      Left            =   3120
      TabIndex        =   5
      Top             =   720
      Width           =   2500
   End
   Begin VB.ComboBox CboKTM 
      Height          =   345
      Left            =   3120
      TabIndex        =   6
      Text            =   "Jurusan"
      Top             =   240
      Width           =   2500
   End
   Begin VB.ComboBox CboAbsen 
      Height          =   345
      Left            =   240
      TabIndex        =   7
      Text            =   "Kelas"
      Top             =   240
      Width           =   2000
   End
   Begin VB.CommandButton CmdCetakAbsenKelas 
      Caption         =   "Cetak Absen Kelas"
      Height          =   500
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   2000
   End
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1155
      ScaleWidth      =   2235
      TabIndex        =   9
      Top             =   120
      Width           =   2295
   End
   Begin VB.PictureBox Picture2 
      Height          =   1215
      Left            =   3000
      ScaleHeight     =   1155
      ScaleWidth      =   2715
      TabIndex        =   10
      Top             =   120
      Width           =   2775
   End
   Begin VB.PictureBox Picture4 
      Height          =   1815
      Left            =   3000
      ScaleHeight     =   1755
      ScaleWidth      =   2715
      TabIndex        =   11
      Top             =   1560
      Width           =   2775
   End
   Begin VB.PictureBox Picture3 
      Height          =   1815
      Left            =   120
      ScaleHeight     =   1755
      ScaleWidth      =   2715
      TabIndex        =   12
      Top             =   1560
      Width           =   2775
      Begin VB.ComboBox CboSMT 
         Height          =   345
         Left            =   1800
         TabIndex        =   15
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   " Semester :"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Semester"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      TabIndex        =   14
      Top             =   2760
      Width           =   1005
   End
End
Attribute VB_Name = "CetakBerkas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Call BukaDB
Dim Tabel As New ADODB.Recordset
Tabel.Open "select Distinct Kelas from Mahasiswa", Conn
Tabel.Requery
Do While Not Tabel.EOF
    CboAbsen.AddItem Tabel!kelas
    Tabel.MoveNext
Loop
Conn.Close

Call BukaDB
Tabel.Open "select Distinct Jurusan from Mahasiswa", Conn
Tabel.Requery
Do While Not Tabel.EOF
    CboKTM.AddItem Tabel!Jurusan
    CboKRS.AddItem Tabel!Jurusan
    CboAbsenUjian.AddItem Tabel!Jurusan
    Tabel.MoveNext
Loop
Conn.Close

Call BukaDB
Tabel.Open "select Distinct smt from matakuliah", Conn
Tabel.Requery
Do While Not Tabel.EOF
    CboSMT.AddItem Tabel!smt
    Tabel.MoveNext
Loop

End Sub


Private Sub CmdCetakAbsenKelas_Click()
If CboAbsen = "" Or CboAbsen = "Kelas" Then
    MsgBox "Anda belum memilih Kelasnya"
    Exit Sub
End If
CR.SelectionFormula = "{Mahasiswa.Kelas}='" & CboAbsen & "'"
CR.ReportFileName = App.Path & "\absen kelas.rpt"
CR.WindowState = crptMaximized
CR.RetrieveDataFiles
CR.Action = 1
End Sub

Private Sub CmdCetakKTM_Click()
If CboKTM = "" Or CboKTM = "Jurusan" Then
    MsgBox "Anda belum memilih Jurusannya"
    Exit Sub
End If
CR.SelectionFormula = "{Mahasiswa.Jurusan}='" & CboKTM & "'"
CR.ReportFileName = App.Path & "\KTM.rpt"
CR.WindowState = crptMaximized
CR.RetrieveDataFiles
CR.Action = 1
End Sub

Private Sub CboKRS_Click()
If CboKRS = "MANAJEMEN INFORMATIKA" Then
    Text1 = "1"
ElseIf CboKRS = "KOMPUTER AKUNTANSI" Then
    Text1 = "2"
ElseIf CboKRS = "TEKNIK KOMPUTER" Then
    Text1 = "3"
End If
End Sub

Private Sub CmdCetakKRS_Click()
If CboKRS = "" Or CboKRS = "Jurusan" Or CboSMT = "" Then
    MsgBox "Pilih Jurusan dan semesternya"
    Exit Sub
End If
CR.SelectionFormula = "{Mahasiswa.Jurusan}='" & CboKRS & "' and {matakuliah.kodemk}[1]='" & Text1 & "' and {matakuliah.smt}[1]='" & CboSMT & "'"
CR.ReportFileName = App.Path & "\KPU.rpt"
CR.WindowState = crptMaximized
CR.RetrieveDataFiles
CR.Action = 1
End Sub

Private Sub CmdCetakAbsenUJian_Click()
If CboAbsenUjian = "" Or CboAbsenUjian = "Jurusan" Then
    MsgBox "Anda belum memilih jurusannya"
    Exit Sub
End If
CR.SelectionFormula = "{Mahasiswa.Jurusan}='" & CboAbsenUjian & "'"
CR.ReportFileName = App.Path & "\absen ujian.rpt"
CR.WindowState = crptMaximized
CR.RetrieveDataFiles
CR.Action = 1
End Sub

Private Sub Command1_Click()
CR.ReportFileName = App.Path & "\form daftar nilai.rpt"
CR.WindowState = crptMaximized
CR.RetrieveDataFiles
CR.Action = 1

End Sub

