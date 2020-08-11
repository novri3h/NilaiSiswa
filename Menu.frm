VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Menu 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Menu Utama"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   6855
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
   Picture         =   "Menu.frx":0000
   ScaleHeight     =   5460
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   741
      Appearance      =   1
      _Version        =   327682
   End
   Begin Crystal.CrystalReport CR 
      Left            =   360
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin ComctlLib.StatusBar STBAR 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5085
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ImageList IMG 
      Left            =   840
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Menu mnmaster 
      Caption         =   "Master"
      Begin VB.Menu mnpemakai 
         Caption         =   "Pemakai"
      End
      Begin VB.Menu mnmahasiswa 
         Caption         =   "Mahasiswa"
      End
      Begin VB.Menu mnmtkl 
         Caption         =   "Mata Kuliah"
      End
      Begin VB.Menu mndosen 
         Caption         =   "Dosen"
      End
   End
   Begin VB.Menu mnupdating 
      Caption         =   "Updating"
      Begin VB.Menu mnupdatemaster 
         Caption         =   "Update Data Master"
      End
      Begin VB.Menu mntransmhs 
         Caption         =   "Transfer Data Mahasiswa"
      End
   End
   Begin VB.Menu mnawal 
      Caption         =   "Cetak Berkas Awal"
   End
   Begin VB.Menu mnnilai 
      Caption         =   "Nilai Semester"
      Begin VB.Menu mnolahnilai 
         Caption         =   "Entri Nilai"
      End
      Begin VB.Menu mnnilaismt 
         Caption         =   "Cetak Nilai Semester"
      End
      Begin VB.Menu mnwajibremed 
         Caption         =   "Mahasiswa wajib Remedial"
      End
   End
   Begin VB.Menu mndataher 
      Caption         =   "Remedial"
      Begin VB.Menu mndaftarher 
         Caption         =   "Pendaftaran Ujian Remedial"
      End
      Begin VB.Menu mncetakberkasher 
         Caption         =   "Cetak Berkas Ujian Remedial"
      End
      Begin VB.Menu mnentrinilaiher 
         Caption         =   "Entri Nilai Remedial"
      End
      Begin VB.Menu mncetaknilaiher 
         Caption         =   "Cetak Nilai Remedial"
      End
   End
   Begin VB.Menu Mncetaknilai 
      Caption         =   "Transkrip"
   End
   Begin VB.Menu mnkeluar 
      Caption         =   "Keluar"
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then End
End Sub

Private Sub mnaaa_Click()
    CR.ReportFileName = App.Path & "\Nilai Semester.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub mnbbb_Click()
    CR.ReportFileName = App.Path & "\Nilai Kelas.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub mnawal_Click()
CetakBerkas.Show vbModal
End Sub

Private Sub mncetakberkasher_Click()
CetakDataHer.Show
End Sub

Private Sub Mncetaknilai_Click()
CetakNilaiTranskrip.Show vbModal
End Sub

Private Sub mncetaknilaiher_Click()
CetakNilaiHer.Show
End Sub

Private Sub mnctkberkas_Click()
CetakBerkas.Show
End Sub

Private Sub mnctkmtkl_Click()
CetakMataKuliah.Show
End Sub

Private Sub mndaftarher_Click()
DaftarHer.Show
End Sub

Private Sub mndosen_Click()
Dosen.Show vbModal
End Sub

Private Sub mnentrinilaiher_Click()
EntriNilaiHer.Show vbModal
End Sub

Private Sub mnformulir_Click()
Formulir.Show vbModal
End Sub

Private Sub mnkeluar_Click()
End
End Sub

Private Sub mnmahasiswa_Click()
Mahasiswa.Show vbModal
End Sub

Private Sub mnmtkl_Click()
MataKuliah.Show vbModal
End Sub

Private Sub mnnilaismt_Click()
CetakNilaiSemester.Show
End Sub

Private Sub mnnilaitranskrip_Click()
CetakNilaiTranskrip.Show
End Sub

Private Sub mnolahnilai_Click()
OlahNilai2.Show vbModal
End Sub

Private Sub mnpendaftaran_Click()
Pendaftaran.Show vbModal
End Sub

Private Sub mnpemakai_Click()
Pemakai.Show vbModal
End Sub

Private Sub mnsql_Click()
UjiSQL.Show vbModal
End Sub

Private Sub mntransmhs_Click()
TransferMhs.Show vbModal
End Sub

Private Sub mnupdatemaster_Click()
UpdateMaster.Show
End Sub

Private Sub mnwajibremed_Click()
    CR.SelectionFormula = "{Nilai.Ket}='Kurang' OR {Nilai.Ket}='Gagal'"
    CR.ReportFileName = App.Path & "\wajib remedial.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub
