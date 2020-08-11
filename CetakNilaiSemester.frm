VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form CetakNilaiSemester 
   Caption         =   "Nilai Semester"
   ClientHeight    =   2115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6600
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
   ScaleHeight     =   2115
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Cetak Rincian Nilai"
      Height          =   1815
      Left            =   3600
      TabIndex        =   5
      Top             =   120
      Width           =   2775
      Begin VB.CommandButton Command2 
         Caption         =   "Cetak Rincian Nilai"
         Height          =   400
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   2475
      End
      Begin VB.ComboBox Combo4 
         Height          =   345
         Left            =   120
         TabIndex        =   7
         Text            =   "Semester"
         Top             =   720
         Width           =   2500
      End
      Begin VB.ComboBox Combo3 
         Height          =   345
         Left            =   120
         TabIndex        =   6
         Text            =   "Kelas"
         Top             =   360
         Width           =   2500
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cetak Nilai Semester"
      Height          =   400
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   2475
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   240
      TabIndex        =   2
      Text            =   "Jurusan"
      Top             =   480
      Width           =   2500
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   2400
      TabIndex        =   1
      Top             =   1320
      Width           =   300
   End
   Begin Crystal.CrystalReport CR 
      Left            =   3000
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cetak IPS"
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2775
      Begin VB.ComboBox Combo2 
         Height          =   345
         Left            =   120
         TabIndex        =   4
         Text            =   "Semester"
         Top             =   720
         Width           =   2500
      End
   End
End
Attribute VB_Name = "CetakNilaiSemester"
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
    Combo1.AddItem Tabel!Jurusan
    Tabel.MoveNext
Loop

RSMTKL.Open "select distinct smt from matakuliah", Conn
Do While Not RSMTKL.EOF
    Combo2.AddItem RSMTKL!smt
    Combo4.AddItem RSMTKL!smt
    RSMTKL.MoveNext
Loop

RSMHS.Open "select distinct kelas from mahasiswa", Conn
Do While Not RSMHS.EOF
    Combo3.AddItem RSMHS!kelas
    RSMHS.MoveNext
Loop


Conn.Close

End Sub

Private Sub Combo1_Click()
If Combo1 = "KOMPUTER AKUNTANSI" Then
    Text1 = "2"
ElseIf Combo1 = "MANAJEMEN INFORMATIKA" Then
    Text1 = "1"
ElseIf Combo1 = "TEKNIK KOMPUTER" Then
    Text1 = "3"
End If
End Sub


Private Sub Command1_Click()
If Text1 = "" Or Combo2 = "" Or Combo2 = "Semester" Then
    MsgBox "Pilih Jurusan dan Semesternya..!"
    Combo1.SetFocus
    Exit Sub
End If
    CR.SelectionFormula = "{Mahasiswa.Jurusan}='" & Combo1 & "' and {matakuliah.kodemk}[1]='" & Text1 & "' and {matakuliah.smt}[1]='" & Combo2 & "'"
    CR.ReportFileName = App.Path & "\Nilai Semester.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub Command2_Click()
If Combo3 = "Kelas" Or Combo4 = "Semester" Then
    MsgBox "Pilih Kelas dan semester..!"
    Combo3.SetFocus
    Exit Sub
End If
    CR.SelectionFormula = "{matakuliah.smt}[1]='" & Combo4 & "' and {mahasiswa.kelas}='" & Combo3 & "'"
    CR.ReportFileName = App.Path & "\Nilai Kelas.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1

End Sub

