VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form CetakDataHer 
   Caption         =   "Cetak Data Her"
   ClientHeight    =   2115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2805
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
   ScaleWidth      =   2805
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo2 
      Height          =   345
      Left            =   120
      TabIndex        =   4
      Text            =   "Semester"
      Top             =   480
      Width           =   2500
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   120
      TabIndex        =   2
      Text            =   "Jurusan"
      Top             =   120
      Width           =   2500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cetak Absen Peserta HER"
      Height          =   500
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   2500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cetak Kartu Peserta HER"
      Height          =   500
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   2500
   End
   Begin Crystal.CrystalReport CR 
      Left            =   1320
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Left            =   600
      TabIndex        =   3
      Top             =   2400
      Width           =   495
   End
End
Attribute VB_Name = "CetakDataHer"
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
Combo2.AddItem "1"
Combo2.AddItem "2"
End Sub

Private Sub Combo1_Click()
If Combo1 = "MANAJEMEN INFORMATIKA" Then
    Label1 = "1"
ElseIf Combo1 = "KOMPUTER AKUNTANSI" Then
    Label1 = "2"
ElseIf Combo1 = "TEKNIK KOMPUTER" Then
    Label1 = "3"
End If
End Sub


Private Sub Command1_Click()
If Label1 = "" Or Combo2 = "Semester" Then
    MsgBox "Pilih Jurusan dan semesternya..!"
    Combo1.SetFocus
    Exit Sub
End If
'filter laporan berdasarkan jurusan dan semesternya bernilai 1
CR.SelectionFormula = "{Mahasiswa.Jurusan}='" & Combo1 & "' and {matakuliah.kodemk}[1]='" & Label1.Caption & "' and {matakuliah.smt}[1]='" & Combo2 & "'"
CR.ReportFileName = App.Path & "\KPU Her.rpt"
CR.WindowState = crptMaximized
CR.RetrieveDataFiles
CR.Action = 1
End Sub

Private Sub Command2_Click()
If Label1 = "" Or Combo2 = "Semester" Then
    MsgBox "Pilih Jurusan dan semesternya..!"
    Combo1.SetFocus
    Exit Sub
End If
'filter laporan berdasarkan jurusan dan semesternya bernilai 1
CR.SelectionFormula = "{Mahasiswa.Jurusan}='" & Combo1 & "' and {matakuliah.kodemk}[1]='" & Label1 & "' and {matakuliah.smt}[1]='" & Combo2 & "'"
CR.ReportFileName = App.Path & "\peserta her.rpt"
CR.WindowState = crptMaximized
CR.RetrieveDataFiles
CR.Action = 1
End Sub

