VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form CetakNilaiTranskrip 
   Caption         =   "Nilai Transkrip"
   ClientHeight    =   1890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3450
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
   ScaleHeight     =   1890
   ScaleWidth      =   3450
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   240
      TabIndex        =   1
      Text            =   "Jurusan"
      Top             =   480
      Width           =   2800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cetak Nilai Transkrip"
      Height          =   400
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   2800
   End
   Begin Crystal.CrystalReport CR 
      Left            =   1200
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "CetakNilaiTranskrip"
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
End Sub

Private Sub Command1_Click()
If Combo1 = "" Or Combo1 = "Jurusan" Then
    MsgBox "Pilih Jurusannya"
    Combo1.SetFocus
    Exit Sub
End If
    CR.SelectionFormula = "{Mahasiswa.Jurusan}='" & Combo1 & "'"
    CR.ReportFileName = App.Path & "\Nilai Transkrip.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

