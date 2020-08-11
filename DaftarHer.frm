VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form DaftarHer 
   Caption         =   "Pendaftaran Ujian Her"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7350
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
   ScaleHeight     =   11055
   ScaleWidth      =   20220
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "DaftarHer.frx":0000
      Height          =   4095
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   7223
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   2
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Nomor"
         Caption         =   "No"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Kode"
         Caption         =   "Kode"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Nama"
         Caption         =   "Nama"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "SKS"
         Caption         =   "SKS"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   494,929
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   2849,953
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   494,929
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3720
      Top             =   5160
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Height          =   4335
      Left            =   5400
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "Tutup"
      Height          =   350
      Left            =   2640
      TabIndex        =   9
      Top             =   5160
      Width           =   850
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "Batal"
      Height          =   350
      Left            =   1800
      TabIndex        =   8
      Top             =   5160
      Width           =   850
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "Simpan"
      Height          =   350
      Left            =   960
      TabIndex        =   7
      Top             =   5160
      Width           =   850
   End
   Begin VB.TextBox TxtNIM 
      Height          =   350
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   1000
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode Mata Kuliah"
      Height          =   345
      Left            =   5400
      TabIndex        =   13
      Top             =   120
      Width           =   1785
   End
   Begin VB.Label JmlList 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   5400
      TabIndex        =   12
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label JmlKode 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   120
      TabIndex        =   10
      Top             =   5160
      Width           =   750
   End
   Begin VB.Label LblJurusan 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   2160
      TabIndex        =   5
      Top             =   480
      Width           =   3100
   End
   Begin VB.Label LblKelas 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1080
      TabIndex        =   4
      Top             =   480
      Width           =   1005
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kelas"
      Height          =   345
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   950
   End
   Begin VB.Label LblNamaMhs 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   3100
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " NIM"
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   950
   End
End
Attribute VB_Name = "DaftarHer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub Form_activate()
Call BukaDB
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBnilai.mdb"
Adodc1.RecordSource = "TRDaftarher"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
Call TabelKosong
TxtNIM = ""
JmlList = List1.ListCount
End Sub
 
Private Sub Form_Load()
Call BukaDB
End Sub

Private Sub TxtNIM_KeyPress(KeyAscii As Integer)
TxtNIM.MaxLength = 7
If KeyAscii = 13 Then
    If Len(TxtNIM) < 7 Then
        MsgBox "nim harus 7 digit"
        Exit Sub
    End If
    Call BukaDB
    RSMHS.Open "select * from mahasiswa where nim='" & TxtNIM & "'", Conn
    If RSMHS.EOF Then
        MsgBox "Nim tidak terdaftar"
        TxtNIM.SetFocus
    Else
        LblNamaMhs = Space(1) & RSMHS!namamhs
        LblKelas = Space(1) & RSMHS!kelas
        LblJurusan = Space(1) & RSMHS!Jurusan
        DataGrid1.SetFocus
        DataGrid1.Col = 1
        
        Dim RScarimk As New ADODB.Recordset
        RScarimk.Open "SELECT DISTINCT MATAKULIAH.KODEMK FROM MATAKULIAH,MAHASISWA,NILAI WHERE NILAI.KODEMK=MATAKULIAH.KODEMK AND MAHASISWA.NIM=NILAI.NIM AND NILAI.TOTAL<60 AND MAHASISWA.NIM='" & TxtNIM & "'", Conn
        List1.Clear
        If Not RScarimk.EOF Then
            Do While Not RScarimk.EOF
                List1.AddItem RScarimk!kodemk
                RScarimk.MoveNext
            Loop
        Else
            MsgBox "Nim ini sebenarnya tidak harus ikut remedial"
            TxtNIM.SetFocus
            Exit Sub
        End If
    End If
End If
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
If DataGrid1.Col = 1 Then
    Call BukaDB
    RSMTKL.Open "select * from matakuliah where kodemk='" & Adodc1.Recordset!Kode & "'", Conn
    RSMTKL.Requery
    If RSMTKL.EOF Then
        MsgBox "Kode Mata Kuliah Tidak Terdaftar"
        DataGrid1.Col = 1
        Exit Sub
    End If
    Adodc1.Recordset!Kode = RSMTKL!kodemk
    Adodc1.Recordset!Nama = RSMTKL!namamk
    Adodc1.Recordset!SKS = RSMTKL!SKS
    Adodc1.Recordset.Update
    Adodc1.Recordset.MoveNext
    DataGrid1.Col = 1
    JmlKode.Caption = Str(JmlData)
    Call JmlData
End If
End Sub

Function JmlData()

On Error Resume Next
Adodc1.Recordset.MoveFirst
Kode = 0
Do While Not Adodc1.Recordset.EOF And Adodc1.Recordset!Kode <> vbNullString ' 0
    Kode = Kode + 1
    Adodc1.Recordset.MoveNext
    JmlKode = Kode
Loop

End Function

Private Sub CmdSimpan_Click()
If JmlKode.Caption = "" Or TxtNIM = "" Then
    MsgBox "Data Belum Lengkap"
    TxtNIM.SetFocus
    Exit Sub
Else
    Adodc1.Recordset.MoveFirst
    Do While Not Adodc1.Recordset.EOF
        If Adodc1.Recordset!Kode <> vbNullString Then
            SQLTambah = "Insert Into PesertaHer(Nim,KodeMK) values ('" & TxtNIM & "','" & Adodc1.Recordset!Kode & "')"
            Conn.Execute (SQLTambah)
        End If
    Adodc1.Recordset.MoveNext
    Loop
    Blank
    TxtNIM = ""
    TxtNIM.SetFocus
End If

End Sub

Sub datagrid1_Keypress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then KeyAscii = 0
End Sub

Private Sub CmdBatal_Click()
Blank
TxtNIM.SetFocus
End Sub

Private Sub CmdTutup_Click()
Unload Me
End Sub

Sub Blank()
Call TabelKosong
LblNamaMhs = ""
LblKelas = ""
LblJurusan = ""
JmlKode = ""
JmlKode = ""
TxtNIM = ""
List1.Clear
End Sub

Function TabelKosong()
If Not Adodc1.Recordset.RecordCount = 0 Then
    Adodc1.Recordset.MoveFirst
    Do While Not Adodc1.Recordset.EOF
        Adodc1.Recordset.Delete
        Adodc1.Recordset.MoveNext
    Loop
End If
For i = 1 To 15
    Adodc1.Recordset.AddNew
    Adodc1.Recordset!Nomor = i
    Adodc1.Recordset.Update
Next i
Adodc1.Recordset.MoveFirst
End Function

