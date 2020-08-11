VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Mahasiswa 
   Caption         =   "Data Mahasiswa"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5865
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
   ScaleHeight     =   5445
   ScaleWidth      =   5865
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "&Tutup"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   1800
      Width           =   1250
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   1800
      Width           =   1250
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1800
      Width           =   1250
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Input"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   1250
   End
   Begin VB.ComboBox CBJurusan 
      Height          =   345
      Left            =   1200
      TabIndex        =   4
      Top             =   240
      Width           =   1250
   End
   Begin VB.TextBox TNama 
      Height          =   350
      Left            =   1200
      TabIndex        =   6
      Top             =   960
      Width           =   4500
   End
   Begin VB.TextBox TNIM 
      Height          =   350
      Left            =   1200
      TabIndex        =   5
      Top             =   600
      Width           =   1250
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Mahasiswa.frx":0000
      Height          =   1995
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   3519
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
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "NIM"
         Caption         =   "NIM"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "NAMAMHS"
         Caption         =   "NAMA"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "KELAS"
         Caption         =   "KELAS"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "JURUSAN"
         Caption         =   "JURUSAN"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "SMT1"
         Caption         =   "SMT1"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "SMT2"
         Caption         =   "SMT2"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "SMT3"
         Caption         =   "SMT3"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "SMT4"
         Caption         =   "SMT4"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "SMT5"
         Caption         =   "SMT5"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "SMT6"
         Caption         =   "SMT6"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1995,024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1995,024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   464,882
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   464,882
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   464,882
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   464,882
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   464,882
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   464,882
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   345
      Left            =   120
      Top             =   4560
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   609
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HT"
      Height          =   345
      Left            =   4200
      TabIndex        =   21
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label LBHT 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   345
      Left            =   4680
      TabIndex        =   20
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label LBTK 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   3720
      TabIndex        =   19
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label LBJurusan 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   2640
      TabIndex        =   18
      Top             =   240
      Width           =   3045
   End
   Begin VB.Label LBKA 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   2640
      TabIndex        =   17
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label LBMI 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1680
      TabIndex        =   16
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label LBKelas 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   3720
      TabIndex        =   15
      Top             =   600
      Width           =   1965
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TK"
      Height          =   345
      Left            =   3240
      TabIndex        =   14
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "KA"
      Height          =   345
      Left            =   2160
      TabIndex        =   13
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MI"
      Height          =   345
      Left            =   1200
      TabIndex        =   12
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kelas"
      Height          =   345
      Left            =   2640
      TabIndex        =   11
      Top             =   600
      Width           =   1005
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama"
      Height          =   345
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   1005
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " NIM"
      Height          =   345
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   1005
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Jurusan"
      Height          =   345
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   1005
   End
End
Attribute VB_Name = "Mahasiswa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_activate()
Call BukaDB
Adodc1.ConnectionString = pathdata
Adodc1.RecordSource = "MAHASISWA"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
Call JumlahMI
Call JumlahKA
Call JumlahTK
Call JumlahHT
End Sub

Private Sub Form_Load()
Call BukaDB
Call KondisiAwal
TNIM.MaxLength = 7
Call KondisiAwal
Call ListJurusan
End Sub

Private Sub CBJurusan_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 27 Then Unload Me
If KeyAscii = 13 Then
    If CBJurusan = "MI" Then
        LBJurusan = "MANAJEMEN INFORMATIKA"
        Call Nim_OTO_MI
        Call KelasMI
    ElseIf CBJurusan = "KA" Then
        LBJurusan = "KOMPUTER AKUNTANSI"
        Call Nim_OTO_KA
        Call KelasKA
    ElseIf CBJurusan = "TK" Then
        LBJurusan = "TEKNIK KOMPUTER"
        Call Nim_OTO_TK
        Call KelasTK
    ElseIf CBJurusan = "HT" Then
        LBJurusan = "KARANTINA HEWAN DAN TUMBUHAN"
        Call Nim_OTO_HT
        Call KelasHT
    End If
    
    TNIM.Enabled = False
    If CBJurusan <> "MI" And CBJurusan <> "KA" And CBJurusan <> "TK" And CBJurusan <> "HT" Then
        MsgBox ("Jurusan tidak terdaftar, harusnya MI, KA, TK atau HT")
        CBJurusan.SetFocus
        Exit Sub
    End If
    TNama.SetFocus
End If
End Sub


Private Sub CBJurusan_Click()
If CBJurusan = "MI" Then
    LBJurusan = "MANAJEMEN INFORMATIKA"
    Call Nim_OTO_MI
    Call KelasMI
ElseIf CBJurusan = "KA" Then
    LBJurusan = "KOMPUTER AKUNTANSI"
    Call Nim_OTO_KA
    Call KelasKA
ElseIf CBJurusan = "TK" Then
    LBJurusan = "TEKNIK KOMPUTER"
    Call Nim_OTO_TK
    Call KelasTK
ElseIf CBJurusan = "HT" Then
    LBJurusan = "KARANTINA HEWAN DAN TUMBUHAN"
    Call Nim_OTO_HT
    Call KelasHT
End If
TNIM.Enabled = False
End Sub

Private Sub Command1_Click()
If Command1.Caption = "&Input" Then
    Command1.Caption = "Simpan"
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Caption = "&Batal"
    Call Terang
    CBJurusan.SetFocus
    Exit Sub
Else
    If CBJurusan = "" Or TNIM = "" Or TNama = "" Or LBKelas = "" Then
        MsgBox "Data belum lengkap"
        Exit Sub
    Else
        Dim Simpan As String
        Simpan = "insert into mahasiswa(NIM,NamaMhs,Jurusan,Kelas,SMT1,SMT2,SMT3,SMT4,SMT5,SMT6) values ('" & TNIM & "','" & TNama & "','" & LBJurusan & "','" & LBKelas & "',1,0,0,0,0,0)"
        Conn.Execute Simpan
        Form_activate
        KondisiAwal
        Command1.SetFocus
    End If
End If
End Sub

Private Sub Command2_Click()
If Command2.Caption = "&Edit" Then
    Command2.Caption = "Simpan"
    Command1.Enabled = False
    Command3.Enabled = False
    Command4.Caption = "&Batal"
    Call Terang
    TNIM.SetFocus
    Exit Sub
Else
    If TNama = "" Then
        MsgBox "Data belum lengkap"
        Exit Sub
    Else
        
        Dim Simpan As String
        Simpan = "Update mahasiswa set Namamhs='" & TNama & "' where nim='" & TNIM & "'"
        Conn.Execute Simpan
        Call KondisiAwal
        Adodc1.Refresh
        DataGrid1.Refresh
        Command2.SetFocus
    End If
End If
End Sub

Private Sub Command3_Click()
If Command3.Caption = "&Hapus" Then
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Caption = "&Hapus"
    Command4.Caption = "&Batal"
    TNIM.Enabled = True
    TNIM.SetFocus
End If
End Sub

Private Sub Command4_Click()
Select Case Command4.Caption
    Case "&Tutup"
        'End
        Unload Me
    Case "&Batal"
        Call KondisiAwal
End Select
End Sub

Sub KelasMI()
If Val(LBMI) < 5 And CBJurusan = "MI" Then
    LBKelas = "MI1A"
ElseIf Val(LBMI) = 5 And CBJurusan = "MI" Then
    LBKelas = "MI1B"
ElseIf Val(LBMI) >= 6 And Val(LBMI) < 10 And CBJurusan = "MI" Then
    LBKelas = "MI1B"
ElseIf Val(LBMI) = 10 And CBJurusan = "MI" Then
    LBKelas = "MI1C"
ElseIf Val(LBMI) > 10 And CBJurusan = "MI" Then
    LBKelas = "MI1C"
End If
End Sub

Sub KelasKA()
If LBKA < 5 And CBJurusan = "KA" Then
    LBKelas = "KA1A"
ElseIf LBKA = 5 And CBJurusan = "KA" Then
    LBKelas = "KA1B"
ElseIf LBKA >= 6 And LBKA < 10 And CBJurusan = "KA" Then
    LBKelas = "KA1B"
ElseIf LBKA = 10 And CBJurusan = "KA" Then
    LBKelas = "KA1C"
ElseIf LBKA > 10 And CBJurusan = "KA" Then
    LBKelas = "KA1C"
End If
End Sub

Sub KelasTK()
If LBTK < 5 And CBJurusan = "TK" Then
    LBKelas = "TK1A"
ElseIf LBTK = 5 And CBJurusan = "TK" Then
    LBKelas = "TK1B"
ElseIf LBTK >= 6 And LBTK < 10 And CBJurusan = "TK" Then
    LBKelas = "TK1B"
ElseIf LBTK = 10 And CBJurusan = "TK" Then
    LBKelas = "TK1C"
ElseIf LBTK > 10 And CBJurusan = "TK" Then
    LBKelas = "TK1C"
End If
End Sub


Sub KelasHT()
If LBHT < 5 And CBJurusan = "HT" Then
    LBKelas = "HT1A"
ElseIf LBHT = 5 And CBJurusan = "HT" Then
    LBKelas = "HT1B"
ElseIf LBHT >= 6 And LBHT < 10 And CBJurusan = "HT" Then
    LBKelas = "HT1B"
ElseIf LBHT = 10 And CBJurusan = "HT" Then
    LBKelas = "HT1C"
ElseIf LBHT > 10 And CBJurusan = "HT" Then
    LBKelas = "HT1C"
End If
End Sub

Private Sub Nim_OTO_MI()
Call BukaDB
Dim RS As New ADODB.Recordset
RS.Open "select NIM from Mahasiswa where Jurusan='MANAJEMEN INFORMATIKA' order by nim desc", Conn
RS.Requery
If RS.EOF Then
    Urutan = Format(Date, "YY") + "01" + "001"
    TNIM = Urutan
    Exit Sub
Else
    Hitung = Right(RS!nim, 3) + 1
    Urutan = Format(Date, "YY") + "01" + Right("000" & Hitung, 3)
End If
TNIM = Urutan
End Sub

Sub Nim_OTO_KA()
Call BukaDB
Dim RS As New ADODB.Recordset
RS.Open "select NIM from Mahasiswa where Jurusan='KOMPUTER AKUNTANSI' order by nim desc", Conn
RS.Requery
If RS.EOF Then
    Urutan = Format(Date, "YY") + "02" + "001"
    TNIM = Urutan
Else
    Hitung = Right(RS!nim, 3) + 1
    Urutan = Format(Date, "YY") + "02" + Right("000" & Hitung, 3)
End If
TNIM = Urutan
End Sub

Sub Nim_OTO_TK()
Call BukaDB
Dim RS As New ADODB.Recordset
RS.Open "select NIM from Mahasiswa where Jurusan='TEKNIK KOMPUTER' order by nim desc", Conn
RS.Requery
If RS.EOF Then
    Urutan = Format(Date, "YY") + "03" + "001"
    TNIM = Urutan
Else
    Hitung = Right(RS!nim, 3) + 1
    Urutan = Format(Date, "YY") + "03" + Right("000" & Hitung, 3)
End If
TNIM = Urutan
End Sub

Sub Nim_OTO_HT()
Call BukaDB
Dim RS As New ADODB.Recordset
RS.Open "select NIM from Mahasiswa where Jurusan='KARANTINA HEWAN DAN TUMBUHAN' order by nim desc", Conn
RS.Requery
If RS.EOF Then
    Urutan = Format(Date, "YY") + "04" + "001"
    TNIM = Urutan
Else
    Hitung = Right(RS!nim, 3) + 1
    Urutan = Format(Date, "YY") + "04" + Right("000" & Hitung, 3)
End If
TNIM = Urutan
End Sub


Function JumlahMI()
Dim RS As New ADODB.Recordset
RS.Open "select count(NIM) as JMLMI from Mahasiswa where jurusan='MANAJEMEN INFORMATIKA'", Conn
LBMI = RS!JMLMI
End Function

Function JumlahKA()
Dim RS As New ADODB.Recordset
 RS.Open "select count(NIM) as JMLKA from Mahasiswa where jurusan='KOMPUTER AKUNTANSI'", Conn
LBKA = RS!JMLKA
End Function

Function JumlahTK()
Dim RS As New ADODB.Recordset
RS.Open "select count(NIM) as JMLTK from Mahasiswa where jurusan='TEKNIK KOMPUTER'", Conn
LBTK = RS!JMLTK
End Function

Function JumlahHT()
Dim RS As New ADODB.Recordset
RS.Open "select count(NIM) as JMLHT from Mahasiswa where jurusan='KARANTINA HEWAN DAN TUMBUHAN'", Conn
LBHT = RS!JMLHT
End Function


Sub ListJurusan()
CBJurusan.AddItem ("MI")
CBJurusan.AddItem ("KA")
CBJurusan.AddItem ("TK")
CBJurusan.AddItem ("HT")
End Sub


Sub KondisiAwal()
Adodc1.ConnectionString = pathdata
Adodc1.RecordSource = "Mahasiswa"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
Call Gelap
Call Kosongkan
Call JumlahMI
Call JumlahKA
Call JumlahTK
Command1.Caption = "&Input"
Command2.Caption = "&Edit"
Command3.Caption = "&Hapus"
Command4.Caption = "&Tutup"
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
End Sub

Sub Tampilkan()
With RSMHS
    CBJurusan = Left(!kelas, 2)
    TNama = !namamhs
    LBKelas = !kelas
    LBJurusan = !Jurusan
End With
End Sub


Private Sub TNama_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    If Command1.Caption = "Simpan" Then Command1.SetFocus
    If Command2.Caption = "Simpan" Then Command2.SetFocus
End If
End Sub

Private Sub TNIM_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Len(TNIM) < 7 Then
        MsgBox "NIM harus 7 digit"
        TNIM.SetFocus
        Exit Sub
    End If
    
    'untuk &Input
    If Command1.Caption = "&Simpan" Then
        Call CariNIM
            If Not RSMHS.EOF Then
                Gelap
                Tampilkan
                MsgBox "Nomor Mahasiswa Sudah Ada"
                Kosongkan
                Terang
                TNIM.SetFocus
            Else
                Terang
                Gelap
                TNama.SetFocus
            End If
    
    'untuk &Edit
    ElseIf Command2.Caption = "&Simpan" Then
            Call CariNIM
            If Not RSMHS.EOF Then
                Tampilkan
                Terang
                TNIM.Enabled = False
                TNama.SetFocus
            Else
                MsgBox "Nomor Mahasiswa Tidak Ditemukan"
                Kosongkan
                Terang
                TNIM.SetFocus
            End If
        
    'untuk hapus
    ElseIf Command3.Caption = "&Hapus" Then
        With RSMHS
            Call CariNIM
            If Not RSMHS.EOF Then
                Tampilkan
                Gelap
                Pesan = MsgBox("Yakin Data Ini Akan Dihapus...?", vbYesNo)
                If Pesan = vbYes Then
                    
                    Dim HapusMhs As String
                    HapusMhs = "delete * from mahasiswa where nim='" & TNIM & "'"
                    Conn.Execute (HapusMhs)
                    Adodc1.Refresh
                    DataGrid1.Refresh
                    KondisiAwal
                    Command3.SetFocus
                Else
                    KondisiAwal
                    Command3.SetFocus
                End If
            Else
                MsgBox "NIM Tidak Ditemukan"
                Kosongkan
                Terang
                TNIM.SetFocus
            End If
        End With
   End If
End If
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub


Private Sub Kosongkan()
Dim Ctl As Control
For Each Ctl In Me
    If TypeName(Ctl) = "TextBox" Or TypeName(Ctl) = "ComboBox" Then
        Ctl.Text = ""
    End If
Next
LBJurusan = ""
LBKelas = ""
End Sub

Private Sub Terang()
Dim Ctl As Control
For Each Ctl In Me
    If TypeName(Ctl) = "TextBox" Or TypeName(Ctl) = "ComboBox" Then
        Ctl.Enabled = True
    End If
Next
End Sub

Private Sub Gelap()
Dim Ctl As Control
For Each Ctl In Me
    If TypeName(Ctl) = "TextBox" Or TypeName(Ctl) = "ComboBox" Then
        Ctl.Enabled = False
    End If
Next
End Sub

Sub CariNIM()
Call BukaDB
RSMHS.Open "Select * From Mahasiswa where NIM='" & TNIM & "'", Conn
End Sub
