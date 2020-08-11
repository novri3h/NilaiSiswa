VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Pendaftaran 
   Caption         =   "Pendaftaran"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TKota 
      Height          =   350
      Left            =   1320
      TabIndex        =   37
      Top             =   3960
      Width           =   4500
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1995
      Left            =   240
      TabIndex        =   33
      Top             =   5040
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   3519
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   2
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   405
      Left            =   4080
      Top             =   120
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   714
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
   Begin VB.TextBox TNIM 
      Height          =   350
      Left            =   1320
      TabIndex        =   5
      Top             =   1080
      Width           =   1250
   End
   Begin VB.TextBox TNama 
      Height          =   350
      Left            =   1320
      TabIndex        =   6
      Top             =   1440
      Width           =   4500
   End
   Begin VB.TextBox TempatLhr 
      Height          =   350
      Left            =   1320
      TabIndex        =   7
      Top             =   1800
      Width           =   4500
   End
   Begin VB.TextBox TTelepon 
      Height          =   350
      Left            =   1320
      TabIndex        =   11
      Top             =   2880
      Width           =   4500
   End
   Begin VB.TextBox TOrtu 
      Height          =   350
      Left            =   1320
      TabIndex        =   12
      Top             =   3240
      Width           =   4500
   End
   Begin VB.TextBox TAlamat 
      Height          =   350
      Left            =   1320
      TabIndex        =   13
      Top             =   3600
      Width           =   4500
   End
   Begin VB.ComboBox CBJurusan 
      Height          =   315
      Left            =   1320
      TabIndex        =   4
      Top             =   720
      Width           =   1250
   End
   Begin VB.ComboBox CBGender 
      Height          =   315
      Left            =   1320
      TabIndex        =   9
      Top             =   2520
      Width           =   1250
   End
   Begin VB.ComboBox CBAgama 
      Height          =   315
      Left            =   3840
      TabIndex        =   10
      Top             =   2520
      Width           =   1965
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Input"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Tutup"
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   4440
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker TanggalLhr 
      Height          =   345
      Left            =   1320
      TabIndex        =   8
      Top             =   2160
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
      _Version        =   393216
      Format          =   20578305
      CurrentDate     =   39302
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kota"
      Height          =   345
      Left            =   240
      TabIndex        =   38
      Top             =   3960
      Width           =   1005
   End
   Begin VB.Label KodeOpr 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   2760
      TabIndex        =   36
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label TglDaftar 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1320
      TabIndex        =   35
      Top             =   120
      Width           =   1250
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tgl Daftar"
      Height          =   345
      Left            =   240
      TabIndex        =   34
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Jurusan"
      Height          =   345
      Left            =   240
      TabIndex        =   32
      Top             =   720
      Width           =   1005
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " NIM"
      Height          =   345
      Left            =   240
      TabIndex        =   31
      Top             =   1080
      Width           =   1005
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama"
      Height          =   345
      Left            =   240
      TabIndex        =   30
      Top             =   1440
      Width           =   1005
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kelas"
      Height          =   345
      Left            =   2760
      TabIndex        =   29
      Top             =   1080
      Width           =   1005
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tgl Lahir"
      Height          =   345
      Left            =   240
      TabIndex        =   28
      Top             =   2160
      Width           =   1005
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tempat Lhr"
      Height          =   345
      Left            =   240
      TabIndex        =   27
      Top             =   1800
      Width           =   1005
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Gender"
      Height          =   345
      Left            =   240
      TabIndex        =   26
      Top             =   2520
      Width           =   1005
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Agama"
      Height          =   345
      Left            =   2760
      TabIndex        =   25
      Top             =   2520
      Width           =   1005
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Telepon"
      Height          =   345
      Left            =   240
      TabIndex        =   24
      Top             =   2880
      Width           =   1005
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama Ortu"
      Height          =   345
      Left            =   240
      TabIndex        =   23
      Top             =   3240
      Width           =   1005
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Alamat"
      Height          =   345
      Left            =   240
      TabIndex        =   22
      Top             =   3600
      Width           =   1005
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MI"
      Height          =   345
      Left            =   2760
      TabIndex        =   21
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "KA"
      Height          =   345
      Left            =   3720
      TabIndex        =   20
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TK"
      Height          =   345
      Left            =   4800
      TabIndex        =   19
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label LBKelas 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   3840
      TabIndex        =   18
      Top             =   1080
      Width           =   1965
   End
   Begin VB.Label LBMI 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   3240
      TabIndex        =   17
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label LBKA 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   4200
      TabIndex        =   16
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label LBJurusan 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   2760
      TabIndex        =   15
      Top             =   720
      Width           =   3045
   End
   Begin VB.Label LBTK 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   5280
      TabIndex        =   14
      Top             =   2160
      Width           =   495
   End
End
Attribute VB_Name = "Pendaftaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_activate()
Call BukaDB
'Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\database\DBSPMB.mdb"
Adodc1.ConnectionString = pathdata
Adodc1.RecordSource = "pendaftaran"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
Call JumlahMI
Call JumlahKA
Call JumlahTK
End Sub

Private Sub Form_Load()
'Pendaftaran.Caption = "Operator : " & Login.TxtNamaOpr & " >>> Daftar Ulang <<<"
'Pendaftaran.KodeOpr = Login.TxtKodeOpr
TglDaftar.Caption = Date
Call BukaDB
Call KondisiAwal
TNIM.MaxLength = 7
TTelepon.MaxLength = 15
Call KondisiAwal
Call ListJurusan
Call ListGender
Call ListAgama
'RSMHS.Open "SELECT * FROM MAHASISWA", Conn
'Combo1.Clear
'Do While Not RSMHS.EOF
'    Combo1.AddItem RSMHS!nim & vbTab & RSMHS!namamhs
'    RSMHS.MoveNext
'Loop
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
    End If
    
    TNIM.Enabled = False
    If CBJurusan <> "MI" And CBJurusan <> "KA" And CBJurusan <> "TK" Then
        MsgBox ("Jurusan tidak terdaftar, harusnya MI, KA atau TK")
        CBJurusan.SetFocus
        Exit Sub
    End If
    TNama.SetFocus
End If
End Sub

Private Sub CBAgama_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If CBAgama <> "ISLAM" And CBAgama <> "KRISTEN" And CBAgama <> "HINDU" And CBAgama <> "BUDHA" Then
        MsgBox ("agama tidak terdaftar, harusnya ISLAM, KRISTEN, HINDU atau BUDHA")
        CBAgama.SetFocus
        Exit Sub
    End If
    TTelepon.SetFocus
End If
End Sub

Private Sub CBGender_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If CBGender <> "PRIA" And CBGender <> "WANITA" Then
        MsgBox ("Gender harus PRIA atau WANITA")
        CBGender.SetFocus
        Exit Sub
    End If
    CBAgama.SetFocus
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
    If CBJurusan = "" Or TNIM = "" Or TNama = "" Or LBKelas = "" Or TanggalLhr = "" Or TempatLhr = "" Or CBGender = "" Or CBAgama = "" Or TTelepon = "" Or TOrtu = "" Or TAlamat = "" Or TKota = "" Then
        MsgBox "Data belum lengkap"
        Exit Sub
    Else
        Dim aa As String
        aa = "insert into Pendaftaran(Tanggal,NIM,NamaMhs,Jurusan,Kelas,TanggalLhr,TempatLhr,Gender,Agama,Telepon,Ortu,Alamat,Kota,kodeopr) values ('" & TglDaftar & "','" & TNIM & "','" & TNama & "','" & LBJurusan & "','" & LBKelas & "','" & CDate(TanggalLhr) & "','" & TempatLhr & "','" & CBGender & "','" & CBAgama & "','" & TTelepon & "','" & TOrtu & "','" & TAlamat & "','" & TKota & "','" & Menu.STBAR.Panels(1).Text & "')"
        Conn.Execute aa
        Adodc1.Refresh
        DataGrid1.Refresh
        
        Dim bb As String
        bb = "insert into mahasiswa(NIM,NamaMhs,Jurusan,Kelas,Semester) values ('" & TNIM & "','" & TNama & "','" & LBJurusan & "','" & LBKelas & "',1)"
        Conn.Execute bb
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
    If TNama = "" Or TanggalLhr = "" Or TempatLhr = "" Or CBGender = "" Or CBAgama = "" Or TTelepon = "" Or TOrtu = "" Or TAlamat = "" Or TKota = "" Then
        MsgBox "Data belum lengkap"
        Exit Sub
    Else
        Dim cc As String
        cc = "Update pendaftaran set Namamhs='" & TNama & "',TanggalLhr='" & TanggalLhr & "',TempatLhr='" & TempatLhr & "',Gender='" & CBGender & "',agama='" & CBAgama & "',telepon='" & TTelepon & "',ortu='" & TOrtu & "',alamat='" & TAlamat & "',kota='" & TKota & "',kodeopr='" & Menu.STBAR.Panels(1).Text & "' where nim='" & TNIM & "'"
        Conn.Execute cc
        
        Dim DD As String
        DD = "Update mahasiswa set Namamhs='" & TNama & "' where nim='" & TNIM & "'"
        Conn.Execute DD
        Call KondisiAwal
        Adodc1.Refresh
        DataGrid1.Refresh
        Command2.SetFocus
    End If
End If
End Sub

'Private Sub Command3_Click()
'If Command3.Caption = "Hapus" Then
'    Command1.Enabled = False
'    Command2.Enabled = False
'    Command3.Caption = "Hapus"
'    Command4.Caption = "Batal"
'    TNIM.Enabled = True
'    TNIM.SetFocus
'End If
'End Sub

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

Private Sub Nim_OTO_MI()
Call BukaDB
Dim RS As New ADODB.Recordset
RS.Open "select NIM from Pendaftaran where Jurusan='MANAJEMEN INFORMATIKA' order by nim desc", Conn
RS.Requery
If RS.EOF Then
    Urutan = Right(TglDaftar, 2) + "01" + "001"
    TNIM = Urutan
    Exit Sub
Else
    Hitung = Right(RS!nim, 3) + 1
    Urutan = Right(TglDaftar, 2) + "01" + Right("000" & Hitung, 3)
End If
TNIM = Urutan
End Sub

Sub Nim_OTO_KA()
Call BukaDB
Dim RS As New ADODB.Recordset
RS.Open "select NIM from Pendaftaran where Jurusan='KOMPUTER AKUNTANSI' order by nim desc", Conn
RS.Requery
If RS.EOF Then
    Urutan = Right(TglDaftar, 2) + "02" + "001"
    TNIM = Urutan
Else
    Hitung = Right(RS!nim, 3) + 1
    Urutan = Right(TglDaftar, 2) + "02" + Right("000" & Hitung, 3)
End If
TNIM = Urutan
End Sub

Sub Nim_OTO_TK()
Call BukaDB
Dim RS As New ADODB.Recordset
RS.Open "select NIM from Pendaftaran where Jurusan='TEKNIK KOMPUTER' order by nim desc", Conn
RS.Requery
If RS.EOF Then
    Urutan = Right(TglDaftar, 2) + "03" + "001"
    TNIM = Urutan
Else
    Hitung = Right(RS!nim, 3) + 1
    Urutan = Right(TglDaftar, 2) + "03" + Right("000" & Hitung, 3)
End If
TNIM = Urutan
End Sub

Function JumlahMI()
Dim RS As New ADODB.Recordset
RS.Open "select count(NIM) as JMLMI from Pendaftaran where jurusan='MANAJEMEN INFORMATIKA'", Conn
LBMI = RS!JMLMI
End Function

Function JumlahKA()
Dim RS As New ADODB.Recordset
 RS.Open "select count(NIM) as JMLKA from Pendaftaran where jurusan='KOMPUTER AKUNTANSI'", Conn
LBKA = RS!JMLKA
End Function

Function JumlahTK()
Dim RS As New ADODB.Recordset
RS.Open "select count(NIM) as JMLTK from Pendaftaran where jurusan='TEKNIK KOMPUTER'", Conn
LBTK = RS!JMLTK
End Function

Sub ListJurusan()
CBJurusan.AddItem ("MI")
CBJurusan.AddItem ("KA")
CBJurusan.AddItem ("TK")
End Sub

Sub ListGender()
CBGender.AddItem ("PRIA")
CBGender.AddItem ("WANITA")
End Sub

Sub ListAgama()
CBAgama.AddItem ("ISLAM")
CBAgama.AddItem ("KRISTEN")
CBAgama.AddItem ("HINDU")
CBAgama.AddItem ("BUDHA")
End Sub

Sub KondisiAwal()

'Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\database\DBSPMB.mdb"
Adodc1.ConnectionString = pathdata
Adodc1.RecordSource = "pendaftaran"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
Call Gelap
Call Kosongkan
Call JumlahMI
Call JumlahKA
Call JumlahTK
Command1.Caption = "&Input"
Command2.Caption = "&Edit"
Command3.Caption = "Hapus"
Command4.Caption = "&Tutup"
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
End Sub

Sub Tampilkan()
With RSPendaftaran
    CBJurusan = Left(!Kelas, 2)
    TNama = !namamhs
    LBKelas = !Kelas
    LBJurusan = !Jurusan
    TempatLhr = !TempatLhr
    TanggalLhr = !TanggalLhr
    CBGender = !Gender
    CBAgama = !Agama
    TTelepon = !Telepon
    TOrtu = !Ortu
    TAlamat = !Alamat
    TKota = !KOTA
End With
End Sub

Private Sub TAlamat_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then TKota.SetFocus
End Sub

Private Sub TKota_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    If Command1.Caption = "Simpan" Then Command1.SetFocus
    If Command2.Caption = "Simpan" Then Command2.SetFocus
End If
End Sub

Private Sub TNama_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then TempatLhr.SetFocus
End Sub

Private Sub TNIM_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Len(TNIM) < 7 Then
        MsgBox "NIM harus 7 digit"
        TNIM.SetFocus
        Exit Sub
    End If
    
    'untuk &Input
    If Command1.Caption = "Simpan" Then
        Call CariNIM
            If Not RSPendaftaran.EOF Then
                Gelap
                Tampilkan
                MsgBox "Nomor Pendaftaran Sudah Ada"
                Kosongkan
                Terang
                TNIM.SetFocus
            Else
                Terang
                Gelap
                TNama.SetFocus
            End If
    
    'untuk &Edit
    ElseIf Command2.Caption = "Simpan" Then
            Call CariNIM
            If Not RSPendaftaran.EOF Then
                Tampilkan
                Terang
                TNIM.Enabled = False
                TNama.SetFocus
            Else
                MsgBox "Nomor Pendaftaran Tidak Ditemukan"
                Kosongkan
                Terang
                TNIM.SetFocus
            End If
        
    'untuk hapus
'    ElseIf Command3.Caption = "Hapus" Then
'        With RSPendaftaran
'            Call CariNIM
'            If Not rspendaftaran.eof Then
'                Tampilkan
'                Gelap
'                pesan = MsgBox("Yakin Data Ini Akan Dihapus...?", vbYesNo)
'                If pesan = vbYes Then
'                    .Delete
'                    Dim HapusMhs As String
'                    HapusMhs = "delete * from mahasiswa where nim='" & TNIM & "'"
'                    Conn.Execute (HapusMhs)
'                    Adodc1.Refresh
'                    DataGrid1.Refresh
'                    KondisiAwal
'                    Command3.SetFocus
'                Else
'                    KondisiAwal
'                    Command3.SetFocus
'                End If
'            Else
'                MsgBox "Nomor Formulir Tidak Ditemukan"
'                Kosongkan
'                Terang
'                TNIM.SetFocus
'            End If
'        End With
   End If
End If
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub TOrtu_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then TAlamat.SetFocus
End Sub

Private Sub TTelepon_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then TOrtu.SetFocus
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub TanggalLhr_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CBGender.SetFocus
End Sub

Private Sub TempatLhr_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then TanggalLhr.SetFocus
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
RSPendaftaran.Open "Select * From pendaftaran where NIM='" & TNIM & "'", Conn
End Sub


