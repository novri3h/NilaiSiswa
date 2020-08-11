VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Formulir 
   Caption         =   "Penjualan Formulir"
   ClientHeight    =   2145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   ScaleHeight     =   2145
   ScaleWidth      =   5430
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Formulir.frx":0000
      Height          =   1935
      Left            =   120
      TabIndex        =   15
      Top             =   2160
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3413
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   2
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "NOMOR"
         Caption         =   "Nomor"
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
         DataField       =   "TANGGAL"
         Caption         =   "Tanggal"
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
         DataField       =   "NAMA"
         Caption         =   "Nama Calom Mahasiswa"
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
         DataField       =   "ALAMAT"
         Caption         =   "Alamat"
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
      BeginProperty Column04 
         DataField       =   "TELEPON"
         Caption         =   "Telepon"
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
      BeginProperty Column05 
         DataField       =   "HARGA"
         Caption         =   "Harga"
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
      BeginProperty Column06 
         DataField       =   "KODEOPR"
         Caption         =   "Kode Opr"
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
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   345
      Left            =   120
      Top             =   4200
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   609
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
   Begin VB.CommandButton Command3 
      Caption         =   "&Tutup"
      Height          =   375
      Left            =   2760
      TabIndex        =   14
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   1440
      TabIndex        =   13
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Input"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Telepon 
      Height          =   350
      Left            =   1440
      TabIndex        =   10
      Top             =   1200
      Width           =   1250
   End
   Begin VB.TextBox Alamat 
      Height          =   350
      Left            =   1440
      TabIndex        =   9
      Top             =   840
      Width           =   3900
   End
   Begin VB.TextBox Nama 
      Height          =   350
      Left            =   1440
      TabIndex        =   8
      Top             =   480
      Width           =   3900
   End
   Begin VB.TextBox Nomor 
      Height          =   350
      Left            =   1440
      TabIndex        =   7
      Top             =   120
      Width           =   1250
   End
   Begin VB.Label Harga 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   4080
      TabIndex        =   11
      Top             =   1200
      Width           =   1245
   End
   Begin VB.Label Tanggal 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   4080
      TabIndex        =   6
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Harga"
      Height          =   345
      Left            =   2760
      TabIndex        =   5
      Top             =   1200
      Width           =   1245
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Telepon"
      Height          =   345
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1245
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Alamat"
      Height          =   345
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1245
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama"
      Height          =   345
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1245
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nomor"
      Height          =   345
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tanggal"
      Height          =   345
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   1245
   End
End
Attribute VB_Name = "Formulir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_activate()
Call BukaDB
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBSPMB.mdb"
Adodc1.RecordSource = "Formulir"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
End Sub

Sub Form_Load()
'Formulir.Caption = "Operator : " & Login.TxtNamaOpr & " >>> Penjualan Formulir <<<"
'Formulir.KodeOpr = Login.TxtKodeOpr
Tanggal = Format(Date, "DD-MM-YYYY")
Harga.Caption = 80000
Nomor.MaxLength = 4
Nama.MaxLength = 30
Alamat.MaxLength = 40
Telepon.MaxLength = 8
KondisiAwal
End Sub

Private Sub AutoNomor()
Call BukaDB
RSFormulir.Open ("select * from Formulir Where Nomor In(Select Max(Nomor)From Formulir)Order By Nomor Desc"), Conn
RSFormulir.Requery
    'Dim Urutan As String * 9
    Dim Hitung As Long
    With RSFormulir
        If .EOF Then
            'Urutan = Format(Date, "YYMMDD") + "001"
            Urutan = "0001"
            Nomor = Urutan
            Exit Sub
        Else
            'If Left(Nomor, 6) <> Format(Date, "YYMMDD") Then
            '    Urutan = Format(Date, "YYMMDD") + "001"
            'Else
                Hitung = !Nomor + 1
                Urutan = Right("0000" & Hitung, 4)
            End If
        'End If
        Nomor = Urutan
    End With
End Sub


'Private Sub AutoNomor()
'Call BukaDB
'RSFormulir.Open ("select * from Formulir Where Nomor In(Select Max(Nomor)From Formulir)Order By Nomor Desc"), Conn
'RSFormulir.Requery
'    Dim Urutan As String * 8
'    Dim Hitung As Integer
'    With RSFormulir
'        If .EOF Then
'            Urutan = "0001"
'            Nomor = Urutan
'        Else
'            Hitung = Right(!Nomor, 4) + 1
'            Urutan = Right("0000" & Hitung, 4)
'        End If
'        Nomor = Urutan
'    End With
'End Sub

Private Sub KosongkanText()
    Nomor = ""
    Nama = ""
    Alamat = ""
    Telepon = ""
End Sub

Private Sub SiapIsi()
    Nomor.Enabled = True
    Nama.Enabled = True
    Alamat.Enabled = True
    Telepon.Enabled = True
End Sub

Private Sub TidakSiapIsi()
    Nomor.Enabled = False
    Nama.Enabled = False
    Alamat.Enabled = False
    Telepon.Enabled = False
End Sub

Private Sub KondisiAwal()
    KosongkanText
    TidakSiapIsi
    Command1.Caption = "&Input"
    Command2.Caption = "&Edit"
    Command3.Caption = "&Tutup"
    Command1.Enabled = True
    Command2.Enabled = True
End Sub

Private Sub TampilkanData()
    With RSFormulir
        If Not RSFormulir.EOF Then
            Tanggal = RSFormulir!Tanggal
            Nama = RSFormulir!Nama
            Alamat = RSFormulir!Alamat
            Telepon = RSFormulir!Telepon
            Harga = RSFormulir!Harga
        End If
    End With
End Sub

Private Sub Command1_Click()
    If Command1.Caption = "&Input" Then
        Command1.Caption = "&Simpan"
        Command2.Enabled = False
        Command3.Caption = "&Batal"
        SiapIsi
        KosongkanText
        Call AutoNomor
        Nomor.Enabled = False
        Nama.SetFocus
    Else
        If Nomor = "" Or Nama = "" Or Alamat = "" Or Telepon = "" Or Harga = "" Then
            MsgBox "Data Belum Lengkap...!"
        Else
            Dim SQLTambah As String
            SQLTambah = "Insert Into Formulir (Tanggal,Nomor,Nama,alamat,Telepon,Harga,kodeopr) values ('" & Tanggal & "','" & Nomor & "','" & Nama & "','" & Alamat & "','" & Telepon & "','" & Harga & "','" & Menu.STBAR.Panel(1).Text & "')"
            Conn.Execute SQLTambah
            Form_activate
            KondisiAwal
        End If
    End If
End Sub

Private Sub Command2_Click()
    If Command2.Caption = "&Edit" Then
        Command1.Enabled = False
        Command2.Caption = "&Simpan"
        Command3.Caption = "&Batal"
        SiapIsi
        Nomor.SetFocus
    Else
        If Nama = "" Or Alamat = "" Or Telepon = "" Or Harga = "" Then
            MsgBox "Masih Ada Data Yang Kosong"
        Else
            Dim SQLEdit As String
            SQLEdit = "Update Formulir Set Nama= '" & Nama & "', alamat='" & Alamat & "', Telepon='" & Telepon & "',Harga='" & Harga & "',kodeopr ='" & Menu.STBAR.Panels(1).Text & "' where Nomor='" & Nomor & "'"
            Conn.Execute SQLEdit
            Form_activate
            KondisiAwal
        End If
    End If
End Sub

'Private Sub CmdHapus_Click()
'    If Cmdhapus.Caption = "&Hapus" Then
'        Command1.Enabled = False
'        Command2.Enabled = False
'        Command3.Caption = "&Batal"
'        KosongkanText
'        SiapIsi
'        Nomor.SetFocus
'    End If
'End Sub

Private Sub Command3_Click()
    Select Case Command3.Caption
        Case "&Tutup"
            Unload Me
        Case "&Batal"
            TidakSiapIsi
            KondisiAwal
    End Select
End Sub

Function CariData()
    Call BukaDB
    RSFormulir.Open "Select * From Formulir where Nomor='" & Nomor & "'", Conn
End Function

Private Sub Nomor_Keypress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    If Len(Nomor) < 4 Then
        MsgBox "Kode Harus 4 Digit"
        Nomor.SetFocus
        Exit Sub
    Else
        Nama.SetFocus
    End If

    If Command1.Caption = "&Simpan" Then
        Call CariData
        If Not RSFormulir.EOF Then
            TampilkanData
            MsgBox "Kode Formulir Sudah Ada"
            KosongkanText
            Nomor.SetFocus
        Else
            Nama.SetFocus
        End If
    End If
    
    If Command2.Caption = "&Simpan" Then
        Call CariData
        If Not RSFormulir.EOF Then
            TampilkanData
            Nomor.Enabled = False
            Nama.SetFocus
        Else
            MsgBox "Kode Formulir Tidak Ada"
            Nomor = ""
            Nomor.SetFocus
        End If
    End If
    
'    If Cmdhapus.Enabled = True Then
'        Call CariData
'        If Not RSFormulir.EOF Then
'            TampilkanData
'            Pesan = MsgBox("Yakin akan dihapus", vbYesNo)
'            If Pesan = vbYes Then
'                Dim SQLHapus As String
'                SQLHapus = "Delete From Formulir where Nomor= '" & Nomor & "'"
'                Conn.Execute SQLHapus
'                Form_activate
'                Kondisiawal
'            Else
'                Form_activate
'                Kondisiawal
'            End If
'        Else
'            MsgBox "Data Tidak ditemukan"
'            Nomor.SetFocus
'        End If
'    End If
End If
End Sub

Private Sub Nama_Keypress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Alamat.SetFocus
End Sub

Private Sub alamat_Keypress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Telepon.SetFocus
    'If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub telepon_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Command1.Enabled = True Then
            Command1.SetFocus
        ElseIf Command2.Enabled = True Then
            Command2.SetFocus
        End If
    End If
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

'Private Sub Harga_Keypress(Keyascii As Integer)
'    If Keyascii = 13 Then
'        If Command1.Enabled = True Then
'            Command1.SetFocus
'        ElseIf Command2.Enabled = True Then
'            Command2.SetFocus
'        End If
'    End If
'    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
'End Sub


