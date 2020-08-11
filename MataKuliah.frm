VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form MataKuliah 
   Caption         =   "Data Mata Kuliah"
   ClientHeight    =   3885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5385
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
   ScaleHeight     =   3885
   ScaleWidth      =   5385
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "MataKuliah.frx":0000
      Height          =   1900
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   3360
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
         DataField       =   "KodeMK"
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
      BeginProperty Column01 
         DataField       =   "NamaMK"
         Caption         =   "Nama Mata Kuliah"
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
      BeginProperty Column03 
         DataField       =   "SMT"
         Caption         =   "SMT"
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
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2745,071
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   494,929
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   494,929
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   350
      Left            =   3360
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
   Begin VB.CommandButton Command4 
      Caption         =   "&Tutup"
      Height          =   350
      Left            =   3360
      TabIndex        =   3
      Top             =   1320
      Width           =   1000
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Hapus"
      Height          =   350
      Left            =   2280
      TabIndex        =   2
      Top             =   1320
      Width           =   1000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Edit"
      Height          =   350
      Left            =   1200
      TabIndex        =   1
      Top             =   1320
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Input"
      Height          =   350
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   1000
   End
   Begin VB.TextBox Text4 
      Height          =   350
      Left            =   4200
      TabIndex        =   7
      Top             =   840
      Width           =   1000
   End
   Begin VB.TextBox Text3 
      Height          =   350
      Left            =   1200
      TabIndex        =   6
      Top             =   840
      Width           =   1000
   End
   Begin VB.TextBox Text2 
      Height          =   350
      Left            =   1200
      TabIndex        =   5
      Top             =   480
      Width           =   4000
   End
   Begin VB.TextBox Text1 
      Height          =   350
      Left            =   1200
      TabIndex        =   4
      Top             =   120
      Width           =   1000
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SMT"
      Height          =   345
      Left            =   3120
      TabIndex        =   11
      Top             =   840
      Width           =   1005
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SKS"
      Height          =   345
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   1005
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama"
      Height          =   345
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   1005
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode"
      Height          =   345
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1005
   End
End
Attribute VB_Name = "MataKuliah"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_activate()
Call BukaDB
Adodc1.ConnectionString = pathdata
Adodc1.RecordSource = "matakuliah"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
End Sub

Private Sub Form_Load()
Call BukaDB
Text1.MaxLength = 4
Text2.MaxLength = 30
Text3.MaxLength = 1
Text4.MaxLength = 1
KondisiAwal
End Sub

Private Sub Command1_Click()
If Command1.Caption = "&Input" Then
    Command1.Caption = "&Simpan"
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Caption = "&Batal"
    Terang
    Text1.Enabled = True
    Text1.SetFocus
Else
     If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
        MsgBox "Data Belum Lengkap...!"
        Exit Sub
    Else
        Dim Tambah As String
        Tambah = "insert into matakuliah (kodemk,namamk,sks,smt) values " & _
        "('" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Text4 & "')"
        Conn.Execute Tambah
        Form_activate
        Kosong
        Gelap
        KondisiAwal
    End If
End If
End Sub

Private Sub Command2_Click()
If Command2.Caption = "&Edit" Then
    Command1.Enabled = False
    Command2.Caption = "&Simpan"
    Command3.Enabled = False
    Command4.Caption = "&Batal"
    Terang
    Text1.Enabled = True
    Text1.SetFocus
Else
    If Text2 = "" Or Text3 = "" Or Text4 = "" Then
        MsgBox "Data Belum Lengkap...!"
        Exit Sub
    Else
        Dim edit As String
        edit = "update matakuliah set namamk='" & Text2 & "',sks='" & Text3 & "',smt='" & Text4 & "' where kodemk='" & Text1 & "'"
        Conn.Execute edit
        Form_activate
        Kosong
        Gelap
        KondisiAwal
    End If
End If
End Sub

Private Sub Command3_Click()
If Command3.Caption = "&Hapus" Then
    Command1.Enabled = False
    Command2.Enabled = False
    Command4.Caption = "&Batal"
    Text1.Enabled = True
    Text1.SetFocus
End If
End Sub

Private Sub Command4_Click()
Select Case Command4.Caption
    Case "&Tutup"
        Unload Me
    Case "&Batal"
        Kosong
        Gelap
        KondisiAwal
    End Select
End Sub

Sub CariMK()
Call BukaDB
RSMTKL.Open "select * from matakuliah where KodeMK='" & Text1 & "'", Conn
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Len(Text1) = 4 Then
        'With RSMTKL
        If Command1.Caption = "&Simpan" Then
            Call CariMK
            If Not RSMTKL.EOF Then
                Tampilkan
                MsgBox "Kode MataKuliah Sudah Terdaftar"
                Kosong
                Text1.SetFocus
            Else
                Text4 = Mid(Text1, 2, 1)
                Text2.SetFocus
            End If
        
        ElseIf Command2.Caption = "&Simpan" Then
            Call CariMK
            If Not RSMTKL.EOF Then
                Text1.Enabled = False
                Tampilkan
                Terang
                Text2.SetFocus
            Else
                MsgBox "Kode MataKuliah Tidak Terdaftar"
                Text1.SetFocus
            End If
            
        ElseIf Command3.Caption = "&Hapus" Then
            Call CariMK
            If Not RSMTKL.EOF Then
                Tampilkan
                Pesan = MsgBox("Yakin data akan di hapus", vbYesNo, "Konfirmasi")
                If Pesan = vbYes Then
                    RSMTKL.Delete
                    Adodc1.Refresh
                    DataGrid1.Refresh
                    KondisiAwal
                Else
                    KondisiAwal
                End If
            Else
                Pesan = MsgBox("Kode MataKuliah Tidak Terdaftar")
                Text1.SetFocus
            End If
        End If
        
    End If
    End If

If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then Text3.SetFocus
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text4.SetFocus
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub Text4_keypress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Command1.Caption = "&Simpan" Then
        Command1.SetFocus
    ElseIf Command2.Caption = "&Simpan" Then
        Command2.SetFocus
    End If
End If
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub Kosong()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
End Sub

Private Sub Terang()
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
End Sub

Private Sub Gelap()
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
End Sub

Private Sub Tampilkan()
With RSMTKL
    Text2 = !namamk
    Text3 = !SKS
    Text4 = !smt
End With
End Sub

Private Sub KondisiAwal()
Command1.Caption = "&Input"
Command2.Caption = "&Edit"
Command3.Caption = "&Hapus"
Command4.Caption = "&Tutup"
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Text1.Enabled = False
Gelap
Kosong
End Sub

