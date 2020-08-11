VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Dosen 
   Caption         =   "Data Dosen"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6930
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
   ScaleHeight     =   6570
   ScaleWidth      =   6930
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   1455
      Left            =   120
      TabIndex        =   19
      Top             =   960
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2566
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   18
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
   Begin VB.TextBox Text6 
      Height          =   350
      Left            =   5400
      TabIndex        =   17
      Top             =   6120
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   350
      Left            =   4560
      TabIndex        =   13
      Top             =   6120
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Height          =   350
      Left            =   1320
      TabIndex        =   12
      Top             =   6120
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      Height          =   350
      Left            =   120
      TabIndex        =   11
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton CmdTambahData 
      Caption         =   "Tambah Data"
      Height          =   350
      Left            =   4080
      TabIndex        =   10
      Top             =   5400
      Width           =   1250
   End
   Begin VB.CommandButton CmdHapusData 
      Caption         =   "Hapus Data"
      Height          =   350
      Left            =   5400
      TabIndex        =   9
      Top             =   5400
      Width           =   1250
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Dosen.frx":0000
      Height          =   2775
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4895
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "Nomor"
         Caption         =   "Nomor"
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
         DataField       =   "Kode"
         Caption         =   "Kode"
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
         DataField       =   "Nama"
         Caption         =   "Nama"
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
         DataField       =   "SKS"
         Caption         =   "SKS"
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
         DataField       =   "Kelas"
         Caption         =   "Kelas"
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
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   2505,26
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column04 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Tutup"
      Height          =   350
      Left            =   2640
      TabIndex        =   3
      Top             =   5400
      Width           =   800
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Hapus"
      Height          =   350
      Left            =   1800
      TabIndex        =   2
      Top             =   5400
      Width           =   800
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Edit"
      Height          =   350
      Left            =   960
      TabIndex        =   1
      Top             =   5400
      Width           =   800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Input"
      Height          =   350
      Left            =   120
      TabIndex        =   0
      Top             =   5400
      Width           =   800
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   120
      Top             =   6720
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
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
   Begin VB.TextBox Text2 
      Height          =   350
      Left            =   1440
      TabIndex        =   5
      Top             =   480
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      Height          =   350
      Left            =   1440
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   2040
      Top             =   6720
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   3960
      Top             =   6720
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
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
   Begin MSDataGridLib.DataGrid DataGrid3 
      Height          =   1455
      Left            =   3360
      TabIndex        =   20
      Top             =   960
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   2566
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   18
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
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Kelas"
      Height          =   225
      Left            =   5520
      TabIndex        =   18
      Top             =   5880
      Width           =   435
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "SKS"
      Height          =   225
      Left            =   4560
      TabIndex        =   16
      Top             =   5880
      Width           =   345
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Nama Mata Kuliah"
      Height          =   225
      Left            =   1320
      TabIndex        =   15
      Top             =   5880
      Width           =   1470
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Kode"
      Height          =   225
      Left            =   120
      TabIndex        =   14
      Top             =   5880
      Width           =   405
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama Dosen"
      Height          =   345
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   1245
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode Dosen"
      Height          =   350
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1245
   End
End
Attribute VB_Name = "Dosen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_activate()
Call BukaDB
Adodc1.ConnectionString = pathdata
Adodc1.RecordSource = "TRDosen"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh

'==============
Adodc2.ConnectionString = pathdata
Adodc2.RecordSource = "dosen"
Adodc2.Refresh
Set DataGrid2.DataSource = Adodc2
DataGrid2.Refresh
'====================
'==============
Adodc3.ConnectionString = pathdata
Adodc3.RecordSource = "matakuliah"
Adodc3.Refresh
Set DataGrid3.DataSource = Adodc3
DataGrid3.Refresh
'====================

Call HapusTabel
Call TbhNmr
Adodc1.Recordset.MoveFirst
End Sub

Private Sub Form_Load()
'Call BukaDB
'RSDosen.Open "dosen", Conn
'Combo1.Clear
'Do While Not RSDosen.EOF
'    Combo1.AddItem RSDosen!kodedsn & Space(5) & RSDosen!namadsn
'    RSDosen.MoveNext
'Loop
Text1.MaxLength = 3
Text2.MaxLength = 30
KondisiAwal
End Sub

Sub ReffHapus()
Call BukaDB
Dim RSCari As New ADODB.Recordset
RSCari.Open "SELECT DETAILDOSEN.KODEMK,NAMAMK,SKS FROM DETAILDOSEN,DOSEN,MATAKULIAH WHERE DETAILDOSEN.KODEMK=MATAKULIAH.KODEMK AND dosen.kodedsn=detaildosen.kodedsn and DOSEN.KODEDSN='" & Text1 & "'", Conn
RSCari.Requery
Call HapusTabel
RSCari.MoveFirst
Nomor = 0
Do While Not RSCari.EOF
    Nomor = Nomor + 1
    Adodc1.Recordset.AddNew
    Adodc1.Recordset!Nomor = Nomor
    Adodc1.Recordset!Kode = RSCari!kodemk
    Adodc1.Recordset!Nama = RSCari!namamk
    Adodc1.Recordset!SKS = RSCari!SKS
    Adodc1.Recordset.Update
    RSCari.MoveNext
Loop
End Sub

Private Sub CmdHapusData_Click()
Text3.Enabled = True
Text3.SetFocus
CmdTambahData.Enabled = False
End Sub

Private Sub CmdTambahData_Click()
Text3.Enabled = True
Text3.SetFocus
CmdHapusData.Enabled = False
End Sub

Sub HapusTabel()
Call BukaDB
Dim Hapus As String
Hapus = "delete * from trdosen"
Conn.Execute Hapus
Call BukaDB
Adodc1.Refresh
DataGrid1.Refresh
End Sub

Sub TbhNmr()
For i = 1 To 10
    Adodc1.Recordset.AddNew
    Adodc1.Recordset!Nomor = i
    Adodc1.Recordset.Update
Next i
Adodc1.Refresh
DataGrid1.Refresh
End Sub

Function Tambah_Baris()
    For i = Adodc1.Recordset.RecordCount To Adodc1.Recordset.RecordCount
        Adodc1.Recordset.AddNew
        Adodc1.Recordset!Nomor = i + 1
        Adodc1.Recordset.Update
    Next i
End Function

Sub HapusTabel1()
On Error Resume Next
Call BukaDB
If Adodc1.Recordset.RecordCount <> 0 Then
    Adodc1.Recordset.MoveFirst
    Do While Not Adodc1.Recordset.EOF
        Adodc1.Recordset.Delete
        Adodc1.Recordset.MoveNext
    Loop
End If

End Sub

Private Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
If DataGrid1.Col = 1 Then
    Call BukaDB
    RSMTKL.Open "select * from matakuliah where kodemk='" & Adodc1.Recordset!Kode & "'", Conn
    RSMTKL.Requery
    If RSMTKL.EOF Then
        MsgBox "Kode mata kuliah tidak terdaftar"
        DataGrid1.Col = 1
        Exit Sub
    Else
        Adodc1.Recordset!Nama = RSMTKL!namamk
        Adodc1.Recordset!SKS = RSMTKL!SKS
        DataGrid1.Col = 4
        Exit Sub
    End If
    
End If

If DataGrid1.Col = 4 Then
    Call BukaDB
    RSMHS.Open "select distinct kelas from mahasiswa where kelas='" & Adodc1.Recordset!kelas & "'", Conn
    RSMHS.Requery
    If RSMHS.EOF Then
        MsgBox "Kelas tidak terdaftar"
        DataGrid1.Col = 4
        Exit Sub
    Else
        Adodc1.Recordset!kelas = RSMHS!kelas
        Adodc1.Recordset.Update
        Adodc1.Recordset.MoveNext
        DataGrid1.Col = 1
    End If
End If

End Sub

Private Sub Command1_Click()
If Command1.Caption = "&Input" Then
    Command1.Caption = "&Simpan"
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Caption = "&Batal"
    Text1.Enabled = True
    Text1.SetFocus
Else
     If Text1 = "" Or Text2 = "" Then
        MsgBox "Data Belum Lengkap...!"
        Exit Sub
    Else
        Dim Tambah As String
        Tambah = "insert into Dosen (KodeDsn,NamaDsn) values " & _
        "('" & Text1 & "','" & Text2 & "')"
        Conn.Execute Tambah
        
        Adodc1.Recordset.MoveFirst
        Do While Not Adodc1.Recordset.EOF
            If Adodc1.Recordset!Kode <> vbNullString Then
                Dim tambahdetail As String
                tambahdetail = "insert into detaildosen(kodedsn,kodemk,kelas) values " & _
                "('" & Text1 & "','" & Adodc1.Recordset!Kode & "','" & Adodc1.Recordset!kelas & "')"
                Conn.Execute tambahdetail
            End If
        Adodc1.Recordset.MoveNext
        Loop
        KondisiAwal
        Call HapusTabel
        Call TbhNmr
    End If
End If
End Sub

Private Sub Command2_Click()
If Command2.Caption = "&Edit" Then
    Command1.Enabled = False
    Command2.Caption = "&Simpan"
    Command3.Enabled = False
    Command4.Caption = "&Batal"
    Text1.Enabled = True
    Text1.SetFocus
Else
    If Text2 = "" Then
        MsgBox "Data Belum Lengkap...!"
        Exit Sub
    Else
        Dim edit As String
        edit = "update Dosen set NamaDsn='" & Text2 & "' where KodeDsn='" & Text1 & "'"
        Conn.Execute edit
        
        Dim HapusDulu As String
        HapusDulu = "delete * from detaildosen where kodedsn='" & Text1 & "'"
        Conn.Execute HapusDulu
                    
        Adodc1.Recordset.MoveFirst
        Do While Not Adodc1.Recordset.EOF
            If Adodc1.Recordset!Kode <> vbNullString Then
                Dim tambahdetail As String
                tambahdetail = "insert into detaildosen(kodedsn,kodemk,KELAS) values " & _
                "('" & Text1 & "','" & Adodc1.Recordset!Kode & "','" & Adodc1.Recordset!kelas & "')"
                Conn.Execute tambahdetail
            End If
        Adodc1.Recordset.MoveNext
        Loop
        Call HapusTabel
        Call TbhNmr
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
        Call HapusTabel
        Call TbhNmr
        KondisiAwal
    End Select
End Sub

Sub CariDosen()
Call BukaDB
RSDosen.Open "select * from Dosen where KodeDsn='" & Text1 & "'", Conn
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    If Command1.Caption = "&Simpan" Then
        Call CariDosen
        If Not RSDosen.EOF Then
            Text2 = RSDosen!namadsn
            MsgBox "Kode Dosen Sudah Terdaftar"
            Text1 = ""
            Text2 = ""
            Text1.SetFocus
        Else
            Text2.Enabled = True
            Text2.SetFocus
        End If
    
    ElseIf Command2.Caption = "&Simpan" Then
        Call CariDosen
        If Not RSDosen.EOF Then
            Text1.Enabled = False
            Text2.Enabled = True
            Text2 = RSDosen!namadsn
            Text2.SetFocus
            Call BukaDB
            Dim RSCari As New ADODB.Recordset
            RSCari.Open "SELECT DETAILDOSEN.KODEMK,NAMAMK,SKS,DETAILDOSEN.KELAS FROM DETAILDOSEN,DOSEN,MATAKULIAH WHERE DETAILDOSEN.KODEMK=MATAKULIAH.KODEMK AND dosen.kodedsn=detaildosen.kodedsn and DOSEN.KODEDSN='" & Text1 & "'", Conn
            RSCari.Requery
            Call HapusTabel
            Adodc1.Refresh
            DataGrid1.Refresh
            RSCari.MoveFirst
            Nomor = 0
            Do While Not RSCari.EOF
                Nomor = Nomor + 1
                Adodc1.Recordset.AddNew
                Adodc1.Recordset!Nomor = Nomor
                Adodc1.Recordset!Kode = RSCari!kodemk
                Adodc1.Recordset!Nama = RSCari!namamk
                Adodc1.Recordset!SKS = RSCari!SKS
                Adodc1.Recordset!kelas = RSCari!kelas
                Adodc1.Recordset.Update
                RSCari.MoveNext
            Loop
            CmdTambahData.Enabled = True
            CmdHapusData.Enabled = True
        Else
            MsgBox "Kode Dosen Tidak Terdaftar"
            Text1.SetFocus
        End If
        
    ElseIf Command3.Enabled = True Then
        Call CariDosen
        If Not RSDosen.EOF Then
            Text2 = RSDosen!namadsn
            Pesan = MsgBox("Yakin data akan di hapus", vbYesNo, "Konfirmasi")
            If Pesan = vbYes Then
                Dim Hapus As String
                Hapus = "delete * from dosen where kodedsn='" & Text1 & "'"
                Conn.Execute Hapus
                
                Dim hapusdetail As String
                hapusdetail = "delete * from detaildosen where kodedsn='" & Text1 & "'"
                Conn.Execute hapusdetail
                KondisiAwal
            Else
                KondisiAwal
            End If
        Else
            Pesan = MsgBox("Kode Dosen Tidak Terdaftar")
            Text1.SetFocus
        End If
    End If
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    DataGrid1.SetFocus
    DataGrid1.Col = 1
End If
End Sub

Private Sub KondisiAwal()
Command1.Caption = "&Input"
Command2.Caption = "&Edit"
Command3.Caption = "&Hapus"
Command4.Caption = "&Tutup"
Command1.Enabled = True: Command2.Enabled = True
Command3.Enabled = True: Command4.Enabled = True
CmdHapusData.Enabled = False
CmdTambahData.Enabled = False
Text1 = "": Text2 = ""
Text3 = "": Text4 = ""
Text5 = "": Text6 = ""
Text1.Enabled = False: Text2.Enabled = False
Text3.Enabled = False: Text4.Enabled = False
Text5.Enabled = False: Text6.Enabled = False

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If CmdTambahData.Enabled = True Then
        Call BukaDB
        RSMTKL.Open "select * from matakuliah where kodemk='" & Text3 & "'", Conn
        If Not RSMTKL.EOF Then
            Text4 = RSMTKL!namamk
            Text5 = RSMTKL!SKS
            Text6.Enabled = True
            Text6.SetFocus
        Else
            MsgBox "kode matakuliah tidak terdaftar"
            Text3.SetFocus
        End If
    
    ElseIf CmdHapusData.Enabled = True Then
        Call BukaDB
        Dim RSCari As New ADODB.Recordset
        RSCari.Open "select * from trdosen where kode='" & Text3 & "'", Conn
        If Not RSCari.EOF Then
            Pesan = MsgBox("yakin akan dihapus", vbYesNo)
            If Pesan = vbYes Then
                Dim Hapus As String
                Hapus = "delete * from detaildosen where kodemk='" & Text3 & "' and kodedsn='" & Text1 & "'"
                Conn.Execute Hapus
                Call HapusTabel
                Call ReffHapus
                Text3 = "": Text4 = "": Text5 = ""
            End If
        Else
            MsgBox "kode matakuliah tidak terdaftar"
            Text3.SetFocus
        End If
    End If
End If
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then KeyAscii = 0
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    If CmdTambahData.Enabled = True Then
        Call BukaDB
        RSMHS.Open "select DISTINCT KELAS from MAHASISWA where KELAS ='" & Text6 & "'", Conn
        If Not RSMHS.EOF Then
            Adodc1.Recordset.AddNew
            Adodc1.Recordset!Nomor = Adodc1.Recordset.RecordCount - 1 + 1
            Adodc1.Recordset!Kode = Text3
            Adodc1.Recordset!Nama = Text4
            Adodc1.Recordset!SKS = Text5
            Adodc1.Recordset!kelas = Text6
            Adodc1.Recordset.Update
            Text3 = "": Text4 = "": Text5 = "": Text6 = ""
            Text3.SetFocus
        Else
            MsgBox "kELAS tidak terdaftar"
            Text6.SetFocus
        End If
    End If
End If
End Sub
