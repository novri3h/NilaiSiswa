VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form OlahNilai1 
   Caption         =   "Pengolahan Nilai"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   6780
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   350
      Left            =   1800
      TabIndex        =   11
      Top             =   3720
      Width           =   750
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   840
      Width           =   5000
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   5000
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   5000
   End
   Begin VB.CommandButton CmdSimpanData 
      Caption         =   "Simpan"
      Height          =   350
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   750
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "OlahNilai1.frx":0000
      Height          =   2175
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   3836
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
         DataField       =   "NIM"
         Caption         =   "NIM"
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
         DataField       =   "NamaMhs"
         Caption         =   "Nama Mahasiswa"
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
         DataField       =   "Kelas"
         Caption         =   "Kelas"
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
         DataField       =   "Absen"
         Caption         =   "Absen"
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
         DataField       =   "Tugas"
         Caption         =   "Tugas"
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
         DataField       =   "UTS"
         Caption         =   "UTS"
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
         DataField       =   "UAS"
         Caption         =   "UAS"
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
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   750.047
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2760
      Top             =   3720
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
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
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   350
      Left            =   960
      TabIndex        =   4
      Top             =   3720
      Width           =   750
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kelas"
      Height          =   345
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   1500
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mata Kuliah"
      Height          =   345
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   1500
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Dosen"
      Height          =   350
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label LBJumlah 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   5640
      TabIndex        =   7
      Top             =   3720
      Width           =   750
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Jumlah"
      Height          =   345
      Left            =   4800
      TabIndex        =   6
      Top             =   3720
      Width           =   750
   End
End
Attribute VB_Name = "OlahNilai1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdBatal_Click()
Combo1 = ""
Combo2 = ""
Combo3 = ""
Call KosongkanGrid
Combo1.SetFocus
End Sub

Private Sub Combo1_Click()
Call BukaDB
Dim CariMatKul As New ADODB.Recordset
CariMatKul.Open "select DETAILDOSEN.KODEMK,MATAKULIAH.NAMAMK from DETAILDOSEN,matakuliah where matakuliah.kodemk=detaildosen.kodemk and detaildosen.kodedsn='" & Left(Combo1, 3) & "'", Conn
CariMatKul.MoveFirst
Combo2.Clear
Do While Not CariMatKul.EOF
    Combo2.AddItem CariMatKul!kodemk & Space(5) & CariMatKul!namamk
    CariMatKul.MoveNext
Loop
End Sub

Private Sub Combo2_Click()
Call BukaDB
Dim CariKelas As New ADODB.Recordset
CariKelas.Open "select kelas from detaildosen where kodedsn='" & Left(Combo1, 3) & "' and kodemk='" & Left(Combo2, 4) & " '", Conn

Combo3.Clear
Do While Not CariKelas.EOF
    Combo3.AddItem CariKelas!kelas
    CariKelas.MoveNext
Loop
End Sub

Private Sub Combo3_Click()
Call BukaDB
Dim TampilSiswa As New ADODB.Recordset
TampilSiswa.Open "Select NamaMK, Jurusan From matakuliah,mahasiswa Where Kodemk='" & Trim(Left(Combo2, 4)) & "' And kelas='" & Combo3 & "'", Conn

Adodc1.RecordSource = "Select Nim,NamaMhs,kelas,kodemk,Absen,Tugas,UTS,UAS From Nilai Where Kodemk='" & Trim(Left(Combo2, 4)) & "' And Kelas='" & Combo3 & "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
    Adodc1.RecordSource = "Select Nim,NamaMhs,Absen,Tugas,UTS,UAS From transNilai Where Kelas='" & Combo3 & "'"
    Adodc1.Refresh
End If
LBJumlah = Adodc1.Recordset.RecordCount

End Sub


Private Sub Form_activate()
Call BukaDB
Adodc1.ConnectionString = pathdata
Call KosongkanNilai
Adodc1.RecordSource = "select * from transnilai Where nim='xxx'"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
End Sub

Private Sub Form_Load()
Call BukaDB

RSDosen.Open "SELECT * from dosen", Conn
RSDosen.MoveFirst
Combo1.Clear
Do While Not RSDosen.EOF
    Combo1.AddItem RSDosen!kodedsn & Space(5) & RSDosen!NamaDsn
    RSDosen.MoveNext
Loop
Call Semula
End Sub

Private Sub CmdSimpanData_Click()
If Combo1 = "" Or Combo2 = "" Or Combo3 = "" Or LBJumlah = "" Then
    MsgBox "Data belum lengkap"
    Exit Sub
End If

Call BukaDB
RSNilai.Open "select * from nilai where kodemk='" & Trim(Left(Combo2, 4)) & "' and kelas ='" & Combo3 & "'", Conn
If RSNilai.EOF Then
    Adodc1.Recordset.MoveFirst
    Do While Not Adodc1.Recordset.EOF
        Dim Simpan As String
        Simpan = "Insert Into Nilai(Kelas,KodeMK,NIM,namamhs,absen,tugas,uts,uas) values " & _
        "('" & Combo3 & "','" & Left(Combo2, 4) & "','" & Adodc1.Recordset!nim & "','" & Adodc1.Recordset!namamhs & "','" & Adodc1.Recordset!absen & "','" & Adodc1.Recordset!tugas & "','" & Adodc1.Recordset!uts & "','" & Adodc1.Recordset!uas & "')"
        Conn.Execute (Simpan)
        Call KosongkanNilai
        Adodc1.Recordset.MoveNext
    Loop
Else
    Adodc1.Recordset.MoveFirst
    Do While Not Adodc1.Recordset.EOF
    Dim edit As String
    edit = "Update nilai set absen='" & Adodc1.Recordset!absen & "',Tugas='" & Adodc1.Recordset!tugas & "',UTS='" & Adodc1.Recordset!uts & "',UAS='" & Adodc1.Recordset!uas & "' where nim='" & Adodc1.Recordset!nim & "' and kelas='" & Combo3 & "' and kodemk='" & Trim(Left(Combo2, 4)) & "'"
    Conn.Execute edit
    Adodc1.Recordset.MoveNext
    Loop
End If


Call Updating
Call Semula
Call KosongkanGrid
Call KosongkanNilai
Form_activate
End Sub


Private Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
If DataGrid1.Col = 3 Then
    If Adodc1.Recordset!absen > 100 Then
        MsgBox "Nilai maksimal adalah 100"
        Exit Sub
    Else
        Adodc1.Recordset.MoveNext
    End If
End If

If DataGrid1.Col = 4 Then
    If Adodc1.Recordset!tugas > 100 Then
        MsgBox "Nilai maksimal adalah 100"
        Exit Sub
    Else
        Adodc1.Recordset.MoveNext
    End If
End If

If DataGrid1.Col = 5 Then
    If Adodc1.Recordset!uts > 100 Then
        MsgBox "Nilai maksimal adalah 100"
        Exit Sub
    Else
        Adodc1.Recordset.MoveNext
    End If
End If

If DataGrid1.Col = 6 Then
    If Adodc1.Recordset!uas > 100 Then
        MsgBox "Nilai maksimal adalah 100"
        Exit Sub
    Else
        Adodc1.Recordset.MoveNext
    End If
End If
    
End Sub



Sub Semula()
CmdTutup.Caption = "&Tutup"
CmdTutup.Enabled = True

End Sub


Private Sub CmdTutup_Click()
Select Case CmdTutup.Caption
    Case "&Tutup"
        'End
        Unload Me
    Case "&Batal"
        Semula
        Call KosongkanNilai
        Call KosongkanGrid
End Select
End Sub

Private Sub datagrid1_Keypress(KeyAscii As Integer)
On Error Resume Next
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then KeyAscii = 0
End Sub

Sub GridEntri()
    Call KosongkanNilai
    Adodc1.RecordSource = "Select NIM,NamaMhs,absen,tugas,uts,uas  from TransNilai where kelas='" & Combo3 & "'"
    Adodc1.Refresh
    LBJumlah = Adodc1.Recordset.RecordCount
End Sub

Sub TampilNilai()
    Adodc1.RecordSource = "Select Nim,NamaMhs,Absen,Tugas,UTS,UAS From Nilai Where Kodemk='" & Trim(Left(Combo2, 4)) & "' And Kelas='" & Combo3 & "'"
    Adodc1.Refresh
    LBJumlah = Adodc1.Recordset.RecordCount
End Sub

Sub KosongkanGrid()
    Adodc1.RecordSource = "Select nim,namamhs From TransNilai Where kelas='" & Combo3 & "'"
    Adodc1.Refresh
    LBJumlah = ""
End Sub

Sub CariData()
    Adodc1.RecordSource = "Select Nim,NamaMhs,kelas,kodemk,Absen,Tugas,UTS,UAS From Nilai Where Kodemk='" & Trim(Left(Combo2, 4)) & "' And Kelas='" & Combo3 & "'"
    Adodc1.Refresh
    LBJumlah = Adodc1.Recordset.RecordCount
End Sub

Sub Updating()
'On Error Resume Next
    Dim RSNilai As New ADODB.Recordset
    RSNilai.Open "select * from Nilai", Conn
    If Not RSNilai.EOF Then
        Dim aa As String
        aa = "Update Nilai Set Total=(Absen*0.1) + (tugas* 0.2) + (uts*0.3) + (uas*0.4) where kodemk='" & Trim(Left(Combo2, 4)) & "' and kelas ='" & Combo3 & "'"
        Conn.Execute aa
        
        Dim bb As String
        bb = "Update Nilai Set Grade=iif (val(Total)=0,'E',iif(val(Total)>0 and val(Total)<60,'D',iif(val(Total)>=60 and val(Total)<75,'C',iif(val(Total)>=75 and val(Total)<85,'B','A')))) where kodemk='" & Trim(Left(Combo2, 4)) & "' and kelas ='" & Combo3 & "'"
        Conn.Execute bb
        
        Dim cc As String
        cc = "Update Nilai Set Ket=iif (Grade='E' or Grade='D','Her',iif(Grade='A','Memuaskan',iif(Grade='B','Baik','Cukup'))) where kodemk='" & Trim(Left(Combo2, 4)) & "' and kelas ='" & Combo3 & "'"
        Conn.Execute cc
    End If
    MsgBox "Penyimpanan dan Updating Data Sukses"
    Combo1 = "":    Combo2 = "":    Combo3 = ""
End Sub

Sub KosongkanNilai()
'On Error Resume Next
Dim Nolkan As String
Nolkan = "update transnilai set absen=0,tugas=0,uts=0,uas=0 where kelas ='" & Trim(Combo3.Text) & "'"
Conn.Execute Nolkan
End Sub

