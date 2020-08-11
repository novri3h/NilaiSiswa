VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form EntriNilaiHer 
   Caption         =   "Entri Nilai Her"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7875
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
   ScaleHeight     =   5430
   ScaleWidth      =   7875
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "EntriNilaiHer.frx":0000
      Height          =   4215
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   7435
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
      BeginProperty Column02 
         DataField       =   "Nama"
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
      BeginProperty Column03 
         DataField       =   "Nilai"
         Caption         =   "Nilai"
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
            ColumnWidth     =   494,929
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3000,189
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   750,047
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   5640
      Top             =   4920
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
   Begin VB.ListBox List1 
      Height          =   3660
      Left            =   6000
      TabIndex        =   8
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox TxtKodeMK 
      Height          =   350
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   885
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "Simpan"
      Height          =   400
      Left            =   960
      TabIndex        =   2
      Top             =   4920
      Width           =   1200
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "Batal"
      Height          =   400
      Left            =   2160
      TabIndex        =   3
      Top             =   4920
      Width           =   1200
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "Tutup"
      Height          =   400
      Left            =   3360
      TabIndex        =   4
      Top             =   4920
      Width           =   1200
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NIM"
      Height          =   345
      Left            =   6000
      TabIndex        =   9
      Top             =   600
      Width           =   1725
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kod MT Kuliah"
      Height          =   345
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1250
   End
   Begin VB.Label LblNamaMk 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   2400
      TabIndex        =   6
      Top             =   120
      Width           =   5325
   End
   Begin VB.Label LblJumlah 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   120
      TabIndex        =   5
      Top             =   4920
      Width           =   750
   End
End
Attribute VB_Name = "EntriNilaiHer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_activate()
Call BukaDB
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBnilai.mdb"
Adodc1.RecordSource = "TRNilaiHer"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh

Call TabelKosong
TxtKodeMK = ""
End Sub

Private Sub TxtKodeMK_KeyPress(KeyAscii As Integer)
TxtKodeMK.MaxLength = 4
If KeyAscii = 13 Then
    Call BukaDB
    RSMTKL.Open "select * from matakuliah where kodemk='" & TxtKodeMK & "'", Conn
    If RSMTKL.EOF Then
        MsgBox "Kode Mata Kuliah tidak terdaftar"
        TxtKodeMK.SetFocus
    Else
        LblNamaMk = Space(1) & RSMTKL!namamk
        DataGrid1.SetFocus
        DataGrid1.Col = 1
        
        Dim CariNIM As New ADODB.Recordset
        CariNIM.Open "SELECT DISTINCT MAHASISWA.NIM FROM MATAKULIAH,MAHASISWA,NILAI WHERE NILAI.KODEMK=MATAKULIAH.KODEMK AND MAHASISWA.NIM=NILAI.NIM AND NILAI.TOTAL<60 AND MATAKULIAH.KODEMK='" & TxtKodeMK & "'", Conn
        List1.Clear
        If Not CariNIM.EOF Then
            Do While Not CariNIM.EOF
                List1.AddItem CariNIM!nim
                CariNIM.MoveNext
            Loop
        Else
            MsgBox "Data tidak ditemukan"
            TxtKodeMK.SetFocus
            Exit Sub
        End If
        
    End If
End If
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub datagrid1_Keypress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then KeyAscii = 0
End Sub

Private Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
If DataGrid1.Col = 1 Then
    Call BukaDB
    RSMHS.Open "select * from mahasiswa where nim='" & Adodc1.Recordset!nim & "'", Conn
    RSMHS.Requery
    If RSMHS.EOF Then
        MsgBox "NIM Tidak Terdaftar"
        DataGrid1.Col = 1
        Exit Sub
    End If

    Adodc1.Recordset!nim = RSMHS!nim
    Adodc1.Recordset!Nama = RSMHS!namamhs
    DataGrid1.Col = 3
    DataGrid1.Refresh
    Exit Sub
End If

If DataGrid1.Col = 3 Then
    Adodc1.Recordset!Nilai = Adodc1.Recordset!Nilai
    Adodc1.Recordset.Update
    Adodc1.Recordset.MoveNext
    DataGrid1.Col = 1
    Call JmlData
End If
End Sub

Private Sub CmdSimpan_Click()
If LblJumlah.Caption = "" Or TxtKodeMK = "" Then
    MsgBox "Data Belum Lengkap"
    TxtKodeMK.SetFocus
    Exit Sub
Else
    Adodc1.Recordset.MoveFirst
    Do While Not Adodc1.Recordset.EOF
        If Adodc1.Recordset!nim <> vbNullString Then
            Call BukaDB
            RSNilaiHer.Open "Select * from nilaiher where kodemk='" & TxtKodeMK & "' and nim='" & Adodc1.Recordset!nim & "'", Conn
            If RSNilaiHer.EOF Then
                SQLTambah = "Insert Into NilaiHer(KodeMK,NIM,Nilai) values ('" & TxtKodeMK & "','" & Adodc1.Recordset!nim & "','" & Adodc1.Recordset!Nilai & "')"
                Conn.Execute (SQLTambah)
            Else
                sqlupdate = "Update nilaiher set nilai='" & Adodc1.Recordset!Nilai & "' where kodemk='" & TxtKodeMK & "' and nim='" & Adodc1.Recordset!nim & "'"
                Conn.Execute (sqlupdate)
            End If
        End If
    Adodc1.Recordset.MoveNext
    Loop
    
    Adodc1.Recordset.MoveFirst
    Do While Not Adodc1.Recordset.EOF
        'rasio nilai 30% nilai awal dan 70% nilai hasil her
        sqlupdate = "Update Nilai Set Total= (Total*0.3) + '" & Adodc1.Recordset!Nilai * 0.7 & "' where NIM='" & Adodc1.Recordset!nim & "' and Kodemk='" & TxtKodeMK & "'"
        Conn.Execute (sqlupdate)
    Adodc1.Recordset.MoveNext
    Loop
    
    SQLGrade = "Update Nilai Set Grade=iif (val(Total)=0,'E',iif(val(Total)>0 and val(Total)<60,'D',iif(val(Total)>=60 and val(Total)<75,'C',iif(val(Total)>=75 and val(Total)<85,'B','A')))) where ket ='KURANG' OR KET='GAGAL'"
    Conn.Execute (SQLGrade)
    SQLKet = "Update Nilai Set ket=iif (Grade='E' or Grade='D','GAGAL',iif(grade='A','MEMUASKAN',iif(grade='B','BAIK','CUKUP'))) where ket ='KURANG' OR KET='GAGAL'"
    Conn.Execute (SQLKet)
    Blank
    TxtKodeMK = ""
    TxtKodeMK.SetFocus
    List1.Clear
End If
End Sub

Private Sub CmdBatal_Click()
Blank
TxtKodeMK.SetFocus
End Sub

Private Sub CmdTutup_Click()
Unload Me
End Sub

Function JmlData()
On Error Resume Next
Adodc1.Recordset.MoveFirst
Item = 0
Do While Not Adodc1.Recordset.EOF And Adodc1.Recordset!nim <> vbNullString
    Item = Item + 1
    Adodc1.Recordset.MoveNext
    LblJumlah = Item
Loop
End Function

Sub Blank()
Call TabelKosong
LblNamaMk = ""
LblJumlah = ""
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

