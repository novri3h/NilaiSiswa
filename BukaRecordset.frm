VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form BukaRecordset 
   Caption         =   "Form1"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1995
      Left            =   120
      TabIndex        =   8
      Top             =   3480
      Width           =   7455
      _ExtentX        =   13150
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
      Height          =   375
      Left            =   120
      Top             =   3000
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin VB.ListBox List2 
      Height          =   1815
      Left            =   3960
      TabIndex        =   6
      Top             =   1080
      Width           =   3615
   End
   Begin VB.TextBox Text4 
      Height          =   350
      Left            =   6000
      TabIndex        =   5
      Top             =   120
      Width           =   1500
   End
   Begin VB.TextBox Text3 
      Height          =   350
      Left            =   3960
      TabIndex        =   4
      Top             =   120
      Width           =   1500
   End
   Begin VB.TextBox Text2 
      Height          =   350
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   1500
   End
   Begin VB.TextBox Text1 
      Height          =   350
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1500
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   3615
   End
End
Attribute VB_Name = "BukaRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then End
If KeyAscii = 13 Then
    
    Dim Cari As String
    Cari = "select distinct matakuliah.kodemk,matakuliah.namamk,mahasiswa.namamhs from matakuliah,master,mahasiswa where mahasiswa.nim=master.nim and matakuliah.kodemk=master.kodemk and mahasiswa.nim='" & Text1 & "'"
    
    Dim aa As New ADODB.Recordset
    Call BukaDB
    
    aa.Open (Cari), Conn
    If aa.EOF Then
        Label1 = ""
        List1.Clear
        MsgBox "nim salah"
    Else
        Label1 = aa!namamhs
        List1.Clear
        Do While Not aa.EOF
            List1.AddItem aa!kodemk & vbTab & aa!namamk
            aa.MoveNext
        Loop
        Text2 = List1.ListCount
    End If
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then End
If KeyAscii = 13 Then
    
    Dim Cari As String
    Cari = "select distinct mahasiswa.nim,mahasiswa.namamhs,matakuliah.namamk from matakuliah,master,mahasiswa where mahasiswa.nim=master.nim and matakuliah.kodemk=master.kodemk and matakuliah.kodemk='" & Text3 & "'"
    
    Dim aa As New ADODB.Recordset
    Call BukaDB
    
    aa.Open (Cari), Conn
    If aa.EOF Then
        Label2 = ""
        List2.Clear
        MsgBox "kode matakuliah salah"
    Else
        Label2 = aa!namamk
        List2.Clear
        Do While Not aa.EOF
            List2.AddItem aa!nim & vbTab & aa!namamhs
            aa.MoveNext
        Loop
        Text4 = List2.ListCount
        
        Dim Cari2 As String
        Cari2 = "select distinct mahasiswa.nim,mahasiswa.namamhs from matakuliah,master,mahasiswa where mahasiswa.nim=master.nim and matakuliah.kodemk=master.kodemk and matakuliah.kodemk='" & Text3 & "'"
        
        Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBSPMB.mdb"
        Adodc1.RecordSource = Cari2
        Adodc1.Refresh
        Set DataGrid1.DataSource = Adodc1
        DataGrid1.Refresh
          
    End If
End If
End Sub

