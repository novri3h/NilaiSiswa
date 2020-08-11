VERSION 5.00
Begin VB.Form TransferMhs 
   Caption         =   "Transfer Data Mahasiswa"
   ClientHeight    =   1170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3750
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
   ScaleHeight     =   1170
   ScaleWidth      =   3750
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Transfer Data Mahasiswa"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3500
   End
End
Attribute VB_Name = "TransferMhs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Call BukaDB
End Sub

Private Sub Command1_Click()
Dim Hapus As String
Hapus = "Delete * from TransNilai"
Conn.Execute Hapus

Dim TransferData As New ADODB.Recordset
TransferData.Open "SELECT * from mahasiswa", Conn
TransferData.MoveFirst
Do While Not TransferData.EOF
    Dim Transfer1 As String
    Transfer1 = "Insert Into TransNilai(NIM,NamaMhs,Kelas) values " & _
    "('" & TransferData!nim & "','" & TransferData!namamhs & "','" & TransferData!kelas & "')"
    Conn.Execute (Transfer1)
TransferData.MoveNext
Loop
MsgBox "Transfer Data Mahasiswa sukses"
Unload Me
End Sub


