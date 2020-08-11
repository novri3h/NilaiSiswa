VERSION 5.00
Begin VB.Form Login 
   Caption         =   "Login"
   ClientHeight    =   1650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3720
   BeginProperty Font 
      Name            =   "Century"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   ScaleHeight     =   1650
   ScaleWidth      =   3720
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3375
      Begin VB.TextBox TxtPasswordOpr 
         Height          =   350
         Left            =   1200
         TabIndex        =   1
         Top             =   720
         Width           =   2000
      End
      Begin VB.TextBox TxtNamaOpr 
         Height          =   350
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   2000
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Password"
         Height          =   345
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1000
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nama"
         Height          =   345
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1000
      End
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim A As Byte
Dim B As Byte

Private Sub Form_Load()
TxtNamaOpr.MaxLength = 30
TxtPasswordOpr.MaxLength = 10
'TxtNamaOpr.PasswordChar = "X"
TxtPasswordOpr.PasswordChar = "X"
TxtPasswordOpr.Enabled = False
'TxtKodeOpr.Enabled = False
End Sub

Private Sub TxtNamaOpr_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 27 Then Unload Me
If KeyAscii = 13 Then
    Call BukaDB
    RSOperator.Open "Select NamaOpr from Operator where NamaOpr ='" & TxtNamaOpr & "'", Conn
    If RSOperator.EOF Then
        A = A + 1
        If 1 - A = 0 Then
            MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
                    "Nama '" & TxtNamaOpr & "' tidak dikenal"
            TxtNamaOpr = ""
            TxtNamaOpr.SetFocus
        ElseIf 2 - A = 0 Then
            MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
                    "Nama '" & TxtNamaOpr & "' tidak dikenal"
            TxtNamaOpr = ""
            TxtNamaOpr.SetFocus
        ElseIf 3 - A = 0 Then
            MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
                    "Nama '" & TxtNamaOpr & "' tidak dikenal" & Chr(13) & _
                    "Kesempatan habis, Ulangi dari awal"
            Unload Me
        End If
    Else
        TxtNamaOpr.Enabled = False
        TxtPasswordOpr.Enabled = True
        TxtPasswordOpr.SetFocus
    End If
End If
End Sub

Private Sub txtpasswordOpr_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 27 Then Unload Me
Dim KodeOperator As String
Dim NamaOperator As String
If KeyAscii = 13 Then
    Call BukaDB
    RSOperator.Open "Select * from Operator where NamaOpr ='" & TxtNamaOpr & "' and PasswordOpr='" & TxtPasswordOpr & "'", Conn
    If RSOperator.EOF Then
        B = B + 1
        If 1 - B = 0 Then
            MsgBox "Kesempatan ke " & B & " Salah"
            TxtPasswordOpr = ""
            TxtPasswordOpr.SetFocus
        ElseIf 2 - B = 0 Then
            MsgBox "Kesempatan ke " & B & " Salah"
            TxtPasswordOpr = ""
            TxtPasswordOpr.SetFocus
        ElseIf 3 - B = 0 Then
            MsgBox "Kesempatan ke " & B & " Salah"
            Unload Me
        End If
    Else
        Unload Me
        Menu.Show
        Menu.STBAR.Panels(1).Visible = False
        Menu.STBAR.Panels(2).Text = RSOperator!NAMAOPR
        Menu.STBAR.Panels(3).Text = RSOperator!Status
        If Menu.STBAR.Panels(3).Text <> "ADMINISTRATOR" Then
            Menu.mnmaster.Enabled = False
            Menu.mnupdating.Enabled = False
        End If
    End If
    
End If
End Sub

