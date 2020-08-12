VERSION 5.00
Begin VB.Form Login 
   Caption         =   "Login"
   ClientHeight    =   1530
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
   ScaleHeight     =   1530
   ScaleWidth      =   3720
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtKodeKsr 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1320
      TabIndex        =   6
      Top             =   1800
      Width           =   2000
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3375
      Begin VB.TextBox TxtPasswordKsr 
         Height          =   350
         Left            =   1200
         TabIndex        =   1
         Top             =   720
         Width           =   2000
      End
      Begin VB.TextBox TxtNamaKsr 
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
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   1005
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
'batasi jumlah karakter
TxtNamaKsr.MaxLength = 30
TxtPasswordKsr.MaxLength = 10
'nama dan password diubah menjadi karakter X
'TxtNamaKsr.PasswordChar = "X"
TxtPasswordKsr.PasswordChar = "X"
TxtPasswordKsr.Enabled = False
TxtKodeKsr.Enabled = False
End Sub

Private Sub TxtNamaKsr_KeyPress(Keyascii As Integer)
'ubah karakter jadi besar semua
Keyascii = Asc(UCase(Chr(Keyascii)))
'jika menekan ESC form ditutup
If Keyascii = 27 Then Unload Me
'jika menekan enter setelah mengisi nama, maka..
If Keyascii = 13 Then
    'buka database
    Call BukaDB
    'cari nama kasir yang diketik
    RSKasir.Open "Select NamaKsr from Kasir where NamaKsr ='" & TxtNamaKsr & "'", Conn
    'jika tidak ditemukan, maka
    If RSKasir.EOF Then
        'batasi akses ke nama kasir 3 kali kesempatan
        A = A + 1
        If 1 - A = 0 Then
            MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
                    "Nama '" & TxtNamaKsr & "' tidak dikenal"
            TxtNamaKsr = ""
            TxtNamaKsr.SetFocus
        ElseIf 2 - A = 0 Then
            MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
                    "Nama '" & TxtNamaKsr & "' tidak dikenal"
            TxtNamaKsr = ""
            TxtNamaKsr.SetFocus
        ElseIf 3 - A = 0 Then
            MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
                    "Nama '" & TxtNamaKsr & "' tidak dikenal" & Chr(13) & _
                    "Kesempatan habis, Ulangi dari awal"
            'End
            Unload Me
        End If
    Else
        'jika nama kasir benar, maka nama kasir menjadi false
        TxtNamaKsr.Enabled = False
        'password kasir menjadi true dan menjadi fokus kursor
        TxtPasswordKsr.Enabled = True
        TxtPasswordKsr.SetFocus
    End If
End If
End Sub

'coding ini sama dengan nama kasir
Private Sub txtpasswordksr_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 27 Then Unload Me
Dim KodeKasir As String
Dim NamaKasir As String
If Keyascii = 13 Then
    Call BukaDB
    RSKasir.Open "Select * from Kasir where NamaKsr ='" & TxtNamaKsr & "' and PasswordKsr='" & TxtPasswordKsr & "'", Conn
    If RSKasir.EOF Then
        B = B + 1
        If 1 - B = 0 Then
            MsgBox "Kesempatan ke " & B & " Salah"
            TxtPasswordKsr = ""
            TxtPasswordKsr.SetFocus
        ElseIf 2 - B = 0 Then
            MsgBox "Kesempatan ke " & B & " Salah"
            TxtPasswordKsr = ""
            TxtPasswordKsr.SetFocus
        ElseIf 3 - B = 0 Then
            MsgBox "Kesempatan ke " & B & " Salah"
            'End
            Unload Me
        End If
    Else
        'jika nama dan password benar, maka...tutup form login
        Unload Me
        'panggil menu utama
        Menu.Show
    End If
End If
End Sub

