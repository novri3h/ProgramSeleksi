VERSION 5.00
Begin VB.Form Transfer 
   Caption         =   "Transfer Data Pelamar"
   ClientHeight    =   1425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4245
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
   ScaleHeight     =   1425
   ScaleWidth      =   4245
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Transfer Nomor dan Nama Pelamar Ke Tabel Nilai"
      Height          =   1000
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   4000
   End
End
Attribute VB_Name = "Transfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Call BukaDB 'buka database
Dim HapusNilai As String 'definisikan variabel string
HapusNilai = "delete * from Nilai" 'hapus dulu isi tabel nilai
Conn.Execute (HapusNilai) 'eksekusi penghapusan

RSPelamar.Open "Select * from pelamar", Conn 'buka tabel pelamar
RSPelamar.Requery
RSPelamar.MoveFirst 'mulailah dari reord pertama
Do While Not RSPelamar.EOF 'baca tabel pelamar hingga record terakhir
    Dim Transfer As String
    'simpan nomor lamaran dan nama pelamar ke tabel nilai
    Transfer = "insert into Nilai(Nomorlmr,Nama) " & Chr(13) & _
    "values ('" & RSPelamar!NomorLmr & "','" & RSPelamar!Nama & "')"
    Conn.Execute (Transfer)
    RSPelamar.MoveNext
Loop
MsgBox "Transfer Data Pelamar Sukses" 'tampilkan pesan sukses
Menu.mntransfer.Visible = False 'menu transfer data disembunyikan
Unload Me 'form langsung ditutup
End Sub

