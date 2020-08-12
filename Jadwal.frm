VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Jadwal 
   Caption         =   "Transfer Jadwal Testing"
   ClientHeight    =   1275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3795
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
   ScaleHeight     =   1275
   ScaleWidth      =   3795
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Transfer Jadwal Testing"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   3500
   End
   Begin MSAdodcLib.Adodc DT 
      Height          =   405
      Left            =   120
      Top             =   1320
      Visible         =   0   'False
      Width           =   3500
      _ExtentX        =   6165
      _ExtentY        =   714
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
      Caption         =   "Jadwal"
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
End
Attribute VB_Name = "Jadwal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
DT.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\ADOSeleksi.mdb"
DT.RecordSource = "Jadwal"
DT.Refresh
End Sub

Private Sub Command1_Click()
Call BukaDB 'buka database
Dim Hapus As String 'defiisikan sebuah variabel
Hapus = "delete from jadwal" 'hapus isi data jadwal terlebih dahulu
Conn.Execute (Hapus) 'eksekusi penghapusan

RSPelamar.Open "Select * from pelamar", Conn 'buka tabelpelamar
RSPelamar.MoveFirst 'mulailah dari data pertama
Do While Not RSPelamar.EOF 'baca tabel pelamar sampai record terakhir
    Dim Tgl As Date 'definisikan tanggal
    DT.Refresh 'data jadwal direfresh
    If DT.Recordset.RecordCount < 4 Then 'jika jumlah data 5..
        Tgl = Date + 10 'tanggal test 10 hari kemudian
        Grup = 1 'grup 1
    ElseIf DT.Recordset.RecordCount >= 4 And DT.Recordset.RecordCount < 8 Then 'jika data 6-10
        Tgl = Date + 11 'tanggal test 11 hari kemudian
        Grup = 2 'grup 2
    ElseIf DT.Recordset.RecordCount >= 9 Then 'jika jumlah data 11 lebih
        Tgl = Date + 12 'tgl test 12 hari kemudian
        Grup = 3 'grup 3
    End If
    Tempat = "Aula" 'tempat test di Aula
    Test1 = "Ruang 1"
    Test2 = "Ruang 2"
    Test3 = "Ruang 3"
    Test4 = "Ruang 4"
    'simpan data pelamar ke tabel jadwal sesuai kriteria di atas
    Dim Simpan As String
    Simpan = "insert into jadwal(Nomorlmr,Tanggal,Tempat,Grup,Test1,Test2,Test3,Test4) " & Chr(13) & _
    "values ('" & RSPelamar!NomorLmr & "','" & Tgl & "','" & Tempat & "','" & Grup & "','" & Test1 & "','" & Test2 & "','" & Test3 & "','" & Test4 & "')"
    Conn.Execute (Simpan)
    RSPelamar.MoveNext
Loop
Conn.Close
MsgBox "Transfer Jadwal Sukses" 'tampilkan pesan sukses
Menu.mnjadwal.Visible = False 'menu transfer disembunyikan
Unload Me 'form langsung ditutup
End Sub

