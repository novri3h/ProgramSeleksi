VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Nilai 
   Caption         =   "Entri Nilai Model Pertama"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8970
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
   ScaleHeight     =   4320
   ScaleWidth      =   8970
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Simpan Data Yang Lulus"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   2000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hapus Semua Data Nilai"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   2000
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   400
      Left            =   120
      Top             =   120
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   5741
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
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "NOMORLMR"
         Caption         =   "NOMOR"
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
         DataField       =   "NAMA"
         Caption         =   "NAMA"
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
         DataField       =   "NILAI1"
         Caption         =   "NILAI1"
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
         DataField       =   "NILAI2"
         Caption         =   "NILAI2"
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
         DataField       =   "NILAI3"
         Caption         =   "NILAI3"
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
      BeginProperty Column05 
         DataField       =   "NILAI4"
         Caption         =   "NILAI4"
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
      BeginProperty Column06 
         DataField       =   "TOTAL"
         Caption         =   "TOTAL"
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
      BeginProperty Column07 
         DataField       =   "SKOR"
         Caption         =   "SKOR"
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
      BeginProperty Column08 
         DataField       =   "KET"
         Caption         =   "KET"
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
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   1995,024
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   705,26
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   705,26
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   705,26
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   705,26
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   705,26
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   705,26
         EndProperty
         BeginProperty Column08 
            Locked          =   -1  'True
            ColumnWidth     =   1005,165
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Nilai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\ADOSeleksi.mdb"
Adodc1.RecordSource = "Nilai"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
End Sub

Private Sub Command1_Click() 'jika command1 diklik
Pesan = MsgBox("yakin akan dihapus..?", vbYesNo, "konfirmasi") 'tampilkan pesan..
If Pesan = vbYes Then 'jika dijawab YES
    Adodc1.Recordset.MoveFirst 'mulailah dari record pertama
    Do Until Adodc1.Recordset.EOF 'baca tabel sampai record terakhir
        For i = 1 To Adodc1.Recordset.RecordCount 'lakukan berulang-ulang sebanyak data nilai
            Adodc1.Recordset!Nilai1 = 0 'nilai1 kosongkan (dan seterusnya)
            Adodc1.Recordset!Nilai2 = 0
            Adodc1.Recordset!Nilai3 = 0
            Adodc1.Recordset!Nilai4 = 0
            Adodc1.Recordset!Total = 0
            Adodc1.Recordset!Skor = 0
            Adodc1.Recordset!Ket = vbNullString
        Next i
        Adodc1.Recordset.MoveNext
    Loop
End If
End Sub

Private Sub Command2_Click() 'jika command2 diklik
Call BukaDB 'buka database
Dim HapusHasil As String
HapusHasil = "delete * from  Hasil" 'hapus dulu isi tabel nilai
Conn.Execute (HapusHasil) 'eksekusi penghapusan
Adodc1.Recordset.MoveFirst 'mulailah dari baris awal
Do Until Adodc1.Recordset.EOF 'lakukan sampai baris akhir
    If Adodc1.Recordset!Ket = "LULUS" Then 'jika keteragannya LULUS
        Dim SimpanHasil As String
        'simpan data yang lulus tersebut
        SimpanHasil = "insert into hasil(Nomorlmr,nama,Nilai1,Nilai2,Nilai3,Nilai4,Total,Skor,ket) values " & Chr(13) & _
        "('" & Adodc1.Recordset!NomorLmr & "','" & Adodc1.Recordset!Nama & "', " & Chr(13) & _
        "'" & Adodc1.Recordset!Nilai1 & "','" & Adodc1.Recordset!Nilai2 & "', " & Chr(13) & _
        "'" & Adodc1.Recordset!Nilai3 & "','" & Adodc1.Recordset!Nilai4 & "', " & Chr(13) & _
        "'" & Adodc1.Recordset!Total & "','" & Adodc1.Recordset!Skor & "','" & Adodc1.Recordset!Ket & "')"
        Conn.Execute (SimpanHasil) 'eksekusi penyimpanan
    End If
    Adodc1.Recordset.MoveNext
Loop
MsgBox "Data Yang Lulus Berhasil Disimpan" 'tampilkan pesan
End Sub

Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer) 'setelah kolom diedit (diisi data)
On Error GoTo salah
If DataGrid1.Col = 2 Then 'jika saat itu kursor berada di kolom 3 (nilai1)
    Nilai 'jalankan prosedur pencari nilai
    Adodc1.Recordset.MoveNext 'kursor pindah ke baris berikutnya
    DataGrid1.Col = 2 'kursor pindah ke kolom 3 (nilai1)
ElseIf DataGrid1.Col = 3 Then
    Nilai
    Adodc1.Recordset.MoveNext
    DataGrid1.Col = 3
ElseIf DataGrid1.Col = 4 Then
    Nilai
    Adodc1.Recordset.MoveNext
    DataGrid1.Col = 4
ElseIf DataGrid1.Col = 5 Then
    Nilai
    Adodc1.Recordset.MoveNext
    DataGrid1.Col = 5
End If
On Error GoTo 0
Exit Sub
salah:
MsgBox "Cek isi data, harus angka"
End Sub

Private Sub Form_Load()
DataGrid1.Col = 2
End Sub

'prosedur mencari nilai
Sub Nilai()
    'Total=(Nilai1+Nilai2+Nilai3+Nilai4)
    DataGrid1.Columns(6) = Val(Adodc1.Recordset!Nilai1) + Val(Adodc1.Recordset!Nilai2) + Val(Adodc1.Recordset!Nilai3) + Val(Adodc1.Recordset!Nilai4)
    'Skor=Total/4
    Adodc1.Recordset!Skor = DataGrid1.Columns(6) / 4
    'Ket=Jika skor >=80 maka LULUS
    If Val(Adodc1.Recordset!Skor) >= 80 Then Adodc1.Recordset!Ket = "LULUS" Else Adodc1.Recordset!Ket = "GAGAL"
End Sub
