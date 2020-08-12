VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Nilai1 
   Caption         =   "Entri Nilai Model Kedua"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4230
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
   ScaleHeight     =   4920
   ScaleWidth      =   4230
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1850
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3254
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
      ColumnCount     =   3
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
         DataField       =   "Test"
         Caption         =   "Test"
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
         DataField       =   "Nilai"
         Caption         =   "Nilai"
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
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1005,165
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   405
      Left            =   1680
      Top             =   4440
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
   Begin VB.CommandButton Cmdtutup 
      Caption         =   "&Tutup"
      Height          =   350
      Left            =   120
      TabIndex        =   12
      Top             =   3960
      Width           =   800
   End
   Begin VB.CommandButton Cmdbatal 
      Caption         =   "&Batal"
      Height          =   350
      Left            =   120
      TabIndex        =   11
      Top             =   3600
      Width           =   800
   End
   Begin VB.CommandButton Cmdsimpan 
      Caption         =   "&Simpan"
      Height          =   350
      Left            =   120
      TabIndex        =   10
      Top             =   3240
      Width           =   800
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   1560
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   480
      Width           =   1500
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nomor ID"
      Height          =   345
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   1395
   End
   Begin VB.Label ID 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1560
      TabIndex        =   14
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label Ket 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   2760
      TabIndex        =   9
      Top             =   3960
      Width           =   1245
   End
   Begin VB.Label Skore 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   2760
      TabIndex        =   8
      Top             =   3600
      Width           =   1245
   End
   Begin VB.Label Total 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   2760
      TabIndex        =   7
      Top             =   3240
      Width           =   1245
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Keterangan"
      Height          =   345
      Left            =   1680
      TabIndex        =   6
      Top             =   3960
      Width           =   1005
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Skore"
      Height          =   345
      Left            =   1680
      TabIndex        =   5
      Top             =   3600
      Width           =   1005
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Nilai"
      Height          =   345
      Left            =   1680
      TabIndex        =   4
      Top             =   3240
      Width           =   1005
   End
   Begin VB.Label Nama 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1560
      TabIndex        =   3
      Top             =   840
      Width           =   2580
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama Pelamar"
      Height          =   345
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1395
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nomor Lamaran"
      Height          =   345
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1395
   End
End
Attribute VB_Name = "Nilai1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'cari nomor id otomatis
Private Sub Auto()
Call BukaDB
RSNilai1.Open "select * from Nilai1 Where ID In(Select Max(ID)From Nilai1)Order By ID Desc", Conn
RSNilai1.Requery
    Dim Urutan As String * 10
    Dim Hitung As Long
    With RSNilai1
        If .EOF Then
            Urutan = Right(Date, 2) + Mid(Date, 4, 2) + Left(Date, 2) + "0001"
            LblID = Urutan
        Else
            If Left(!ID, 6) <> Right(Date, 2) + Mid(Date, 4, 2) + Left(Date, 2) Then
                Urutan = Right(Date, 2) + Mid(Date, 4, 2) + Left(Date, 2) + "0001"
            Else
                Hitung = (!ID) + 1
                Urutan = (Right(Date, 2) + Mid(Date, 4, 2) + Left(Date, 2)) + Right("0000" & Hitung, 4)
            End If
        End If
        ID = Urutan
    End With
End Sub

Private Sub Form_Activate()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\ADOSeleksi.mdb"
Adodc1.RecordSource = "Transaksi"
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh

Call BukaDB
RSPelamar.Open "Pelamar", Conn
Combo1.Clear
Do Until RSPelamar.EOF
    Combo1.AddItem RSPelamar!NomorLmr
    RSPelamar.MoveNext
Loop

Call Tabel_Kosong
Adodc1.Recordset.MoveFirst
Call Auto
End Sub

Private Sub Combo1_Keypress(Keyascii As Integer)
On Error Resume Next
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Combo1 = "" Then
        MsgBox "pilih kode Pelamar...!"
        Combo1.SetFocus
        Exit Sub
    Else
        Call BukaDB
        RSPelamar.Open "Select * from Pelamar where NomorLmr='" & Combo1 & "'", Conn
        If Not RSPelamar.EOF Then
            RSDetail.Open "Select distinct nomorlmr from DetailNilai1 where nomorlmr='" & Combo1 & "'", Conn
            If Not RSDetail.EOF Then
                MsgBox "Data nilai sudah dientri"
                Combo1.SetFocus
                Exit Sub
            Else
                Nama = RSPelamar!Nama
                DataGrid1.SetFocus
                DataGrid1.Col = 2
            End If
        Else
            MsgBox "nomor pelamar tidak terdaftar"
            Combo1.SetFocus
            Exit Sub
        End If
    End If
End If
If Keyascii = 27 Then Unload Me
End Sub

Private Sub Combo1_Click()
    Call BukaDB
    RSPelamar.Open "Select * from Pelamar where NomorLmr='" & Combo1 & "'", Conn
    If Not RSPelamar.EOF Then
        Nama = RSPelamar!Nama
    End If
    Conn.Close
End Sub

Function Tabel_Kosong()
    Adodc1.Recordset.MoveFirst
    Do While Not Adodc1.Recordset.EOF
        Adodc1.Recordset.Delete
        Adodc1.Recordset.MoveNext
    Loop
    For i = 1 To 1
        Adodc1.Recordset.AddNew
        Adodc1.Recordset!nomor = i
        Adodc1.Recordset!Test = "Nilai" + Str(i)
        Adodc1.Recordset.Update
    Next i
    DataGrid1.Col = 1
End Function

Function Tambah_Baris()
    For i = Adodc1.Recordset.RecordCount To Adodc1.Recordset.RecordCount
        Adodc1.Recordset.AddNew
        Adodc1.Recordset!nomor = i + 1
        Adodc1.Recordset!Test = "Nilai " + Adodc1.Recordset!nomor
        Adodc1.Recordset.Update
    Next i
End Function

Private Sub DataGrid1_Keypress(Keyascii As Integer)
If DataGrid1.Col = 2 Then
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack Or Keyascii = vbKeyReturn) Then Keyascii = 0
End If
End Sub

Private Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
    'On Error Resume Next
    If DataGrid1.Col = 2 Then
        Adodc1.Recordset!Nilai = Adodc1.Recordset!Nilai
        Adodc1.Recordset.Update
        DataGrid1.Refresh
        Call Tambah_Baris
        Adodc1.Recordset.MoveNext
        DataGrid1.Col = 2
        Adodc1.Recordset.MoveLast
        Call TotalNilai
        'Total = Format(TotalNilai, "#,###,###")
        Skore = Round(Val(Total) / 5, 2)
        If Skore >= 75 Then
            Ket = "Lulus"
        Else
            Ket = "Gagal"
        End If
        If Adodc1.Recordset.RecordCount = 6 Then
            DataGrid1.Enabled = False
            Cmdsimpan.SetFocus
        End If
    End If
End Sub

Private Sub Bersihkan()
    Total = ""
    Skore = ""
    Ket = ""
    Nama = ""
End Sub

Private Sub CmdSimpan_Click()
If Combo1 = "" Or Total = "" Or Skore = "" Or Ket = "" Then
    MsgBox "Data belum lengkap"
    Exit Sub
Else
    If Total = "" Then
        MsgBox "tidak ada entri nilai"
        Exit Sub
    End If
End If

    Call BukaDB
        
    'simpan transaksi ke tbl Nilai1
    Dim SimpanNilai1 As String
    SimpanNilai1 = "Insert Into Nilai1(id,Total,Skor,Ket)" & _
    "values('" & ID & "','" & Total & "','" & Skore & "','" & Ket & "')"
    Conn.Execute (SimpanNilai1)
    
    'simpan data transaksi ke tabel DetailNilai1
    Adodc1.Recordset.MoveFirst
    Do While Not Adodc1.Recordset.EOF
        If Adodc1.Recordset!Nilai <> vbNullString Then
            Dim SQLDetail As String
            SQLDetail = "Insert Into DetailNilai1(id,Nomorlmr,Test,Nilai) " & _
            "values ('" & ID & "','" & Combo1 & "','" & Adodc1.Recordset!Test & "','" & Adodc1.Recordset!Nilai & "')"
            Conn.Execute (SQLDetail)
        End If
    Adodc1.Recordset.MoveNext
    Loop
    Bersihkan
    Form_Activate
    Combo1.SetFocus
    DataGrid1.Enabled = True
End Sub

Private Sub CmdBatal_Click()
Call Bersihkan
Combo1.SetFocus
Form_Activate
DataGrid1.Enabled = True
End Sub

Function TotalNilai()
On Error Resume Next
Adodc1.Recordset.MoveFirst
Total = 0
Do While Not Adodc1.Recordset.EOF And Adodc1.Recordset!Nilai <> 0
    Total = Total + Adodc1.Recordset!Nilai
    Adodc1.Recordset.MoveNext
    Total = Format(Total, "#,###,###")
Loop
End Function

Private Sub Cmdtutup_Click()
Bersihkan
Call Tabel_Kosong
Unload Me
End Sub
