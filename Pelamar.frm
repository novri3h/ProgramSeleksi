VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Pelamar 
   Caption         =   "Data Pelamar"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   5910
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Pelamar.frx":0000
      Height          =   1845
      Left            =   120
      TabIndex        =   14
      Top             =   2520
      Width           =   5655
      _ExtentX        =   9975
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "NOMORLMR"
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
         DataField       =   "NAMA"
         Caption         =   "Nama Pelamar"
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
         DataField       =   "ALAMAT"
         Caption         =   "Alamat"
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
         DataField       =   "TELEPON"
         Caption         =   "Telepon"
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
            ColumnWidth     =   975,118
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2505,26
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2505,26
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1005,165
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.TextBox Text2 
         DataField       =   "NAMA"
         DataSource      =   "DTPelamar"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   1080
         TabIndex        =   9
         Top             =   600
         Width           =   4400
      End
      Begin VB.TextBox Text3 
         DataField       =   "ALAMAT"
         DataSource      =   "DTPelamar"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   1080
         TabIndex        =   8
         Top             =   960
         Width           =   4400
      End
      Begin VB.TextBox Text4 
         DataField       =   "TELEPON"
         DataSource      =   "DTPelamar"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   1080
         TabIndex        =   7
         Top             =   1320
         Width           =   1500
      End
      Begin VB.TextBox Text1 
         DataField       =   "NOMORLMR"
         DataSource      =   "DTPelamar"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   1080
         TabIndex        =   6
         Top             =   240
         Width           =   1500
      End
      Begin VB.CommandButton Cmdrefresh 
         Caption         =   "&Refresh"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   5
         Top             =   1800
         Width           =   1000
      End
      Begin VB.CommandButton Cmdtutup 
         Caption         =   "&Tutup"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   4
         Top             =   1800
         Width           =   1000
      End
      Begin VB.CommandButton Cmdhapus 
         Caption         =   "&Hapus"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   3
         Top             =   1800
         Width           =   1000
      End
      Begin VB.CommandButton Cmdedit 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   1800
         Width           =   1000
      End
      Begin VB.CommandButton Cmdinput 
         Caption         =   "&Input"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   1800
         Width           =   1000
      End
      Begin MSAdodcLib.Adodc DT 
         Height          =   405
         Left            =   3480
         Top             =   1320
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
         Caption         =   "Pelamar"
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
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " No Pelamar"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nama"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1005
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Alamat"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Telepon"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   1005
      End
   End
End
Attribute VB_Name = "Pelamar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
DT.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\ADOSeleksi.mdb"
DT.RecordSource = "Pelamar"
DT.Refresh
Set DataGrid1.DataSource = DT
DataGrid1.Refresh
End Sub

Private Sub AutoNomor()
Call BukaDB
RSPelamar.Open ("select * from Pelamar Where NomorLmr In(Select Max(NomorLmr)From Pelamar)Order By NomorLmr Desc"), Conn
RSPelamar.Requery
    Dim Urutan As String * 4
    Dim Hitung As Long
    With RSPelamar
        If .EOF Then
            Urutan = "0001"
            Text1 = Urutan
        Else
            Hitung = Str(!NomorLmr) + 1
            Urutan = Right("0000" & Hitung, 4)
        End If
        Text1 = Urutan
    End With
End Sub

Sub Form_Load()
Text1.MaxLength = 4
Text2.MaxLength = 30
Text3.MaxLength = 30
Text4.MaxLength = 15
Kondisiawal
End Sub

Function CariData()
    Call BukaDB
    RSPelamar.Open "Select * From Pelamar where NomorLmr='" & Text1 & "'", Conn
End Function

Private Sub KosongkanText()
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
End Sub

Private Sub SiapIsi()
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
End Sub

Private Sub TidakSiapIsi()
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
End Sub

Private Sub Kondisiawal()
    KosongkanText
    TidakSiapIsi
    Cmdinput.Caption = "&Input"
    Cmdedit.Caption = "&Edit"
    Cmdhapus.Caption = "&Hapus"
    Cmdtutup.Caption = "&Tutup"
    Cmdinput.Enabled = True
    Cmdedit.Enabled = True
    Cmdhapus.Enabled = True
End Sub

Private Sub TampilkanData()
    With RSPelamar
        If Not RSPelamar.EOF Then
            Text2 = RSPelamar!Nama
            Text3 = RSPelamar!Alamat
            Text4 = RSPelamar!Telepon
        End If
    End With
End Sub

Private Sub CmdRefresh_Click()
    If Cmdinput.Caption = "&Simpan" Then
        Cmdinput.SetFocus
    ElseIf Cmdedit.Caption = "&Simpan" Then
        Cmdedit.SetFocus
    End If
    Call Kondisiawal
    Form_Activate
End Sub

Private Sub CmdInput_click()
    If Cmdinput.Caption = "&Input" Then
        Cmdinput.Caption = "&Simpan"
        Cmdedit.Enabled = False
        Cmdhapus.Enabled = False
        Cmdtutup.Caption = "&Batal"
        SiapIsi
        KosongkanText
        Call AutoNomor
        Text1.Enabled = False
        Text2.SetFocus
    Else
        If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
            MsgBox "Data Belum Lengkap...!"
        Else
            Dim SQLTambah As String
            SQLTambah = "Insert Into Pelamar (NomorLmr,Nama,Alamat,Telepon) values ('" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Text4 & "')"
            Conn.Execute SQLTambah
            Cmdrefresh.SetFocus
        End If
    End If
End Sub

Private Sub CmdEdit_Click()
    If Cmdedit.Caption = "&Edit" Then
        Cmdinput.Enabled = False
        Cmdedit.Caption = "&Simpan"
        Cmdhapus.Enabled = False
        Cmdtutup.Caption = "&Batal"
        SiapIsi
        Text1.SetFocus
    Else
        If Text2 = "" Or Text3 = "" Or Text4 = "" Then
            MsgBox "Masih Ada Data Yang Kosong"
        Else
            Dim SQLEdit As String
            SQLEdit = "Update Pelamar Set Nama= '" & Text2 & "', Alamat='" & Text3 & "', Telepon='" & Text4 & "' where NomorLmr='" & Text1 & "'"
            Conn.Execute SQLEdit
            Cmdrefresh.SetFocus
        End If
    End If
End Sub

Private Sub CmdHapus_Click()
    If Cmdhapus.Caption = "&Hapus" Then
        Cmdinput.Enabled = False
        Cmdedit.Enabled = False
        Cmdtutup.Caption = "&Batal"
        KosongkanText
        SiapIsi
        Text1.SetFocus
    End If
End Sub

Private Sub Cmdtutup_Click()
    Select Case Cmdtutup.Caption
        Case "&Tutup"
            Unload Me
        Case "&Batal"
            TidakSiapIsi
            Kondisiawal
    End Select
End Sub

Private Sub Text1_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Len(Text1) < 4 Then
        MsgBox "Nomor Lamaran Harus 4 Digit"
        Text1.SetFocus
        Exit Sub
    Else
        Text2.SetFocus
    End If

    If Cmdinput.Caption = "&Simpan" Then
        Call CariData
            If Not RSPelamar.EOF Then
                TampilkanData
                MsgBox "Nomor pelamar Sudah Ada"
                KosongkanText
                Text1.SetFocus
            Else
                Text2.SetFocus
            End If
    End If
    
    If Cmdedit.Caption = "&Simpan" Then
        Call CariData
            If Not RSPelamar.EOF Then
                TampilkanData
                Text1.Enabled = False
                Text2.SetFocus
            Else
                MsgBox "Nomor pelamar Tidak Ada"
                Text1 = ""
                Text1.SetFocus
            End If
    End If
    
    If Cmdhapus.Enabled = True Then
        Call CariData
            If Not RSPelamar.EOF Then
                TampilkanData
                Pesan = MsgBox("Yakin akan dihapus", vbYesNo)
                If Pesan = vbYes Then
                    Dim SQLHapus As String
                    SQLHapus = "Delete From Pelamar where NomorLmr= '" & Text1 & "'"
                    Conn.Execute SQLHapus
                    Kondisiawal
                    Cmdrefresh.SetFocus
                Else
                    Kondisiawal
                    Cmdhapus.SetFocus
                End If
            Else
                MsgBox "Data Tidak ditemukan"
                Text1.SetFocus
            End If
    End If
End If
End Sub

Private Sub Text2_Keypress(Keyascii As Integer)
    Keyascii = Asc(UCase(Chr(Keyascii)))
    If Keyascii = 13 Then Text3.SetFocus
End Sub

Private Sub Text3_Keypress(Keyascii As Integer)
    Keyascii = Asc(UCase(Chr(Keyascii)))
    If Keyascii = 13 Then Text4.SetFocus
End Sub

Private Sub Text4_Keypress(Keyascii As Integer)
    If Keyascii = 13 Then
        If Cmdinput.Enabled = True Then
            Cmdinput.SetFocus
        ElseIf Cmdedit.Enabled = True Then
            Cmdedit.SetFocus
        End If
    End If
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub


