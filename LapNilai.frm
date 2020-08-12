VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form LapNilai 
   Caption         =   "Laporan Nilai"
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2835
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
   ScaleHeight     =   2880
   ScaleWidth      =   2835
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Pelamar Cadangan"
      Height          =   500
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   2500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Laporan Nilai Ujian"
      Height          =   500
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   2500
   End
   Begin Crystal.CrystalReport CR 
      Left            =   2040
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.ListBox List1 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   2500
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cari Data Berdasarkan Kriteria :"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2505
   End
End
Attribute VB_Name = "LapNilai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'panggil file laporan
CR.ReportFileName = App.Path & "\Lap Nilai.rpt"
'jika ada perubahan data direfresh
CR.WindowState = crptMaximized
'tampilkan satu layar penuh
CR.RetrieveDataFiles
'tampilkan ke layar
CR.Action = 0
End Sub

Private Sub Command2_Click()
CR.ReportFileName = App.Path & "\Lap Cadangan.rpt"
CR.WindowState = crptMaximized
CR.RetrieveDataFiles
CR.Action = 0
End Sub

Private Sub Form_Load()
List1.AddItem "LULUS"
List1.AddItem "GAGAL"
End Sub

Private Sub List1_Click()
'panggil laporan nilai yang keterangannya = list1
CR.SelectionFormula = "{Nilai.Ket}='" & List1 & "'"
CR.ReportFileName = App.Path & "\Lap Nilai.rpt"
CR.WindowState = crptMaximized
CR.RetrieveDataFiles
CR.Action = 0
End Sub
