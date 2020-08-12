VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Menu 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Menu Utama"
   ClientHeight    =   3900
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   5880
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
   Picture         =   "Menu.frx":0000
   ScaleHeight     =   3900
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3525
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   1320
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   1920
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":ADC76
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":ADF90
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":AE2AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":AE5C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":AE8DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":AEBF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":AEF12
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":AF22C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":AF546
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":AF860
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnfile 
      Caption         =   "&File"
      Begin VB.Menu mnpelamar 
         Caption         =   "&Pelamar"
      End
   End
   Begin VB.Menu mnproses 
      Caption         =   "&Proses"
      Begin VB.Menu mnjadwal 
         Caption         =   "&Transfer Jadwal"
      End
      Begin VB.Menu mntransfer 
         Caption         =   "Transfer Data &Pelamar"
      End
      Begin VB.Menu mnmodel1 
         Caption         =   "Input Nilai Model &Pertama"
      End
      Begin VB.Menu mnmodel2 
         Caption         =   "Input Nilai Model &Kedua"
      End
   End
   Begin VB.Menu mnlaporan 
      Caption         =   "&Laporan"
      Begin VB.Menu mnlappelamar 
         Caption         =   "Data Pelamar"
      End
      Begin VB.Menu mnlapjadwal 
         Caption         =   "Data Jadwal"
      End
      Begin VB.Menu mnlapnilai 
         Caption         =   "Data Nilai"
      End
   End
   Begin VB.Menu mnkeluar 
      Caption         =   "&Keluar"
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyPress(Keyascii As Integer)
If Keyascii = 27 Then End
End Sub

Private Sub mnjadwal_Click()
Jadwal.Show
End Sub

Private Sub mnkeluar_Click()
End
End Sub

Private Sub mnlapjadwal_Click()
'memanggil file laporan
CR1.ReportFileName = App.Path & "\Lap data Jadwal.rpt"
'tampilkan satu layar penuh
CR1.WindowState = crptMaximized
'jika ada perubahan data direfresh
CR1.RetrieveDataFiles
'tampilkan le layar
CR1.Action = 0
End Sub

Private Sub mnlapnilai_Click()
LapNilai.Show
End Sub

Private Sub mnlappelamar_Click()
CR1.ReportFileName = App.Path & "\Lap Pelamar.rpt"
CR1.WindowState = crptMaximized
CR1.RetrieveDataFiles
CR1.Action = 0
End Sub

Private Sub mnmodel1_Click()
Nilai.Show
End Sub

Private Sub mnmodel2_Click()
Nilai1.Show
End Sub

Private Sub mnpelamar_Click()
Pelamar.Show
End Sub

Private Sub mntransfer_Click()
Transfer.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Key
    Case "a"
        Pelamar.Show
    Case "b"
        Jadwal.Show
    Case "c"
        Transfer.Show
    Case "d"
        Nilai.Show
    Case "e"
        Nilai1.Show
    Case "f"
       CR1.ReportFileName = App.Path & "\Lap Pelamar.rpt"
        CR1.WindowState = crptMaximized
        CR1.RetrieveDataFiles
        CR1.Action = 0
    Case "g"
        'memanggil file laporan
        CR1.ReportFileName = App.Path & "\Lap data Jadwal.rpt"
        'tampilkan satu layar penuh
        CR1.WindowState = crptMaximized
        'jika ada perubahan data direfresh
        CR1.RetrieveDataFiles
        'tampilkan le layar
        CR1.Action = 0

    Case "h"
        LapNilai.Show
    Case "i"
        End
End Select
End Sub
