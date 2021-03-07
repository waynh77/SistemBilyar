VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form cetakLap_frm 
   Caption         =   "Cetak Laporan"
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3960
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2175
   ScaleWidth      =   3960
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton command1 
      Caption         =   "BATAL"
      Height          =   615
      Index           =   1
      Left            =   2280
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   2040
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton command1 
      Caption         =   "CETAK"
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   840
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   101974017
      CurrentDate     =   39629
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   101974017
      CurrentDate     =   39629
   End
   Begin VB.Label Label1 
      Caption         =   "Masukan Akhir Periode"
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Masukan Awal Periode"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "cetakLap_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
        CrystalReport1.ReportFileName = App.Path & "\laporan penjualan.rpt"
        CrystalReport1.SelectionFormula = "{transaksi.tanggal}>= date(" & Format(DTPicker1, "yyyy,m,d") & ") and {transaksi.tanggal}<= date(" & Format(DTPicker2, "yyyy,m,d") & ")"
        CrystalReport1.RetrieveDataFiles
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Action = 1
Case 1
    Unload Me
End Select
End Sub

Private Sub Form_Load()
DTPicker1 = Date
DTPicker2 = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
menu_utama.Enabled = True
End Sub
