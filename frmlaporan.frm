VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmlaporan 
   Caption         =   "LAPORAN PENJUALAN"
   ClientHeight    =   5505
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   8880
   StartUpPosition =   3  'Windows Default
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "frmlaporan.frx":0000
      Height          =   1815
      Left            =   5040
      OleObjectBlob   =   "frmlaporan.frx":0014
      TabIndex        =   17
      Top             =   1920
      Width           =   4575
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "setingharga"
      Top             =   4560
      Width           =   2535
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "pool"
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Transaksi"
      Top             =   4080
      Width           =   2535
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "produk"
      Top             =   4080
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2880
      TabIndex        =   15
      Text            =   "Text2"
      Top             =   2520
      Width           =   975
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   6480
      Top             =   4320
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5160
      Top             =   3840
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4200
      Top             =   2160
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3360
      Top             =   1920
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmlaporan.frx":09E7
      Height          =   1815
      Left            =   5040
      OleObjectBlob   =   "frmlaporan.frx":09FB
      TabIndex        =   11
      Top             =   0
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "PILIH JENIS LAPORAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4695
      Begin VB.OptionButton Option1 
         Caption         =   "Kwartalan"
         Height          =   255
         Index           =   3
         Left            =   2640
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Tahunan"
         Height          =   255
         Index           =   5
         Left            =   2640
         TabIndex        =   13
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Semesteran"
         Height          =   255
         Index           =   4
         Left            =   2640
         TabIndex        =   12
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Triwulan"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   9
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Bulanan"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Pertanggal"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "KELUAR"
      Height          =   495
      Index           =   3
      Left            =   3720
      TabIndex        =   5
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CETAK"
      Height          =   495
      Index           =   2
      Left            =   2520
      TabIndex        =   4
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BATAL"
      Height          =   495
      Index           =   1
      Left            =   1320
      TabIndex        =   3
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PROSES"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Thn"
      Height          =   195
      Left            =   2520
      TabIndex        =   16
      Top             =   2520
      Width           =   285
   End
   Begin VB.Line Line3 
      X1              =   4920
      X2              =   4920
      Y1              =   120
      Y2              =   3720
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   4920
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "MASUKAN PERIODE LAPORAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   2730
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4920
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "label1"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   420
   End
   Begin VB.Menu keluar 
      Caption         =   "&Keluar"
   End
End
Attribute VB_Name = "frmlaporan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim opt As Integer
Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
If Not Option1(5).Value = True And Text1 = "" Or Text2 = "" Then
            x = MsgBox("Input belum lengkap ", vbOKOnly, "Kurang Input")
    If Text2 = "" Then
        Text2.SetFocus
    ElseIf Not Option1(5).Value = True And Text1 = "" Then
        Text1.SetFocus
    End If
Else
    With Data2
    If Option1(0).Value = True Then
            .RecordSource = "Select jenis_produk,klasifikasi_produk,nama_produk,qty,harga_produk from transaksi,produk where cdate(transaksi.tanggal) = '" & Text1 & "/" & Text2 & "' and transaksi.kode_produk=produk.kode_produk "
            Data3.RecordSource = "Select no_pool,nama_costumer,waktu_mulai,waktu_selesai,harga_jam,jumlah_bayar,user from pool where cdate(tanggal) = '" & Text1 & "/" & Text2 & "'"
    ElseIf Option1(1).Value = True Then
        If Text1 > 0 And Text1 < 13 Then
            .RecordSource = "Select jenis_produk,klasifikasi_produk,nama_produk,qty,harga_produk from transaksi,produk where cdate(month(transaksi.tanggal)) = " & Val(Text1) & " and cdate(year(tanggal))= " & Val(Text2) & " and transaksi.kode_produk=produk.kode_produk "
            Data3.RecordSource = "Select no_pool,nama_costumer,waktu_mulai,waktu_selesai,harga_jam,jumlah_bayar,user from pool where cdate(month(tanggal)) = " & Val(Text1) & " and cdate(year(tanggal))=" & Val(Text2)
        Else
            x = MsgBox("Tidak ada Bulan " & Text1, vbOKOnly, "Salah Input")
            Text1 = ""
            Text2 = "2007"
            Text1.SetFocus
            Exit Sub
        End If
    ElseIf Option1(2).Value = True Then
        If Text1 > 0 And Text1 < 5 Then
            Select Case Text1
            Case 1
                .RecordSource = "Select jenis_produk,klasifikasi_produk,nama_produk,qty,harga_produk from transaksi,produk where cdate(month(tanggal)) >= 1 and cdate(month(tanggal)) <= 3 and cdate(year(tanggal))=" & Text2 & " and transaksi.kode_produk=produk.kode_produk "
            Data3.RecordSource = "Select no_pool,nama_costumer,waktu_mulai,waktu_selesai,harga_jam,jumlah_bayar,user from pool where cdate(month(tanggal)) >= 1 and cdate(month(tanggal)) <= 3 and cdate(year(tanggal))=" & Text2
            Case 2
                .RecordSource = "Select jenis_produk,klasifikasi_produk,nama_produk,qty,harga_produk from transaksi,produk where cdate(month(tanggal)) >= 4 and cdate(month(tanggal)) <=6  and cdate(year(tanggal))=" & Text2 & " and transaksi.kode_produk=produk.kode_produk "
            Data3.RecordSource = "Select no_pool,nama_costumer,waktu_mulai,waktu_selesai,harga_jam,jumlah_bayar,user from pool where cdate(month(tanggal)) >= 4 and cdate(month(tanggal)) <= 6 and cdate(year(tanggal))=" & Text2
            Case 3
                .RecordSource = "Select jenis_produk,klasifikasi_produk,nama_produk,qty,harga_produk from transaksi,produk where cdate(month(tanggal)) >= 7 and cdate(month(tanggal)) <= 9 and cdate(year(tanggal))=" & Text2 & " and transaksi.kode_produk=produk.kode_produk "
            Data3.RecordSource = "Select no_pool,nama_costumer,waktu_mulai,waktu_selesai,harga_jam,jumlah_bayar,user from pool where cdate(month(tanggal)) >= 7 and cdate(month(tanggal)) <= 9 and cdate(year(tanggal))=" & Text2
            Case 4
                .RecordSource = "Select jenis_produk,klasifikasi_produk,nama_produk,qty,harga_produk from transaksi,produk where cdate(month(tanggal)) >= 10 and cdate(month(tanggal)) <=12 and cdate(year(tanggal))=" & Text2 & " and transaksi.kode_produk=produk.kode_produk "
            Data3.RecordSource = "Select no_pool,nama_costumer,waktu_mulai,waktu_selesai,harga_jam,jumlah_bayar,user from pool where cdate(month(tanggal)) >= 10 and cdate(month(tanggal)) <= 12 and cdate(year(tanggal))=" & Text2
            End Select
        Else
            x = MsgBox("Tidak ada Triwulan " & Text1, vbOKOnly, "Salah Input")
            Text1 = ""
            Text2 = "2007"
            Text1.SetFocus
            Exit Sub
        End If
    ElseIf Option1(3).Value = True Then
        If Text1 > 0 And Text1 < 4 Then
            Select Case Text1
            Case 1
                .RecordSource = "Select jenis_produk,klasifikasi_produk,nama_produk,qty,harga_produk from transaksi,produk where cdate(month(tanggal)) >= 1 and cdate(month(tanggal)) <= 4 and cdate(year(tanggal))=" & Text2 & " and transaksi.kode_produk=produk.kode_produk "
            Data3.RecordSource = "Select no_pool,nama_costumer,waktu_mulai,waktu_selesai,harga_jam,jumlah_bayar,user from pool where cdate(month(tanggal)) >= 1 and cdate(month(tanggal)) <= 4 and cdate(year(tanggal))=" & Text2
            Case 2
                .RecordSource = "Select jenis_produk,klasifikasi_produk,nama_produk,qty,harga_produk from transaksi,produk where cdate(month(tanggal)) >= 5 and cdate(month(tanggal)) <= 8  and cdate(year(tanggal))=" & Text2 & " and transaksi.kode_produk=produk.kode_produk "
            Data3.RecordSource = "Select no_pool,nama_costumer,waktu_mulai,waktu_selesai,harga_jam,jumlah_bayar,user from pool where cdate(month(tanggal)) >= 5 and cdate(month(tanggal)) <= 8 and cdate(year(tanggal))=" & Text2
            Case 3
                .RecordSource = "Select jenis_produk,klasifikasi_produk,nama_produk,qty,harga_produk from transaksi,produk where cdate(month(tanggal)) >= 9 and cdate(month(tanggal)) <= 12  and cdate(year(tanggal))=" & Text2 & " and transaksi.kode_produk=produk.kode_produk "
            Data3.RecordSource = "Select no_pool,nama_costumer,waktu_mulai,waktu_selesai,harga_jam,jumlah_bayar,user from pool where cdate(month(tanggal)) >= 9 and cdate(month(tanggal)) <= 12 and cdate(year(tanggal))=" & Text2
            End Select
        Else
            x = MsgBox("Tidak ada Kwartal " & Text1, vbOKOnly, "Salah Input")
            Text1 = ""
            Text2 = "2007"
            Text1.SetFocus
            Exit Sub
        End If
    ElseIf Option1(4).Value = True Then
        If Text1 > 0 And Text1 < 3 Then
            Select Case Text1
            Case 1
                .RecordSource = "Select jenis_produk,klasifikasi_produk,nama_produk,qty,harga_produk from transaksi,produk where cdate(month(tanggal)) >= 1 and cdate(month(tanggal)) <= 6  and cdate(year(tanggal))=" & Text2 & " and transaksi.kode_produk=produk.kode_produk "
            Data3.RecordSource = "Select no_pool,nama_costumer,waktu_mulai,waktu_selesai,harga_jam,jumlah_bayar,user from pool where cdate(month(tanggal)) >= 1 and cdate(month(tanggal)) <= 6 and cdate(year(tanggal))=" & Text2
            Case 2
                .RecordSource = "Select jenis_produk,klasifikasi_produk,nama_produk,qty,harga_produk from transaksi,produk where cdate(month(tanggal)) >= 7 and cdate(month(tanggal)) <= 12  and cdate(year(tanggal))=" & Text2 & " and transaksi.kode_produk=produk.kode_produk "
            Data3.RecordSource = "Select no_pool,nama_costumer,waktu_mulai,waktu_selesai,harga_jam,jumlah_bayar,user from pool where cdate(month(tanggal)) >= 7 and cdate(month(tanggal)) <= 12 and cdate(year(tanggal))=" & Text2
            End Select
        Else
            x = MsgBox("Tidak ada Semester " & Text1, vbOKOnly, "Salah Input")
            Text1 = ""
            Text2 = "2007"
            Text1.SetFocus
            Exit Sub
        End If
    ElseIf Option1(5).Value = True Then
        .RecordSource = "Select jenis_produk,klasifikasi_produk,nama_produk,qty,harga_produk from transaksi,produk where cdate(year(tanggal))=" & Text2 & " and transaksi.kode_produk=produk.kode_produk "
    Data3.RecordSource = "Select no_pool,nama_costumer,waktu_mulai,waktu_selesai,harga_jam,jumlah_bayar,user from pool where cdate(year(tanggal))=" & Text2
    End If
    Data2.Refresh
    Data3.Refresh
    If .Recordset.RecordCount = 0 And Data3.Recordset.RecordCount = 0 Then
        x = MsgBox("Tidak ada data/transaksi", vbOKOnly, "Data Kosong")
        Text1 = ""
        Text2 = "2007"
        If Not Option1(5).Value = True Then
            Text1.SetFocus
        Else
            Text2.SetFocus
        End If
    Else
        Timer3.Enabled = True
        Command1(2).Enabled = True
        Text1.Enabled = False
        Text2.Enabled = False
        Command1(0).Enabled = False
    End If
    End With
End If
Case 1
    Timer3.Enabled = False
    Timer2.Enabled = True
    Timer4.Enabled = True
    Option1(opt).Value = False
    Frame1.Enabled = True
    Text1.Enabled = True
    Text2.Enabled = True
    Command1(0).Enabled = True
Case 2
    Timer3.Enabled = False
    Timer2.Enabled = True
    Timer4.Enabled = True
    Option1(opt).Value = False
    Frame1.Enabled = True
    Text1.Enabled = True
    Text2.Enabled = True
    Command1(0).Enabled = True
    tampillaporan
Case 3
'With frmlaporan
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
frmlaporan.Height = 2295
frmlaporan.Width = 5040
Option1(0).Value = False
Option1(1).Value = False
Option1(2).Value = False
Option1(3).Value = False
Option1(4).Value = False
Option1(5).Value = False
Frame1.Enabled = True
Text2 = "2007"
Text1 = ""
'.Data1.Refresh
'.Data2.Refresh
'.Data3.Refresh
'.Data4.Refresh
'End With
    Me.Hide
    menu_utama.Show
'    End
End Select
End Sub

Private Sub tampillaporan()

End Sub
Private Sub Form_Load()
frmlaporan.Height = 2295
frmlaporan.Width = 5040
Option1(Index).Value = False
Text2 = "2007"
Text1 = ""
'Data1.Refresh
'Data2.Refresh
'Data3.Refresh
'Data4.Refresh
End Sub

Private Sub keluar_Click()
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
frmlaporan.Height = 2295
frmlaporan.Width = 5040
Option1(0).Value = False
Option1(1).Value = False
Option1(2).Value = False
Option1(3).Value = False
Option1(4).Value = False
Option1(5).Value = False
Frame1.Enabled = True
Text2 = "2007"
Text1 = ""
    Me.Hide
    menu_utama.Show
End Sub

Private Sub Option1_Click(Index As Integer)
Command1(0).Enabled = True
If frmlaporan.Height < 4530 Then
    Timer1.Enabled = True
End If
Command1(2).Enabled = False
Frame1.Enabled = False
opt = Index
Select Case Index
Case 0
Text1 = ""
Text1.Enabled = True
Text1.SetFocus
Label1 = "Input Tgl dan Bln(Format : MM/DD, contoh : 3/13 = 13 Mar)"
Case 1
Text1 = ""
Text1.Enabled = True
Text1.SetFocus
Label1 = "Input Bulan(Format : 1 sd 12)"
Case 2
Text1 = ""
Text1.Enabled = True
Text1.SetFocus
Label1 = "Input Triwulan(Format : 1/2/3/4)"
Case 3
Text1 = ""
Text1.Enabled = True
Text1.SetFocus
Label1 = "Input Kwartal(Format : 1/2/3)"
Case 4
Text1 = ""
Text1.Enabled = True
Text1.SetFocus
Label1 = "Input Semester(Format : 1/2)"
Text1 = ""
Text1.Enabled = True
Case 5
Text1 = ""
Text2.SetFocus
Text1.Enabled = False
Label1 = "Input Tahun"
End Select
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Option1(0).Value = True Then
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc("/") Or KeyAscii = vbKeyBack) Then
        Beep
        KeyAscii = 0
    End If
Else
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
        Beep
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
    Beep
    KeyAscii = 0
End If
End Sub

Private Sub Timer1_Timer()
If frmlaporan.Height > 4530 Then
    Timer1.Enabled = False
End If
frmlaporan.Height = frmlaporan.Height + 50
End Sub

Private Sub Timer2_Timer()
If frmlaporan.Height < 2295 Then
    Timer2.Enabled = False
frmlaporan.Height = 2295
End If
frmlaporan.Height = frmlaporan.Height - 50
End Sub

Private Sub Timer3_Timer()
If frmlaporan.Width > 9885 Then
    Timer3.Enabled = False
End If
frmlaporan.Width = frmlaporan.Width + 50
End Sub

Private Sub Timer4_Timer()
If frmlaporan.Width < 5040 Then
    Timer4.Enabled = False
frmlaporan.Width = 5040
End If
frmlaporan.Width = frmlaporan.Width - 50
End Sub
