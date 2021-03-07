VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form test_sql 
   Caption         =   "TEST SQL"
   ClientHeight    =   7155
   ClientLeft      =   1740
   ClientTop       =   690
   ClientWidth     =   8955
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   8955
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
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
      Left            =   7080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6600
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   14
      Text            =   "test_sql.frx":0000
      Top             =   6000
      Width           =   8535
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "test_sql.frx":0006
      Height          =   1935
      Left            =   240
      OleObjectBlob   =   "test_sql.frx":001A
      TabIndex        =   13
      Top             =   2640
      Width           =   8535
   End
   Begin VB.CommandButton Command4 
      Caption         =   "cetak"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   11
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "proses input text sql"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   10
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Text            =   "Text4"
      Top             =   4680
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   4680
      Width           =   495
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   4680
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Selesai"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "proses sql"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Data Data1 
      Caption         =   "master penjualan"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6600
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "test_sql.frx":09FD
      Top             =   5280
      Width           =   8535
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "test_sql.frx":0A03
      Height          =   2295
      Left            =   240
      OleObjectBlob   =   "test_sql.frx":0A17
      TabIndex        =   0
      Top             =   360
      Width           =   8535
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   0
      Width           =   3855
   End
   Begin VB.Label Label3 
      Caption         =   "tahun"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "bulan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "tanggal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   4680
      Width           =   735
   End
End
Attribute VB_Name = "test_sql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim datasql As String
Dim datasql2 As String
Dim tgl As String

Private Sub Command1_Click()
Dim x As String
On Error GoTo salah
Data1.RecordSource = Text1
Data1.Refresh
datasql = Text1
Data2.RecordSource = Text5
Data2.Refresh
datasql2 = Text5
Text1 = ""
Text5 = ""
Text2 = ""
Text3 = ""
Command1.Enabled = False
Text2.SetFocus
If Data1.Recordset.EOF Then
    x = MsgBox("data tidak diketemukan", 0, "info")
    Data1.Refresh
    Data2.Refresh
    Text2 = ""
    Text3 = ""
    Text2.SetFocus
    Text1 = ""
    Command1.Enabled = False
End If
On Error GoTo 0
Exit Sub
salah:
x = MsgBox("he3x... baru belajar yah, tulisan sql salah", 0, "dasar cupu lo")
End Sub

Private Sub Command2_Click()
Me.Hide
menu_utama.Show
End Sub

Private Sub Command3_Click()
If Text2 = "" Or Text3 = "" Or Text4 = "" Then
    If Text2 = "" Then
        Text2.SetFocus
    ElseIf Text3 = "" Then
        Text3.SetFocus
    ElseIf Text4 = "" Then
        Text4 = "2007"
    End If
Else
'    Text1 = "select *,(harga_produk*qty)-(harga_produk*qty*DISCOUNT/100)+(harga_produk*qty*15.5/100) as jumlah from transaksi,produk,pool where cdate(day(transaksi.tanggal) = " & Text2 & " and cdate(month(transaksi.tanggal))= " & Text3 & " and cdate(year(transaksi.tanggal))= " & Text4 & " order by waktu asc"
    Text1 = "select waktu,nama_produk,harga_produk,discount,qty,(harga_produk*qty)-(harga_produk*qty*DISCOUNT/100)+(harga_produk*qty*15.5/100) as jumlah, user from transaksi,produk where transaksi.kode_produk=produk.kode_produk and cdate(day(tanggal)) = " & Text2 & " and cdate(month(tanggal))= " & Text3 & " and cdate(year(tanggal))= " & Text4 & " order by waktu asc"
    Text5 = "select * from pool where cdate(day(tanggal)) = " & Text2 & " and cdate(month(tanggal))= " & Text3 & " and cdate(year(tanggal))= " & Text4 & " order by waktu_mulai asc"
    Label4 = "Tanggal : " & Val(Text2) & "/" & Val(Text3) & "/" & Val(Text4)
    tgl = Label4.Caption
    Command1.Enabled = True
    Command1.SetFocus
    Command3.Enabled = False
End If
End Sub

Private Sub Command4_Click()
testReport.Title = tgl
DataEnvironment1.Commands(2).CommandType = adCmdText
DataEnvironment1.Commands(2).CommandText = datasql
'DataEnvironment1.rsCommand2.Update
DataEnvironment1.Connection2.Properties.Refresh
DataEnvironment1.Commands(2).Properties.Refresh
'DataEnvironment1.Commands(3).CommandType = adCmdText
'DataEnvironment1.Commands(3).CommandText = datasql2
'DataEnvironment1.Commands(3).Properties.Refresh
testReport.Refresh
testReport.WindowState = 2
testReport.Show
End Sub

Private Sub Form_Activate()
Text2.MaxLength = 2
Text3.MaxLength = 2
Text4.MaxLength = 4
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = "2007"
Text2.SetFocus
tgl = ""
datasql = ""
datasql2 = ""
Label4 = ""
Text5 = ""
Data1.RecordSource = ""
Data2.RecordSource = ""
Data1.Refresh
Data2.Refresh
If Text1 = "" Then
    Command1.Enabled = False
    Command3.Enabled = False
End If
'Command3.Default = True
End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
    Beep
    KeyAscii = 0
End If
If KeyAscii = 13 Then
    If Text2 = "" Then
        x = MsgBox("tanggal masih kosong coy...", 0, "diisi dulu dnk bos")
        Text2.SetFocus
    Else
        Text3.SetFocus
    End If
End If
If Val(Text2) > 31 Then
    Beep
    x = MsgBox("tanggal tdk lebih dari 31", 0, "gila kali lo ya...")
    Text2 = ""
    Text2.SetFocus
    KeyAscii = 0
End If
End Sub



Private Sub Text3_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
    Beep
    KeyAscii = 0
End If
If KeyAscii = 13 Then
    If Text3 = "" Then
        x = MsgBox("bulan masih kosong coy...", 0, "diisi dulu dnk bos")
        Text3.SetFocus
    Else
        Text4.SetFocus
    End If
End If
If Val(Text3) > 12 Then
    Beep
    x = MsgBox("bulan tdk lebih dari 12", 0, "gila kali lo ya...")
    KeyAscii = 0
    Text3 = ""
End If
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
    Beep
    KeyAscii = 0
End If
If KeyAscii = 13 Then
    If Text4 = "" Then
        x = MsgBox("tahun masih kosong coy...", 0, "diisi dulu dnk bos")
        Text4.SetFocus
    Else
        Command3.Enabled = True
        Command3.SetFocus
    End If
End If
If Val(Text4) > 2007 Then
    Beep
    x = MsgBox("nyang bener neng... skrng aja baru taon 2007", 0, "gila kali lo ya...")
    KeyAscii = 0
End If
End Sub
