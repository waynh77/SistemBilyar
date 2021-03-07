VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form DBproduk_form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DATA PRODUK"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10920
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   10920
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2040
      TabIndex        =   3
      Text            =   "KLASIFIKASI PRODUK"
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox Kode_produk 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "DBproduk.frx":0000
      Height          =   3735
      Left            =   120
      OleObjectBlob   =   "DBproduk.frx":0019
      TabIndex        =   10
      Top             =   2400
      Width           =   10695
   End
   Begin VB.CommandButton cmdproses 
      Caption         =   "Simpan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   5295
   End
   Begin VB.ComboBox cbo_jenis_prod 
      Height          =   315
      ItemData        =   "DBproduk.frx":09FC
      Left            =   2040
      List            =   "DBproduk.frx":09FE
      TabIndex        =   2
      Text            =   "JENIS PRODUK"
      Top             =   720
      Width           =   3375
   End
   Begin VB.CommandButton cmdbatal 
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   7
      Top             =   1680
      Width           =   5295
   End
   Begin VB.TextBox Harga_produk 
      Height          =   375
      Left            =   7440
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Nama_produk 
      Height          =   285
      Left            =   7440
      TabIndex        =   4
      Top             =   720
      Width           =   3375
   End
   Begin VB.Data DataProduk 
      Caption         =   "DATABASE PRODUK"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "produk"
      Top             =   240
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "KLASIFIKASI PRODUK"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   1680
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "KODE PRODUK"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   240
      Width           =   1185
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "HARGA SATUAN"
      Height          =   195
      Left            =   5520
      TabIndex        =   9
      Top             =   1200
      Width           =   1275
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "NAMA PRODUK"
      Height          =   195
      Left            =   5520
      TabIndex        =   8
      Top             =   720
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "JENIS PRODUK"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1185
   End
End
Attribute VB_Name = "DBproduk_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbo_jenis_prod_Click()
Combo1.Clear
Select Case cbo_jenis_prod.ListIndex
Case 0
    Combo1.AddItem "SALAD"
    Combo1.AddItem "SOUP"
'    Combo1.AddItem "OTHERS..."
    Combo1.ListIndex = (0)
Case 1
    Combo1.AddItem "GRILL"
    Combo1.AddItem "SANDWICH"
    Combo1.AddItem "FAVOURITE"
    Combo1.AddItem "PASTA"
'    Combo1.AddItem "OTHERS..."
    Combo1.ListIndex = (0)
Case 2
    Combo1.AddItem "LIGHT MEAL/SNACK"
    Combo1.AddItem "DESSERT"
'    Combo1.AddItem "OTHERS..."
    Combo1.ListIndex = (0)
Case 3
    Combo1.AddItem "APERITIF"
    Combo1.AddItem "COGNAC"
    Combo1.AddItem "WHISKY"
    Combo1.AddItem "TEQUILA"
    Combo1.AddItem "VODKA"
    Combo1.AddItem "GIN"
    Combo1.AddItem "RUM"
    Combo1.AddItem "PREMIUM LIQUERS"
    Combo1.AddItem "REGULAR LIQUERS"
    Combo1.AddItem "BEER'S"
    Combo1.AddItem "WINES"
    Combo1.ListIndex = (0)
Case 5
    Combo1.AddItem "MOCKTAIL'S"
    Combo1.AddItem "LAZY SMOOTHIES"
    Combo1.AddItem "MILKSHAKE"
    Combo1.AddItem "COFFEE"
    Combo1.AddItem "TEA"
    Combo1.AddItem "FRESH JUICE"
    Combo1.AddItem "SOFT DRINK"
'    Combo1.AddItem "OTHERS..."
    Combo1.ListIndex = (0)
Case 4
    Combo1.AddItem "FANCY SHOOTERS"
    Combo1.AddItem "LONG BALL DRINK'S"
    Combo1.AddItem "FIRE BALL"
    Combo1.AddItem "CRUSHED LOVER"
    Combo1.AddItem "BEER'S COCKTAIL'S"
    Combo1.AddItem "CLASSIC COCKTAIL'S"
    Combo1.AddItem "THE LADIES"
    Combo1.AddItem "CHAMPAGNE COKTAIL'S"
    Combo1.AddItem "MARGARITA'S"
    Combo1.AddItem "ADICTIF COFFEE"
'    Combo1.AddItem "OTHERS..."
    Combo1.ListIndex = (0)
Case 6
    Combo1.AddItem "MAKANAN"
    Combo1.AddItem "MINUMAN"
    Combo1.AddItem "LAIN-LAIN"
    Combo1.ListIndex = (0)
End Select
End Sub

Private Sub cbo_jenis_prod_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        Combo1.SetFocus
    End If
End Sub

Private Sub cmdproses_Click()
If cmdproses.Caption = "&Tambah" Then
    cmdproses.Caption = "&Simpan Tambahan"
    cmdbatal.Caption = "&Batal"
    jelas
    Kode_produk.Text = ""
    auto
    kosong
    Kode_produk.Enabled = False
'    cbo_jenis_prod.SetFocus
    cbo_jenis_prod.ListIndex = (0)
ElseIf cmdproses.Caption = "&Simpan Tambahan" Then
    If Kode_produk = "" Or cbo_jenis_prod = "" Or Combo1 = "" Or Nama_produk = "" Or Harga_produk = "" Then
        x = MsgBox("Maaf Data Belum Lengkap!!!", 0, "PERINGATAN!!!")
        If Kode_produk = "" Then
            Kode_produk.SetFocus
        ElseIf cbo_jenis_prod = "" Then
            cbo_jenis_prod.SetFocus
        ElseIf Combo1 = "" Then
            Combo1.SetFocus
        ElseIf Nama_produk = "" Then
            Nama_produk.SetFocus
        ElseIf Harga_produk = "" Then
            Harga_produk.SetFocus
        End If
    Else
        With DataProduk.Recordset
        a = 0
        If Not .RecordCount = 0 Then
        .MoveFirst
        Do While Not .EOF
            If cbo_jenis_prod = !jenis_produk And Combo1 = !KLASIFIKASI_produk And Nama_produk = !Nama_produk Then
                x = MsgBox("Data Sudah Ada...!!! Silahkan Masukan Data yg lain", vbOKOnly, "Peringatan...!!!")
                Kode_produk = !Kode_produk
                Combo1 = !KLASIFIKASI_produk
                Nama_produk = !Nama_produk
                cbo_jenis_prod = !jenis_produk
                Harga_produk = !Harga_produk
                burem
                a = 1
                .MoveLast
            End If
            .MoveNext
        Loop
        End If
        If a = 0 Then
            .AddNew
            !Kode_produk = Kode_produk
            !jenis_produk = cbo_jenis_prod
            !KLASIFIKASI_produk = Combo1
            !Nama_produk = Nama_produk
            !Harga_produk = Val(Harga_produk)
            .Update
        End If
        End With
        DataProduk.Refresh
        Kode_produk = ""
        cmdproses.Caption = "&Tambah"
        cmdbatal.Caption = "&Keluar"
        kosong
        burem
        awal
    End If
ElseIf cmdproses.Caption = "&Edit" Then
    cmdproses.Caption = "&Simpan Edit"
    cmdbatal.Caption = "&Batal"
    burem
    Kode_produk.Enabled = True
    Kode_produk.Text = ""
    Kode_produk.SetFocus
ElseIf cmdproses.Caption = "&Simpan Edit" Then
    If Kode_produk = "" Or cbo_jenis_prod = "" Or Nama_produk = "" Or Harga_produk = "" Then
        x = MsgBox("Maaf Data Belum Lengkap!!!", 0, "PERINGATAN!!!")
        If Kode_produk = "" Then
            Kode_produk.SetFocus
        ElseIf cbo_jenis_prod = "" Then
            cbo_jenis_prod.SetFocus
        ElseIf Nama_produk = "" Then
            Nama_produk.SetFocus
        ElseIf Harga_produk = "" Then
            Harga_produk.SetFocus
        End If
    Else
        With DataProduk.Recordset
        .Edit
        !Kode_produk = Kode_produk
        !jenis_produk = cbo_jenis_prod
        !KLASIFIKASI_produk = Combo1
        !Nama_produk = Nama_produk
        !Harga_produk = Val(Harga_produk)
        .Update
        End With
        DataProduk.Refresh
        Kode_produk = ""
        kosong
        burem
        cmdproses.Caption = "&Edit"
        cmdbatal.Caption = "&Keluar"
        awal
    End If
ElseIf cmdproses.Caption = "&Cari" Then
    Kode_produk.Text = ""
    kosong
    burem
    Kode_produk.Enabled = True
    Kode_produk.SetFocus
ElseIf cmdproses.Caption = "&Hapus" Then
    cmdbatal.Caption = "&Batal"
    Kode_produk.Text = ""
    Kode_produk.Enabled = True
    Kode_produk.SetFocus
    cmdproses.Caption = "&Hapus Data"
ElseIf cmdproses.Caption = "&Hapus Data" Then
    If Not Nama_produk = "" Then
        x = MsgBox("Apakah Anda Yakin Data Ini Akan Dihapus", vbYesNo, "Konfirmasi")
        If x = vbYes Then
            If DataProduk.Recordset.RecordCount = 0 Then
                y = MsgBox("Maaf Data Masih Kosong, Silahkan Mengisi Data Terlebih dahulu!", vbOKOnly, "Informasi")
                cmdproses.Caption = "&Hapus"
                cmdbatal.Caption = "&Keluar"
                Exit Sub
            Else
                DataProduk.Recordset.Delete
                DataProduk.Refresh
            End If
        End If
    End If
    cmdproses.Caption = "&Hapus"
    cmdbatal.Caption = "&Keluar"
    kosong
    Kode_produk.Text = ""
    burem
End If
End Sub

Private Sub Cmdbatal_Click()
Select Case cmdbatal.Caption
    Case "&Keluar"
        kosong
        Kode_produk = ""
        DBproduk_form.Visible = False
        menu_utama.Visible = True
    Case "&Batal"
    If cmdproses.Caption = "&Simpan Tambahan" Then
        cmdproses.Caption = "&Tambah"
    ElseIf cmdproses.Caption = "&Simpan Edit" Then
        cmdproses.Caption = "&Edit"
        Kode_produk.Enabled = False
        cmdproses.SetFocus
    ElseIf cmdproses.Caption = "&Hapus Data" Then
        cmdproses.Caption = "&Hapus"
        Kode_produk.Enabled = False
        cmdproses.SetFocus
    End If
        kode_barang = ""
        kosong
        burem
        awal
    End Select
End Sub


Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Nama_produk.SetFocus
End If
End Sub

Private Sub Form_Activate()
    Kode_produk.MaxLength = 6
    Nama_produk.MaxLength = 30
    Harga_produk.MaxLength = 8
    burem
    cmdproses.SetFocus
    DataProduk.Refresh
End Sub

Private Sub Form_Load()
cbo_jenis_prod.AddItem "VARIETY OF SALAD & SOUP"
cbo_jenis_prod.AddItem "MAIN COURSE"
cbo_jenis_prod.AddItem "CHOICE OF LIGHT MEAL/SNACK"
cbo_jenis_prod.AddItem "SHOOTER'S ALCOHOLIC DRINK LIST"
cbo_jenis_prod.AddItem "COCKTAILS"
cbo_jenis_prod.AddItem "NON ALCOHOLIC DRINKS"
cbo_jenis_prod.AddItem "OTHER..."
cbo_jenis_prod.ListIndex = (0)
End Sub

Private Sub Harga_produk_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If cmdproses.Caption = "&Simpan" Then
        cmdproses.SetFocus
    End If
End If
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
    Beep
    KeyAscii = 0
End If
End Sub

Private Sub Kode_produk_Change()
    Dim panjang As Byte
    panjang = Len(Kode_produk)
    If panjang < 6 Then
        Exit Sub
    End If
    With DataProduk.Recordset
    If cmdproses.Caption = "&Tambah" Or cmdproses.Caption = "&Simpan Tambahan" Then
        .Index = "idxproduk"
        .Seek "=", Kode_produk
        If .NoMatch Then
            jelas
            kosong
            cbo_jenis_prod.SetFocus
            Exit Sub
        End If
        cmdproses.Caption = "&Tambah"
        burem
        Kode_produk = !Kode_produk
        cbo_jenis_prod = !jenis_produk
        Combo1 = !KLASIFIKASI_produk
        Nama_produk = !Nama_produk
        Harga_produk = !Harga_produk
    ElseIf cmdproses.Caption = "&Edit" Or cmdproses.Caption = "&Simpan Edit" Then
        .Index = "idxproduk"
        .Seek "=", Kode_produk
        If .NoMatch Then
            cmdproses.Caption = "&Edit"
            x = MsgBox("Maaf Data Tidak Ada!!!", 0, "PERINGATAN!!!")
            Exit Sub
        End If
        Kode_produk = !Kode_produk
        cbo_jenis_prod = !jenis_produk
        Combo1 = !KLASIFIKASI_produk
        Nama_produk = !Nama_produk
        Harga_produk = !Harga_produk
        jelas
        cbo_jenis_prod.SetFocus
        cmdproses.Caption = "&Simpan Edit"
    ElseIf cmdproses.Caption = "&Hapus" Or cmdproses.Caption = "&Hapus Data" Then
        .Index = "idxproduk"
        .Seek "=", Kode_produk
        If .NoMatch Then
            cmdproses.Caption = "&Hapus"
            x = MsgBox("Maaf Data Tidak Ada!!!", 0, "PERINGATAN!!!")
            Exit Sub
        End If
        Kode_produk = !Kode_produk
        cbo_jenis_prod = !jenis_produk
        Combo1 = !KLASIFIKASI_produk
        Nama_produk = !Nama_produk
        Harga_produk = !Harga_produk
        cmdproses.Caption = "&Hapus Data"
    ElseIf cmdproses.Caption = "&Cari" Then
        .Index = "idxproduk"
        .Seek "=", Kode_produk
        If .NoMatch Then
            x = MsgBox("Maaf Data Tidak Ada!!!", 0, "PERINGATAN!!!")
            Kode_produk.Text = ""
            Kode_produk.SetFocus
            Exit Sub
        End If
        x = MsgBox("Data Ketemu!!!", 0, "CARI DATA")
        Kode_produk = !Kode_produk
        cbo_jenis_prod = !jenis_produk
        Combo1 = !KLASIFIKASI_produk
        Nama_produk = !Nama_produk
        Harga_produk = !Harga_produk
    End If
    End With
End Sub

Private Sub kosong()
    Nama_produk = ""
    cbo_jenis_prod = ""
    Combo1 = ""
    Harga_produk = ""
End Sub

Private Sub jelas()
    Nama_produk.Enabled = True
    cbo_jenis_prod.Enabled = True
    Combo1.Enabled = True
    Harga_produk.Enabled = True
End Sub

Private Sub burem()
    Nama_produk.Enabled = False
    cbo_jenis_prod.Enabled = False
    Combo1.Enabled = False
    Harga_produk.Enabled = False
End Sub

Private Sub kode_produk_keypress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Nama_produk_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
    Harga_produk.SetFocus
    End If
End Sub

Sub awal()
'cmdproses.Caption = "&Tambah"
cmdbatal.Caption = "&Keluar"
End Sub

Private Sub auto()
Dim urutan As String * 6
Dim hitung As Integer
With DataProduk.Recordset
If .RecordCount = 0 Then
    .AddNew
    urutan = "PRO" & "001"
Else
    .MoveLast
    If Val(Left(.Fields("kode_produk"), 3)) <> "000" Then
        urutan = "000" & "001"
    Else
        hitung = Val(Right(.Fields("kode_produk"), 3)) + 1
        urutan = "PRO" & Right("000" & hitung, 3)
    End If
End If
Kode_produk = urutan
End With
End Sub
