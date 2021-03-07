VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Transaksi_form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRANSAKSI"
   ClientHeight    =   7395
   ClientLeft      =   3585
   ClientTop       =   495
   ClientWidth     =   8340
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7265.883
   ScaleMode       =   0  'User
   ScaleWidth      =   8340
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1080
      Width           =   615
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Text            =   "Combo4"
      Top             =   2040
      Width           =   2775
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Proses"
      Height          =   735
      Left            =   6840
      TabIndex        =   30
      Top             =   5520
      Width           =   1215
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Transaksi_form.frx":0000
      Height          =   1335
      Left            =   120
      OleObjectBlob   =   "Transaksi_form.frx":0014
      TabIndex        =   29
      Top             =   5280
      Width           =   6255
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&EDIT"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   4560
      Width           =   1335
   End
   Begin VB.ListBox List2 
      Height          =   3570
      Left            =   4680
      TabIndex        =   24
      Top             =   0
      Width           =   3495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "E&XIT"
      Height          =   495
      Left            =   3120
      TabIndex        =   10
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&PRINT"
      Height          =   975
      Left            =   3480
      TabIndex        =   7
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&CLOSE"
      Height          =   495
      Left            =   1680
      TabIndex        =   9
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Tetap"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Transaksi"
      Top             =   5280
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data Data2 
      Caption         =   "Temp"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "temp_trans"
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data DB_Prod 
      Caption         =   "Produk"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "produk"
      Top             =   3480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&SAVE"
      Height          =   975
      Left            =   2640
      TabIndex        =   6
      Top             =   3000
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   5400
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "Transaksi_form.frx":09E7
      Left            =   1560
      List            =   "Transaksi_form.frx":09E9
      TabIndex        =   5
      Text            =   "Combo3"
      Top             =   3480
      Width           =   975
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Text            =   "Combo2"
      Top             =   2520
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Transaksi_form.frx":09EB
      Left            =   1560
      List            =   "Transaksi_form.frx":09ED
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Line Line7 
      X1              =   120
      X2              =   4320
      Y1              =   1414.858
      Y2              =   1414.858
   End
   Begin VB.Label Label33 
      Caption         =   "%"
      Height          =   255
      Left            =   2280
      TabIndex        =   47
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label32 
      Caption         =   "Discount"
      Height          =   255
      Left            =   120
      TabIndex        =   46
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label31 
      Caption         =   "Label31"
      Height          =   255
      Left            =   4680
      TabIndex        =   45
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Line Line6 
      X1              =   120
      X2              =   4320
      Y1              =   589.524
      Y2              =   589.524
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   120
      X2              =   4320
      Y1              =   4126.668
      Y2              =   4126.668
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   5880
      X2              =   8040
      Y1              =   4952.001
      Y2              =   4952.001
   End
   Begin VB.Line Line3 
      X1              =   4680
      X2              =   8160
      Y1              =   4598.287
      Y2              =   4598.287
   End
   Begin VB.Line Line2 
      X1              =   4680
      X2              =   8160
      Y1              =   3890.858
      Y2              =   3890.858
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      Caption         =   "Label30"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7320
      TabIndex        =   44
      Top             =   4440
      Width           =   690
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "Label28"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5520
      TabIndex        =   43
      Top             =   4440
      Width           =   690
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "Label27"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7320
      TabIndex        =   42
      Top             =   4080
      Width           =   690
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "Label26"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6120
      TabIndex        =   41
      Top             =   4080
      Width           =   690
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "Label25"
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
      Height          =   195
      Left            =   7320
      TabIndex        =   40
      Top             =   4800
      Width           =   690
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "Grand Total"
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
      Height          =   195
      Left            =   6000
      TabIndex        =   39
      Top             =   4800
      Width           =   1020
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "Discount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4680
      TabIndex        =   38
      Top             =   4440
      Width           =   765
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "Tax and Service"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4680
      TabIndex        =   37
      Top             =   4080
      Width           =   1410
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   4335
      Y1              =   943.238
      Y2              =   957.977
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "Klasifikasi"
      Height          =   195
      Left            =   120
      TabIndex        =   36
      Top             =   2040
      Width           =   690
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "Kasir :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   35
      Top             =   4200
      Width           =   645
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Costumer Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   34
      Top             =   720
      Width           =   1545
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   6840
      TabIndex        =   33
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label user 
      AutoSize        =   -1  'True
      Caption         =   "user"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   840
      TabIndex        =   32
      Top             =   4200
      Width           =   465
   End
   Begin VB.Label pembeli 
      AutoSize        =   -1  'True
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1800
      TabIndex        =   31
      Top             =   720
      Width           =   675
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Tanggal :"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   120
      TabIndex        =   28
      Top             =   360
      Width           =   675
   End
   Begin VB.Label Label18 
      Caption         =   "Label18"
      Height          =   255
      Left            =   6840
      TabIndex        =   27
      Top             =   5280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Label17"
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
      Height          =   195
      Left            =   7320
      TabIndex        =   26
      Top             =   3720
      Width           =   690
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Sub Total"
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
      Height          =   195
      Left            =   4680
      TabIndex        =   25
      Top             =   3720
      Width           =   840
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Sub Total"
      Height          =   195
      Left            =   120
      TabIndex        =   23
      Top             =   3960
      Width           =   690
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Label14"
      Height          =   195
      Left            =   1560
      TabIndex        =   22
      Top             =   3960
      Width           =   570
   End
   Begin VB.Label Label13 
      Caption         =   "Label13"
      Height          =   375
      Left            =   6720
      TabIndex        =   21
      Top             =   5760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label12 
      Caption         =   "meja18"
      Height          =   255
      Left            =   3240
      TabIndex        =   20
      Top             =   4320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label no_faktur 
      Caption         =   "Label12"
      Height          =   375
      Left            =   6600
      TabIndex        =   19
      Top             =   6000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jam :"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   2880
      TabIndex        =   18
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Jumlah Pembelian"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   3480
      Width           =   1275
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Label8"
      Height          =   195
      Left            =   1560
      TabIndex        =   16
      Top             =   3000
      Width           =   480
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Harga Satuan"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   3000
      Width           =   990
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Nama Produk"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Jenis Produk"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "JAM"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   3480
      TabIndex        =   12
      Top             =   360
      Width           =   315
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "tgl"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   1080
      TabIndex        =   11
      Top             =   360
      Width           =   165
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
   Begin VB.Menu keluar 
      Caption         =   "&Keluar"
   End
End
Attribute VB_Name = "Transaksi_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub burem()
Combo1.Enabled = True
Combo4.Enabled = False
Combo2.Enabled = False
Combo3.Enabled = False
Text1.Enabled = False
End Sub
Private Sub Combo1_Click()
Combo4.Clear
Select Case Combo1.ListIndex
Case 0
    Combo4.AddItem "SALAD"
    Combo4.AddItem "SOUP"
'    Combo4.AddItem "OTHERS..."
    Combo4.ListIndex = (0)
Case 1
    Combo4.AddItem "GRILL"
    Combo4.AddItem "SANDWICH"
    Combo4.AddItem "FAVOURITE"
    Combo4.AddItem "PASTA"
'    Combo4.AddItem "OTHERS..."
    Combo4.ListIndex = (0)
Case 2
    Combo4.AddItem "LIGHT MEAL/SNACK"
    Combo4.AddItem "DESSERT"
'    Combo4.AddItem "OTHERS..."
    Combo4.ListIndex = (0)
Case 3
    Combo4.AddItem "APERITIF"
    Combo4.AddItem "COGNAC"
    Combo4.AddItem "WHISKY"
    Combo4.AddItem "TEQUILA"
    Combo4.AddItem "VODKA"
    Combo4.AddItem "GIN"
    Combo4.AddItem "RUM"
    Combo4.AddItem "PREMIUM LIQUERS"
    Combo4.AddItem "REGULAR LIQUERS"
    Combo4.AddItem "BEER'S"
    Combo4.AddItem "WINES"
    Combo4.ListIndex = (0)
Case 5
    Combo4.AddItem "MOCKTAIL'S"
    Combo4.AddItem "LAZY SMOOTHIES"
    Combo4.AddItem "MILKSHAKE"
    Combo4.AddItem "COFFEE"
    Combo4.AddItem "TEA"
    Combo4.AddItem "FRESH JUICE"
    Combo4.AddItem "SOFT DRINK"
    Combo4.ListIndex = (0)
Case 4
    Combo4.AddItem "FANCY SHOOTERS"
    Combo4.AddItem "LONG BALL DRINK'S"
    Combo4.AddItem "FIRE BALL"
    Combo4.AddItem "CRUSHED LOVER"
    Combo4.AddItem "BEER'S COCKTAIL'S"
    Combo4.AddItem "CLASSIC COCKTAIL'S"
    Combo4.AddItem "THE LADIES"
    Combo4.AddItem "CHAMPAGNE COKTAIL'S"
    Combo4.AddItem "MARGARITA'S"
    Combo4.AddItem "ADICTIF COFFEE"
'    Combo4.AddItem "BEER'S"
    Combo4.ListIndex = (0)
Case 6
    Combo4.AddItem "MAKANAN"
    Combo4.AddItem "MINUMAN"
    Combo4.AddItem "LAIN-LAIN"
    Combo4.ListIndex = (0)
End Select
Combo4.Enabled = True
If Combo3 <> "" Then
    Label14 = rkanan(Label8 * Combo3, "###,###,###")
End If
Combo4.SetFocus
'Combo4.ListIndex = (0)
'Combo2.Enabled = False
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo2_Click()
With DB_Prod.Recordset
.MoveFirst
Do While Not .EOF
    If Combo1.Text = !jenis_produk And Combo4 = !KLASIFIKASI_produk And Combo2.Text = !Nama_produk Then
        Label8 = rkanan(!Harga_produk, "###,###,###")
        Label13 = !Kode_produk
    End If
    .MoveNext
Loop
End With
Combo3.Enabled = True
If Combo3 <> "" Then
    Label14 = rkanan(Label8 * Combo3, "###,###,###")
End If
Combo3.SetFocus
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo3_Click()
Label14 = rkanan(Label8 * Combo3, "###,###,###")
Command1.SetFocus
End Sub


Private Sub Combo3_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo4_Click()
Combo2.Clear
With DB_Prod.Recordset
.MoveFirst
Do While Not .EOF
    If Combo4.Text = !KLASIFIKASI_produk Then
        Combo2.AddItem !Nama_produk
    End If
    .MoveNext
Loop
End With
Combo2.Enabled = True
Combo2.ListIndex = (0)
If Combo3 <> "" Then
    Label14 = rkanan(Label8 * Combo3, "###,###,###")
End If
Combo2.SetFocus
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Command1_Click()
Dim total As Double
If Combo1 = "" Or Combo1 = "Pilih Salah Satu" Or Combo2 = "" Or Combo3 = "" Then
    x = MsgBox("Maaf data masih ada yang kosong, silahkan diisi dahulu !!!", vbOKOnly, "DATA KOSONG")
    If Combo1 = "" Or Combo1 = "Pilih Salah Satu" Then
        Combo1.SetFocus
    ElseIf Combo2 = "" Then
        Combo2.SetFocus
    ElseIf Combo3 = "" Then
        Combo3.SetFocus
    End If
ElseIf Text1 > 100 Then
        x = MsgBox("discount tidak lebih dari 100%", vbOKOnly, "Peringatan")
        Text1 = 0
        Text1.SetFocus
Else
With Data2.Recordset
.AddNew
!no_meja = Label12
!nama_pembeli = pembeli.Caption
!Kode_produk = Label13
!qty = Val(Combo3)
!discount = Text1
.Update
'Data2.Refresh
'Text1 = !discount
End With
Data2.Refresh
List2.AddItem Combo2
List2.AddItem "    Price @ : Rp " & Format(Label8, "###,###,###")
List2.AddItem "      Qty : " & Combo3 & "  Sub Total : Rp " & Format(Label14, "###,###,###")
kosong
With Data2.Recordset
.MoveFirst
total = 0
Do While Not .EOF
    If !no_meja = Label12 Then
    DB_Prod.Recordset.Index = "idxproduk"
    DB_Prod.Recordset.Seek "=", !Kode_produk
    total = total + DB_Prod.Recordset!Harga_produk * !qty
    Text1 = !discount
    End If
    .MoveNext
Loop
End With
Label17 = rkanan(total, "###,###,###")
frmseting.Data1.Refresh
Label26 = "(" & rkanan(frmseting.Data1.Recordset!tax, "##.##") & "%)"
Label28 = "(" & rkanan(Text1, "##.##") & "%)"
Label27 = rkanan(frmseting.Data1.Recordset!tax * total / 100, "###,###,###")
Label30 = rkanan(total * Text1 / 100, "###,###,###")
Label25 = rkanan(total - (total * Text1 / 100) + (frmseting.Data1.Recordset!tax * total / 100), "###,###,###")
burem
End If
End Sub

Private Sub kosong()
Combo1 = "Pilih Salah Satu"
Combo4 = ""
Combo2 = ""
Combo3 = ""
Label8 = ""
Label14 = ""
Text1 = 0
End Sub


Private Sub Command3_Click()
With Data2.Recordset
If Not .BOF Then
List2.Clear
List2.AddItem Label12
kosong
burem
.MoveFirst
Do While Not .EOF
    Data1.Recordset.AddNew
    Data1.Recordset!tanggal = Label2.Caption
    Data1.Recordset!waktu = Label3
    Data1.Recordset!no_meja = !no_meja
    Data1.Recordset!nama_pembeli = !nama_pembeli
    Data1.Recordset!Kode_produk = !Kode_produk
    Data1.Recordset!qty = !qty
    Data1.Recordset!user = user
    Data1.Recordset!discount = !discount
    Data1.Recordset!Status = True
    Data1.Recordset.Update
    .Delete
    .MoveNext
Loop
Data1.Refresh
Data2.Refresh
Label17 = 0
Label30 = 0
Label27 = 0
Label25 = 0
Text1 = 0
Else
    x = MsgBox("Maaf data masih kosong, silahkan diisi dahulu !!!", vbOKOnly, "DATA KOSONG")
End If
End With
If Label4 = "Meja" Then
    menu_utama.Command6(Label18).FontStrikethru = False
    menu_utama.Command6(Label18).FontBold = False
Else
    menu_utama.Command8(Label18).FontStrikethru = False
    menu_utama.Command8(Label18).FontBold = False
End If
End Sub

Private Sub Command4_Click()
If Not Data2.Recordset.BOF Then
    cetaktrans
Else
    x = MsgBox("Maaf data masih kosong, silahkan diisi dahulu !!!", vbOKOnly, "DATA KOSONG")
End If
End Sub
 
Private Sub cetaktrans()
Printer.CurrentX = 0
Printer.CurrentY = 0
'Printer.Show
Printer.Font = "MS sans serif"
'Printer.Cls
'Printer.Caption = "TRANSAKSI " & Label12
With Data2.Recordset
.MoveFirst
Printer.Print
Printer.Print
Printer.FontBold = True
Printer.FontSize = 10
Printer.Print ; "SLINGSHOT"
Printer.FontUnderline = True
Printer.Print ; "Transaksi "; Label12
Printer.FontBold = False
Printer.FontSize = 8
Printer.FontUnderline = False
Printer.Print ; Label2; "     "; Label3
Printer.Print ; "Nama Customer : "; pembeli
Printer.Print
'x = 0
'Do While x < 21
Do While Not .EOF
    DB_Prod.Recordset.MoveFirst
    DB_Prod.Recordset.Index = "idxproduk"
    DB_Prod.Recordset.Seek "=", !Kode_produk
    Printer.Print ; DB_Prod.Recordset!Nama_produk
    'x = x + 1
    Printer.Print Tab(3); "   @Rp "; rkanan(DB_Prod.Recordset!Harga_produk, "###,###,###");
    Printer.Print Tab(23); "  Qty : "; !qty;
    Printer.Print Tab(33); "   Rp "; rkanan(DB_Prod.Recordset!Harga_produk * !qty, "###,###,###")
    .MoveNext
    'x = x + 1
Loop
'printer.Print
'x = x + 1
'Loop
Printer.Print
Printer.FontBold = True
Printer.Print ; "Sub Total ";
Printer.FontUnderline = True
Printer.Print Tab(33); "   Rp "; rkanan(Label17, "###,###,###")
Printer.FontUnderline = False
'printer.FontItalic = True
Printer.Print Tab(5); "Tax & Service ("; frmseting.Data1.Recordset!tax; "%)";
'printer.FontItalic = False
Printer.Print Tab(33); "   Rp "; rkanan(frmseting.Data1.Recordset!tax * Label17 / 100, "###,###,###")
'printer.FontItalic = True
Printer.Print Tab(5); "Discount ( "; Text1; "%)";
Printer.FontUnderline = True
'printer.FontItalic = False
Printer.Print Tab(33); "   Rp "; rkanan(Text1 * Label17 / 100, "###,###,###")
Printer.FontUnderline = False
Printer.Print Tab(15); "Grand Total ";
Printer.FontUnderline = True
Printer.Print Tab(33); "   Rp "; rkanan(Label17 + frmseting.Data1.Recordset!tax * Label17 / 100 - Text1 * Label17 / 100, "###,###,###")
Printer.FontBold = False
Printer.FontUnderline = False
Printer.Print ; "TERIMA KASIH"
Printer.Print ; "Kasir : "; user
End With
Printer.EndDoc
End Sub
Private Sub tampiltrans()
'tampil.CurrentX = 0
'tampil.CurrentY = 0
tampil.Show
tampil.Font = "MS sans serif"
tampil.Cls
tampil.Caption = "TRANSAKSI " & Label12
With Data2.Recordset
.MoveFirst
tampil.Print
tampil.Print
tampil.FontBold = True
tampil.FontSize = 10
tampil.FontUnderline = True
tampil.Print ; "Transaksi "; Label12
tampil.FontBold = False
tampil.FontSize = 8
tampil.FontUnderline = False
tampil.Print ; Label2; "     "; Label3
tampil.Print ; "Nama customer : "; pembeli
tampil.Print
'x = 0
'Do While x < 21
Do While Not .EOF
    DB_Prod.Recordset.MoveFirst
    DB_Prod.Recordset.Index = "idxproduk"
    DB_Prod.Recordset.Seek "=", !Kode_produk
    tampil.Print ; DB_Prod.Recordset!Nama_produk
    'x = x + 1
    tampil.Print Tab(3); "   @Rp "; rkanan(DB_Prod.Recordset!Harga_produk, "###,###,###");
    tampil.Print Tab(23); "  Qty : "; !qty;
    tampil.Print Tab(33); "   Rp "; rkanan(DB_Prod.Recordset!Harga_produk * !qty, "###,###,###")
    .MoveNext
    'x = x + 1
Loop
'tampil.Print
'x = x + 1
'Loop
tampil.Print
tampil.FontBold = True
tampil.Print ; "Sub Total ";
tampil.FontUnderline = True
tampil.Print Tab(33); "   Rp "; rkanan(Label17, "###,###,###")
tampil.FontUnderline = False
'tampil.FontItalic = True
tampil.Print Tab(5); "Tax & Service ("; frmseting.Data1.Recordset!tax; "%)";
'tampil.FontItalic = False
tampil.Print Tab(33); "   Rp "; rkanan(frmseting.Data1.Recordset!tax * Label17 / 100, "###,###,###")
'tampil.FontItalic = True
tampil.Print Tab(5); "Discount ( "; Text1; "%)";
tampil.FontUnderline = True
'tampil.FontItalic = False
tampil.Print Tab(33); "   Rp "; rkanan(Text1 * Label17 / 100, "###,###,###")
tampil.FontUnderline = False
tampil.Print Tab(15); "Grand Total ";
tampil.FontUnderline = True
tampil.Print Tab(33); "   Rp "; rkanan(Label17 + frmseting.Data1.Recordset!tax * Label17 / 100 - Text1 * Label17 / 100, "###,###,###")
tampil.FontBold = False
tampil.FontUnderline = False
tampil.Print ; "Kasir : "; user
End With
'tampil.EndDoc
End Sub
Private Function rkanan(ndata, cformat) As String
    rkanan = Format(ndata, cformat)
    rkanan = Space(Len(cformat) - Len(rkanan)) + rkanan
End Function
Private Sub Command5_Click()
If Label31 = "pool" Then
    If Not Data2.Recordset.BOF Then
        Dim hitpool, hitmakanan, subtotal, hittax, hitdiscount, gt, hargaperjam, hitmenit As Double
        Label8 = "Pool" & Index + 1
        hitpool = 0
        hitmakanan = 0
        subtotal = 0
        hittax = 0
        hitdiscount = 0
        gt = 0
        hargaperjam = 0
        hitmenit = 0
        If frmPool.lbl19 = "Pool10" Then
            hargaperjam = frmseting.Data1.Recordset!spool
        Else
            hargaperjam = frmseting.Data1.Recordset!hargapool
        End If
        With frmPool
            .Label1 = 0
            .Label2 = 0
            .Label3 = 0
            .Label4 = 0
            .Label5 = 0
            .lbl13 = rkanan(hargaperjam, "###,###,###")
            a = 0
            If Not .dt1.Recordset.BOF Then
                .dt1.Recordset.MoveFirst
                Do While Not .dt1.Recordset.EOF
                    If .dt1.Recordset!no_pool = Label12 Then
                        .lblcost = .dt1.Recordset!nama_costumer
                        .Text1 = .dt1.Recordset!waktu_mulai
                        If .dt1.Recordset!waktu_akhir <> blank Then
                            .Text2 = .dt1.Recordset!waktu_akhir
                            hitmenit = (Minute(.Text2) - Minute(.Text1)) + ((Hour(.Text2) - Hour(.Text1)) * 60)
                            .lbl11 = hitmenit
                            If .lbl11 < 60 Then
                                hitpool = hargaperjam
                                .lbl15 = rkanan(hitpool, "###,###,###")
                            Else
                                hitpool = hargapool + ((hitmenit - 60) * (hargapool / 60))
                                .lbl15 = rkanan(hitpool, "###,###,###")
                            End If
                        Else
                            .Text2 = ""
                            .lbl15 = 0
                        End If
                        If .Text2 = "" Then
                            .cmdpool(0).Caption = "Berhenti"
                        Else
                            .cmdpool(0).Caption = "Mulai"
                        End If
                        a = 1
                        .dt1.Recordset.MoveLast
                    End If
                    .dt1.Recordset.MoveNext
                Loop
            Else
                Text1 = ""
            End If
            If Not Transaksi_form.Data2.Recordset.BOF Then
                .List1.Clear
                .List1.AddItem Label12
                b = 0
                Transaksi_form.Data2.Recordset.MoveFirst
                Do While Not Transaksi_form.Data2.Recordset.EOF
                    If Transaksi_form.Data2.Recordset!no_meja = Label12 Then
                        Transaksi_form.DB_Prod.Recordset.Index = "idxproduk"
                        Transaksi_form.DB_Prod.Recordset.Seek "=", Transaksi_form.Data2.Recordset!Kode_produk
                        .List1.AddItem Transaksi_form.DB_Prod.Recordset!Nama_produk
                        .List1.AddItem "   Price @ : Rp " & Format(Transaksi_form.DB_Prod.Recordset!Harga_produk, "###,###,###")
                        .List1.AddItem "      Qty : " & Transaksi_form.Data2.Recordset!qty & "  Sub Total : Rp " & Format(Transaksi_form.Data2.Recordset!qty * Transaksi_form.DB_Prod.Recordset!Harga_produk, "###,###,###")
                        b = 1
                        .lblcost = Transaksi_form.Data2.Recordset!nama_pembeli
                        hitmakanan = hitmakanan + Transaksi_form.Data2.Recordset!qty * Transaksi_form.DB_Prod.Recordset!Harga_produk
                    End If
                    Transaksi_form.Data2.Recordset.MoveNext
                Loop
                .Label1 = rkanan(hitmakanan, "###,###,###")
                frmseting.Data1.Refresh
                subtotal = hitmakanan + hitpool
                .Label2 = rkanan(subtotal, "###,###,###")
                .lbl14(3) = "Tax & Service(" & rkanan(frmseting.Data1.Recordset!tax, "##.##") & "%)"
                .lbl14(5) = "Discount(" & rkanan(.Text3, "##.##") & "%)"
                hittax = frmseting.Data1.Recordset!tax * subtotal / 100
                .Label3 = rkanan(hittax, "###,###,###")
                hitdiscount = subtotal * .Text3 / 100
                .Label4 = rkanan(hitdiscount, "###,###,###")
                gt = subtotal + hittax - hitdiscount
                .Label5 = rkanan(gt, "###,###,###")
                .Command1.FontBold = True
                .Command1.FontBold = True
                If b = 0 Then
                    .Command1.FontBold = False
                End If
            End If
            If a = 0 And b = 1 Then
                .Text1 = ""
            End If
        End With
    End If
    frmPool.Show
    Me.Hide
Else
    menu_utama.Visible = True
    menu_utama.Show
    Transaksi_form.Visible = False
    kosong
    burem
    menu_utama.SSTab1.SetFocus
End If
End Sub


Private Sub Command6_Click()
If Not Data2.Recordset.BOF Then
    frmLogin3.Show
    burem
    Command1.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
    Command6.Enabled = False
    Combo1.Enabled = False
    Me.Enabled = False
Else
    x = MsgBox("Maaf data masih kosong, silahkan diisi dahulu !!!", vbOKOnly, "DATA KOSONG")
End If
End Sub

Private Sub Command7_Click()
Transaksi_form.Height = 5955
'Data2.Recordset.Edit
Data2.Refresh
List2.Clear
List2.AddItem Label12
total = 0
    pembeli = Data2.Recordset!nama_pembeli
    Do While Not Data2.Recordset.EOF
        If Data2.Recordset!no_meja = Label12 Then
            DB_Prod.Recordset.Index = "idxproduk"
            DB_Prod.Recordset.Seek "=", Data2.Recordset!Kode_produk
            List2.AddItem DB_Prod.Recordset!Nama_produk
            List2.AddItem "   Price @ : Rp " & Format(DB_Prod.Recordset!Harga_produk, "###,###,###")
            List2.AddItem "      Qty : " & Data2.Recordset!qty & "  Sub Total : Rp " & Format(Data2.Recordset!qty * DB_Prod.Recordset!Harga_produk, "###,###,###")
            total = total + Data2.Recordset!qty * DB_Prod.Recordset!Harga_produk
        End If
        Data2.Recordset.MoveNext
    Loop
    Label17 = rkanan(total, "###,###,###")
frmseting.Data1.Refresh
Data2.Recordset.MoveFirst
Text1 = Data2.Recordset!discount
Label26 = "(" & rkanan(frmseting.Data1.Recordset!tax, "##.##") & "%)"
Label28 = "(" & rkanan(Text1, "##.##") & "%)"
Label27 = rkanan(frmseting.Data1.Recordset!tax * total / 100, "###,###,###")
Label30 = rkanan(total * Text1 / 100, "###,###,###")
Label25 = rkanan(total - (total * Text1 / 100) + (frmseting.Data1.Recordset!tax * total / 100), "###,###,###")
Combo1.Enabled = True
Command1.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command6.Enabled = True
If Label31 = "pool" Then
Command3.Enabled = False
Command4.Enabled = False
End If
End Sub



Private Sub Form_Activate()
Data1.Refresh
Data2.Refresh
DB_Prod.Refresh
End Sub

Private Sub Form_Load()
    kosong
    burem
    Transaksi_form.Height = 5955
    Transaksi_form.Width = 8430
    Label2 = Date
    Label3 = Time
    Label8 = ""
    Text1 = 0
    Combo3.Text = ""
    x = 1
    Do While x < 101
        Combo3.AddItem x
        x = x + 1
    Loop
    Combo1.AddItem "VARIETY OF SALAD & SOUP"
    Combo1.AddItem "MAIN COURSE"
    Combo1.AddItem "CHOICE OF LIGHT MEAL/SNACK"
    Combo1.AddItem "SHOOTER'S ALCOHOLIC DRINK LIST"
    Combo1.AddItem "COCKTAILS"
    Combo1.AddItem "NON ALCOHOLIC DRINKS"
    Combo1.AddItem "OTHER..."
    'Combo1.ListIndex = (0)
    Combo2.Text = ""
End Sub


Private Sub keluar_Click()
    menu_utama.Show
    Transaksi_form.Visible = False
End Sub

Private Sub Text1_Change()
If Text1 <> "" Then
If Text1 > 100 Then
    x = MsgBox("Discount tidak lebih dari 100%", vbOKOnly, "Peringatan")
    Text1 = 100
End If
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Or KeyAscii = vbKeyBack) Then
    Beep
    KeyAscii = 0
End If
Text1.MaxLength = 5
If KeyAscii = 13 Then
Combo1.SetFocus
End If
End Sub

Private Sub Timer1_Timer()
    Label3 = Time
End Sub
