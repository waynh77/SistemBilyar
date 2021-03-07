VERSION 5.00
Begin VB.Form frmPool 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaksi Pool"
   ClientHeight    =   5310
   ClientLeft      =   1575
   ClientTop       =   2325
   ClientWidth     =   9570
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   9570
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Menit"
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   4680
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   1800
      Width           =   615
   End
   Begin VB.ListBox List1 
      Height          =   4740
      Left            =   6480
      TabIndex        =   35
      Top             =   240
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pesan Makanan/Minuman"
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   3000
      Width           =   3135
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3000
      Top             =   5400
   End
   Begin VB.Data dt2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "pool"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Data dt1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "temp_pool"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Timer tmr1 
      Interval        =   1000
      Left            =   2280
      Top             =   840
   End
   Begin VB.CommandButton cmdpool 
      Caption         =   "Keluar"
      Height          =   975
      Index           =   3
      Left            =   1560
      TabIndex        =   6
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdpool 
      Caption         =   "Cetak"
      Height          =   975
      Index           =   2
      Left            =   1560
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdpool 
      Caption         =   "Simpan"
      Height          =   975
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdpool 
      Caption         =   "Mulai"
      Height          =   975
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "HH:mm:ss"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   4
      EndProperty
      Height          =   285
      Left            =   4680
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "hh:mm:ss"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   4
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lbl22 
      AutoSize        =   -1  'True
      Caption         =   "="
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   3
      Left            =   1320
      TabIndex        =   47
      Top             =   2040
      Width           =   90
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Label11"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   3000
      TabIndex        =   46
      Top             =   2040
      Width           =   570
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Label11"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   1680
      TabIndex        =   45
      Top             =   2040
      Width           =   570
   End
   Begin VB.Label lbl22 
      AutoSize        =   -1  'True
      Caption         =   "Jam"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   2
      Left            =   2400
      TabIndex        =   44
      Top             =   2040
      Width           =   285
   End
   Begin VB.Label lbl22 
      AutoSize        =   -1  'True
      Caption         =   "Menit"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   1
      Left            =   3600
      TabIndex        =   43
      Top             =   2040
      Width           =   390
   End
   Begin VB.Label Label7 
      Caption         =   "%"
      Height          =   255
      Left            =   5400
      TabIndex        =   42
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "Discount"
      Height          =   255
      Left            =   3600
      TabIndex        =   41
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   195
      Left            =   5160
      TabIndex        =   40
      Top             =   5040
      Width           =   585
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   195
      Left            =   5160
      TabIndex        =   39
      Top             =   4680
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   195
      Left            =   5160
      TabIndex        =   38
      Top             =   4440
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   195
      Left            =   5160
      TabIndex        =   37
      Top             =   4080
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   195
      Left            =   5160
      TabIndex        =   36
      Top             =   3720
      Width           =   585
   End
   Begin VB.Line Line5 
      X1              =   3120
      X2              =   6240
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Label lbl14 
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
      ForeColor       =   &H00800080&
      Height          =   195
      Index           =   5
      Left            =   3120
      TabIndex        =   34
      Top             =   4680
      Width           =   765
   End
   Begin VB.Label lbl14 
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
      ForeColor       =   &H00800080&
      Height          =   195
      Index           =   4
      Left            =   3720
      TabIndex        =   33
      Top             =   5040
      Width           =   1020
   End
   Begin VB.Line Line4 
      X1              =   3120
      X2              =   6240
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label lbl14 
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
      ForeColor       =   &H00800080&
      Height          =   195
      Index           =   3
      Left            =   3120
      TabIndex        =   32
      Top             =   4440
      Width           =   1410
   End
   Begin VB.Label lbl14 
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
      ForeColor       =   &H00800080&
      Height          =   195
      Index           =   2
      Left            =   3840
      TabIndex        =   31
      Top             =   4080
      Width           =   840
   End
   Begin VB.Label lbl14 
      AutoSize        =   -1  'True
      Caption         =   "Makanan/Minuman"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   195
      Index           =   1
      Left            =   3120
      TabIndex        =   30
      Top             =   3720
      Width           =   1635
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   240
      X2              =   6240
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   6240
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   240
      X2              =   6240
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lbl22 
      Caption         =   "Menit"
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   29
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label lbl21 
      AutoSize        =   -1  'True
      Caption         =   "Jam"
      Height          =   195
      Left            =   4680
      TabIndex        =   28
      Top             =   360
      Width           =   285
   End
   Begin VB.Label lbl 
      Caption         =   "Tanggal"
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lbl19 
      Caption         =   "Label19"
      Height          =   375
      Left            =   3720
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lbl18 
      Caption         =   "Label18"
      Height          =   255
      Left            =   2520
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lbl17 
      Caption         =   "Label17"
      Height          =   255
      Left            =   4800
      TabIndex        =   24
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label lbl16 
      Caption         =   "Harga/Menit"
      Height          =   255
      Left            =   3600
      TabIndex        =   23
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lbl15 
      AutoSize        =   -1  'True
      Caption         =   "Label15"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   195
      Left            =   5160
      TabIndex        =   22
      Top             =   3480
      Width           =   690
   End
   Begin VB.Label lbl14 
      AutoSize        =   -1  'True
      Caption         =   "Sewa Pool"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   195
      Index           =   0
      Left            =   3840
      TabIndex        =   21
      Top             =   3480
      Width           =   915
   End
   Begin VB.Label lbl13 
      Caption         =   "Label13"
      Height          =   255
      Left            =   1680
      TabIndex        =   20
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label lbl12 
      Caption         =   "Harga/Jam"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label lbl11 
      AutoSize        =   -1  'True
      Caption         =   "Label11"
      ForeColor       =   &H00800080&
      Height          =   195
      Left            =   1680
      TabIndex        =   18
      Top             =   1800
      Width           =   570
   End
   Begin VB.Label lbl10 
      Caption         =   "Jumlah Waktu"
      ForeColor       =   &H00800080&
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label lbl9 
      Caption         =   "Waktu Akhir"
      Height          =   255
      Left            =   3600
      TabIndex        =   16
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lbl8 
      Caption         =   "Waktu Awal"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lbl7 
      Caption         =   "Label7"
      Height          =   255
      Left            =   5160
      TabIndex        =   14
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lbl6 
      Caption         =   "Label6"
      Height          =   255
      Left            =   1200
      TabIndex        =   13
      Top             =   360
      Width           =   975
   End
   Begin VB.Label lblpool 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   240
      Left            =   240
      TabIndex        =   12
      Top             =   0
      Width           =   720
   End
   Begin VB.Label lblopt 
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
      ForeColor       =   &H00800080&
      Height          =   195
      Left            =   4800
      TabIndex        =   11
      Top             =   720
      Width           =   675
   End
   Begin VB.Label lbl2 
      AutoSize        =   -1  'True
      Caption         =   "Nama Kasir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   195
      Left            =   3600
      TabIndex        =   10
      Top             =   720
      Width           =   1065
   End
   Begin VB.Label lblcost 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   195
      Left            =   1680
      TabIndex        =   9
      Top             =   720
      Width           =   675
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Caption         =   "Nama Costumer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   1425
   End
End
Attribute VB_Name = "frmPool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdpool_Click(Index As Integer)
Dim nil1 As Double
Dim nil2 As Double
Dim nil3 As Double
Dim sewapool, perjam, tax, gt, discount, total As Double
Dim rjam As Double
Select Case Index
Case 0
    sewapool = 0
    frmseting.Data1.Refresh
    If lbl19 = "Pool10" Then
        perjam = Val(frmseting.Data1.Recordset!spool)
    Else
        perjam = Val(frmseting.Data1.Recordset!hargapool)
    End If
    tax = 0
    gt = 0
    discount = 0
    If cmdpool(0).Caption = "Mulai" And Text2 = "" Then
        Text1 = ""
        Text2 = ""
        menu_utama.Timer13(lbl18 - 1).Enabled = True
'        Text3 = 0
        cmdpool(0).Caption = "Berhenti"
        cmdpool(1).Enabled = False
        cmdpool(2).Enabled = False
        cmdpool(0).Enabled = True
        
        Text1 = lbl7
        lbl15 = 0
        With dt1.Recordset
            .AddNew
            !no_pool = lbl19
            !nama_costumer = lblcost
            !waktu_mulai = Text1
            !discount = Text3
            !tanggal = lbl6
            .Update
        End With
    ElseIf cmdpool(0).Caption = "Berhenti" Then
        cmdpool(0).Caption = "Mulai"
        cmdpool(1).Enabled = False
        cmdpool(2).Enabled = True
        cmdpool(0).Enabled = False
        Text2 = lbl7
        Text3.Enabled = False
        With dt1.Recordset
            .Index = "temppoolidx"
            .Seek "=", lbl19
            .Edit
            !waktu_akhir = Text2
            !discount = Text3
            menu_utama.Timer13(lbl18 - 1).Enabled = False
        'If dt1.Recordset!tanggal <> lbl6 Then
        'lbl11 = (60 - Minute(Text1)) + ((11 - Hour(Text1)) * 60) + (Minute(Text2) + (Hour(Text2) * 60))
        'Else
        'lbl11 = (Minute(Text2) - Minute(Text1)) + ((Hour(Text2) - Hour(Text1)) * 60)
        'End If
            lbl11 = menu_utama.Label10(lbl18 - 1)
            !lama_main = lbl11
            .Update
            menu_utama.Label10(lbl18 - 1) = 0
        End With
            If lbl11 < 60 Then
                sewapool = perjam
                lbl15 = rkanan(perjam, "###,###,###")
                If Not Transaksi_form.Data2.Recordset.BOF Then
                    total = 0
                    Transaksi_form.Data2.Recordset.MoveFirst
                    Do While Not Transaksi_form.Data2.Recordset.EOF
                        If Transaksi_form.Data2.Recordset!no_meja = lbl19 Then
                            Transaksi_form.DB_Prod.Recordset.Index = "idxproduk"
                            Transaksi_form.DB_Prod.Recordset.Seek "=", Transaksi_form.Data2.Recordset!Kode_produk
                            total = total + Transaksi_form.Data2.Recordset!qty * Transaksi_form.DB_Prod.Recordset!Harga_produk
                        End If
                        Transaksi_form.Data2.Recordset.MoveNext
                    Loop
                End If
                    Label1 = rkanan(total, "###,###,###")
                    frmseting.Data1.Refresh
                    Label2 = rkanan(total + sewapool, "###,###,###")
                    gt = total + sewapool
                    lbl14(3) = "Tax & Service(" & rkanan(frmseting.Data1.Recordset!tax, "##.##") & "%)"
                    lbl14(5) = "Discount(" & rkanan(Text3, "##.##") & "%)"
                    tax = frmseting.Data1.Recordset!tax * gt / 100
                    Label3 = rkanan(tax, "###,###,###")
                    discount = gt * Text3 / 100
                    Label4 = rkanan(discount, "###,###,###")
                    Label5 = rkanan(gt - discount + tax, "###,###,###")
            Else
'                lbl15 = rkanan((frmseting.Data1.Recordset!hargapool) + ((lbl11 - 60) * (frmseting.Data1.Recordset!hargapool / 60)), "###,###,###")
                nil1 = (Val(lbl11) - 60) / 30
                nil2 = Round(nil1)
                nil3 = nil1 - nil2
                If nil3 < 0 Then
                    nil3 = (nil3 + 1)
                End If
                If nil3 <> 0 Then
                    If nil3 > 0.5 Then
                        nil3 = nil2
                    Else
                        nil3 = nil2 + 1
                    End If
                Else
                    nil3 = nil2
                End If
                sewapool = perjam + (nil3 * perjam / 2)
                lbl15 = rkanan(sewapool, "###,###,###")
                If Not Transaksi_form.Data2.Recordset.BOF Then
                    total = 0
                    'gt = 0
                    Transaksi_form.Data2.Recordset.MoveFirst
                    Do While Not Transaksi_form.Data2.Recordset.EOF
                        If Transaksi_form.Data2.Recordset!no_meja = lbl19 Then
                            Transaksi_form.DB_Prod.Recordset.Index = "idxproduk"
                            Transaksi_form.DB_Prod.Recordset.Seek "=", Transaksi_form.Data2.Recordset!Kode_produk
                            total = total + Transaksi_form.Data2.Recordset!qty * Transaksi_form.DB_Prod.Recordset!Harga_produk
                        End If
                        Transaksi_form.Data2.Recordset.MoveNext
                    Loop
                    Label1 = rkanan(total, "###,###,###")
                    frmseting.Data1.Refresh
                    Label2 = rkanan(total + sewapool, "###,###,###")
                    gt = total + sewapool
                    lbl14(3) = "Tax & Service(" & rkanan(frmseting.Data1.Recordset!tax, "##.##") & "%)"
                    lbl14(5) = "Discount(" & rkanan(Text3, "##.##") & "%)"
                    tax = frmseting.Data1.Recordset!tax * gt / 100
                    Label3 = rkanan(tax, "###,###,###")
                    discount = gt * Text3 / 100
                    Label4 = rkanan(discount, "###,###,###")
                    Label5 = rkanan(gt - discount + tax, "###,###,###")
                End If
            End If
    End If
Case 1
If Text2 <> "" Then
    dt2.Refresh
    With dt2.Recordset
            .AddNew
            !no_pool = lbl19
            !nama_costumer = lblcost
            !user = lblopt
            !tanggal = lbl6.Caption
            !waktu_mulai = Text1
            !waktu_selesai = Text2
            !harga_jam = lbl13
            !jumlah_bayar = lbl15 + (lbl15 * frmseting.Data1.Recordset!tax / 100) - (lbl15 * Val(Text3) / 100)
            !discount = Text3
            !lama_main = lbl11
            !Status = True
            .Update
    End With
    With dt1.Recordset
        .Index = "temppoolidx"
        .Seek "=", lbl19
        .Delete
    End With
End If
With Transaksi_form.Data2.Recordset
        If Not .BOF Then
            List1.Clear
            List1.AddItem lbl19
            .MoveFirst
            Do While Not .EOF
                If !no_meja = lbl19 Then
                    Transaksi_form.Data1.Recordset.AddNew
                    Transaksi_form.Data1.Recordset!tanggal = lbl6
                    Transaksi_form.Data1.Recordset!waktu = lbl7
                    Transaksi_form.Data1.Recordset!no_meja = lbl19
                    Transaksi_form.Data1.Recordset!nama_pembeli = lblcost
                    Transaksi_form.Data1.Recordset!Kode_produk = !Kode_produk
                    Transaksi_form.Data1.Recordset!qty = !qty
                    Transaksi_form.Data1.Recordset!user = lblopt
                    Transaksi_form.Data1.Recordset!discount = Text3
                    Transaksi_form.Data1.Recordset!Status = True
                    Transaksi_form.Data1.Recordset.Update
                    .Delete
                End If
                .MoveNext
            Loop
            Transaksi_form.Data1.Refresh
            Transaksi_form.Data2.Refresh
            Label1 = 0
        End If
End With
    cmdpool(1).Enabled = False
    cmdpool(0).Enabled = True
    cmdpool(2).Enabled = False
    Text1 = ""
    Text2 = ""
    lbl11 = 0
    lbl15 = 0
    Text3 = 0
    dt2.Refresh
    dt1.Refresh
    cmdpool(0).Enabled = True
    cmdpool(1).Enabled = False
    cmdpool(2).Enabled = False
    Label1 = 0
    Label2 = 0
    Label3 = 0
    Label4 = 0
    Label5 = 0
    Label8 = 0
    Label9 = 0
    menu_utama.Timer13(Val(lbl18) - 1).Enabled = False
    Text3.Enabled = True
    Text3 = 0
    List1.Clear
    List1.AddItem (lbl19)
    command1.FontBold = False
    cmdpool(0).SetFocus
    menu_utama.Command9(Val(lbl18) - 1).FontBold = False
    menu_utama.Command9(Val(lbl18) - 1).FontStrikethru = False
Case 2
    cetakpool
Case 3
    Me.Hide
    menu_utama.Show
'    cmdpool(0).Enabled = True
    dt1.Refresh
    dt2.Refresh
    Transaksi_form.Data2.RecordSource = "select * from temp_trans"
    Transaksi_form.Data2.Refresh
    If cmdpool(0).Caption = "Berhenti" Then
        With dt1.Recordset
            .Index = "temppoolidx"
            .Seek "=", lbl19
            .Edit
'            !waktu_akhir = Text2
            !discount = Text3
            .Update
            dt1.Refresh
        End With
    End If
End Select
End Sub
Private Sub tampilpool()
If Text2 <> "" Then
'tampil.CurrentX = 0
'tampil.CurrentY = 0
tampil.Show
tampil.Font = "MS sans serif"
tampil.Cls
tampil.Caption = "TRANSAKSI " & lbl19
With Transaksi_form
tampil.Print
tampil.Print
tampil.FontBold = True
tampil.FontSize = 10
tampil.FontUnderline = True
tampil.Print ; "Transaksi "; lbl19
tampil.FontBold = False
tampil.FontSize = 8
tampil.FontUnderline = False
tampil.Print ; lbl6; "     "; lbl7
tampil.Print ; "Nama customer : "; lblcost
tampil.Print
'x = 0
'Do While x < 21
.Data2.RecordSource = "select * from temp_trans where no_meja = " & "'" & lbl19 & "'"
.Data2.Refresh
tampil.Print ; "POOL"
tampil.Print Tab(3); "Sewa/jam Rp "; lbl13
tampil.Print Tab(3); "Waktu Sewa "; lbl11; " Menit";
tampil.Print Tab(33); "   Rp "; rkanan(lbl15, "###,###,###")
tampil.Print
If Not .Data2.Recordset.BOF Then
.Data2.Recordset.MoveFirst
Do While Not .Data2.Recordset.EOF
    .DB_Prod.Recordset.MoveFirst
    .DB_Prod.Recordset.Index = "idxproduk"
    .DB_Prod.Recordset.Seek "=", .Data2.Recordset!Kode_produk
    tampil.Print ; .DB_Prod.Recordset!Nama_produk
    'x = x + 1
    tampil.Print Tab(3); "  @Rp "; rkanan(.DB_Prod.Recordset!Harga_produk, "###,###,###");
    tampil.Print Tab(23); "  Qty : "; .Data2.Recordset!qty;
    tampil.Print Tab(33); "   Rp "; rkanan(.DB_Prod.Recordset!Harga_produk * .Data2.Recordset!qty, "###,###,###")
    .Data2.Recordset.MoveNext
    'x = x + 1
Loop
End If
'tampil.Print
'x = x + 1
'Loop
tampil.Print
tampil.FontBold = True
tampil.Print ; "Sub Total ";
tampil.FontUnderline = True
tampil.Print Tab(33); "   Rp "; rkanan(Label2, "###,###,###")
tampil.FontUnderline = False
'tampil.FontItalic = True
tampil.Print Tab(5); "Tax & Service ("; frmseting.Data1.Recordset!tax; "%)";
'tampil.FontItalic = False
tampil.Print Tab(33); "   Rp "; rkanan(frmseting.Data1.Recordset!tax * Label2 / 100, "###,###,###")
'tampil.FontItalic = True
tampil.Print Tab(5); "Discount ( "; Text3; "%)";
tampil.FontUnderline = True
'tampil.FontItalic = False
tampil.Print Tab(33); "   Rp "; rkanan(Text3 * Label2 / 100, "###,###,###")
tampil.FontUnderline = False
tampil.Print Tab(15); "Grand Total ";
tampil.FontUnderline = True
tampil.Print Tab(33); "   Rp "; rkanan(Label5, "###,###,###")
tampil.FontBold = False
tampil.FontUnderline = False
tampil.Print ; "Kasir : "; lblopt
End With
'tampil.EndDoc
cmdpool(1).Enabled = True
cmdpool(0).Enabled = False
'tampil.EndDoc
Else
x = MsgBox("Transaksi belum selesai...!!!", vbOKOnly, "Peringatan!")
End If
End Sub

Private Sub cetakpool()
If Text2 <> "" Then
Printer.CurrentX = 0
Printer.CurrentY = 0
'printer.Show
Printer.Font = "MS sans serif"
'printer.Cls
'printer.Caption = "TRANSAKSI " & lbl19
With Transaksi_form
Printer.Print
Printer.Print
Printer.FontBold = True
Printer.FontSize = 10
Printer.Print ; "SLINGSHOT"
Printer.FontUnderline = True
Printer.Print ; "Transaksi "; lbl19
Printer.FontBold = False
Printer.FontSize = 8
Printer.FontUnderline = False
Printer.Print ; lbl6; "     "; lbl7
Printer.Print ; "Nama Customer : "; lblcost
Printer.Print
'x = 0
'Do While x < 21
.Data2.RecordSource = "select * from temp_trans where no_meja = " & "'" & lbl19 & "'"
.Data2.Refresh
Printer.Print ; "POOL"
Printer.Print Tab(3); "Sewa/jam Rp "; lbl13
Printer.Print Tab(3); "Waktu Sewa "; lbl11; " Menit";
Printer.Print Tab(33); "   Rp "; rkanan(lbl15, "###,###,###")
Printer.Print
If Not .Data2.Recordset.BOF Then
.Data2.Recordset.MoveFirst
Do While Not .Data2.Recordset.EOF
    .DB_Prod.Recordset.MoveFirst
    .DB_Prod.Recordset.Index = "idxproduk"
    .DB_Prod.Recordset.Seek "=", .Data2.Recordset!Kode_produk
    Printer.Print ; .DB_Prod.Recordset!Nama_produk
    'x = x + 1
    Printer.Print Tab(3); "  @Rp "; rkanan(.DB_Prod.Recordset!Harga_produk, "###,###,###");
    Printer.Print Tab(23); "  Qty : "; .Data2.Recordset!qty;
    Printer.Print Tab(33); "   Rp "; rkanan(.DB_Prod.Recordset!Harga_produk * .Data2.Recordset!qty, "###,###,###")
    .Data2.Recordset.MoveNext
    'x = x + 1
Loop
End If
'printer.Print
'x = x + 1
'Loop
Printer.Print
Printer.FontBold = True
Printer.Print ; "Sub Total ";
Printer.FontUnderline = True
Printer.Print Tab(33); "   Rp "; rkanan(Label2, "###,###,###")
Printer.FontUnderline = False
'printer.FontItalic = True
Printer.Print Tab(5); "Tax & Service ("; frmseting.Data1.Recordset!tax; "%)";
'printer.FontItalic = False
Printer.Print Tab(33); "   Rp "; rkanan(frmseting.Data1.Recordset!tax * Label2 / 100, "###,###,###")
'printer.FontItalic = True
Printer.Print Tab(5); "Discount ( "; Text3; "%)";
Printer.FontUnderline = True
'printer.FontItalic = False
Printer.Print Tab(33); "   Rp "; rkanan(Text3 * Label2 / 100, "###,###,###")
Printer.FontUnderline = False
Printer.Print Tab(15); "Grand Total ";
Printer.FontUnderline = True
Printer.Print Tab(33); "   Rp "; rkanan(Label5, "###,###,###")
Printer.FontBold = False
Printer.FontUnderline = False
Printer.Print ; "TERIMA KASIH"
Printer.Print ; "Kasir : "; lblopt
End With
Printer.EndDoc
cmdpool(1).Enabled = True
cmdpool(0).Enabled = False
'printer.EndDoc
Else
x = MsgBox("Transaksi belum selesai...!!!", vbOKOnly, "Peringatan!")
End If
End Sub
Private Sub Command1_Click()
    Dim y As String
    Transaksi_form.Show
    'Transaksi_form.StartUpPosition = 2
    Transaksi_form.List2.Clear
    Transaksi_form.Caption = "POOL " & lbl18
    Transaksi_form.Label1.Caption = "Transaksi Pool " & lbl18
    Transaksi_form.Label12.Caption = "Pool" & lbl18
    Transaksi_form.Combo1.Text = "Pilih Salah Satu"
    frmPool.Visible = False
    Transaksi_form.Visible = True
    With Transaksi_form
    .Command3.Enabled = False
    .Command4.Enabled = False
        .user.Caption = lblopt.Caption
        .Label17 = 0
        .Text1 = Text3
        .Label27 = 0
        .Label25 = 0
        .Label30 = 0
        .pembeli = lblcost
        .Label1.ForeColor = QBColor(5)
        .Label16.ForeColor = QBColor(5)
        .Label17.ForeColor = QBColor(5)
        .Label4 = "Pool"
        .Label18.Caption = lbl18
        .List2.Clear
        .Data2.RecordSource = "select * from temp_trans"
        .Data2.Refresh
        Transaksi_form.List2.AddItem "Pool" & lbl18
        If .Data2.Recordset.RecordCount <> 0 Then
            .Data2.Recordset.MoveFirst
            total = 0
            .Data2.RecordSource = "select * from temp_trans where no_meja=" & "'" & .Label12 & "'"
            .Data2.Refresh
            If Not .Data2.Recordset.BOF Then
                .pembeli.Caption = .Data2.Recordset!nama_pembeli
                Do While Not .Data2.Recordset.EOF
                    If .Data2.Recordset!no_meja = "Pool" & lbl18 Then
                        .DB_Prod.Recordset.Index = "idxproduk"
                        .DB_Prod.Recordset.Seek "=", .Data2.Recordset!Kode_produk
                        .List2.AddItem .DB_Prod.Recordset!Nama_produk
                        .List2.AddItem "   Price @ : Rp " & Format(.DB_Prod.Recordset!Harga_produk, "###,###,###")
                        .List2.AddItem "      Qty : " & .Data2.Recordset!qty & "  Sub Total : Rp " & Format(.Data2.Recordset!qty * .DB_Prod.Recordset!Harga_produk, "###,###,###")
                        total = total + .Data2.Recordset!qty * .DB_Prod.Recordset!Harga_produk
                    End If
                    .Data2.Recordset.MoveNext
                Loop
            End If
            .pembeli = lblcost
            .pembeli.ForeColor = QBColor(5)
            .user.ForeColor = QBColor(5)
            .Label24.ForeColor = QBColor(5)
            .Label25.ForeColor = QBColor(5)
            .Label17 = rkanan(total, "###,###,###")
            frmseting.Data1.Refresh
            .Label27 = 0
            .Label30 = 0
            .Label26 = "(" & rkanan(frmseting.Data1.Recordset!tax, "##.##") & "%)"
            .Label28 = "(" & rkanan(Text3, "##.##") & "%)"
            .Label27 = rkanan((frmseting.Data1.Recordset!tax * total / 100), "###,###,###")
            .Label30 = rkanan((total * Text3 / 100), "###,###,###")
            .Label25 = rkanan((total + (frmseting.Data1.Recordset!tax * total / 100) - (total * frmseting.Data1.Recordset!discount / 100)), "###,###,###")
        End If
        .Label31 = "pool"
    End With
    Transaksi_form.Height = 5760
End Sub
Private Function rkanan(ndata, cformat) As String
    rkanan = Format(ndata, cformat)
    rkanan = Space(Len(cformat) - Len(rkanan)) + rkanan
End Function



Private Sub Form_Load()
Text3 = 0
Label8 = 0
Label9 = 0
End Sub

Private Sub lbl11_Change()
Dim x, y As Single
If Val(lbl11) > 60 Then
x = Round(Val(lbl11) / 60)
y = Val(lbl11) - x * 60
Label8 = Format(x, "#,###")
Label9 = Format(y, "###")
Else
Label8 = 0
Label9 = lbl11
End If
End Sub

Private Sub Text3_Change()
If Text3 <> "" Then
If Text3 > 100 Then
    x = MsgBox("Discount tidak lebih dari 100%", vbOKOnly, "Peringatan")
    Text3 = 100
Else
    lbl14(5) = "Discount(" & rkanan(Text3, "##.##") & "%)"
'    Label4 = rkanan(Label2 * Text3, "###,###,###")
End If
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Or KeyAscii = vbKeyBack) Then
    Beep
    KeyAscii = 0
End If
Text3.MaxLength = 5
End Sub

Private Sub tmr1_Timer()
lbl7 = Time
End Sub
