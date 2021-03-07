VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form DbTransFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Transaksi"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "Hapus"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Keluar"
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "DbTransFrm.frx":0000
      Height          =   6975
      Left            =   120
      OleObjectBlob   =   "DbTransFrm.frx":0014
      TabIndex        =   0
      Top             =   720
      Width           =   11535
   End
   Begin VB.Data Data1 
      Caption         =   "Data Transaksi"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Transaksi"
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "DbTransFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Me.Hide
    menu_utama.Visible = True
    Data1.RecordSource = "select * from transaksi"
    Data1.Refresh
End Sub

Private Sub Command4_Click()
If Not Data1.Recordset.BOF Then
    x = MsgBox("Apakah anda yakin data akan dihapus?", vbOKCancel, "Hapus Data")
    If x = vbOK Then
        Data1.Recordset.Delete
        Data1.Refresh
    End If
End If
End Sub

Private Sub Form_Activate()
'    Data1.Refresh
'    DBGrid1.Refresh
End Sub

