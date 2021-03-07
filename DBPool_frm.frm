VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form DBPool_frm 
   Caption         =   "Data Base Pool"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Keluar"
      Height          =   375
      Index           =   1
      Left            =   5280
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "DBPool_frm.frx":0000
      Height          =   6975
      Left            =   240
      OleObjectBlob   =   "DBPool_frm.frx":0014
      TabIndex        =   1
      Top             =   720
      Width           =   11295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hapus"
      Height          =   375
      Index           =   0
      Left            =   3360
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.Data Data1 
      Caption         =   "Data Base Pool"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "pool"
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "DBPool_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
If Not Data1.Recordset.BOF Then
    x = MsgBox("Apakah anda yakin data akan dihapus?", vbOKCancel, "Hapus Data")
    If x = vbOK Then
        Data1.Recordset.Delete
        Data1.Refresh
    End If
End If
Case 1
    Me.Hide
    menu_utama.Visible = True
    Data1.Refresh
End Select
End Sub
