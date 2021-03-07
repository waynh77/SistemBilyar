VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form DBadmin_frm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Base Administrator"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   4080
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data Admin"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "data_admin"
      Top             =   2400
      Width           =   3015
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "DBadmin_frm.frx":0000
      Height          =   2535
      Left            =   120
      OleObjectBlob   =   "DBadmin_frm.frx":0014
      TabIndex        =   8
      Top             =   2880
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Keluar"
      Height          =   495
      Index           =   3
      Left            =   2160
      TabIndex        =   6
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hapus"
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Edit"
      Height          =   495
      Index           =   1
      Left            =   2160
      TabIndex        =   4
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tambah"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "User Name"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "DBadmin_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
With Data1.Recordset
Select Case Index
Case 0
    If Command1(0).Caption = "Simpan Tambahan" Then
        If Text1 = "" Or Text2 = "" Then
            x = MsgBox("data blm lengkap", vbOKOnly, "Peringatan")
            If Text1 = "" Then
                Text1.SetFocus
            ElseIf Text2 = "" Then
                Text2.SetFocus
            End If
        Else
            a = 0
            .MoveFirst
            Do While Not .EOF
            If !user = Text1 Then
                a = 1
                b = MsgBox("user sudah ada, silahkan masukan yg lain", vbOKOnly, "Peringatan")
                Text1.SetFocus
                .MoveLast
            End If
            .MoveNext
            Loop
            If a = 0 Then
            .AddNew
            !user = Text1
            !Password = Text2
            .Update
            Data1.Refresh
            Command1(0).Caption = "Tambah"
            Command1(1).Enabled = True
            Command1(2).Enabled = True
            kosong
            burem
            End If
        End If
    Else
        terang
        Command1(0).Caption = "Simpan Tambahan"
        Text1.SetFocus
        Command1(1).Enabled = False
        Command1(2).Enabled = False
    End If
Case 1
    If Not .BOF Then
    If Command1(1).Caption = "Simpan Edit" Then
            If Text1 = "" Or Text2 = "" Then
                x = MsgBox("data blm lengkap", vbOKOnly, "Peringatan")
                If Text1 = "" Then
                    Text1.SetFocus
                ElseIf Text2 = "" Then
                    Text2.SetFocus
                End If
            Else
                .Edit
                !user = Text1
                !Password = Text2
                .Update
                Data1.Refresh
                Command1(1).Caption = "Edit"
                kosong
                burem
                Command1(0).Enabled = True
                Command1(2).Enabled = True
            End If
        Else
        kosong
        terang
        Command1(1).Caption = "Simpan Edit"
        Text1.SetFocus
        Command1(0).Enabled = False
        Command1(2).Enabled = False
        Text2.Enabled = False
        Command1(1).Enabled = False
    End If
    End If
Case 2
    If Not .BOF Then
    If Command1(2).Caption = "Hapus Data" Then
        .Delete
        Data1.Refresh
        Command1(2).Caption = "Hapus"
        kosong
        burem
        Command1(0).Enabled = True
        Command1(1).Enabled = True
    Else
        kosong
        terang
        Command1(2).Caption = "Hapus Data"
        Text1.SetFocus
        Command1(0).Enabled = False
        Command1(1).Enabled = False
        Text2.Enabled = False
        Command1(2).Enabled = False
    End If
    End If
Case 3
    Me.Hide
    kosong
    burem
    Command1(0).Enabled = True
    Command1(1).Enabled = True
    Command1(2).Enabled = True
    Command1(3).Enabled = True
    Data1.Refresh
    menu_utama.Visible = True
End Select
End With
End Sub

Private Sub Form_Activate()
Text1 = ""
Text2 = ""
Data1.Refresh
burem
Command1(0).Caption = "Tambah"
Command1(1).Caption = "Edit"
Command1(2).Caption = "Hapus"
Command1(0).Enabled = True
Command1(1).Enabled = True
Command1(2).Enabled = True
End Sub

Private Sub burem()
Text1.Enabled = False
Text2.Enabled = False
End Sub

Private Sub terang()
Text1.Enabled = True
Text2.Enabled = True
End Sub

Private Sub kosong()
Text1 = ""
Text2 = ""
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Not Data1.Recordset.BOF And Command1(0).Enabled = False Then
    With Data1.Recordset
    .MoveFirst
    .Index = "useridx"
    .Seek "=", Text1
        If .NoMatch Then
            x = MsgBox("data user tidak ada", vbOKOnly, "Peringatan")
            Text1.SetFocus
        Else
            terang
            Text1 = !user
            Text2 = !Password
            If Command1(1).Caption = "Simpan Edit" Then
            Command1(1).Enabled = True
            ElseIf Command1(2).Caption = "Hapus Data" Then
            Command1(2).Enabled = True
            End If
            Text1.SetFocus
        End If
End With
ElseIf Command1(0).Caption = "Simpan Tambahan" Then
    Text2.Enabled = True
    Text2.SetFocus
End If
End If
End Sub
