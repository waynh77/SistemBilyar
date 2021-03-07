VERSION 5.00
Begin VB.Form frmseting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seting Harga"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   3750
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Text            =   "Text4"
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&KELUAR"
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&EDIT"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
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
      RecordSource    =   "setingharga"
      Top             =   1920
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "PRICE FOR S. POOL"
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
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   1800
   End
   Begin VB.Label Label3 
      Caption         =   "%"
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
      Index           =   0
      Left            =   2640
      TabIndex        =   7
      Top             =   960
      Width           =   120
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "TAX AND SERVICE"
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
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PRICE FOR POOL"
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1560
   End
End
Attribute VB_Name = "frmseting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Command1.Caption = "&EDIT" Then
    Text1.Enabled = True
    Text2.Enabled = True
'    Text3.Enabled = True
    Text4.Enabled = True
    Text1.SetFocus
    Command1.Caption = "&SIMPAN"
Else
    If Text1 = "" Or Text2 = "" Or Text4 = "" Then
        If Text1 = "" Then
            Text1.SetFocus
        ElseIf Text2 = "" Then
            Text2.SetFocus
        ElseIf Text4 = "" Then
            Text4.SetFocus
        End If
    Else
        Data1.Recordset.Edit
        Data1.Recordset!hargapool = Text1
        Data1.Recordset!tax = Text2
        Data1.Recordset!spool = Text4
        Data1.Recordset.Update
        Data1.Refresh
        Text1.Enabled = False
        Text2.Enabled = False
'        Text3.Enabled = False
        Text4.Enabled = False
        Command1.Caption = "&EDIT"
    End If
End If
End Sub

Private Sub Command2_Click()
Me.Hide
menu_utama.Show
End Sub

Private Sub Form_Activate()
Data1.Refresh
If Not Data1.Recordset.BOF Then
    Data1.Recordset.MoveFirst
    Text1 = Data1.Recordset!hargapool
    Text4 = Data1.Recordset!spool
'    Text3 = Format(Data1.Recordset!discount, "##.##")
    Text2 = Format(Data1.Recordset!tax, "##.##")
End If

End Sub

Private Sub Form_Load()
    Command1.Caption = "&EDIT"
    Text1.Enabled = False
    Text2.Enabled = False
'    Text3.Enabled = False
    Text4.Enabled = False
'    Data1.Refresh
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text4.SetFocus
End If
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
    Beep
    KeyAscii = 0
End If
End Sub




Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text2.SetFocus
End If
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
    Beep
    KeyAscii = 0
End If
End Sub
