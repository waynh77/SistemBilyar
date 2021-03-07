VERSION 5.00
Begin VB.Form frmcost 
   Caption         =   "Customer Name"
   ClientHeight    =   2100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4515
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2100
   ScaleWidth      =   4515
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Batal"
      Height          =   615
      Left            =   2520
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   720
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Masukan Nama Customer (max 30 karakter)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmcost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
With Transaksi_form
        If Text1 = "" Then
            .pembeli = "Customer"
            Me.Hide
            Transaksi_form.Show
            .Text1.Enabled = True
            .Text1.SetFocus
        Else
            Me.Hide
            .pembeli = Text1
            Transaksi_form.Show
            .Text1.Enabled = True
            .Text1.SetFocus
        End If
End With
End Sub

Private Sub Command2_Click()
Me.Hide
menu_utama.Show
Transaksi_form.Hide
End Sub

Private Sub Form_Activate()
Text1.MaxLength = 30
Text1 = ""
Text1.SetFocus
End Sub

Private Sub Form_Load()
Text1 = ""
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
Command1.SetFocus
End If
End Sub
