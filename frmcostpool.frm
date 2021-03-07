VERSION 5.00
Begin VB.Form frmcostpool 
   Caption         =   "Customer Name"
   ClientHeight    =   2130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2130
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   720
      Width           =   4215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Batal"
      Height          =   615
      Left            =   2640
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
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmcostpool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
With frmPool
        If Text1 = "" Then
            .lblcost = "Customer"
            .cmdpool(0).Caption = "Mulai"
            .Text1 = ""
            .Text2 = ""
            .Show
            .Command1.FontBold = False
            Me.Hide
            .Label1 = 0
            .Label2 = 0
            .Label3 = 0
            .Label4 = 0
            .Label5 = 0
            .lbl15 = 0
            .Text3.Enabled = True
            .Text3 = 0
        Else
            .lblcost = Text1
            .cmdpool(0).Caption = "Mulai"
            .Text1 = ""
            .Text2 = ""
            .Text3.Enabled = True
            .Text3 = 0
            .Label1 = 0
            .Label2 = 0
            .Label3 = 0
            .Label4 = 0
            .Label5 = 0
            .lbl15 = 0
            .Show
            .Command1.FontBold = False
            Me.Hide
        End If
End With
End Sub

Private Sub Command2_Click()
Me.Hide
menu_utama.Show
frmPool.Hide
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
End Sub

