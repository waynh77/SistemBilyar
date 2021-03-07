VERSION 5.00
Begin VB.Form frmloginuser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Masukan User dan Password"
   ClientHeight    =   1500
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   886.25
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "data_admin"
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   3
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   5
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmloginuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    If menu_utama.Enabled = False Then
        menu_utama.Enabled = True
    Else
        End
    End If
    Me.Hide
        txtUserName = ""
    txtPassword = ""
End Sub

Private Sub cmdOK_Click()
Dim a As Double
With Data1.Recordset
If Not .BOF Then
    .MoveFirst
    a = 0
    Do While Not .EOF
    'check for correct password
        If txtUserName = !user And txtPassword = !Password Then
        'place code to here to pass the
        'success to the calling sub
        'setting a global var is the easiest
            LoginSucceeded = True
            Me.Hide
            menu_utama.Enabled = True
            menu_utama.Show
            menu_utama.Label5.Caption = !user
            menu_utama.SSTab1.Tab = 1
                txtUserName = ""
    txtPassword = ""
            a = 1
        End If
        .MoveNext
    Loop
    If a = 0 Then
        MsgBox "Invalid User and Password, try again!", , "Login"
        txtPassword.SetFocus
'        SendKeys "{Home}+{End}"
    End If
End If
End With
End Sub

Private Sub Form_Load()
Call open_db
End Sub
