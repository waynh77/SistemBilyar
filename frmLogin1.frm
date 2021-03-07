VERSION 5.00
Begin VB.Form frmLogin1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administrator Login"
   ClientHeight    =   1155
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   682.412
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   2
      Top             =   660
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   3
      Top             =   660
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   165
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   0
      Top             =   180
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin1"
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
    Me.Hide
    menu_utama.Visible = True
    menu_utama.SSTab1.Tab = 1
End Sub

Private Sub cmdOK_Click()
    'check for correct password
    If txtPassword = "a" Then
        'place code to here to pass the
        'success to the calling sub
        'setting a global var is the easiest
        LoginSucceeded = True
        txtPassword = ""
        Me.Hide
        menu_utama.SSTab1.Tab = 0
        menu_utama.Visible = True
        Else
        MsgBox "Invalid Password, try again!", , "Login"
        txtPassword.SetFocus
'        SendKeys "{Home}+{End}"
    End If
End Sub
