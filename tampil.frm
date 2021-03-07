VERSION 5.00
Begin VB.Form tampil 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DATA BASE PRODUK"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11625
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   11625
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu exit 
      Caption         =   "&Keluar"
   End
End
Attribute VB_Name = "tampil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub exit_Click()
If tampil.Caption = "LAPORAN PENJUALAN" Then
End
Else
Me.Hide
tampil.Cls
End If
End Sub
