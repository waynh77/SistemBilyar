VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form menu_utama 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SISTEM PENJUALAN"
   ClientHeight    =   5685
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5655
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "menu_utama.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   2
      Left            =   3960
      Top             =   960
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   3120
      Top             =   960
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   192
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "DATABASE PRODUK"
      TabPicture(0)   =   "menu_utama.frx":1CCA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Command11"
      Tab(0).Control(1)=   "Command7"
      Tab(0).Control(2)=   "Command5"
      Tab(0).Control(3)=   "Command4"
      Tab(0).Control(4)=   "Command3"
      Tab(0).Control(5)=   "Command2"
      Tab(0).Control(6)=   "Command1"
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "RESTAURANT"
      TabPicture(1)   =   "menu_utama.frx":1CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Line1(0)"
      Tab(1).Control(1)=   "Line2"
      Tab(1).Control(2)=   "Label1"
      Tab(1).Control(3)=   "Label2"
      Tab(1).Control(4)=   "Line1(1)"
      Tab(1).Control(5)=   "Line1(2)"
      Tab(1).Control(6)=   "Line1(3)"
      Tab(1).Control(7)=   "Line1(4)"
      Tab(1).Control(8)=   "command6(0)"
      Tab(1).Control(9)=   "command6(1)"
      Tab(1).Control(10)=   "command6(2)"
      Tab(1).Control(11)=   "command6(3)"
      Tab(1).Control(12)=   "command6(4)"
      Tab(1).Control(13)=   "command6(5)"
      Tab(1).Control(14)=   "command6(6)"
      Tab(1).Control(15)=   "command6(7)"
      Tab(1).Control(16)=   "command6(8)"
      Tab(1).Control(17)=   "command6(9)"
      Tab(1).Control(18)=   "command6(25)"
      Tab(1).Control(19)=   "command6(26)"
      Tab(1).Control(20)=   "command6(27)"
      Tab(1).Control(21)=   "command6(28)"
      Tab(1).Control(22)=   "command6(29)"
      Tab(1).Control(23)=   "command6(10)"
      Tab(1).Control(24)=   "command6(11)"
      Tab(1).Control(25)=   "command6(12)"
      Tab(1).Control(26)=   "command6(13)"
      Tab(1).Control(27)=   "command6(14)"
      Tab(1).Control(28)=   "command6(15)"
      Tab(1).Control(29)=   "command6(16)"
      Tab(1).Control(30)=   "command6(17)"
      Tab(1).Control(31)=   "command6(18)"
      Tab(1).Control(32)=   "command6(19)"
      Tab(1).Control(33)=   "command6(20)"
      Tab(1).Control(34)=   "command6(21)"
      Tab(1).Control(35)=   "command6(22)"
      Tab(1).Control(36)=   "command6(23)"
      Tab(1).Control(37)=   "command6(24)"
      Tab(1).ControlCount=   38
      TabCaption(2)   =   "BAR"
      TabPicture(2)   =   "menu_utama.frx":1D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Timer12"
      Tab(2).Control(1)=   "Timer11"
      Tab(2).Control(2)=   "Timer4"
      Tab(2).Control(3)=   "Timer3"
      Tab(2).Control(4)=   "Command8(29)"
      Tab(2).Control(5)=   "Command8(28)"
      Tab(2).Control(6)=   "Command8(27)"
      Tab(2).Control(7)=   "Command8(26)"
      Tab(2).Control(8)=   "Command8(25)"
      Tab(2).Control(9)=   "Command8(24)"
      Tab(2).Control(10)=   "Command8(23)"
      Tab(2).Control(11)=   "Command8(22)"
      Tab(2).Control(12)=   "Command8(21)"
      Tab(2).Control(13)=   "Command8(20)"
      Tab(2).Control(14)=   "Command8(19)"
      Tab(2).Control(15)=   "Command8(18)"
      Tab(2).Control(16)=   "Command8(17)"
      Tab(2).Control(17)=   "Command8(16)"
      Tab(2).Control(18)=   "Command8(15)"
      Tab(2).Control(19)=   "Command8(14)"
      Tab(2).Control(20)=   "Command8(13)"
      Tab(2).Control(21)=   "Command8(12)"
      Tab(2).Control(22)=   "Command8(11)"
      Tab(2).Control(23)=   "Command8(10)"
      Tab(2).Control(24)=   "Command8(9)"
      Tab(2).Control(25)=   "Command8(8)"
      Tab(2).Control(26)=   "Command8(7)"
      Tab(2).Control(27)=   "Command8(6)"
      Tab(2).Control(28)=   "Command8(5)"
      Tab(2).Control(29)=   "Command8(4)"
      Tab(2).Control(30)=   "Command8(3)"
      Tab(2).Control(31)=   "Command8(2)"
      Tab(2).Control(32)=   "Command8(1)"
      Tab(2).Control(33)=   "Command8(0)"
      Tab(2).Control(34)=   "Label3"
      Tab(2).Control(35)=   "Line3(5)"
      Tab(2).Control(36)=   "Line3(4)"
      Tab(2).Control(37)=   "Line3(3)"
      Tab(2).Control(38)=   "Line3(2)"
      Tab(2).Control(39)=   "Line3(1)"
      Tab(2).Control(40)=   "Line3(0)"
      Tab(2).ControlCount=   41
      TabCaption(3)   =   "POOL"
      TabPicture(3)   =   "menu_utama.frx":1D1E
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label6"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Line4"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Shape1"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Line5(0)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Line5(1)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Line5(2)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Line5(3)"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Label7(0)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Label7(1)"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Label8"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "Label10(1)"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "Label10(2)"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "Label10(3)"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "Label10(4)"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "Label10(5)"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "Label10(6)"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "Label10(7)"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).Control(17)=   "Label10(8)"
      Tab(3).Control(17).Enabled=   0   'False
      Tab(3).Control(18)=   "Label10(9)"
      Tab(3).Control(18).Enabled=   0   'False
      Tab(3).Control(19)=   "Label10(0)"
      Tab(3).Control(19).Enabled=   0   'False
      Tab(3).Control(20)=   "Command9(0)"
      Tab(3).Control(20).Enabled=   0   'False
      Tab(3).Control(21)=   "Timer5"
      Tab(3).Control(21).Enabled=   0   'False
      Tab(3).Control(22)=   "Timer6"
      Tab(3).Control(22).Enabled=   0   'False
      Tab(3).Control(23)=   "Command9(1)"
      Tab(3).Control(23).Enabled=   0   'False
      Tab(3).Control(24)=   "Command9(2)"
      Tab(3).Control(24).Enabled=   0   'False
      Tab(3).Control(25)=   "Command9(3)"
      Tab(3).Control(25).Enabled=   0   'False
      Tab(3).Control(26)=   "Command9(4)"
      Tab(3).Control(26).Enabled=   0   'False
      Tab(3).Control(27)=   "Command9(5)"
      Tab(3).Control(27).Enabled=   0   'False
      Tab(3).Control(28)=   "Command9(6)"
      Tab(3).Control(28).Enabled=   0   'False
      Tab(3).Control(29)=   "Command9(7)"
      Tab(3).Control(29).Enabled=   0   'False
      Tab(3).Control(30)=   "Command9(8)"
      Tab(3).Control(30).Enabled=   0   'False
      Tab(3).Control(31)=   "Timer7"
      Tab(3).Control(31).Enabled=   0   'False
      Tab(3).Control(32)=   "Timer8"
      Tab(3).Control(32).Enabled=   0   'False
      Tab(3).Control(33)=   "Timer9"
      Tab(3).Control(33).Enabled=   0   'False
      Tab(3).Control(34)=   "Timer10"
      Tab(3).Control(34).Enabled=   0   'False
      Tab(3).Control(35)=   "Command9(9)"
      Tab(3).Control(35).Enabled=   0   'False
      Tab(3).Control(36)=   "Timer13(0)"
      Tab(3).Control(36).Enabled=   0   'False
      Tab(3).Control(37)=   "Timer13(1)"
      Tab(3).Control(37).Enabled=   0   'False
      Tab(3).Control(38)=   "Timer13(2)"
      Tab(3).Control(38).Enabled=   0   'False
      Tab(3).Control(39)=   "Timer13(3)"
      Tab(3).Control(39).Enabled=   0   'False
      Tab(3).Control(40)=   "Timer13(4)"
      Tab(3).Control(40).Enabled=   0   'False
      Tab(3).Control(41)=   "Timer13(5)"
      Tab(3).Control(41).Enabled=   0   'False
      Tab(3).Control(42)=   "Timer13(6)"
      Tab(3).Control(42).Enabled=   0   'False
      Tab(3).Control(43)=   "Timer13(7)"
      Tab(3).Control(43).Enabled=   0   'False
      Tab(3).Control(44)=   "Timer13(8)"
      Tab(3).Control(44).Enabled=   0   'False
      Tab(3).Control(45)=   "Timer13(9)"
      Tab(3).Control(45).Enabled=   0   'False
      Tab(3).ControlCount=   46
      Begin VB.Timer Timer13 
         Index           =   9
         Interval        =   60000
         Left            =   4320
         Top             =   3360
      End
      Begin VB.Timer Timer13 
         Index           =   8
         Interval        =   60000
         Left            =   3360
         Top             =   3360
      End
      Begin VB.Timer Timer13 
         Index           =   7
         Interval        =   60000
         Left            =   2400
         Top             =   3360
      End
      Begin VB.Timer Timer13 
         Index           =   6
         Interval        =   60000
         Left            =   1440
         Top             =   3360
      End
      Begin VB.Timer Timer13 
         Index           =   5
         Interval        =   60000
         Left            =   480
         Top             =   3480
      End
      Begin VB.Timer Timer13 
         Index           =   4
         Interval        =   60000
         Left            =   4320
         Top             =   1680
      End
      Begin VB.Timer Timer13 
         Index           =   3
         Interval        =   60000
         Left            =   3360
         Top             =   1680
      End
      Begin VB.Timer Timer13 
         Index           =   2
         Interval        =   60000
         Left            =   2400
         Top             =   1680
      End
      Begin VB.Timer Timer13 
         Index           =   1
         Interval        =   60000
         Left            =   1440
         Top             =   1800
      End
      Begin VB.Timer Timer13 
         Index           =   0
         Interval        =   60000
         Left            =   600
         Top             =   1680
      End
      Begin VB.CommandButton Command9 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   9
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   3240
         Width           =   735
      End
      Begin VB.Timer Timer12 
         Interval        =   200
         Left            =   -73920
         Top             =   720
      End
      Begin VB.Timer Timer11 
         Interval        =   200
         Left            =   -74640
         Top             =   840
      End
      Begin VB.CommandButton Command11 
         Caption         =   "&Seting Harga"
         Height          =   735
         Left            =   -72120
         TabIndex        =   85
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Timer Timer10 
         Enabled         =   0   'False
         Interval        =   3
         Left            =   3480
         Top             =   600
      End
      Begin VB.Timer Timer9 
         Enabled         =   0   'False
         Interval        =   2
         Left            =   2880
         Top             =   600
      End
      Begin VB.Timer Timer8 
         Enabled         =   0   'False
         Interval        =   2
         Left            =   2280
         Top             =   600
      End
      Begin VB.Timer Timer7 
         Interval        =   2
         Left            =   1680
         Top             =   600
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Command9"
         Height          =   1215
         Index           =   8
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   3240
         Width           =   735
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Command9"
         Height          =   1215
         Index           =   7
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   3240
         Width           =   735
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Command9"
         Height          =   1215
         Index           =   6
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   3240
         Width           =   735
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Command9"
         Height          =   1215
         Index           =   5
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   3240
         Width           =   735
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Command9"
         Height          =   1215
         Index           =   4
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   1560
         Width           =   735
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Command9"
         Height          =   1215
         Index           =   3
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   1560
         Width           =   735
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Command9"
         Height          =   1215
         Index           =   2
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   1560
         Width           =   735
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Command9"
         Height          =   1215
         Index           =   1
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   1560
         Width           =   735
      End
      Begin VB.Timer Timer6 
         Interval        =   2
         Left            =   3000
         Top             =   0
      End
      Begin VB.Timer Timer5 
         Interval        =   2
         Left            =   3840
         Top             =   0
      End
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   2
         Left            =   -71520
         Top             =   720
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   2
         Left            =   -70560
         Top             =   840
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   495
         Index           =   29
         Left            =   -70560
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   4020
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   495
         Index           =   28
         Left            =   -71400
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   4020
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   495
         Index           =   27
         Left            =   -72240
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   4020
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   495
         Index           =   26
         Left            =   -73080
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   4020
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   495
         Index           =   25
         Left            =   -73920
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   4020
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   495
         Index           =   24
         Left            =   -74760
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   4020
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   495
         Index           =   23
         Left            =   -70560
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   3300
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   495
         Index           =   22
         Left            =   -71400
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   3300
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   495
         Index           =   21
         Left            =   -72240
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   3300
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   495
         Index           =   20
         Left            =   -73080
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   3300
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   495
         Index           =   19
         Left            =   -73920
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   3300
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   495
         Index           =   18
         Left            =   -74760
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   3300
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   495
         Index           =   17
         Left            =   -70560
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   2580
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   495
         Index           =   16
         Left            =   -71400
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   2580
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   495
         Index           =   15
         Left            =   -72240
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   2580
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   495
         Index           =   14
         Left            =   -73080
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   2580
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   495
         Index           =   13
         Left            =   -73920
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   2580
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   495
         Index           =   12
         Left            =   -74760
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   2580
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   495
         Index           =   11
         Left            =   -70560
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   1860
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   495
         Index           =   10
         Left            =   -71400
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   1860
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   495
         Index           =   9
         Left            =   -72240
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   1860
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   495
         Index           =   8
         Left            =   -73080
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   1860
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   495
         Index           =   7
         Left            =   -73920
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   1860
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   495
         Index           =   6
         Left            =   -74760
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   1860
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   495
         Index           =   5
         Left            =   -70560
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   1140
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   495
         Index           =   4
         Left            =   -71400
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   1140
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   495
         Index           =   3
         Left            =   -72240
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   1140
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   495
         Index           =   2
         Left            =   -73080
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   1140
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   495
         Index           =   1
         Left            =   -73920
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   1140
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   495
         Index           =   0
         Left            =   -74760
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   1140
         Width           =   615
      End
      Begin VB.CommandButton Command7 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   -70800
         TabIndex        =   36
         Top             =   4260
         Width           =   855
      End
      Begin VB.CommandButton command6 
         Caption         =   "MEJA 25"
         Height          =   495
         Index           =   24
         Left            =   -74760
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   4020
         Width           =   615
      End
      Begin VB.CommandButton command6 
         Caption         =   "MEJA 24"
         Height          =   495
         Index           =   23
         Left            =   -70560
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   3300
         Width           =   615
      End
      Begin VB.CommandButton command6 
         Caption         =   "MEJA 23"
         Height          =   495
         Index           =   22
         Left            =   -71400
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   3300
         Width           =   615
      End
      Begin VB.CommandButton command6 
         Caption         =   "MEJA 22"
         Height          =   495
         Index           =   21
         Left            =   -72240
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   3300
         Width           =   615
      End
      Begin VB.CommandButton command6 
         Caption         =   "MEJA 21"
         Height          =   495
         Index           =   20
         Left            =   -73080
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   3300
         Width           =   615
      End
      Begin VB.CommandButton command6 
         Caption         =   "MEJA 20"
         Height          =   495
         Index           =   19
         Left            =   -73920
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   3300
         Width           =   615
      End
      Begin VB.CommandButton command6 
         Caption         =   "MEJA 19"
         Height          =   495
         Index           =   18
         Left            =   -74760
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   3300
         Width           =   615
      End
      Begin VB.CommandButton command6 
         Caption         =   "MEJA 18"
         Height          =   495
         Index           =   17
         Left            =   -70560
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   2580
         Width           =   615
      End
      Begin VB.CommandButton command6 
         Caption         =   "MEJA 17"
         Height          =   495
         Index           =   16
         Left            =   -71400
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2580
         Width           =   615
      End
      Begin VB.CommandButton command6 
         Caption         =   "MEJA 16"
         Height          =   495
         Index           =   15
         Left            =   -72240
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2580
         Width           =   615
      End
      Begin VB.CommandButton command6 
         Caption         =   "MEJA 15"
         Height          =   495
         Index           =   14
         Left            =   -73080
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   2580
         Width           =   615
      End
      Begin VB.CommandButton command6 
         Caption         =   "MEJA 14"
         Height          =   495
         Index           =   13
         Left            =   -73920
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2580
         Width           =   615
      End
      Begin VB.CommandButton command6 
         Caption         =   "MEJA 13"
         Height          =   495
         Index           =   12
         Left            =   -74760
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2580
         Width           =   615
      End
      Begin VB.CommandButton command6 
         Caption         =   "MEJA 12"
         Height          =   495
         Index           =   11
         Left            =   -70560
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1860
         Width           =   615
      End
      Begin VB.CommandButton command6 
         Caption         =   "MEJA 11"
         Height          =   495
         Index           =   10
         Left            =   -71400
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1860
         Width           =   615
      End
      Begin VB.CommandButton command6 
         Caption         =   "MEJA 30"
         Height          =   495
         Index           =   29
         Left            =   -70560
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   4020
         Width           =   615
      End
      Begin VB.CommandButton command6 
         Caption         =   "MEJA 29"
         Height          =   495
         Index           =   28
         Left            =   -71400
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   4020
         Width           =   615
      End
      Begin VB.CommandButton command6 
         Caption         =   "MEJA 28"
         Height          =   495
         Index           =   27
         Left            =   -72240
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   4020
         Width           =   615
      End
      Begin VB.CommandButton command6 
         Caption         =   "MEJA 27"
         Height          =   495
         Index           =   26
         Left            =   -73080
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   4020
         Width           =   615
      End
      Begin VB.CommandButton command6 
         Caption         =   "MEJA 26"
         Height          =   495
         Index           =   25
         Left            =   -73920
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   4020
         Width           =   615
      End
      Begin VB.CommandButton command6 
         Caption         =   "MEJA 10"
         Height          =   495
         Index           =   9
         Left            =   -72240
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1860
         Width           =   615
      End
      Begin VB.CommandButton command6 
         Caption         =   "MEJA 9"
         Height          =   495
         Index           =   8
         Left            =   -73080
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1860
         Width           =   615
      End
      Begin VB.CommandButton command6 
         Caption         =   "MEJA 8"
         Height          =   495
         Index           =   7
         Left            =   -73920
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1860
         Width           =   615
      End
      Begin VB.CommandButton command6 
         Caption         =   "MEJA 7"
         Height          =   495
         Index           =   6
         Left            =   -74760
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1860
         Width           =   615
      End
      Begin VB.CommandButton command6 
         Caption         =   "MEJA 6"
         Height          =   495
         Index           =   5
         Left            =   -70560
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1140
         Width           =   615
      End
      Begin VB.CommandButton command6 
         Caption         =   "MEJA 5"
         Height          =   495
         Index           =   4
         Left            =   -71400
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1140
         Width           =   615
      End
      Begin VB.CommandButton command6 
         Caption         =   "MEJA 4"
         Height          =   495
         Index           =   3
         Left            =   -72240
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1140
         Width           =   615
      End
      Begin VB.CommandButton command6 
         Caption         =   "MEJA 3"
         Height          =   495
         Index           =   2
         Left            =   -73080
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1140
         Width           =   615
      End
      Begin VB.CommandButton command6 
         Caption         =   "MEJA 2"
         Height          =   495
         Index           =   1
         Left            =   -73920
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1140
         Width           =   615
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "C&ARI"
         Height          =   735
         Left            =   -72120
         TabIndex        =   5
         Top             =   1920
         Width           =   1935
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&CETAK"
         Height          =   735
         Left            =   -74520
         TabIndex        =   4
         Top             =   3000
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&HAPUS"
         Height          =   735
         Left            =   -74520
         TabIndex        =   3
         Top             =   1920
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&EDIT"
         Height          =   735
         Left            =   -72120
         TabIndex        =   2
         Top             =   900
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&TAMBAH"
         Height          =   735
         Left            =   -74520
         TabIndex        =   1
         Top             =   900
         Width           =   1935
      End
      Begin VB.CommandButton command6 
         Caption         =   "MEJA 1"
         Height          =   495
         Index           =   0
         Left            =   -74760
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1140
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Command9"
         Height          =   1215
         Index           =   0
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Label10"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   88
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Label10"
         Height          =   255
         Index           =   9
         Left            =   4200
         TabIndex        =   97
         Top             =   4440
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Label10"
         Height          =   255
         Index           =   8
         Left            =   3240
         TabIndex        =   96
         Top             =   4440
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Label10"
         Height          =   255
         Index           =   7
         Left            =   2280
         TabIndex        =   95
         Top             =   4440
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Label10"
         Height          =   255
         Index           =   6
         Left            =   1320
         TabIndex        =   94
         Top             =   4440
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Label10"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   93
         Top             =   4440
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Label10"
         Height          =   255
         Index           =   4
         Left            =   4200
         TabIndex        =   92
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Label10"
         Height          =   255
         Index           =   3
         Left            =   3240
         TabIndex        =   91
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Label10"
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   90
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Label10"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   89
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Left            =   240
         TabIndex        =   84
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "toc"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   4800
         TabIndex        =   83
         Top             =   960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "tic"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   2520
         TabIndex        =   82
         Top             =   960
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00800080&
         BorderWidth     =   2
         Index           =   3
         X1              =   360
         X2              =   4920
         Y1              =   4725
         Y2              =   4725
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00800080&
         BorderWidth     =   2
         Index           =   2
         X1              =   360
         X2              =   4920
         Y1              =   3900
         Y2              =   3900
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00800080&
         BorderWidth     =   2
         Index           =   1
         X1              =   360
         X2              =   4920
         Y1              =   3045
         Y2              =   3045
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00800080&
         BorderWidth     =   2
         Index           =   0
         X1              =   360
         X2              =   4920
         Y1              =   2205
         Y2              =   2205
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H00C000C0&
         Height          =   255
         Left            =   2640
         Shape           =   3  'Circle
         Top             =   1080
         Width           =   375
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   3
         X1              =   120
         X2              =   1320
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "POOL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   300
         Left            =   4320
         TabIndex        =   72
         Top             =   650
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   300
         Left            =   -72720
         TabIndex        =   69
         Top             =   645
         Width           =   705
      End
      Begin VB.Line Line3 
         BorderColor     =   &H0000C000&
         BorderWidth     =   2
         Index           =   5
         X1              =   -74760
         X2              =   -69960
         Y1              =   4620
         Y2              =   4620
      End
      Begin VB.Line Line3 
         BorderColor     =   &H0000C000&
         BorderWidth     =   2
         Index           =   4
         X1              =   -74760
         X2              =   -69960
         Y1              =   3900
         Y2              =   3900
      End
      Begin VB.Line Line3 
         BorderColor     =   &H0000C000&
         BorderWidth     =   2
         Index           =   3
         X1              =   -74760
         X2              =   -69960
         Y1              =   1740
         Y2              =   1740
      End
      Begin VB.Line Line3 
         BorderColor     =   &H0000C000&
         BorderWidth     =   2
         Index           =   2
         X1              =   -74760
         X2              =   -69960
         Y1              =   2460
         Y2              =   2460
      End
      Begin VB.Line Line3 
         BorderColor     =   &H0000C000&
         BorderWidth     =   2
         Index           =   1
         X1              =   -74760
         X2              =   -69960
         Y1              =   3180
         Y2              =   3180
      End
      Begin VB.Line Line3 
         BorderColor     =   &H0000C000&
         BorderWidth     =   2
         Index           =   0
         X1              =   -74760
         X2              =   -69960
         Y1              =   1020
         Y2              =   1020
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000C0&
         BorderWidth     =   2
         Index           =   4
         X1              =   -74760
         X2              =   -69960
         Y1              =   4620
         Y2              =   4620
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000C0&
         BorderWidth     =   2
         Index           =   3
         X1              =   -74760
         X2              =   -69960
         Y1              =   3900
         Y2              =   3900
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000C0&
         BorderWidth     =   2
         Index           =   2
         X1              =   -74760
         X2              =   -69960
         Y1              =   1020
         Y2              =   1020
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000C0&
         BorderWidth     =   2
         Index           =   1
         X1              =   -74760
         X2              =   -69960
         Y1              =   3180
         Y2              =   3180
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Left            =   -71520
         TabIndex        =   38
         Top             =   780
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "RESTAURANT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   -74760
         TabIndex        =   37
         Top             =   650
         Width           =   1935
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000C0&
         BorderWidth     =   2
         X1              =   -74760
         X2              =   -69960
         Y1              =   1740
         Y2              =   1740
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000C0&
         BorderWidth     =   2
         Index           =   0
         X1              =   -74760
         X2              =   -69960
         Y1              =   2460
         Y2              =   2460
      End
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "WaynhSoft"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   4560
      TabIndex        =   87
      Top             =   5400
      Width           =   1035
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   " "
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
      Left            =   720
      TabIndex        =   71
      Top             =   0
      Width           =   75
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Kasir : "
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
      TabIndex        =   70
      Top             =   0
      Width           =   615
   End
   Begin VB.Menu dbadmin_mnu 
      Caption         =   "&Data Admin"
      Begin VB.Menu dbproduk 
         Caption         =   "DataBase Produk"
      End
      Begin VB.Menu db_trans 
         Caption         =   "File Transaksi"
         Begin VB.Menu food 
            Caption         =   "Makanan dan Minuman"
            Begin VB.Menu transrest 
               Caption         =   "Restaurant"
            End
            Begin VB.Menu transbar 
               Caption         =   "Bar"
            End
            Begin VB.Menu Transpool 
               Caption         =   "Pool"
            End
         End
         Begin VB.Menu Pool 
            Caption         =   "Pool"
         End
      End
      Begin VB.Menu dbuser 
         Caption         =   "DataBase User"
      End
   End
   Begin VB.Menu lap 
      Caption         =   "&Laporan"
   End
   Begin VB.Menu calk 
      Caption         =   "Kalkulator"
   End
   Begin VB.Menu gantiuser 
      Caption         =   "&Ganti User"
   End
   Begin VB.Menu carimnu 
      Caption         =   "cari data"
      Visible         =   0   'False
   End
   Begin VB.Menu keluar 
      Caption         =   "&Keluar"
   End
End
Attribute VB_Name = "menu_utama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Integer

Private Sub calk_Click()
AppActivate Shell("calc.exe", 1)
End Sub

Private Sub carimnu_Click()
test_sql.Show
Me.Hide
End Sub

Private Sub Command1_Click()
        DBproduk_form.cmdproses.Caption = "&Tambah"
        DBproduk_form.cmdbatal.Caption = "&Keluar"
        DBproduk_form.Caption = "TAMBAH DATA"
        DBproduk_form.Visible = True
        menu_utama.Visible = False
End Sub

Private Sub Command11_Click()
    Me.Hide
    frmseting.Show
End Sub

Private Sub Command2_Click()
        DBproduk_form.cmdproses.Caption = "&Edit"
        DBproduk_form.cmdbatal.Caption = "&Keluar"
        DBproduk_form.Caption = "EDIT DATA"
        DBproduk_form.Visible = True
        menu_utama.Visible = False
End Sub

Private Sub Command3_Click()
        DBproduk_form.cmdproses.Caption = "&Hapus"
        DBproduk_form.cmdbatal.Caption = "&Keluar"
        DBproduk_form.Caption = "HAPUS DATA"
        DBproduk_form.Visible = True
        menu_utama.Visible = False
End Sub

Private Sub Command4_Click()
'Printer.Show
cetakkeprinter
'With DBProdukRpt
'    .WindowState = 2
'    DBproduk_form.DataProduk.Refresh
'    .Show
'End With
End Sub
Private Function rkanan(ndata, cformat) As String
    rkanan = Format(ndata, cformat)
    rkanan = Space(Len(cformat) - Len(rkanan)) + rkanan
End Function

Private Sub cetakkelayar()
    Dim mno, mhal, mbaris As Integer
    Dim mjumlah As Double
    Dim mgrs As String
'    tampil.CurrentX = 0
'    tampil.CurrentY = 0
    tampil.Font = "courier new"
    tampil.Cls
    tampil.Show
    tampil.Caption = "DATA BASE PRODUK"
    With DBproduk_form.DataProduk
    If .Recordset.RecordCount = 0 Then
        x = MsgBox("Maaf data masih kosong, silahkan diisi dahulu !!!", vbOKOnly, "DATA KOSONG")
        Exit Sub
    End If
    .Recordset.MoveFirst
    mno = 0
    mhal = 0
    Do While Not .Recordset.EOF
        mhal = mhal + 1
        tampil.Print
        tampil.Print
        tampil.FontBold = True
        tampil.FontSize = 10
        tampil.Print Tab(5); "DATA BASE PRODUK"
        tampil.Print Tab(5); "Tanggal Cetak : "; Format(Date, "dd/mm/yyyy")
        tampil.FontBold = False
        tampil.FontSize = 8
        tampil.Print Tab(95); "Hal :"; Format(mhal, "###")
        mgrs = String$(100, "-")
        tampil.Print Tab(5); mgrs
        tampil.Print Tab(5); "No.";
        tampil.Print Tab(10); "KODE";
        tampil.Print Tab(20); "JENIS";
        tampil.Print Tab(45); "KLASIFIKASI";
        tampil.Print Tab(65); "NAMA PRODUK";
        tampil.Print Tab(95); "HARGA"
        tampil.Print Tab(5); mgrs
        mbaris = 0
        Do While Not .Recordset.EOF And mbaris <= 49
            mno = mno + 1
            tampil.Print Tab(5); rkanan(mno, "###");
            tampil.Print Tab(10); .Recordset!Kode_produk;
            tampil.Print Tab(20); .Recordset!jenis_produk;
            tampil.Print Tab(45); .Recordset!KLASIFIKASI_produk;
            tampil.Print Tab(65); .Recordset!Nama_produk;
            tampil.Print Tab(95); rkanan(.Recordset!Harga_produk, "###,###,###")
            mbaris = mbaris + 1
            .Recordset.MoveNext
        Loop
        tampil.Print Tab(5); mgrs
 '       tampil.NewPage
    Loop
    End With
'    tampil.EndDoc
End Sub

Private Sub cetakkeprinter()
    Dim mno, mhal, mbaris As Integer
    Dim mjumlah As Double
    Dim mgrs As String
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.Font = "courier new"
'    printer.Cls
'    printer.Show
'    printer.Caption = "DATA BASE PRODUK"
    With DBproduk_form.DataProduk
    If .Recordset.RecordCount = 0 Then
        x = MsgBox("Maaf data masih kosong, silahkan diisi dahulu !!!", vbOKOnly, "DATA KOSONG")
        Exit Sub
    End If
    .Recordset.MoveFirst
    mno = 0
    mhal = 0
    Do While Not .Recordset.EOF
        mhal = mhal + 1
        Printer.Print
        Printer.Print
        Printer.FontBold = True
        Printer.FontSize = 10
        Printer.Print Tab(5); "DATA BASE PRODUK"
        Printer.Print Tab(5); "Tanggal Cetak : "; Format(Date, "dd/mm/yyyy")
        Printer.FontBold = False
        Printer.FontSize = 8
        Printer.Print Tab(95); "Hal :"; Format(mhal, "###")
        mgrs = String$(100, "-")
        Printer.Print Tab(5); mgrs
        Printer.Print Tab(5); "No.";
        Printer.Print Tab(10); "KODE";
        Printer.Print Tab(20); "JENIS";
        Printer.Print Tab(45); "KLASIFIKASI";
        Printer.Print Tab(65); "NAMA PRODUK";
        Printer.Print Tab(95); "HARGA"
        Printer.Print Tab(5); mgrs
        mbaris = 0
        Do While Not .Recordset.EOF And mbaris <= 49
            mno = mno + 1
            Printer.Print Tab(5); rkanan(mno, "###");
            Printer.Print Tab(10); .Recordset!Kode_produk;
            Printer.Print Tab(20); .Recordset!jenis_produk;
            Printer.Print Tab(45); .Recordset!KLASIFIKASI_produk;
            Printer.Print Tab(65); .Recordset!Nama_produk;
            Printer.Print Tab(95); rkanan(.Recordset!Harga_produk, "###,###,###")
            mbaris = mbaris + 1
            .Recordset.MoveNext
        Loop
        Printer.Print Tab(5); mgrs
 '       printer.NewPage
    Loop
    End With
    Printer.EndDoc
End Sub

Private Sub Command5_Click()
    DBproduk_form.cmdproses.Caption = "&Cari"
    DBproduk_form.cmdbatal.Caption = "&Keluar"
    DBproduk_form.Caption = "Cari DATA"
    DBproduk_form.Visible = True
    menu_utama.Visible = False
End Sub

Private Sub Command6_Click(Index As Integer)
    Dim y As String
    Transaksi_form.List2.Clear
    Transaksi_form.Caption = "MEJA " & Index + 1
    Transaksi_form.Label1.Caption = "Transaksi Meja " & Index + 1
    Transaksi_form.Label12.Caption = "Meja" & Index + 1
    Transaksi_form.Combo1.Text = "Pilih Salah Satu"
    menu_utama.Visible = False
    Transaksi_form.Visible = True
    With Transaksi_form
    .user.Caption = Label5.Caption
    .Text1 = 0
    .Command3.Enabled = True
    .Command4.Enabled = True
    .Label17 = 0
    .Label1.ForeColor = QBColor(12)
    .Label16.ForeColor = QBColor(12)
    .Label17.ForeColor = QBColor(12)
    .Label4 = "Meja"
    .Label18.Caption = Index
    .List2.Clear
    .Data2.RecordSource = "select * from temp_trans"
    .Data2.Refresh
    Transaksi_form.List2.AddItem "Meja" & Index + 1
    If .Data2.Recordset.RecordCount <> 0 Then
        .Data2.Recordset.MoveFirst
        total = 0
        .Data2.RecordSource = "select * from temp_trans where no_meja=" & "'" & .Label12 & "'"
        .Data2.Refresh
        If Not .Data2.Recordset.BOF Then
            Transaksi_form.Show
            .Text1 = .Data2.Recordset!discount
            .Text1.Enabled = False
            .pembeli.Caption = .Data2.Recordset!nama_pembeli
            Do While Not .Data2.Recordset.EOF
            If .Data2.Recordset!no_meja = "Meja" & Index + 1 Then
                .DB_Prod.Recordset.Index = "idxproduk"
                .DB_Prod.Recordset.Seek "=", .Data2.Recordset!Kode_produk
                .List2.AddItem .DB_Prod.Recordset!Nama_produk
                .List2.AddItem "   Price @ : Rp " & Format(.DB_Prod.Recordset!Harga_produk, "###,###,###")
                .List2.AddItem "      Qty : " & .Data2.Recordset!qty & "  Sub Total : Rp " & Format(.Data2.Recordset!qty * .DB_Prod.Recordset!Harga_produk, "###,###,###")
                total = total + .Data2.Recordset!qty * .DB_Prod.Recordset!Harga_produk
            End If
            .Data2.Recordset.MoveNext
            Loop
        Else
            frmcost.Show
            Me.Hide
            Transaksi_form.Hide
        End If
    Else
        frmcost.Show
        Me.Hide
        Transaksi_form.Hide
    End If
    .pembeli.ForeColor = QBColor(12)
    .user.ForeColor = QBColor(12)
    .Label24.ForeColor = QBColor(12)
    .Label25.ForeColor = QBColor(12)
    .Label17 = rkanan(total, "###,###,###")
    frmseting.Data1.Refresh
    '.Text1 = 0
    .Label27 = 0
    .Label30 = 0
    .Label26 = "(" & rkanan(frmseting.Data1.Recordset!tax, "##.##") & "%)"
    .Label28 = "(" & rkanan(.Text1, "##.##") & "%)"
    .Label27 = rkanan((frmseting.Data1.Recordset!tax * total / 100), "###,###,###")
    .Label30 = rkanan((total * .Text1 / 100), "###,###,###")
    .Label25 = rkanan(total + (frmseting.Data1.Recordset!tax * total / 100) - (total * .Text1 / 100), "###,###,###")
    .Label31 = "meja"
    End With
    Transaksi_form.Height = 5955
End Sub

Private Sub Command7_Click()
If Transaksi_form.Data2.Recordset.RecordCount = 0 And frmPool.dt1.Recordset.RecordCount = 0 Then
    x = MsgBox("Laporan Harian Akan Dicetak, Apakah Printer Sudah Siap?", vbOKCancel, "Cetak Laporan")
    If x = vbOK Then
        tampillaporanharian
        End
    End If
Else
    x = MsgBox("Masih ada transaksi yg belum selesai...!!!", vbOKOnly, "Peringatan!")
End If
End Sub

Private Sub Command8_Click(Index As Integer)
    Dim y As String
    Transaksi_form.Show
    Transaksi_form.List2.Clear
    Transaksi_form.Caption = "BAR " & Index + 1
    Transaksi_form.Label1.Caption = "Transaksi Bar " & Index + 1
    Transaksi_form.Label12.Caption = "Bar" & Index + 1
    Transaksi_form.Combo1.Text = "Pilih Salah Satu"
    menu_utama.Visible = False
    Transaksi_form.Visible = True
    With Transaksi_form
    .user.Caption = Label5.Caption
    .Text1 = 0
    .Command3.Enabled = True
    .Command4.Enabled = True
    .Label17 = 0
    .Label1.ForeColor = QBColor(12)
    .Label16.ForeColor = QBColor(12)
    .Label17.ForeColor = QBColor(12)
    .Label4 = "Bar"
    .Label18.Caption = Index
    .List2.Clear
    .Data2.RecordSource = "select * from temp_trans"
    .Data2.Refresh
    Transaksi_form.List2.AddItem "Bar" & Index + 1
    If .Data2.Recordset.RecordCount <> 0 Then
    .Data2.Recordset.MoveFirst
    total = 0
    .Data2.RecordSource = "select * from temp_trans where no_meja=" & "'" & .Label12 & "'"
    .Data2.Refresh
    If Not .Data2.Recordset.BOF Then
        .Text1 = .Data2.Recordset!discount
        .Text1.Enabled = False
        .pembeli.Caption = .Data2.Recordset!nama_pembeli
        Do While Not .Data2.Recordset.EOF
            If .Data2.Recordset!no_meja = "Bar" & Index + 1 Then
                .DB_Prod.Recordset.Index = "idxproduk"
                .DB_Prod.Recordset.Seek "=", .Data2.Recordset!Kode_produk
                .List2.AddItem .DB_Prod.Recordset!Nama_produk
                .List2.AddItem "   Price @ : Rp " & Format(.DB_Prod.Recordset!Harga_produk, "###,###,###")
                .List2.AddItem "      Qty : " & .Data2.Recordset!qty & "  Sub Total : Rp " & Format(.Data2.Recordset!qty * .DB_Prod.Recordset!Harga_produk, "###,###,###")
                total = total + .Data2.Recordset!qty * .DB_Prod.Recordset!Harga_produk
            End If
            .Data2.Recordset.MoveNext
        Loop
    Else
            frmcost.Show
            Me.Hide
            Transaksi_form.Hide
        End If
    Else
        frmcost.Show
        Me.Hide
        Transaksi_form.Hide
    End If
    .pembeli.ForeColor = QBColor(12)
    .user.ForeColor = QBColor(12)
    .Label24.ForeColor = QBColor(12)
    .Label25.ForeColor = QBColor(12)
    .Label17 = rkanan(total, "###,###,###")
    frmseting.Data1.Refresh
    '.Text1 = 0
    .Label27 = 0
    .Label30 = 0
    .Label26 = "(" & rkanan(frmseting.Data1.Recordset!tax, "##.##") & "%)"
    .Label28 = "(" & rkanan(.Text1, "##.##") & "%)"
    .Label27 = rkanan((frmseting.Data1.Recordset!tax * total / 100), "###,###,###")
    .Label30 = rkanan((total * .Text1 / 100), "###,###,###")
    .Label25 = rkanan((total + (frmseting.Data1.Recordset!tax * total / 100) - (total * .Text1 / 100)), "###,###,###")
    .Label31 = "bar"
    End With
    Transaksi_form.Height = 5955
End Sub

Private Sub Command9_Click(Index As Integer)
Dim hitpool, hitmakanan, subtotal, hittax, hitdiscount, gt, hargaperjam, hitmenit As Double
Dim nil1 As Double
Dim nil2 As Double
Dim nil3 As Double
'Select Case Index
hitpool = 0
hitmakanan = 0
subtotal = 0
hittax = 0
hitdiscount = 0
gt = 0
hargaperjam = 0
hitmenit = 0
frmseting.Data1.Refresh
frmPool.dt1.Refresh
If Index = 9 Then
    hargaperjam = frmseting.Data1.Recordset!spool
Else
    hargaperjam = frmseting.Data1.Recordset!hargapool
End If
    Label8 = "Pool" & Index + 1
    With frmPool
    .cmdpool(0).Enabled = True
    .cmdpool(1).Enabled = False
    .cmdpool(2).Enabled = False
    .List1.Clear
    .List1.AddItem ("Pool" & Index + 1)
    .Label1 = 0
    .Label2 = 0
    .Label3 = 0
    .Label4 = 0
    .Label5 = 0
    .lbl6 = Date
    .lbl11 = 0
    .Text3 = 0
    .lblpool = "Transaksi Pool " & Index + 1
    .lbl17 = rkanan(hargaperjam / 60, "###,###.##")
    .lbl18 = Index + 1
    .lbl19 = "Pool" & .lbl18
    .lblopt = Label5
    .Text1.Enabled = False
    .Text2.Enabled = False
    .lbl13 = rkanan(hargaperjam, "###,###,###")
    a = 0
    If Not .dt1.Recordset.BOF Then
        .dt1.Recordset.MoveFirst
        Do While Not .dt1.Recordset.EOF
            If .dt1.Recordset!no_pool = Label8 Then
                .lblcost = .dt1.Recordset!nama_costumer
                .Text1 = .dt1.Recordset!waktu_mulai
                .Text3 = .dt1.Recordset!discount
                a = 1
                If .dt1.Recordset!waktu_akhir <> blank Then
                    .Text2 = .dt1.Recordset!waktu_akhir
                    'hitmenit = (Minute(.Text2) - Minute(.Text1)) + ((Hour(.Text2) - Hour(.Text1)) * 60)
                    .lbl11 = .dt1.Recordset!lama_main
                    If .lbl11 < 60 Then
                        hitpool = hargaperjam
                    Else
                        nil1 = (Val(.dt1.Recordset!lama_main) - 60) / 30
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
                        hitpool = hargaperjam + (nil3 * (hargaperjam / 2))
                    End If
                    .lbl15 = rkanan(hitpool, "###,###,###")
                    .Label1 = rkanan(hitmakanan, "###,###,###")
                    subtotal = hitmakanan + hitpool
                    .Label2 = rkanan(subtotal, "###,###,###")
                    .lbl14(3) = "Tax & Service(" & rkanan(frmseting.Data1.Recordset!tax, "##.##") & "%)"
                    .lbl14(5) = "Discount(" & rkanan(.Text3, "##.##") & "%)"
                    hittax = frmseting.Data1.Recordset!tax * subtotal / 100
                    .Label3 = rkanan(hittax, "###,###,###")
                    hitdiscount = subtotal * .Text3 / 100
                    .Label4 = rkanan(hitdiscount, "###,###,###")
                    gt = subtotal + hittax - hitdiscount
                    .Label5 = rkanan(gt, "###,###,###")
                    .cmdpool(0).Caption = "Mulai"
                    .cmdpool(0).Enabled = False
                    .cmdpool(1).Enabled = True
                    .cmdpool(2).Enabled = True
                    .Text3.Enabled = False
                Else
                    .Text2 = ""
                    .lbl15 = 0
                    .cmdpool(0).Caption = "Berhenti"
                    .cmdpool(0).Enabled = True
                    .cmdpool(1).Enabled = False
                    .cmdpool(2).Enabled = False
                    .Text3.Enabled = True
                End If
                '.cmdpool(1).Enabled = True
                '.cmdpool(2).Enabled = True
                a = 1
                .dt1.Recordset.MoveLast
            End If
            .dt1.Recordset.MoveNext
        Loop
    Else
        .Text1 = ""
        .Text2 = ""
        .Text3.Enabled = True
        .Text3 = 0
    End If
    If Not Transaksi_form.Data2.Recordset.BOF Then
    With frmPool
    b = 0
    Transaksi_form.Data2.Recordset.MoveFirst
    Do While Not Transaksi_form.Data2.Recordset.EOF
        If Transaksi_form.Data2.Recordset!no_meja = Label8 Then
            Transaksi_form.DB_Prod.Recordset.Index = "idxproduk"
            Transaksi_form.DB_Prod.Recordset.Seek "=", Transaksi_form.Data2.Recordset!Kode_produk
            .List1.AddItem Transaksi_form.DB_Prod.Recordset!Nama_produk
            .List1.AddItem "   Price @ : Rp " & Format(Transaksi_form.DB_Prod.Recordset!Harga_produk, "###,###,###")
            .List1.AddItem "      Qty : " & Transaksi_form.Data2.Recordset!qty & "  Sub Total : Rp " & Format(Transaksi_form.Data2.Recordset!qty * Transaksi_form.DB_Prod.Recordset!Harga_produk, "###,###,###")
            b = 1
            .lblcost = Transaksi_form.Data2.Recordset!nama_pembeli
            hitmakanan = hitmakanan + Transaksi_form.Data2.Recordset!qty * Transaksi_form.DB_Prod.Recordset!Harga_produk
        End If
        Transaksi_form.Data2.Recordset.MoveNext
    Loop
    .Label1 = rkanan(hitmakanan, "###,###,###")
    frmseting.Data1.Refresh
    subtotal = hitmakanan + hitpool
    .Label2 = rkanan(subtotal, "###,###,###")
    .lbl14(3) = "Tax & Service(" & rkanan(frmseting.Data1.Recordset!tax, "##.##") & "%)"
    .lbl14(5) = "Discount(" & rkanan(.Text3, "##.##") & "%)"
    hittax = frmseting.Data1.Recordset!tax * subtotal / 100
    .Label3 = rkanan(hittax, "###,###,###")
    hitdiscount = subtotal * .Text3 / 100
    .Label4 = rkanan(hitdiscount, "###,###,###")
    gt = subtotal + hittax - hitdiscount
    .Label5 = rkanan(gt, "###,###,###")
    .command1.FontBold = True
    End With
    frmPool.Show
    .command1.FontBold = True
    Me.Hide
    If b = 0 Then
        .command1.FontBold = False
    End If
    End If
    If a = 0 And b = 1 Then
        .Text1 = ""
        .Text2 = ""
        .cmdpool(0).Caption = "Mulai"
    End If
    If a = 0 And b = 0 Then
        frmcostpool.Show
        Me.Hide
    Else
        .Show
        Me.Hide
    End If
    End With
End Sub

Private Sub dbproduk_Click()
    frmLogin1.Show
    menu_utama.Visible = False
End Sub

Private Sub dbuser_Click()
    frmloginadmin.Show
    menu_utama.Visible = False
End Sub

Private Sub Form_Activate()
Transaksi_form.Data2.Refresh
    With Transaksi_form.Data2.Recordset
    If Not .RecordCount = 0 Then
    x = 0
    Do While Not x = 30
        .MoveFirst
        Do While Not .EOF
            If !no_meja = "Meja" & x + 1 Then
                Command6(x).FontStrikethru = True
                Command6(x).Font.Bold = True
                Label2 = x
                .MoveLast
            End If
            .MoveNext
        Loop
        x = x + 1
    Loop
    x = 0
    Do While Not x = 30
        .MoveFirst
        Do While Not .EOF
            If !no_meja = "Bar" & x + 1 Then
                Command8(x).FontStrikethru = True
                Command8(x).Font.Bold = True
                Label2 = x
                .MoveLast
            End If
            .MoveNext
        Loop
        x = x + 1
    Loop
    x = 0
    Do While Not x = 10
        .MoveFirst
        a = 0
        Do While Not .EOF
            If !no_meja = "Pool" & x + 1 Then
                Command9(x).FontStrikethru = True
                Command9(x).Font.Bold = True
                Label8 = x
                a = 1
                .MoveLast
            End If
            .MoveNext
        Loop
        
        x = x + 1
    Loop
    End If
    End With
cek_pool
End Sub

Sub cek_pool()
frmPool.dt1.DatabaseName = App.Path & "\master penjualan.mdb"
frmPool.dt1.RecordSource = "temp_pool"
frmPool.dt1.Refresh
With frmPool.dt1.Recordset
If Not .BOF Then
    x = 0
    Do While Not x = 10
        .MoveFirst
        Do While Not .EOF
            If !no_pool = "Pool" & x + 1 Then
                Command9(x).FontStrikethru = True
                Command9(x).Font.Bold = True
                Label8 = x
                .MoveLast
            End If
            .MoveNext
        Loop
        x = x + 1
    Loop
End If
End With

End Sub

Private Sub Form_Load()
Call open_db
Timer1.Enabled = True
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False
Timer9.Enabled = False
Timer10.Enabled = False
Timer11.Enabled = False
Timer12.Enabled = False
x = 0
Do While Not x = Command8.Count
Command8(x).Caption = "BAR  " & x + 1
x = x + 1
Loop
x = 0
Do While Not x = Command9.Count - 1
Command9(x).Caption = "POOL " & x + 1
Timer13(x).Enabled = False
Label10(x) = 0
x = x + 1
Loop
Label10(9) = 0
Timer13(9).Enabled = False
End Sub

Private Sub HapusDB_Click()
    DBproduk_form.cmdproses.Caption = "&Hapus"
    DBproduk_form.cmdbatal.Caption = "&Keluar"
    DBproduk_form.Caption = "TAMBAH DATA"
    DBproduk_form.Visible = True
    menu_utama.Visible = False
End Sub

Private Sub gantiuser_Click()
    frmloginuser.Show
    frmloginuser.txtUserName.SetFocus
    Me.Enabled = False
End Sub

Private Sub keluar_Click()
If Transaksi_form.Data2.Recordset.RecordCount = 0 And frmPool.dt1.Recordset.RecordCount = 0 Then
    x = MsgBox("Apakah Laporan Harian Akan Dicetak ?", vbYesNo, "Cetak Laporan")
    If x = vbYes Then
        cetaklaporanharian
    End If
    End
Else
    x = MsgBox("Masih ada transaksi yg belum selesai...!!!", vbOKOnly, "Peringatan!")
End If
End Sub
Private Sub cetaklaporanharian()
    'x = MsgBox("Sedang Mencetak Laporan", vbInformation, "Cetak Laporan")
    Dim mno, mhal, mbaris As Integer
    Dim mjumlah, total, totqty, grandqty, grandtot, totpool As Double
    Dim mgrs As String
    Printer.Font = "courier new"
'    printer.Cls
'    printer.Show
    Printer.CurrentX = 0
    Printer.CurrentY = 0
'    Printer.Caption = "LAPORAN PENJUALAN"
    DBproduk_form.DataProduk.Refresh
    frmseting.Data1.Refresh
    With DbTransFrm.Data1
    .RecordSource = "select * from transaksi where status=true"
    .Refresh
    If .Recordset.RecordCount = 0 Then
        x = MsgBox("Maaf data masih kosong, silahkan diisi dahulu !!!", vbOKOnly, "DATA KOSONG")
        Exit Sub
    End If
    .Recordset.MoveFirst
    mno = 0
    mhal = 0
    Do While Not .Recordset.EOF
        mhal = mhal + 1
        Printer.Print
        Printer.Print
        Printer.FontBold = True
        Printer.FontSize = 10
        Printer.Print Tab(5); "LAPORAN PENJUALAN HARIAN"
        Printer.Print Tab(5); "Tanggal Cetak : "; Format(Date, "dd/mm/yyyy")
        Printer.FontBold = False
        Printer.FontSize = 8
        Printer.Print Tab(95); "Hal :"; Format(mhal, "###")
        mgrs = String$(100, "-")
        Printer.Print Tab(5); mgrs
        Printer.Print Tab(5); "No.";
        Printer.Print Tab(10); "WAKTU";
        Printer.Print Tab(25); "NO. MEJA";
        Printer.Print Tab(35); "PRODUK";
        Printer.Print Tab(65); "HARGA";
        Printer.Print Tab(75); "QTY";
        Printer.Print Tab(83); "DSKN";
        Printer.Print Tab(88); "JUMLAH";
        Printer.Print Tab(97); "PETUGAS"
        Printer.Print Tab(5); mgrs
        mbaris = 0
        Do While Not .Recordset.EOF And mbaris <= 60
            mjumlah = 0
            subsumqty = 0
            mno = mno + 1
            DBproduk_form.DataProduk.Recordset.Index = "idxproduk"
            DBproduk_form.DataProduk.Recordset.Seek "=", .Recordset!Kode_produk
            Printer.Print Tab(5); rkanan(mno, "###");
            Printer.Print Tab(10); .Recordset!waktu;
            Printer.Print Tab(25); .Recordset!no_meja;
            'printer.Print Tab(35); DBproduk_form.DataProduk.Recordset!jenis_produk
            Printer.Print Tab(35); DBproduk_form.DataProduk.Recordset!Kode_produk; "  "; DBproduk_form.DataProduk.Recordset!Nama_produk;
            Printer.Print Tab(60); rkanan(DBproduk_form.DataProduk.Recordset!Harga_produk, "###,###,###");
            Printer.Print Tab(75); rkanan(.Recordset!qty, "#,###");
            Printer.Print Tab(80); rkanan(.Recordset!discount, "##.##") & "%";
            mjumlah = (.Recordset!qty * DBproduk_form.DataProduk.Recordset!Harga_produk) + (.Recordset!qty * DBproduk_form.DataProduk.Recordset!Harga_produk * frmseting.Data1.Recordset!tax / 100) - (.Recordset!qty * DBproduk_form.DataProduk.Recordset!Harga_produk * .Recordset!discount / 100)
            Printer.Print Tab(86); rkanan(mjumlah, "###,###,###");
            Printer.Print Tab(98); .Recordset!user
            subsumqty = subsumqty + .Recordset!qty
            totqty = totqty + .Recordset!qty
            mbaris = mbaris + 1
            total = total + mjumlah
            .Recordset.Edit
            .Recordset!Status = False
            .Recordset.Update
            .Recordset.MoveNext
        Loop
        Printer.Print Tab(5); mgrs
        'printer.Print Tab(60); "SUB TOTAL";
        'printer.Print Tab(80); rkanan(totqty, "#,###");
        'printer.Print Tab(85); rkanan(total, "###,###,###")
        grandqty = grandqty + totqty
        grandtot = grandtot + total
        totqty = 0
        total = 0
    Loop
    Printer.Print Tab(5); mgrs
    Printer.Print Tab(60); "TOTAL FOOD";
    Printer.Print Tab(75); rkanan(grandqty, "#,###");
    Printer.Print Tab(86); rkanan(grandtot, "###,###,###")
    Printer.Print Tab(5); mgrs
    Printer.Print
    Printer.Print
    End With
    Printer.Print Tab(5); mgrs
    'frmPool.dt2.Refresh
    With frmPool.dt2
    .RecordSource = "select * from pool where status=true " 'cdate(tanggal) = '" & Date & "'"
    .Refresh
    If .Recordset.RecordCount = 0 Then
        x = MsgBox("Tidak ada transaksi Pool !!!", vbOKOnly, "DATA KOSONG")
'        Exit Sub
    Else
    .Recordset.MoveFirst
    totpool = 0
    mno = 0
    Do While Not .Recordset.EOF
        mno = mno + 1
        Printer.Print Tab(5); mno;
        Printer.Print Tab(10); .Recordset!waktu_mulai;
        Printer.Print Tab(25); .Recordset!no_pool;
        Printer.Print Tab(35); .Recordset!nama_costumer;
        Printer.Print Tab(65); rkanan(.Recordset!harga_jam, "###,###");
        Printer.Print Tab(80); rkanan(.Recordset!discount, "##.##") & "%";
        'printer.Print Tab(80); (Minute(!waktu_selesai) - Minute(!waktu_mulai)) + ((Hour(!waktu_selesai) - Hour(waktu_mulai)) * 60);
        Printer.Print Tab(85); rkanan(.Recordset!jumlah_bayar, "###,###,###");
        Printer.Print Tab(97); .Recordset!user
        totpool = totpool + .Recordset!jumlah_bayar
        grandtot = grandtot + .Recordset!jumlah_bayar
        .Recordset.Edit
        .Recordset!Status = False
        .Recordset.Update
        .Recordset.MoveNext
    Loop
    End If
    Printer.Print Tab(5); mgrs
    Printer.Print Tab(60); "TOTAL POOL";
    Printer.Print Tab(85); rkanan(totpool, "###,###,###")
    Printer.Print Tab(5); mgrs
    frmseting.Data1.Refresh
    Printer.Print
    Printer.Print Tab(60); "TOTAL FOOD+POOL";
'    printer.Print
    Printer.Print Tab(85); rkanan(grandtot, "###,###,###")
'    printer.Print Tab(60); "TAX & SERVICE("; rkanan(frmseting.Data1.Recordset!tax, "##.##"); "%)";
    End With
    
'        printer.Print Tab(85); rkanan(grandtot * frmseting.Data1.Recordset!tax / 100, "###,###,###")
        'printer.Print Tab(60); "DISCOUNT("; rkanan(frmseting.Data1.Recordset!discount, "##.##"); "%)";
'        printer.Print Tab(85); rkanan(grandtot * frmseting.Data1.Recordset!discount / 100, "###,###,###")
'        printer.Print
'        printer.Print Tab(60); "GRAND TOTAL";
'        printer.Print Tab(85); rkanan(grandtot + grandtot * frmseting.Data1.Recordset!tax / 100 - grandtot * frmseting.Data1.Recordset!discount / 100, "###,###,###")
        grandqty = 0
        grandtot = 0
    Printer.EndDoc
End Sub

Private Sub tampillaporanharian()
    'x = MsgBox("Sedang Mencetak Laporan", vbInformation, "Cetak Laporan")
    Dim mno, mhal, mbaris As Integer
    Dim mjumlah, total, totqty, grandqty, grandtot, totpool As Double
    Dim mgrs As String
    tampil.Font = "courier new"
    tampil.Cls
    tampil.Show
    tampil.CurrentX = 0
    tampil.CurrentY = 0
    tampil.Caption = "LAPORAN PENJUALAN"
    With DbTransFrm.Data1
    .RecordSource = "select * from transaksi where status=true"
    .Refresh
    If .Recordset.RecordCount = 0 Then
        x = MsgBox("Maaf data masih kosong, silahkan diisi dahulu !!!", vbOKOnly, "DATA KOSONG")
        Exit Sub
    End If
    .Recordset.MoveFirst
    mno = 0
    mhal = 0
    Do While Not .Recordset.EOF
        mhal = mhal + 1
        tampil.Print
        tampil.Print
        tampil.FontBold = True
        tampil.FontSize = 10
        tampil.Print Tab(5); "LAPORAN PENJUALAN HARIAN"
        tampil.Print Tab(5); "Tanggal Cetak : "; Format(Date, "dd/mm/yyyy")
        tampil.FontBold = False
        tampil.FontSize = 8
        tampil.Print Tab(95); "Hal :"; Format(mhal, "###")
        mgrs = String$(100, "-")
        tampil.Print Tab(5); mgrs
        tampil.Print Tab(5); "No.";
        tampil.Print Tab(10); "WAKTU";
        tampil.Print Tab(25); "NO. MEJA";
        tampil.Print Tab(35); "PRODUK";
        tampil.Print Tab(65); "HARGA";
        tampil.Print Tab(75); "QTY";
        tampil.Print Tab(83); "DSKN";
        tampil.Print Tab(88); "JUMLAH";
        tampil.Print Tab(97); "PETUGAS"
        tampil.Print Tab(5); mgrs
        mbaris = 0
        Do While Not .Recordset.EOF And mbaris <= 60
            mjumlah = 0
            subsumqty = 0
            mno = mno + 1
            DBproduk_form.DataProduk.Recordset.Index = "idxproduk"
            DBproduk_form.DataProduk.Recordset.Seek "=", .Recordset!Kode_produk
            tampil.Print Tab(5); rkanan(mno, "###");
            tampil.Print Tab(10); .Recordset!waktu;
            tampil.Print Tab(25); .Recordset!no_meja;
            'tampil.Print Tab(35); DBproduk_form.DataProduk.Recordset!jenis_produk
            tampil.Print Tab(35); DBproduk_form.DataProduk.Recordset!Kode_produk; "  "; DBproduk_form.DataProduk.Recordset!Nama_produk;
            tampil.Print Tab(60); rkanan(DBproduk_form.DataProduk.Recordset!Harga_produk, "###,###,###");
            tampil.Print Tab(75); rkanan(.Recordset!qty, "#,###");
            tampil.Print Tab(80); rkanan(.Recordset!discount, "##.##") & "%";
            mjumlah = (.Recordset!qty * DBproduk_form.DataProduk.Recordset!Harga_produk) + (.Recordset!qty * DBproduk_form.DataProduk.Recordset!Harga_produk * frmseting.Data1.Recordset!tax / 100) - (.Recordset!qty * DBproduk_form.DataProduk.Recordset!Harga_produk * .Recordset!discount / 100)
            tampil.Print Tab(86); rkanan(mjumlah, "###,###,###");
            tampil.Print Tab(98); .Recordset!user
            subsumqty = subsumqty + .Recordset!qty
            totqty = totqty + .Recordset!qty
            mbaris = mbaris + 1
            total = total + mjumlah
            .Recordset.Edit
            .Recordset!Status = False
            .Recordset.Update
            .Recordset.MoveNext
        Loop
        tampil.Print Tab(5); mgrs
        'tampil.Print Tab(60); "SUB TOTAL";
        'tampil.Print Tab(80); rkanan(totqty, "#,###");
        'tampil.Print Tab(85); rkanan(total, "###,###,###")
        grandqty = grandqty + totqty
        grandtot = grandtot + total
        totqty = 0
        total = 0
    Loop
    tampil.Print Tab(5); mgrs
    tampil.Print Tab(60); "TOTAL FOOD";
    tampil.Print Tab(75); rkanan(grandqty, "#,###");
    tampil.Print Tab(86); rkanan(grandtot, "###,###,###")
    tampil.Print Tab(5); mgrs
    tampil.Print
    tampil.Print
    End With
    tampil.Print Tab(5); mgrs
    'frmPool.dt2.Refresh
    With frmPool.dt2
    .RecordSource = "select * from pool where status=true " 'cdate(tanggal) = '" & Date & "'"
    .Refresh
    If .Recordset.RecordCount = 0 Then
        x = MsgBox("Tidak ada transaksi Pool !!!", vbOKOnly, "DATA KOSONG")
'        Exit Sub
    Else
    .Recordset.MoveFirst
    totpool = 0
    mno = 0
    Do While Not .Recordset.EOF
        mno = mno + 1
        tampil.Print Tab(5); mno;
        tampil.Print Tab(10); .Recordset!waktu_mulai;
        tampil.Print Tab(25); .Recordset!no_pool;
        tampil.Print Tab(35); .Recordset!nama_costumer;
        tampil.Print Tab(65); rkanan(.Recordset!harga_jam, "###,###");
        tampil.Print Tab(83); rkanan(.Recordset!discount, "##.##");
        'tampil.Print Tab(80); (Minute(!waktu_selesai) - Minute(!waktu_mulai)) + ((Hour(!waktu_selesai) - Hour(waktu_mulai)) * 60);
        tampil.Print Tab(85); rkanan(.Recordset!jumlah_bayar, "###,###,###");
        tampil.Print Tab(97); .Recordset!user
        totpool = totpool + .Recordset!jumlah_bayar
        grandtot = grandtot + .Recordset!jumlah_bayar
        .Recordset.Edit
        .Recordset!Status = False
        .Recordset.Update
        .Recordset.MoveNext
    Loop
    End If
    tampil.Print Tab(5); mgrs
    tampil.Print Tab(60); "TOTAL POOL";
    tampil.Print Tab(85); rkanan(totpool, "###,###,###")
    tampil.Print Tab(5); mgrs
    frmseting.Data1.Refresh
    tampil.Print
    tampil.Print Tab(60); "TOTAL FOOD+POOL";
'    tampil.Print
    tampil.Print Tab(85); rkanan(grandtot, "###,###,###")
'    tampil.Print Tab(60); "TAX & SERVICE("; rkanan(frmseting.Data1.Recordset!tax, "##.##"); "%)";
    End With
    
'        tampil.Print Tab(85); rkanan(grandtot * frmseting.Data1.Recordset!tax / 100, "###,###,###")
        'tampil.Print Tab(60); "DISCOUNT("; rkanan(frmseting.Data1.Recordset!discount, "##.##"); "%)";
'        tampil.Print Tab(85); rkanan(grandtot * frmseting.Data1.Recordset!discount / 100, "###,###,###")
'        tampil.Print
'        tampil.Print Tab(60); "GRAND TOTAL";
'        tampil.Print Tab(85); rkanan(grandtot + grandtot * frmseting.Data1.Recordset!tax / 100 - grandtot * frmseting.Data1.Recordset!discount / 100, "###,###,###")
        grandqty = 0
        grandtot = 0
'    tampil.EndDoc
End Sub

Private Sub lap_Click()
'    frmlaporan.Show
    Me.Enabled = False
    cetakLap_frm.Show
End Sub

Private Sub Pool_Click()
frmLogin2.Show
Me.Hide
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Select Case SSTab1.Tab
Case 0
    frmLogin1.Show
    menu_utama.Visible = False
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Timer4.Enabled = False
    Timer5.Enabled = False
    Timer6.Enabled = False
    Timer7.Enabled = False
    Timer8.Enabled = False
    Timer9.Enabled = False
    Timer10.Enabled = False
    Timer11.Enabled = False
    Timer12.Enabled = False
    SSTab1.ForeColor = QBColor(1)
Case 1
    Label1.Left = 1560
    Timer1.Enabled = True
    Timer2.Enabled = False
    Timer3.Enabled = False
    Timer4.Enabled = False
    Timer5.Enabled = False
    Timer6.Enabled = False
    Timer7.Enabled = False
    Timer8.Enabled = False
    Timer9.Enabled = False
    Timer10.Enabled = False
    Timer11.Enabled = False
    Timer12.Enabled = False
    SSTab1.ForeColor = QBColor(12)
Case 2
    Label3.Left = 1560
    Timer3.Enabled = True
    Timer4.Enabled = False
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer5.Enabled = False
    Timer6.Enabled = False
    Timer7.Enabled = False
    Timer8.Enabled = False
    Timer9.Enabled = False
    Timer10.Enabled = False
    Timer11.Enabled = True
    Timer12.Enabled = False
    SSTab1.ForeColor = QBColor(2)
Case 3
    Label6.Left = 1560
    Timer3.Enabled = False
    Timer4.Enabled = False
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer5.Enabled = False
    Timer6.Enabled = False
    Timer5.Enabled = True
    Timer7.Enabled = True
    Timer8.Enabled = False
    Timer9.Enabled = False
    Timer10.Enabled = False
    Timer6.Enabled = False
    Timer11.Enabled = False
    Timer12.Enabled = False
    Shape1.Left = 2640
    Label7(0).Visible = False
    Label7(1).Visible = False
    Line4.X1 = 120
    Line4.X2 = 1320
    SSTab1.ForeColor = QBColor(5)
End Select
End Sub

Private Sub TambahDB_Click()
    DBproduk_form.cmdproses.Caption = "&Tambah"
    DBproduk_form.cmdbatal.Caption = "&Keluar"
    DBproduk_form.Caption = "TAMBAH DATA"
    DBproduk_form.Visible = True
    menu_utama.Visible = False
End Sub

Private Sub Timer1_Timer()
    Label1.Left = Label1.Left + 20
    If Label1.Left > 3239 Then
        Timer1.Enabled = False
        Timer2.Enabled = True
    End If
End Sub

Private Sub Timer10_Timer()
    Line4.X1 = Line4.X1 - 20
    Line4.X2 = Line4.X2 - 20
    If Line4.X1 < 100 Then
    Timer7.Enabled = True
    Timer10.Enabled = False
    End If
End Sub

Private Sub Timer11_Timer()
Label3.Visible = True
Timer11.Enabled = False
Timer12.Enabled = True
End Sub

Private Sub Timer12_Timer()
Label3.Visible = False
Timer11.Enabled = True
Timer12.Enabled = False
End Sub

Private Sub Timer13_Timer(Index As Integer)
    Label10(Index) = Label10(Index) + 1
End Sub

Private Sub Timer2_Timer()
    Label1.Left = Label1.Left - 20
    If Label1.Left < 121 Then
        Timer1.Enabled = True
        Timer2.Enabled = False
    End If
End Sub

Private Sub Timer3_Timer()
    Label3.Left = Label3.Left + 20
    If Label3.Left > 4440 Then
        Timer3.Enabled = False
        Timer4.Enabled = True
    End If
End Sub


Private Sub Timer4_Timer()
    Label3.Left = Label3.Left - 20
    If Label3.Left < 121 Then
        Timer3.Enabled = True
        Timer4.Enabled = False
    End If
End Sub

Private Sub Timer5_Timer()
    Label6.Left = Label6.Left + 20
    If Label6.Left > 4320 Then
        Timer5.Enabled = False
        Timer6.Enabled = True
    End If
End Sub

Private Sub Timer6_Timer()
    Label6.Left = Label6.Left - 20
    If Label6.Left < 121 Then
        Timer5.Enabled = True
        Timer6.Enabled = False
    End If
End Sub

Private Sub updateDB_Click()
    DBproduk_form.cmdproses.Caption = "&Edit"
    DBproduk_form.cmdbatal.Caption = "&Keluar"
    DBproduk_form.Caption = "TAMBAH DATA"
    DBproduk_form.Visible = True
    menu_utama.Visible = False
End Sub

Private Sub Timer7_Timer()
    Line4.X1 = Line4.X1 + 40
    Line4.X2 = Line4.X2 + 40
    If Line4.X2 > 2620 Then
        Label7(0).Visible = True
        Timer8.Enabled = True
        Timer7.Enabled = False
    End If
End Sub

Private Sub Timer8_Timer()
    Shape1.Left = Shape1.Left + 40
    If Shape1.Left > 4780 Then
        Label7(0).Visible = False
        Label7(1).Visible = True
        Timer8.Enabled = False
        Timer9.Enabled = True
    End If
End Sub

Private Sub Timer9_Timer()
    Shape1.Left = Shape1.Left - 20
    If Shape1.Left < 2620 Then
        Label7(1).Visible = False
        Timer9.Enabled = False
        Timer10.Enabled = True
    End If
End Sub

Private Sub transbar_Click()
With DbTransFrm
    .Data1.RecordSource = "select * from transaksi"
    .Data1.Refresh
If Not .Data1.Recordset.BOF Then
.Data1.Recordset.MoveFirst
.Data1.Refresh
.Data1.RecordSource = "select * from transaksi where no_meja like " & "'Bar*'" & " order by No_meja asc"
.Data1.Refresh
End If
frmLogin.Show
End With
End Sub

Private Sub Transpool_Click()
With DbTransFrm
    .Data1.RecordSource = "select * from transaksi"
    .Data1.Refresh
If Not .Data1.Recordset.BOF Then
.Data1.Recordset.MoveFirst
.Data1.Refresh
.Data1.RecordSource = "select * from transaksi where no_meja like " & "'Pool*'" & " order by No_meja asc"
.Data1.Refresh
End If
frmLogin.Show
End With
End Sub

Private Sub transrest_Click()
With DbTransFrm
    .Data1.RecordSource = "select * from transaksi"
    .Data1.Refresh
If Not .Data1.Recordset.BOF Then
.Data1.Recordset.MoveFirst
.Data1.Refresh
.Data1.RecordSource = "select * from transaksi where no_meja like " & "'Meja*'" & " order by no_meja asc"
.Data1.Refresh
End If
frmLogin.Show
menu_utama.Visible = False
End With
End Sub
