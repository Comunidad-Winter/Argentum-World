VERSION 5.00
Begin VB.Form frmZonaInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Información de la zona"
   ClientHeight    =   9360
   ClientLeft      =   11505
   ClientTop       =   3030
   ClientWidth     =   4680
   Icon            =   "frmMapInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   624
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   Begin VB.CheckBox chkInterdimensional 
      Caption         =   "Interdimensional"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2280
      TabIndex        =   105
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CheckBox chkNewbie 
      Caption         =   "Newbie"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   104
      Top             =   5640
      Width           =   855
   End
   Begin VB.TextBox txtMap 
      Height          =   285
      Left            =   600
      TabIndex        =   102
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox txtSalidaY 
      Height          =   285
      Left            =   3120
      TabIndex        =   99
      Top             =   8400
      Width           =   615
   End
   Begin VB.TextBox txtSalidaX 
      Height          =   285
      Left            =   2160
      TabIndex        =   98
      Top             =   8400
      Width           =   615
   End
   Begin VB.TextBox txtSalidaMap 
      Height          =   285
      Left            =   1200
      TabIndex        =   96
      Top             =   8400
      Width           =   615
   End
   Begin VB.CommandButton btnCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2760
      TabIndex        =   95
      Top             =   8880
      Width           =   1455
   End
   Begin VB.CheckBox chkSoloClanes 
      Caption         =   "Solo clanes"
      Height          =   210
      Left            =   240
      TabIndex        =   94
      Top             =   6360
      Width           =   1335
   End
   Begin VB.ComboBox cmbFaccion 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmMapInfo.frx":628A
      Left            =   1200
      List            =   "frmMapInfo.frx":629D
      Style           =   2  'Dropdown List
      TabIndex        =   92
      Top             =   7080
      Width           =   2655
   End
   Begin VB.CheckBox chkNieve 
      Caption         =   "Nieve"
      Height          =   210
      Left            =   240
      TabIndex        =   91
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CheckBox chkNiebla 
      Caption         =   "Niebla"
      Height          =   210
      Left            =   240
      TabIndex        =   90
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CheckBox chkLluvia 
      Caption         =   "Lluvia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   89
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CheckBox chkSoloFaccion 
      Caption         =   "Solo facción"
      Height          =   210
      Left            =   240
      TabIndex        =   88
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton cmdIraZona 
      Caption         =   "Ir a Zona"
      Height          =   375
      Left            =   120
      TabIndex        =   87
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtIrZona 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1200
      TabIndex        =   86
      Text            =   "1"
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtLvlMin 
      Height          =   285
      Left            =   1200
      TabIndex        =   85
      Top             =   8040
      Width           =   615
   End
   Begin VB.TextBox txtLvlMax 
      Height          =   285
      Left            =   1200
      TabIndex        =   83
      Top             =   7680
      Width           =   615
   End
   Begin VB.CheckBox chkSinMascotas 
      Caption         =   "Sin Mascotas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2280
      TabIndex        =   81
      Top             =   6120
      Width           =   2295
   End
   Begin VB.CommandButton btnwav 
      Height          =   255
      Index           =   4
      Left            =   6840
      Picture         =   "frmMapInfo.frx":62CE
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   7320
      Width           =   255
   End
   Begin VB.TextBox tSonidos 
      Height          =   285
      Index           =   4
      Left            =   5640
      TabIndex        =   78
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton Command39 
      Height          =   255
      Left            =   7680
      Picture         =   "frmMapInfo.frx":682A
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   7320
      Width           =   375
   End
   Begin VB.CommandButton Command38 
      Height          =   255
      Left            =   8640
      Picture         =   "frmMapInfo.frx":6D86
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   7320
      Width           =   375
   End
   Begin VB.CommandButton Command37 
      Height          =   255
      Left            =   8160
      Picture         =   "frmMapInfo.frx":72E2
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   7320
      Width           =   375
   End
   Begin VB.CommandButton btnwav 
      Height          =   255
      Index           =   18
      Left            =   7200
      Picture         =   "frmMapInfo.frx":783E
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   7320
      Width           =   375
   End
   Begin VB.CommandButton btnwav 
      Height          =   255
      Index           =   2
      Left            =   6840
      Picture         =   "frmMapInfo.frx":7D9A
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   6600
      Width           =   255
   End
   Begin VB.TextBox tSonidos 
      Height          =   285
      Index           =   2
      Left            =   5640
      TabIndex        =   71
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command36 
      Height          =   255
      Left            =   7680
      Picture         =   "frmMapInfo.frx":82F6
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton Command35 
      Height          =   255
      Left            =   8640
      Picture         =   "frmMapInfo.frx":8852
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton Command34 
      Height          =   255
      Left            =   8160
      Picture         =   "frmMapInfo.frx":8DAE
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton btnwav 
      Height          =   255
      Index           =   16
      Left            =   7200
      Picture         =   "frmMapInfo.frx":930A
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton btnwav 
      Height          =   255
      Index           =   3
      Left            =   6840
      Picture         =   "frmMapInfo.frx":9866
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   6960
      Width           =   255
   End
   Begin VB.TextBox tSonidos 
      Height          =   285
      Index           =   3
      Left            =   5640
      TabIndex        =   64
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton Command25 
      Height          =   255
      Left            =   7680
      Picture         =   "frmMapInfo.frx":9DC2
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton Command24 
      Height          =   255
      Left            =   8640
      Picture         =   "frmMapInfo.frx":A31E
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton Command23 
      Height          =   255
      Left            =   8160
      Picture         =   "frmMapInfo.frx":A87A
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton btnwav 
      Height          =   255
      Index           =   14
      Left            =   7200
      Picture         =   "frmMapInfo.frx":ADD6
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton btnwav 
      Height          =   255
      Index           =   0
      Left            =   6840
      Picture         =   "frmMapInfo.frx":B332
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   5880
      Width           =   255
   End
   Begin VB.TextBox tSonidos 
      Height          =   285
      Index           =   0
      Left            =   5640
      TabIndex        =   57
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton Command22 
      Height          =   255
      Left            =   7680
      Picture         =   "frmMapInfo.frx":B88E
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   5880
      Width           =   375
   End
   Begin VB.CommandButton Command21 
      Height          =   255
      Left            =   8640
      Picture         =   "frmMapInfo.frx":BDEA
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   5880
      Width           =   375
   End
   Begin VB.CommandButton Command20 
      Height          =   255
      Left            =   8160
      Picture         =   "frmMapInfo.frx":C346
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   5880
      Width           =   375
   End
   Begin VB.CommandButton btnwav 
      Height          =   255
      Index           =   12
      Left            =   7200
      Picture         =   "frmMapInfo.frx":C8A2
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   5880
      Width           =   375
   End
   Begin VB.CommandButton btnwav 
      Height          =   255
      Index           =   1
      Left            =   6840
      Picture         =   "frmMapInfo.frx":CDFE
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   6240
      Width           =   255
   End
   Begin VB.TextBox tSonidos 
      Height          =   285
      Index           =   1
      Left            =   5640
      TabIndex        =   50
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton Command19 
      Height          =   255
      Left            =   7680
      Picture         =   "frmMapInfo.frx":D35A
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   6240
      Width           =   375
   End
   Begin VB.CommandButton Command18 
      Height          =   255
      Left            =   8640
      Picture         =   "frmMapInfo.frx":D8B6
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   6240
      Width           =   375
   End
   Begin VB.CommandButton Command17 
      Height          =   255
      Left            =   8160
      Picture         =   "frmMapInfo.frx":DE12
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   6240
      Width           =   375
   End
   Begin VB.CommandButton btnwav 
      Height          =   255
      Index           =   10
      Left            =   7200
      Picture         =   "frmMapInfo.frx":E36E
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   6240
      Width           =   375
   End
   Begin VB.CommandButton Command9 
      Height          =   255
      Left            =   3000
      Picture         =   "frmMapInfo.frx":E8CA
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   3840
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      Height          =   255
      Left            =   3960
      Picture         =   "frmMapInfo.frx":EE26
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   3840
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Height          =   255
      Left            =   3480
      Picture         =   "frmMapInfo.frx":F382
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   3840
      Width           =   375
   End
   Begin VB.CommandButton btnMidi 
      Height          =   255
      Index           =   7
      Left            =   2520
      Picture         =   "frmMapInfo.frx":F8DE
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   3840
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Height          =   255
      Left            =   3000
      Picture         =   "frmMapInfo.frx":FE3A
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Height          =   255
      Left            =   3960
      Picture         =   "frmMapInfo.frx":10396
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Height          =   255
      Left            =   3480
      Picture         =   "frmMapInfo.frx":108F2
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton btnMidi 
      Height          =   255
      Index           =   6
      Left            =   2520
      Picture         =   "frmMapInfo.frx":10E4E
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton btnMidi 
      Height          =   255
      Index           =   5
      Left            =   2520
      Picture         =   "frmMapInfo.frx":113AA
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   3120
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Height          =   255
      Left            =   3480
      Picture         =   "frmMapInfo.frx":11906
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   3120
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Height          =   255
      Left            =   3960
      Picture         =   "frmMapInfo.frx":11E62
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   3120
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Left            =   3000
      Picture         =   "frmMapInfo.frx":123BE
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   3120
      Width           =   375
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Guardar Cambios"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   8880
      Width           =   1455
   End
   Begin VB.CheckBox chkSegura 
      BackColor       =   &H80000004&
      Caption         =   "Segura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   210
      Left            =   240
      TabIndex        =   20
      Top             =   5880
      Width           =   1095
   End
   Begin VB.TextBox txtMapNombre 
      Height          =   285
      Left            =   1680
      TabIndex        =   19
      Text            =   "Nombre"
      Top             =   600
      Width           =   2655
   End
   Begin VB.TextBox tMusica 
      Height          =   285
      Index           =   0
      Left            =   960
      TabIndex        =   18
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox tMusica 
      Height          =   285
      Index           =   1
      Left            =   960
      TabIndex        =   17
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox tMusica 
      Height          =   285
      Index           =   2
      Left            =   960
      TabIndex        =   16
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton btnMidi 
      Height          =   255
      Index           =   0
      Left            =   2160
      Picture         =   "frmMapInfo.frx":1291A
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton btnMidi 
      Height          =   255
      Index           =   1
      Left            =   2160
      Picture         =   "frmMapInfo.frx":12E76
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3480
      Width           =   255
   End
   Begin VB.CommandButton btnMidi 
      Height          =   255
      Index           =   2
      Left            =   2160
      Picture         =   "frmMapInfo.frx":133D2
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3840
      Width           =   255
   End
   Begin VB.TextBox tZY2 
      Height          =   285
      Left            =   2040
      TabIndex        =   12
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox tZX2 
      Height          =   285
      Left            =   2040
      TabIndex        =   11
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox tZY1 
      Height          =   285
      Left            =   600
      TabIndex        =   10
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox tZX1 
      Height          =   285
      Left            =   600
      TabIndex        =   9
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox TxtR 
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   2280
      TabIndex        =   8
      Text            =   "0"
      Top             =   4680
      Width           =   615
   End
   Begin VB.TextBox TxtG 
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   3120
      TabIndex        =   7
      Text            =   "0"
      Top             =   4680
      Width           =   615
   End
   Begin VB.TextBox TxtB 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   3960
      TabIndex        =   6
      Text            =   "0"
      Top             =   4680
      Width           =   615
   End
   Begin VB.CheckBox chkSinResu 
      Caption         =   "Sin Resucitar"
      Height          =   210
      Left            =   2280
      TabIndex        =   5
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CheckBox chkSinInvi 
      Caption         =   "Sin Invi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2280
      TabIndex        =   4
      Top             =   5880
      Width           =   1455
   End
   Begin VB.ComboBox txtMapZona 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmMapInfo.frx":1392E
      Left            =   1680
      List            =   "frmMapInfo.frx":1394A
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4320
      Width           =   2655
   End
   Begin VB.CheckBox chkMapBackup 
      Caption         =   "Backup"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   1
      Top             =   6120
      Width           =   855
   End
   Begin VB.CheckBox chkSinMagia 
      Caption         =   "Sin Magia"
      Height          =   210
      Left            =   2280
      TabIndex        =   0
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Map:"
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   103
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label17 
      Caption         =   "Y:"
      Height          =   255
      Left            =   2880
      TabIndex        =   101
      Top             =   8400
      Width           =   255
   End
   Begin VB.Label Label11 
      Caption         =   "X:"
      Height          =   255
      Left            =   1920
      TabIndex        =   100
      Top             =   8400
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "Salida map:"
      Height          =   255
      Left            =   240
      TabIndex        =   97
      Top             =   8400
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Facción:"
      Height          =   255
      Left            =   480
      TabIndex        =   93
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label lblMapNivelMinimo 
      Caption         =   "Lvl Minimo"
      Height          =   255
      Left            =   240
      TabIndex        =   84
      Top             =   8040
      Width           =   975
   End
   Begin VB.Label lblMapNivelMaximo 
      Caption         =   "Lvl Maximo"
      Height          =   210
      Left            =   240
      TabIndex        =   82
      Top             =   7680
      Width           =   975
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      X1              =   8
      X2              =   8
      Y1              =   96
      Y2              =   200
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      X1              =   184
      X2              =   8
      Y1              =   200
      Y2              =   200
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      X1              =   184
      X2              =   184
      Y1              =   96
      Y2              =   200
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      X1              =   8
      X2              =   184
      Y1              =   96
      Y2              =   96
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Sonido 5"
      ForeColor       =   &H80000001&
      Height          =   255
      Index           =   15
      Left            =   4920
      TabIndex        =   80
      Top             =   7320
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Sonido 3"
      ForeColor       =   &H80000001&
      Height          =   255
      Index           =   14
      Left            =   4920
      TabIndex        =   73
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Sonido 4"
      ForeColor       =   &H80000001&
      Height          =   255
      Index           =   13
      Left            =   4920
      TabIndex        =   66
      Top             =   6960
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Sonido 1"
      ForeColor       =   &H80000001&
      Height          =   255
      Index           =   12
      Left            =   4920
      TabIndex        =   59
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Sonido 2"
      ForeColor       =   &H80000001&
      Height          =   255
      Index           =   11
      Left            =   4920
      TabIndex        =   52
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   33
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "B"
      Height          =   255
      Left            =   3720
      TabIndex        =   32
      Top             =   4680
      Width           =   135
   End
   Begin VB.Label Label9 
      Caption         =   "G"
      Height          =   255
      Left            =   2880
      TabIndex        =   31
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label8 
      Caption         =   "R"
      Height          =   255
      Left            =   2040
      TabIndex        =   30
      Top             =   4680
      Width           =   135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Musica 1"
      ForeColor       =   &H80000001&
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   29
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Musica 2"
      ForeColor       =   &H80000001&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   28
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Musica 3"
      ForeColor       =   &H80000001&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   27
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Y2:"
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   5
      Left            =   1680
      TabIndex        =   26
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "X2:"
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   6
      Left            =   1680
      TabIndex        =   25
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Y1:"
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   24
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "X1:"
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   23
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label tNumZona 
      BackColor       =   &H80000004&
      Caption         =   "Zona Nº:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   2040
      TabIndex        =   22
      Top             =   120
      Width           =   2295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   8
      X2              =   264
      Y1              =   584
      Y2              =   584
   End
   Begin VB.Label Label3 
      Caption         =   "Terreno Tipo:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   4320
      Width           =   1095
   End
End
Attribute VB_Name = "frmZonaInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nZona As t_ZonaInfo
Dim OpenedZona As Integer
Private Sub btnCancelar_Click()
Unload Me
End Sub

Private Sub btnMidi_Click(Index As Integer)

    Dim ret As Integer

    Dim Num As Integer

    Num = Val(frmZonaInfo.tMusica(Index).Text)

    If IsPlaying Then
        ret = mciSendString("close mus", 0&, 0, 0)
        IsPlaying = False
    Else
        ret = mciSendString("open " & """" & App.Path & "\..\Resources\midi\" & Num & ".mid" & """" & " type sequencer alias mus", 0&, 0, 0)
        ret = mciSendString("play mus", 0&, 0, 0)
        IsPlaying = True
    End If

End Sub

Private Sub btnwav_Click(Index As Integer)

'Dim ret As Integer
'
'Dim Num As Integer
'
'Num = Val(frmMapInfo.tSonidos(Index).Text)
'
'If IsPlaying Then
'   ret = mciSendString("close mus", 0&, 0, 0)
'   IsPlaying = False
'Else
'   ret = mciSendString("open " & """" & App.Path & "\Wav\" & Num & ".wav" & """" & " type sequencer alias mus", 0&, 0, 0)
'   ret = mciSendString("play mus", 0&, 0, 0)
'   IsPlaying = True
'End If


   ' Close CANYON.MID file and sequencer device
End Sub

Public Sub cmdIraZona_Click()
    Dim i As Integer
    Dim e As Integer

    If txtIrZona.Text > NumZonas Then
        Exit Sub
    Else
        i = Val(txtIrZona.Text)
    End If

    If i < 0 Then i = 0

    OpenZona (i)

End Sub


Public Sub OpenZona(id As Integer)
    Dim i As Integer
    If id > 0 Then
        nZona = Zona(id)
    Else
        Dim newZona As t_ZonaInfo
        nZona = newZona
        nZona.Map = UserMap
        If RectanguloX > RectanguloX2 Then
            nZona.X2 = RectanguloX
            nZona.X = RectanguloX2
        Else
            nZona.X = RectanguloX
            nZona.X2 = RectanguloX2
        End If
        
        If RectanguloY > RectanguloY2 Then
            nZona.Y2 = RectanguloY
            nZona.Y = RectanguloY2
        Else
            nZona.Y = RectanguloY
            nZona.Y2 = RectanguloY2
        End If
        
        nZona.Lluvia = 1
        nZona.Terreno = "BOSQUE"
    End If
    
    OpenedZona = id
    
    With nZona
        txtMapNombre.Text = .Zona_name
        txtMap.Text = .Map
        tZX1.Text = .X
        tZY1.Text = .Y
        tZX2.Text = .X2
        tZY2.Text = .Y2
        txtMapZona.ListIndex = 0
        For i = 0 To txtMapZona.ListCount - 1
            If UCase$(txtMapZona.List(i)) = .Terreno Then
                txtMapZona.ListIndex = i
                Exit For
            End If
        Next i
        
        tMusica(0).Text = .Musica1
        tMusica(1).Text = .Musica2
        tMusica(2).Text = .Musica3
        
        
        chkLluvia.value = .Lluvia
        chkNiebla.value = .Niebla
        chkNieve.value = .Nieve
        chkSinMagia.value = .SinMagia
        chkSinInvi.value = .SinInvi
        chkSinMascotas.value = .SinMascotas
        chkSinResu.value = .SinResucitar
        chkSegura.value = .Segura
        chkNewbie.value = .Newbie
        chkMapBackup.value = .Backup
        chkSoloClanes.value = .SoloClanes
        chkSoloFaccion.value = .SoloFaccion
        chkInterdimensional.value = .Interdimensional

        cmbFaccion.ListIndex = .Faccion
        txtLvlMax.Text = .MaxLevel
        txtLvlMin.Text = .MinLevel
        
        txtSalidaMap.Text = .SalidaMap
        txtSalidaX.Text = .SalidaX
        txtSalidaY.Text = .SalidaY
        

    End With
    If id > 0 Then
        tNumZona.Caption = "Zona N°: " & id
    Else
        tNumZona.Caption = "Zona Nueva"
    End If
    txtIrZona.Text = id
End Sub

Private Sub cmdMusica_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmMusica.Show
End Sub

Private Sub Command1_Click()
    Dim ret As Integer

    Dim Num As Integer
    frmZonaInfo.tMusica(0) = frmZonaInfo.tMusica(0) - 1
    Num = frmZonaInfo.tMusica(0)
    If IsPlaying Then
        ret = mciSendString("close mus", 0&, 0, 0)
        IsPlaying = False
        ret = mciSendString("open " & """" & App.Path & "\..\Resources\Midi\" & Num & ".mid" & """" & " type sequencer alias mus", 0&, 0, 0)
        ret = mciSendString("play mus", 0&, 0, 0)
        IsPlaying = True
    End If
End Sub

Private Sub Command2_Click()
    Dim ret As Integer

    Dim Num As Integer
    frmZonaInfo.tMusica(0) = frmZonaInfo.tMusica(0) + 1
    Num = frmZonaInfo.tMusica(0)
    If IsPlaying Then
        ret = mciSendString("close mus", 0&, 0, 0)
        IsPlaying = False
        ret = mciSendString("open " & """" & App.Path & "\..\Resources\Midi\" & Num & ".mid" & """" & " type sequencer alias mus", 0&, 0, 0)
        ret = mciSendString("play mus", 0&, 0, 0)
        IsPlaying = True
    End If
End Sub


Private Sub Command26_Click()

If Trim$(txtMapNombre.Text) = "" Then
    Call MsgBox("El nombre de la zona no puede estar vacio.", vbExclamation)
    Exit Sub
End If

    With nZona
        .Zona_name = Trim$(txtMapNombre.Text)
        .Map = Val(txtMap.Text)
        .X = Val(tZX1.Text)
        .Y = Val(tZY1.Text)
        .X2 = Val(tZX2.Text)
        .Y2 = Val(tZY2.Text)
        .Terreno = UCase$(txtMapZona.Text)
        .Musica1 = Val(tMusica(0).Text)
        .Musica2 = Val(tMusica(1).Text)
        .Musica3 = Val(tMusica(2).Text)
        
        
        .Lluvia = chkLluvia.value
        .Niebla = chkNiebla.value
        .Nieve = chkNieve.value
        .SinMagia = chkSinMagia.value
        .SinInvi = chkSinInvi.value
        .SinMascotas = chkSinMascotas.value
        .SinResucitar = chkSinResu.value
        .Segura = chkSegura.value
        .Newbie = chkNewbie.value
        .Backup = chkMapBackup.value
        .SoloClanes = chkSoloClanes.value
        .SoloFaccion = chkSoloFaccion.value
        .Interdimensional = chkInterdimensional.value

        .Faccion = cmbFaccion.ListIndex
        .MaxLevel = Val(txtLvlMax.Text)
        .MinLevel = Val(txtLvlMin.Text)
        
        .SalidaMap = Val(txtSalidaMap.Text)
        .SalidaX = Val(txtSalidaX.Text)
        .SalidaY = Val(txtSalidaY.Text)
        

    End With

OpenedZona = SaveZona(OpenedZona, nZona)
Call FrmMain.DibujarZonas
Unload Me
End Sub


Private Sub Command3_Click()
Dim Num As Integer
Dim ret As Integer

Num = frmZonaInfo.tMusica(0)

If IsPlaying Then
   ret = mciSendString("close mus", 0&, 0, 0)
   IsPlaying = False
   ret = mciSendString("play mus", 0&, 0, 0)
   IsPlaying = True
End If
End Sub




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If UnloadMode = vbFormControlMenu Then
    Cancel = True
    Me.Hide
End If
End Sub
