VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000000&
   Caption         =   "TxtWav.Text = ""508-509"""
   ClientHeight    =   14610
   ClientLeft      =   2085
   ClientTop       =   750
   ClientWidth     =   28560
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   974
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1904
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar PB 
      Height          =   255
      Left            =   6240
      TabIndex        =   134
      Top             =   1200
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.TextBox txtnamemapa 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   23640
      TabIndex        =   129
      Top             =   2940
      Width           =   1935
   End
   Begin VB.ListBox lstMaps 
      Height          =   2595
      Left            =   24360
      Sorted          =   -1  'True
      TabIndex        =   127
      Top             =   240
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Magia"
      Height          =   375
      Index           =   0
      Left            =   4800
      TabIndex        =   126
      Top             =   1080
      Width           =   735
   End
   Begin VB.PictureBox renderer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   12975
      Left            =   4800
      ScaleHeight     =   865
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1145
      TabIndex        =   124
      Top             =   1560
      Width           =   17175
   End
   Begin VB.CommandButton BloqAll 
      Caption         =   "X"
      Height          =   255
      Left            =   2040
      TabIndex        =   123
      Top             =   5040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkBloqueo 
      BackColor       =   &H80000000&
      Caption         =   "N"
      Height          =   255
      Index           =   2
      Left            =   2040
      TabIndex        =   122
      Top             =   4680
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chkBloqueo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000000&
      Caption         =   "O"
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   121
      Top             =   5040
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chkBloqueo 
      BackColor       =   &H80000000&
      Caption         =   "S"
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   120
      Top             =   5400
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chkBloqueo 
      BackColor       =   &H80000000&
      Caption         =   "E"
      Height          =   255
      Index           =   3
      Left            =   2400
      TabIndex        =   119
      Top             =   5040
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000000&
      Caption         =   "Opción Grh"
      Height          =   1095
      Left            =   16200
      TabIndex        =   114
      Top             =   0
      Width           =   5775
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   18
         Left            =   240
         TabIndex        =   115
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "Grh Normal"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   19
         Left            =   240
         TabIndex        =   116
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "Dia / Noche"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   21
         Left            =   2160
         TabIndex        =   118
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Caption         =   "Remplazo Grh"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   22
         Left            =   2160
         TabIndex        =   125
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         Caption         =   "Limpiar Luz,Particula,Trigger's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
   End
   Begin VB.Frame FraOpciones 
      BackColor       =   &H80000000&
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   22080
      TabIndex        =   97
      Top             =   120
      Width           =   2175
      Begin VB.CommandButton cmdDM 
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   102
         Top             =   480
         Width           =   240
      End
      Begin VB.CommandButton cmdDM 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   600
         Picture         =   "frmMain.frx":628A
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   720
         Width           =   240
      End
      Begin VB.CommandButton cmdDM 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   840
         Picture         =   "frmMain.frx":6571
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   480
         Width           =   240
      End
      Begin VB.CommandButton cmdDM 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   360
         Picture         =   "frmMain.frx":6860
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   480
         Width           =   240
      End
      Begin VB.CommandButton cmdDM 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   600
         Picture         =   "frmMain.frx":6B50
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   240
         Width           =   240
      End
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   103
         Top             =   1080
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmMain.frx":6E42
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   104
         Top             =   1080
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmMain.frx":7A94
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   105
         Top             =   1440
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmMain.frx":86E6
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   3
         Left            =   720
         TabIndex        =   106
         Top             =   1440
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmMain.frx":9338
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   4
         Left            =   1200
         TabIndex        =   107
         Top             =   1080
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   "1"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   5
         Left            =   1680
         TabIndex        =   108
         Top             =   1080
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   "2"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   6
         Left            =   1200
         TabIndex        =   109
         Top             =   1440
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   "3"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   7
         Left            =   1680
         TabIndex        =   110
         Top             =   1440
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   "4"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   11
         Left            =   1200
         TabIndex        =   111
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "Ir Map"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   12
         Left            =   240
         TabIndex        =   112
         Top             =   1920
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "Ambientacion"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   17
         Left            =   1200
         TabIndex        =   113
         Top             =   600
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "Bloq"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   20
         Left            =   240
         TabIndex        =   117
         Top             =   2280
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "Sup x Bloques"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
   End
   Begin VB.CheckBox Check6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Techos transparentes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   16440
      TabIndex        =   95
      Top             =   1080
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   120
      ScaleHeight     =   273
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   289
      TabIndex        =   94
      Top             =   10440
      Width           =   4335
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1215
      Left            =   1680
      TabIndex        =   89
      Top             =   120
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   2143
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      TextRTF         =   $"frmMain.frx":9F8A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox MiniMap 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H8000000B&
      Height          =   11175
      Left            =   22080
      ScaleHeight     =   745
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   420
      TabIndex        =   88
      Top             =   3270
      Width           =   6300
      Begin VB.Shape ApuntadorRadar 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   6  'Inside Solid
         DrawMode        =   6  'Mask Pen Not
         FillColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   600
         Top             =   480
         Width           =   375
      End
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   6
      Left            =   11040
      TabIndex        =   29
      Top             =   30
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1826
      Caption         =   "Tri&gger's (F12)"
      CapAlign        =   2
      BackStyle       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":A001
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   5
      Left            =   10080
      TabIndex        =   28
      Top             =   30
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1826
      Caption         =   "&Objetos (F11)"
      CapAlign        =   2
      BackStyle       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":A5C7
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   3
      Left            =   9120
      TabIndex        =   27
      Top             =   30
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1826
      Caption         =   "&NPC's (F8)"
      CapAlign        =   2
      BackStyle       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":AAC8
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   2
      Left            =   8160
      TabIndex        =   26
      Top             =   30
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1826
      Caption         =   "&Bloqueos (F7)"
      CapAlign        =   2
      BackStyle       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":AE7C
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   1
      Left            =   7200
      TabIndex        =   25
      Top             =   30
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1826
      Caption         =   "&Translados (F6)"
      CapAlign        =   2
      BackStyle       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   5
      Image           =   "frmMain.frx":B1FD
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   0
      Left            =   6240
      TabIndex        =   24
      Top             =   30
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1826
      Caption         =   "&Superficie (F5)"
      CapAlign        =   2
      BackStyle       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   0
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   5
      Image           =   "frmMain.frx":E85D
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H cmdQuitarFunciones 
      Height          =   435
      Left            =   13920
      TabIndex        =   23
      ToolTipText     =   "Quitar Todas las Funciones Activadas"
      Top             =   600
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   767
      Caption         =   "&Quitar Funciones (F4)"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   12632319
   End
   Begin VB.Timer TimAutoGuardarMapa 
      Enabled         =   0   'False
      Interval        =   40000
      Left            =   1440
      Top             =   2400
   End
   Begin VB.TextBox StatTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3435
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "frmMain.frx":11DA3
      Top             =   6360
      Width           =   4395
   End
   Begin VB.PictureBox pPaneles 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   4395
      Left            =   120
      ScaleHeight     =   4365
      ScaleWidth      =   4365
      TabIndex        =   4
      Top             =   1800
      Width           =   4395
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   600
         Left            =   1320
         Top             =   120
      End
      Begin WorldEditor.lvButtons_H insertarParticula 
         Height          =   375
         Left            =   120
         TabIndex        =   83
         Top             =   3840
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Caption         =   "&Insertar Particula"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Timer Timer2 
         Interval        =   1000
         Left            =   240
         Top             =   3120
      End
      Begin VB.TextBox ColorLuz 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         TabIndex        =   85
         Text            =   "0"
         Top             =   2640
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.PictureBox LuzColor 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   600
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   84
         Top             =   960
         Visible         =   0   'False
         Width           =   375
      End
      Begin WorldEditor.lvButtons_H quitarparticula 
         Height          =   375
         Left            =   2280
         TabIndex        =   82
         Top             =   3840
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Caption         =   "&Quitar Particula"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   240
         Top             =   120
      End
      Begin VB.TextBox RangoLuz 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1200
         TabIndex        =   79
         Text            =   "0"
         Top             =   960
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox numerodeparticula 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1680
         TabIndex        =   78
         Text            =   "0"
         Top             =   3480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.PictureBox Picture5 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   6
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture6 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   7
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture7 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   8
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture8 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   9
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture9 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   10
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture11 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   30
         Top             =   0
         Width           =   0
      End
      Begin WorldEditor.lvButtons_H cAgregarFuncalAzar 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "Insetar NPC's al &Azar"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarFunc 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   40
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar NPC's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarFunc 
         Height          =   735
         Index           =   0
         Left            =   2400
         TabIndex        =   41
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar NPC's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cAgregarFuncalAzar 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   48
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "Insetar OBJ's al &Azar"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarFunc 
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   49
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar OBJ's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarFunc 
         Height          =   735
         Index           =   2
         Left            =   2400
         TabIndex        =   50
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar Objetos"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarFunc 
         Height          =   735
         Index           =   1
         Left            =   240
         TabIndex        =   63
         Top             =   3360
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar NPC's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarFunc 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   62
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar NPC's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cAgregarFuncalAzar 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   61
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "Insetar NPC's al &Azar"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.ComboBox cNumFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   1
         ItemData        =   "frmMain.frx":11DE3
         Left            =   3360
         List            =   "frmMain.frx":11DE5
         TabIndex        =   60
         Text            =   "500"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin WorldEditor.lvButtons_H insertarLuz 
         Height          =   375
         Left            =   240
         TabIndex        =   80
         Top             =   1800
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "&Insertar Luz"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H QuitarLuz 
         Height          =   375
         Left            =   240
         TabIndex        =   81
         Top             =   2160
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "&Quitar Luz"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.ComboBox cFiltro 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   0
         Left            =   600
         TabIndex        =   52
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ComboBox cCapas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         ItemData        =   "frmMain.frx":11DE7
         Left            =   1080
         List            =   "frmMain.frx":11DF7
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cCantFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   2
         ItemData        =   "frmMain.frx":11E07
         Left            =   840
         List            =   "frmMain.frx":11E09
         TabIndex        =   0
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cGrh 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Left            =   2880
         TabIndex        =   53
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cCantFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   0
         ItemData        =   "frmMain.frx":11E0B
         Left            =   840
         List            =   "frmMain.frx":11E0D
         TabIndex        =   38
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cCantFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   1
         ItemData        =   "frmMain.frx":11E0F
         Left            =   840
         List            =   "frmMain.frx":11E11
         TabIndex        =   57
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cFiltro 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   3
         Left            =   600
         TabIndex        =   45
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ComboBox cNumFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   2
         ItemData        =   "frmMain.frx":11E13
         Left            =   3360
         List            =   "frmMain.frx":11E15
         TabIndex        =   47
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2580
         Index           =   2
         ItemData        =   "frmMain.frx":11E17
         Left            =   4440
         List            =   "frmMain.frx":11E19
         Sorted          =   -1  'True
         TabIndex        =   59
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.ComboBox cFiltro 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   1
         Left            =   600
         TabIndex        =   36
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
      End
      Begin WorldEditor.lvButtons_H cSeleccionarSuperficie 
         Height          =   735
         Left            =   2400
         TabIndex        =   56
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar Superficie"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarEnEstaCapa 
         Height          =   375
         Left            =   120
         TabIndex        =   55
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar en esta Capa"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarEnTodasLasCapas 
         Height          =   375
         Left            =   120
         TabIndex        =   54
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "Quitar en &Capas 2 y 3"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cUnionManual 
         Height          =   375
         Left            =   240
         TabIndex        =   69
         Top             =   2160
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "&Union con Mapa Adyacente (manual)"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.ComboBox cFiltro 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   2
         Left            =   600
         TabIndex        =   58
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ComboBox cNumFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   0
         ItemData        =   "frmMain.frx":11E1B
         Left            =   3360
         List            =   "frmMain.frx":11E1D
         TabIndex        =   37
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin WorldEditor.lvButtons_H cInsertarTrans 
         Height          =   375
         Left            =   240
         TabIndex        =   67
         Top             =   1440
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "&Insertar Translado"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.TextBox tTMapa 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Left            =   1200
         TabIndex        =   64
         Text            =   "1"
         Top             =   240
         Visible         =   0   'False
         Width           =   2900
      End
      Begin VB.TextBox tTX 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Left            =   1200
         TabIndex        =   65
         Text            =   "1"
         Top             =   600
         Visible         =   0   'False
         Width           =   2900
      End
      Begin VB.TextBox tTY 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Left            =   1200
         TabIndex        =   66
         Text            =   "1"
         Top             =   960
         Visible         =   0   'False
         Width           =   2900
      End
      Begin WorldEditor.lvButtons_H cVerBloqueos 
         Height          =   495
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   873
         Caption         =   "&Mostrar Bloqueos"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarBloqueo 
         Height          =   975
         Left            =   120
         TabIndex        =   44
         Top             =   840
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1720
         Caption         =   "&Quitar Bloqueos"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarTransOBJ 
         Height          =   375
         Left            =   240
         TabIndex        =   68
         Top             =   1800
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "Colocar automaticamente &Objeto"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cUnionAuto 
         Height          =   375
         Left            =   240
         TabIndex        =   70
         Top             =   2520
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "Union con Mapas &Adyacentes (auto)"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarTrans 
         Height          =   375
         Left            =   240
         TabIndex        =   71
         Top             =   2880
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "&Quitar Translados"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2580
         Index           =   0
         ItemData        =   "frmMain.frx":11E1F
         Left            =   120
         List            =   "frmMain.frx":11E21
         Sorted          =   -1  'True
         TabIndex        =   51
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2580
         Index           =   3
         ItemData        =   "frmMain.frx":11E23
         Left            =   120
         List            =   "frmMain.frx":11E25
         Sorted          =   -1  'True
         TabIndex        =   46
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2580
         Index           =   1
         ItemData        =   "frmMain.frx":11E27
         Left            =   120
         List            =   "frmMain.frx":11E29
         Sorted          =   -1  'True
         TabIndex        =   35
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   3210
         Index           =   4
         ItemData        =   "frmMain.frx":11E2B
         Left            =   120
         List            =   "frmMain.frx":11E2D
         TabIndex        =   34
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin WorldEditor.lvButtons_H cInsertarBloqueo 
         Height          =   615
         Left            =   120
         TabIndex        =   43
         Top             =   2040
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1085
         Caption         =   "&Insertar Bloqueos"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarTrigger 
         Height          =   375
         Left            =   2400
         TabIndex        =   33
         Top             =   3480
         Visible         =   0   'False
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   661
         Caption         =   "&Insertar Trigger"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cVerTriggers 
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Mostrar Trigger's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarTrigger 
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar Trigger's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H TiggerEspecial 
         Height          =   375
         Left            =   2400
         TabIndex        =   91
         Top             =   3840
         Visible         =   0   'False
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   661
         Caption         =   "&Trigger Especial"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.ListBox ListaParticulas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3390
         Left            =   0
         Sorted          =   -1  'True
         TabIndex        =   87
         Top             =   0
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Si el rango es mayor a 100 la luz se convierte en redonda."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   96
         Top             =   3120
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label lYver 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Y vertical:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   240
         TabIndex        =   74
         Top             =   1005
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lXhor 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "X horizontal:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   240
         TabIndex        =   73
         Top             =   645
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label lMapN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Mapa:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   240
         TabIndex        =   72
         Top             =   285
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label lbFiltrar 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   2820
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lbCapas 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Capa Actual:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   120
         TabIndex        =   21
         Top             =   3195
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label lbGrh 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Sup Actual:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   2040
         TabIndex        =   20
         Top             =   3195
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lNumFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Numero de NPC:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   1
         Left            =   2160
         TabIndex        =   19
         Top             =   3195
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lCantFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   3195
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lbFiltrar 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   2820
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lNumFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Numero de OBJ:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   2
         Left            =   2160
         TabIndex        =   16
         Top             =   3195
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lCantFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   3195
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lbFiltrar 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   2820
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lCantFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   3195
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lNumFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Numero de NPC:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   0
         Left            =   2160
         TabIndex        =   12
         Top             =   3195
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lbFiltrar 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   2820
         Visible         =   0   'False
         Width           =   450
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      FillColor       =   &H80000000&
      ForeColor       =   &H00000000&
      Height          =   3660
      Left            =   120
      ScaleHeight     =   244
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   293
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6240
      Width           =   4395
      Begin VB.PictureBox PreviewGrh 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         FillColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   3300
         Left            =   0
         ScaleHeight     =   220
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   293
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   4395
         Begin VB.Shape Cual 
            BorderColor     =   &H0000FF00&
            BorderStyle     =   3  'Dot
            DrawMode        =   7  'Invert
            FillColor       =   &H0080FF80&
            Height          =   495
            Left            =   0
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
      End
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   2565
      Top             =   2025
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   675
      Index           =   4
      Left            =   10080
      TabIndex        =   75
      Top             =   360
      Visible         =   0   'False
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1191
      Caption         =   "none"
      CapAlign        =   2
      BackStyle       =   3
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":11E2F
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   7
      Left            =   12000
      TabIndex        =   76
      Top             =   30
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1826
      Caption         =   "Particulas"
      CapAlign        =   2
      BackStyle       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":121E3
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   8
      Left            =   12960
      TabIndex        =   77
      Top             =   30
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1826
      Caption         =   "Luces"
      CapAlign        =   2
      BackStyle       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":125A2
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   480
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin WorldEditor.lvButtons_H cmdInformacionDelMapa 
      Height          =   375
      Left            =   13920
      TabIndex        =   90
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Caption         =   "&Información del Mapa"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H lvButtons_H1 
      Height          =   915
      Left            =   120
      TabIndex        =   128
      ToolTipText     =   "Quitar Todas las Funciones Activadas"
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1614
      Caption         =   "Guardar Cambios (CTRL + G)"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   12632319
   End
   Begin WorldEditor.lvButtons_H LvBOpcion 
      Height          =   375
      Index           =   8
      Left            =   26610
      TabIndex        =   131
      Top             =   2880
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "Zonas"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H LvBOpcion 
      Height          =   375
      Index           =   9
      Left            =   27480
      TabIndex        =   132
      Top             =   2880
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "Spawns"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H LvBOpcion 
      Height          =   375
      Index           =   10
      Left            =   25680
      TabIndex        =   133
      Top             =   2880
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "Hostiles"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del mapa"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   22200
      TabIndex        =   130
      Top             =   2985
      Width           =   1455
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      Caption         =   "Label15"
      Height          =   255
      Left            =   4560
      TabIndex        =   93
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      Caption         =   "Mapa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   92
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label POSX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1680
      TabIndex        =   86
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Shape MainViewShp 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0E0FF&
      Height          =   10965
      Left            =   4680
      Top             =   1440
      Width           =   11325
   End
   Begin VB.Menu FileMnu 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuNuevoMapa 
         Caption         =   "&Nuevo Mapa"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuAbrirMapa 
         Caption         =   "&Abrir Mapa"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuAbrirMapaLong 
         Caption         =   "&Abrir Mapa Long"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuReAbrirMapa 
         Caption         =   "&Re-Abrir Mapa"
      End
      Begin VB.Menu mnuArchivoLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGuardarMapa 
         Caption         =   "&Guardar Mapa"
         Shortcut        =   ^G
      End
      Begin VB.Menu mmGuardarCliente 
         Caption         =   "&Guardar Cliente"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuEdicion 
      Caption         =   "&Edición"
      Begin VB.Menu mnuCortar 
         Caption         =   "C&ortar Selección"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopiar 
         Caption         =   "&Copiar Selección"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPegar 
         Caption         =   "&Pegar Selección"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuBloquearS 
         Caption         =   "&Bloquear Selección"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuRealizarOperacion 
         Caption         =   "&Realizar Operación en Selección"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuDeshacerPegado 
         Caption         =   "Deshacer P&egado de Selección"
         Shortcut        =   ^S
      End
      Begin VB.Menu mmCapasCopiar 
         Caption         =   "Capas que se copian"
         Begin VB.Menu mmCopiarCapa1 
            Caption         =   "Capa 1"
            Checked         =   -1  'True
            Shortcut        =   ^{F1}
         End
         Begin VB.Menu mmCopiarCapa2 
            Caption         =   "Capa 2"
            Checked         =   -1  'True
            Shortcut        =   ^{F2}
         End
         Begin VB.Menu mmCopiarCapa3 
            Caption         =   "Capa 3"
            Checked         =   -1  'True
            Shortcut        =   ^{F3}
         End
         Begin VB.Menu mmCopiarCapa4 
            Caption         =   "Capa 4"
            Checked         =   -1  'True
            Shortcut        =   ^{F4}
         End
      End
      Begin VB.Menu mnuLineEdicion0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeshacer 
         Caption         =   "&Deshacer"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuUtilizarDeshacer 
         Caption         =   "&Utilizar Deshacer"
         Checked         =   -1  'True
      End
      Begin VB.Menu mmMapSize 
         Caption         =   "&Tamaño del Mapa"
      End
      Begin VB.Menu mnuLineEdicion1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuitar 
         Caption         =   "&Quitar"
         Begin VB.Menu Todas_las_Particulas 
            Caption         =   "Todas las Particulas"
         End
         Begin VB.Menu Todas_las_luces 
            Caption         =   "Todas las luces"
         End
         Begin VB.Menu mnuQuitarTranslados 
            Caption         =   "Todos los &Translados"
         End
         Begin VB.Menu mnuQuitarBloqueos 
            Caption         =   "Todos los &Bloqueos"
         End
         Begin VB.Menu mnuQuitarNPCs 
            Caption         =   "Todos los &NPC's"
         End
         Begin VB.Menu mnuQuitarNPCsHostiles 
            Caption         =   "Todos los NPC's &Hostiles"
         End
         Begin VB.Menu mnuQuitarObjetos 
            Caption         =   "Todos los &Objetos"
         End
         Begin VB.Menu mnuQuitarTriggers 
            Caption         =   "Todos los Tri&gger's"
         End
         Begin VB.Menu mnuLineEdicion2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuQuitarTODO 
            Caption         =   "TODO"
         End
      End
      Begin VB.Menu mnuLineEdicion3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFunciones 
         Caption         =   "&Funciones"
         Begin VB.Menu mnuQuitarFunciones 
            Caption         =   "&Quitar Funciones"
            Shortcut        =   {F4}
         End
         Begin VB.Menu mnuAutoQuitarFunciones 
            Caption         =   "Auto-&Quitar Funciones"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuConfigAvanzada 
         Caption         =   "Configuracion A&vanzada de Superficie"
      End
      Begin VB.Menu mnuLineEdicion4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoCompletarSuperficies 
         Caption         =   "Auto-Completar &Superficies"
      End
      Begin VB.Menu mnuAutoCapturarSuperficie 
         Caption         =   "Auto-C&apturar información de la Superficie"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuAutoCapturarTranslados 
         Caption         =   "Auto-&Capturar información de los Translados"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuAutoGuardarMapas 
         Caption         =   "Configuración de Auto-&Guardar Mapas"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "&Ver"
      Begin VB.Menu mnuCapas 
         Caption         =   "...&Capas"
         Begin VB.Menu mnuVerCapa1 
            Caption         =   "Capa &1 (Piso)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuVerCapa2 
            Caption         =   "Capa &2 (costas, etc)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuVerCapa3 
            Caption         =   "Capa &3 (arboles, etc)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuVerCapa4 
            Caption         =   "Capa &4 (techos, etc)"
         End
      End
      Begin VB.Menu mnuVerTranslados 
         Caption         =   "...&Translados"
      End
      Begin VB.Menu mnuVerBloqueos 
         Caption         =   "...&Bloqueos"
      End
      Begin VB.Menu mnuVerNPCs 
         Caption         =   "...&NPC's"
      End
      Begin VB.Menu mnuVerObjetos 
         Caption         =   "...&Objetos"
      End
      Begin VB.Menu mnuVerTriggers 
         Caption         =   "...Tri&gger's"
      End
      Begin VB.Menu mnuVerMarco 
         Caption         =   "...Marco"
      End
      Begin VB.Menu mnuVerGrilla 
         Caption         =   "...Gri&lla"
      End
      Begin VB.Menu mnuVerLuces 
         Caption         =   "...Luces"
      End
      Begin VB.Menu mnuVerParticulas 
         Caption         =   "...Particulas"
      End
      Begin VB.Menu mnuZonas 
         Caption         =   "...Zonas"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuNpcSpawn 
         Caption         =   "...NpcSpawns"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuVerAutomatico 
         Caption         =   "Control &Automaticamente"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuPaneles 
      Caption         =   "&Paneles"
      Begin VB.Menu mnuSuperficie 
         Caption         =   "&Superficie"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuTranslados 
         Caption         =   "&Translados"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuBloquear 
         Caption         =   "&Bloquear"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuNpcS 
         Caption         =   "&NPC's"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuNPCsHostiles 
         Caption         =   "NPC's &Hostiles"
         Shortcut        =   {F9}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuObjetos 
         Caption         =   "&Objetos"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuTriggers 
         Caption         =   "Tri&gger's"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuQSuperficie 
         Caption         =   "Ocultar Superficie"
         Shortcut        =   +{F5}
      End
      Begin VB.Menu mnuQTranslados 
         Caption         =   "Ocultar Translados"
         Shortcut        =   +{F6}
      End
      Begin VB.Menu mnuQBloquear 
         Caption         =   "Ocultar Bloquear"
         Shortcut        =   +{F7}
      End
      Begin VB.Menu mnuQNPCs 
         Caption         =   "Ocultar NPC's"
         Shortcut        =   +{F8}
      End
      Begin VB.Menu mnuQNPCsHostiles 
         Caption         =   "Ocultar NPC's Hostiles"
         Shortcut        =   +{F9}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuQObjetos 
         Caption         =   "Ocultar Objetos"
         Shortcut        =   +{F11}
      End
      Begin VB.Menu mnuQTriggers 
         Caption         =   "Ocultar Trigger's"
         Shortcut        =   +{F12}
      End
   End
   Begin VB.Menu menuZonas 
      Caption         =   "&Zonas"
      Begin VB.Menu mnuAddZona 
         Caption         =   "Agregar nueva zona"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnuEditarZona 
         Caption         =   "Editar zona"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEliminarZona 
         Caption         =   "Eliminar zona"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCopiarZonas 
         Caption         =   "Copiar configuración de zona actual a todas con mismo nombre"
      End
      Begin VB.Menu mmmgui 
         Caption         =   "-"
      End
      Begin VB.Menu menuAddNpcSpawn 
         Caption         =   "Agregar npc spawn"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuEditarSpawn 
         Caption         =   "Editar spawn"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEliminarSpawn 
         Caption         =   "Eliminar spawn"
         Enabled         =   0   'False
      End
      Begin VB.Menu mmEliminarHostiles 
         Caption         =   "Eliminar npcs hostiles en areas"
      End
      Begin VB.Menu mmmgui2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReloadZonas 
         Caption         =   "Recargar archivo de zonas y spawns"
      End
   End
   Begin VB.Menu mnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnuInformes 
         Caption         =   "&Informes"
      End
      Begin VB.Menu mnuModoCaminata 
         Caption         =   "Modalidad &Caminata"
      End
      Begin VB.Menu mnuGRHaBMP 
         Caption         =   "&GRH => BMP"
      End
      Begin VB.Menu mnuOptimizar 
         Caption         =   "Optimi&zar Mapa"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuEditarIndices 
         Caption         =   "Editar Indices.ini"
      End
      Begin VB.Menu mnuActualizarIndices 
         Caption         =   "Actualizar índices..."
      End
   End
   Begin VB.Menu mnuObjSc 
      Caption         =   "mnuObjSc"
      Visible         =   0   'False
      Begin VB.Menu mnuConfigObjTrans 
         Caption         =   "&Utilizar como Objeto de Translados"
      End
   End
   Begin VB.Menu ladder 
      Caption         =   "Tools"
      Begin VB.Menu mmGenMini 
         Caption         =   "Generar minimapa"
      End
      Begin VB.Menu mmGenMini2 
         Caption         =   "Generar minimapas desde el 2+"
      End
      Begin VB.Menu vergraficoslistado 
         Caption         =   "Ver Graficos"
      End
      Begin VB.Menu Ambientacones 
         Caption         =   "Ambientaciones"
      End
   End
   Begin VB.Menu mapppear 
      Caption         =   "Mapear"
      Begin VB.Menu agua 
         Caption         =   "Agua"
      End
      Begin VB.Menu pasto 
         Caption         =   "Pasto"
      End
      Begin VB.Menu arena 
         Caption         =   "Arena"
      End
      Begin VB.Menu hielo 
         Caption         =   "Hielo"
      End
      Begin VB.Menu ins_ladder 
         Caption         =   "Insertar"
         Begin VB.Menu objalazar 
            Caption         =   "Objeto al Azar"
         End
         Begin VB.Menu arbolazar 
            Caption         =   "Arboles al azar"
         End
      End
      Begin VB.Menu blqq 
         Caption         =   "Bloquear"
         Begin VB.Menu blqspaciosvacios 
            Caption         =   "Espacios vacios"
         End
      End
      Begin VB.Menu BloquesOpen 
         Caption         =   "Bloques"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   
'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************

'MOTOR DX8 POR LADDER
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Option Explicit
Public tX         As Integer
Public tY         As Integer
Public LastX      As Integer
Public LastY      As Integer
Public MouseX     As Long
Public MouseY     As Long
Public MouseBoton As Long
Public MouseShift As Long
Private clicX     As Long
Private clicY     As Long

Private shlShell  As Shell32.Shell
Private shlFolder As Shell32.Folder

Private Sub PonerAlAzar(ByVal n As Integer, T As Byte)
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06 by GS
    '*************************************************
    
    On Error GoTo PonerAlAzar_Err
    
    Dim ObjIndex As Long
    Dim NpcIndex As Long
    Dim X, Y, i
    Dim Head    As Integer
    Dim Body    As Integer
    Dim Heading As Byte
    Dim Leer    As New clsIniReader
    i = n

     Call modEdicion.Deshacer_Add(tX, tY, 1, 1) ' Hago deshacer

    Do While i > 0
        X = CInt(RandomNumber(1, MapSize.Width - 1))
        Y = CInt(RandomNumber(1, MapSize.Height - 1))
    
        Select Case T

            Case 0

                If MapData(X, Y).OBJInfo.ObjIndex = 0 Then
                    i = i - 1

                    If cInsertarBloqueo.value = True Then
                        MapData(X, Y).Blocked = 1
                    Else
                        MapData(X, Y).Blocked = 0

                    End If

                    If cNumFunc(2).Text > 0 Then
                        ObjIndex = cNumFunc(2).Text
                        InitGrh MapData(X, Y).ObjGrh, ObjData(ObjIndex).grhindex
                        MapData(X, Y).OBJInfo.ObjIndex = ObjIndex
                        MapData(X, Y).OBJInfo.Amount = Val(cCantFunc(2).Text)

                        Select Case ObjData(ObjIndex).ObjType ' GS

                            Case 4, 8, 10, 22 ' Arboles, Carteles, Foros, Yacimientos
                                MapData(X, Y).Graphic(3) = MapData(X, Y).ObjGrh

                        End Select

                    End If

                End If

            Case 1

                If (MapData(X, Y).Blocked And &HF) <> &HF Then
                    i = i - 1

                    If cNumFunc(T - 1).Text > 0 Then
                        NpcIndex = cNumFunc(T - 1).Text
                        Body = NpcData(NpcIndex).Body
                        Head = NpcData(NpcIndex).Head
                        Heading = NpcData(NpcIndex).Heading
                        Call MakeChar(NextOpenChar(), Body, Head, Heading, CInt(X), CInt(Y))
                        MapData(X, Y).NpcIndex = NpcIndex

                    End If

                End If

            Case 2

                If (MapData(X, Y).Blocked And &HF) <> &HF Then
                    i = i - 1

                    If cNumFunc(T - 1).Text >= 0 Then
                        NpcIndex = cNumFunc(T - 1).Text
                        Body = NpcData(NpcIndex).Body
                        Head = NpcData(NpcIndex).Head
                        Heading = NpcData(NpcIndex).Heading
                        Call MakeChar(NextOpenChar(), Body, Head, Heading, CInt(X), CInt(Y))
                        MapData(X, Y).NpcIndex = NpcIndex

                    End If

                End If

        End Select

        DoEvents
    Loop

    
    Exit Sub

PonerAlAzar_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.PonerAlAzar", Erl)
    Resume Next
    
End Sub

Private Sub bloqqq_Click()
    
    On Error GoTo bloqqq_Click_Err
    
    Dim X As Integer
    Dim Y As Integer
    Dim i As Long

    For X = 1 To MapSize.Width
        For Y = 1 To MapSize.Height

            If MapData(X, Y).Graphic(1).grhindex = 1 Then
                MapData(X, Y).Blocked = 1

            End If

            ' If MapData(X, y).OBJInfo.objindex = 472 Then
            ' MapData(X, y).OBJInfo.objindex = 0
            ' MapData(X, y).Graphic(3).grhindex = 735
            '  MapData(x, y).Graphic(3).grhindex = 738
            
            ' End If
        Next Y
    Next X

    
    Exit Sub

bloqqq_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.bloqqq_Click", Erl)
    Resume Next
    
End Sub


Private Sub agua_Click()
    
    On Error GoTo agua_Click_Err
    
    cGrh.Text = DameGrhIndex(137)

    Call modPaneles.VistaPreviaDeSup

    
    Exit Sub

agua_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.agua_Click", Erl)
    Resume Next
    
End Sub

Private Sub Ambientacones_Click()
    
    On Error GoTo Ambientacones_Click_Err
    
    AmbientacionesForm.Show , FrmMain

    
    Exit Sub

Ambientacones_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Ambientacones_Click", Erl)
    Resume Next
    
End Sub



Private Sub arena_Click()
    
    On Error GoTo arena_Click_Err
    
    cGrh.Text = DameGrhIndex(245)

    Call modPaneles.VistaPreviaDeSup

    
    Exit Sub

arena_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.arena_Click", Erl)
    Resume Next
    
End Sub

Private Sub BloqAll_Click()
    
    On Error GoTo BloqAll_Click_Err
    
    Dim i As Integer
    
    If maskBloqueo = &HF Then
        For i = 0 To 3
            chkBloqueo(i).value = vbUnchecked
        Next
        maskBloqueo = 0

    Else
        For i = 0 To 3
            chkBloqueo(i).value = vbChecked
        Next
        maskBloqueo = &HF
    End If
    
    Exit Sub

BloqAll_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.BloqAll_Click", Erl)
    Resume Next
    
End Sub

Private Sub BloquesOpen_Click()
    
    On Error GoTo BloquesOpen_Click_Err
    
    Call CargarBloq

    
    Exit Sub

BloquesOpen_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.BloquesOpen_Click", Erl)
    Resume Next
    
End Sub

Private Sub blqspaciosvacios_Click()
    
    On Error GoTo blqspaciosvacios_Click_Err
    
    Dim X As Integer
    Dim Y As Integer
    Dim i As Long

    For Y = 1 To MapSize.Height
        For X = 1 To MapSize.Width

            If MapData(X, Y).Graphic(1).grhindex = 0 Or MapData(X, Y).Graphic(1).grhindex = 1 Then
                MapData(X, Y).Blocked = 1

            End If

        Next X
    Next Y

    Exit Sub

blqspaciosvacios_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.blqspaciosvacios_Click", Erl)
    Resume Next
    
End Sub

Private Sub borrarnegros_Click()
    
    On Error GoTo borrarnegros_Click_Err
    
    Dim X As Integer
    Dim Y As Integer
    Dim i As Long

    For Y = 1 To MapSize.Height
        For X = 1 To MapSize.Width

            If MapData(X, Y).Graphic(2).grhindex = 7284 Or MapData(X, Y).Graphic(2).grhindex = 7303 Or MapData(X, Y).Graphic(2).grhindex = 7304 _
               Or MapData(X, Y).Graphic(2).grhindex = 7308 Or MapData(X, Y).Graphic(2).grhindex = 7310 Or MapData(X, Y).Graphic(2).grhindex = 7315 Or MapData(X, Y).Graphic(2).grhindex = 7316 _
               Or MapData(X, Y).Graphic(2).grhindex = 7306 Or MapData(X, Y).Graphic(2).grhindex = 7328 Or MapData(X, Y).Graphic(2).grhindex = 7327 Or MapData(X, Y).Graphic(2).grhindex = 7357 _
               Or MapData(X, Y).Graphic(2).grhindex = 29382 Or MapData(X, Y).Graphic(2).grhindex = 29384 Or MapData(X, Y).Graphic(2).grhindex = 29383 Or MapData(X, Y).Graphic(2).grhindex = 7290 Or MapData(X, Y).Graphic(2).grhindex = 7291 Or MapData(X, Y).Graphic(2).grhindex = 7358 Or MapData(X, Y).Graphic(2).grhindex = 7376 _
               Or MapData(X, Y).Graphic(2).grhindex = 7313 Or MapData(X, Y).Graphic(2).grhindex = 7314 _
               Or MapData(X, Y).Graphic(2).grhindex = 29379 Or MapData(X, Y).Graphic(2).grhindex = 29649 Or MapData(X, Y).Graphic(2).grhindex = 29393 Or MapData(X, Y).Graphic(2).grhindex = 29401 Or MapData(X, Y).Graphic(2).grhindex = 29403 Or MapData(X, Y).Graphic(2).grhindex = 29366 Or MapData(X, Y).Graphic(2).grhindex = 29388 Or MapData(X, Y).Graphic(2).grhindex = 29390 Or MapData(X, Y).Graphic(2).grhindex = 29392 Or MapData(X, Y).Graphic(2).grhindex = 29395 Or MapData(X, Y).Graphic(2).grhindex = 29396 Or MapData(X, Y).Graphic(2).grhindex = 29399 Or MapData(X, Y).Graphic(2).grhindex = 29398 Or MapData(X, Y).Graphic(2).grhindex = 29397 Or MapData(X, Y).Graphic(2).grhindex = 29407 Or MapData(X, Y).Graphic(2).grhindex = 29408 Or MapData(X, Y).Graphic(2).grhindex = 29409 Or MapData(X, Y).Graphic(2).grhindex = 29410 Or MapData(X, Y).Graphic(2).grhindex = 29373 Or MapData(X, Y).Graphic(2).grhindex = 29372 _
               Or MapData(X, Y).Graphic(2).grhindex = 7321 Or MapData(X, Y).Graphic(2).grhindex = 7297 Or MapData(X, Y).Graphic(2).grhindex = 7300 Or MapData(X, Y).Graphic(2).grhindex = 7301 _
               Or MapData(X, Y).Graphic(2).grhindex = 7302 Or MapData(X, Y).Graphic(2).grhindex = 29619 Or MapData(X, Y).Graphic(2).grhindex = 7311 _
               Or MapData(X, Y).Graphic(2).grhindex = 29612 Or MapData(X, Y).Graphic(2).grhindex = 29630 Or MapData(X, Y).Graphic(2).grhindex = 29618 Or MapData(X, Y).Graphic(2).grhindex = 29634 Or MapData(X, Y).Graphic(2).grhindex = 29625 Or MapData(X, Y).Graphic(2).grhindex = 29628 Or MapData(X, Y).Graphic(2).grhindex = 29629 Or MapData(X, Y).Graphic(2).grhindex = 29631 Or MapData(X, Y).Graphic(2).grhindex = 29632 Or MapData(X, Y).Graphic(2).grhindex = 29637 Or MapData(X, Y).Graphic(2).grhindex = 29638 Or MapData(X, Y).Graphic(2).grhindex = 29640 Or MapData(X, Y).Graphic(2).grhindex = 29642 Or MapData(X, Y).Graphic(2).grhindex = 29643 Or MapData(X, Y).Graphic(2).grhindex = 29645 Or MapData(X, Y).Graphic(2).grhindex = 29646 Or MapData(X, Y).Graphic(2).grhindex = 29655 Or MapData(X, Y).Graphic(2).grhindex = 29656 Or MapData(X, Y).Graphic(2).grhindex = 29647 Or MapData(X, Y).Graphic(2).grhindex = 29648 Or MapData(X, Y).Graphic(2).grhindex = 29651 Or MapData(X, Y).Graphic(2).grhindex = 29653 _
               Or MapData(X, Y).Graphic(2).grhindex = 7325 Or MapData(X, Y).Graphic(2).grhindex = 7326 Or MapData(X, Y).Graphic(2).grhindex = 7354 _
               Or MapData(X, Y).Graphic(2).grhindex = 7373 Or MapData(X, Y).Graphic(2).grhindex = 7371 Or MapData(X, Y).Graphic(2).grhindex = 7365 _
               Or MapData(X, Y).Graphic(2).grhindex = 29597 Or MapData(X, Y).Graphic(2).grhindex = 29595 Or MapData(X, Y).Graphic(2).grhindex = 29596 _
               Or MapData(X, Y).Graphic(2).grhindex = 29571 Or MapData(X, Y).Graphic(2).grhindex = 29608 Or MapData(X, Y).Graphic(2).grhindex = 29607 _
               Or MapData(X, Y).Graphic(2).grhindex = 29588 Or MapData(X, Y).Graphic(2).grhindex = 29590 Or MapData(X, Y).Graphic(2).grhindex = 29583 _
               Or MapData(X, Y).Graphic(2).grhindex = 29584 Or MapData(X, Y).Graphic(2).grhindex = 29586 _
               Or MapData(X, Y).Graphic(2).grhindex = 7369 Or MapData(X, Y).Graphic(2).grhindex = 7367 Or MapData(X, Y).Graphic(2).grhindex = 7352 _
               Or MapData(X, Y).Graphic(2).grhindex = 7375 Or MapData(X, Y).Graphic(2).grhindex = 7351 Or MapData(X, Y).Graphic(2).grhindex = 7368 _
               Or MapData(X, Y).Graphic(2).grhindex = 7332 Or MapData(X, Y).Graphic(2).grhindex = 7339 Or MapData(X, Y).Graphic(2).grhindex = 7366 _
               Or MapData(X, Y).Graphic(2).grhindex = 7360 Or MapData(X, Y).Graphic(2).grhindex = 7338 Or MapData(X, Y).Graphic(2).grhindex = 7363 Or MapData(X, Y).Graphic(2).grhindex = 29582 Or MapData(X, Y).Graphic(2).grhindex = 29581 Or MapData(X, Y).Graphic(2).grhindex = 29580 _
               Or MapData(X, Y).Graphic(2).grhindex = 29593 Or MapData(X, Y).Graphic(2).grhindex = 29594 Or MapData(X, Y).Graphic(2).grhindex = 29570 _
               Or MapData(X, Y).Graphic(2).grhindex = 29599 Or MapData(X, Y).Graphic(2).grhindex = 29601 Or MapData(X, Y).Graphic(2).grhindex = 29591 _
               Or MapData(X, Y).Graphic(2).grhindex = 7349 Or MapData(X, Y).Graphic(2).grhindex = 7348 Or MapData(X, Y).Graphic(2).grhindex = 7345 _
               Or MapData(X, Y).Graphic(2).grhindex = 29606 Or MapData(X, Y).Graphic(2).grhindex = 29605 Or MapData(X, Y).Graphic(2).grhindex = 29577 _
               Or MapData(X, Y).Graphic(2).grhindex = 7350 Or MapData(X, Y).Graphic(2).grhindex = 7362 Or MapData(X, Y).Graphic(2).grhindex = 7338 _
               Or MapData(X, Y).Graphic(2).grhindex = 7317 Or MapData(X, Y).Graphic(2).grhindex = 7319 Or MapData(X, Y).Graphic(2).grhindex = 8272 Or MapData(X, Y).Graphic(2).grhindex = 8263 Then
                Rem 7357 Or 7358 Or 7375 Or 7376 Or 22590 Or 22588 Or 22594 Or 22595 Or 22582 Or 22583 Then
                MapData(X, Y).Graphic(2).grhindex = 0

            End If
        
            If MapData(X, Y).Graphic(1).grhindex = 0 Then
                MapData(X, Y).Graphic(1).grhindex = 1

            End If

        Next X
    Next Y
      
    
    Exit Sub

borrarnegros_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.borrarnegros_Click", Erl)
    Resume Next
    
End Sub

Private Sub cAgregarFuncalAzar_Click(Index As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cAgregarFuncalAzar_Click_Err
    

    If IsNumeric(cCantFunc(Index).Text) = False Or cCantFunc(Index).Text > 200 Then
        MsgBox "El Valor de Cantidad introducido no es soportado!" & vbCrLf & "El valor maximo es 200.", vbCritical
        Exit Sub

    End If

    cAgregarFuncalAzar(Index).Enabled = False
    Call PonerAlAzar(CInt(cCantFunc(Index).Text), 1 + (IIf(Index = 2, -1, Index)))
    cAgregarFuncalAzar(Index).Enabled = True

    
    Exit Sub

cAgregarFuncalAzar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cAgregarFuncalAzar_Click", Erl)
    Resume Next
    
End Sub

Private Sub cCantFunc_Change(Index As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cCantFunc_Change_Err
    

    If Val(cCantFunc(Index)) < 1 Then
        cCantFunc(Index).Text = 1

    End If

    If Val(cCantFunc(Index)) > 10000 Then
        cCantFunc(Index).Text = 10000

    End If

    
    Exit Sub

cCantFunc_Change_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cCantFunc_Change", Erl)
    Resume Next
    
End Sub

Private Sub cCapas_Change()
    
    On Error GoTo cCapas_Change_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 31/05/06
    '*************************************************
    If Val(cCapas.Text) < 1 Then
        cCapas.Text = 1

    End If

    If Val(cCapas.Text) > 4 Then
        cCapas.Text = 4

    End If

    cCapas.Tag = vbNullString

    
    Exit Sub

cCapas_Change_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cCapas_Change", Erl)
    Resume Next
    
End Sub

Private Sub cCapas_KeyPress(KeyAscii As Integer)
    
    On Error GoTo cCapas_KeyPress_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If IsNumeric(Chr(KeyAscii)) = False Then KeyAscii = 0

    
    Exit Sub

cCapas_KeyPress_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cCapas_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub cFiltro_GotFocus(Index As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cFiltro_GotFocus_Err
    
    HotKeysAllow = False

    
    Exit Sub

cFiltro_GotFocus_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cFiltro_GotFocus", Erl)
    Resume Next
    
End Sub

Private Sub cFiltro_KeyPress(Index As Integer, KeyAscii As Integer)
    
    On Error GoTo cFiltro_KeyPress_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If KeyAscii = 13 Then
        Call Filtrar(Index)

    End If

    
    Exit Sub

cFiltro_KeyPress_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cFiltro_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub cFiltro_LostFocus(Index As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cFiltro_LostFocus_Err
    
    HotKeysAllow = True

    
    Exit Sub

cFiltro_LostFocus_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cFiltro_LostFocus", Erl)
    Resume Next
    
End Sub

Private Sub cGrh_KeyPress(KeyAscii As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************

    On Error GoTo Fallo

    If KeyAscii = 13 Then
        Call fPreviewGrh(cGrh.Text)

        If FrmMain.cGrh.ListCount > 5 Then
            FrmMain.cGrh.RemoveItem 0

        End If

        FrmMain.cGrh.AddItem FrmMain.cGrh.Text

    End If

    Exit Sub
Fallo:
    cGrh.Text = 1

End Sub

Private Sub Check1_Click()
    
    On Error GoTo Check1_MouseUp_Err
    
    If LoadingMap Then Exit Sub

    If MapDat.Lluvia = 0 Then

        MapDat.Lluvia = 1
        Call AddtoRichTextBox(FrmMain.RichTextBox1, "Lluvia en mapa activada.", 255, 255, 255, False, True, False)
    Else
        MapDat.Lluvia = 0
        Call AddtoRichTextBox(FrmMain.RichTextBox1, "Lluvia en mapa desactivada.", 255, 255, 255, False, True, False)

    End If

    MapInfo.Changed = 1

    
    Exit Sub

Check1_MouseUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Check1_MouseUp", Erl)
    Resume Next
    
End Sub



Private Sub Check2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo Check2_MouseUp_Err
    

    If Nieba = 0 Then
        Nieba = 1
        Call AddtoRichTextBox(FrmMain.RichTextBox1, "Nieve en mapa activada.", 255, 255, 255, False, True, False)
    Else
        Nieba = 0
        Call AddtoRichTextBox(FrmMain.RichTextBox1, "Nieve en mapa desactivada.", 255, 255, 255, False, True, False)

    End If

    MapInfo.Changed = 1

    
    Exit Sub

Check2_MouseUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Check2_MouseUp", Erl)
    Resume Next
    
End Sub





Private Sub ListaParticulas_Click()
        FrmMain.numerodeparticula.Text = ReadField(2, FrmMain.ListaParticulas.List(FrmMain.ListaParticulas.ListIndex), Asc("#"))
End Sub

Private Sub lstMaps_DblClick()
Dim Num As Integer
Num = mid$(lstMaps.Text, 1, InStr(1, lstMaps.Text, " ") - 1)

If MapInfo.Changed = 1 Then
    Call mnuGuardarMapa_Click
End If
modMapIO.AbrirMapa App.Path & "\..\Resources\Mapas\Mapa" & Num & ".csm"
DoEvents
mnuReAbrirMapa.Enabled = True
EngineRun = True
End Sub

Private Sub lvButtons_H1_Click()
mnuGuardarMapa_Click
End Sub



Private Sub menuAddNpcSpawn_Click()
mnuNpcSpawn.Checked = True
mnuQuitarFunciones_Click
RectanguloModo = 2
End Sub

Private Sub MiniMap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    On Error GoTo MiniMap_MouseDown_Err
    
    If X < 0 Then X = 0
    If Y < 0 Then Y = 0
    If X > MiniMap.Width Then X = MiniMap.Width
    If Y > MiniMap.Height Then Y = MiniMap.Height
    
    UserPos.X = MapSize.Width * X / MiniMap.Width
    UserPos.Y = MapSize.Height * Y / MiniMap.Height
    bRefreshRadar = True

    
    Exit Sub

MiniMap_MouseDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.MiniMap_MouseDown", Erl)
    Resume Next
    
End If
End Sub

Private Sub mmCopiarCapa1_Click()
mmCopiarCapa1.Checked = Not mmCopiarCapa1.Checked
End Sub

Private Sub mmCopiarCapa2_Click()
mmCopiarCapa2.Checked = Not mmCopiarCapa2.Checked
End Sub

Private Sub mmCopiarCapa3_Click()
mmCopiarCapa3.Checked = Not mmCopiarCapa3.Checked
End Sub

Private Sub mmCopiarCapa4_Click()
mmCopiarCapa4.Checked = Not mmCopiarCapa4.Checked
End Sub

Private Sub mmEliminarHostiles_Click()
Dim i As Integer
Dim X As Integer
Dim Y As Integer
For i = 1 To NumSpawns
    With NpcSpawn(i)
        If .Map = UserMap And .Deleted = 0 And .X > 0 Then
            For Y = .Y To .Y2
                For X = .X To .X2
                    If MapData(X, Y).NpcIndex > 0 Then
                        If NpcData(MapData(X, Y).NpcIndex).Hostile = 1 Then
                            EraseChar (MapData(X, Y).CharIndex)
                            MapData(X, Y).NpcIndex = 0
                        End If
                    End If
                Next X
            Next Y
        End If
    End With
Next i

End Sub

Private Sub mmGenMini_Click()
    Dim tmpPic As StdPicture
    Dim picNr As Long
    Dim sFileName As String
    Dim maxCx As Long, maxCy As Long
    Dim picWidth As Long, picHeight As Long
    
   
    frmRenderer.Show
   
    frmRenderer.PicGrande.ScaleMode = vbPixels
    frmRenderer.PicGrande.AutoRedraw = True
    frmRenderer.PicGrande.BorderStyle = 0&
    Dim X2 As Integer
    Dim Y2 As Integer
    
    
     maxCy = MapSize.Width * 4
     maxCx = MapSize.Height * 4
     
     
    For Y2 = 0 To ((MapSize.Height - 1) \ 100)
       For X2 = 0 To ((MapSize.Width - 1) \ 100)
         Set SurfaceDB.Used = New Collection
         Call engine.MapCapture(False, True, X2 * 100, Y2 * 100)
         SurfaceDB.UnloadUnused
      Next X2
    Next Y2
    
    frmRenderer.PicGrande.AutoRedraw = False
    
    
    Unload frmRenderer
    
    Shell (App.Path & "\UnirMinimapa.exe " & UserMap)
    DoEvents
    Sleep 1000
    DoEvents
    CargarMinimap
End Sub

Private Sub mmGenMini2_Click()
Dim i As Integer

For i = 2 To FrmMain.lstMaps.ListCount
    UserMap = mid$(lstMaps.List(i), 1, InStr(1, lstMaps.List(i), " ") - 1)
    FrmMain.Label16.Caption = "Map " & UserMap
    modMapIO.AbrirMapa App.Path & "\..\Resources\Mapas\Mapa" & i & ".csm"
    DoEvents
    mmGenMini_Click
Next i
End Sub

Private Sub mmGuardarCliente_Click()
    Call SaveMapMagicClient(App.Path & "\..\Resources\Mapas\Mapa" & UserMap & ".map")
End Sub

Private Sub mmMapSize_Click()
frmMapSize.Show
End Sub

Private Sub mnuAddZona_Click()
mnuZonas.Checked = True
mnuQuitarFunciones_Click
RectanguloModo = 1
End Sub


Private Sub mnuCopiarZonas_Click()
If SelectedZona > 0 Then
    Dim i As Integer
    With Zona(SelectedZona)
        For i = 1 To NumZonas
            If Zona(i).Map = UserMap Then
                If Zona(i).Zona_name = Zona(SelectedZona).Zona_name Then
                    Zona(i).Terreno = .Terreno
                    Zona(i).Ambient = .Ambient
                    Zona(i).Backup = .Backup
                    Zona(i).Base_light = .Base_light
                    Zona(i).Faccion = .Faccion
                    Zona(i).Lluvia = .Lluvia
                    Zona(i).MaxLevel = .MaxLevel
                    Zona(i).MinLevel = .MinLevel
                    Zona(i).Musica1 = .Musica1
                    Zona(i).Musica2 = .Musica2
                    Zona(i).Musica3 = .Musica3
                    Zona(i).Newbie = .Newbie
                    Zona(i).Niebla = .Niebla
                    Zona(i).Nieve = .Nieve
                    Zona(i).SalidaMap = .SalidaMap
                    Zona(i).SalidaX = .SalidaX
                    Zona(i).SalidaY = .SalidaY
                    Zona(i).Segura = .Segura
                    Zona(i).SinInvi = .SinInvi
                    Zona(i).SinMagia = .SinMagia
                    Zona(i).SinMascotas = .SinMascotas
                    Zona(i).SinResucitar = .SinResucitar
                    Zona(i).SoloClanes = .SoloClanes
                    Zona(i).SoloFaccion = .SoloFaccion
                    Call SaveZona(i, Zona(i))
                End If
            End If
        Next i
    End With
End If
End Sub

Private Sub mnuEditarZona_Click()
If SelectedZona > 0 Then
    frmZonaInfo.OpenZona (SelectedZona)
    frmZonaInfo.Show
End If
End Sub

Private Sub mnuEditarSpawn_Click()
If SelectedSpawn > 0 Then
    frmSpawnInfo.OpenSpawn (SelectedSpawn)
    frmSpawnInfo.Show
End If
End Sub

Private Sub mnuEliminarSpawn_Click()
If SelectedSpawn > 0 Then
    If MsgBox("¿Realmente desea eliminar el spawn " & SelectedSpawn & "?", vbExclamation + vbYesNo) = vbYes Then
        Call DeleteSpawn(SelectedSpawn)
        SelectedSpawn = 0
        DibujarSpawns
    End If
End If
End Sub

Private Sub mnuEliminarZona_Click()
If SelectedZona > 0 Then
    If MsgBox("¿Realmente desea eliminar la zona " & SelectedZona & " - " & Zona(SelectedZona).Zona_name & "?", vbExclamation + vbYesNo) = vbYes Then
        Call DeleteZona(SelectedZona)
        SelectedZona = 0
        DibujarZonas
    End If
End If
End Sub


Private Sub mnuInsertarTransladosAdyasentes_Click()

End Sub

Private Sub mnuNpcSpawn_Click()
mnuNpcSpawn.Checked = (mnuNpcSpawn.Checked = False)
End Sub

Private Sub mnuReloadZonas_Click()
Call LoadZonas
Call LoadNpcSpawn
End Sub

Private Sub mnuVerMarco_Click()

    On Error GoTo mnuVerMarco_Click_Err
    
    VerMarco = (VerMarco = False)
    mnuVerMarco.Checked = VerMarco

    
    Exit Sub

mnuVerMarco_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuVerMarco_Click", Erl)
    Resume Next
End Sub

Private Sub mnuZonas_Click()
mnuZonas.Checked = (mnuZonas.Checked = False)
End Sub

Private Sub numerodeparticula_Change()
'FrmMain.ListaParticulas.ListIndex = FrmMain.numerodeparticula.Text
End Sub

Private Sub Seguro_Click()
    
    On Error GoTo Check4_MouseUp_Err
    
    If LoadingMap Then Exit Sub
    

    If MapDat.Seguro = 1 Then
        MapDat.Seguro = 0
        Call AddtoRichTextBox(FrmMain.RichTextBox1, "Mapa inseguro", 255, 255, 255, False, True, False)
    Else
        MapDat.Seguro = 1
        Call AddtoRichTextBox(FrmMain.RichTextBox1, "Mapa seguro.", 255, 255, 255, False, True, False)

    End If
    
    MapInfo.Changed = 1

    
    Exit Sub

Check4_MouseUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Check4_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub BackUp_Click()
    
    On Error GoTo Check5_MouseUp_Err
    
    If LoadingMap Then Exit Sub
    

    If MapDat.backup_mode = 1 Then
        MapDat.backup_mode = 0
        Call AddtoRichTextBox(FrmMain.RichTextBox1, "Backup de mapa desactivado.", 255, 255, 255, False, True, False)
    Else
        MapDat.backup_mode = 1
        Call AddtoRichTextBox(FrmMain.RichTextBox1, "Backup de mapa activado.", 255, 255, 255, False, True, False)

    End If
    
    MapInfo.Changed = 1

    
    Exit Sub

Check5_MouseUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Check5_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub Check6_Click()
    
    On Error GoTo Check6_Click_Err
    
    AlphaTecho = Not AlphaTecho

    
    Exit Sub

Check6_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Check6_Click", Erl)
    Resume Next
    
End Sub


Private Sub Command14_Click()
    
    On Error GoTo Command14_Click_Err
    
    Dim Y As Integer
    Dim X As Integer

    For Y = 1 To MapSize.Height
        For X = 1 To MapSize.Width

            If MapData(X, Y).particle_Index = 180 Then
                MapData(X, Y).particle_Index = 0

            End If

        Next X
    Next Y

    
    Exit Sub

Command14_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Command14_Click", Erl)
    Resume Next
    
End Sub


Private Sub chkBloqueo_Click(Index As Integer)
    
    On Error GoTo chkBloqueo_Click_Err
    
    maskBloqueo = maskBloqueo Xor 2 ^ Index

    
    Exit Sub

chkBloqueo_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.chkBloqueo_Click", Erl)
    Resume Next
    
End Sub



Private Sub cmdDM_Click(Index As Integer)
    
    On Error GoTo cmdDM_Click_Err
    
    frmConfigSup.DespMosaic.value = vbChecked

    Select Case Index

        Case 0 'A
    
            frmConfigSup.DMLargo.Text = Val(frmConfigSup.DMLargo.Text) + 1

        Case 1 '<
            frmConfigSup.DMAncho.Text = Val(frmConfigSup.DMAncho.Text) + 1

        Case 2 '>
            frmConfigSup.DMAncho.Text = Val(frmConfigSup.DMAncho.Text) - 1

        Case 3 'V
            frmConfigSup.DMLargo.Text = Val(frmConfigSup.DMLargo.Text) - 1

        Case 4 '0
            frmConfigSup.DMAncho.Text = 0
            frmConfigSup.DMLargo.Text = 0

    End Select

    
    Exit Sub

cmdDM_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cmdDM_Click", Erl)
    Resume Next
    
End Sub

Private Sub Remplazograficos()
    
    On Error GoTo Remplazograficos_Err
    

    Dim Y As Integer
    Dim X As Integer
    Dim c As Integer
    Dim D As Integer
    

    
    For Y = 1 To MapSize.Height
        For X = 1 To MapSize.Width
    
            ' If MapData(X, y).OBJInfo.objindex > 0 Then
            '  If ObjData(MapData(X, y).OBJInfo.objindex).ObjType = 4 Then
            '   If MapData(X, y).Graphic(3).grhindex = MapData(X, y).ObjGrh.grhindex Then MapData(X, y).Graphic(3).grhindex = 0
            '   MapData(X, y).OBJInfo.objindex = 0
            '   MapData(X, y).OBJInfo.Amount = 0
            '   MapData(X, y).Blocked = 0
            ' End If
            '  End If
        
'            If MapData(X, y).Graphic(c).grhindex = txtGRH.Text Then
'                MapData(X, y).Graphic(D).grhindex = TxtGrh2.Text
            
'                'InitGrh MapData(X, y).Graphic(2), 0
'                MapData(X, y).Graphic(2).grhindex = TxtGrh.Text
'                InitGrh MapData(X, y).Graphic(2), TxtGrh2.Text
            
'            End If
        
            '        If MapData(X, y).Graphic(3).grhindex = 12445 Then
            '            MapData(X, y).Graphic(3).grhindex = 0
            '            'InitGrh MapData(X, y).Graphic(2), 0
            '            MapData(X, y).Graphic(2).grhindex = 12445
            '            InitGrh MapData(X, y).Graphic(2), 12445
            '        End If
        
            ' Dim num As Long
        
            ' For num = 943 To 950
            '   If MapData(X, y).Graphic(3).grhindex = num Then
            ' MapData(X, y).Graphic(3).grhindex = 0
            'InitGrh MapData(X, y).Graphic(2), 0
            'MapData(X, y).Graphic(2).grhindex = num
            ' InitGrh MapData(X, y).Graphic(2), num
            ' End If
            ' Next num
        
        Next X
    Next Y

    
    Exit Sub

Remplazograficos_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Remplazograficos", Erl)
    Resume Next
    
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo Form_KeyDown_Err
    

'    If KeyCode = vbKeySpace Then
'        If FrmBloques.Visible = True Then
'            Call InsertarBloque
'
'        End If
'
'    End If

    
    Exit Sub

Form_KeyDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Form_KeyDown", Erl)
    Resume Next
    
End Sub

Private Sub hielo_Click()
    
    On Error GoTo hielo_Click_Err
    
    cGrh.Text = DameGrhIndex(621)

    Call modPaneles.VistaPreviaDeSup

    
    Exit Sub

hielo_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.hielo_Click", Erl)
    Resume Next
    
End Sub

Private Sub Label16_Click()
    
    On Error GoTo Label16_Click_Err
    
   ' Timer4.Enabled = True

    
    Exit Sub

Label16_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Label16_Click", Erl)
    Resume Next
    
End Sub

Private Sub LuzMapa_Change()
    'MapInfo.Changed = 1
End Sub

Private Sub LvBOpcion_Click(Index As Integer)
    
    On Error GoTo LvBOpcion_Click_Err
    

    Select Case Index

        Case 0
            cVerBloqueos.value = (cVerBloqueos.value = False)
            mnuVerBloqueos.Checked = cVerBloqueos.value
            
                If mnuVerBloqueos.Checked = False Then
                    LvBOpcion(0).BackColor = &H80000000
                Else
                    LvBOpcion(0).BackColor = &H80FF80
                End If

        Case 1
            mnuVerTranslados.Checked = (mnuVerTranslados.Checked = False)
            
                If mnuVerTranslados.Checked = False Then
                    LvBOpcion(1).BackColor = &H80000000
                Else
                    LvBOpcion(1).BackColor = &H80FF80
                End If

        Case 2
            mnuVerObjetos.Checked = (mnuVerObjetos.Checked = False)
            
                If mnuVerObjetos.Checked = False Then
                    LvBOpcion(2).BackColor = &H80000000
                Else
                    LvBOpcion(2).BackColor = &H80FF80
                End If

        Case 3
            cVerTriggers.value = (cVerTriggers.value = False)
            mnuVerTriggers.Checked = cVerTriggers.value
            
                If mnuVerTriggers.Checked = False Then
                    LvBOpcion(3).BackColor = &H80000000
                Else
                    LvBOpcion(3).BackColor = &H80FF80
                End If

        Case 4
            mnuVerCapa1.Checked = (mnuVerCapa1.Checked = False)
            
                If mnuVerCapa1.Checked = False Then
                    LvBOpcion(4).BackColor = &H80000000
                Else
                    LvBOpcion(4).BackColor = &H80FF80
                End If

        Case 5
            mnuVerCapa2.Checked = (mnuVerCapa2.Checked = False)
            
                If mnuVerCapa2.Checked = False Then
                    LvBOpcion(5).BackColor = &H80000000
                Else
                    LvBOpcion(5).BackColor = &H80FF80
                End If
        Case 6
            mnuVerCapa3.Checked = (mnuVerCapa3.Checked = False)
            
                If mnuVerCapa3.Checked = False Then
                    LvBOpcion(6).BackColor = &H80000000
                Else
                    LvBOpcion(6).BackColor = &H80FF80
                End If
                
        Case 7
            mnuVerCapa4.Checked = (mnuVerCapa4.Checked = False)
            
                If mnuVerCapa4.Checked = False Then
                    LvBOpcion(7).BackColor = &H80000000
                Else
                    LvBOpcion(7).BackColor = &H80FF80
                End If
        Case 8
            DibujarZonas
        Case 9
            DibujarSpawns
        Case 10
            DibujarHostiles
        Case 12
            AmbientacionesForm.Show , FrmMain
            Call SelectPanel_Click(0)
            modPaneles.VerFuncion 0, True
            cSeleccionarSuperficie.Enabled = True

        Case 17
            
            Call frmOptimizar.cOptimizar_Click

        Case 18
            mnuAutoCompletarSuperficies_Click

        Case 19
            cVerBloqueos.value = False
            cVerTriggers.value = False
            mnuVerParticulas.Checked = True
            
        Case 20
            Call InsertarBloque

        Case 21
            Call frmRemplazo.Show
        Case 22
            Call Todas_las_Particulas_Click
            Call Todas_las_luces_Click
            Call mnuQuitarTriggers_Click
    End Select

    
    Exit Sub

LvBOpcion_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.LvBOpcion_Click", Erl)
    Resume Next
    
End Sub

Private Sub MapFlags_Click(Index As Integer)

    If Not LoadingMap Then

        Dim Flag As Byte
        Flag = 2 ^ Index
    
        MapDat.restrict_mode = Val(MapDat.restrict_mode) Xor Flag
    
        MapInfo.Changed = 1
        
    End If

End Sub


Private Sub mnuAbrirMapaLong_Click()
    Dialog.CancelError = True

    On Error GoTo ErrHandler

    FormatoIAO = False

    DeseaGuardarMapa Dialog.Filename

    ObtenerNombreArchivo False

    If Len(Dialog.Filename) < 3 Then Exit Sub

    If WalkMode = True Then
        Call modGeneral.ToggleWalkMode

    End If
    
    modMapIO.AbrirMapa Dialog.Filename
    DoEvents
    mnuReAbrirMapa.Enabled = True
    EngineRun = True
    
    Exit Sub
ErrHandler:

End Sub

Private Sub mnuActualizarIndices_Click()
    
    On Error GoTo mnuActualizarIndices_Click_Err
    
    frmActualizarIndices.Show , Me

    
    Exit Sub

mnuActualizarIndices_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuActualizarIndices_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuEditarIndices_Click()
Shell "C:\WINDOWS\System32\notepad.exe " & App.Path & "\..\Resources\init\indices.ini", vbNormalFocus


End Sub


Private Sub niebla_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo niebla_MouseUp_Err
    

    If nieblaV = 0 Then
        nieblaV = 1
        Call AddtoRichTextBox(FrmMain.RichTextBox1, "Niebla en mapa activada.", 255, 255, 255, False, True, False)
    Else
        nieblaV = 0
        Call AddtoRichTextBox(FrmMain.RichTextBox1, "Niebla en mapa desactivada.", 255, 255, 255, False, True, False)

    End If

    MapInfo.Changed = 1

    
    Exit Sub

niebla_MouseUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.niebla_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub cInsertarFunc_Click(Index As Integer)
    
    On Error GoTo cInsertarFunc_Click_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If cInsertarFunc(Index).value = True Then
        cQuitarFunc(Index).Enabled = False
        cAgregarFuncalAzar(Index).Enabled = False

        If Index <> 2 Then cCantFunc(Index).Enabled = False
        Call modPaneles.EstSelectPanel((Index) + 3, True)
    Else
        cQuitarFunc(Index).Enabled = True
        cAgregarFuncalAzar(Index).Enabled = True

        If Index <> 2 Then cCantFunc(Index).Enabled = True
        Call modPaneles.EstSelectPanel((Index) + 3, False)

    End If

    
    Exit Sub

cInsertarFunc_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cInsertarFunc_Click", Erl)
    Resume Next
    
End Sub

Private Sub cInsertarTrans_Click()
    
    On Error GoTo cInsertarTrans_Click_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 22/05/06
    '*************************************************
    If cInsertarTrans.value = True Then
        cQuitarTrans.Enabled = False
        Call modPaneles.EstSelectPanel(1, True)
    Else
        cQuitarTrans.Enabled = True
        Call modPaneles.EstSelectPanel(1, False)

    End If

    
    Exit Sub

cInsertarTrans_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cInsertarTrans_Click", Erl)
    Resume Next
    
End Sub

Private Sub cInsertarTrigger_Click()
    
    On Error GoTo cInsertarTrigger_Click_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If cInsertarTrigger.value = True Then
        cQuitarTrigger.Enabled = False
        Call modPaneles.EstSelectPanel(6, True)
    Else
        cQuitarTrigger.Enabled = True
        Call modPaneles.EstSelectPanel(6, False)

    End If

    
    Exit Sub

cInsertarTrigger_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cInsertarTrigger_Click", Erl)
    Resume Next
    
End Sub

Private Sub cmdQuitarFunciones_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cmdQuitarFunciones_Click_Err
    
    Call mnuQuitarFunciones_Click

    
    Exit Sub

cmdQuitarFunciones_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cmdQuitarFunciones_Click", Erl)
    Resume Next
    
End Sub






Private Sub Command1_Click(Index As Integer)
  
    Dim X As Integer
    Dim Y As Integer
    Dim Map As Integer
    
    Map = 400 'initial map
    
    Dim W As Integer
    Dim H As Integer
    
    Dim MapasX As Integer
    Dim MapasY As Integer
    
    MapasX = 3
    MapasY = 3
    
    W = 24 + (MapasX * 86)
    H = 20 + (MapasY * 90)
    
    Dim MapasMagic() As Integer
    
    ReDim MapasMagic(1 To MapasX, 1 To MapasY)
    
    MapasMagic(2, 1) = 315
    MapasMagic(2, 2) = 314
    MapasMagic(1, 3) = 313
    MapasMagic(2, 3) = 311
    MapasMagic(3, 3) = 312
    
    ReDim MapDataMagic(1 To W, 1 To H)
  
  
If Index = 0 Then

   'Para regenerar los mapas de clientes
   'For X = 2 To 35
   '     modMapIO.AbrirMapa App.Path & "\..\Resources\Mapas\Mapa" & X & ".csm"
   '     Call SaveMapMagicClient(App.Path & "\..\Resources\Mapas\Mapa" & X & ".map")
   'Next X
   'Exit Sub

    
    For Y = 1 To MapasY
        For X = 1 To MapasX
            
            Map = MapasMagic(X, Y)
            If Map > 0 Then
                Dim esBordeX As Boolean
                Dim esBordeY As Boolean
                
                If X > 1 Then
                    esBordeX = MapasMagic(X - 1, Y) = 0
                End If
                If Y > 1 Then
                    esBordeY = MapasMagic(X, Y - 1) = 0
                End If
                Call LoadMapMagic(Map, X, Y, 0, esBordeX, esBordeY)
            End If
            
            
            'Map = NextMapMagic
        Next X
    Next Y
    
    MapName = InputBox("Nombre del mapa:", MapName)
    
    Call SaveMapMagic(W, H, MapName)
    
    modMapIO.AbrirMapa App.Path & "\..\Resources\Mapas\" & MapName & ".csm"
    DoEvents
    mnuReAbrirMapa.Enabled = True
    EngineRun = True
    
    MapSize.Width = W
    MapSize.Height = H
    
    Call SaveMapMagicClient(App.Path & "\..\Resources\Mapas\" & MapName & ".map")
    
ElseIf Index = 1 Then
    
End If

End Sub

Private Sub cUnionManual_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cUnionManual_Click_Err
    
    cInsertarTrans.value = (cUnionManual.value = True)
    Call cInsertarTrans_Click

    
    Exit Sub

cUnionManual_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cUnionManual_Click", Erl)
    Resume Next
    
End Sub

Private Sub cverBloqueos_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cverBloqueos_Click_Err
    
    mnuVerBloqueos.Checked = cVerBloqueos.value

    
    Exit Sub

cverBloqueos_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cverBloqueos_Click", Erl)
    Resume Next
    
End Sub

Private Sub cverTriggers_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cverTriggers_Click_Err
    
    mnuVerTriggers.Checked = cVerTriggers.value

    
    Exit Sub

cverTriggers_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cverTriggers_Click", Erl)
    Resume Next
    
End Sub

Private Sub cNumFunc_KeyPress(Index As Integer, KeyAscii As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cNumFunc_KeyPress_Err
    

    If KeyAscii = 13 Then
        Dim Cont As String
        Cont = FrmMain.cNumFunc(Index).Text
        Call cNumFunc_LostFocus(Index)

        If Cont <> FrmMain.cNumFunc(Index).Text Then Exit Sub
        If FrmMain.cNumFunc(Index).ListCount > 5 Then
            FrmMain.cNumFunc(Index).RemoveItem 0

        End If

        FrmMain.cNumFunc(Index).AddItem FrmMain.cNumFunc(Index).Text
        Exit Sub
    ElseIf KeyAscii = 8 Then
    
    ElseIf IsNumeric(Chr(KeyAscii)) = False Then
        KeyAscii = 0
        Exit Sub

    End If

    
    Exit Sub

cNumFunc_KeyPress_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cNumFunc_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub cNumFunc_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cNumFunc_KeyUp_Err
    

    If cNumFunc(Index).Text = vbNullString Then
        FrmMain.cNumFunc(Index).Text = IIf(Index = 1, 500, 1)

    End If

    
    Exit Sub

cNumFunc_KeyUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cNumFunc_KeyUp", Erl)
    Resume Next
    
End Sub

Private Sub cNumFunc_LostFocus(Index As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cNumFunc_LostFocus_Err
    

    If Index = 0 Then
        If FrmMain.cNumFunc(Index).Text > 499 Or FrmMain.cNumFunc(Index).Text < 1 Then
            FrmMain.cNumFunc(Index).Text = 1

        End If

    ElseIf Index = 1 Then

        If FrmMain.cNumFunc(Index).Text < 500 Or FrmMain.cNumFunc(Index).Text > 32000 Then
            FrmMain.cNumFunc(Index).Text = 500

        End If

    ElseIf Index = 2 Then

        If FrmMain.cNumFunc(Index).Text < 1 Or FrmMain.cNumFunc(Index).Text > 32000 Then
            FrmMain.cNumFunc(Index).Text = 1

        End If

    End If

    
    Exit Sub

cNumFunc_LostFocus_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cNumFunc_LostFocus", Erl)
    Resume Next
    
End Sub

Private Sub cInsertarBloqueo_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 29/05/06
    '*************************************************
    
    On Error GoTo cInsertarBloqueo_Click_Err
    cInsertarBloqueo.Tag = vbNullString

    If cInsertarBloqueo.value = True Then
        cQuitarBloqueo.Enabled = False
        Call modPaneles.EstSelectPanel(2, True)
    Else
        cQuitarBloqueo.Enabled = True
        Call modPaneles.EstSelectPanel(2, False)

    End If

    
    Exit Sub

cInsertarBloqueo_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cInsertarBloqueo_Click", Erl)
    Resume Next
    
End Sub

Private Sub cQuitarBloqueo_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cQuitarBloqueo_Click_Err
    
    cInsertarBloqueo.Tag = vbNullString

    If cQuitarBloqueo.value = True Then
        cInsertarBloqueo.Enabled = False
        Call modPaneles.EstSelectPanel(2, True)
    Else
        cInsertarBloqueo.Enabled = True
        Call modPaneles.EstSelectPanel(2, False)

    End If

    
    Exit Sub

cQuitarBloqueo_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cQuitarBloqueo_Click", Erl)
    Resume Next
    
End Sub

Private Sub cQuitarEnEstaCapa_Click()
    
    On Error GoTo cQuitarEnEstaCapa_Click_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If cQuitarEnEstaCapa.value = True Then
        lListado(0).Enabled = False
        cFiltro(0).Enabled = False
        cGrh.Enabled = False
        cSeleccionarSuperficie.Enabled = False
        cQuitarEnTodasLasCapas.Enabled = False
        Call modPaneles.EstSelectPanel(0, True)
    Else
        lListado(0).Enabled = True
        cFiltro(0).Enabled = True
        cGrh.Enabled = True
        cSeleccionarSuperficie.Enabled = True
        cQuitarEnTodasLasCapas.Enabled = True
        Call modPaneles.EstSelectPanel(0, False)

    End If

    
    Exit Sub

cQuitarEnEstaCapa_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cQuitarEnEstaCapa_Click", Erl)
    Resume Next
    
End Sub

Private Sub cQuitarEnTodasLasCapas_Click()
    
    On Error GoTo cQuitarEnTodasLasCapas_Click_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If cQuitarEnTodasLasCapas.value = True Then
        cCapas.Enabled = False
        lListado(0).Enabled = False
        cFiltro(0).Enabled = False
        cGrh.Enabled = False
        cSeleccionarSuperficie.Enabled = False
        cQuitarEnEstaCapa.Enabled = False
        Call modPaneles.EstSelectPanel(0, True)
    Else
        cCapas.Enabled = True
        lListado(0).Enabled = True
        cFiltro(0).Enabled = True
        cGrh.Enabled = True
        cSeleccionarSuperficie.Enabled = True
        cQuitarEnEstaCapa.Enabled = True
        Call modPaneles.EstSelectPanel(0, False)

    End If

    
    Exit Sub

cQuitarEnTodasLasCapas_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cQuitarEnTodasLasCapas_Click", Erl)
    Resume Next
    
End Sub

Private Sub cQuitarFunc_Click(Index As Integer)
    
    On Error GoTo cQuitarFunc_Click_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If cQuitarFunc(Index).value = True Then
        cInsertarFunc(Index).Enabled = False
        cAgregarFuncalAzar(Index).Enabled = False
        cCantFunc(Index).Enabled = False
        cNumFunc(Index).Enabled = False
        cFiltro((Index) + 1).Enabled = False
        lListado((Index) + 1).Enabled = False
        Call modPaneles.EstSelectPanel((Index) + 3, True)
    Else
        cInsertarFunc(Index).Enabled = True
        cAgregarFuncalAzar(Index).Enabled = True
        cCantFunc(Index).Enabled = True
        cNumFunc(Index).Enabled = True
        cFiltro((Index) + 1).Enabled = True
        lListado((Index) + 1).Enabled = True
        Call modPaneles.EstSelectPanel((Index) + 3, False)

    End If

    
    Exit Sub

cQuitarFunc_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cQuitarFunc_Click", Erl)
    Resume Next
    
End Sub

Private Sub cQuitarTrans_Click()
    
    On Error GoTo cQuitarTrans_Click_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If cQuitarTrans.value = True Then
        cInsertarTransOBJ.Enabled = False
        cInsertarTrans.Enabled = False
        cUnionManual.Enabled = False
        cUnionAuto.Enabled = False
        tTMapa.Enabled = False
        tTX.Enabled = False
        tTY.Enabled = False
        Call modPaneles.EstSelectPanel(1, True)
    Else
        tTMapa.Enabled = True
        tTX.Enabled = True
        tTY.Enabled = True
        cUnionAuto.Enabled = True
        cUnionManual.Enabled = True
        cInsertarTrans.Enabled = True
        cInsertarTransOBJ.Enabled = True
        Call modPaneles.EstSelectPanel(1, False)

    End If

    
    Exit Sub

cQuitarTrans_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cQuitarTrans_Click", Erl)
    Resume Next
    
End Sub

Private Sub cQuitarTrigger_Click()
    
    On Error GoTo cQuitarTrigger_Click_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If cQuitarTrigger.value = True Then
        lListado(4).Enabled = False
        cInsertarTrigger.Enabled = False
        Call modPaneles.EstSelectPanel(6, True)
    Else
        lListado(4).Enabled = True
        cInsertarTrigger.Enabled = True
        Call modPaneles.EstSelectPanel(6, False)

    End If

    
    Exit Sub

cQuitarTrigger_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cQuitarTrigger_Click", Erl)
    Resume Next
    
End Sub

Public Sub cSeleccionarSuperficie_Click()
    
    On Error GoTo cSeleccionarSuperficie_Click_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If cSeleccionarSuperficie.value = True Then
        cQuitarEnTodasLasCapas.Enabled = False
        cQuitarEnEstaCapa.Enabled = False
        Call modPaneles.EstSelectPanel(0, True)
    Else
        cQuitarEnTodasLasCapas.Enabled = True
        cQuitarEnEstaCapa.Enabled = True
        Call modPaneles.EstSelectPanel(0, False)

    End If

    
    Exit Sub

cSeleccionarSuperficie_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cSeleccionarSuperficie_Click", Erl)
    Resume Next
    
End Sub


Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Me.Caption = "WorldEditor Argentum United"
    Call LoadMapList
    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Form_Load", Erl)
    Resume Next
    
End Sub


Sub LoadMapList()
On Error Resume Next
lstMaps.Clear
Dim sFileName As String
sFileName = Dir(App.Path & "\..\Resources\Mapas\")
Dim fh As Integer
Dim MH As tMapHeader
Dim MS As tMapSize
Dim MD As tMapDat

Do While sFileName > ""


    Dim ind As Integer
    Dim Num As Integer

    fh = FreeFile
    
    
    Open App.Path & "\..\Resources\Mapas\" & sFileName For Binary As fh

        Get #fh, , MH
        Get #fh, , MS
        Get #fh, , MD
        
       
        ind = InStr(sFileName, ".csm")
        If ind > 0 Then
            Num = mid$(sFileName, 5, Len(sFileName) - 8)
        
            lstMaps.AddItem Format(Num, "00") & " - " & MD.map_name
        End If
        
    Close fh



  sFileName = Dir()

Loop

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'If Seleccionando Then CopiarSeleccion
End Sub

Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

    Rem Estado Climatico
End Sub

Private Sub insertarLuz_Click()
    
    On Error GoTo insertarLuz_Click_Err
    

    If insertarLuz.value = True Then
        QuitarLuz.Enabled = False
    Else
        QuitarLuz.Enabled = True

    End If

    
    Exit Sub

insertarLuz_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.insertarLuz_Click", Erl)
    Resume Next
    
End Sub

Private Sub insertarParticula_Click()
    
    On Error GoTo insertarParticula_Click_Err
        
    If insertarParticula.value = True Then
        quitarparticula.Enabled = False
    Else
        quitarparticula.Enabled = True

    End If

    
    Exit Sub

insertarParticula_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.insertarParticula_Click", Erl)
    Resume Next
    
End Sub

Private Sub insnpcrandom_Click()
    
    On Error GoTo insnpcrandom_Click_Err
    
    Dim Cantidad As Byte
    Cantidad = InputBox("Ingrese la cantidad de npcs ingresamos")

    
    Exit Sub

insnpcrandom_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.insnpcrandom_Click", Erl)
    Resume Next
    
End Sub

Private Sub lListado_Click(Index As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 29/05/06
    '*************************************************
    
    On Error GoTo lListado_Click_Err
    

    If HotKeysAllow = False Then
        lListado(Index).Tag = lListado(Index).ListIndex

        Select Case Index
    
            Case 0
                cGrh.Text = DameGrhIndex(ReadField(2, lListado(Index).Text, Asc("#")))

                If SupData(ReadField(2, lListado(Index).Text, Asc("#"))).Capa <> 0 Then
                    If LenB(ReadField(2, lListado(Index).Text, Asc("#"))) = 0 Then cCapas.Tag = cCapas.Text
                    cCapas.Text = SupData(ReadField(2, lListado(Index).Text, Asc("#"))).Capa
                Else

                    If LenB(cCapas.Tag) <> 0 Then
                        cCapas.Text = cCapas.Tag
                        cCapas.Tag = vbNullString

                    End If

                End If

                'If SupData(ReadField(2, lListado(index).Text, Asc("#"))).Block = True Then
                '   If LenB(cInsertarBloqueo.Tag) = 0 Then cInsertarBloqueo.Tag = IIf(cInsertarBloqueo.value = True, 1, 0)
                '    cInsertarBloqueo.value = True
                '   Call cInsertarBloqueo_Click
                ' Else
                '    If LenB(cInsertarBloqueo.Tag) <> 0 Then
                '        cInsertarBloqueo.value = IIf(Val(cInsertarBloqueo.Tag) = 1, True, False)
                '       cInsertarBloqueo.Tag = vbNullString
                '       Call cInsertarBloqueo_Click
                '   End If
                'End If
                Call fPreviewGrh(cGrh.Text)

            Case 1
                cNumFunc(0).Text = ReadField(2, lListado(Index).Text, Asc("#"))
                Call Grh_Render_To_Hdc(picture1, BodyData(NpcData(cNumFunc(0).Text).Body).Walk(3).grhindex, 0, 0, False)

            Case 2
                cNumFunc(1).Text = ReadField(2, lListado(Index).Text, Asc("#"))

            Case 3
                cNumFunc(2).Text = ReadField(2, lListado(Index).Text, Asc("#"))
                Call Grh_Render_To_Hdc(picture1, ObjData(cNumFunc(2).Text).grhindex, 0, 0, False)

            Case 4
                TriggerBox = FrmMain.lListado(4).ListIndex

        End Select

    Else

        Rem lListado(index).ListIndex = lListado(index).Tag
    End If

    Call modPaneles.VistaPreviaDeSup

    
    Exit Sub

lListado_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.lListado_Click", Erl)
    Resume Next
    
End Sub

Private Sub lListado_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo lListado_MouseDown_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 29/05/06
    '*************************************************
    If Index = 3 And Button = 2 Then
        If lListado(3).ListIndex > -1 Then Me.PopupMenu mnuObjSc

    End If

    
    Exit Sub

lListado_MouseDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.lListado_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub lListado_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 22/05/06
    '*************************************************
    
    On Error GoTo lListado_MouseMove_Err
    

    HotKeysAllow = False

    
    Exit Sub

lListado_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.lListado_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub LuzColor_Click()
    
    On Error GoTo LuzColor_Click_Err
    
    ColorLuz.Text = Selected_Color()

    
    Exit Sub

LuzColor_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.LuzColor_Click", Erl)
    Resume Next
    
End Sub

Private Sub MiniMap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo MiniMap_MouseDown_Err
    
    UserPos.X = MapSize.Width * X / MiniMap.Width
    UserPos.Y = MapSize.Height * Y / MiniMap.Height
    
    If RectanguloModo = 2 And Button = 2 Then
    
        If RectanguloX = 0 Then
            RectanguloX = UserPos.X
            RectanguloY = UserPos.Y
            RectanguloX2 = UserPos.X
            RectanguloY2 = UserPos.Y
        Else
            'Termina el rectangulo
            If RectanguloModo = 2 Then
                'Es un spawn
                RectanguloX2 = UserPos.X
                RectanguloY2 = UserPos.Y
                frmSpawnInfo.OpenSpawn (0)
                frmSpawnInfo.Show
            End If
            RectanguloModo = 0
            RectanguloX = 0
            RectanguloY = 0
        End If
    End If
    
    
    bRefreshRadar = True

    
    Exit Sub

MiniMap_MouseDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.MiniMap_MouseDown", Erl)
    Resume Next
    
End Sub


Private Sub mnuAbrirMapa_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Dialog.CancelError = True

    On Error GoTo ErrHandler

    DeseaGuardarMapa Dialog.Filename

    ObtenerNombreArchivo False

    If Len(Dialog.Filename) < 3 Then Exit Sub

    If WalkMode = True Then
        Call modGeneral.ToggleWalkMode

    End If

    modMapIO.AbrirMapa Dialog.Filename
    DoEvents
    mnuReAbrirMapa.Enabled = True
    EngineRun = True
    
    Exit Sub
ErrHandler:

End Sub

Private Sub mnuacercade_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuacercade_Click_Err
    
    frmAbout.Show

    
    Exit Sub

mnuacercade_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuacercade_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuAutoCapturarTranslados_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 28/05/06
    '*************************************************
    
    On Error GoTo mnuAutoCapturarTranslados_Click_Err
    
    mnuAutoCapturarTranslados.Checked = (mnuAutoCapturarTranslados.Checked = False)

    
    Exit Sub

mnuAutoCapturarTranslados_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuAutoCapturarTranslados_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuAutoCapturarSuperficie_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 28/05/06
    '*************************************************
    
    On Error GoTo mnuAutoCapturarSuperficie_Click_Err
    
    mnuAutoCapturarSuperficie.Checked = (mnuAutoCapturarSuperficie.Checked = False)

    
    Exit Sub

mnuAutoCapturarSuperficie_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuAutoCapturarSuperficie_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuAutoCompletarSuperficies_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuAutoCompletarSuperficies_Click_Err
    
    mnuAutoCompletarSuperficies.Checked = (mnuAutoCompletarSuperficies.Checked = False)

    If mnuAutoCompletarSuperficies.Checked = False Then
        FrmMain.LvBOpcion(18).Caption = "Grh Normal"
    Else
        FrmMain.LvBOpcion(18).Caption = "AutoCompletar"

    End If

    
    Exit Sub

mnuAutoCompletarSuperficies_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuAutoCompletarSuperficies_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuAutoGuardarMapas_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuAutoGuardarMapas_Click_Err
    
    frmAutoGuardarMapa.Show

    
    Exit Sub

mnuAutoGuardarMapas_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuAutoGuardarMapas_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuAutoQuitarFunciones_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuAutoQuitarFunciones_Click_Err
    
    mnuAutoQuitarFunciones.Checked = (mnuAutoQuitarFunciones.Checked = False)

    
    Exit Sub

mnuAutoQuitarFunciones_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuAutoQuitarFunciones_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuBloquear_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuBloquear_Click_Err
    
    Dim i As Byte

    For i = 0 To 6

        If i <> 2 Then
            FrmMain.SelectPanel(i).value = False
            Call VerFuncion(i, False)

        End If

    Next

    modPaneles.VerFuncion 2, True

    
    Exit Sub

mnuBloquear_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuBloquear_Click", Erl)
    Resume Next
    
End Sub



Private Sub mnuBloquearS_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    
    On Error GoTo mnuBloquearS_Click_Err
    
    Call BlockearSeleccion

    
    Exit Sub

mnuBloquearS_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuBloquearS_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuConfigAvanzada_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuConfigAvanzada_Click_Err
    
    frmConfigSup.Show

    
    Exit Sub

mnuConfigAvanzada_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuConfigAvanzada_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuConfigObjTrans_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 29/05/06
    '*************************************************
    
    On Error GoTo mnuConfigObjTrans_Click_Err
    
    Cfg_TrOBJ = cNumFunc(2).Text

    
    Exit Sub

mnuConfigObjTrans_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuConfigObjTrans_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuCopiar_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    
    On Error GoTo mnuCopiar_Click_Err
    
    Call CopiarSeleccion

    
    Exit Sub

mnuCopiar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuCopiar_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuCortar_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    
    On Error GoTo mnuCortar_Click_Err
    
    Call CortarSeleccion

    
    Exit Sub

mnuCortar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuCortar_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuDeshacer_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 15/10/06
    '*************************************************
    
    On Error GoTo mnuDeshacer_Click_Err
    
    Call modEdicion.Deshacer_Recover

    
    Exit Sub

mnuDeshacer_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuDeshacer_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuDeshacerPegado_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    
    On Error GoTo mnuDeshacerPegado_Click_Err
    
    Call DePegar

    
    Exit Sub

mnuDeshacerPegado_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuDeshacerPegado_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuGRHaBMP_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    
    On Error GoTo mnuGRHaBMP_Click_Err
    
    frmGRHaBMP.Show

    
    Exit Sub

mnuGRHaBMP_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuGRHaBMP_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuGuardarcomoBMP_Click()
    '*************************************************
    'Author: Salvito
    'Last modified: 01/05/2008 - ^[GS]^
    '*************************************************
    
    On Error GoTo mnuGuardarcomoBMP_Click_Err
    
    Dim Ratio As Integer
    Ratio = CInt(Val(InputBox("En que escala queres Renderizar? Entre 1 y 20.", "Elegi Escala", "1")))

    If Ratio < 1 Then Ratio = 1
    If Ratio >= 1 And Ratio <= 20 Then

    End If

    
    Exit Sub

mnuGuardarcomoBMP_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuGuardarcomoBMP_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuGuardarcomoJPG_Click()
    '*************************************************
    'Author: Salvito
    'Last modified: 01/05/2008 - ^[GS]^
    '*************************************************
    
    On Error GoTo mnuGuardarcomoJPG_Click_Err
    
    Dim Ratio As Integer
    Ratio = CInt(Val(InputBox("En que escala queres Renderizar? Entre 1 y 20.", "Elegi Escala", "1")))

    If Ratio < 1 Then Ratio = 1
    If Ratio >= 1 And Ratio <= 20 Then
  
    End If

    
    Exit Sub

mnuGuardarcomoJPG_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuGuardarcomoJPG_Click", Erl)
    Resume Next
    
End Sub

Public Sub mnuGuardarMapa_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuGuardarMapa_Click_Err
    
    If MsgBox("¿Desea guardar los cambios del mapa?", vbYesNo + vbInformation) = vbYes Then
        Call Save_Map_Data(App.Path & "\..\Resources\Mapas\Mapa" & UserMap & ".csm")
  
        MapInfo.Changed = 0
        
        If UserMap > Val(mid$(lstMaps.List(lstMaps.ListCount - 1), 1, InStr(1, lstMaps.List(lstMaps.ListCount - 1), " ") - 1)) Then
            lstMaps.AddItem UserMap & " - " & MapInfo.name
        End If
    End If
    

    
    Exit Sub

mnuGuardarMapa_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuGuardarMapa_Click", Erl)
    Resume Next
    
End Sub


Private Sub mnuGuardarUltimaConfig_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 23/05/06
    '*************************************************
    Rem mnuGuardarUltimaConfig.Checked = (mnuGuardarUltimaConfig.Checked = False)
End Sub

Private Sub mnuInformes_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuInformes_Click_Err
    
    frmInformes.Show

    
    Exit Sub

mnuInformes_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuInformes_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuManual_Click()
    
    On Error GoTo mnuManual_Click_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 24/11/08
    '*************************************************
    If LenB(Dir(App.Path & "\manual\index.html", vbArchive)) <> 0 Then
        Call Shell("explorer " & App.Path & "\manual\index.html")
        DoEvents

    End If

    
    Exit Sub

mnuManual_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuManual_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuModoCaminata_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 28/05/06
    '*************************************************
    
    On Error GoTo mnuModoCaminata_Click_Err
    
    ToggleWalkMode

    
    Exit Sub

mnuModoCaminata_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuModoCaminata_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuNPCs_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuNPCs_Click_Err
    
    Dim i As Byte

    For i = 0 To 6

        If i <> 3 Then
            FrmMain.SelectPanel(i).value = False
            Call VerFuncion(i, False)

        End If

    Next
    modPaneles.VerFuncion 3, True

    
    Exit Sub

mnuNPCs_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuNPCs_Click", Erl)
    Resume Next
    
End Sub

'Private Sub mnuNPCsHostiles_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
'Dim i As Byte
'For i = 0 To 6
'    If i <> 4 Then
'        frmMain.SelectPanel(i).value = False
'        Call VerFuncion(i, False)
'    End If
'Next
'modPaneles.VerFuncion 4, True
'End Sub

Private Sub mnuNuevoMapa_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuNuevoMapa_Click_Err
    

    Dim loopc As Integer

    DeseaGuardarMapa Dialog.Filename

    FrmMain.Dialog.Filename = Empty


    UserMap = Val(mid$(lstMaps.List(lstMaps.ListCount - 1), 1, InStr(1, lstMaps.List(lstMaps.ListCount - 1), " ") - 1)) + 1
    FrmMain.Label16.Caption = "Map " & UserMap
    Call modMapIO.NuevoMapa
    
    Exit Sub

mnuNuevoMapa_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuNuevoMapa_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuObjetos_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuObjetos_Click_Err
    
    Dim i As Byte

    For i = 0 To 6

        If i <> 5 Then
            FrmMain.SelectPanel(i).value = False
            Call VerFuncion(i, False)

        End If

    Next
    modPaneles.VerFuncion 5, True

    
    Exit Sub

mnuObjetos_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuObjetos_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuOptimizar_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 22/09/06
    '*************************************************
    
    On Error GoTo mnuOptimizar_Click_Err
    
    frmOptimizar.Show

    
    Exit Sub

mnuOptimizar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuOptimizar_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuPegar_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    
    On Error GoTo mnuPegar_Click_Err
    
    
    ModoPegar = True

    
    Exit Sub

mnuPegar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuPegar_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuQBloquear_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuQBloquear_Click_Err
    
    modPaneles.VerFuncion 2, False

    
    Exit Sub

mnuQBloquear_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuQBloquear_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuQNPCs_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuQNPCs_Click_Err
    
    modPaneles.VerFuncion 3, False

    
    Exit Sub

mnuQNPCs_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuQNPCs_Click", Erl)
    Resume Next
    
End Sub

'Private Sub mnuQNPCsHostiles_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
'modPaneles.VerFuncion 4, False
'End Sub

Private Sub mnuQObjetos_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuQObjetos_Click_Err
    
    modPaneles.VerFuncion 5, False

    
    Exit Sub

mnuQObjetos_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuQObjetos_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuQSuperficie_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuQSuperficie_Click_Err
    
    modPaneles.VerFuncion 0, False

    
    Exit Sub

mnuQSuperficie_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuQSuperficie_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuQTranslados_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuQTranslados_Click_Err
    
    modPaneles.VerFuncion 1, False

    
    Exit Sub

mnuQTranslados_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuQTranslados_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuQTriggers_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuQTriggers_Click_Err
    
    modPaneles.VerFuncion 6, False

    
    Exit Sub

mnuQTriggers_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuQTriggers_Click", Erl)
    Resume Next
    
End Sub


Private Sub mnuQuitarFunciones_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuQuitarFunciones_Click_Err
    

    ' Superficies
    cSeleccionarSuperficie.value = False
    Call cSeleccionarSuperficie_Click
    cQuitarEnEstaCapa.value = False
    Call cQuitarEnEstaCapa_Click
    cQuitarEnTodasLasCapas.value = False
    Call cQuitarEnTodasLasCapas_Click

    ' Translados
    cQuitarTrans.value = False
    Call cQuitarTrans_Click
    cInsertarTrans.value = False
    Call cInsertarTrans_Click

    ' Bloqueos
    cQuitarBloqueo.value = False
    Call cQuitarBloqueo_Click
    cInsertarBloqueo.value = False
    Call cInsertarBloqueo_Click

    ' Otras funciones
    cInsertarFunc(0).value = False
    Call cInsertarFunc_Click(0)
    cInsertarFunc(1).value = False
    Call cInsertarFunc_Click(1)
    cInsertarFunc(2).value = False
    Call cInsertarFunc_Click(2)
    cQuitarFunc(0).value = False
    Call cQuitarFunc_Click(0)
    cQuitarFunc(1).value = False
    Call cQuitarFunc_Click(1)
    cQuitarFunc(2).value = False
    Call cQuitarFunc_Click(2)

    ' Triggers
    cInsertarTrigger.value = False
    Call cInsertarTrigger_Click
    cQuitarTrigger.value = False
    Call cQuitarTrigger_Click

    ' particulas
    insertarParticula.value = False
    Call insertarParticula_Click
    quitarparticula.value = False
    Call quitarparticula_Click

    ' Luces
    insertarLuz.value = False
    Call insertarLuz_Click
    QuitarLuz.value = False
    Call QuitarLuz_Click

    
    Exit Sub

mnuQuitarFunciones_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuQuitarFunciones_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuQuitarNPCs_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuQuitarNPCs_Click_Err
    
    Call modEdicion.Quitar_NPCs(False)

    
    Exit Sub

mnuQuitarNPCs_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuQuitarNPCs_Click", Erl)
    Resume Next
    
End Sub

'Private Sub mnuQuitarNPCsHostiles_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
'Call modEdicion.Quitar_NPCs(True)
'End Sub

Private Sub mnuQuitarObjetos_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuQuitarObjetos_Click_Err
    
    Call modEdicion.Quitar_Objetos

    
    Exit Sub

mnuQuitarObjetos_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuQuitarObjetos_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuQuitarTODO_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuQuitarTODO_Click_Err
    
    Call modEdicion.Borrar_Mapa

    
    Exit Sub

mnuQuitarTODO_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuQuitarTODO_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuQuitarTranslados_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 16/10/06
    '*************************************************
    
    On Error GoTo mnuQuitarTranslados_Click_Err
    
    Call modEdicion.Quitar_Translados

    
    Exit Sub

mnuQuitarTranslados_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuQuitarTranslados_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuQuitarTriggers_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuQuitarTriggers_Click_Err
    
    Call modEdicion.Quitar_Triggers

    
    Exit Sub

mnuQuitarTriggers_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuQuitarTriggers_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuReAbrirMapa_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    On Error GoTo ErrHandler

    If FileExist(Dialog.Filename, vbArchive) = False Then Exit Sub
    If MapInfo.Changed = 1 Then
        If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then
            'modMapIO.GuardarMapa Dialog.FileName
            'Call modMapIO.GuardarMapa(PATH_Save & MapName)
            mnuGuardarMapa_Click
        End If

    End If

    modMapIO.AbrirMapa Dialog.Filename
    DoEvents
    mnuReAbrirMapa.Enabled = True
    EngineRun = True
    Exit Sub
ErrHandler:

End Sub

Private Sub mnuRealizarOperacion_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    
    On Error GoTo mnuRealizarOperacion_Click_Err
    

    
    mnuAutoCompletarSuperficies.Checked = False

    Call AccionSeleccion

    
    Exit Sub

mnuRealizarOperacion_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuRealizarOperacion_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuSalir_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuSalir_Click_Err
    
    Unload Me

    
    Exit Sub

mnuSalir_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuSalir_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuSuperficie_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuSuperficie_Click_Err
    
    Dim i As Byte

    For i = 0 To 6

        If i <> 0 Then
            FrmMain.SelectPanel(i).value = False
            Call VerFuncion(i, False)

        End If

    Next
    modPaneles.VerFuncion 0, True

    
    Exit Sub

mnuSuperficie_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuSuperficie_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuTranslados_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuTranslados_Click_Err
    
    Dim i As Byte

    For i = 0 To 6

        If i <> 1 Then
            FrmMain.SelectPanel(i).value = False
            Call VerFuncion(i, False)

        End If

    Next
    modPaneles.VerFuncion 1, True

    
    Exit Sub

mnuTranslados_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuTranslados_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuTriggers_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuTriggers_Click_Err
    
    Dim i As Byte

    For i = 0 To 6

        If i <> 6 Then
            FrmMain.SelectPanel(i).value = False
            Call VerFuncion(i, False)

        End If

    Next
    modPaneles.VerFuncion 6, True

    
    Exit Sub

mnuTriggers_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuTriggers_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuUtilizarDeshacer_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 16/10/06
    '*************************************************
    
    On Error GoTo mnuUtilizarDeshacer_Click_Err
    
    mnuUtilizarDeshacer.Checked = (mnuUtilizarDeshacer.Checked = False)

    
    Exit Sub

mnuUtilizarDeshacer_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuUtilizarDeshacer_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuVerAutomatico_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuVerAutomatico_Click_Err
    
    mnuVerAutomatico.Checked = (mnuVerAutomatico.Checked = False)

    
    Exit Sub

mnuVerAutomatico_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuVerAutomatico_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuVerBloqueos_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuVerBloqueos_Click_Err
    
    cVerBloqueos.value = (cVerBloqueos.value = False)
    mnuVerBloqueos.Checked = cVerBloqueos.value

    
    Exit Sub

mnuVerBloqueos_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuVerBloqueos_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuVerCapa1_Click()
    
    On Error GoTo mnuVerCapa1_Click_Err
    
    mnuVerCapa1.Checked = (mnuVerCapa1.Checked = False)

    
    Exit Sub

mnuVerCapa1_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuVerCapa1_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuVerCapa2_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuVerCapa2_Click_Err
    
    mnuVerCapa2.Checked = (mnuVerCapa2.Checked = False)

    
    Exit Sub

mnuVerCapa2_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuVerCapa2_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuVerCapa3_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuVerCapa3_Click_Err
    
    mnuVerCapa3.Checked = (mnuVerCapa3.Checked = False)

    
    Exit Sub

mnuVerCapa3_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuVerCapa3_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuVerCapa4_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuVerCapa4_Click_Err
    
    mnuVerCapa4.Checked = (mnuVerCapa4.Checked = False)

    
    Exit Sub

mnuVerCapa4_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuVerCapa4_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuVerGrilla_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 24/11/08
    '*************************************************
    
    On Error GoTo mnuVerGrilla_Click_Err
   
    VerGrilla = (VerGrilla = False)
    mnuVerGrilla.Checked = VerGrilla


    
    Exit Sub

mnuVerGrilla_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuVerGrilla_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuVerLuces_Click()
    
    On Error GoTo mnuVerLuces_Click_Err
    
    mnuVerLuces.Checked = (mnuVerLuces.Checked = False)

    
    Exit Sub

mnuVerLuces_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuVerLuces_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuVerNPCs_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 26/05/06
    '*************************************************
    
    On Error GoTo mnuVerNPCs_Click_Err
    
    mnuVerNPCs.Checked = (mnuVerNPCs.Checked = False)

    
    Exit Sub

mnuVerNPCs_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuVerNPCs_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuVerObjetos_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 26/05/06
    '*************************************************
    
    On Error GoTo mnuVerObjetos_Click_Err
    
    mnuVerObjetos.Checked = (mnuVerObjetos.Checked = False)

    
    Exit Sub

mnuVerObjetos_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuVerObjetos_Click", Erl)
    Resume Next
    
End Sub

Public Sub mnuVerParticulas_Click()
    
    On Error GoTo mnuVerParticulas_Click_Err
    

    mnuVerParticulas.Checked = (mnuVerParticulas.Checked = False)

    
    Exit Sub

mnuVerParticulas_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuVerParticulas_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuVerTranslados_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 26/05/06
    '*************************************************
    
    On Error GoTo mnuVerTranslados_Click_Err
    
    mnuVerTranslados.Checked = (mnuVerTranslados.Checked = False)

    
    Exit Sub

mnuVerTranslados_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuVerTranslados_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuVerTriggers_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuVerTriggers_Click_Err
    
    cVerTriggers.value = (cVerTriggers.value = False)
    mnuVerTriggers.Checked = cVerTriggers.value

    
    Exit Sub

mnuVerTriggers_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuVerTriggers_Click", Erl)
    Resume Next
    
End Sub

Private Sub picRadar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo picRadar_MouseDown_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 29/05/06
    '*************************************************
    If X < 11 Then X = 11
    If X > 89 Then X = 89
    If Y < 10 Then Y = 10
    If Y > 92 Then Y = 92
    UserPos.X = X
    UserPos.Y = Y
    bRefreshRadar = True

    
    Exit Sub

picRadar_MouseDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.picRadar_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub picRadar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 28/05/06
    '*************************************************
    
    On Error GoTo picRadar_MouseMove_Err
    
    MiRadarX = X
    MiRadarY = Y

    
    Exit Sub

picRadar_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.picRadar_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 24/11/08
    '*************************************************
    
    On Error GoTo Form_QueryUnload_Err
    

    ' Guardar configuración
    Rem WriteVar IniPath & "WorldEditor.ini", "CONFIGURACION", "GuardarConfig", IIf(FrmMain.mnuGuardarUltimaConfig.Checked = True, "1", "0")

    WriteVar IniPath & "WorldEditor.ini", "PATH", "UltimoMapa", Dialog.Filename
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "ControlAutomatico", IIf(FrmMain.mnuVerAutomatico.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "Capa2", IIf(FrmMain.mnuVerCapa2.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "Capa3", IIf(FrmMain.mnuVerCapa3.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "Capa4", IIf(FrmMain.mnuVerCapa4.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "Translados", IIf(FrmMain.mnuVerTranslados.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "Objetos", IIf(FrmMain.mnuVerObjetos.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "NPCs", IIf(FrmMain.mnuVerNPCs.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "Triggers", IIf(FrmMain.mnuVerTriggers.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "Marco", IIf(FrmMain.mnuVerMarco.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "Grilla", IIf(FrmMain.mnuVerGrilla.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "Bloqueos", IIf(FrmMain.mnuVerBloqueos.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "LastPos", UserPos.X & "-" & UserPos.Y
    WriteVar IniPath & "WorldEditor.ini", "CONFIGURACION", "UtilizarDeshacer", IIf(FrmMain.mnuUtilizarDeshacer.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "CONFIGURACION", "AutoCapturarTrans", IIf(FrmMain.mnuAutoCapturarTranslados.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "CONFIGURACION", "AutoCapturarSup", IIf(FrmMain.mnuAutoCapturarSuperficie.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "CONFIGURACION", "ObjTranslado", Val(Cfg_TrOBJ)

    'Allow MainLoop to close program
    If prgRun = True Then
        prgRun = False
        Cancel = 1

    End If

    
    Exit Sub

Form_QueryUnload_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Form_QueryUnload", Erl)
    Resume Next
    
End Sub
Private Sub objalazar_Click()
    
    On Error GoTo objalazar_Click_Err
    

    Dim Cantidad As Long
    Dim bloquear As Byte
    Dim objeto   As Long
    Dim X        As Integer
    Dim Y        As Integer
    Dim i        As Long

    Cantidad = InputBox("Ingrese la cantidad de objetos a mapear")

    If Cantidad <= 0 Then Exit Sub
    bloquear = InputBox("¿Desea bloquear los obejtos? (1= SI | 0 = NO")

    If bloquear > 1 Then Exit Sub
    objeto = FrmMain.cNumFunc(2).Text

    For i = 1 To Cantidad
        X = RandomNumber(10, 91)
        Y = RandomNumber(8, 93)

        If MapData(X, Y).Graphic(1).grhindex < 1505 Or MapData(X, Y).Graphic(1).grhindex > 1520 Then
            
            MapInfo.Changed = 1 'Set changed flag
                
            MapData(X, Y).Blocked = bloquear * &HF
        
            InitGrh MapData(X, Y).ObjGrh, ObjData(objeto).grhindex
            MapData(X, Y).OBJInfo.ObjIndex = objeto
            MapData(X, Y).OBJInfo.Amount = 1

        End If
            
    Next i

    Call AddtoRichTextBox(FrmMain.RichTextBox1, "Se agregaron " & Cantidad & " " & ObjData(objeto).name & " al mapa.", 255, 255, 255, False, True, False)

    
    Exit Sub

objalazar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.objalazar_Click", Erl)
    Resume Next
    
End Sub

Private Sub Objeto_Click()
    
    On Error GoTo Objeto_Click_Err
    
    Dim Y As Integer
    Dim X As Integer

    For Y = 1 To MapSize.Height
        For X = 1 To MapSize.Width
            'If MapData(X, Y).OBJInfo.objindex = Text1 Then
            '         InitGrh MapData(X, Y).ObjGrh, 1
            '        MapData(X, Y).OBJInfo.objindex = Text2
            '         MapData(X, Y).OBJInfo.Amount = 1
            ' End If
        Next X
    Next Y

    
    Exit Sub

Objeto_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Objeto_Click", Erl)
    Resume Next
    
End Sub

Private Sub pasto_Click()
    
    On Error GoTo pasto_Click_Err
    
    cGrh.Text = DameGrhIndex(0)

    Call modPaneles.VistaPreviaDeSup

    
    Exit Sub

pasto_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.pasto_Click", Erl)
    Resume Next
    
End Sub

Private Sub QuitarLuz_Click()
    
    On Error GoTo QuitarLuz_Click_Err
    

    If QuitarLuz.value = True Then
        insertarLuz.Enabled = False
    Else
        insertarLuz.Enabled = True

    End If

    
    Exit Sub

QuitarLuz_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.QuitarLuz_Click", Erl)
    Resume Next
    
End Sub

Private Sub quitarparticula_Click()
    
    On Error GoTo quitarparticula_Click_Err
    

    If quitarparticula.value = True Then
        insertarParticula.Enabled = False
    Else
        insertarParticula.Enabled = True

    End If

    
    Exit Sub

quitarparticula_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.quitarparticula_Click", Erl)
    Resume Next
    
End Sub


Private Sub renderer_Click()
    
    On Error GoTo renderer_Click_Err
    

    Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
    UltimoClickX = tX
    UltimoClickY = tY

    If DesdeBloq = True Then
        RepetirSup = False
        modEdicion.Deshacer_Add UltimoClickX, UltimoClickY, 1, 1
        DesdeBloq = False
        Call PonerGrh
        
        If RepetirSup Then
            Call InsertarBloque

        End If

    End If
    
    

    
    Exit Sub

renderer_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.renderer_Click", Erl)
    Resume Next
    
End Sub

Private Sub renderer_DblClick()
    
    On Error GoTo renderer_DblClick_Err
    
    Dim tX As Integer
    Dim tY As Integer

    If Not MapaCargado Then Exit Sub

    If SobreX > 0 And SobreY > 0 Then
        DobleClick Val(SobreX), Val(SobreY)
    End If

    
    Exit Sub

renderer_DblClick_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.renderer_DblClick", Erl)
    Resume Next
    
End Sub

Private Sub renderer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo renderer_MouseDown_Err
    

    If Not MapaCargado Then Exit Sub

    Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
    
    If ModoPegar Then
        PegarSeleccion
        ModoPegar = False
        Exit Sub
    End If
    
    
    Dim i As Integer
    SelectedZona = 0
    For i = 1 To NumZonas
        With Zona(i)
            If .Deleted = 0 And .Map = UserMap And tX >= .X And tX <= .X2 And tY >= .Y And tY <= .Y2 Then
                SelectedZona = i
                Exit For
            End If
        End With
    Next i
    mnuEditarZona.Enabled = SelectedZona > 0
    mnuEliminarZona.Enabled = SelectedZona > 0
    
    SelectedSpawn = 0
    For i = 1 To NumSpawns
        With NpcSpawn(i)
            If .Deleted = 0 And .Map = UserMap And tX >= .X And tX <= .X2 And tY >= .Y And tY <= .Y2 Then
                SelectedSpawn = i
                Exit For
            End If
        End With
    Next i
    mnuEditarSpawn.Enabled = SelectedSpawn > 0
    mnuEliminarSpawn.Enabled = SelectedSpawn > 0
    
    If RectanguloModo > 0 Then
    
        If RectanguloX = 0 Then
            RectanguloX = tX
            RectanguloY = tY
            RectanguloX2 = tX
            RectanguloY2 = tY
        Else
            'Termina el rectangulo
            If RectanguloModo = 1 Then
                frmZonaInfo.OpenZona (0)
                frmZonaInfo.Show
            ElseIf RectanguloModo = 2 Then
                'Es un spawn
                frmSpawnInfo.OpenSpawn (0)
                frmSpawnInfo.Show
            End If
            RectanguloModo = 0
            RectanguloX = 0
            RectanguloY = 0
        End If
    End If

    'If Shift = 1 And Button = 2 Then PegarSeleccion tX, tY: Exit Sub
    If Shift = 1 And Button = 1 Then
        Seleccionando = True
        SeleccionIX = tX '+ UserPos.X
        SeleccionIY = tY '+ UserPos.Y
        SeleccionFX = tX
        SeleccionFY = tY
    Else
        ClickEdit Button, tX, tY

    End If

    
    Exit Sub

renderer_MouseDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.renderer_MouseDown", Erl)
    Resume Next
    
End Sub

Public Sub renderer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo renderer_MouseMove_Err
    

    MouseX = X
    MouseY = Y
    MouseBoton = Button
    MouseShift = Shift

    'Make sure map is loaded
    If Not MapaCargado Then Exit Sub
    HotKeysAllow = True
    
    Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
    
    POSX.Caption = "X: " & tX & " - Y: " & tY
    
    
    If RectanguloModo > 0 Then
        If RectanguloX > 0 Then
            RectanguloX2 = tX
            RectanguloY2 = tY
        End If
    End If

    If tX < 1 Or tY < 1 Or tX > MapSize.Width Or tY > MapSize.Height Then
        POSX.ForeColor = vbRed
    Else
        POSX.ForeColor = vbWhite

    End If

    If Shift = 1 And Button = 1 Then
        Seleccionando = True
        SeleccionFX = tX '+ TileX
        SeleccionFY = tY '+ TileY
    Else

        If tX = 0 Then Exit Sub
        If tY = 0 Then Exit Sub
        If tX = LastX And tY = LastY Then Exit Sub
        
        ClickEdit Button, tX, tY
        
        LastX = tX
        LastY = tY

    End If

    
    Exit Sub

renderer_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.renderer_MouseMove", Erl)
    Resume Next
    
End Sub


Private Sub renderer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim temp As Integer
If SeleccionIX > SeleccionFX Then
    temp = SeleccionFX
    SeleccionFX = SeleccionIX
    SeleccionIX = temp
End If

If SeleccionIY > SeleccionFY Then
    temp = SeleccionFY
    SeleccionFY = SeleccionIY
    SeleccionIY = temp
End If
End Sub

Public Sub SelectPanel_Click(Index As Integer)
    
    On Error GoTo SelectPanel_Click_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Dim i As Byte

    For i = 0 To 8

        If i <> Index Then
            SelectPanel(i).value = False
            Call VerFuncion(i, False)

        End If

    Next

    If mnuAutoQuitarFunciones.Checked = True Then Call mnuQuitarFunciones_Click

    Call VerFuncion(Index, SelectPanel(Index).value)

    
    Exit Sub

SelectPanel_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.SelectPanel_Click", Erl)
    Resume Next
    
End Sub




Private Sub Text3_Change()

End Sub

Private Sub TiggerEspecial_Click()

    On Error Resume Next

    TriggerBox = InputBox("Ingrese el numero de trigger a usar.")

End Sub


Public Sub ObtenerNombreArchivo(ByVal Guardar As Boolean)

    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************
    On Error Resume Next

    With Dialog

        If FormatoIAO Then
            .Filter = "Mapas de RevolucionAO (*.csm)|*.csm"
        Else
            .Filter = "Mapas de ArgentumOnline (*.map)|*.map"

        End If

        If Guardar Then
            .DialogTitle = "Guardar"
            .DefaultExt = ".txt"
            .Filename = vbNullString
            .FLAGS = cdlOFNPathMustExist
            .ShowSave
        Else
            .DialogTitle = "Cargar"
            .Filename = vbNullString
            .FLAGS = cdlOFNFileMustExist
            .ShowOpen

        End If

    End With

End Sub


Private Sub Timer2_Timer()
    
    On Error GoTo Timer2_Timer_Err
    

    If engine.bRunning Then engine.Engine_ActFPS

    
    Exit Sub

Timer2_Timer_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Timer2_Timer", Erl)
    Resume Next
    
End Sub

Private Sub Todas_las_luces_Click()
    
    On Error GoTo Todas_las_luces_Click_Err
    
    Dim X As Integer
    Dim Y As Integer
    Dim i As Long

    For X = 1 To MapSize.Width
        For Y = 1 To MapSize.Height

            MapData(X, Y).luz.Rango = 0
        Next Y
    Next X

    engine.Light_Remove_All
    MapInfo.Changed = 1

    
    Exit Sub

Todas_las_luces_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Todas_las_luces_Click", Erl)
    Resume Next
    
End Sub

Private Sub Todas_las_Particulas_Click()
    
    On Error GoTo Todas_las_Particulas_Click_Err
    
    Dim X As Integer
    Dim Y As Integer
    Dim i As Long

    For X = 1 To MapSize.Width
        For Y = 1 To MapSize.Height
            MapData(X, Y).particle_Index = 0
        Next Y
    Next X

    engine.Particle_Group_Remove_All
    MapInfo.Changed = 1

    
    Exit Sub

Todas_las_Particulas_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Todas_las_Particulas_Click", Erl)
    Resume Next
    
End Sub


Private Sub txtnamemapa_Change()
    
    On Error GoTo txtnamemapa_Change_Err
    
    MapDat.map_name = txtnamemapa
    Call AddtoRichTextBox(FrmMain.RichTextBox1, "Nombre de mapa cambiado a:  " & MapDat.map_name, 255, 255, 255, False, True, False)
    MapInfo.Changed = 1

    
    Exit Sub

txtnamemapa_Change_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.txtnamemapa_Change", Erl)
    Resume Next
    
End Sub


Private Sub vergraficoslistado_Click()
    
    On Error GoTo vergraficoslistado_Click_Err
    
    Form1.Show , FrmMain

    
    Exit Sub

vergraficoslistado_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.vergraficoslistado_Click", Erl)
    Resume Next
    
End Sub

Sub DibujarSpawns()
    Dim i As Integer
    Dim r As Single
    r = MiniMap.Width / MapSize.Width
    MiniMap.Cls
    MiniMap.ForeColor = vbGreen
    For i = 1 To NumSpawns
        If NpcSpawn(i).Map = UserMap And NpcSpawn(i).Deleted = 0 Then
            MiniMap.Line ((NpcSpawn(i).X - 1) * r, (NpcSpawn(i).Y - 1) * r)-((NpcSpawn(i).X2 - 1) * r, (NpcSpawn(i).Y - 1) * r)
            MiniMap.Line ((NpcSpawn(i).X - 1) * r, (NpcSpawn(i).Y2 - 1) * r)-((NpcSpawn(i).X2 - 1) * r, (NpcSpawn(i).Y2 - 1) * r)
            MiniMap.Line ((NpcSpawn(i).X - 1) * r, (NpcSpawn(i).Y - 1) * r)-((NpcSpawn(i).X - 1) * r, (NpcSpawn(i).Y2 - 1) * r)
            MiniMap.Line ((NpcSpawn(i).X2 - 1) * r, (NpcSpawn(i).Y - 1) * r)-((NpcSpawn(i).X2 - 1) * r, (NpcSpawn(i).Y2 - 1) * r)
        End If
    Next i
End Sub

Sub DibujarZonas()
    Dim i As Integer
    Dim r As Single
    r = MiniMap.Width / MapSize.Width
    MiniMap.Cls
    MiniMap.ForeColor = vbYellow
    For i = 1 To NumZonas
        If Zona(i).Map = UserMap And Zona(i).Deleted = 0 Then
            MiniMap.Line ((Zona(i).X - 1) * r, (Zona(i).Y - 1) * r)-((Zona(i).X2 - 1) * r, (Zona(i).Y - 1) * r)
            MiniMap.Line ((Zona(i).X - 1) * r, (Zona(i).Y2 - 1) * r)-((Zona(i).X2 - 1) * r, (Zona(i).Y2 - 1) * r)
            MiniMap.Line ((Zona(i).X - 1) * r, (Zona(i).Y - 1) * r)-((Zona(i).X - 1) * r, (Zona(i).Y2 - 1) * r)
            MiniMap.Line ((Zona(i).X2 - 1) * r, (Zona(i).Y - 1) * r)-((Zona(i).X2 - 1) * r, (Zona(i).Y2 - 1) * r)
        End If
    Next i
End Sub


Sub DibujarHostiles()
    Dim i As Integer
    Dim r As Single
    r = MiniMap.Width / MapSize.Width
    MiniMap.Cls
    MiniMap.ForeColor = vbWhite
    
    Dim X As Integer
    Dim Y As Integer
     MiniMap.ForeColor = vbRed
    For Y = 1 To MapSize.Height
        For X = 1 To MapSize.Width
            If MapData(X, Y).NpcIndex > 0 Then
                If NpcData(MapData(X, Y).NpcIndex).Hostile = 1 Then
                    MiniMap.PSet (((X - 1) * r), ((Y - 1) * r))
                    MiniMap.PSet (((X) * r), ((Y - 1) * r))
                    MiniMap.PSet (((X - 1) * r), ((Y) * r))
                    MiniMap.PSet (((X) * r), ((Y) * r))

                End If
            End If
        Next X
    Next Y
    
End Sub

