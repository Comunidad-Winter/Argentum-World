VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   11535
   ClientLeft      =   345
   ClientTop       =   240
   ClientWidth     =   15330
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   769
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1022
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer timerCooldownCombo 
      Interval        =   20
      Left            =   9120
      Top             =   2400
   End
   Begin VB.PictureBox panel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5355
      Left            =   11340
      ScaleHeight     =   357
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   266
      TabIndex        =   37
      Top             =   2340
      Width           =   3990
      Begin VB.PictureBox picHechiz 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         ClipControls    =   0   'False
         DrawStyle       =   3  'Dash-Dot
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3855
         Left            =   360
         MousePointer    =   99  'Custom
         ScaleHeight     =   257
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   216
         TabIndex        =   39
         Top             =   645
         Visible         =   0   'False
         Width           =   3240
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3825
         Left            =   420
         ScaleHeight     =   255
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   210
         TabIndex        =   38
         Top             =   675
         Width           =   3150
      End
      Begin VB.Label ObjLbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ItemData"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   615
         Left            =   240
         TabIndex        =   40
         Top             =   4560
         Visible         =   0   'False
         Width           =   3540
      End
      Begin VB.Image imgDeleteItem 
         Height          =   375
         Left            =   3480
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image cmdlanzar 
         Height          =   525
         Left            =   165
         Tag             =   "1"
         Top             =   4665
         Visible         =   0   'False
         Width           =   3675
      End
      Begin VB.Image cmdMoverHechi 
         Height          =   300
         Index           =   0
         Left            =   3660
         MouseIcon       =   "frmMain.frx":57E2
         MousePointer    =   99  'Custom
         Tag             =   "0"
         Top             =   1455
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Image cmdMoverHechi 
         Height          =   300
         Index           =   1
         Left            =   3660
         MouseIcon       =   "frmMain.frx":5934
         MousePointer    =   99  'Custom
         Tag             =   "0"
         Top             =   1140
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Image imgHechizos 
         Height          =   495
         Left            =   1995
         Tag             =   "0"
         Top             =   15
         Width           =   1995
      End
      Begin VB.Image imgInventario 
         Height          =   495
         Left            =   0
         Tag             =   "0"
         Top             =   15
         Width           =   1995
      End
      Begin VB.Image imgSpellInfo 
         Height          =   300
         Left            =   3660
         Tag             =   "1"
         Top             =   1755
         Width           =   300
      End
      Begin VB.Image imgInvLock 
         Height          =   210
         Index           =   2
         Left            =   75
         Top             =   4125
         Width           =   210
      End
      Begin VB.Image imgInvLock 
         Height          =   210
         Index           =   1
         Left            =   75
         Top             =   3600
         Width           =   210
      End
      Begin VB.Image imgInvLock 
         Height          =   210
         Index           =   0
         Left            =   75
         Top             =   3090
         Width           =   210
      End
   End
   Begin VB.PictureBox shapexy 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   16920
      ScaleHeight     =   150
      ScaleWidth      =   150
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   7080
      Width           =   180
   End
   Begin VB.Timer dobleclick 
      Left            =   8520
      Top             =   2400
   End
   Begin VB.TextBox SendTxtCmsg 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   360
      Left            =   120
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   2100
      Visible         =   0   'False
      Width           =   11040
   End
   Begin VB.Timer Evento 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   6360
      Top             =   2400
   End
   Begin VB.Timer UpdateDaytime 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3480
      Top             =   2400
   End
   Begin VB.Timer Efecto 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2040
      Top             =   2400
   End
   Begin VB.Timer MacroLadder 
      Enabled         =   0   'False
      Interval        =   1300
      Left            =   1560
      Top             =   2400
   End
   Begin VB.Timer TimerNiebla 
      Interval        =   100
      Left            =   1080
      Top             =   2400
   End
   Begin VB.Timer TimerLluvia 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   600
      Top             =   2400
   End
   Begin VB.Timer UpdateLight 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3000
      Top             =   2400
   End
   Begin VB.Timer Contadores 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3960
      Top             =   2400
   End
   Begin VB.Timer cerrarcuenta 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   5400
      Top             =   2400
   End
   Begin VB.Timer LlamaDeclan 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   4920
      Top             =   2400
   End
   Begin VB.TextBox SendTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00202020&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   360
      Left            =   120
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   2100
      Visible         =   0   'False
      Width           =   11040
   End
   Begin VB.Timer macrotrabajo 
      Enabled         =   0   'False
      Left            =   2520
      Top             =   2400
   End
   Begin VB.Timer ShowFPS 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5880
      Top             =   2400
   End
   Begin VB.PictureBox renderer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   10980
      Left            =   135
      ScaleHeight     =   732
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   736
      TabIndex        =   3
      Top             =   480
      Width           =   11040
   End
   Begin VB.Shape shpCooldownComboBlock 
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   11505
      Top             =   9315
      Width           =   15
   End
   Begin VB.Image btnMenu 
      Height          =   375
      Left            =   11505
      Tag             =   "0"
      Top             =   11010
      Width           =   1410
   End
   Begin VB.Image Image4 
      Height          =   270
      Index           =   0
      Left            =   14535
      Tag             =   "0"
      Top             =   90
      Width           =   300
   End
   Begin VB.Label lblLvl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B6B6B6&
      Height          =   345
      Left            =   14670
      TabIndex        =   9
      Top             =   1560
      Width           =   555
   End
   Begin VB.Label lblClase 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Guerrero"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BFBFBF&
      Height          =   270
      Left            =   11760
      TabIndex        =   36
      Top             =   1140
      Width           =   3165
   End
   Begin VB.Image ImgSegClan 
      Appearance      =   0  'Flat
      Height          =   510
      Left            =   13755
      Picture         =   "frmMain.frx":5A86
      ToolTipText     =   "Seguro de clan"
      Top             =   10950
      Width           =   510
   End
   Begin VB.Label lblResis 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B6B6B6&
      Height          =   240
      Left            =   14430
      TabIndex        =   35
      ToolTipText     =   "Resistencia m�gica"
      Top             =   10395
      Width           =   600
   End
   Begin VB.Image ImgSeg 
      Appearance      =   0  'Flat
      Height          =   510
      Left            =   14280
      Picture         =   "frmMain.frx":6898
      ToolTipText     =   "Seguro de ataque"
      Top             =   10950
      Width           =   510
   End
   Begin VB.Image ImgSegParty 
      Height          =   510
      Left            =   13230
      Picture         =   "frmMain.frx":76AA
      ToolTipText     =   "Seguro de grupo"
      Top             =   10950
      Width           =   510
   End
   Begin VB.Label lblArmor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B6B6B6&
      Height          =   240
      Left            =   11970
      TabIndex        =   34
      ToolTipText     =   "Defensa armadura"
      Top             =   10395
      Width           =   600
   End
   Begin VB.Image ImgSegResu 
      Appearance      =   0  'Flat
      Height          =   510
      Left            =   14805
      Picture         =   "frmMain.frx":84BC
      ToolTipText     =   "Seguro de resurrecci�n"
      Top             =   10950
      Width           =   510
   End
   Begin VB.Image imgOro 
      Height          =   375
      Left            =   13200
      Top             =   7920
      Width           =   375
   End
   Begin VB.Label GldLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100.000"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   270
      Left            =   13680
      TabIndex        =   33
      ToolTipText     =   "Monedas de oro"
      Top             =   7995
      Width           =   690
   End
   Begin VB.Label AgilidadLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "40"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   270
      Left            =   12690
      TabIndex        =   32
      ToolTipText     =   "Agilidad"
      Top             =   8010
      Width           =   210
   End
   Begin VB.Label Fuerzalbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "40"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006C9A28&
      Height          =   270
      Left            =   12030
      TabIndex        =   31
      ToolTipText     =   "Fuerza"
      Top             =   8010
      Width           =   210
   End
   Begin VB.Label lblHelm 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B6B6B6&
      Height          =   240
      Left            =   14430
      TabIndex        =   30
      ToolTipText     =   "Defensa casco"
      Top             =   10005
      Width           =   600
   End
   Begin VB.Label lblShielder 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B6B6B6&
      Height          =   240
      Left            =   13170
      TabIndex        =   29
      ToolTipText     =   "Defensa escudo"
      Top             =   10005
      Width           =   600
   End
   Begin VB.Label lblWeapon 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B6B6B6&
      Height          =   240
      Left            =   11970
      TabIndex        =   28
      ToolTipText     =   "Da�o f�sico arma"
      Top             =   10005
      Width           =   600
   End
   Begin VB.Label lbldm 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+0%"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B6B6B6&
      Height          =   240
      Left            =   13170
      TabIndex        =   27
      ToolTipText     =   "Aumento de da�o m�gico"
      Top             =   10395
      Width           =   600
   End
   Begin VB.Label AGUbar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   13455
      TabIndex        =   26
      Top             =   9525
      Width           =   675
   End
   Begin VB.Label hambar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   14505
      TabIndex        =   25
      Top             =   9525
      Width           =   675
   End
   Begin VB.Label stabar 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   11685
      TabIndex        =   24
      Top             =   9525
      Width           =   1350
   End
   Begin VB.Label manabar 
      Alignment       =   2  'Center
      BackColor       =   &H80000006&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B6B6B6&
      Height          =   240
      Left            =   11550
      TabIndex        =   23
      Top             =   9015
      Width           =   3585
   End
   Begin VB.Label HpBar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B6B6B6&
      Height          =   240
      Left            =   11550
      TabIndex        =   22
      Top             =   8505
      Width           =   3585
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   7680
      Top             =   120
      Width           =   375
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderStyle     =   6  'Inside Solid
      X1              =   512
      X2              =   537
      Y1              =   20
      Y2              =   20
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      BorderStyle     =   6  'Inside Solid
      X1              =   512
      X2              =   537
      Y1              =   12
      Y2              =   12
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderStyle     =   6  'Inside Solid
      X1              =   512
      X2              =   537
      Y1              =   16
      Y2              =   16
   End
   Begin VB.Label btnInvisible 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Invisible"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   375
      Left            =   6000
      TabIndex        =   20
      Top             =   75
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label btnSpawn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Spawn NPC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   375
      Left            =   4560
      TabIndex        =   19
      Top             =   75
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label createObj 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Crear Obj"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   375
      Left            =   3120
      TabIndex        =   18
      Top             =   75
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label panelGM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PanelGM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   375
      Left            =   1800
      TabIndex        =   17
      Top             =   75
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   495
      Left            =   9360
      TabIndex        =   16
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   11880
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label lblhora 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B6B6B6&
      Height          =   225
      Left            =   13170
      TabIndex        =   14
      Top             =   2010
      Width           =   495
   End
   Begin VB.Label ms 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "30 ms"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   8520
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label fps 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fps: 200"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   8490
      TabIndex        =   12
      ToolTipText     =   "Numero de usuarios online"
      Top             =   150
      Width           =   645
   End
   Begin VB.Label Coord 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000 X:00 Y: 00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   9720
      TabIndex        =   11
      Top             =   210
      Width           =   1215
   End
   Begin VB.Label NombrePJ 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del pj"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BFBFBF&
      Height          =   495
      Left            =   11760
      TabIndex        =   10
      Top             =   540
      Width           =   3150
   End
   Begin VB.Label lblPorcLvl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "33.33%"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   11640
      TabIndex        =   8
      Top             =   1605
      Visible         =   0   'False
      Width           =   2820
   End
   Begin VB.Image OpcionesBoton 
      Height          =   315
      Left            =   11431
      Tag             =   "0"
      Top             =   65
      Width           =   315
   End
   Begin VB.Image CombateIcon 
      Height          =   180
      Left            =   8828
      Picture         =   "frmMain.frx":92CE
      Top             =   1812
      Width           =   555
   End
   Begin VB.Image globalIcon 
      Height          =   180
      Left            =   8828
      Picture         =   "frmMain.frx":9852
      Top             =   2008
      Width           =   555
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   10320
      TabIndex        =   5
      ToolTipText     =   "Activar / desactivar chat globales"
      Top             =   1800
      Width           =   855
   End
   Begin VB.Image Image3 
      Height          =   210
      Left            =   284
      Tag             =   "0"
      ToolTipText     =   "Modo de chat"
      Top             =   1894
      Width           =   255
   End
   Begin VB.Image Image4 
      Height          =   270
      Index           =   1
      Left            =   14910
      Tag             =   "0"
      Top             =   90
      Width           =   300
   End
   Begin VB.Label onlines 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Online: 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   12480
      TabIndex        =   4
      ToolTipText     =   "Numero de usuarios online"
      Top             =   120
      Width           =   1665
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7680
      TabIndex        =   2
      Top             =   1680
      Width           =   450
   End
   Begin VB.Label NameMapa 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mapa Desconocido"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   165
      Left            =   9765
      TabIndex        =   1
      Top             =   45
      Width           =   1125
   End
   Begin VB.Label exp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99999/99999"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   11580
      TabIndex        =   7
      Top             =   1605
      Width           =   2940
   End
   Begin VB.Image ExpBar 
      Height          =   180
      Left            =   11385
      Picture         =   "frmMain.frx":9DD6
      Top             =   1635
      Width           =   3255
   End
   Begin VB.Image STAShp 
      Height          =   225
      Left            =   11610
      Top             =   9510
      Width           =   1455
   End
   Begin VB.Image AGUAsp 
      Height          =   225
      Left            =   13425
      Top             =   9510
      Width           =   705
   End
   Begin VB.Image COMIDAsp 
      Height          =   225
      Left            =   14475
      Top             =   9510
      Width           =   705
   End
   Begin VB.Image Hpshp 
      Height          =   330
      Left            =   11505
      Top             =   8475
      Width           =   3645
   End
   Begin VB.Image MANShp 
      Height          =   330
      Left            =   11505
      Top             =   8985
      Width           =   3645
   End
   Begin VB.Image EstadisticasBoton 
      Height          =   690
      Left            =   14640
      Tag             =   "0"
      Top             =   1320
      Width           =   675
   End
   Begin VB.Menu mnuObj 
      Caption         =   "Objeto"
      Visible         =   0   'False
      Begin VB.Menu mnuTirar 
         Caption         =   "Tirar"
      End
      Begin VB.Menu mnuUsar 
         Caption         =   "Usar"
      End
      Begin VB.Menu mnuEquipar 
         Caption         =   "Equipar"
      End
   End
   Begin VB.Menu mnuNpc 
      Caption         =   "NPC"
      Visible         =   0   'False
      Begin VB.Menu mnuNpcDesc 
         Caption         =   "Descripcion"
      End
      Begin VB.Menu mnuNpcComerciar 
         Caption         =   "Comerciar"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'You can contact me at:
'morgolock@speedy.com.ar
'
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez
'Call ParseUserCommand("/CMSG " & stxtbuffercmsg)
Option Explicit

Private Declare Sub svb_shutdown_steam Lib "steam_vb.dll" ()


Public WithEvents Inventario As clsGrapchicalInventory
Attribute Inventario.VB_VarHelpID = -1

Private Const WS_EX_TRANSPARENT = &H20&
Private totalclicks As Integer
Private Const GWL_EXSTYLE = (-20)

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

' Constantes para SendMessage
Const WM_SYSCOMMAND As Long = &H112&

Const MOUSE_MOVE    As Long = &HF012&

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private MenuNivel As Byte

Private Type POINTAPI

    x As Long
    y As Long

End Type


Private Declare Function ReleaseCapture Lib "user32" () As Long

Public MouseBoton As Long

Public MouseShift As Long

Public IsPlaying  As Byte

Public ShowPercentage As Boolean

Public bmoving    As Boolean

Public dX         As Integer

Public dY         As Integer

Private Const IntervaloEntreClicks As Long = 50

Dim TempTick As Long

Private iClickTick As Long

' Constantes para SendMessage

Const HWND_TOPMOST = -1

Const HWND_NOTOPMOST = -2

Const SWP_NOSIZE = &H1

Const SWP_NOMOVE = &H2

Const SWP_NOACTIVATE = &H10

Const SWP_SHOWWINDOW = &H40

Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As String) As Long

Private Const EM_GETLINE = &HC4

Private Const EM_LINELENGTH = &HC1
Private cBotonEliminarItem As clsGraphicalButton

Private Sub btnInvisible_Click()
    
    On Error GoTo btnInvisible_Click_Err
    
    Call ParseUserCommand("/INVISIBLE")
    
    Me.SetFocus
    
    Exit Sub

btnInvisible_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.btnInvisible_Click", Erl)
    Resume Next
    
End Sub


Private Sub loadButtons()

    Set cBotonEliminarItem = New clsGraphicalButton
                                                
    Call cBotonEliminarItem.Initialize(imgDeleteItem, "boton-borrar-item-default.bmp", _
                                                "boton-borrar-item-over.bmp", _
                                                "boton-borrar-item-off.bmp", Me)
End Sub

Private Sub btnMenu_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo Handler
    
    btnMenu.Picture = LoadInterface("boton-menu-off.bmp")
    btnMenu.Tag = "2"

    
    Exit Sub

Handler:
    Call RegistrarError(err.Number, err.Description, "frmMain.btnMenu_MouseDown", Erl)
    Resume Next
End Sub

Private Sub btnMenu_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo Handler
    
    If btnMenu.Tag = "0" Then
        btnMenu.Picture = LoadInterface("boton-menu-over.bmp")
        btnMenu.Tag = "1"
    End If
    
    Exit Sub

Handler:
    Call RegistrarError(err.Number, err.Description, "frmMain.btnMenu_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub btnMenu_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo Handler
    
    btnMenu.Picture = Nothing
    btnMenu.Tag = "0"

    
    Exit Sub

Handler:
    Call RegistrarError(err.Number, err.Description, "frmMain.btnMenu_MouseDown", Erl)
    Resume Next
End Sub

Private Sub btnSpawn_Click()
    
    On Error GoTo btnSpawn_Click_Err
    
    Me.SetFocus
    
    Call WriteSpawnListRequest
    
    Exit Sub

btnSpawn_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.btnSpawn_Click", Erl)
    Resume Next
    
End Sub



Private Sub clanimg_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo clanimg_MouseUp_Err
    
    If pausa Then Exit Sub

    If frmGuildLeader.visible Then Unload frmGuildLeader
    Call WriteRequestGuildLeaderInfo

    
    Exit Sub

clanimg_MouseUp_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.clanimg_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub cmdlanzar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo cmdlanzar_MouseDown_Err
    
    
    If ModoHechizos = BloqueoLanzar Then
        If Not MainTimer.Check(TimersIndex.AttackSpell, False) Or Not MainTimer.Check(TimersIndex.CastSpell, False) Then
            Exit Sub
        End If
    End If
    
    cmdlanzar.Picture = LoadInterface("boton-lanzar-off.bmp")
    cmdlanzar.Tag = "1"

    
    Exit Sub

cmdlanzar_MouseDown_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.cmdlanzar_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub cmdlanzar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo cmdlanzar_MouseUp_Err
    
    If ModoHechizos = BloqueoLanzar Then
        If Not MainTimer.Check(TimersIndex.AttackSpell, False) Or Not MainTimer.Check(TimersIndex.CastSpell, False) Then
            Exit Sub
        End If
    End If
    
    cmdlanzar.Picture = LoadInterface("boton-lanzar-over.bmp")
    cmdlanzar.Tag = "1"
    
    Exit Sub

cmdlanzar_MouseUp_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.cmdlanzar_MouseUp", Erl)
    Resume Next
    
End Sub
Private Sub cmdMoverHechi_Click(Index As Integer)
    
    On Error GoTo cmdMoverHechi_Click_Err
    

    If hlst.ListIndex = -1 Then Exit Sub

    Dim sTemp As String

    Select Case Index

        Case 1 'subir

            If hlst.ListIndex = 0 Then Exit Sub

        Case 0 'bajar

            If hlst.ListIndex = hlst.ListCount - 1 Then Exit Sub

    End Select

    Call WriteMoveSpell(Index, hlst.ListIndex + 1)
    
    Select Case Index

        Case 1 'subir
            sTemp = hlst.List(hlst.ListIndex - 1)
            hlst.List(hlst.ListIndex - 1) = hlst.List(hlst.ListIndex)
            hlst.List(hlst.ListIndex) = sTemp
            hlst.ListIndex = hlst.ListIndex - 1

        Case 0 'bajar
            sTemp = hlst.List(hlst.ListIndex + 1)
            hlst.List(hlst.ListIndex + 1) = hlst.List(hlst.ListIndex)
            hlst.List(hlst.ListIndex) = sTemp
            hlst.ListIndex = hlst.ListIndex + 1

    End Select

    
    Exit Sub

cmdMoverHechi_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.cmdMoverHechi_Click", Erl)
    Resume Next
    
End Sub

Public Sub ControlSeguroParty(ByVal Mostrar As Boolean)
    
    On Error GoTo ControlSeguroParty_Err

    If Mostrar Then
        ImgSegParty = LoadInterface("boton-seguro-party-on.bmp")
        SeguroParty = True
    Else
        ImgSegParty = LoadInterface("boton-seguro-party-off.bmp")
        SeguroParty = False
    End If
    
    Exit Sub

ControlSeguroParty_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.ControlSeguroParty", Erl)
    Resume Next
    
End Sub

Public Sub DibujarSeguro()
    
    On Error GoTo DibujarSeguro_Err
    
    ImgSeg = LoadInterface("boton-seguro-ciudadano-on.bmp")
    
    SeguroGame = True

    
    Exit Sub

DibujarSeguro_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.DibujarSeguro", Erl)
    Resume Next
    
End Sub

Public Sub DesDibujarSeguro()
    
    On Error GoTo DesDibujarSeguro_Err
    
    ImgSeg = LoadInterface("boton-seguro-ciudadano-off.bmp")
    
    SeguroGame = False

    
    Exit Sub

DesDibujarSeguro_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.DesDibujarSeguro", Erl)
    Resume Next
    
End Sub

Private Sub cmdMoverHechi_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    
    On Error GoTo cmdMoverHechi_MouseMove_Err
    

    Select Case Index

        Case 0

            cmdMoverHechi(Index).Picture = LoadInterface("boton-sm-flecha-aba-off.bmp")
            cmdMoverHechi(Index).Tag = "2"


        Case 1

            cmdMoverHechi(Index).Picture = LoadInterface("boton-sm-flecha-arr-off.bmp")
            cmdMoverHechi(Index).Tag = "2"

    
    End Select
    
    
    Exit Sub

cmdMoverHechi_MouseMove_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.cmdMoverHechi_MouseMove", Erl)
    Resume Next
End Sub

Private Sub cmdMoverHechi_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo cmdMoverHechi_MouseMove_Err
    

    Select Case Index

        Case 0

            If cmdMoverHechi(Index).Tag = "0" Then
                cmdMoverHechi(Index).Picture = LoadInterface("boton-sm-flecha-aba-over.bmp")
                cmdMoverHechi(Index).Tag = "1"

            End If
            
            cmdMoverHechi(1).Picture = Nothing
            imgSpellInfo.Picture = Nothing
            cmdMoverHechi(1).Tag = "0"
            imgSpellInfo.Tag = "0"

        Case 1

            If cmdMoverHechi(Index).Tag = "0" Then
                cmdMoverHechi(Index).Picture = LoadInterface("boton-sm-flecha-arr-over.bmp")
                cmdMoverHechi(Index).Tag = "1"

            End If
            
            cmdMoverHechi(0).Picture = Nothing
            imgSpellInfo.Picture = Nothing
            cmdMoverHechi(0).Tag = "0"
            imgSpellInfo.Tag = "0"
    
    End Select
    
    
    Exit Sub

cmdMoverHechi_MouseMove_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.cmdMoverHechi_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub cmdMoverHechi_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo cmdMoverHechi_MouseMove_Err
    

    Select Case Index

        Case 0

            Set cmdMoverHechi(Index).Picture = Nothing
            cmdMoverHechi(Index).Tag = "0"


        Case 1

            Set cmdMoverHechi(Index).Picture = Nothing
            cmdMoverHechi(Index).Tag = "0"

    
    End Select
    
    
    Exit Sub

cmdMoverHechi_MouseMove_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.cmdMoverHechi_MouseMove", Erl)
    Resume Next
End Sub

Private Sub CombateIcon_Click()
    
    On Error GoTo CombateIcon_Click_Err
    

    If ChatCombate = 0 Then
        ChatCombate = 1
        CombateIcon.Picture = LoadInterface("infoapretado.bmp")
    Else
        ChatCombate = 0
        CombateIcon.Picture = LoadInterface("info.bmp")

    End If

    Call WriteMacroPos

    
    Exit Sub

CombateIcon_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.CombateIcon_Click", Erl)
    Resume Next
    
End Sub



Private Sub Command2_Click()

End Sub

Private Sub Contadores_Timer()
    
    On Error GoTo Contadores_Timer_Err
    

    If UserEstado = 1 Then Exit Sub
    If InviCounter > 0 Then
        InviCounter = InviCounter - 1

    End If


    If DrogaCounter > 0 Then
        DrogaCounter = DrogaCounter - 1

        If DrogaCounter <= 12 And DrogaCounter > 0 Then
                Call Sound.Sound_Stop(SND_DOPA)
                Call Sound.Sound_Play(SND_DOPA)
            If DrogaCounter Mod 2 = 0 Then
                frmMain.Fuerzalbl.ForeColor = vbWhite
                frmMain.AgilidadLbl.ForeColor = vbWhite
            Else
                frmMain.Fuerzalbl.ForeColor = RGB(204, 0, 0)
                frmMain.AgilidadLbl.ForeColor = RGB(204, 0, 0)
            End If
        End If
        
        If DrogaCounter <= 12 And DrogaCounter > 0 Then
        End If
        
    End If

    If InviCounter = 0 And DrogaCounter = 0 Then
        Contadores.enabled = False

    End If

    
    Exit Sub

Contadores_Timer_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.Contadores_Timer", Erl)
    Resume Next
    
End Sub

Private Sub createObj_Click()
    
    On Error GoTo createObj_Click_Err
    
    Dim i As Long
    For i = 1 To NumOBJs

        If ObjData(i).Name <> "" Then

            Dim subelemento As ListItem

            Set subelemento = frmObjetos.ListView1.ListItems.Add(, , ObjData(i).Name)
            
            subelemento.SubItems(1) = i

        End If

    Next i
    
    Me.SetFocus
    
    frmObjetos.Show , Me
    
    Exit Sub

createObj_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.createObj_Click", Erl)
    Resume Next
    
End Sub

Private Sub dobleclick_Timer()
    Static segundo As Long
    segundo = segundo + 1
    If segundo = 2 And totalclicks > 20 Then
    Call WriteLogMacroClickHechizo(tMacro.dobleclick, totalclicks)
    totalclicks = 0
    segundo = 0
    dobleclick.Interval = 0
    'Label10.Caption = 0
    End If
    If segundo = 2 And totalclicks <= 20 Then
    totalclicks = 0
    segundo = 0
    dobleclick.Interval = 0
    End If
End Sub

Private Sub Efecto_Timer()
    
    On Error GoTo Efecto_Timer_Err
    
    If MapDat.Base_light > 0 Then
        Call SetGlobalLight(MapDat.Base_light)
    Else
        Call RestaurarLuz
    End If

    Efecto.enabled = False
    EfectoEnproceso = False

    
    Exit Sub

Efecto_Timer_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.Efecto_Timer", Erl)
    Resume Next
    
End Sub

Private Sub hlst_Click()
    

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Form_KeyDown_Err
    


    If Not SendTxt.visible And Not SendTxtCmsg.visible Then
        If Not pausa And frmMain.visible And Not frmComerciar.visible And Not frmComerciarUsu.visible And Not frmBancoObj.visible And Not frmGoliath.visible Then
            If KeyCode = BindKeys(24).KeyCode Then
                
                If UserCharIndex > 0 Then
                    If Not minimapa_visible Then
                        minimapa_visible = True
                        MouseZona = 0
                        MouseMapX = 0
                        MouseMapY = 0
                        Call Sound.Sound_Play(CStr(213), False, Sound.Calculate_Volume(UserPos.x, UserPos.y), Sound.Calculate_Pan(UserPos.x, UserPos.y))
                    End If
                   
                End If
            End If
            
        End If
    End If
                
    
    Exit Sub

Form_KeyDown_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.Form_KeyDown", Erl)
    Resume Next
End Sub

Private Sub Image1_Click()
    ConsoleDialog.menu_consola_visible = Not ConsoleDialog.menu_consola_visible
    ConsoleDialog.cambiando_intensidad = False
End Sub

Private Sub EstadisticasBoton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo EstadisticasBoton_MouseDown_Err
    
    
    
    EstadisticasBoton.Picture = LoadInterface("boton-skills-off.bmp")
    EstadisticasBoton.Tag = "1"
    
    Exit Sub

EstadisticasBoton_MouseDown_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.EstadisticasBoton_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub EstadisticasBoton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo EstadisticasBoton_MouseMove_Err
    

    If EstadisticasBoton.Tag = "0" Then
        EstadisticasBoton.Picture = LoadInterface("boton-skills-over.bmp")
        EstadisticasBoton.Tag = "1"
    End If

    
    Exit Sub

EstadisticasBoton_MouseMove_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.EstadisticasBoton_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub EstadisticasBoton_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo EstadisticasBoton_MouseUp_Err
    
    
    If pausa Or tutorial_index > 0 Then Exit Sub
    
    If MostrarTutorial And tutorial_index <= 0 Then
        If tutorial(4).Activo = 1 Then
            tutorial_index = e_tutorialIndex.TUTORIAL_SkillPoints
            'TUTORIAL MAPA INSEGURO
            Call mostrarCartel(tutorial(tutorial_index).titulo, tutorial(tutorial_index).textos(1), tutorial(tutorial_index).grh, -1, &H164B8A, , , False, 100, 629, 100, 685, 640, 530, 64, 64)
            Exit Sub
        End If
    End If
    
    LlegaronSkills = True
    Call WriteRequestSkills
    
    Exit Sub

EstadisticasBoton_MouseUp_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.EstadisticasBoton_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub Evento_Timer()

    InvasionActual = 0
    
    Evento.enabled = False

End Sub

Private Sub exp_Click()
    
    On Error GoTo exp_Click_Err

    'Call WriteScrollInfo

    ShowPercentage = Not ShowPercentage
    
    Exit Sub

exp_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.exp_Click", Erl)
    Resume Next
    
End Sub

Private Sub exp_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Not ShowPercentage Then
        lblPorcLvl.visible = True
        exp.visible = False
    End If
    
End Sub

Private Sub Form_Activate()
    
    On Error GoTo Form_Activate_Err
    
    Exit Sub

Form_Activate_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.Form_Activate", Erl)
    Resume Next
    
End Sub


Private Sub imgHechizos_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo imgHechizos_MouseMove_Err
    
    imgHechizos.Picture = Nothing
    imgHechizos.Tag = "1"

    
    Exit Sub

imgHechizos_MouseMove_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.imgHechizos_MouseMove", Erl)
    Resume Next
End Sub

Private Sub imgInventario_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo imgInventario_MouseMove_Err

    imgInventario.Picture = LoadInterface("boton-inventory-over.bmp")
    imgInventario.Tag = "1"

    
    Exit Sub

imgInventario_MouseMove_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.imgInventario_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub mapMundo_Click()
    On Error GoTo mapMundo_Click_Err
    
    ExpMult = 1
    OroMult = 1
    Call frmMapaGrande.CalcularPosicionMAPA
    frmMapaGrande.Picture = LoadInterface("ventanamapa.bmp")
    frmMapaGrande.Show , frmMain

    
    Exit Sub

mapMundo_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.mapMundo_Click", Erl)
    Resume Next
End Sub

Private Sub imgSpellInfo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
imgSpellInfo.Picture = LoadInterface("boton-sm-ayuda-off.bmp")
imgSpellInfo.Tag = "2"
End Sub

Private Sub imgSpellInfo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If imgSpellInfo.Tag = "0" Then
    imgSpellInfo.Picture = LoadInterface("boton-sm-ayuda-over.bmp")
    imgSpellInfo.Tag = "1"
End If
cmdMoverHechi(0).Picture = Nothing
cmdMoverHechi(1).Picture = Nothing
cmdMoverHechi(0).Tag = "0"
cmdMoverHechi(1).Tag = "0"
End Sub

Private Sub imgSpellInfo_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Set imgSpellInfo.Picture = Nothing
imgSpellInfo.Tag = "0"
End Sub

Private Sub PicCorreo_Click()

End Sub

Private Sub lblLvl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Call EstadisticasBoton_MouseUp(Button, Shift, x, y)
End Sub

Private Sub picHechiz_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If y < 0 Then y = 0
If y > Int(picHechiz.ScaleHeight / hlst.Pixel_Alto) * hlst.Pixel_Alto - 1 Then y = Int(picHechiz.ScaleHeight / hlst.Pixel_Alto) * hlst.Pixel_Alto - 1
If x < picHechiz.ScaleWidth - 10 Then
    hlst.ListIndex = Int(y / hlst.Pixel_Alto) + hlst.Scroll
    hlst.DownBarrita = 0
    If Seguido = 1 Then
        Call WriteNotifyInventarioHechizos(2, hlst.ListIndex, hlst.Scroll)
    End If
Else
    hlst.DownBarrita = y - hlst.Scroll * (picHechiz.ScaleHeight - hlst.BarraHeight) / (hlst.ListCount - hlst.VisibleCount)
End If
End Sub

Private Sub picHechiz_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    Dim yy As Integer
    yy = y
    If yy < 0 Then yy = 0
    If yy > Int(picHechiz.ScaleHeight / hlst.Pixel_Alto) * hlst.Pixel_Alto - 1 Then yy = Int(picHechiz.ScaleHeight / hlst.Pixel_Alto) * hlst.Pixel_Alto - 1
    If hlst.DownBarrita > 0 Then
        hlst.Scroll = (y - hlst.DownBarrita) * (hlst.ListCount - hlst.VisibleCount) / (picHechiz.ScaleHeight - hlst.BarraHeight)
    Else
        hlst.ListIndex = Int(yy / hlst.Pixel_Alto) + hlst.Scroll
        If Seguido = 1 Then
            Call WriteNotifyInventarioHechizos(2, hlst.ListIndex, hlst.Scroll)
        End If
        If ScrollArrastrar = 0 Then
            If (y < yy) Then hlst.Scroll = hlst.Scroll - 1
            If (y > yy) Then hlst.Scroll = hlst.Scroll + 1
        End If
    End If
ElseIf Button = 0 Then
    hlst.ShowBarrita = x > picHechiz.ScaleWidth - hlst.BarraWidth * 2
End If
End Sub

Private Sub picHechiz_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
hlst.DownBarrita = 0
End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not picInv.visible Then Exit Sub
    
    If dobleclick.Interval = 0 Then dobleclick.Interval = 1000
    
    If Button = 1 Then
        dobleclick.Interval = 1000
        totalclicks = totalclicks + 1
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo Form_KeyUp_Err
    


    If Not SendTxt.visible And Not SendTxtCmsg.visible Then
        If Not pausa And frmMain.visible And Not frmComerciar.visible And Not frmComerciarUsu.visible And Not frmBancoObj.visible And Not frmGoliath.visible Then
            
            
            
            If Accionar(KeyCode) Then
                Exit Sub
            ElseIf KeyCode = BindKeys(24).KeyCode Then
                If minimapa_visible Then
                    Call Sound.Sound_Play(CStr(214), False, Sound.Calculate_Volume(UserPos.x, UserPos.y), Sound.Calculate_Pan(UserPos.x, UserPos.y))
                    minimapa_visible = False
                End If
            ElseIf KeyCode = vbKeyReturn Then

                If Not frmCantidad.visible Then
                    
                    Call CompletarEnvioMensajes
                    'HarThaoS: Al abrir el textBox de escritura tomo el tiempo de inicio para controlar macro de cartel
                    StartOpenChatTime = GetTickCount
                    SendTxt.visible = True
                    SendTxt.SetFocus
                End If
                
            ElseIf KeyCode = vbKeyDelete Then
                If Not SendTxt.visible Then
                    SendTxtCmsg.visible = True
                    SendTxtCmsg.SetFocus
                End If
            ElseIf KeyCode = vbKeyEscape And Not UserSaliendo Then
                frmCerrar.Show , frmMain
                ' Call WriteQuit
        
            ElseIf KeyCode = 27 And UserSaliendo Then
                Call WriteCancelarExit

                Rem  Call SendData("CU")
            ElseIf KeyCode = 80 And PescandoEspecial Then
                Call IntentarObtenerPezEspecial
            End If

        End If

    Else
    
        If SendTxt.visible Then
            SendTxt.SetFocus
        End If
        
        If SendTxtCmsg.visible Then
            SendTxtCmsg.SetFocus
        End If
        
    End If

    
    Exit Sub

Form_KeyUp_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.Form_KeyUp", Erl)
    Resume Next
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Form_MouseDown_Err
    
    
    If SendTxt.visible Then SendTxt.SetFocus
    MouseBoton = Button
    MouseShift = Shift
    
    If frmComerciar.visible Then
        Unload frmComerciar

    End If
    
    If frmBancoObj.visible Then
        Unload frmBancoObj

    End If
    
    If frmQuestInfo.visible Then
        Unload frmQuestInfo
    End If
    
    Exit Sub

Form_MouseDown_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.Form_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Form_MouseUp_Err
    
    clicX = x
    clicY = y
    
    Call Minimap.MouseUp(0, x, y)
    
    Exit Sub

Form_MouseUp_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.Form_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    On Error GoTo Form_QueryUnload_Err
    

    If prgRun = True Then
        prgRun = False
        Cancel = 1

    End If

    
    Exit Sub

Form_QueryUnload_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.Form_QueryUnload", Erl)
    Resume Next
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error GoTo Form_Unload_Err
    
    Call svb_shutdown_steam
    Call DisableURLDetect

    Exit Sub

Form_Unload_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.Form_Unload", Erl)
    Resume Next
    
End Sub

Private Sub GldLbl_Click()
    
    On Error GoTo GldLbl_Click_Err
    
    Inventario.SelectGold

    If UserGLD > 0 Then
        frmCantidad.Show , frmMain

    End If

    
    Exit Sub

GldLbl_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.GldLbl_Click", Erl)
    Resume Next
    
End Sub

Private Sub GlobalIcon_Click()
    
    On Error GoTo GlobalIcon_Click_Err
    

    If ChatGlobal = 0 Then
        ChatGlobal = 1
        globalIcon.Picture = LoadInterface("globalapretado.bmp")
    Else
        ChatGlobal = 0
        globalIcon.Picture = LoadInterface("global.bmp")

    End If

    Call WriteMacroPos

    
    Exit Sub

GlobalIcon_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.GlobalIcon_Click", Erl)
    Resume Next
    
End Sub

Private Sub Image2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image2_MouseDown_Err
    

    If Index = 0 Then
        If OpcionMenu <> 0 Then

            ' Image2(Index).Tag = "1"
            ' Image2(Index).Picture = LoadInterface("botoninventarioapretado.bmp")
            Rem    Image2(1).Picture = LoadInterface("botonconjuros.bmp")
            Rem   Image2(2).Picture = LoadInterface("botonmenu.bmp")
        End If

    End If

    If Index = 1 Then
        If OpcionMenu <> 1 Then

            ' Image2(Index).Tag = "1"
            '  Image2(1).Picture = LoadInterface("botonconjurosapretado.bmp")
            Rem   Image2(2).Picture = LoadInterface("botonmenu.bmp")
            Rem    Image2(0).Picture = LoadInterface("botoninventario.bmp")
        End If

    End If

    
    Exit Sub

Image2_MouseDown_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.Image2_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub HpBar_Click()
    Call ParseUserCommand("/PROMEDIO")
End Sub

Private Sub Hpshp_Click()
    HpBar_Click
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    '    Image3.Picture = LoadInterface("elegirchatapretado.bmp")
    
    On Error GoTo Image3_MouseDown_Err
    
    frmMensaje.PopupMenuMensaje

    
    Exit Sub

Image3_MouseDown_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.Image3_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image3_MouseMove_Err
    

    If Image3.Tag = "0" Then
        Image3.Picture = LoadInterface("elegirchatmarcado.bmp")
        Image3.Tag = "1"

    End If

    
    Exit Sub

Image3_MouseMove_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.Image3_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Image4_Click(Index As Integer)
    
    On Error GoTo Image4_Click_Err
    

    Select Case Index

        Case 0
            Me.WindowState = vbMinimized

        Case 1
            If frmCerrar.visible Then Exit Sub
            Dim mForm As Form
            For Each mForm In Forms
                If mForm.hWnd <> Me.hWnd Then Unload mForm
                Set mForm = Nothing
            Next
            frmCerrar.Show , Me

    End Select

    
    Exit Sub

Image4_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.Image4_Click", Erl)
    Resume Next
    
End Sub

Private Sub Image4_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image4_MouseDown_Err
    

    Select Case Index

        Case 0
            Image4(0).Picture = LoadInterface("boton-sm-minimizar-off.bmp")

        Case 1
            Image4(1).Picture = LoadInterface("boton-sm-cerrar-off.bmp")

    End Select

    
    Exit Sub

Image4_MouseDown_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.Image4_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub Image4_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image4_MouseMove_Err
    

    Select Case Index

        Case 0

            If Image4(Index).Tag = "0" Then
                Image4(Index).Picture = LoadInterface("boton-sm-minimizar-over.bmp")
                Image4(Index).Tag = "1"
                Image4(1).Picture = Nothing

            End If

            If Image4(1).Tag = "1" Then
                Image4(1).Picture = Nothing
                Image4(1).Tag = "0"

            End If

        Case 1

            If Image4(Index).Tag = "0" Then
                Image4(Index).Picture = LoadInterface("boton-sm-cerrar-over.bmp")
                Image4(Index).Tag = "1"

            End If

            If Image4(0).Tag = "1" Then
                Image4(0).Picture = Nothing
                Image4(0).Tag = "0"

            End If

    End Select
         
    
    Exit Sub

Image4_MouseMove_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.Image4_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Image5_Click()
    
    On Error GoTo Image5_Click_Err
    

    If frmGrupo.visible = False Then
        Call WriteRequestGrupo

    End If

    
    Exit Sub

Image5_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.Image5_Click", Erl)
    Resume Next
    
End Sub


Private Sub Image6_Click()
    
    On Error GoTo Image6_Click_Err
    
    Call WriteSafeToggle

    
    Exit Sub

Image6_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.Image6_Click", Erl)
    Resume Next
    
End Sub

Private Sub imgBugReport_Click()
    
    On Error GoTo imgBugReport_Click_Err
    
    frmGmAyuda.Show vbModeless, frmMain

    
    Exit Sub

imgBugReport_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.imgBugReport_Click", Erl)
    Resume Next
    
End Sub

Private Sub imgHechizos_Click()
  Call hechizosClick
End Sub

Public Sub hechizosClick()
  
    On Error GoTo hechizosClick_Err
    

    If picHechiz.visible Then Exit Sub
    
    TempTick = GetTickCount And &H7FFFFFFF
    
    If TempTick - iClickTick < IntervaloEntreClicks And Not iClickTick = 0 And LastMacroButton <> tMacroButton.Hechizos Then
        Call WriteLogMacroClickHechizo(tMacro.Coordenadas)
    End If
    
    iClickTick = TempTick
    
    LastMacroButton = tMacroButton.Hechizos
    
    panel.Picture = LoadInterface("centrohechizo.bmp")
    
    
    If Seguido = 1 Then
        Call WriteNotifyInventarioHechizos(2, hlst.ListIndex, hlst.Scroll)
    End If
    
    picInv.visible = False
    
    picHechiz.visible = True

    cmdlanzar.visible = True

    imgSpellInfo.visible = True

    cmdMoverHechi(0).visible = True
    cmdMoverHechi(1).visible = True

    frmMain.imgInvLock(0).visible = False
    frmMain.imgInvLock(1).visible = False
    frmMain.imgInvLock(2).visible = False
    imgDeleteItem.visible = False
    
    Exit Sub

hechizosClick_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.hechizosClick", Erl)
    Resume Next
End Sub

Private Sub imgHechizos_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo imgHechizos_MouseDown_Err
       
    imgHechizos.Picture = LoadInterface("boton-hechizos-off.bmp")
    imgHechizos.Tag = "2"

    
    Exit Sub

imgHechizos_MouseDown_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.imgHechizos_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub imgHechizos_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo imgHechizos_MouseMove_Err
    
    
    If imgHechizos.Tag <> "1" And Button = 0 Then
        imgHechizos.Picture = LoadInterface("boton-hechizos-over.bmp")
        imgHechizos.Tag = "1"
        imgInventario.Picture = Nothing
        imgInventario.Tag = "0"
    End If

    
    Exit Sub

imgHechizos_MouseMove_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.imgHechizos_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub ImgHogar_Click()
    Call ParseUserCommand("/HOGAR")
End Sub


Private Sub imgInventario_Click()
    Call inventoryClick
End Sub

Public Sub inventoryClick()
    On Error GoTo inventoryClick_Err
    

    If picInv.visible Then Exit Sub

    TempTick = GetTickCount And &H7FFFFFFF
    
    If TempTick - iClickTick < IntervaloEntreClicks And Not iClickTick = 0 And LastMacroButton <> tMacroButton.Inventario Then
        Call WriteLogMacroClickHechizo(tMacro.Coordenadas)
    End If
    
    iClickTick = TempTick
    
    LastMacroButton = tMacroButton.Inventario

    panel.Picture = LoadInterface("centroinventario.bmp")
    'Call Audio.PlayWave(SND_CLICK)
    picInv.visible = True
    picHechiz.visible = False
    cmdlanzar.visible = False
    imgSpellInfo.visible = False
    If Seguido = 1 Then
        Call WriteNotifyInventarioHechizos(1, hlst.ListIndex, hlst.Scroll)
    End If

    cmdMoverHechi(0).visible = False
    cmdMoverHechi(1).visible = False
   ' Call Inventario.ReDraw
    frmMain.imgInvLock(0).visible = True
    frmMain.imgInvLock(1).visible = True
    frmMain.imgInvLock(2).visible = True
    imgDeleteItem.visible = True

    
    Exit Sub

inventoryClick_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.inventoryClick", Erl)
    Resume Next
End Sub

Private Sub imgInventario_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    

    On Error GoTo imgInventario_MouseDown_Err
    
    imgInventario.Picture = LoadInterface("boton-inventory-off.bmp")
    imgInventario.Tag = "2"

    'Call Inventario.DrawInventory
    
    Exit Sub

imgInventario_MouseDown_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.imgInventario_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub imgInventario_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo imgInventario_MouseMove_Err

    If imgInventario.Tag <> "1" And Button = 0 Then
        imgInventario.Picture = LoadInterface("boton-inventory-over.bmp")
        imgInventario.Tag = "1"
        imgHechizos.Picture = Nothing
        imgHechizos.Tag = "0"
    End If

    
    Exit Sub

imgInventario_MouseMove_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.imgInventario_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub imgOro_Click()
    
    On Error GoTo imgOro_Click_Err
    
    GldLbl_Click
    
    Exit Sub

imgOro_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.imgOro_Click", Erl)
    Resume Next
    
End Sub



Private Sub ImgSeg_Click()
    Call WriteSafeToggle
End Sub

Private Sub ImgSegClan_Click()
    Call WriteSeguroClan
End Sub

Private Sub ImgSegParty_Click()
    Call WriteParyToggle
End Sub

Private Sub ImgSegResu_Click()
    Call WriteSeguroResu
End Sub

Private Sub Label1_Click()
    frmBancoCuenta.Show , frmMain
End Sub

Private Sub lblPorcLvl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If ShowPercentage Then
        lblPorcLvl.visible = False
        exp.visible = True
    End If

End Sub

Private Sub LlamaDeclan_Timer()
    
    On Error GoTo LlamaDeclan_Timer_Err
    
    frmMapaGrande.llamadadeclan.visible = False
    frmMapaGrande.Shape2.visible = False
    LlamaDeclan.enabled = False

    
    Exit Sub

LlamaDeclan_Timer_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.LlamaDeclan_Timer", Erl)
    Resume Next
    
End Sub

Private Sub MANShp_Click()
    
    On Error GoTo MANShp_Click_Err
    
    manabar_Click
    
    Exit Sub

MANShp_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.MANShp_Click", Erl)
    Resume Next
    
End Sub



Private Sub OpcionesBoton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo OpcionesBoton_MouseDown_Err
    
    OpcionesBoton.Picture = LoadInterface("opcionesoverdown.bmp")
    OpcionesBoton.Tag = "1"

    Call frmOpciones.Init
    
    Exit Sub

OpcionesBoton_MouseDown_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.OpcionesBoton_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub panelGM_Click()
    
    On Error GoTo panelGM_Click_Err
    
    frmPanelgm.Width = 4860
    Call WriteSOSShowList
    Call WriteGMPanel
    
    Me.SetFocus
    
    Exit Sub

panelGM_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.panelGM_Click", Erl)
    Resume Next
    
End Sub

Private Sub Inventario_ItemDropped(ByVal Drag As Integer, ByVal Drop As Integer, ByVal x As Integer, ByVal y As Integer)
    
    On Error GoTo Inventario_ItemDropped_Err
    

    ' Si solt� un item en un slot v�lido
    
    If Drop > 0 Then
        If Drag <> Drop Then
            ' Muevo el item dentro del iventario
            Call WriteItemMove(Drag, Drop)
        End If
    End If

    
    Exit Sub

Inventario_ItemDropped_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.Inventario_ItemDropped", Erl)
    Resume Next
    
End Sub

Private Sub picInv_Paint()
    
    On Error GoTo picInv_Paint_Err
    
    Inventario.ReDraw
    
    Exit Sub

picInv_Paint_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.picInv_Paint", Erl)
    Resume Next
    
End Sub


Private Sub QuestBoton_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo QuestBoton_MouseUp_Err
    
    
    If pausa Then Exit Sub
    
    Call WriteQuestListRequest

    
    Exit Sub

QuestBoton_MouseUp_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.QuestBoton_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub Label6_Click()
    
    On Error GoTo Label6_Click_Err
    
    Inventario.SelectGold

    If UserGLD > 0 Then
        frmCantidad.Show , frmMain

    End If

    
    Exit Sub

Label6_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.Label6_Click", Erl)
    Resume Next
    
End Sub

Private Sub Label7_Click()
    
    On Error GoTo Label7_Click_Err
    
    
    Call AddToConsole("No tenes mensajes nuevos.", 255, 255, 255, False, False, False)

    
    Exit Sub

Label7_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.Label7_Click", Erl)
    Resume Next
    
End Sub

Private Sub lblPorcLvl_Click()
    
    On Error GoTo lblPorcLvl_Click_Err
    
    'Call WriteScrollInfo
    
    ShowPercentage = Not ShowPercentage
    
    Exit Sub

lblPorcLvl_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.lblPorcLvl_Click", Erl)
    Resume Next
    
End Sub

Private Sub MacroLadder_Timer()
    
    
    On Error GoTo MacroLadder_Timer_Err
    
    If pausa Then Exit Sub
    
    If UserMacro.cantidad > 0 And UserMacro.Activado And UserMinSTA > 0 Then
    
        Select Case UserMacro.TIPO

            Case 1 'Alquimia
                Call WriteCraftAlquimista(UserMacro.Index)
                UserMacro.cantidad = UserMacro.cantidad - 1

            Case 3 'Sasteria
                Call WriteCraftSastre(UserMacro.Index)
                UserMacro.cantidad = UserMacro.cantidad - 1

            Case 4 'Herreria
                Call WriteCraftBlacksmith(UserMacro.Index)
                UserMacro.cantidad = UserMacro.cantidad - 1

            Case 6
                ' Jopi: Esto se hace desde el servidor
                'Call WriteWorkLeftClick(TargetXMacro, TargetYMacro, UsingSkill)

        End Select

    Else
        Call ResetearUserMacro

    End If

    'UserMacro.cantidad = UserMacro.cantidad - 1
    
    Exit Sub

MacroLadder_Timer_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.MacroLadder_Timer", Erl)
    Resume Next
    
End Sub

Private Sub macrotrabajo_Timer()
    'If Inventario.SelectedItem = 0 Then
    '   DesactivarMacroTrabajo
    '   Exit Sub
    'End If
    
    On Error GoTo macrotrabajo_Timer_Err
    
    
    If pausa Then Exit Sub
    
    
    'Macros are disabled if not using Argentum!
    If Not Application.IsAppActive() Then
        DesactivarMacroTrabajo
        Exit Sub

    End If
    
    If (UsingSkill = eSkill.Herreria And Not frmHerrero.visible) Then
        Call WriteWorkLeftClick(TargetXMacro, TargetYMacro, UsingSkill)
        UsingSkill = 0

    End If
    
    'If Inventario.OBJType(Inventario.SelectedItem) = eObjType.otWeapon Then
    If Not (frmCarp.visible = True) Then
        If frmMain.Inventario.IsItemSelected Then Call WriteUseItem(frmMain.Inventario.SelectedItem)
    End If

    
    Exit Sub

macrotrabajo_Timer_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.macrotrabajo_Timer", Erl)
    Resume Next
    
End Sub

Public Sub ActivarMacroTrabajo()
    
    On Error GoTo ActivarMacroTrabajo_Err
    
    TargetXMacro = tX
    TargetYMacro = tY
    macrotrabajo.Interval = IntervaloTrabajoConstruir
    macrotrabajo.enabled = True
    Call AddToConsole("Macro Trabajo ACTIVADO", 0, 200, 200, False, True, False)

    
    Exit Sub

ActivarMacroTrabajo_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.ActivarMacroTrabajo", Erl)
    Resume Next
    
End Sub

Public Sub DesactivarMacroTrabajo()
    
    On Error GoTo DesactivarMacroTrabajo_Err
    
    TargetXMacro = 0
    TargetYMacro = 0
    macrotrabajo.enabled = False
    MacroBltIndex = 0
    UsingSkill = 0
    MousePointer = vbDefault
    Call AddToConsole("Macro Trabajo DESACTIVADO", 0, 200, 200, False, True, False)

    
    Exit Sub

DesactivarMacroTrabajo_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.DesactivarMacroTrabajo", Erl)
    Resume Next
    
End Sub

Private Sub MenuOpciones_Click()

End Sub

Private Sub manabar_Click()
    
    On Error GoTo manabar_Click_Err
        
    Call WriteMeditate

    
    Exit Sub

manabar_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.manabar_Click", Erl)
    Resume Next
    
End Sub


Private Sub mnuEquipar_Click()
    
    On Error GoTo mnuEquipar_Click_Err
    

    If MainTimer.Check(TimersIndex.UseItemWithU) Then Call WriteEquipItem(Inventario.SelectedItem)

    
    Exit Sub

mnuEquipar_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.mnuEquipar_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuNPCComerciar_Click()
    
    On Error GoTo mnuNPCComerciar_Click_Err
    
    Call WriteLeftClick(tX, tY)
    Call WriteCommerceStart

    
    Exit Sub

mnuNPCComerciar_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.mnuNPCComerciar_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuNpcDesc_Click()
    
    On Error GoTo mnuNpcDesc_Click_Err
    
    Call WriteLeftClick(tX, tY)

    
    Exit Sub

mnuNpcDesc_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.mnuNpcDesc_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuTirar_Click()
    
    On Error GoTo mnuTirar_Click_Err
    
    Call TirarItem

    
    Exit Sub

mnuTirar_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.mnuTirar_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuUsar_Click()
    
    On Error GoTo mnuUsar_Click_Err
    

    If frmMain.Inventario.IsItemSelected Then Call WriteUseItem(frmMain.Inventario.SelectedItem)

    
    Exit Sub

mnuUsar_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.mnuUsar_Click", Erl)
    Resume Next
    
End Sub


Private Sub onlines_Click()
    
    On Error GoTo onlines_Click_Err
    
    Call WriteOnline

    
    Exit Sub

onlines_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.onlines_Click", Erl)
    Resume Next
    
End Sub

Private Sub OpcionesBoton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo OpcionesBoton_MouseMove_Err
    

    If OpcionesBoton.Tag = "0" Then
        OpcionesBoton.Picture = LoadInterface("opcionesover.bmp")
        OpcionesBoton.Tag = "1"

    End If

    
    Exit Sub

OpcionesBoton_MouseMove_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.OpcionesBoton_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub OpcionesBoton_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo OpcionesBoton_MouseUp_Err
    
    Call frmOpciones.Init

    
    Exit Sub

OpcionesBoton_MouseUp_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.OpcionesBoton_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub Panel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Panel_MouseMove_Err
    

    ObjLbl.visible = False

    If cmdlanzar.Tag = "1" Then
        cmdlanzar.Picture = Nothing
        cmdlanzar.Tag = "0"

    End If
    
    If imgInventario.Tag <> "0" Then
        imgInventario.Picture = Nothing
        imgInventario.Tag = "0"

    End If

    If imgHechizos.Tag <> "0" Then
        imgHechizos.Picture = Nothing
        imgHechizos.Tag = "0"

    End If

    If cmdMoverHechi(1).Tag = "1" Then
        cmdMoverHechi(1).Picture = Nothing
        cmdMoverHechi(1).Tag = "0"

    End If
    
    If cmdMoverHechi(0).Tag = "1" Then
        cmdMoverHechi(0).Picture = Nothing
        cmdMoverHechi(0).Tag = "0"

    End If
    
    If imgSpellInfo.Tag = "1" Then
        imgSpellInfo.Picture = Nothing
        imgSpellInfo.Tag = "0"
    End If

    
    Exit Sub

Panel_MouseMove_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.Panel_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub picInv_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo picInv_MouseMove_Err
    

    Dim Slot As Byte

    UsaMacro = False
    
    Slot = Inventario.GetSlot(x, y)
    
    If Slot <= 0 Then
        ObjLbl.visible = False
        Exit Sub
    End If
    
    If Inventario.Amount(Slot) > 0 Then
    
        ObjLbl.visible = True
        
        Select Case ObjData(Inventario.ObjIndex(Slot)).ObjType

            Case eObjType.otWeapon
                ObjLbl = Inventario.ItemName(Slot) & " (" & Inventario.Amount(Slot) & ")" & vbCrLf & "Da�o: " & ObjData(Inventario.ObjIndex(Slot)).MinHit & "/" & ObjData(Inventario.ObjIndex(Slot)).MaxHit

            Case eObjType.otArmadura
                ObjLbl = Inventario.ItemName(Slot) & " (" & Inventario.Amount(Slot) & ")" & vbCrLf & "Defensa: " & ObjData(Inventario.ObjIndex(Slot)).MinDef & "/" & ObjData(Inventario.ObjIndex(Slot)).MaxDef

            Case eObjType.otCASCO
                ObjLbl = Inventario.ItemName(Slot) & " (" & Inventario.Amount(Slot) & ")" & vbCrLf & "Defensa: " & ObjData(Inventario.ObjIndex(Slot)).MinDef & "/" & ObjData(Inventario.ObjIndex(Slot)).MaxDef

            Case eObjType.otESCUDO
                ObjLbl = Inventario.ItemName(Slot) & " (" & Inventario.Amount(Slot) & ")" & vbCrLf & "Defensa: " & ObjData(Inventario.ObjIndex(Slot)).MinDef & "/" & ObjData(Inventario.ObjIndex(Slot)).MaxDef

            Case Else
                ObjLbl = Inventario.ItemName(Slot) & " (" & Inventario.Amount(Slot) & ")" & vbCrLf & ObjData(Inventario.ObjIndex(Slot)).Texto

        End Select
        
        If Len(ObjLbl.Caption) < 100 Then
            ObjLbl.FontSize = 7
            
        ElseIf Len(ObjLbl.Caption) > 100 And Len(ObjLbl.Caption) < 150 Then
            ObjLbl.FontSize = 6

            '
            ' Else
            '  ObjLbl.FontSize = 5
        End If

    Else
        ObjLbl.visible = False

    End If

    
    Exit Sub

picInv_MouseMove_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.picInv_MouseMove", Erl)
    Resume Next
    
End Sub
Private Sub CompletarEnvioMensajes()
    
    On Error GoTo CompletarEnvioMensajes_Err
    

    Select Case SendingType

        Case 1
            SendTxt.Text = ""

        Case 2
            SendTxt.Text = "-"

        Case 3
            SendTxt.Text = ("\" & sndPrivateTo & " ")

        Case 5
            SendTxt.Text = "/GRUPO "

        Case 6
            SendTxt.Text = "/GRMG "

        Case 7
            SendTxt.Text = ";"

        Case 8
            SendTxt.Text = "/RMSG "

    End Select

    stxtbuffer = SendTxt.Text
    SendTxt.SelStart = Len(SendTxt.Text)

    
    Exit Sub

CompletarEnvioMensajes_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.CompletarEnvioMensajes", Erl)
    Resume Next
    
End Sub


Private Sub refuerzolanzar_Click()
    
    On Error GoTo refuerzolanzar_Click_Err
    
    Call cmdLanzar_Click

    
    Exit Sub

refuerzolanzar_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.refuerzolanzar_Click", Erl)
    Resume Next
    
End Sub

Private Sub refuerzolanzar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo refuerzolanzar_MouseMove_Err
    
    UsaMacro = False
    CnTd = 0

    If cmdlanzar.Tag = "0" Then
        cmdlanzar.Picture = LoadInterface("lanzarmarcado.bmp")
        cmdlanzar.Tag = "1"

    End If

    
    Exit Sub

refuerzolanzar_MouseMove_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.refuerzolanzar_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub renderer_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo renderer_MouseUp_Err
    

    'If DropItem Then
    '    frmMain.UsandoDrag = False
    '    DropItem = False
    '    DropIndex = 0
    '    DropActivo = False
    '    Call FormParser.Parse_Form(Me)
    'End If
    
    If Minimap.MouseUp(Button, MouseX, MouseY) Then Exit Sub

    clicX = x
    clicY = y
    If Button = vbLeftButton Then
        If ConsoleDialog.menu_consola_visible And ConsoleDialog.cambiando_intensidad Then
            ConsoleDialog.cambiando_intensidad = False
        End If
        If Pregunta Then
            If x >= 395 And x <= 412 And y >= 243 And y <= 260 Then
                If PreguntaLocal Then
    
                    Select Case PreguntaNUM
    
                        Case 1
                            Pregunta = False
                            DestItemSlot = 0
                            DestItemCant = 0
                            PreguntaLocal = False
                            
                        Case 2 ' Denunciar
                            Pregunta = False
                            PreguntaLocal = False
    
                    End Select
    
                Else
                    Call WriteResponderPregunta(False)
                    Pregunta = False
    
                End If
                
                Exit Sub
    
            ElseIf x >= 420 And x <= 436 And y >= 243 And y <= 260 Then
                If PreguntaLocal Then
    
                    Select Case PreguntaNUM
    
                        Case 1 '�Destruir item?
                            Call WriteDrop(DestItemSlot, DestItemCant)
                            Pregunta = False
                            PreguntaLocal = False
                            
                        Case 2 ' Denunciar
                            Call WriteDenounce(TargetName)
                            Pregunta = False
                            PreguntaLocal = False
    
                    End Select
    
                Else
                    Call WriteResponderPregunta(True)
                    Pregunta = False
    
                End If
                
                Exit Sub
    
            End If
    
        End If
    
    ElseIf Button = vbRightButton Then
        
        Dim charindex As Integer
        charindex = MapData(rrX(tX), rrY(tY)).charindex
        
        If charindex = 0 Then
            charindex = MapData(rrX(tX), rrY(tY + 1)).charindex
        End If
        
        If charindex <> 0 And charindex <> UserCharIndex Then
            Dim Frm As Form
            
            Call WriteLeftClick(tX, tY)

            TargetX = tX
            TargetY = tY
        
            If charlist(charindex).EsMascota Then
                Set Frm = frmMenuNPC
            
            ElseIf Not charlist(charindex).esNpc Then
                
                TargetName = charlist(charindex).nombre
                
                If charlist(UserCharIndex).priv > 0 And Shift = 0 Then
                    Set Frm = frmMenuGM
                Else
                    Set Frm = frmMenuUser
                End If
            End If
            
            If Not Frm Is Nothing Then
                Call Frm.Show
            
                Frm.Left = Me.Left + (renderer.Left + x + 1) * Screen.TwipsPerPixelX
                    
                If (renderer.Top + y) * Screen.TwipsPerPixelY + Frm.Height > Me.Height Then
                    Frm.Top = Me.Top + (renderer.Top + y) * Screen.TwipsPerPixelY - Frm.Height
                Else
                    Frm.Top = Me.Top + (renderer.Top + y) * Screen.TwipsPerPixelY
                End If
                
                Set Frm = Nothing
            End If
        End If

    End If
    
    
    Exit Sub

renderer_MouseUp_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.renderer_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub renderer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo renderer_MouseMove_Err

    DisableURLDetect

    Call Form_MouseMove(Button, Shift, renderer.Left + x, renderer.Top + y)
    'If DropItem Then

    ' frmMain.UsandoDrag = False
    ' Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
    'Call WriteDropItem(DropIndex, tX, tY, CantidadDrop)
    ' DropItem = False
    ' DropIndex = 0
    ' TimeDrop = 0
    ' DropActivo = False
    ' CantidadDrop = 0
    ' Call FormParser.Parse_Form(frmMain)
    
    ' End If
    
    'LucesCuadradas.Light_Remove (10)
    
    'LucesCuadradas.Light_Create tX, tY, &HFFFFFFF, 1, 10
    'LucesCuadradas.Light_Render_All
    If ConsoleDialog.menu_consola_visible And ConsoleDialog.cambiando_intensidad Then
        If x >= 512 And x <= 724 Then
            
            If x <= 563 Then
                
                ConsoleDialog.consoleAlpha_min = x
                Dim minAlpha As Single
                minAlpha = ((x - 512) * 100) / 212
                ConsoleDialog.consoleAlpha_min = minAlpha * 2.55
                ConsoleDialog.consoleAlpha_min_pos = x
            ElseIf x > 564 Then
                Dim maxAlpha As Single
                maxAlpha = ((x - 564 + 52) * 100) / 212
                ConsoleDialog.consoleAlpha_max = maxAlpha * 2.55
                ConsoleDialog.consoleAlpha_max_pos = x
            End If
        End If
    End If
    
    Exit Sub

renderer_MouseMove_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.renderer_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub renderer_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo renderer_MouseDown_Err
    
    If SendTxt.visible Then SendTxt.SetFocus
    MouseBoton = Button
    MouseShift = Shift
    If frmComerciar.visible Then Unload frmComerciar
    If frmBancoObj.visible Then Unload frmBancoObj
    If frmEstadisticas.visible Then Unload frmEstadisticas
    If frmGoliath.visible Then Unload frmGoliath
    If frmMapaGrande.visible Then frmMapaGrande.visible = False
    If frmViajes.visible Then Unload frmViajes
    If frmCantidad.visible Then Unload frmCantidad
    If frmGrupo.visible Then Unload frmGrupo
    If frmGmAyuda.visible Then Unload frmGmAyuda
    If frmGuildAdm.visible Then Unload frmGuildAdm
    If frmHerrero.visible Then Unload frmHerrero
    If frmSastre.visible Then Unload frmSastre
    If frmAlqui.visible Then Unload frmAlqui
    If frmCarp.visible Then Unload frmCarp
    If frmMenuUser.visible Then Unload frmMenuUser
    If frmMenuGM.visible Then Unload frmMenuGM
    If frmMenuNPC.visible Then Unload frmMenuNPC
    
    If ConsoleDialog.menu_consola_visible Then
        If x >= 512 And x <= 723 And y >= 115 And y <= 130 Then
            ConsoleDialog.cambiando_intensidad = True
             If x <= 563 Then
                ConsoleDialog.consoleAlpha_min = x
                Dim minAlpha As Single
                minAlpha = ((x - 512) * 100) / 212
                ConsoleDialog.consoleAlpha_min = minAlpha * 2.55
                ConsoleDialog.consoleAlpha_min_pos = x
            ElseIf x > 564 Then
                Dim maxAlpha As Single
                maxAlpha = ((x - 564 + 52) * 100) / 212
                ConsoleDialog.consoleAlpha_max = maxAlpha * 2.55
                ConsoleDialog.consoleAlpha_max_pos = x
            End If
        End If
    End If
    
    Call Minimap.MouseDown(Button, MouseX, MouseY)
    
    Exit Sub

renderer_MouseDown_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.renderer_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub renderer_DblClick()
    
    On Error GoTo renderer_DblClick_Err
    
    Call Form_DblClick

    
    Exit Sub

renderer_DblClick_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.renderer_DblClick", Erl)
    Resume Next
    
End Sub

Private Sub renderer_Click()
    'Call addCooldown(713, 15000)
    On Error GoTo renderer_Click_Err
    Call Form_Click
    If SendTxt.visible Then SendTxt.SetFocus
    If SendTxtCmsg.visible Then SendTxtCmsg.SetFocus
    Exit Sub

renderer_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.renderer_Click", Erl)
    Resume Next
End Sub

Private Sub Retar_Click()
    Call ParseUserCommand("/RETAR")
End Sub


Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo SendTxt_KeyUp_Err

    Dim str1 As String

    Dim str2 As String

    'Send text
    If KeyCode = vbKeyReturn Then
        
        If LenB(stxtbuffer) <> 0 Then
        
            ' If Right$(stxtbuffer, 1) = " " Or left(stxtbuffer, 1) = " " Then
            ' stxtbuffer = Trim(stxtbuffer)
            ' End If
        
            If Left$(stxtbuffer, 1) = "/" Then
                If UCase$(Left$(stxtbuffer, 7)) = "/GRUPO " Then
                    SendingType = 5
                ElseIf UCase$(Left$(stxtbuffer, 6)) = "/CMSG " Then
                    SendingType = 4
                ElseIf UCase$(Left$(stxtbuffer, 6)) = "/GRMG " Then
                    SendingType = 6
                ElseIf UCase$(Left$(stxtbuffer, 6)) = "/RMSG " Then
                    SendingType = 8
                Else
                    SendingType = 1
                End If
            
                If stxtbuffer <> "" Then Call ParseUserCommand(stxtbuffer)
    
                'Shout
            ElseIf Left$(stxtbuffer, 1) = "-" Then

                If Right$(stxtbuffer, Len(stxtbuffer) - 1) <> "" Then Call ParseUserCommand("-" & Right$(stxtbuffer, Len(stxtbuffer) - 1))
                SendingType = 2
            
                'Global
            ElseIf Left$(stxtbuffer, 1) = ";" Then

                If Right$(stxtbuffer, Len(stxtbuffer) - 1) <> "" Then Call ParseUserCommand("/CONSOLA " & Right$(stxtbuffer, Len(stxtbuffer) - 1))
                sndPrivateTo = ""
            
            ElseIf Left$(stxtbuffer, 1) = "/RMSG" Then

                If Right$(stxtbuffer, Len(stxtbuffer) - 1) <> "" Then Call ParseUserCommand("/RMSG " & Right$(stxtbuffer, Len(stxtbuffer) - 1))
                SendingType = 8
                sndPrivateTo = ""

                'Privado
            ElseIf Left$(stxtbuffer, 1) = "\" Then

                Dim mensaje As String
 
                str1 = Right$(stxtbuffer, Len(stxtbuffer) - 1)
                str2 = ReadField(1, str1, 32)
                mensaje = Right$(stxtbuffer, Len(str1) - Len(str2) - 1)
                sndPrivateTo = str2
                SendingType = 3
    
                If str1 <> "" Then Call WriteWhisper(sndPrivateTo, mensaje)
                    
                'Say
            Else

                If stxtbuffer <> "" Then Call ParseUserCommand(stxtbuffer)
                SendingType = 1
                sndPrivateTo = ""

            End If
            
            

        Else
            SendingType = 1
            sndPrivateTo = ""

        End If
        stxtbuffer = ""
        SendTxt.Text = ""
        KeyCode = 0
        
        
3        Dim tiempoTranscurridoCartel As Double
        
        tiempoTranscurridoCartel = GetTickCount - StartOpenChatTime

        Call computeLastElapsedTimeChat(tiempoTranscurridoCartel)
        
        
        
        SendTxt.visible = False
        
    End If

    Exit Sub

SendTxt_KeyUp_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.SendTxt_KeyUp", Erl)
    Resume Next
    
End Sub

Private Sub computeLastElapsedTimeChat(ByVal tiempoTranscurridoCartel As Double)
    Dim i As Long
    
    For i = 2 To 6
        LastElapsedTimeChat(i - 1) = LastElapsedTimeChat(i)
    Next i
    
    LastElapsedTimeChat(6) = tiempoTranscurridoCartel
        
    'HarThaoS: Calculo el m�nimo y m�ximo de mis carteleos
    Dim min As Double, max As Double
    
    min = LastElapsedTimeChat(6)
    max = LastElapsedTimeChat(6)
    
    For i = 1 To 6
        If LastElapsedTimeChat(i) > max Then max = LastElapsedTimeChat(i)
        If LastElapsedTimeChat(i) < min Then min = LastElapsedTimeChat(i)
    Next i
    
    If (max - min) > 0 And (max - min) < 12 Then
         Call WriteLogMacroClickHechizo(tMacro.borrarCartel)
    End If
    
    
    
End Sub

Private Sub SendTxtCmsg_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    If SendTxtCmsg.SelStart > 2 Then Call ParseUserCommand("/CMSG " & SendTxtCmsg.Text)
    SendTxtCmsg.visible = False
    SendTxtCmsg.Text = ""
  End If
End Sub

Private Sub ShowFPS_Timer()
    
    On Error GoTo ShowFPS_Timer_Err
    
    fps.Caption = "FPS: " & engine.fps
    
    Exit Sub

ShowFPS_Timer_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.ShowFPS_Timer", Erl)
    Resume Next
    
End Sub

Private Sub cerrarcuenta_Timer()
    
    On Error GoTo cerrarcuenta_Timer_Err
    
    Unload frmConnect
    Unload frmCrearPersonaje
    cerrarcuenta.enabled = False

    
    Exit Sub

cerrarcuenta_Timer_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.cerrarcuenta_Timer", Erl)
    Resume Next
    
End Sub

Private Sub timerCooldownCombo_Timer()
    If GetTickCount - StartComboCooldownTime > ComboCooldownTime Then
        shpCooldownComboBlock.Width = 1
    Else
        shpCooldownComboBlock.Width = 243 - (GetTickCount - StartComboCooldownTime) / ComboCooldownTime * 243
    End If
End Sub

Private Sub TimerLluvia_Timer()
    
    On Error GoTo TimerLluvia_Timer_Err
    

    If bRain Then

        If CantPartLLuvia < 250 Then

            CantPartLLuvia = CantPartLLuvia + 1
            Graficos_Particulas.Particle_Group_Edit (MeteoIndex)
        Else
            CantPartLLuvia = 250
            TimerLluvia.enabled = False

        End If

    Else

        If CantPartLLuvia > 0 Then
            CantPartLLuvia = CantPartLLuvia - 1
            Graficos_Particulas.Particle_Group_Edit (MeteoIndex)
        Else
    
            Call Graficos_Particulas.Engine_MeteoParticle_Set(-1)
            CantPartLLuvia = 0
            TimerLluvia.enabled = False

        End If

    End If

    
    Exit Sub

TimerLluvia_Timer_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.TimerLluvia_Timer", Erl)
    Resume Next
    
End Sub

Private Sub TimerMusica_Timer()

End Sub

Private Sub TimerNiebla_Timer()
    
    On Error GoTo TimerNiebla_Timer_Err
    

    If bNiebla Then

        If AlphaNiebla < MaxAlphaNiebla Then
            AlphaNiebla = AlphaNiebla + 1
        Else
            AlphaNiebla = MaxAlphaNiebla
            TimerNiebla.enabled = False

        End If

    Else

        If AlphaNiebla > 0 Then
            AlphaNiebla = AlphaNiebla - 1

        Else
            AlphaNiebla = 0
            MaxAlphaNiebla = 0
            TimerNiebla.enabled = False
        End If

    End If

    
    Exit Sub

TimerNiebla_Timer_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.TimerNiebla_Timer", Erl)
    Resume Next
    
End Sub


Private Sub cmdLanzar_Click()
    
    On Error GoTo cmdLanzar_Click_Err
    
    If pausa Then Exit Sub

    TempTick = GetTickCount And &H7FFFFFFF
    
    If TempTick - iClickTick < IntervaloEntreClicks And Not iClickTick = 0 And LastMacroButton <> tMacroButton.Lanzar Then
        
        Call WriteLogMacroClickHechizo(tMacro.Coordenadas)
    End If
    
    iClickTick = TempTick
    
    LastMacroButton = tMacroButton.Lanzar

    If hlst.List(hlst.ListIndex) <> "(Vac�o)" Then
        If UserEstado = 1 Then

            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("�Est�s muerto!", .red, .green, .blue, .bold, .italic)

            End With

        Else
        
            If ModoHechizos = BloqueoLanzar Then
                If Not MainTimer.Check(TimersIndex.AttackSpell, False) Or Not MainTimer.Check(TimersIndex.CastSpell, False) Then
                    Exit Sub
                End If
            End If

            
            Call WriteCastSpell(hlst.ListIndex + 1)
            'Call WriteWork(eSkill.Magia)
            UsaMacro = True
            UsaLanzar = True

        End If

    End If

    
    Exit Sub

cmdLanzar_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.cmdLanzar_Click", Erl)
    Resume Next
    
End Sub

Private Sub CmdLanzar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo CmdLanzar_MouseMove_Err
    
    UsaMacro = False
    CnTd = 0

    If cmdlanzar.Tag = "0" Then
        cmdlanzar.Picture = LoadInterface("boton-lanzar-over.bmp")
        cmdlanzar.Tag = "1"

    End If

    
    Exit Sub

CmdLanzar_MouseMove_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.CmdLanzar_MouseMove", Erl)
    Resume Next
    
End Sub

Public Sub Form_Click()
    
    On Error GoTo Form_Click_Err

    If pausa Then Exit Sub
        
    If mascota.visible Then
        If Sqr((MouseX - mascota.PosX) ^ 2 + (MouseY - mascota.PosY) ^ 2) < 30 Then
            mascota.dialog = ""
        End If
    End If

    If ConsoleDialog.MouseClick(MouseX, MouseY) Then
        frmMain.SendTxt.Top = frmMain.renderer.Top + ConsoleDialog.console_height
        frmMain.SendTxtCmsg.Top = frmMain.renderer.Top + ConsoleDialog.console_height
    End If
    
    If cartel_visible Then
        If MouseX > 50 And MouseY > 478 And MouseX < 671 And MouseY < 585 Then
            'Debug.Print tutorial_texto_actual
            If tutorial_index > 0 Then
                Call nextCartel
            Else
                Call cerrarCartel
            End If
        End If
    End If
    If MouseBoton = vbLeftButton And ACCION1 = 0 Or MouseBoton = vbRightButton And ACCION2 = 0 Or MouseBoton = 4 And ACCION3 = 0 Then
        If Not Comerciando Then
            ' Fix: game area esta mal
            'If Not InGameArea() Then Exit Sub

            If MouseShift = 0 Then
                If UsingSkill = 0 Or MacroLadder.enabled Then
                    Call CountPacketIterations(packetControl(ClientPacketID.LeftClick), 150)
                    'Debug.Print "click"
                    Call WriteLeftClick(tX, tY)
                Else

                    'If macrotrabajo.Enabled Then DesactivarMacroTrabajo
                    
                    Dim SendSkill As Boolean
                    
                    If UsingSkill = magia Then
                        
                        If ModoHechizos = BloqueoLanzar Then
                            SendSkill = IIf((MouseX >= renderer.ScaleLeft And MouseX <= MainViewWidth + renderer.ScaleLeft And MouseY >= renderer.ScaleTop And MouseY <= renderer.ScaleTop + MainViewHeight), True, False)
                            
                            If Not SendSkill Then
                                Exit Sub
                            End If
                            
                            Call MainTimer.Restart(TimersIndex.CastAttack)
                            Call MainTimer.Restart(TimersIndex.CastSpell)
                        Else
                            If MainTimer.Check(TimersIndex.AttackSpell, False) Then
                                If MainTimer.Check(TimersIndex.CastSpell) Then
                                    SendSkill = IIf((MouseX >= renderer.ScaleLeft And MouseX <= MainViewWidth + renderer.ScaleLeft And MouseY >= renderer.ScaleTop And MouseY <= renderer.ScaleTop + MainViewHeight), True, False)
                                    
                                    If Not SendSkill Then
                                        Exit Sub
                                    End If
                                    
                                '    Set cooldown_hechizo = New clsCooldown
                                '    Call cooldown_hechizo.Cooldown_Initialize(IntervaloMagia, 26018)
                                '    Call addCooldown(cooldown_hechizo)
                                    Call MainTimer.Restart(TimersIndex.CastAttack)
                                
                                ElseIf ModoHechizos = SinBloqueo Then
                                    SendSkill = IIf((MouseX >= renderer.ScaleLeft And MouseX <= MainViewWidth + renderer.ScaleLeft And MouseY >= renderer.ScaleTop And MouseY <= renderer.ScaleTop + MainViewHeight), True, False)
                                    
                                    If Not SendSkill Then
                                        Exit Sub
                                    End If
                                
                                    With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                        Call ShowConsoleMsg("No puedes lanzar hechizos tan r�pido.", .red, .green, .blue, .bold, .italic)
                                    End With
                                Else
                                    Exit Sub
                                End If
                                
                            ElseIf ModoHechizos = SinBloqueo Then
                                SendSkill = IIf((MouseX >= renderer.ScaleLeft And MouseX <= MainViewWidth + renderer.ScaleLeft And MouseY >= renderer.ScaleTop And MouseY <= renderer.ScaleTop + MainViewHeight), True, False)
                                    
                                If Not SendSkill Then
                                    Exit Sub
                                End If
                                
                                With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                    Call ShowConsoleMsg("No puedes lanzar tan r�pido despu�s de un golpe.", .red, .green, .blue, .bold, .italic)
                                End With
                            Else
                                Exit Sub
                            End If
                        End If

                    End If
                    
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Proyectiles Then
                        If MainTimer.Check(TimersIndex.AttackSpell, False) Then
                            If MainTimer.Check(TimersIndex.CastAttack, False) Then
                                If MainTimer.Check(TimersIndex.Arrows) Then
                                    SendSkill = True
                                    Call MainTimer.Restart(TimersIndex.Attack) ' Prevengo flecha-golpe
                                    Call MainTimer.Restart(TimersIndex.CastSpell) ' flecha-hechizo

                                End If

                            End If

                        End If

                    End If
                
                    'Splitted because VB isn't lazy!
                    If (UsingSkill = Robar Or UsingSkill = Domar Or UsingSkill = Grupo Or UsingSkill = MarcaDeClan Or UsingSkill = MarcaDeGM) Then
                        If MainTimer.Check(TimersIndex.CastSpell) Then
                            If UsingSkill = MarcaDeGM Then

                                Dim Pos As Integer

                                If MapData(rrX(tX), rrY(tY)).charindex <> 0 Then
                                    Pos = InStr(charlist(MapData(rrX(tX), rrY(tY)).charindex).nombre, "<")
                                
                                    If Pos = 0 Then Pos = LenB(charlist(MapData(rrX(tX), rrY(tY)).charindex).nombre) + 2
                                    frmPanelgm.cboListaUsus.Text = Left$(charlist(MapData(rrX(tX), rrY(tY)).charindex).nombre, Pos - 2)

                                End If

                            Else
                                SendSkill = True

                            End If

                        End If

                    End If
                    
                    If (UsingSkill = eSkill.Pescar Or UsingSkill = eSkill.Talar Or UsingSkill = eSkill.Mineria Or UsingSkill = eSkill.Herreria) Then
                        
                        If MainTimer.Check(TimersIndex.CastSpell) Then
                            Call WriteWorkLeftClick(tX, tY, UsingSkill)
                            Call FormParser.Parse_Form(frmMain)

                            If CursoresGraficos = 0 Then
                                frmMain.MousePointer = vbDefault

                            End If

                            Exit Sub

                        End If

                    End If
                   
                    If SendSkill Then
                        If UsingSkill = eSkill.magia Then
                            If ComprobarPosibleMacro(MouseX, MouseY) Then
                                Call WriteWorkLeftClick(tX + RandomNumber(-2, 2), tY + RandomNumber(-2, 2), UsingSkill)
                            Else
                                Call WriteWorkLeftClick(tX, tY, UsingSkill)
                            End If
                        Else
                            Call WriteWorkLeftClick(tX, tY, UsingSkill)
                        End If

                    End If

                    Call FormParser.Parse_Form(frmMain)

                    If CursoresGraficos = 0 Then
                        frmMain.MousePointer = vbDefault

                    End If
                    
                    UsaLanzar = False
                    UsingSkill = 0

                End If

            Else
                If minimapa_visible Then
                    Call WriteWarpChar("YO", UserMap, MouseMapX, MouseMapY)
                Else
                    Call WriteWarpChar("YO", UserMap, tX, tY)
                End If
            End If
            
            If cartel Then cartel = False
            
        End If
    
    ElseIf MouseBoton = vbLeftButton And ACCION1 = 1 Or MouseBoton = vbRightButton And ACCION2 = 1 Or MouseBoton = 4 And ACCION3 = 1 Then
        'Call WriteDoubleClick(tX, tY)
    
    ElseIf MouseBoton = vbLeftButton And ACCION1 = 2 Or MouseBoton = vbRightButton And ACCION2 = 2 Or MouseBoton = 4 And ACCION3 = 2 Then

        If UserDescansar Or UserMeditar Then Exit Sub
        If MainTimer.Check(TimersIndex.CastAttack, False) Then
            If MainTimer.Check(TimersIndex.Attack) Then
                Call MainTimer.Restart(TimersIndex.AttackSpell)
                Call WriteAttack

            End If

        End If
    
    ElseIf MouseBoton = vbLeftButton And ACCION1 = 3 Or MouseBoton = vbRightButton And ACCION2 = 3 Or MouseBoton = 4 And ACCION3 = 3 Then

            If frmMain.Inventario.IsItemSelected Then Call WriteUseItem(frmMain.Inventario.SelectedItem)

    
    ElseIf MouseBoton = vbLeftButton And ACCION1 = 4 Or MouseBoton = vbRightButton And ACCION2 = 4 Or MouseBoton = 4 And ACCION3 = 4 Then

        If MapData(rrX(tX), rrY(tY)).charindex <> 0 Then
            If charlist(MapData(rrX(tX), rrY(tY)).charindex).nombre <> charlist(MapData(rrX(UserPos.x), rrY(UserPos.y)).charindex).nombre Then
                If charlist(MapData(rrX(tX), rrY(tY)).charindex).esNpc = False Then
                    SendTxt.Text = "\" & charlist(MapData(rrX(tX), rrY(tY)).charindex).nombre & " "

                 
                    If SendTxtCmsg.visible = False Then
                        SendTxt.visible = True
                        SendTxt.SetFocus
                        SendTxt.SelStart = Len(SendTxt.Text)
                    End If
                End If

            End If

        End If

    End If

    
    Exit Sub

Form_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.Form_Click", Erl)
    Resume Next
    
End Sub

Private Sub Form_DblClick()
    
    On Error GoTo Form_DblClick_Err
    

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 12/27/2007
    '12/28/2007: ByVal - Chequea que la ventana de comercio y boveda no este abierta al hacer doble clic a un comerciante, sobrecarga la lista de items.
    '**************************************************************
    If Not frmComerciar.visible And Not frmBancoObj.visible Then
        If MouseBoton = vbLeftButton Then

            Call WriteDoubleClick(tX, tY)
        End If

    End If

    
    Exit Sub

Form_DblClick_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.Form_DblClick", Erl)
    Resume Next
    
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err

    Call FormParser.Parse_Form(frmMain)
    
    MenuNivel = 1
    Me.Caption = "ArgentumWorld"
    
    Set hlst = New clsGraphicalList
    Call hlst.Initialize(Me.picHechiz, RGB(200, 190, 190))
    
    If Not RunningInVB Then
        Call WheelHook(frmMain.hWnd)
    End If
    
    loadButtons
    
    frmMain.SendTxt.Top = frmMain.renderer.Top + ConsoleDialog.console_height
    frmMain.SendTxtCmsg.Top = frmMain.renderer.Top + ConsoleDialog.console_height
    Exit Sub

Form_Load_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Form_MouseMove_Err

    ' Disable links checking (not over consola)
    StopCheckingLinks
    
    Call ConsoleDialog.MouseMove(x, y)
    Call Minimap.MouseMove(MouseX, MouseY)
    
    If PantallaCompleta = 0 And Button = vbLeftButton Then
        If MoverVentana = 1 Then
            If Not UserMoving Then
                ' Mover form s�lo en la parte superior
                If y < 30 Then
                    Call Minimap.MouseUp(0, x, y)
                    MoverForm
                End If

                'Call Auto_Drag(Me.hwnd)
            End If

        End If

    End If

    MouseX = x - renderer.Left
    MouseY = y - renderer.Top
    
    If minimapa_visible Then

        Dim mapX As Integer
        Dim mapY As Integer
        Dim PosX As Integer
        Dim PosY As Integer

        
        If MapSize.Width * 2 < MainViewWidth Then
            mapX = (MainViewWidth - MapSize.Width * 2) / 2
        Else
            If UserPos.x > MapSize.Width - MainViewWidth / 4 Then
                PosX = MapSize.Width - MainViewWidth / 2
            ElseIf UserPos.x > MainViewWidth / 4 Then
                PosX = UserPos.x - MainViewWidth / 4
            End If
        End If
        If MapSize.Height * 2 < MainViewHeight Then
            mapY = (MainViewHeight - MapSize.Height * 2) / 2
        Else
            If UserPos.y > MapSize.Height - MainViewHeight / 4 Then
                PosY = MapSize.Height - MainViewHeight / 2
            ElseIf UserPos.y > MainViewHeight / 4 Then
                PosY = UserPos.y - MainViewHeight / 4
            End If
        End If
        
        
        MouseMapX = ((MouseX - mapX) / 2) + PosX
        MouseMapY = ((MouseY - mapY) / 2) + PosY
        
        If MouseMapX < 1 Then MouseMapX = 0
        If MouseMapY < 1 Then MouseMapY = 0
        If MouseMapX > MapSize.Width Then MouseMapX = 0
        If MouseMapY > MapSize.Height Then MouseMapY = 0
        
        MouseZona = getZona(MouseMapX, MouseMapY)
        
    End If
    
  
    ObjLbl.visible = False
    
    If EstadisticasBoton.Tag = "1" Then
        EstadisticasBoton.Picture = Nothing
        EstadisticasBoton.Tag = "0"

    End If
    
    If cmdlanzar.Tag = "1" Then
        cmdlanzar.Picture = Nothing
        cmdlanzar.Tag = "0"

    End If

    If imgInventario.Tag = "1" Then
        imgInventario.Picture = Nothing
        imgInventario.Tag = "0"

    End If

    If imgHechizos.Tag = "1" Then
        imgHechizos.Picture = Nothing
        imgHechizos.Tag = "0"

    End If
 
    If Image4(0).Tag = "1" Then
        Image4(0).Picture = Nothing
        Image4(0).Tag = "0"

    End If

    If Image4(1).Tag = "1" Then
        Image4(1).Picture = Nothing
        Image4(1).Tag = "0"

    End If

    If Image3.Tag = "1" Then
        Image3.Picture = Nothing
        Image3.Tag = "0"

    End If


    If Image4(1).Tag = "1" Then
        Image4(1).Picture = Nothing
        Image4(1).Tag = "0"

    End If

    If OpcionesBoton.Tag = "1" Then
        OpcionesBoton.Picture = Nothing
        OpcionesBoton.Tag = "0"

    End If
    
    If btnMenu.Tag = "1" Then
        btnMenu.Picture = Nothing
        btnMenu.Tag = "0"
    End If
    
  

    If ShowPercentage Then
        lblPorcLvl.visible = True
        exp.visible = False
    Else
        lblPorcLvl.visible = False
        exp.visible = True
    End If
    

    
    frmMenuUser.LostFocus
    frmMenuGM.LostFocus
    frmMenuNPC.LostFocus
    
    Exit Sub

Form_MouseMove_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.Form_MouseMove", Erl)
    Resume Next
    
End Sub




Private Sub picInv_DblClick()
    
    On Error GoTo picInv_DblClick_Err
    

    If Not picInv.visible Then Exit Sub
    
    If frmCarp.visible Or frmHerrero.visible Or frmComerciar.visible Or frmBancoObj.visible Then Exit Sub
    If pausa Then Exit Sub
    
    If UserMeditar Then Exit Sub
    'If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then Exit Sub
    
    If macrotrabajo.enabled Then DesactivarMacroTrabajo
    
    If Not Inventario.IsItemSelected Then Exit Sub

    ' Hacemos acci�n del doble clic correspondiente
    Dim ObjType As Byte

    ObjType = ObjData(Inventario.ObjIndex(Inventario.SelectedItem)).ObjType

    Select Case ObjType

        Case eObjType.otArmadura, eObjType.otESCUDO, eObjType.otmagicos, eObjType.otFlechas, eObjType.otCASCO, eObjType.otNudillos, eObjType.otAnillos, eObjType.otManchas
            If Not Inventario.Equipped(Inventario.SelectedItem) Then
                Call WriteEquipItem(Inventario.SelectedItem)
            End If
            
        Case eObjType.otWeapon

            If ObjData(Inventario.ObjIndex(Inventario.SelectedItem)).proyectil = 1 And Inventario.Equipped(Inventario.SelectedItem) Then
                Call WriteUseItem(Inventario.SelectedItem)
            Else
                If Not Inventario.Equipped(Inventario.SelectedItem) Then
                    Call WriteEquipItem(Inventario.SelectedItem)
                End If
            End If
            
        Case eObjType.OtHerramientas

            If Inventario.Equipped(Inventario.SelectedItem) Then
                Call WriteUseItem(Inventario.SelectedItem)
            Else
                If Not Inventario.Equipped(Inventario.SelectedItem) Then
                    Call WriteEquipItem(Inventario.SelectedItem)
                End If
            End If
                
        Case Else
            Call CountPacketIterations(packetControl(ClientPacketID.UseItem), 180)
                   ' Debug.Print "QWEASDqweads"
            Call WriteUseItem(Inventario.SelectedItem)
            
    End Select

    
    Exit Sub

picInv_DblClick_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.picInv_DblClick", Erl)
    Resume Next
    
End Sub
Private Function countRepts(ByVal packet As Long)
    
End Function




Private Sub SendTxt_Change()
    
    On Error GoTo SendTxt_Change_Err
    

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 3/06/2006
    '3/06/2006: Maraxus - imped� se inserten caract�res no imprimibles
    '**************************************************************
    If Len(SendTxt.Text) > 160 Then
        stxtbuffer = "Soy un cheater, avisenle a un gm"
    Else

        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i         As Long

        Dim tempStr   As String

        Dim CharAscii As Integer
        
        For i = 1 To Len(SendTxt.Text)
            CharAscii = Asc(mid$(SendTxt.Text, i, 1))

            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempStr = tempStr & Chr$(CharAscii)

            End If

        Next i
        
        If tempStr <> SendTxt.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.Text = tempStr

        End If
        
        stxtbuffer = SendTxt.Text

    End If

    
    Exit Sub

SendTxt_Change_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.SendTxt_Change", Erl)
    Resume Next
    
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    
    On Error GoTo SendTxt_KeyPress_Err
    

    If Not (KeyAscii = vbKeyBack) And Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then KeyAscii = 0

    
    Exit Sub

SendTxt_KeyPress_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.SendTxt_KeyPress", Erl)
    Resume Next
    
End Sub

Private Function InGameArea() As Boolean
    
    On Error GoTo InGameArea_Err
    

    If clicX < renderer.Left Or clicX > renderer.Left + (32 * 23) Then Exit Function
    If clicY < renderer.Top Or clicY > renderer.Top + (32 * 17) Then Exit Function
    InGameArea = True

    
    Exit Function

InGameArea_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.InGameArea", Erl)
    Resume Next
    
End Function

Private Sub MoverForm()
    
    On Error GoTo moverForm_Err
    

    Dim res As Long

    ReleaseCapture
    res = SendMessage(Me.hWnd, WM_SYSCOMMAND, MOUSE_MOVE, 0)

    
    Exit Sub

moverForm_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.moverForm", Erl)
    Resume Next
    
End Sub

Private Sub imgSpellInfo_Click()
    
    On Error GoTo imgSpellInfo_Click_Err
    

    If hlst.ListIndex <> -1 Then
        Call WriteSpellInfo(hlst.ListIndex + 1)

    End If

    
    Exit Sub

imgSpellInfo_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.imgSpellInfo_Click", Erl)
    Resume Next
    
End Sub

Private Sub UpdateDaytime_Timer()
    ' Si no hay luz de mapa, usamos la luz ambiental
    
    On Error GoTo UpdateDaytime_Timer_Err
    
    Call RevisarHoraMundo
    
    Exit Sub

UpdateDaytime_Timer_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.UpdateDaytime_Timer", Erl)
    Resume Next
    
End Sub

Private Sub UpdateLight_Timer()
    
    On Error GoTo UpdateLight_Timer_Err
    
    
    If light_transition < 1# Then
        light_transition = light_transition + STEP_LIGHT_TRANSITION * UpdateLight.Interval
        
        If light_transition > 1# Then light_transition = 1#
        
        Call LerpRGBA(global_light, last_light, next_light, light_transition)
        Call MapUpdateGlobalLight
    End If
    
    
    Exit Sub

UpdateLight_Timer_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.UpdateLight_Timer", Erl)
    Resume Next
    
End Sub


Public Sub OnClientDisconnect(ByVal Error As Long)
    On Error GoTo OnClientDisconnect_Err

    If (Error = 10061) Then
        If frmConnect.visible Then
            Call TextoAlAsistente("Ha ocurrido un error de conexi�n!")
            Call MsgBox("�No me pude conectar! Te recomiendo verificar el estado de los servidores en argentumworld?.com.ar y asegurarse de estar conectado a internet.")
        Else
            Call MsgBox("Ha ocurrido un error al conectar con el servidor. Le recomendamos verificar el estado de los servidores en argentumworld?.com.ar, y asegurarse de estar conectado directamente a internet", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error al conectar")
        End If
    Else
    
        frmConnect.MousePointer = 1
        ShowFPS.enabled = False

        If (Error <> 0 And Error <> 2) Then
            Call TextoAlAsistente("Ha ocurrido un error de conexi�n!")
            Call MsgBox("Ha ocurrido un error al conectar con el servidor. Le recomendamos verificar el estado de los servidores en argentumworld?.com.ar, y asegurarse de estar conectado directamente a internet", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error al conectar")
            
            If frmConnect.visible Then
                Connected = False
            Else
                If (Connected) Then
                    Call HandleDisconnect
                End If
            End If
          
        Else
            Call RegistrarError(Error, "Conexion cerrada", "OnClientDisconnect")
            If frmConnect.visible Then
                Connected = False
            Else
                If (Connected) Then
                    Call HandleDisconnect
                End If
            End If
        End If
    End If


    Exit Sub

OnClientDisconnect_Err:
    Call RegistrarError(err.Number, err.Description, "frmMain.MainSocket_LastError", Erl)
    Resume Next
End Sub

Private Sub imgDeleteItem_Click()
    If Not frmMain.Inventario.IsItemSelected Then
        Call AddToConsole("No tienes seleccionado ning�n item", 255, 255, 255, False, False, False)
    Else
        If MsgBox("Seguro que desea eliminar el item?", vbYesNo, "Eliminar objeto") = vbYes Then
            Call WriteDeleteItem(frmMain.Inventario.SelectedItem)
        End If
    End If
End Sub
