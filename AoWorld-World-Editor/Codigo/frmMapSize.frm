VERSION 5.00
Begin VB.Form frmMapSize 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Map Size"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5655
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   377
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox tHeight 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Text            =   "0"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox tWidth 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Text            =   "0"
      Top             =   600
      Width           =   1215
   End
   Begin WorldEditor.lvButtons_H cmdAceptar 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Caption         =   "&Aceptar y Aplicar"
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
      cBack           =   12648384
   End
   Begin WorldEditor.lvButtons_H lvButtons_H1 
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "&Mover 1 - X"
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
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H lvButtons_H3 
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "&Mover 10 - X"
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
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H lvButtons_H4 
      Height          =   495
      Left            =   3840
      TabIndex        =   7
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "&Mover 1 - Y"
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
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H lvButtons_H5 
      Height          =   495
      Left            =   3840
      TabIndex        =   8
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "&Mover 10 - Y"
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
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H lvButtons_H6 
      Height          =   495
      Left            =   2280
      TabIndex        =   9
      Top             =   1800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "&Achicar 1 - X"
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
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H lvButtons_H7 
      Height          =   495
      Left            =   2280
      TabIndex        =   10
      Top             =   2400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "&Achicar 10 - X"
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
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H lvButtons_H8 
      Height          =   495
      Left            =   3840
      TabIndex        =   11
      Top             =   1800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "&Achicar 1 - Y"
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
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H lvButtons_H9 
      Height          =   495
      Left            =   3840
      TabIndex        =   12
      Top             =   2400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "&Achicar 10 - Y"
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
      cBack           =   -2147483633
   End
   Begin VB.Label Label2 
      Caption         =   "Alto:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Ancho:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "frmMapSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
Dim X As Integer, Y As Integer
Dim OldMapData() As MapBlock
Dim W As Integer
Dim H As Integer
W = MapSize.Width
H = MapSize.Height
ReDim OldMapData(1 To W, 1 To H)

For X = 1 To W
    For Y = 1 To H
        OldMapData(X, Y) = MapData(X, Y)
    Next Y
Next X

MapSize.Width = Int(tWidth.Text)
MapSize.Height = Int(tHeight.Text)
ReDim MapData(1 To MapSize.Width, 1 To MapSize.Height)

For X = 1 To IIf(W > MapSize.Width, MapSize.Width, W)
    For Y = 1 To IIf(H > MapSize.Height, MapSize.Height, H)
        MapData(X, Y) = OldMapData(X, Y)
    Next Y
Next X

Unload Me
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub Form_Load()
tWidth.Text = MapSize.Width
tHeight.Text = MapSize.Height
End Sub

Private Sub lvButtons_H1_Click()
Dim X As Integer
Dim Y As Integer
Dim xx As Integer
Dim yy As Integer

xx = 2
yy = 1

For Y = yy To MapSize.Height
    For X = xx To MapSize.Width
        MapData(X - xx + 1, Y - yy + 1) = MapData(X, Y)
    Next X
Next Y

MapSize.Width = MapSize.Width - xx + 1
MapSize.Height = MapSize.Height - yy + 1

tWidth.Text = MapSize.Width
tHeight.Text = MapSize.Height

End Sub

Private Sub lvButtons_H2_Click()


End Sub

Private Sub lvButtons_H3_Click()
Dim X As Integer
Dim Y As Integer
Dim xx As Integer
Dim yy As Integer

xx = 11
yy = 1

For Y = yy To MapSize.Height
    For X = xx To MapSize.Width
        MapData(X - xx + 1, Y - yy + 1) = MapData(X, Y)
    Next X
Next Y

MapSize.Width = MapSize.Width - xx + 1
MapSize.Height = MapSize.Height - yy + 1

tWidth.Text = MapSize.Width
tHeight.Text = MapSize.Height
End Sub

Private Sub lvButtons_H4_Click()
Dim X As Integer
Dim Y As Integer
Dim xx As Integer
Dim yy As Integer

xx = 1
yy = 2

For Y = yy To MapSize.Height
    For X = xx To MapSize.Width
        MapData(X - xx + 1, Y - yy + 1) = MapData(X, Y)
    Next X
Next Y

MapSize.Width = MapSize.Width - xx + 1
MapSize.Height = MapSize.Height - yy + 1

tWidth.Text = MapSize.Width
tHeight.Text = MapSize.Height
End Sub

Private Sub lvButtons_H5_Click()
Dim X As Integer
Dim Y As Integer
Dim xx As Integer
Dim yy As Integer

xx = 1
yy = 11

For Y = yy To MapSize.Height
    For X = xx To MapSize.Width
        MapData(X - xx + 1, Y - yy + 1) = MapData(X, Y)
    Next X
Next Y

MapSize.Width = MapSize.Width - xx + 1
MapSize.Height = MapSize.Height - yy + 1

tWidth.Text = MapSize.Width
tHeight.Text = MapSize.Height
End Sub

Private Sub lvButtons_H6_Click()
MapSize.Width = MapSize.Width - 1
tWidth.Text = MapSize.Width
End Sub

Private Sub lvButtons_H7_Click()
MapSize.Width = MapSize.Width - 10
tWidth.Text = MapSize.Width
End Sub

Private Sub lvButtons_H8_Click()
MapSize.Height = MapSize.Height - 1
tHeight.Text = MapSize.Height
End Sub

Private Sub lvButtons_H9_Click()
MapSize.Height = MapSize.Height - 10
tHeight.Text = MapSize.Height
End Sub
