VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Begin VB.Form frmOpciones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   7365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7980
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   7320
      Top             =   6600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.HScrollBar Alpha 
      Height          =   315
      LargeChange     =   60
      Left            =   9240
      Max             =   255
      SmallChange     =   2
      TabIndex        =   10
      Top             =   2880
      Value           =   120
      Width           =   2775
   End
   Begin VB.CheckBox Macro 
      Caption         =   "Arriba"
      Height          =   255
      Index           =   0
      Left            =   11400
      TabIndex        =   9
      Top             =   10680
      Width           =   855
   End
   Begin VB.CheckBox Macro 
      Caption         =   "Abajo"
      Height          =   255
      Index           =   1
      Left            =   12600
      TabIndex        =   8
      Top             =   10680
      Width           =   855
   End
   Begin VB.CheckBox Macro 
      Caption         =   "Izquierda"
      Height          =   255
      Index           =   2
      Left            =   11400
      TabIndex        =   7
      Top             =   10920
      Width           =   1095
   End
   Begin VB.CheckBox Macro 
      Caption         =   "Derecha"
      Height          =   255
      Index           =   3
      Left            =   12600
      TabIndex        =   6
      Top             =   10920
      Width           =   1095
   End
   Begin VB.CheckBox Check7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ver nombres del mapa"
      Height          =   255
      Left            =   11160
      TabIndex        =   5
      Tag             =   "13"
      Top             =   9600
      Width           =   1935
   End
   Begin VB.CheckBox Check11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Desvanecimiento de techos"
      Height          =   255
      Left            =   11280
      TabIndex        =   4
      Top             =   9120
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   195
      Left            =   7560
      TabIndex        =   3
      Top             =   10920
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   9600
      Top             =   10080
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Bajar Volumen"
      Height          =   375
      Left            =   7440
      TabIndex        =   1
      Top             =   9960
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Subir Volumen"
      Height          =   375
      Left            =   7440
      TabIndex        =   0
      Top             =   10440
      Width           =   1335
   End
   Begin VB.PictureBox PanelVideo 
      BorderStyle     =   0  'None
      Height          =   4965
      Left            =   240
      ScaleHeight     =   331
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   504
      TabIndex        =   14
      Top             =   1800
      Visible         =   0   'False
      Width           =   7560
      Begin VB.HScrollBar MinimapOpacity 
         Height          =   255
         LargeChange     =   50
         Left            =   1800
         Max             =   255
         SmallChange     =   10
         TabIndex        =   29
         Top             =   4200
         Value           =   255
         Width           =   1935
      End
      Begin VB.ComboBox cboLuces 
         Height          =   315
         ItemData        =   "frmOpciones.frx":0152
         Left            =   240
         List            =   "frmOpciones.frx":015F
         TabIndex        =   21
         Top             =   3600
         Width           =   2175
      End
      Begin VB.Label LabelMinimapOpacity 
         BackStyle       =   0  'Transparent
         Caption         =   "Opacidad Minimapa"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Image num_comp_inv 
         Height          =   255
         Left            =   270
         Top             =   3090
         Width           =   255
      End
      Begin VB.Image chkItemsEnRender 
         Height          =   255
         Left            =   270
         Top             =   2715
         Width           =   255
      End
      Begin VB.Image Fullscreen 
         Height          =   255
         Left            =   270
         Top             =   2310
         Width           =   255
      End
      Begin VB.Image Respiracion 
         Height          =   255
         Left            =   270
         Top             =   1905
         Width           =   255
      End
      Begin VB.Image VSync 
         Height          =   255
         Left            =   270
         Top             =   1500
         Width           =   255
      End
      Begin VB.Image Check5 
         Height          =   255
         Left            =   270
         Top             =   1095
         Width           =   255
      End
      Begin VB.Image Check6 
         Height          =   255
         Left            =   270
         Top             =   690
         Width           =   255
      End
   End
   Begin VB.PictureBox PanelJugabilidad 
      BorderStyle     =   0  'None
      Height          =   4845
      Left            =   240
      ScaleHeight     =   323
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   504
      TabIndex        =   12
      Top             =   1800
      Width           =   7560
      Begin VB.TextBox txlCantidadTextosConsola 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   2325
         TabIndex        =   25
         Text            =   "8"
         Top             =   1485
         Width           =   495
      End
      Begin VB.ComboBox cbTutorial 
         BackColor       =   &H80000007&
         ForeColor       =   &H8000000B&
         Height          =   315
         ItemData        =   "frmOpciones.frx":017C
         Left            =   4800
         List            =   "frmOpciones.frx":0186
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   4290
         Width           =   1695
      End
      Begin VB.ComboBox cbRenderNpcs 
         BackColor       =   &H80000007&
         ForeColor       =   &H8000000B&
         Height          =   315
         ItemData        =   "frmOpciones.frx":01A0
         Left            =   1440
         List            =   "frmOpciones.frx":01AA
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   4290
         Width           =   1695
      End
      Begin VB.ComboBox cbLenguaje 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         ForeColor       =   &H80000005&
         Height          =   315
         ItemData        =   "frmOpciones.frx":01C2
         Left            =   3960
         List            =   "frmOpciones.frx":01CC
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   2880
         Width           =   3255
      End
      Begin VB.HScrollBar scrSens 
         Height          =   315
         LargeChange     =   5
         Left            =   240
         Max             =   20
         Min             =   1
         TabIndex        =   19
         Top             =   2520
         Value           =   10
         Width           =   3375
      End
      Begin VB.ComboBox cbBloqueoHechizos 
         BackColor       =   &H80000007&
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1080
         Width           =   3375
      End
      Begin VB.Label lblLineasConsolaModificador 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   2040
         TabIndex        =   27
         Top             =   1515
         Width           =   180
      End
      Begin VB.Label lblLineasConsolaModificador 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   240
         Index           =   1
         Left            =   2910
         TabIndex        =   26
         Top             =   1515
         Width           =   180
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Lineas de consola"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   360
         TabIndex        =   24
         Top             =   1500
         Width           =   1695
      End
      Begin VB.Image Check8 
         Height          =   255
         Left            =   4005
         Top             =   1455
         Width           =   255
      End
      Begin VB.Image Check2 
         Height          =   255
         Left            =   1875
         Top             =   3195
         Width           =   255
      End
      Begin VB.Image Check3 
         Height          =   255
         Left            =   270
         Top             =   3195
         Width           =   255
      End
      Begin VB.Image Check4 
         Height          =   255
         Left            =   270
         Top             =   1050
         Width           =   255
      End
      Begin VB.Image Check9 
         Height          =   255
         Left            =   270
         Top             =   645
         Width           =   255
      End
   End
   Begin VB.PictureBox PanelAudio 
      BorderStyle     =   0  'None
      Height          =   4845
      Left            =   240
      ScaleHeight     =   323
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   504
      TabIndex        =   15
      Top             =   1800
      Visible         =   0   'False
      Width           =   7560
      Begin VB.HScrollBar HScroll1 
         Height          =   315
         LargeChange     =   1000
         Left            =   3960
         Max             =   0
         Min             =   -4000
         SmallChange     =   2
         TabIndex        =   18
         Top             =   3000
         Width           =   3375
      End
      Begin VB.HScrollBar scrVolume 
         Height          =   315
         LargeChange     =   1000
         Left            =   3960
         Max             =   0
         Min             =   -4000
         SmallChange     =   2
         TabIndex        =   17
         Top             =   1800
         Width           =   3375
      End
      Begin VB.HScrollBar scrMidi 
         Height          =   315
         LargeChange     =   1000
         Left            =   3960
         Max             =   0
         Min             =   -4000
         SmallChange     =   2
         TabIndex        =   16
         Top             =   600
         Width           =   3375
      End
      Begin VB.Image chko 
         Height          =   255
         Index           =   2
         Left            =   255
         Top             =   1905
         Width           =   255
      End
      Begin VB.Image chko 
         Height          =   255
         Index           =   0
         Left            =   255
         Top             =   690
         Width           =   255
      End
      Begin VB.Image chkInvertir 
         Height          =   255
         Left            =   255
         Top             =   2310
         Width           =   255
      End
      Begin VB.Image chko 
         Height          =   255
         Index           =   3
         Left            =   255
         Top             =   1500
         Width           =   255
      End
      Begin VB.Image chko 
         Height          =   255
         Index           =   1
         Left            =   255
         Top             =   1095
         Width           =   255
      End
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.Image BtnSolapa 
      Height          =   420
      Index           =   2
      Left            =   5280
      Tag             =   "0"
      Top             =   1200
      Width           =   2460
   End
   Begin VB.Image BtnSolapa 
      Height          =   420
      Index           =   1
      Left            =   2790
      Tag             =   "0"
      Top             =   1200
      Width           =   2460
   End
   Begin VB.Image BtnSolapa 
      Height          =   420
      Index           =   0
      Left            =   300
      Tag             =   "2"
      Top             =   1200
      Width           =   2460
   End
   Begin VB.Image facebook 
      Height          =   300
      Left            =   4270
      Tag             =   "0"
      Top             =   6840
      Width           =   300
   End
   Begin VB.Image instagram 
      Height          =   300
      Left            =   4650
      Tag             =   "0"
      Top             =   6840
      Width           =   300
   End
   Begin VB.Image discord 
      Height          =   300
      Left            =   3520
      Tag             =   "0"
      Top             =   6840
      Width           =   300
   End
   Begin VB.Image cmdChangePassword 
      Height          =   420
      Left            =   5290
      Tag             =   "0"
      Top             =   6790
      Width           =   2265
   End
   Begin VB.Image Command1 
      Height          =   420
      Left            =   480
      Tag             =   "0"
      Top             =   6790
      Width           =   2265
   End
   Begin VB.Image cmdcerrar 
      Height          =   360
      Left            =   7560
      Tag             =   "0"
      Top             =   0
      Width           =   435
   End
   Begin VB.Image cmdweb 
      Height          =   300
      Left            =   3100
      Tag             =   "0"
      Top             =   6840
      Width           =   300
   End
   Begin VB.Image cmdayuda 
      Height          =   435
      Left            =   9120
      Tag             =   "0"
      Top             =   5880
      Width           =   1395
   End
   Begin VB.Label txtMSens 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "10"
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
      Height          =   255
      Left            =   4200
      TabIndex        =   11
      Top             =   5680
      Width           =   3375
   End
   Begin VB.Image check1 
      Height          =   210
      Left            =   8760
      Top             =   3840
      Width           =   180
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   9000
      TabIndex        =   2
      Top             =   10680
      Width           =   1455
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
Option Explicit

Private Bajar       As Boolean

Private Subir       As Boolean

Public bmoving      As Boolean

Public dX           As Integer

Public dY           As Integer

' Constantes para SendMessage
Const WM_SYSCOMMAND As Long = &H112&

Const MOUSE_MOVE    As Long = &HF012&

Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long

' funci�n Api para aplicar la transparencia a la ventana
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

' Funciones api para los estilos de la ventana
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'constantes
Private Const GWL_EXSTYLE = (-20)

Private Const LWA_ALPHA = &H2

Private Const WS_EX_LAYERED = &H80000

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST = -1

Private Const HWND_NOTOPMOST = -2

Private Const SWP_NOMOVE = &H2

Private Const SWP_NOSIZE = &H1

Public Function Is_Transparent(ByVal hWnd As Long) As Boolean
    
    On Error GoTo Is_Transparent_Err
    

    
  
    Dim msg As Long
  
    msg = GetWindowLong(hWnd, GWL_EXSTYLE)
         
    If (msg And WS_EX_LAYERED) = WS_EX_LAYERED Then
        Is_Transparent = True
    Else
        Is_Transparent = False

    End If
  
    If err Then
        Is_Transparent = False

    End If
  
    
    Exit Function

Is_Transparent_Err:
    Call RegistrarError(err.Number, err.Description, "frmOpciones.Is_Transparent", Erl)
    Resume Next
    
End Function
  
'Funci�n que aplica la transparencia, se le pasa el hwnd del form y un valor de 0 a 255
Public Function Aplicar_Transparencia(ByVal hWnd As Long, Valor As Integer) As Long
    
    On Error GoTo Aplicar_Transparencia_Err
    
  
    Dim msg As Long
  
    
  
    If Valor < 0 Or Valor > 255 Then
        Aplicar_Transparencia = 1
    Else
        msg = GetWindowLong(hWnd, GWL_EXSTYLE)
        msg = msg Or WS_EX_LAYERED
     
        SetWindowLong hWnd, GWL_EXSTYLE, msg
     
        'Establece la transparencia
        SetLayeredWindowAttributes hWnd, 0, Valor, LWA_ALPHA
  
        Aplicar_Transparencia = 0
  
    End If
  
    If err Then
        Aplicar_Transparencia = 2

    End If
  
    
    Exit Function

Aplicar_Transparencia_Err:
    Call RegistrarError(err.Number, err.Description, "frmOpciones.Aplicar_Transparencia", Erl)
    Resume Next
    
End Function

Private Sub Alpha_Change()
    
    On Error GoTo Alpha_Change_Err
    
    AlphaMacro = Alpha.Value

    
    Exit Sub

Alpha_Change_Err:
    Call RegistrarError(err.Number, err.Description, "frmOpciones.Alpha_Change", Erl)
    Resume Next
    
End Sub

Private Sub BtnSolapa_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim Name As String

    Select Case Index
    
        Case 0
            Name = "jugabilidad"
            PanelJugabilidad.visible = True
            PanelVideo.visible = False
            PanelAudio.visible = False
            Call SetSolapa(0, 2)
            Call SetSolapa(1, 0)
            Call SetSolapa(2, 0)
            
        Case 1
            Name = "video"
            PanelJugabilidad.visible = False
            PanelVideo.visible = True
            PanelAudio.visible = False
            Call SetSolapa(0, 0)
            Call SetSolapa(1, 2)
            Call SetSolapa(2, 0)
            
        Case 2
            Name = "audio"
            PanelJugabilidad.visible = False
            PanelVideo.visible = False
            PanelAudio.visible = True
            Call SetSolapa(0, 0)
            Call SetSolapa(1, 0)
            Call SetSolapa(2, 2)
    
    End Select
    
End Sub

Private Sub BtnSolapa_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    If BtnSolapa(Index).Tag = "0" Then
        Call SetSolapa(Index, 1)
    End If

End Sub

Private Sub SetSolapa(Index As Integer, ByVal Tag As String)

    Dim Name As String, estado As String
    
    Select Case Index
        Case 0: Name = "jugabilidad"
        Case 1: Name = "video"
        Case 2: Name = "audio"
    End Select
    
    Select Case Tag
        Case "0": estado = "off"
        Case "1": estado = "over"
        Case "2": estado = "default"
    End Select
    
    BtnSolapa(Index).Picture = LoadInterface("boton-" & Name & "-" & estado & ".bmp")
    BtnSolapa(Index).Tag = Tag

End Sub

Private Sub cbBloqueoHechizos_Click()

    ModoHechizos = cbBloqueoHechizos.ListIndex

End Sub


Private Sub cbLenguaje_Click()

    Dim message As String, title As String
       
    If cbLenguaje.ListIndex + 1 <> language Then
       
        Select Case cbLenguaje.ListIndex
        
            Case 0
                message = "Para que los cambios surjan efecto deber� volver a abrir el cliente."
                title = "Cambiar Idioma"
            
            Case 1
                message = "You must restart the game to apply the changes."
                title = "Change language"
            
        
        End Select
        
        If MsgBox(message, vbYesNo, title) = vbYes Then
            Call WriteVar(RESOURCES_PATH & "/OUTPUT/Configuracion.ini", "OPCIONES", "Localization", cbLenguaje.ListIndex + 1)
        End If
    End If
    
End Sub

Private Sub cboLuces_Click()

    Call WriteVar(RESOURCES_PATH & "/OUTPUT/Configuracion.ini", "VIDEO", "LuzGlobal", cboLuces.ListIndex)
    selected_light = cboLuces.ListIndex
End Sub



Private Sub cbTutorial_Click()
    If cbTutorial.ListIndex <> MostrarTutorial Then
        MostrarTutorial = cbTutorial.ListIndex
        If MostrarTutorial Then
            Dim i As Long
            
            For i = 1 To UBound(tutorial)
                Call WriteVar(RESOURCES_PATH & "/OUTPUT/Configuracion.ini", "TUTORIAL" & i, "Activo", 1)
                tutorial(i).Activo = 1
            Next i
        End If
        Call WriteVar(RESOURCES_PATH & "/OUTPUT/Configuracion.ini", "INITTUTORIAL", "MostrarTutorial", cbTutorial.ListIndex)
    End If
End Sub
Private Sub cbRenderNpcs_Click()
           
    If cbRenderNpcs.ListIndex <> npcs_en_render Then
        npcs_en_render = cbRenderNpcs.ListIndex
        Call WriteVar(RESOURCES_PATH & "/OUTPUT/Configuracion.ini", "OPCIONES", "NpcsEnRender", cbRenderNpcs.ListIndex)
    End If
End Sub


Private Sub Check4_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Check4_MouseUp_Err
    

    If PermitirMoverse = 1 Then
        PermitirMoverse = 0
    Else
        PermitirMoverse = 1

    End If

    If PermitirMoverse = 0 Then
        Check4.Picture = Nothing
    Else
        Check4.Picture = LoadInterface("check-amarillo.bmp")
    End If

    
    Exit Sub

Check4_MouseUp_Err:
    Call RegistrarError(err.Number, err.Description, "frmOpciones.Check4_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub Check5_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Check5_MouseUp_Err
    

    If MoverVentana = 1 Then
        MoverVentana = 0
    Else
        MoverVentana = 1

    End If

    If MoverVentana = 0 Then
        Check5.Picture = Nothing
    Else
        Check5.Picture = LoadInterface("check-amarillo.bmp")

    End If

    
    Exit Sub

Check5_MouseUp_Err:
    Call RegistrarError(err.Number, err.Description, "frmOpciones.Check5_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub Check2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Check2_MouseUp_Err
    CursoresGraficos = Not CursoresGraficos

    If CursoresGraficos Then
        Check2.Picture = LoadInterface("check-amarillo.bmp")
        Call WriteVar(RESOURCES_PATH & "/OUTPUT/" & "Configuracion.ini", "VIDEO", "CursoresGraficos", 1)
    Else
        Check2.Picture = Nothing
        Call WriteVar(RESOURCES_PATH & "/OUTPUT/" & "Configuracion.ini", "VIDEO", "CursoresGraficos", 0)
    End If
    
    MsgBox "Para que los cambios en esta opci�n sean reflejados, deber� reiniciar el cliente.", vbQuestion, "ArgentumWorld - Advertencia"

    Exit Sub

Check2_MouseUp_Err:
    Call RegistrarError(err.Number, err.Description, "frmOpciones.Check2_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub Check8_Click()
    On Error GoTo Check8_MouseUp_Err
    
    If ScrollArrastrar = 1 Then
        ScrollArrastrar = 0
        Check8.Picture = Nothing
    Else
        ScrollArrastrar = 1
        Check8.Picture = LoadInterface("check-amarillo.bmp")
    End If

    
    Exit Sub

Check8_MouseUp_Err:
    Call RegistrarError(err.Number, err.Description, "frmOpciones.Check8_MouseUp", Erl)
    Resume Next
End Sub

Private Sub chkInvertir_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo chkInvertir_MouseUp_Err
    

    If InvertirSonido = 1 Then
        InvertirSonido = 0
            
        Sound.InvertirSonido = False
    Else
        InvertirSonido = 1
        Sound.InvertirSonido = True

    End If
        
    If InvertirSonido = 0 Then
        chkInvertir.Picture = Nothing
    Else
        chkInvertir.Picture = LoadInterface("check-amarillo.bmp")

    End If

    
    Exit Sub

chkInvertir_MouseUp_Err:
    Call RegistrarError(err.Number, err.Description, "frmOpciones.chkInvertir_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub chkItemsEnRender_Click()
    InfoItemsEnRender = Not InfoItemsEnRender
    If InfoItemsEnRender Then
        chkItemsEnRender.Picture = LoadInterface("check-amarillo.bmp")
    Else
        chkItemsEnRender.Picture = Nothing
    End If
End Sub

Private Sub chkO_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo chkO_MouseUp_Err
    

    Call Sound.Sound_Play(SND_CLICK)

    Select Case Index

        Case 0

            If Musica <> CONST_DESHABILITADA Then
                Sound.Music_Stop
                Musica = CONST_DESHABILITADA
                scrMidi.enabled = False
            Else
                Musica = CONST_MP3
                scrMidi.enabled = True
                Sound.NextMusic = MapDat.music_numberHi
                Sound.Fading = 100

            End If

            If Musica = 0 Then
                chko(0).Picture = Nothing
            Else
                chko(0).Picture = LoadInterface("check-amarillo.bmp")

            End If

        Case 1

            If fX = 1 Then
                fX = 0
                chko(2).enabled = False
                scrVolume.enabled = False
            
                Call Sound.Sound_Stop_All
            Else
                fX = 1
                chko(2).enabled = True
                scrVolume.enabled = True

            End If
        
            If fX = 0 Then
                chko(1).Picture = Nothing
            Else
                chko(1).Picture = LoadInterface("check-amarillo.bmp")

            End If

        Case 2

            If FxNavega = 1 Then
                FxNavega = 0
            Else
                FxNavega = 1

            End If
        
            If FxNavega = 0 Then
                chko(2).Picture = Nothing
            Else
                chko(2).Picture = LoadInterface("check-amarillo.bmp")

            End If

        Case 3

            If AmbientalActivated = 1 Then
                HScroll1.enabled = False
                AmbientalActivated = 0
                Sound.LastAmbienteActual = 0
                Sound.AmbienteActual = 0
                Sound.Ambient_Stop
            Else
                HScroll1.enabled = True
                AmbientalActivated = 1
                Call AmbientarAudio(UserMap)

            End If

            If AmbientalActivated = 0 Then
                chko(3).Picture = Nothing
            Else
                chko(3).Picture = LoadInterface("check-amarillo.bmp")

            End If

    End Select

    
    Exit Sub

chkO_MouseUp_Err:
    Call RegistrarError(err.Number, err.Description, "frmOpciones.chkO_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub cmdayuda_Click()
    
    On Error GoTo cmdayuda_Click_Err
    
    Call FrmGmAyuda.Show

    
    Exit Sub

cmdayuda_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmOpciones.cmdayuda_Click", Erl)
    Resume Next
    
End Sub

Private Sub cmdayuda_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo cmdayuda_MouseMove_Err
    

    If cmdayuda.Tag = "0" Then
        cmdayuda.Picture = LoadInterface("config_ayuda.bmp")
        cmdayuda.Tag = "1"

    End If

    
    Exit Sub

cmdayuda_MouseMove_Err:
    Call RegistrarError(err.Number, err.Description, "frmOpciones.cmdayuda_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Command1_MouseMove_Err
    

    If Command1.Tag = "0" Then
        Command1.Picture = LoadInterface("boton-config-teclas-over.bmp")
        Command1.Tag = "1"

    End If

    cmdcerrar = Nothing
    cmdcerrar.Tag = "0"
    
    
    Exit Sub

Command1_MouseMove_Err:
    Call RegistrarError(err.Number, err.Description, "frmOpciones.Command1_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub cmdCerrar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo cmdCerrar_MouseMove_Err
    

    If cmdcerrar.Tag = "0" Then
        'cmdCerrar.Picture = LoadInterface("config_cerrar.bmp")
        cmdcerrar.Tag = "1"

    End If

    cmdChangePassword = Nothing
    cmdChangePassword.Tag = "0"
    Command1 = Nothing
    Command1.Tag = "0"

    
    Exit Sub

cmdCerrar_MouseMove_Err:
    Call RegistrarError(err.Number, err.Description, "frmOpciones.cmdCerrar_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub cmdChangePassword_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo cmdChangePassword_MouseMove_Err
    

    If cmdChangePassword.Tag = "0" Then
        cmdChangePassword.Picture = LoadInterface("boton-cambiar-pass-over.bmp")
        cmdChangePassword.Tag = "1"

    End If

    cmdcerrar = Nothing
    cmdcerrar.Tag = "0"

    
    Exit Sub

cmdChangePassword_MouseMove_Err:
    Call RegistrarError(err.Number, err.Description, "frmOpciones.cmdChangePassword_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub cmdWeb_Click()
    
    On Error GoTo cmdWeb_Click_Err
    
    ShellExecute Me.hWnd, "open", "https://argentumworld?.com.ar/", "", "", 0

    
    Exit Sub

cmdWeb_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmOpciones.cmdWeb_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command5_Click()
    
    On Error GoTo Command5_Click_Err
    
    MsgBox ("Proximamente")

    
    Exit Sub

Command5_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmOpciones.Command5_Click", Erl)
    Resume Next
    
End Sub

Private Sub discord_Click()
    
    On Error GoTo discord_Click_Err
    
    ShellExecute Me.hWnd, "open", "https://discord.gg/e3juVbF", "", "", 0

    
    Exit Sub

discord_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmOpciones.discord_Click", Erl)
    Resume Next
    
End Sub

Private Sub facebook_Click()
    
    On Error GoTo facebook_Click_Err
    
    ShellExecute Me.hWnd, "open", "https://argentumworld?.com.ar/", "", "", 0

    
    Exit Sub

facebook_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmOpciones.facebook_Click", Erl)
    Resume Next
    
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Call Aplicar_Transparencia(Me.hWnd, 240)
'    Call FormParser.Parse_Form(Me)
    Me.Picture = LoadInterface("configuracion-vacio.bmp")
    
    PanelJugabilidad.Picture = LoadInterface("configuracion-jugabilidad.bmp")
    PanelVideo.Picture = LoadInterface("configuracion-video.bmp")
    PanelAudio.Picture = LoadInterface("configuracion-audio.bmp")
    
    
    selected_light = GetVar(RESOURCES_PATH & "/OUTPUT/Configuracion.ini", "VIDEO", "LuzGlobal")
    
    If LenB(selected_light) = 0 Then selected_light = 0
    
    cboLuces.ListIndex = selected_light

    BtnSolapa(0).Picture = LoadInterface("boton-jugabilidad-default.bmp")
    BtnSolapa(1).Picture = LoadInterface("boton-video-off.bmp")
    BtnSolapa(2).Picture = LoadInterface("boton-audio-off.bmp")
    txlCantidadTextosConsola.Text = ConsoleDialog.RENDER_DIALOGS_MAX
    Exit Sub

Form_Load_Err:
    Call RegistrarError(err.Number, err.Description, "frmOpciones.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Form_KeyPress_Err
    

    If (KeyAscii = 27) Then
        Unload Me
    End If

    
    Exit Sub

Form_KeyPress_Err:
    Call RegistrarError(err.Number, err.Description, "frmOpciones.Form_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub Check3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

   ' If Check3 Then
      '  SwapMouseButton 1
       ' Check3.Picture = LoadInterface("check-amarillo.bmp")
  '   Else
      ' SwapMouseButton 0
      ' Check3.Picture = Nothing
  '  End If
End Sub

Private Sub Check6_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Check6_MouseUp_Err
    

    If FPSFLAG = 1 Then
        FPSFLAG = 0
    Else
        FPSFLAG = 1

    End If
        
    If FPSFLAG = 0 Then
        Check6.Picture = Nothing
    Else
        Check6.Picture = LoadInterface("check-amarillo.bmp")

    End If

    
    Exit Sub

Check6_MouseUp_Err:
    Call RegistrarError(err.Number, err.Description, "frmOpciones.Check6_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub Check9_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Check9_MouseUp_Err
    

    If CopiarDialogoAConsola = 1 Then
        CopiarDialogoAConsola = 0
    Else
        CopiarDialogoAConsola = 1

    End If
        
    If CopiarDialogoAConsola = 0 Then
        Check9.Picture = Nothing
    Else
        Check9.Picture = LoadInterface("check-amarillo.bmp")

    End If

    
    Exit Sub

Check9_MouseUp_Err:
    Call RegistrarError(err.Number, err.Description, "frmOpciones.Check9_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub Command2_Click()
    
    On Error GoTo Command2_Click_Err
    
    Bajar = True
    Subir = False
    Timer1.enabled = True

    
    Exit Sub

Command2_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmOpciones.Command2_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command3_Click()
    
    On Error GoTo Command3_Click_Err
    
    Subir = True
    Bajar = False
    Timer1.enabled = True

    
    Exit Sub

Command3_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmOpciones.Command3_Click", Erl)
    Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Form_MouseMove_Err

    MoverForm Me.hWnd
    cmdayuda = Nothing
    cmdayuda.Tag = "0"
    discord = Nothing
    discord.Tag = "0"
    cmdweb = Nothing
    cmdweb.Tag = "0"
    instagram = Nothing
    instagram.Tag = "0"
    facebook = Nothing
    facebook.Tag = "0"
    Command1 = Nothing
    Command1.Tag = "0"
    cmdcerrar = Nothing
    cmdcerrar.Tag = "0"
    cmdChangePassword = Nothing
    cmdChangePassword.Tag = "0"

    If BtnSolapa(0).Tag = "1" Then
        Call SetSolapa(0, 0)
    End If
    
    If BtnSolapa(1).Tag = "1" Then
        Call SetSolapa(1, 0)
    End If
    
    If BtnSolapa(2).Tag = "1" Then
        Call SetSolapa(2, 0)
    End If
    
    Exit Sub

Form_MouseMove_Err:
    Call RegistrarError(err.Number, err.Description, "frmOpciones.Form_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub cmdcerrar_Click()
    
    On Error GoTo cmdcerrar_Click_Err
    
    Call GuardarOpciones
    Me.visible = False
    frmMain.SetFocus

    
    Exit Sub

cmdcerrar_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmOpciones.cmdcerrar_Click", Erl)
    Resume Next
    
End Sub

Private Sub cmdChangePassword_Click()
    
    On Error GoTo cmdChangePassword_Click_Err
    
    Call ShellExecute(0, "open", "http://argentumworld?.com.ar/recuperar", 0, 0, 1)
    
    Exit Sub

cmdChangePassword_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmOpciones.cmdChangePassword_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command1_Click()
    
    On Error GoTo Command1_Click_Err
    
    Call frmCustomKeys.Show(vbModeless, Me)

    
    Exit Sub

Command1_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmOpciones.Command1_Click", Erl)
    Resume Next
    
End Sub

Public Sub Init()
    
    On Error GoTo Init_Err
    
    
    If CopiarDialogoAConsola = 0 Then
        Check9.Picture = Nothing
    Else
        Check9.Picture = LoadInterface("check-amarillo.bmp")

    End If
    
    If MoverVentana = 0 Then
        Check5.Picture = Nothing
    Else
        Check5.Picture = LoadInterface("check-amarillo.bmp")
    End If
    
    If InfoItemsEnRender Then
        chkItemsEnRender.Picture = LoadInterface("check-amarillo.bmp")
    Else
        chkItemsEnRender.Picture = Nothing
    End If

    If CursoresGraficos = 0 Then
        Check2.Picture = Nothing
    Else
        Check2.Picture = LoadInterface("check-amarillo.bmp")
    End If
    
    If PantallaCompleta Then
        Fullscreen.Picture = LoadInterface("check-amarillo.bmp")
    Else
        Fullscreen.Picture = Nothing
    End If

    If PermitirMoverse = 0 Then
        Check4.Picture = Nothing
    Else
        Check4.Picture = LoadInterface("check-amarillo.bmp")

    End If
    
    If ScrollArrastrar = 0 Then
        Check8.Picture = Nothing
    Else
        Check8.Picture = LoadInterface("check-amarillo.bmp")

    End If
    
    If Musica = 0 Then
        chko(0).Picture = Nothing
    Else
        chko(0).Picture = LoadInterface("check-amarillo.bmp")

    End If
    
    If FxNavega = 0 Then
        chko(2).Picture = Nothing
    Else
        chko(2).Picture = LoadInterface("check-amarillo.bmp")
    End If
    

    If NumerosCompletosInventario = 0 Then
        num_comp_inv.Picture = Nothing
    Else
        num_comp_inv.Picture = LoadInterface("check-amarillo.bmp")
    End If
    
    If MostrarRespiracion Then
        Respiracion.Picture = LoadInterface("check-amarillo.bmp")
    Else
        Respiracion.Picture = Nothing
    End If
    
    If AmbientalActivated = 0 Then
        chko(3).Picture = Nothing
    Else
        chko(3).Picture = LoadInterface("check-amarillo.bmp")
    End If

    If fX = 0 Then
        chko(1).Picture = Nothing
    Else
        chko(1).Picture = LoadInterface("check-amarillo.bmp")
    End If
    
    If InvertirSonido = 0 Then
        chkInvertir.Picture = Nothing
    Else
        chkInvertir.Picture = LoadInterface("check-amarillo.bmp")
    End If
    
    If FPSFLAG = 0 Then
        Check6.Picture = Nothing
    Else
        Check6.Picture = LoadInterface("check-amarillo.bmp")

    End If

    If VSyncActivado Then
        VSync.Picture = LoadInterface("check-amarillo.bmp")
    Else
        VSync.Picture = Nothing
    End If
    
    scrVolume.Value = VolFX
    HScroll1.Value = VolAmbient
    scrMidi.Value = VolMusic
    
    Alpha.Value = AlphaMacro
    
    Call cbBloqueoHechizos.Clear
    Call cbBloqueoHechizos.AddItem("Bloqueo en soltar")
    Call cbBloqueoHechizos.AddItem("Bloqueo al lanzar")
    Call cbBloqueoHechizos.AddItem("Sin bloqueo")
    cbBloqueoHechizos.ListIndex = ModoHechizos
    scrSens.Value = SensibilidadMouse
    MinimapOpacity.Value = Minimap.Opacity

    Me.Show vbModeless, frmMain

    
    Exit Sub

Init_Err:
    Call RegistrarError(err.Number, err.Description, "frmOpciones.Init", Erl)
    Resume Next
    
End Sub

Private Sub Fullscreen_Click()
   PantallaCompleta = Not PantallaCompleta

    If PantallaCompleta Then
        Fullscreen.Picture = LoadInterface("check-amarillo.bmp")
        Call SetResolution
    Else
        Fullscreen.Picture = Nothing
        Call ResetResolution
    End If
End Sub

Private Sub HScroll1_Change()
    
    On Error GoTo HScroll1_Change_Err
    
    Sound.Ambient_Volume_Set HScroll1.Value
    VolAmbient = HScroll1.Value

    
    Exit Sub

HScroll1_Change_Err:
    Call RegistrarError(err.Number, err.Description, "frmOpciones.HScroll1_Change", Erl)
    Resume Next
    
End Sub



Private Sub instagram_Click()
    
    On Error GoTo instagram_Click_Err
    
    ShellExecute Me.hWnd, "open", "https://argentumworld?.com.ar/", "", "", 0

    
    Exit Sub

instagram_Click_Err:
    Call RegistrarError(err.Number, err.Description, "frmOpciones.instagram_Click", Erl)
    Resume Next
    
End Sub


Private Sub Label3_Click()

End Sub

Private Sub lblIdioma_Click()

End Sub


Private Sub MinimapOpacity_Change()
    Minimap.Opacity = MinimapOpacity.Value
End Sub

Private Sub num_comp_inv_Click()
    If NumerosCompletosInventario = 0 Then
        NumerosCompletosInventario = 1
        num_comp_inv.Picture = LoadInterface("check-amarillo.bmp")
    Else
        NumerosCompletosInventario = 0
        num_comp_inv.Picture = Nothing
    End If
End Sub

Private Sub Respiracion_Click()
    MostrarRespiracion = Not MostrarRespiracion

    If MostrarRespiracion Then
        Respiracion.Picture = LoadInterface("check-amarillo.bmp")
    Else
        Respiracion.Picture = Nothing
    End If
End Sub

Private Sub scrMidi_Change()
    
    On Error GoTo scrMidi_Change_Err
    

    If Musica <> CONST_DESHABILITADA Then
        Sound.Music_Volume_Set scrMidi.Value
        Sound.VolumenActualMusicMax = scrMidi.Value
        VolMusic = Sound.VolumenActualMusicMax

    End If

    
    Exit Sub

scrMidi_Change_Err:
    Call RegistrarError(err.Number, err.Description, "frmOpciones.scrMidi_Change", Erl)
    Resume Next
    
End Sub

Private Sub scrSens_Change()
    
    On Error GoTo scrSens_Change_Err
    
    MouseS = scrSens.Value
    SensibilidadMouse = MouseS
    Call General_Set_Mouse_Speed(MouseS)
    txtMSens.Caption = scrSens.Value

    
    Exit Sub

scrSens_Change_Err:
    Call RegistrarError(err.Number, err.Description, "frmOpciones.scrSens_Change", Erl)
    Resume Next
    
End Sub

Private Sub scrVolume_Change()
    
    On Error GoTo scrVolume_Change_Err
    
    Sound.VolumenActual = scrVolume.Value
    VolFX = Sound.VolumenActual

    
    Exit Sub

scrVolume_Change_Err:
    Call RegistrarError(err.Number, err.Description, "frmOpciones.scrVolume_Change", Erl)
    Resume Next
    
End Sub

Private Sub lblLineasConsolaModificador_Click(Index As Integer)
    If Index = 0 Then 'resto
        If Val(txlCantidadTextosConsola.Text) > 4 Then
            txlCantidadTextosConsola.Text = Val(txlCantidadTextosConsola.Text) - 1
        End If
    Else 'sumo
        If Val(txlCantidadTextosConsola.Text) < 15 Then
            txlCantidadTextosConsola.Text = Val(txlCantidadTextosConsola.Text) + 1
        End If
    End If
    
    Call WriteVar(RESOURCES_PATH & "/OUTPUT/Configuracion.ini", "OPCIONES", "DialogosEnConsola", txlCantidadTextosConsola.Text)
    ConsoleDialog.RENDER_DIALOGS = Val(txlCantidadTextosConsola.Text)
    ConsoleDialog.RENDER_DIALOGS_MAX = Val(txlCantidadTextosConsola.Text)
    ConsoleDialog.console_height = 14 * ConsoleDialog.RENDER_DIALOGS + 14
    ConsoleDialog.SetupConsoleTexture
    frmMain.SendTxt.Top = frmMain.renderer.Top + ConsoleDialog.console_height
    frmMain.SendTxtCmsg.Top = frmMain.renderer.Top + ConsoleDialog.console_height
End Sub


Private Sub VSync_Click()
    VSyncActivado = Not VSyncActivado
    Call engine.engine_init
    prgRun = True
    pausa = False
    QueRender = 0
     FramesPerSecCounter = 60
    
    If VSyncActivado Then
        VSync.Picture = LoadInterface("check-amarillo.bmp")
    Else
        VSync.Picture = Nothing
    End If
End Sub
