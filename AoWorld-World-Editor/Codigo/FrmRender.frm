VERSION 5.00
Begin VB.Form FrmRender 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   15615
   ClientLeft      =   9495
   ClientTop       =   2025
   ClientWidth     =   19035
   LinkTopic       =   "Form4"
   ScaleHeight     =   1041
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1269
   Visible         =   0   'False
   Begin VB.PictureBox PicGrande 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   3600
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   144
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   12360
      Width           =   2160
   End
   Begin VB.CheckBox chkQuitarNPCs 
      BackColor       =   &H80000016&
      Caption         =   "Quitar NPCs , Bloq y Exit del Borde"
      Height          =   255
      Left            =   3240
      TabIndex        =   22
      Top             =   1080
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton cmdArbolesFix 
      Caption         =   "Informar error del mapa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   19
      Top             =   3960
      Width           =   3045
   End
   Begin VB.Timer SaveallMapaFix 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2640
      Top             =   0
   End
   Begin VB.OptionButton OptMinimapas100 
      BackColor       =   &H80000016&
      Caption         =   "Minimapas 100*100"
      Height          =   195
      Left            =   840
      TabIndex        =   18
      Top             =   6000
      Width           =   2055
   End
   Begin VB.OptionButton OptMapasMundo 
      BackColor       =   &H80000016&
      Caption         =   "Mapas Mundo"
      Height          =   255
      Left            =   1920
      TabIndex        =   17
      Top             =   5640
      Width           =   1575
   End
   Begin VB.OptionButton OptMinimapasWE 
      BackColor       =   &H80000016&
      Caption         =   "Minimapas WE"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   5640
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.Timer ListadoNPCs 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3000
      Top             =   0
   End
   Begin VB.CommandButton cmdListadoDe 
      Caption         =   "Listado de NPCs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   14
      Top             =   3480
      Width           =   3045
   End
   Begin VB.CheckBox chkArreglarTodo 
      BackColor       =   &H80000016&
      Caption         =   "Arreglar todo"
      Height          =   255
      Left            =   6960
      TabIndex        =   13
      Top             =   2160
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdVerErrores 
      Caption         =   "Ver Errores"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   12
      Top             =   3000
      Width           =   3045
   End
   Begin VB.CheckBox chkCasas 
      BackColor       =   &H80000016&
      Caption         =   "Cabañas New"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1080
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CheckBox chkFaroles 
      BackColor       =   &H80000016&
      Caption         =   "Faroles y particulas - Hogar a leña de Comercios"
      Height          =   255
      Left            =   3240
      TabIndex        =   10
      Top             =   720
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CheckBox chkCarteles 
      BackColor       =   &H80000016&
      Caption         =   "Carteles, Objetos"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CheckBox chkPuertas 
      BackColor       =   &H80000016&
      Caption         =   "Puertas"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox chkBloqueosSin 
      BackColor       =   &H80000016&
      Caption         =   "Bloqueos sin acceso"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CheckBox chkNPCsSin 
      BackColor       =   &H80000016&
      Caption         =   "NPCs sin Body o sin GRHs"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CheckBox chkArboles 
      BackColor       =   &H80000016&
      Caption         =   "Arboles , Plantas y Fogatas"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   2160
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton cmdBuscarErrores 
      Caption         =   "Buscar Errores"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   3045
   End
   Begin VB.Timer SaveAllErrores 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1200
      Top             =   360
   End
   Begin VB.Timer SaveAll 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2040
      Top             =   0
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Renderizar minimapas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   3
      Top             =   6360
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Render mapa completo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   2
      Top             =   7320
      Width           =   3015
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Render desde el 2+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   1
      Top             =   7920
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   0
      Top             =   4440
      Width           =   3045
   End
   Begin VB.CheckBox chkGraficosDe 
      BackColor       =   &H80000016&
      Caption         =   "Graficos de arbol"
      Height          =   195
      Left            =   3240
      TabIndex        =   20
      Top             =   1440
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CheckBox chkLuzFalor 
      BackColor       =   &H80000016&
      Caption         =   "Luz Farol"
      Height          =   255
      Left            =   3240
      TabIndex        =   21
      Top             =   1800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox picMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   12000
      Left            =   3600
      ScaleHeight     =   625
      ScaleMode       =   0  'User
      ScaleWidth      =   625
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   120
      Width           =   12000
   End
End
Attribute VB_Name = "FrmRender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'*************************************************************
' Capturar la imagen de controles
       
'  1 - Colocar un picturebox llamado picture1, un Command1 y un Command2 _
   2 - Agragar algunos controles _
   3 - Indicar en la Sub " Capturar_Imagen " .. el control a capturar
'*************************************************************
      
' Declaraciones del Api
      
'*************************************************************
' Función BitBlt para copiar la imagen del control en un picturebox
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
      
' Recupera la imagen del área del control
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long

    Dim handle As Integer
    Dim iz As Integer
    Dim de As Integer
    Dim ar As Integer
    Dim ab As Integer
    Dim Norte As Integer
    Dim Sur As Integer
    Dim Este As Integer
    Dim Oeste As Integer
    Dim MapSup

Private Sub chkArreglarTodo_Click()

    If chkArreglarTodo.value = 1 Then
        chkGraficosDe.value = 1
        chkBloqueosSin.value = 1
        chkArboles.value = 1
        chkNPCsSin.value = 1
        chkPuertas.value = 1
        chkCarteles.value = 1
        chkFaroles.value = 1
        chkCasas.value = 1
        chkLuzFalor.value = 1
        chkQuitarNPCs.value = 1
        cmdBuscarErrores.Caption = "Reparar Errores"
    Else
        chkGraficosDe.value = 0
        chkBloqueosSin.value = 0
        chkArboles.value = 0
        chkNPCsSin.value = 0
        chkPuertas.value = 0
        chkCarteles.value = 0
        chkFaroles.value = 0
        chkCasas.value = 0
        chkLuzFalor.value = 0
        chkQuitarNPCs.value = 0
        cmdBuscarErrores.Caption = "Buscar Errores"
    End If

End Sub
Private Sub chkArboles_Click()
If chkLuzFalor.value = 1 Or chkBloqueosSin.value = 1 Or chkArboles.value = 1 Or chkNPCsSin.value = 1 Or chkPuertas.value = 1 Or chkCarteles.value = 1 Or chkFaroles.value = 1 Or chkCasas.value = 1 Then
    cmdBuscarErrores.Caption = "Reparar Errores"
Else
    cmdBuscarErrores.Caption = "Buscar Errores"
End If
End Sub
Private Sub chkBloqueosSin_Click()
If chkLuzFalor.value = 1 Or chkBloqueosSin.value = 1 Or chkArboles.value = 1 Or chkNPCsSin.value = 1 Or chkPuertas.value = 1 Or chkCarteles.value = 1 Or chkFaroles.value = 1 Or chkCasas.value = 1 Then
    cmdBuscarErrores.Caption = "Reparar Errores"
Else
    cmdBuscarErrores.Caption = "Buscar Errores"
End If
End Sub

Private Sub chkCarteles_Click()
If chkLuzFalor.value = 1 Or chkBloqueosSin.value = 1 Or chkArboles.value = 1 Or chkNPCsSin.value = 1 Or chkPuertas.value = 1 Or chkCarteles.value = 1 Or chkFaroles.value = 1 Or chkCasas.value = 1 Then
    cmdBuscarErrores.Caption = "Reparar Errores"
Else
    cmdBuscarErrores.Caption = "Buscar Errores"
End If
End Sub

Private Sub chkCasas_Click()
If chkLuzFalor.value = 1 Or chkBloqueosSin.value = 1 Or chkArboles.value = 1 Or chkNPCsSin.value = 1 Or chkPuertas.value = 1 Or chkCarteles.value = 1 Or chkFaroles.value = 1 Or chkCasas.value = 1 Then
    cmdBuscarErrores.Caption = "Reparar Errores"
Else
    cmdBuscarErrores.Caption = "Buscar Errores"
End If
End Sub

Private Sub chkFaroles_Click()
If chkBloqueosSin.value = 1 Or chkArboles.value = 1 Or chkNPCsSin.value = 1 Or chkPuertas.value = 1 Or chkCarteles.value = 1 Or chkFaroles.value = 1 Or chkCasas.value = 1 Then
    cmdBuscarErrores.Caption = "Reparar Errores"
Else
    cmdBuscarErrores.Caption = "Buscar Errores"
End If
End Sub

Private Sub chkLuzFalor_Click()
If chkLuzFalor.value = 1 Or chkBloqueosSin.value = 1 Or chkArboles.value = 1 Or chkNPCsSin.value = 1 Or chkPuertas.value = 1 Or chkCarteles.value = 1 Or chkFaroles.value = 1 Or chkCasas.value = 1 Then
    cmdBuscarErrores.Caption = "Reparar Errores"
Else
    cmdBuscarErrores.Caption = "Buscar Errores"
End If
End Sub

Private Sub chkNPCsSin_Click()
If chkLuzFalor.value = 1 Or chkBloqueosSin.value = 1 Or chkArboles.value = 1 Or chkNPCsSin.value = 1 Or chkPuertas.value = 1 Or chkCarteles.value = 1 Or chkFaroles.value = 1 Or chkCasas.value = 1 Then
    cmdBuscarErrores.Caption = "Reparar Errores"
Else
    cmdBuscarErrores.Caption = "Buscar Errores"
End If
End Sub

Private Sub chkPuertas_Click()
If chkLuzFalor.value = 1 Or chkBloqueosSin.value = 1 Or chkArboles.value = 1 Or chkNPCsSin.value = 1 Or chkPuertas.value = 1 Or chkCarteles.value = 1 Or chkFaroles.value = 1 Or chkCasas.value = 1 Then
    cmdBuscarErrores.Caption = "Reparar Errores"
Else
    cmdBuscarErrores.Caption = "Buscar Errores"
End If
End Sub


Private Sub chkSoloMapa_Click()

'    If chkSoloMapa.Value = 1 And chkReparar.Value = 1 Then
'        chkGraficosDe.Visible = True
'        chkArboles.Visible = True
'        chkLuzFalor.Visible = True
'      Else
'        chkGraficosDe.Visible = False
'        chkArboles.Visible = False
'        chkLuzFalor.Visible = False
'    End If

End Sub

Private Sub cmdAceptar_Click()


Dim i As Integer

For i = 2 To FrmMain.lstMaps.ListCount
    UserMap = mid$(lstMaps.List(i), 1, InStr(1, lstMaps.List(i), " ") - 1)
    FrmMain.Label16.Caption = "Map " & UserMap
    modMapIO.AbrirMapa App.Path & "\..\Resources\Mapas\Mapa" & i & ".csm"
    DoEvents
    Command2_Click
Next i

    
End Sub

'*************************************************************
' Sub que copia la imagen del control en un picturebox
'*************************************************************
Public Sub Capturar_Imagen(Control As Control, Destino As Object)
          
    Dim hdc             As Long
    Dim Escala_Anterior As Integer
    Dim Ancho           As Long
    Dim Alto            As Long
          
    ' Para que se mantenga la imagen por si se repinta la ventana
    Destino.AutoRedraw = True
          
    On Error Resume Next

    ' Si da error es por que el control está dentro de un Frame _
      ya que  los Frame no tiene  dicha propiedad
    Escala_Anterior = Control.Container.ScaleMode
          
    If Err.Number = 438 Then
        ' Si el control está en un Frame, convierte la escala
        Ancho = ScaleX(Control.Width, vbTwips, vbPixels)
        Alto = ScaleY(Control.Height, vbTwips, vbPixels)
    Else
        ' Si no cambia la escala del  contenedor a pixeles
        Control.Container.ScaleMode = vbPixels
        Ancho = Control.Width
        Alto = Control.Height

    End If
          
    ' limpia el error
    On Error GoTo 0

    ' Captura el área de pantalla correspondiente al control
    hdc = GetWindowDC(Control.hWnd)
    
    ' Copia esa área al picturebox
    If ToWorldMap2 Then
        'Call BitBlt(Destino.hdc, 0 - 50, 0 - 50, Ancho - 50, Alto - 50, hdc, 0, 0, vbSrcCopy) '
        Call BitBlt(Destino.hdc, 0, 0, Ancho, Alto, hdc, 0, 0, vbSrcCopy)
    Else
        Call BitBlt(Destino.hdc, 0, 0, 3000, 3000, hdc, 0, 0, vbSrcCopy)
        

    End If
    
    ' Convierte la imagen anterior en un Mapa de bits
    Destino.Picture = Destino.Image
    
    ' Borra la imagen ya que ahora usa el Picture
    Call Destino.Cls
          
    On Error Resume Next

    If Err.Number = 0 Then
        ' Si el control no está en un  Frame, restaura la escala del contenedor
        Control.Container.ScaleMode = Escala_Anterior

    End If
          
End Sub

Private Sub cmdArbolesFix_Click()

    If MapInfo.Changed = 1 Then
        If MsgBox("Este mapa fue modificado. ¿Guardo los cambios?", vbYesNo, "Reparar") = vbYes Then
            Call modMapIO.GuardarMapa(PATH_Save & MapName)
        Else
            Exit Sub
        End If
    End If
   
    Dim Filename As String
        

    If FileExist(Filename, vbArchive) = False Then
        Unload Me
        MsgBox "Primero abre algún mapa de la carpeta a convertir.", vbOKOnly, "Error"
        Exit Sub
    End If
    
    FrmRender.Height = 2900
    FrmRender.Top = 7200
    
    Call AbrirMapa(Filename)
    
    SaveallMapaFix.Enabled = True
    
    handle = FreeFile

    If Dir(App.Path & "\errores.txt", vbArchive) <> "" Then
        Kill (App.Path & "\errores.txt")
    End If

    Open App.Path & "\errores.txt" For Append As #handle
End Sub

Private Sub cmdBuscarErrores_Click()

    If MapInfo.Changed = 1 Then
        If MsgBox("Este mapa fue modificado. ¿Guardo los cambios?", vbYesNo, "Reparar") = vbYes Then
            Call modMapIO.GuardarMapa(PATH_Save & MapName)
        Else
            Exit Sub
        End If
    End If
   
    Dim Filename As String
    


    If FileExist(Filename, vbArchive) = False Then
        Unload Me
        MsgBox "Primero abre algún mapa de la carpeta a convertir.", vbOKOnly, "Error"
        Exit Sub
    End If
    
    FrmRender.Height = 2900
    FrmRender.Top = 7200
    
    Call AbrirMapa(Filename)
    
     SaveAllErrores.Enabled = True
    
    handle = FreeFile

    If Dir(App.Path & "\errores.txt", vbArchive) <> "" Then
        Kill (App.Path & "\errores.txt")
    End If

    Open App.Path & "\errores.txt" For Append As #handle
    
End Sub

Private Sub cmdListadoDe_Click()

    Dim Filename As String
    Filename = PATH_Save & NameMap_Save & "1.csm"

    If FileExist(Filename, vbArchive) = False Then
        Unload Me
        MsgBox "Primero abre algún mapa de la carpeta a convertir.", vbOKOnly, "Error"
        Exit Sub
    End If
    
    FrmRender.Height = 2900
    FrmRender.Top = 7200
    
    Call AbrirMapa(Filename)
    
    ListadoNPCs.Enabled = True
    
    handle = FreeFile

    If Dir(App.Path & "\ListadoNPCs.txt", vbArchive) <> "" Then
        Kill (App.Path & "\ListadoNPCs.txt")
    End If

    Open App.Path & "\ListadoNPCs.txt" For Append As #handle


End Sub

Private Sub cmdVerErrores_Click()
Shell "C:\WINDOWS\System32\notepad.exe " & App.Path & "\errores.txt", vbNormalFocus
End Sub

Private Sub Command1_Click()
    
    On Error GoTo Command1_Click_Err
    
    Unload Me

    
    Exit Sub

Command1_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmRender.Command1_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command2_Click()
Dim tmpPic As StdPicture
   Dim picNr As Long
   Dim sFileName As String
   Dim maxCx As Long, maxCy As Long
   Dim picWidth As Long, picHeight As Long
   PicGrande.ScaleMode = vbPixels
   PicGrande.AutoRedraw = True
   PicGrande.BorderStyle = 0&
   Dim X2 As Integer
   Dim Y2 As Integer
   
'UserMap = 1
   
    maxCy = MapSize.Width * 4
    maxCx = MapSize.Height * 4
    'PicGrande.Move PicGrande.Left, PicGrande.Top, ScaleX(maxCx, vbPixels, Me.ScaleMode), ScaleY(maxCy, vbPixels, Me.ScaleMode)
    
    
    
    
   For Y2 = 0 To ((MapSize.Height - 1) \ 100)
      For X2 = 0 To ((MapSize.Width - 1) \ 100)
        Set SurfaceDB.Used = New Collection
        Call engine.MapCapture(False, True, X2 * 100, Y2 * 100)
        SurfaceDB.UnloadUnused
     Next X2
  Next Y2
  'SavePicture PicGrande.Image, App.Path & "\Render\Grande.bmp"
  PicGrande.AutoRedraw = False
  
  
  Shell (App.Path & "\UnirMinimapa.exe " & UserMap)
End Sub

Public Sub Command3_Click()

    Dim Filename As String
    
    FrmMain.cVerBloqueos.value = (FrmMain.cVerBloqueos.value = False)
    'FrmMain.mnuVerBloqueos.Checked = FrmMain.cVerBloqueos.Value
    
    FrmMain.mnuVerTranslados.Checked = (FrmMain.mnuVerTranslados.Checked = False)
    
    FrmMain.cVerTriggers.value = (FrmMain.cVerTriggers.value = False)
    'FrmMain.mnuVerTriggers.Checked = FrmMain.cVerTriggers.Value
            
    Filename = PATH_Save & NameMap_Save & "1.csm"



    If FileExist(Filename, vbArchive) = False Then
        Unload Me
        MsgBox "Primero abre algún mapa de la carpeta a convertir.", vbOKOnly, "Error"
        Exit Sub
    End If
    
    Call AbrirMapa(Filename)
    
    SaveAll.Enabled = True

End Sub

Public Function IsNorte(ByVal X As Integer, ByVal Y As Integer) As Boolean

    If MapData(X, Y).Blocked = 1 Or MapData(X, Y).Blocked = 3 Or MapData(X, Y).Blocked = 5 Or MapData(X, Y).Blocked = 7 Or MapData(X, Y).Blocked = 9 Or MapData(X, Y).Blocked = 11 Or MapData(X, Y).Blocked = 13 Or MapData(X, Y).Blocked = 15 Then
        IsNorte = True
        Norte = 4
        Exit Function
    Else:
        IsNorte = False
        Exit Function
    End If
    
End Function
Public Function IsSur(ByVal X As Integer, ByVal Y As Integer) As Boolean

    If MapData(X, Y).Blocked = 4 Or MapData(X, Y).Blocked = 5 Or MapData(X, Y).Blocked = 6 Or MapData(X, Y).Blocked = 7 Or MapData(X, Y).Blocked = 12 Or MapData(X, Y).Blocked = 13 Or MapData(X, Y).Blocked = 14 Or MapData(X, Y).Blocked = 15 Then
        Sur = 1
        IsSur = True
        Exit Function
    Else
        IsSur = False
        Exit Function
    End If
    
End Function
Public Function IsEste(ByVal X As Integer, ByVal Y As Integer) As Boolean

    If MapData(X, Y).Blocked = 2 Or MapData(X, Y).Blocked = 3 Or MapData(X, Y).Blocked = 6 Or MapData(X, Y).Blocked = 7 Or MapData(X, Y).Blocked = 10 Or MapData(X, Y).Blocked = 11 Or MapData(X, Y).Blocked = 14 Or MapData(X, Y).Blocked = 15 Then
        Este = 8
        IsEste = True
        Exit Function
    Else
        IsEste = False
        Exit Function
    End If
    
End Function
Public Function IsOeste(ByVal X As Integer, ByVal Y As Integer) As Boolean

    If MapData(X, Y).Blocked = 8 Or MapData(X, Y).Blocked = 9 Or MapData(X, Y).Blocked = 10 Or MapData(X, Y).Blocked = 11 Or MapData(X, Y).Blocked = 12 Or MapData(X, Y).Blocked = 13 Or MapData(X, Y).Blocked = 14 Or MapData(X, Y).Blocked = 15 Then
        Oeste = 2
        IsOeste = True
        Exit Function
    Else
        IsOeste = False
        Exit Function
    End If
    
End Function
Private Function IsBlock(ByVal X As Integer, ByVal Y As Integer) As Boolean

    If X - 1 < 1 Or X + 1 > MapSize.Width Then
        IsBlock = True
        Exit Function
    End If
    
    If Y - 1 < 1 Or Y + 1 > MapSize.Height Then
        IsBlock = True
        Exit Function
    End If
    
    IsBlock = (MapData(X, Y).Blocked And &HF) = &HF
    
End Function

Private Sub SaveAllErrores_Timer()

    Dim X As Integer, Y As Integer, BordeX As Integer, BordeY As Integer
    
    Call NPCsBordes(X, Y)
    
    BordeX = 11
    BordeY = 8

    '*****************************************************************************************************
    '*****************************************************************************************************
    'Grhs y Objetos en el mapa
    '*****************************************************************************************************
    '*****************************************************************************************************

    '******************************************************************************************************
    ' Autocompletar Faroles y hogar a leña By ReyarB
    '******************************************************************************************************
    For Y = 1 + BordeY To MapSize.Height - BordeY
        For X = 1 + BordeX To MapSize.Width - BordeX
        
        
            Call FixLuces(X, Y, MapData(X, Y).luz.Rango, MapData(X, Y).luz.color, MapData(X, Y).Graphic(3).grhindex, MapData(X, Y).particle_Index)
            'Call cabañas(X, y, MapData(X, y).Graphic(3).grhindex)
        
            If (MapData(X, Y).Graphic(3).grhindex >= 12747 And MapData(X, Y).Graphic(3).grhindex <= 12750) Or (MapData(X, Y).Graphic(3).grhindex = 49546 Or MapData(X, Y).Graphic(3).grhindex = 28223 Or MapData(X, Y).Graphic(3).grhindex = 2093 Or MapData(X, Y).Graphic(3).grhindex = 2460 Or MapData(X, Y).Graphic(3).grhindex = 2407 Or MapData(X, Y).Graphic(3).grhindex = 2408 Or MapData(X, Y).Graphic(3).grhindex = 12213 Or MapData(X, Y).Graphic(3).grhindex = 12137) Or _
               MapData(X, Y).Graphic(3).grhindex = 9217 Or MapData(X, Y).Graphic(3).grhindex = 9218 Or MapData(X, Y).Graphic(3).grhindex = 49550 Or MapData(X, Y).Graphic(3).grhindex = 49549 Or MapData(X, Y).Graphic(3).grhindex = 12198 Or MapData(X, Y).Graphic(3).grhindex = 12199 Or MapData(X, Y).Graphic(3).grhindex = 12122 Or MapData(X, Y).Graphic(3).grhindex = 12123 Or MapData(X, Y).Graphic(3).grhindex = 49547 Or MapData(X, Y).Graphic(3).grhindex = 9206 Or MapData(X, Y).Graphic(3).grhindex = 12120 Or MapData(X, Y).Graphic(3).grhindex = 12196 Or (MapData(X, Y).Graphic(3).grhindex >= 5624 And MapData(X, Y).Graphic(3).grhindex <= 5627) Or MapData(X, Y).Graphic(3).grhindex = 9205 Or MapData(X, Y).Graphic(3).grhindex = 49546 Or MapData(X, Y).Graphic(3).grhindex = 12213 Or MapData(X, Y).Graphic(3).grhindex = 12137 Or (MapData(X, Y).Graphic(3).grhindex >= 50815 And MapData(X, Y).Graphic(3).grhindex <= 50876) Or (MapData(X, Y).Graphic(3).grhindex >= 47629 And MapData(X, Y).Graphic(3).grhindex <= 47692) Then

                If X > (1 + BordeX) And X < (MapSize.Width - BordeX) And Y > (1 + BordeY) And Y < (MapSize.Height - BordeY) Then

                Call Faroles(X, Y)
                
                End If
            
                '***************************************************************************************
                'Bloqeos casas Paredes DER by ReyarB
                '***************************************************************************************
    
                If MapData(X, Y).Graphic(3).grhindex = 9218 Or MapData(X, Y).Graphic(3).grhindex = 49550 Or MapData(X, Y).Graphic(3).grhindex = 12199 Or MapData(X, Y).Graphic(3).grhindex = 12123 Then
                    Dim Costado As Integer
                    Dim i       As Integer
                    
                    If chkCasas.value = 1 Then
                        iz = 0

                        For i = 1 To 6

                            If IsNorte(X, Y + 3 - i) Then iz = iz + Norte
                            If IsSur(X, Y + 5 - i) Then iz = iz + Sur
                            If IsEste(X + 1, Y + 4 - i) Then iz = iz + Este
                            If iz = 16 Then iz = 0
                            
                            If MapData(X, Y + 4 - i).Blocked <> (2 + iz) And Not MapData(X, Y + 4 - i).Blocked = 15 Then
                                Print #handle, MapName & " ::: Faltan bloqueos en : " & X & ", " & Y + 4 - i & " ::::  SE podia entrar faltaba bloqueo en la Pared de la casa."
                                MapData(X, Y + 4 - i).Blocked = (2 + iz)
                                MapInfo.Changed = 1
                            End If
                            iz = 0
                        Next
                        iz = 0

                        For i = 1 To 4

                            If IsNorte(X - 1, Y + 2 - i) Then iz = iz + Norte
                            If IsSur(X - 1, Y + 4 - i) Then iz = iz + Sur
                            If IsOeste(X - 2, Y + 3 - i) Then iz = iz + Oeste
                            If iz = 16 Then iz = 0
                        
                            If MapData(X - 1, Y + 3 - i).Blocked <> (8 + iz) And Not MapData(X - 1, Y + 3 - i).Blocked = 15 Then
                                Print #handle, MapName & " ::: Faltan bloqueos en : " & X - 1 & ", " & Y + 3 - i & " ::::  SE podia entrar faltaba bloqueo en la Pared de la casa."
                                MapData(X - 1, Y + 3 - i).Blocked = 8 + iz
                                MapInfo.Changed = 1
                            End If
                            iz = 0
                        Next

                    Else
                        iz = 0

                        For i = 1 To 6

                            If IsNorte(X, Y + 3 - i) Then iz = iz + Norte
                            If IsSur(X, Y + 5 - i) Then iz = iz + Sur
                            If IsEste(X + 1, Y + 4 - i) Then iz = iz + Este
                            If iz = 16 Then iz = 0
                            
                            If MapData(X, Y + 4 - i).Blocked <> (2 + iz) Then
                                If MapData(X, Y + 4 - i).Blocked <> 15 Then
                                    Print #handle, MapName & " ::: Faltan bloqueos en : " & X & ", " & Y + 4 - i & " ::::  SE podia entrar faltaba bloqueo en la Pared de la casa."
                                End If
                            End If
                            iz = 0
                        Next
                        iz = 0

                        For i = 1 To 4

                            If IsNorte(X - 1, Y + 2 - i) Then iz = iz + Norte
                            If IsSur(X - 1, Y + 4 - i) Then iz = iz + Sur
                            If IsOeste(X - 2, Y + 3 - i) Then iz = iz + Oeste
                            If iz = 16 Then iz = 0
                        
                            If MapData(X - 1, Y + 3 - i).Blocked <> (8 + iz) Then
                                If MapData(X - 1, Y + 3 - i).Blocked <> 15 Then
                                    Print #handle, MapName & " ::: Faltan bloqueos en : " & X - 1 & ", " & Y + 3 - i & " ::::  SE podia entrar faltaba bloqueo en la Pared de la casa."
                                End If
                            End If
                            iz = 0
                        Next
    
                    End If
                 
                End If
                
                '***************************************************************************************
                'Bloqeos casas Paredes IZQ by ReyarB
                '***************************************************************************************
    
                If MapData(X, Y).Graphic(3).grhindex = 9217 Or MapData(X, Y).Graphic(3).grhindex = 49549 Or MapData(X, Y).Graphic(3).grhindex = 12198 Or MapData(X, Y).Graphic(3).grhindex = 12122 Then
                    If chkCasas.value = 1 Then
                        iz = 0

                        For i = 1 To 6

                            If IsNorte(X, Y + 3 - i) Then iz = iz + Norte
                            If IsSur(X, Y + 5 - i) Then iz = iz + Sur
                            If IsOeste(X - 1, Y + 4 - i) Then iz = iz + Oeste
                            If iz = 16 Then iz = 0
                            If MapData(X, Y + 4 - i).Blocked <> (8 + iz) And Not MapData(X, Y + 4 - i).Blocked = 15 Then
                                Print #handle, MapName & " ::: Faltan bloqueos en : " & X & ", " & Y + 4 - i & " ::::  SE PUEDE Entrar Falta bloqueo en la Pared de la casa."
                                MapData(X, Y + 4 - i).Blocked = 8 + iz
                                MapInfo.Changed = 1
                            End If
                            iz = 0
                        Next
                        iz = 0

                        For i = 1 To 4
                            
                            If IsNorte(X + 1, Y + 2 - i) Then iz = iz + Norte
                            If IsSur(X + 1, Y + 4 - i) Then iz = iz + Sur
                            If IsEste(X + 2, Y + 3 - i) Then iz = iz + Este
                        
                            If MapData(X + 1, Y + 3 - i).Blocked <> (2 + iz) And Not MapData(X + 1, Y + 3 - i).Blocked = 15 Then
                                Print #handle, MapName & " ::: Faltan bloqueos en : " & X + 1 & ", " & Y + 3 - i & " ::::  SE PUEDE Entrar Falta bloqueo en la Pared de la casa."
                                MapData(X + 1, Y + 3 - i).Blocked = 2 + iz
                                MapInfo.Changed = 1
                            End If
                            iz = 0
                        Next

                    Else
                        
                        iz = 0

                        For i = 1 To 6

                            If IsNorte(X, Y + 3 - i) Then iz = iz + Norte
                            If IsSur(X, Y + 5 - i) Then iz = iz + Sur
                            If IsOeste(X - 1, Y + 4 - i) Then iz = iz + Oeste
                            If iz = 16 Then iz = 0
                            If MapData(X, Y + 4 - i).Blocked <> (8 + iz) And Not MapData(X, Y + 4 - i).Blocked = 15 Then
                                Print #handle, MapName & " ::: Faltan bloqueos en : " & X & ", " & Y + 4 - i & " ::::  SE PUEDE Entrar Falta bloqueo en la Pared de la casa."
                            End If
                            iz = 0
                        Next
                        iz = 0

                        For i = 1 To 4
                            
                            If IsNorte(X + 1, Y + 2 - i) Then iz = iz + Norte
                            If IsSur(X + 1, Y + 4 - i) Then iz = iz + Sur
                            If IsEste(X + 2, Y + 3 - i) Then iz = iz + Este
                        
                            If MapData(X + 1, Y + 3 - i).Blocked <> (2 + iz) And Not MapData(X + 1, Y + 3 - i).Blocked = 15 Then
                                Print #handle, MapName & " ::: Faltan bloqueos en : " & X + 1 & ", " & Y + 3 - i & " ::::  SE PUEDE Entrar Falta bloqueo en la Pared de la casa."
                            End If
                            iz = 0
                        Next
    
                    End If

                End If
                
                '***************************************************************************************
                'Bloqeos casas atras
                '***************************************************************************************

                If MapData(X, Y).Graphic(3).grhindex = 9205 Or MapData(X, Y).Graphic(3).grhindex = 49546 Or MapData(X, Y).Graphic(3).grhindex = 12213 Or MapData(X, Y).Graphic(3).grhindex = 12137 And MapData(X, Y + 1).Trigger < 50 Then

                    If chkCasas.value = 1 Then

                        For i = 1 To 8
                            iz = 0

                            If IsNorte(X + 4 - i, Y - 1) Then iz = iz + Norte
                            If IsEste(X + 5 - i, Y) Then iz = iz + Este
                            If IsOeste(X + 3 - i, Y) Then iz = iz + Oeste

                            If MapData(X + 4 - i, Y).Blocked <> (1 + iz) And Not MapData(X + 4 - i, Y).Blocked = 15 Then
                                Print #handle, MapName & " ::: Faltan bloqueos en : " & X + 4 - i & ", " & Y & " ::::  SE PUEDE SALIR POR ATRAS. Falta bloqueo atras de la casa."
                                MapData(X + 4 - i, Y).Blocked = (1 + iz)
                                MapInfo.Changed = 1
                            End If
                            iz = 0

                            If IsSur(X + 4 - i, Y + 2) Then iz = iz + Sur
                            If IsEste(X + 5 - i, Y + 1) Then iz = iz + Este
                            If IsOeste(X + 3 - i, Y + 1) Then iz = iz + Oeste

                            If MapData(X + 4 - i, Y + 1).Blocked <> (4 + iz) And Not MapData(X + 4 - i, Y + 1).Blocked = 15 Then
                                Print #handle, MapName & " ::: Faltan bloqueos en : " & X + 4 - i & ", " & Y + 1 & " ::::  SE PUEDE ENTRAR POR ATRAS. Falta bloqueo dentro de la casa."
                                MapData(X + 4 - i, Y + 1).Blocked = 4 + iz
                                MapInfo.Changed = 1

                            End If
                            iz = 0
                        Next i

                    Else

                        For i = 1 To 8
                            iz = 0

                            If IsNorte(X + 4 - i, Y - 1) Then iz = iz + Norte
                            If IsEste(X + 5 - i, Y) Then iz = iz + Este
                            If IsOeste(X + 3 - i, Y) Then iz = iz + Oeste

                            If MapData(X + 4 - i, Y).Blocked <> (1 + iz) And Not MapData(X + 4 - i, Y).Blocked = 15 Then
                                Print #handle, MapName & " ::: Faltan bloqueos en : " & X + 4 - i & ", " & Y & " ::::  SE PUEDE SALIR POR ATRAS. Falta bloqueo atras de la casa."

                            End If
                            iz = 0

                            If IsSur(X + 4 - i, Y + 2) Then iz = iz + Sur
                            If IsEste(X + 5 - i, Y + 1) Then iz = iz + Este
                            If IsOeste(X + 3 - i, Y + 1) Then iz = iz + Oeste

                            If MapData(X + 4 - i, Y + 1).Blocked <> (4 + iz) And Not MapData(X + 4 - i, Y + 1).Blocked = 15 Then
                                Print #handle, MapName & " ::: Faltan bloqueos en : " & X + 4 - i & ", " & Y + 1 & " ::::  SE PUEDE ENTRAR POR ATRAS. Falta bloqueo dentro de la casa."

                            End If
                            iz = 0
                        Next i

                    End If

                End If

                '********************************************************************************************
                'Bloqeos casas Frente   49547 9206 12120 12196
                '*******************************************************************************************

                If MapData(X, Y).Graphic(3).grhindex = 49547 Or MapData(X, Y).Graphic(3).grhindex = 9206 Or MapData(X, Y).Graphic(3).grhindex = 12120 Or MapData(X, Y).Graphic(3).grhindex = 12196 Then
                    Dim BloqAbajo As Integer
                    i = 0

                    If chkCasas.value = 1 Then

                        For i = 1 To 8

                            If i = 1 Then BloqAbajo = 15
                            If i = 2 Then BloqAbajo = 12
                            If i = 3 Then BloqAbajo = 6
                            If i = 4 Then BloqAbajo = 12
                            If i = 5 Then BloqAbajo = 4
                            If i = 6 Then BloqAbajo = 4
                            If i = 7 Then BloqAbajo = 4
                            If i = 8 Then BloqAbajo = 4

                            If MapData(X + 4 - i, Y + 1).Blocked <> BloqAbajo Then
                                If MapData(X + 4 - i, Y + 1).Blocked >= 15 Then
                                    iz = 15 - BloqAbajo
                                Else
                                    iz = MapData(X + 4 - i, Y + 1).Blocked
                                End If

                                If (MapData(X + 4 - i, Y + 1).Blocked = 4 Or MapData(X + 4 - i, Y + 1).Blocked = 5 Or MapData(X + 4 - i, Y + 1).Blocked = 6 Or MapData(X + 4 - i, Y + 1).Blocked = 7 Or MapData(X + 4 - i, Y + 1).Blocked = 12 Or MapData(X + 4 - i, Y + 1).Blocked = 13 Or MapData(X + 4 - i, Y + 1).Blocked = 14 Or MapData(X + 4 - i, Y + 1).Blocked = 15) Then
                                    iz = MapData(X + 4 - i, Y + 1).Blocked - BloqAbajo
                                End If

                                If Not (MapData(X + 4 - i, Y + 1).Blocked = 4 Or MapData(X + 4 - i, Y + 1).Blocked = 5 Or MapData(X + 4 - i, Y + 1).Blocked = 6 Or MapData(X + 4 - i, Y + 1).Blocked = 7 Or MapData(X + 4 - i, Y + 1).Blocked = 12 Or MapData(X + 4 - i, Y + 1).Blocked = 13 Or MapData(X + 4 - i, Y + 1).Blocked = 14 Or MapData(X + 4 - i, Y + 1).Blocked = 15) Then

                                    If (ObjData(MapData(X + 2, Y).OBJInfo.ObjIndex).Cerrada = 1 And (i <> 2 Or i <> 3)) Then
                                        Print #handle, MapName & " ::: Faltan bloqueos en : " & X + 4 - i & ", " & Y + 1 & " ::::  SE PUEDE SALIR. Falta bloqueo adelante de la casa."
                                        MapData(X + 4 - i, Y + 1).Blocked = (BloqAbajo + iz)
                                        MapInfo.Changed = 1
                                    Else

                                        If Not (i = 2 Or i = 3) Then Print #handle, MapName & " ::: Faltan bloqueos en : " & X + 4 - i & ", " & Y + 1 & " ::::  SE PUEDE SALIR. Falta bloqueo adelante de la casa."
                                        If Not (i = 2 Or i = 3) Then
                                            MapData(X + 4 - i, Y + 1).Blocked = (BloqAbajo + iz)
                                            MapInfo.Changed = 1
                                        End If
                                    End If
                                End If

                            End If

                            If i = 1 Then BloqAbajo = 9
                            If i = 8 Then BloqAbajo = 3
                            If i >= 2 And i <= 7 Then BloqAbajo = 1
                            If MapData(X + 4 - i, Y).Blocked <> BloqAbajo Then
                                If MapData(X + 4 - i, Y).Blocked >= 15 Then
                                    iz = 15 - BloqAbajo
                                Else
                                    iz = MapData(X + 4 - i, Y).Blocked
                                End If

                                If Not ((ObjData(MapData(X + 2, Y).OBJInfo.ObjIndex).Cerrada = 0 And (i = 2 Or i = 3)) Or MapData(X + 4 - i, Y).Blocked = 4 Or MapData(X + 4 - i, Y + 1).Blocked = 5 Or MapData(X + 4 - i, Y).Blocked = 6 Or MapData(X + 4 - i, Y + 1).Blocked = 7 Or MapData(X + 4 - i, Y + 1).Blocked = 12 Or MapData(X + 4 - i, Y + 1).Blocked = 13 Or MapData(X + 4 - i, Y + 1).Blocked = 14 Or MapData(X + 4 - i, Y + 1).Blocked = 15) Then
                                    Print #handle, MapName & " ::: Faltan bloqueos en : " & X + 4 - i & ", " & Y & " ::::  SE PUEDE ENTRAR POR ADELANTE. Falta bloqueo dentro de la casa."
                                End If

                                If (ObjData(MapData(X + 2, Y).OBJInfo.ObjIndex).Cerrada = 1 And (i <> 2 Or i <> 3)) Then
                                    MapData(X + 4 - i, Y).Blocked = (BloqAbajo + iz)
                                    MapInfo.Changed = 1
                                    Print #handle, MapName & " ::: Faltan bloqueos en : " & X + 4 - i & ", " & Y & " ::::  SE PUEDE ENTRAR POR ADELANTE. Falta bloqueo dentro de la casa."
                                Else

                                    If Not (i = 2 Or i = 3) Then
                                        MapData(X + 4 - i, Y).Blocked = (BloqAbajo + iz)
                                        MapInfo.Changed = 1
                                    End If

                                    If Not (i = 2 Or i = 3) Then Print #handle, MapName & " ::: Faltan bloqueos en : " & X + 4 - i & ", " & Y & " ::::  SE PUEDE ENTRAR POR ADELANTE. Falta bloqueo dentro de la casa."
                                End If
                            End If
                        Next i

                    Else

                        For i = 1 To 8

                            If i = 1 Then BloqAbajo = 15
                            If i = 2 Then BloqAbajo = 12
                            If i = 3 Then BloqAbajo = 6
                            If i = 4 Then BloqAbajo = 12
                            If i = 5 Then BloqAbajo = 4
                            If i = 6 Then BloqAbajo = 4
                            If i = 7 Then BloqAbajo = 4
                            If i = 8 Then BloqAbajo = 4

                            If MapData(X + 4 - i, Y + 1).Blocked <> BloqAbajo Then
                                If MapData(X + 4 - i, Y + 1).Blocked >= 15 Then
                                    iz = 15 - BloqAbajo
                                Else
                                    iz = MapData(X + 4 - i, Y + 1).Blocked
                                End If

                                If (MapData(X + 4 - i, Y + 1).Blocked = 4 Or MapData(X + 4 - i, Y + 1).Blocked = 5 Or MapData(X + 4 - i, Y + 1).Blocked = 6 Or MapData(X + 4 - i, Y + 1).Blocked = 7 Or MapData(X + 4 - i, Y + 1).Blocked = 12 Or MapData(X + 4 - i, Y + 1).Blocked = 13 Or MapData(X + 4 - i, Y + 1).Blocked = 14 Or MapData(X + 4 - i, Y + 1).Blocked = 15) Then
                                    iz = MapData(X + 4 - i, Y + 1).Blocked - BloqAbajo
                                End If

                                If Not (MapData(X + 4 - i, Y + 1).Blocked = 4 Or MapData(X + 4 - i, Y + 1).Blocked = 5 Or MapData(X + 4 - i, Y + 1).Blocked = 6 Or MapData(X + 4 - i, Y + 1).Blocked = 7 Or MapData(X + 4 - i, Y + 1).Blocked = 12 Or MapData(X + 4 - i, Y + 1).Blocked = 13 Or MapData(X + 4 - i, Y + 1).Blocked = 14 Or MapData(X + 4 - i, Y + 1).Blocked = 15) Then

                                    If (ObjData(MapData(X + 2, Y).OBJInfo.ObjIndex).Cerrada = 1 And (i <> 2 Or i <> 3)) Then
                                        Print #handle, MapName & " ::: Faltan bloqueos en : " & X + 4 - i & ", " & Y + 1 & " ::::  SE PUEDE SALIR. Falta bloqueo adelante de la casa."
                                    Else

                                        If Not (i = 2 Or i = 3) Then Print #handle, MapName & " ::: Faltan bloqueos en : " & X + 4 - i & ", " & Y + 1 & " ::::  SE PUEDE SALIR. Falta bloqueo adelante de la casa."
                                    End If
                                End If

                            End If

                            If i = 1 Then BloqAbajo = 9
                            If i = 8 Then BloqAbajo = 3
                            If i >= 2 And i <= 7 Then BloqAbajo = 1
                            If MapData(X + 4 - i, Y).Blocked <> BloqAbajo Then
                                If MapData(X + 4 - i, Y).Blocked >= 15 Then
                                    iz = 15 - BloqAbajo
                                Else
                                    iz = MapData(X + 4 - i, Y).Blocked
                                End If

                                If Not ((ObjData(MapData(X + 2, Y).OBJInfo.ObjIndex).Cerrada = 0 And (i = 2 Or i = 3)) Or MapData(X + 4 - i, Y).Blocked = 4 Or MapData(X + 4 - i, Y + 1).Blocked = 5 Or MapData(X + 4 - i, Y).Blocked = 6 Or MapData(X + 4 - i, Y + 1).Blocked = 7 Or MapData(X + 4 - i, Y + 1).Blocked = 12 Or MapData(X + 4 - i, Y + 1).Blocked = 13 Or MapData(X + 4 - i, Y + 1).Blocked = 14 Or MapData(X + 4 - i, Y + 1).Blocked = 15) Then
                                    Print #handle, MapName & " ::: Faltan bloqueos en : " & X + 4 - i & ", " & Y & " ::::  SE PUEDE ENTRAR POR ADELANTE. Falta bloqueo dentro de la casa."
                                End If

                                If (ObjData(MapData(X + 2, Y).OBJInfo.ObjIndex).Cerrada = 1 And (i <> 2 Or i <> 3)) Then
                                    Print #handle, MapName & " ::: Faltan bloqueos en : " & X + 4 - i & ", " & Y & " ::::  SE PUEDE ENTRAR POR ADELANTE. Falta bloqueo dentro de la casa."
                                Else

                                    If Not (i = 2 Or i = 3) Then Print #handle, MapName & " ::: Faltan bloqueos en : " & X + 4 - i & ", " & Y & " ::::  SE PUEDE ENTRAR POR ADELANTE. Falta bloqueo dentro de la casa."
                                End If
                            End If
                        Next i

                    End If

                End If
                
                '*********************************************************************
                'Hogar a leña de casas grh 49546,12213 by ReyarB
                '*********************************************************************
                If ((MapData(X, Y).Graphic(3).grhindex = 49546 Or MapData(X, Y).Graphic(3).grhindex = 9205 Or MapData(X, Y).Graphic(3).grhindex = 12213 Or MapData(X, Y).Graphic(3).grhindex = 12137) And MapData(X, Y + 1).Trigger < 50) Then
                    If chkCasas.value = 1 Then
                        If MapData(X - 3, Y).particle_Index <> 250 Then
                            Print #handle, MapName & " ::: Posición del Particula: " & X - 3 & ", " & Y & " :::: Se puso la Particula = 250"
                            MapData(X - 3, Y).particle_Index = 250
                            MapInfo.Changed = 1
                        End If
    
                        If MapData(X - 3, Y - 2).particle_Index <> 180 Then
                            Print #handle, MapName & " ::: Posición del Particula: " & X - 3 & ", " & Y - 2 & " :::: Se puso la Particula = 180"
                            MapData(X - 3, Y - 2).particle_Index = 180
                            MapInfo.Changed = 1
                        End If
                    Else

                        If MapData(X - 3, Y).particle_Index <> 250 Then
                            Print #handle, MapName & " ::: Posición del Particula: " & X - 3 & ", " & Y & " :::: FALTA la Particula = 250"
                        End If
    
                        If MapData(X - 3, Y - 2).particle_Index <> 180 Then
                            Print #handle, MapName & " ::: Posición del Particula: " & X - 3 & ", " & Y - 2 & " :::: FALTA la Particula = 180"

                        End If
                    End If
                End If
                
                If MapData(X, Y).Graphic(3).grhindex = 2460 Then
                    Call FarolBander(X, Y, MapData(X, Y).Graphic(3).grhindex, MapData(X, Y).particle_Index)
                End If
                
                
                If MapData(X, Y).Graphic(3).grhindex = 2093 Or MapData(X, Y).Graphic(3).grhindex = 28223 Then
                    Call MagiaGas(X, Y, MapData(X, Y).Graphic(3).grhindex, MapData(X, Y).particle_Index)
                End If
                
                If ((MapData(X, Y).Graphic(3).grhindex = 2407 Or MapData(X, Y).Graphic(3).grhindex = 2408) And MapData(X, Y).Trigger < 50) Then
                    Call HogarLeña(X, Y, MapData(X, Y).Graphic(3).grhindex, MapData(X, Y).particle_Index)
                End If


                '***************************************************************************************
                'Carteles de casas by ReyarB bloq parcial en X,Y X,Y-1
                '***************************************************************************************

                If (MapData(X, Y).Graphic(3).grhindex >= 50815 And MapData(X, Y).Graphic(3).grhindex <= 50876) Or (MapData(X, Y).Graphic(3).grhindex >= 47629 And MapData(X, Y).Graphic(3).grhindex <= 47692) Then

                    If X > (1 + BordeX) And X < (MapSize.Width - BordeX) And Y > (1 + BordeY) And Y < (MapSize.Height - BordeY) Then

                        If chkCarteles.value = 1 Then

                            iz = 0

                            If IsNorte(X, Y - 1) Then iz = iz + Norte
                            If IsEste(X + 1, Y) Then iz = iz + Este
                            If IsOeste(X - 1, Y) Then iz = iz + Oeste
                            
                            If (MapData(X, Y).Blocked <> 1 + iz) Then
                                MapData(X, Y).Blocked = 1 + iz
                                MapInfo.Changed = 1
                                Print #handle, MapName & " ::: Posición del Cartel: " & X & ", " & Y & " :::: Faltaba Bloqueo " & MapData(X, Y).Graphic(3).grhindex & " ya fue puesto."
                            End If
                                        
                            iz = 0

                            If IsSur(X, Y + 2) Then iz = iz + Sur
                            If IsOeste(X - 1, Y + 1) Then iz = iz + Oeste
                            If IsEste(X + 1, Y + 1) Then iz = iz + Este

                            If (MapData(X, Y + 1).Blocked <> 4 + iz) Then
                                Print #handle, MapName & " ::: Posición del Cartel: " & X & ", " & Y + 1 & " :::: Faltaba Bloqueo " & MapData(X, Y).Graphic(3).grhindex & " ya fue puesto."
                                MapData(X, Y + 1).Blocked = 4 + iz
                                MapInfo.Changed = 1
                            End If

                        Else
                            iz = 0

                            If IsNorte(X, Y - 1) Then iz = iz + Norte
                            If IsEste(X + 1, Y) Then iz = iz + Este
                            If IsOeste(X - 1, Y) Then iz = iz + Oeste
                            If (MapData(X, Y).Blocked <> 1 + iz) Then
                                Print #handle, MapName & " ::: Posición del Cartel: " & X & ", " & Y & " :::: Falta Bloqueo " & MapData(X, Y).Graphic(3).grhindex
                            End If
                                        
                            iz = 0

                            If IsSur(X, Y + 2) Then iz = iz + Sur
                            If IsOeste(X - 1, Y + 1) Then iz = iz + Oeste
                            If IsEste(X + 1, Y + 1) Then iz = iz + Este
                            If (MapData(X, Y + 1).Blocked <> 4 + iz) Then
                                Print #handle, MapName & " ::: Posición del Cartel: " & X & ", " & Y + 1 & " :::: Falta Bloqueo " & MapData(X, Y).Graphic(3).grhindex
                            End If
                        End If

                    End If
                End If

                iz = 0
                de = 0
                ar = 0
                ab = 0
                
            End If
        
            '**********************************************************************************************************************
            'Puertas Comunes
            '**********************************************************************************************************************
            If MapData(X, Y).OBJInfo.ObjIndex Then
                If ObjData(MapData(X, Y).OBJInfo.ObjIndex).ObjType = 6 And ObjData(MapData(X, Y).OBJInfo.ObjIndex).Subtipo = 0 Then
                    If chkPuertas.value = 1 Then
                        If X > (1 + BordeX) And X < (MapSize.Width - BordeX) And Y > (1 + BordeY) And Y < (MapSize.Height - BordeY) Then
                        
                            If IsNorte(X + 1, Y - 1) Then iz = iz + Norte
                            'If IsSur(X + 1, y + 1) Then iz = iz + Sur
                            If IsOeste(X, Y) Then iz = iz + Oeste
                            If IsEste(X + 2, Y) Then iz = iz + Este
                            If iz = 16 Then iz = 0
                            
                            If Not IsBlock(X + 1, Y) And Not MapData(X + 1, Y).Blocked = iz + 1 Then

                                MapData(X + 1, Y).Blocked = iz + 1
                                MapInfo.Changed = 1
                            End If
            
                            If ObjData(MapData(X, Y).OBJInfo.ObjIndex).Cerrada = 1 Then
                            
                                If MapData(X, Y).Blocked <> 1 Then
                                    Print #handle, MapName & " ::: Posición: " & X & ", " & Y & " :::: Falto bloqueo parcial al Norte"
                                    MapData(X, Y).Blocked = 1
                                    MapInfo.Changed = 1
                                End If
            
                                If MapData(X - 1, Y).Blocked <> 1 Then
                                    Print #handle, MapName & " ::: Posición: " & X - 1 & ", " & Y & " :::: Falto bloqueo parcial al Norte"
                                    MapData(X - 1, Y).Blocked = 1
                                    MapInfo.Changed = 1
                                End If
                                
            
                                If Not (MapData(X - 1, Y + 1).Blocked = 6 Or MapData(X - 1, Y + 1).Blocked = 4) Then
                                    If MapData(X - 1, Y + 1).Blocked < 0 Then
                                        MapData(X - 1, Y + 1).Blocked = 6
                                        Print #handle, MapName & " ::: Posición: " & X - 1 & ", " & Y + 1 & " :::: Falto bloqueo parcial al Sur"
                                    Else
                                        MapData(X - 1, Y + 1).Blocked = 4
                                        Print #handle, MapName & " ::: Posición: " & X - 1 & ", " & Y + 1 & " :::: Falto bloqueo parcial al Sur"
                                    End If
                                    MapInfo.Changed = 1
                                End If

                                If Not (MapData(X, Y + 1).Blocked = 12 Or MapData(X, Y + 1).Blocked = 4) Then
                                    Print #handle, MapName & " ::: Posición: " & X & ", " & Y + 1 & " :::: Falto bloqueo parcial al Sur"

                                    If MapData(X, Y + 1).Blocked < 0 Then
                                        MapData(X, Y + 1).Blocked = 12
                                    Else
                                        MapData(X, Y + 1).Blocked = 4
                                    End If
                                    MapInfo.Changed = 1
                                End If

                                If Not (MapData(X + 1, Y + 1).Blocked = 15 Or MapData(X + 1, Y + 1).Blocked = 4) Then
                                    Print #handle, MapName & " ::: Posición: " & X + 1 & ", " & Y + 1 & " :::: Falto bloqueo parcial al Sur"

                                    If MapData(X + 1, Y + 1).Blocked < 0 Then
                                        MapData(X + 1, Y + 1).Blocked = 15
                                    Else
                                        MapData(X + 1, Y + 1).Blocked = 4
                                    End If
                                    MapInfo.Changed = 1
                                End If
            
                            Else ' Puerta abierta
            
                                If MapData(X, Y).Blocked <> 0 Then
                                    MapData(X, Y).Blocked = 16
                                    Print #handle, MapName & " ::: Posición: " & X & ", " & Y & " :::: BLOQUEO 16"
                                    MapInfo.Changed = 1
                                End If

                                If MapData(X - 1, Y).Blocked <> 0 Then
                                    MapData(X - 1, Y).Blocked = 16
                                    Print #handle, MapName & " ::: Posición: " & X & ", " & Y & " :::: BLOQUEO 16"
                                    MapInfo.Changed = 1
                                End If

                                If MapData(X - 1, Y + 1).Blocked <> 0 Then
                                    MapData(X - 1, Y + 1).Blocked = 16
                                    Print #handle, MapName & " ::: Posición: " & X & ", " & Y & " :::: BLOQUEO 16"
                                    MapInfo.Changed = 1
                                End If

                                If MapData(X, Y + 1).Blocked <> 0 Then
                                    MapData(X, Y + 1).Blocked = 16
                                    Print #handle, MapName & " ::: Posición: " & X & ", " & Y & " :::: BLOQUEO 16"
                                    MapInfo.Changed = 1
                                End If
                                
                            End If
                        End If
            
                    Else
            
                        If X > (1 + BordeX) And X < (MapSize.Width - BordeX) And Y > (1 + BordeY) And Y < (MapSize.Height - BordeY) Then
                            If Not IsBlock(X + 1, Y) And Not (MapData(X + 1, Y).Blocked And 1) <> 0 And ((MapData(X + 1, Y + 1).Blocked And 4) <> 0) Then
                                'Print #handle, MapName & " ::: Posición: " & X + 1 & ", " & y & " :::: FALTA BLOQUEO TOTAL"
                            End If
            
                            If ObjData(MapData(X, Y).OBJInfo.ObjIndex).Cerrada = 1 Then
                                If (MapData(X, Y).Blocked And 1) = 0 Then
                                    Print #handle, MapName & " ::: Posición: " & X & ", " & Y & " :::: FALTA BLOQUEO PARCIAL"
                                End If
            
                                If (MapData(X - 1, Y).Blocked And 1) = 0 Then
                                    Print #handle, MapName & " ::: Posición: " & X - 1 & ", " & Y & " :::: FALTA BLOQUEO PARCIAL"
                                End If
            
                                If (MapData(X - 1, Y + 1).Blocked And 4) = 0 Then
                                    Print #handle, MapName & " ::: Posición: " & X - 1 & ", " & Y + 1 & " :::: FALTA BLOQUEO PARCIAL"
                                End If
            
                                If (MapData(X, Y + 1).Blocked And 4) = 0 Then
                                    Print #handle, MapName & " ::: Posición: " & X & ", " & Y + 1 & " :::: FALTA BLOQUEO PARCIAL"
                                End If
                            Else
            
                                If (MapData(X, Y).Blocked And 1) <> 0 Then
                                    Print #handle, MapName & " ::: Posición: " & X & ", " & Y & " :::: HAY BLOQUEO PARCIAL"
                                End If
            
                                If (MapData(X - 1, Y).Blocked And 1) <> 0 Then
                                    Print #handle, MapName & " ::: Posición: " & X - 1 & ", " & Y & " :::: HAY BLOQUEO PARCIAL"
                                End If
            
                                If (MapData(X - 1, Y + 1).Blocked And 4) <> 0 Then
                                    Print #handle, MapName & " ::: Posición: " & X - 1 & ", " & Y + 1 & " :::: HAY BLOQUEO PARCIAL"
                                End If
            
                                If (MapData(X, Y + 1).Blocked And 4) <> 0 Then
                                    Print #handle, MapName & " ::: Posición: " & X & ", " & Y + 1 & " :::: HAY BLOQUEO PARCIAL"
                                End If
                                
                                If Not (MapData(X, Y + 1).Blocked = 4 Or Not MapData(X, Y + 1).Blocked = 15) Then
                                    Print #handle, MapName & " ::: Posición: " & X + 1 & ", " & Y + 1 & " :::: HAY BLOQUEO "
                                End If
                            End If
                        End If
            
                    End If
                End If
                
                '*********************************************************************************************************
                ' Falta la IA
                '*********************************************************************************************************
                If ObjData(MapData(X, Y).OBJInfo.ObjIndex).ObjType = 6 And ObjData(MapData(X, Y).OBJInfo.ObjIndex).Subtipo = 2 Then
                
                'Call Puertatipo2
                
                End If

                '************************************************************************************************************
                'PUERTA DUCTO
                '************************************************************************************************************
                If ObjData(MapData(X, Y).OBJInfo.ObjIndex).ObjType = 6 And ObjData(MapData(X, Y).OBJInfo.ObjIndex).Subtipo = 3 Then
    
                    If X > (1 + BordeX) And X < (MapSize.Width - BordeX) And Y > (1 + BordeY) And Y < (MapSize.Height - BordeY) Then
                        If chkPuertas.value = 1 Then
                            If Not IsBlock(X + 2, Y) And Not ((MapData(X + 2, Y).Blocked And 1) <> 0 And (MapData(X + 2, Y).Blocked And 4) <> 0) Then
                                Print #handle, MapName & " ::: Posición: " & X + 2 & ", " & Y & " :::: FALTA BLOQUEO TOTAL"
                                MapData(X + 2, Y).Blocked = 15
                                MapInfo.Changed = 1
                            End If
    
                            If Not IsBlock(X - 2, Y) And Not ((MapData(X - 2, Y).Blocked And 1) <> 0 And (MapData(X - 2, Y).Blocked And 4) <> 0) Then
                                Print #handle, MapName & " ::: Posición: " & X - 2 & ", " & Y & " :::: FALTA BLOQUEO TOTAL"
                                MapData(X - 2, Y).Blocked = 15
                                MapInfo.Changed = 1
                            End If
    
                            If ObjData(MapData(X, Y).OBJInfo.ObjIndex).Cerrada = 1 Then
    
                                If (MapData(X - 1, Y).Blocked And 1) = 0 Then
                                    Print #handle, MapName & " ::: Posición: " & X - 1 & ", " & Y & " :::: FALTA BLOQUEO PARCIAL"
                                    MapData(X - 1, Y).Blocked = 1
                                    MapInfo.Changed = 1
                                End If
    
                                If (MapData(X, Y).Blocked And 1) = 0 Then
                                    Print #handle, MapName & " ::: Posición: " & X & ", " & Y & " :::: FALTA BLOQUEO PARCIAL"
                                    MapData(X, Y).Blocked = 1
                                    MapInfo.Changed = 1
                                End If
    
                                If (MapData(X + 1, Y).Blocked And 1) = 0 Then
                                    Print #handle, MapName & " ::: Posición: " & X + 1 & ", " & Y & " :::: FALTA BLOQUEO PARCIAL"
                                    MapData(X + 1, Y).Blocked = 1
                                    MapInfo.Changed = 1
                                End If
    
                                If (MapData(X - 1, Y + 1).Blocked And 4) = 0 Then
                                    Print #handle, MapName & " ::: Posición: " & X - 1 & ", " & Y + 1 & " :::: FALTA BLOQUEO PARCIAL"
                                    MapData(X - 1, Y + 1).Blocked = 4
                                    MapInfo.Changed = 1
                                End If
    
                                If (MapData(X, Y + 1).Blocked And 4) = 0 Then
                                    Print #handle, MapName & " ::: Posición: " & X & ", " & Y + 1 & " :::: FALTA BLOQUEO PARCIAL"
                                    MapData(X, Y + 1).Blocked = 4
                                    MapInfo.Changed = 1
                                End If
    
                                If (MapData(X + 1, Y + 1).Blocked And 4) = 0 Then
                                    Print #handle, MapName & " ::: Posición: " & X + 1 & ", " & Y + 1 & " :::: FALTA BLOQUEO PARCIAL"
                                    MapData(X + 1, Y + 1).Blocked = 4
                                    MapInfo.Changed = 1
                                End If
                            Else

                                If Not IsBlock(X + 2, Y) And Not ((MapData(X + 2, Y).Blocked And 1) <> 0 And (MapData(X + 2, Y).Blocked And 4) <> 0) Then
                                    Print #handle, MapName & " ::: Posición: " & X + 2 & ", " & Y & " :::: FALTA BLOQUEO TOTAL"
                                End If
    
                                If Not IsBlock(X - 2, Y) And Not ((MapData(X - 2, Y).Blocked And 1) <> 0 And (MapData(X - 2, Y).Blocked And 4) <> 0) Then
                                    Print #handle, MapName & " ::: Posición: " & X - 2 & ", " & Y & " :::: FALTA BLOQUEO TOTAL"
                                End If
    
                                If (MapData(X - 1, Y).Blocked And 1) = 0 Then
                                    Print #handle, MapName & " ::: Posición: " & X - 1 & ", " & Y & " :::: FALTA BLOQUEO PARCIAL"
                                End If
    
                                If (MapData(X, Y).Blocked And 1) = 0 Then
                                    Print #handle, MapName & " ::: Posición: " & X & ", " & Y & " :::: FALTA BLOQUEO PARCIAL"

                                End If
    
                                If (MapData(X + 1, Y).Blocked And 1) = 0 Then
                                    Print #handle, MapName & " ::: Posición: " & X + 1 & ", " & Y & " :::: FALTA BLOQUEO PARCIAL"

                                End If
    
                                If (MapData(X - 1, Y + 1).Blocked And 4) = 0 Then
                                    Print #handle, MapName & " ::: Posición: " & X - 1 & ", " & Y + 1 & " :::: FALTA BLOQUEO PARCIAL"

                                End If
    
                                If (MapData(X, Y + 1).Blocked And 4) = 0 Then
                                    Print #handle, MapName & " ::: Posición: " & X & ", " & Y + 1 & " :::: FALTA BLOQUEO PARCIAL"

                                End If
    
                                If (MapData(X + 1, Y + 1).Blocked And 4) = 0 Then
                                    Print #handle, MapName & " ::: Posición: " & X + 1 & ", " & Y + 1 & " :::: FALTA BLOQUEO PARCIAL"

                                End If
                            End If
                        End If
                    End If
                End If

                '***********************************************************************************************************************
                'Puerta de una hoja
                '***********************************************************************************************************************
                If ObjData(MapData(X, Y).OBJInfo.ObjIndex).ObjType = 6 And ObjData(MapData(X, Y).OBJInfo.ObjIndex).Subtipo = 4 Then
                    If X > (1 + BordeX) And X < (MapSize.Width - BordeX) And Y > (1 + BordeY) And Y < (MapSize.Height - BordeY) Then
    
                        If Not IsBlock(X + 1, Y) And Not ((MapData(X + 1, Y).Blocked And 1) <> 0 And (MapData(X + 1, Y).Blocked And 4) <> 0) Then
                            If chkPuertas.value = 1 Then
                                MapData(X + 1, Y).Blocked = 15
                                MapInfo.Changed = 1
                                Print #handle, MapName & " ::: Posición: " & X + 1 & ", " & Y & " :::: SE COLOCA BLOQUEO TOTAL EN LA PARED DE LA PUERTA"
                            Else
                                Print #handle, MapName & " ::: Posición: " & X + 1 & ", " & Y & " :::: FALTA BLOQUEO TOTAL EN LA PARED DE LA PUERTA"
                            End If
                        End If
    
                        If Not IsBlock(X - 1, Y) And Not ((MapData(X - 1, Y).Blocked And 1) <> 0 And (MapData(X - 1, Y).Blocked And 4) <> 0) Then
                            If chkPuertas.value = 1 Then
                                MapData(X - 1, Y).Blocked = 15
                                MapInfo.Changed = 1
                                Print #handle, MapName & " ::: Posición: " & X - 1 & ", " & Y & " :::: SE COLOCA BLOQUEO TOTAL EN LA PARED DE LA PUERTA"
                            Else
                                Print #handle, MapName & " ::: Posición: " & X - 1 & ", " & Y & " :::: FALTA BLOQUEO TOTAL EN LA PARED DE LA PUERTA"
                            End If
                        End If
    
                        If ObjData(MapData(X, Y).OBJInfo.ObjIndex).Cerrada = 1 Then
                                               
                            If chkPuertas.value = 1 Then
                                If (MapData(X, Y).Blocked And 1) = 0 Then
                                    MapData(X, Y).Blocked = 1
                                    MapInfo.Changed = 1
                                    Print #handle, MapName & " ::: Posición: " & X & ", " & Y & " :::: SE COLOCA BLOQUEO EN LA PUERTA"
                                End If
    
                                If (MapData(X, Y + 1).Blocked And 4) = 0 Then
                                    MapData(X, Y + 1).Blocked = 4
                                    MapInfo.Changed = 1
                                    Print #handle, MapName & " ::: Posición: " & X & ", " & Y + 1 & " :::: SE COLOCA BLOQUEO EN LA PUERTA"
                                End If
    
                            Else

                                If (MapData(X, Y).Blocked And 1) = 0 Then
                                    Print #handle, MapName & " ::: Posición: " & X & ", " & Y & " :::: FALTA COLOCAR BLOQUEO NORTE EN LA PUERTA"
                                End If
    
                                If (MapData(X, Y + 1).Blocked And 4) = 0 Then
                                    Print #handle, MapName & " ::: Posición: " & X & ", " & Y + 1 & " :::: FALTA COLOCAR BLOQUEO SUR EN LA PUERTA"
                                End If

                            End If
                        
                        End If
                    End If
                End If
                
            End If
            
            '******************************************************************************************
            ' Árbol con con doble Bloqueo X,Y X-1,Y
            '******************************************************************************************
            If MapData(X, Y).OBJInfo.ObjIndex Then
                If (ObjData(MapData(X, Y).OBJInfo.ObjIndex).ObjType = 4 And (ObjData(MapData(X, Y).OBJInfo.ObjIndex).grhindex = 55638)) Then
                    If chkArboles.value = 1 Then
                        If MapData(X, Y).Blocked <> 15 Or _
                            MapData(X - 1, Y).Blocked <> 15 Then
                            MapData(X, Y).Blocked = 15
                            MapData(X - 1, Y).Blocked = 15
                            MapInfo.Changed = 1
                            Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: Faltaba un bloq :" & MapData(X, Y).Graphic(3).grhindex
                        End If
                    Else
                        If MapData(X, Y).Blocked <> 15 Or _
                            MapData(X - 1, Y).Blocked <> 15 Then
                            Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: Falta un bloq " & MapData(X, Y).Graphic(3).grhindex
                        End If
                    End If
                End If
            End If
            
            
            '******************************************************************************************
            ' Árbol con con doble Bloqueo X,Y X-1,Y
            '******************************************************************************************
            If MapData(X, Y).OBJInfo.ObjIndex Then
                If (ObjData(MapData(X, Y).OBJInfo.ObjIndex).ObjType = 4 And (ObjData(MapData(X, Y).OBJInfo.ObjIndex).grhindex = 463)) Then
                    If chkArboles.value = 1 Then
                        If MapData(X, Y).Blocked <> 15 Or _
                            MapData(X - 1, Y).Blocked <> 15 Or _
                            MapData(X - 2, Y).Blocked <> 15 Or _
                            MapData(X, Y - 1).Blocked <> 15 Or _
                            MapData(X - 1, Y - 1).Blocked <> 15 Or _
                            MapData(X - 2, Y - 1).Blocked <> 15 Then
                            
                            MapData(X, Y).Blocked = 15
                            MapData(X - 1, Y).Blocked = 15
                            MapData(X - 2, Y).Blocked = 15
                            MapData(X, Y - 1).Blocked = 15
                            MapData(X - 1, Y - 1).Blocked = 15
                            MapData(X - 2, Y - 1).Blocked = 15
                            MapInfo.Changed = 1
                        
                            Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: Faltaba un bloq "
                        End If
                    Else
                        If MapData(X, Y).Blocked <> 15 Or _
                            MapData(X - 1, Y).Blocked <> 15 Or _
                            MapData(X - 2, Y).Blocked <> 15 Or _
                            MapData(X, Y - 1).Blocked <> 15 Or _
                            MapData(X - 1, Y - 1).Blocked <> 15 Or _
                            MapData(X - 2, Y - 1).Blocked <> 15 Then
                            Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: Falta un bloq "
                        End If
                    
                    End If
                End If
            End If
            
            
            '******************************************************************************************
            ' Árbol con con doble grafico
            '******************************************************************************************
            If MapData(X, Y).OBJInfo.ObjIndex Then
                If (ObjData(MapData(X, Y).OBJInfo.ObjIndex).ObjType = 4 And (MapData(X, Y).Graphic(3).grhindex > 0)) Then
                    If chkArboles.value = 1 Then
                        Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: Saco Árbol con doble Grafico :" & MapData(X, Y).Graphic(3).grhindex
                        MapData(X, Y).Graphic(3).grhindex = 0
                        MapInfo.Changed = 1
                    Else
                        Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: Árbol con doble Grafico " & MapData(X, Y).Graphic(3).grhindex
                    
                    End If
                End If
            End If

            '******************************************************************************************
            ' saco bloq Especifico busqueda especial y consultas
            '******************************************************************************************
'               Dim igraf As Integer
'
'               Dim PrimerGraficoOLD As Long
'               Dim PrimerGraficoNEW As Long
'
'               PrimerGraficoOLD = 60924
'               PrimerGraficoNEW = 62309
'
'               For igraf = 0 To 1023
'
'                           If MapData(X, y).Graphic(1).grhindex = (PrimerGraficoOLD + igraf) Then
'                                   Print #handle, MapName & " ::: Posición camino : " & X & ", " & y & " ; " & MapData(X, y).Graphic(1).grhindex
'                                   Debug.Print igraf
'
'                                   MapData(X, y).Graphic(1).grhindex = (PrimerGraficoNEW + igraf)
'
'                                   MapInfo.Changed = 1
'
'                           End If
'                       Next igraf

            '********************************************************************************************************
            '***************************************************************************************************

'            If MapData(X, y).Trigger = 8 Then
''                If (MapData(X, y).Graphic(1).grhindex >= 1505 And MapData(X, y).Graphic(1).grhindex <= 1520) Then
''                        Print #handle, MapName & " ::: Posición del Trigger´s: " & X & ", " & y
''                End If
'                MapData(X, y).Trigger = 11
'                MapInfo.Changed = 1
''
'
'            End If
            '**************************************************************************************************
            'Busqueda de GRHs
            '**************************************************************************************************

'                If MapData(X, y).Graphic(1).grhindex = 6544 Then
'                        Print #handle, MapName & " ::: Grh se encuentra en : " & X & ", " & y
'                End If


            '******************************************************************************************
            'Arbol bloq total en X,Y X-1,Y X+1,Y
            '******************************************************************************************
            If MapData(X, Y).Graphic(3).grhindex = 50987 Then
                If Not IsBlock(X, Y) Or Not IsBlock(X - 1, Y) Or Not IsBlock(X + 1, Y) Then

                    If chkArboles.value = 1 Then
                        If (X - 1 > 0 And Y > 0) Then MapData(X - 1, Y).Blocked = 15
                        If (X + 1 > 0 And Y > 0) Then MapData(X + 1, Y).Blocked = 15
                        MapData(X, Y).Blocked = 15
                        MapInfo.Changed = 1
                        Print #handle, MapName & " ::: Posición de la Palmera: " & X & ", " & Y & " :::: Faltaban bloqueos al grafico " & MapData(X, Y).Graphic(3).grhindex&; " ya fue puesto."
                    Else

                        If (X - 1 > 0 And Y > 0) Then MapData(X - 1, Y).Blocked = 15
                        If (X + 1 > 0 And Y > 0) Then MapData(X + 1, Y).Blocked = 15
                        Print #handle, MapName & " ::: Posición de la Palmera: " & X & ", " & Y & " :::: Faltaban bloqueos al grafico " & MapData(X, Y).Graphic(3).grhindex&; " ya fue puesto."
                    End If

                End If
            End If

            '******************************************************************************************
            'Arbol total en X,Y X-1,Y X-3,Y X-3,Y
            '******************************************************************************************
            If MapData(X, Y).Graphic(3).grhindex = 50988 Then 'x 4 a la izq
                If Not IsBlock(X, Y) Or Not IsBlock(X - 1, Y) Or Not IsBlock(X - 2, Y) Or Not IsBlock(X - 3, Y) Then
                    If chkArboles.value = 1 Then
                        If (X - 1 > 0 And Y > 0) Then MapData(X - 1, Y).Blocked = 15
                        If (X - 2 > 0 And Y > 0) Then MapData(X - 2, Y).Blocked = 15
                        If (X - 3 > 0 And Y > 0) Then MapData(X - 3, Y).Blocked = 15
                        MapData(X, Y).Blocked = 15
                        MapInfo.Changed = 1
                        Print #handle, MapName & " ::: Posición de la Palmera: " & X - 4 & ", " & Y & " :::: Faltaban bloqueos al grafico " & MapData(X, Y).Graphic(3).grhindex & " ya fue puesto."
                    Else

                        If (X - 1 > 0 And Y > 0) Then MapData(X - 1, Y).Blocked = 15
                        If (X - 2 > 0 And Y > 0) Then MapData(X - 2, Y).Blocked = 15
                        If (X - 3 > 0 And Y > 0) Then MapData(X - 3, Y).Blocked = 15
                        Print #handle, MapName & " ::: Posición de la Palmera: " & X - 4 & ", " & Y & " :::: Faltaban bloqueos al grafico " & MapData(X, Y).Graphic(3).grhindex & " ya fue puesto."
                    End If

                End If
            End If
            
            '******************************************************************************************
            'Objetos bloq total  X,Y X-1,Y X-2,Y
            '******************************************************************************************
            If MapData(X, Y).Graphic(3).grhindex = 12754 Or MapData(X, Y).Graphic(3).grhindex = 12755 Then
                If Not IsBlock(X, Y) Or Not IsBlock(X - 1, Y) Or Not IsBlock(X - 2, Y) Then
                    If chkCarteles.value = 1 Then
                        If (X > 0 And Y > 0) Then MapData(X, Y).Blocked = 15
                        If (X - 1 > 0 And Y > 0) Then MapData(X - 1, Y).Blocked = 15
                        If (X - 2 > 0 And Y > 0) Then MapData(X - 2, Y).Blocked = 15
                        MapData(X, Y).Blocked = 15
                        MapInfo.Changed = 1
                        Print #handle, MapName & " ::: Posición de la Palmera: " & X - 4 & ", " & Y & " :::: Faltaban bloqueos al grafico " & MapData(X, Y).Graphic(3).grhindex & " ya fue puesto."
                    Else

                        If (X - 1 > 0 And Y > 0) Then MapData(X - 1, Y).Blocked = 15
                        If (X - 2 > 0 And Y > 0) Then MapData(X - 2, Y).Blocked = 15
                        If (X - 3 > 0 And Y > 0) Then MapData(X - 3, Y).Blocked = 15
                        Print #handle, MapName & " ::: Posición de la Palmera: " & X - 4 & ", " & Y & " :::: Faltaban bloqueos al grafico " & MapData(X, Y).Graphic(3).grhindex & " ya fue puesto."
                    End If

                End If
            End If

            '******************************************************************************************
            'Arbol total en X-3,Y
            '******************************************************************************************
            'palmeras en x-3
            If MapData(X, Y).Graphic(3).grhindex = 1879 Then
                If Not IsBlock(X - 3, Y) Then
                    If chkArboles.value = 1 Then
                        MapData(X - 3, Y).Blocked = 15
                        MapInfo.Changed = 1
                        Print #handle, MapName & " ::: Posición de la Palmera: " & X - 3 & ", " & Y & " :::: Falta bloqueos al grafico " & MapData(X, Y).Graphic(3).grhindex
                    Else
                        Print #handle, MapName & " ::: Posición de la Palmera: " & X - 3 & ", " & Y & " :::: Falta bloqueos al grafico " & MapData(X, Y).Graphic(3).grhindex
                    End If
                End If

            End If

            '*******************************************************************************************
            'Palmeras en X-2,Y-1  X-2,Y-1 X-1,Y-2 X-1,Y-2
            '*******************************************************************************************
            If MapData(X, Y).Graphic(3).grhindex = 12174 Then
                If Not IsBlock(X - 2, Y) Or Not IsBlock(X - 2, Y - 1) Or Not IsBlock(X - 2, Y - 2) Or Not IsBlock(X - 1, Y - 2) Or Not IsBlock(X - 1, Y - 1) Then
                    If chkArboles.value = 1 Then
                        MapData(X - 2, Y).Blocked = 15
                        MapInfo.Changed = 1

                        If (X - 2 > 0 And Y - 1 > 0) Then MapData(X - 2, Y - 1).Blocked = 15
                        If (X - 2 > 0 And Y - 2 > 0) Then MapData(X - 2, Y - 2).Blocked = 15
                        If (X - 1 > 0 And Y - 1 > 0) Then MapData(X - 1, Y - 1).Blocked = 15
                        If (X - 1 > 0 And Y - 2 > 0) Then MapData(X - 1, Y - 2).Blocked = 15
                        Print #handle, MapName & " ::: Posición de la Palmera: " & X - 1 & ", " & Y & " :::: Faltaba el bloqueos al grafico " & MapData(X, Y).Graphic(3).grhindex & " ya fue puesto."
                    Else

                        If (X - 2 > 0 And Y - 1 > 0) Then MapData(X - 2, Y - 1).Blocked = 15
                        If (X - 2 > 0 And Y - 2 > 0) Then MapData(X - 2, Y - 2).Blocked = 15
                        If (X - 1 > 0 And Y - 1 > 0) Then MapData(X - 1, Y - 1).Blocked = 15
                        If (X - 1 > 0 And Y - 2 > 0) Then MapData(X - 1, Y - 2).Blocked = 15
                        Print #handle, MapName & " ::: Posición de la Palmera: " & X - 1 & ", " & Y & " :::: Faltaba el bloqueos al grafico " & MapData(X, Y).Graphic(3).grhindex & " ya fue puesto."
                    End If
                End If
            End If

            '**********************************************************************************************
            'Palmeras en X-1,Y
            '**********************************************************************************************
            If MapData(X, Y).Graphic(3).grhindex = 433 Or MapData(X, Y).Graphic(3).grhindex = 460 Or MapData(X, Y).Graphic(3).grhindex = 461 Or MapData(X, Y).Graphic(3).grhindex = 1892 Or MapData(X, Y).Graphic(3).grhindex = 1877 Or MapData(X, Y).Graphic(3).grhindex = 1890 Or MapData(X, Y).Graphic(3).grhindex = 1891 Or MapData(X, Y).Graphic(3).grhindex = 1881 Then
                If Not IsBlock(X - 1, Y) Then
                    If chkArboles.value = 1 Then
                        MapData(X - 1, Y).Blocked = 15
                        MapInfo.Changed = 1
                        
                        Print #handle, MapName & " ::: Posición de la Palmera: " & X - 1 & ", " & Y & " :::: Faltaba el bloqueos al grafico " & MapData(X, Y).Graphic(3).grhindex & " ya fue puesto."
                    Else
                        Print #handle, MapName & " ::: Posición de la Palmera: " & X - 1 & ", " & Y & " :::: Faltaba el bloqueos al grafico " & MapData(X, Y).Graphic(3).grhindex
                    End If
                End If
            End If

            '**************************************************************************************************
            'Pinos del Polo
            '**************************************************************************************************
            If MapData(X, Y).Graphic(3).grhindex = 12166 Or MapData(X, Y).Graphic(3).grhindex = 12168 Or MapData(X, Y).Graphic(3).grhindex = 12165 Or MapData(X, Y).Graphic(3).grhindex = 12169 Then  'bloques de 4 y 3 arriba
                If Not IsBlock(X, Y) Or Not IsBlock(X - 1, Y) Or Not IsBlock(X + 1, Y) Or Not IsBlock(X, Y - 1) Or Not IsBlock(X - 1, Y - 1) Or Not IsBlock(X + 1, Y - 1) Or Not IsBlock(X + 2, Y - 1) Or Not IsBlock(X - 2, Y - 1) Then
                    If chkArboles.value = 1 Then
                        MapData(X, Y).Blocked = 15
                        MapInfo.Changed = 1

                        If X - 1 > 0 Then MapData(X - 1, Y).Blocked = 15
                        If X + 1 < 100 Then MapData(X + 1, Y).Blocked = 15
                        If Y - 1 > 0 Then MapData(X, Y - 1).Blocked = 15
                        If (X - 1 > 0 And Y - 1 > 0) Then MapData(X - 1, Y - 1).Blocked = 15
                        If (X - 2 > 0 And Y - 1 > 0) Then MapData(X - 2, Y - 1).Blocked = 15
                        If (X + 1 < 100 And Y - 1 > 0) Then MapData(X + 1, Y - 1).Blocked = 15
                        If (X + 2 < 100 And Y - 1 > 0) Then MapData(X + 2, Y - 1).Blocked = 15
                        Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: Faltaban Bloqueos al grafico " & MapData(X, Y).Graphic(3).grhindex & " ya fue puesto."
                    Else

                        If X - 1 > 0 Then MapData(X - 1, Y).Blocked = 15
                        If X + 1 < 100 Then MapData(X + 1, Y).Blocked = 15
                        If Y - 1 > 0 Then MapData(X, Y - 1).Blocked = 15
                        If (X - 1 > 0 And Y - 1 > 0) Then MapData(X - 1, Y - 1).Blocked = 15
                        If (X - 2 > 0 And Y - 1 > 0) Then MapData(X - 2, Y - 1).Blocked = 15
                        If (X + 1 < 100 And Y - 1 > 0) Then MapData(X + 1, Y - 1).Blocked = 15
                        If (X + 2 < 100 And Y - 1 > 0) Then MapData(X + 2, Y - 1).Blocked = 15
                        Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: Faltaban Bloqueos al grafico " & MapData(X, Y).Graphic(3).grhindex & " ya fue puesto."
                    End If

                End If
            End If

            '**************************************************************************************************
            'Pinos del Polo
            '**************************************************************************************************
            If MapData(X, Y).Graphic(3).grhindex = 12581 Or MapData(X, Y).Graphic(3).grhindex = 12170 Or MapData(X, Y).Graphic(3).grhindex = 12175 Then   'bloques de 3 y 3 arriba
                If Not IsBlock(X, Y) Or Not IsBlock(X - 1, Y) Or Not IsBlock(X + 1, Y) Or Not IsBlock(X, Y - 1) Or Not IsBlock(X - 1, Y - 1) Or Not IsBlock(X + 1, Y - 1) Then
                    If chkArboles.value = 1 Then
                        MapData(X, Y).Blocked = 15
                        MapInfo.Changed = 1

                        If X - 1 > 0 Then MapData(X - 1, Y).Blocked = 15
                        If X + 1 < 100 Then MapData(X + 1, Y).Blocked = 15
                        If Y - 1 > 0 Then MapData(X, Y - 1).Blocked = 15
                        If (X - 1 > 0 And Y - 1 > 0) Then MapData(X - 1, Y - 1).Blocked = 15
                        If (X + 1 < 100 And Y - 1 > 0) Then MapData(X + 1, Y - 1).Blocked = 15
                        Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: Faltaban Bloqueos al grafico " & MapData(X, Y).Graphic(3).grhindex & " ya fue puesto."
                    Else

                        If X - 1 > 0 Then MapData(X - 1, Y).Blocked = 15
                        If X + 1 < 100 Then MapData(X + 1, Y).Blocked = 15
                        If Y - 1 > 0 Then MapData(X, Y - 1).Blocked = 15
                        If (X - 1 > 0 And Y - 1 > 0) Then MapData(X - 1, Y - 1).Blocked = 15
                        If (X + 1 < 100 And Y - 1 > 0) Then MapData(X + 1, Y - 1).Blocked = 15
                        Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: Faltaban Bloqueos al grafico " & MapData(X, Y).Graphic(3).grhindex & " ya fue puesto."
                    End If

                End If
            End If

            '**************************************************************************************************
            'Palmeras desiereto
            '**************************************************************************************************
            If MapData(X, Y).Graphic(3).grhindex = 1880 Or MapData(X, Y).Graphic(3).grhindex = 1878 Or MapData(X, Y).Graphic(3).grhindex = 55635 Or MapData(X, Y).Graphic(3).grhindex = 1887 Or MapData(X, Y).Graphic(3).grhindex = 1886 Then
                If X > (1 + BordeX) And X < (MapSize.Width - BordeX) And Y > (1 + BordeY) And Y < (MapSize.Height - BordeY) Then
                    If Not IsBlock(X, Y) Then
                        If chkArboles.value = 1 Then
                            MapData(X, Y).Blocked = 15
                            MapInfo.Changed = 1
                            
                            Print #handle, MapName & " ::: Posición de la Árbol o arbusto: " & X & ", " & Y & " :::: Faltaban Bloqueos al grafico " & MapData(X, Y).Graphic(3).grhindex & " ya fue puesto."
                        Else
                            Print #handle, MapName & " ::: Posición de la Árbol o arbusto: " & X & ", " & Y & " :::: Faltaban Bloqueos al grafico " & MapData(X, Y).Graphic(3).grhindex
                        End If
                    End If
                End If
            End If

            '**************************************************************************************************
            'Palmeras desiereto
            '**************************************************************************************************
            If MapData(X, Y).Graphic(3).grhindex = 12581 Or MapData(X, Y).Graphic(3).grhindex = 32145 Or MapData(X, Y).Graphic(3).grhindex = 32160 Then 'bloques de 3 y 3 arriba
                If Not IsBlock(X, Y) Or Not IsBlock(X - 1, Y) Or Not IsBlock(X + 1, Y) Or Not IsBlock(X, Y - 1) Or Not IsBlock(X - 1, Y - 1) Or Not IsBlock(X + 1, Y - 1) Then
                    If chkArboles.value = 1 Then
                        MapData(X, Y).Blocked = 15
                        MapInfo.Changed = 1

                        If X - 1 > 0 Then MapData(X - 1, Y).Blocked = 15
                        If X + 1 < 100 Then MapData(X + 1, Y).Blocked = 15
                        If Y - 1 > 0 Then MapData(X, Y - 1).Blocked = 15
                        If (X - 1 > 0 And Y - 1 > 0) Then MapData(X - 1, Y - 1).Blocked = 15
                        If (X + 1 < 100 And Y - 1 > 0) Then MapData(X + 1, Y - 1).Blocked = 15
                        Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: Faltaban bloqueos al grafico " & MapData(X, Y).Graphic(3).grhindex & " ya fue puesto."
                    Else
                        Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: FALTA Bloqueos al grafico " & MapData(X, Y).Graphic(3).grhindex
                    End If

                End If
            End If

            If MapData(X, Y).Graphic(3).grhindex = 463 Then 'bloques de 3 y 3 arriba
                If Not IsBlock(X, Y) Or Not IsBlock(X - 1, Y) Or Not IsBlock(X - 2, Y) Or Not IsBlock(X, Y - 1) Or Not IsBlock(X - 1, Y - 1) Or Not IsBlock(X - 2, Y - 1) Then
                    If chkArboles.value = 1 Then
                        MapData(X, Y).Blocked = 15
                        MapInfo.Changed = 1

                        If X - 1 > 0 Then MapData(X - 1, Y).Blocked = 15
                        If X - 2 > 0 Then MapData(X - 2, Y).Blocked = 15
                        If Y - 1 > 0 Then MapData(X, Y - 1).Blocked = 15
                        If (X - 1 > 0 And Y - 1 > 0) Then MapData(X - 1, Y - 1).Blocked = 15
                        If (X - 2 > 0 And Y - 1 > 0) Then MapData(X - 2, Y - 1).Blocked = 15
                        Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: Faltaban Bloqueos al grafico " & MapData(X, Y).Graphic(3).grhindex & " ya fue puesto."

                    Else
                        Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: FALTA Bloqueos al grafico " & MapData(X, Y).Graphic(3).grhindex & "."

                    End If

                End If

            End If

            'Arboles x,y y x+1,y
            If MapData(X, Y).Graphic(3).grhindex = 6598 Or MapData(X, Y).Graphic(3).grhindex = 2549 Then
                If Not IsBlock(X, Y) Or MapData(X + 1, Y).Blocked = 0 Then
                    If chkArboles.value = 1 Then
                        MapData(X, Y).Blocked = 15
                        MapData(X + 1, Y).Blocked = 15
                        MapInfo.Changed = 1
                        Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: Faltaba bloqueos al grafico " & MapData(X, Y).Graphic(3).grhindex & " ya fue puesto."

                    Else
                        Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: FALTA bloqueos al grafico " & MapData(X, Y).Graphic(3).grhindex & "."

                    End If

                End If
            End If
            'Arboles x,y y x+1,y
            
            'Arboles x,y y x-1,y
            If MapData(X, Y).Graphic(3).grhindex = 1888 Then
                If Not IsBlock(X, Y) Or MapData(X - 1, Y).Blocked = 0 Then
                    If chkArboles.value = 1 Then
                        MapData(X, Y).Blocked = 15
                        MapData(X - 1, Y).Blocked = 15
                        MapInfo.Changed = 1
                        Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: Faltaba bloqueos al grafico " & MapData(X, Y).Graphic(3).grhindex & " ya fue puesto."

                    Else
                        Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: FALTA bloqueos al grafico " & MapData(X, Y).Graphic(3).grhindex & "."

                    End If

                End If
            End If
            'Arboles x,y y x-1,y
            
'            'Arboles x,y
'            If MapData(X, y).Graphic(3).grhindex = 55635 Then
'                If Not IsBlock(X, y) Then
'                    If chkArboles.Value = 1 Then
'                        MapData(X, y).Blocked = 15
'                        MapInfo.Changed = 1
'                        Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & y & " :::: Faltaba bloqueos al grafico " & MapData(X, y).Graphic(3).grhindex & " ya fue puesto."
'
'                    Else
'                        Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & y & " :::: FALTA bloqueos al grafico " & MapData(X, y).Graphic(3).grhindex & "."
'
'                    End If
'
'                End If
'            End If
'            'Arboles x,y

            'Arbol en su lugar
            '*******************************************************************************************************
            ' Objetos para bloquear en el X,Y lugar
            '*******************************************************************************************************
            If MapData(X, Y).OBJInfo.ObjIndex Then
                If ObjData(MapData(X, Y).OBJInfo.ObjIndex).ObjType = 4 Then

                    If X > (1 + BordeX) And X < (MapSize.Width - BordeX) And Y > (1 + BordeY) And Y < (MapSize.Height - BordeY) Then
                        If Not IsBlock(X, Y) And Not ((MapData(X, Y).Blocked And 1) <> 0) Then
                            If chkArboles.value = 1 Then
                                MapData(X, Y).Blocked = 15
                                MapInfo.Changed = 1
                                
                                Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: Faltaba Bloqueo " & MapData(X, Y).Graphic(3).grhindex&; " ya fue puesto."
                            Else
                                Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: FALTA Bloqueo " & MapData(X, Y).Graphic(3).grhindex
                            End If
                        End If

                    End If
                End If
            End If

            '*******************************************************************************************
            'Tiles sin bloquear
            '*******************************************************************************************

            If Not IsBlock(X, Y) Then
                If IsBlock(X - 1, Y) And IsBlock(X + 1, Y) And IsBlock(X, Y + 1) And IsBlock(X, Y - 1) Then
                    If chkBloqueosSin.value = 1 Then
                        MapData(X, Y).Blocked = 15
                        MapInfo.Changed = 1
                        
                        Print #handle, MapName & " ::: Posición: " & X & ", " & Y & " :::: Fue bloqueado el title sin acceso se Bloqueo ."
                    Else
                        Print #handle, MapName & " ::: Posición: " & X & ", " & Y & " :::: Falta Bloqueo."
                    End If

                End If
            End If

            '*************************************************************************************************
            'NPCS Sin Body
            '*************************************************************************************************
            If chkNPCsSin.value = 1 Then
                If MapData(X, Y).NpcIndex Then
                    If NpcData(MapData(X, Y).NpcIndex).Body = 0 Then
                        Call EraseChar(MapData(X, Y).CharIndex)
                        MapInfo.Changed = 1
                        Print #handle, MapName & " ::: Posición: " & X & ", " & Y & " :::: NPC BODY 0 "; MapData(X, Y).NpcIndex & " se borro."
                    End If

                Else

                    If BodyData(NpcData(MapData(X, Y).NpcIndex).Body).Walk(1).grhindex = 0 Then
                        '                        Call EraseChar(MapData(X, y).CharIndex)
                        '                        Print #handle, MapName & " ::: Posición: " & X & ", " & y & " :::: NPC BODY SIN GRH "; MapData(X, y).NPCIndex
                    End If
                End If
            Else

                If MapData(X, Y).NpcIndex Then
                    If NpcData(MapData(X, Y).NpcIndex).Body = 0 Then
                        'Borrar el npc
                        Print #handle, MapName & " ::: Posición: " & X & ", " & Y & " :::: NPC BODY 0 "; MapData(X, Y).NpcIndex & " se borro."
                    End If
                Else

                    'Borro el npc
                    If BodyData(NpcData(MapData(X, Y).NpcIndex).Body).Walk(1).grhindex = 0 Then
                        'Print #handle, MapName & " ::: Posición: " & X & ", " & y & " :::: NPC BODY SIN GRH "; MapData(X, y).NPCIndex
                    End If
                End If

            End If
            
            '*********************************************************************
            'Antorcha minas 2917 part 249  - 12980 part 183 By ReyarB
            '*********************************************************************
            
            If MapData(X, Y).Graphic(3).grhindex = 12980 Or MapData(X, Y).Graphic(4).grhindex = 12980 Then
                If chkFaroles.value = 1 Then
                    If MapData(X, Y).particle_Index <> 183 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y & " :::: Se puso la Particula = 183"
                        MapData(X, Y).particle_Index = 183
                        MapInfo.Changed = 1
                    End If

                Else

                    If MapData(X, Y).particle_Index <> 183 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y & " :::: FALTA la Particula = 183"
                    End If

                End If
            End If
            
            If MapData(X, Y).Graphic(3).grhindex = 2919 Or MapData(X, Y).Graphic(3).grhindex = 2909 Or MapData(X, Y).Graphic(3).grhindex = 2913 Or MapData(X, Y).Graphic(4).grhindex = 2917 Or MapData(X, Y).Graphic(4).grhindex = 2909 Or MapData(X, Y).Graphic(4).grhindex = 2913 Then
                If chkFaroles.value = 1 Then
                    If MapData(X, Y).particle_Index <> 249 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y & " :::: Se puso la Particula = 249"
                        MapData(X, Y).particle_Index = 249
                        MapInfo.Changed = 1
                    End If

                Else

                    If MapData(X, Y).particle_Index <> 249 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y & " :::: FALTA la Particula = 249"
                    End If

                End If
            
            End If
                        
            '*********************************************************************
            'Candelabros Iglesia 49407 183 By ReyarB
            '*********************************************************************
            
            If (MapData(X, Y).Graphic(3).grhindex = 12716) Then
                If chkFaroles.value = 1 Then
                    If MapData(X, Y - 1).particle_Index <> 239 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y - 1 & " :::: Se puso la Particula = 239"
                        MapData(X, Y - 1).particle_Index = 239
                        MapInfo.Changed = 1
                    End If

                    If MapData(X - 1, Y - 1).particle_Index <> 240 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X - 1 & ", " & Y - 1 & " :::: Se puso la Particula = 240"
                        MapData(X - 1, Y - 1).particle_Index = 240
                        MapInfo.Changed = 1
                    End If

                    If MapData(X + 1, Y - 2).particle_Index <> 241 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X + 1 & ", " & Y - 2 & " :::: Se puso la Particula = 241"
                        MapData(X + 1, Y - 2).particle_Index = 241
                        MapInfo.Changed = 1
                    End If

                    If MapData(X - 1, Y - 2).particle_Index <> 240 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X - 1 & ", " & Y - 2 & " :::: Se puso la Particula = 240"
                        MapData(X - 1, Y - 2).particle_Index = 240
                        MapInfo.Changed = 1
                    End If

                Else

                    If MapData(X, Y - 1).particle_Index <> 239 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y - 1 & " :::: Falta la Particula = 239"
                    End If

                    If MapData(X - 1, Y - 1).particle_Index <> 240 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X - 1 & ", " & Y - 1 & " :::: Falta la Particula = 240"
                    End If

                    If MapData(X + 1, Y - 2).particle_Index <> 241 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X + 1 & ", " & Y - 2 & " :::: Falta la Particula = 241"
                    End If

                    If MapData(X - 1, Y - 2).particle_Index <> 240 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X - 1 & ", " & Y - 2 & " :::: Falta la Particula = 240"
                    End If

                End If
            End If
            
            '*********************************************************************
            'Candelabros Iglesia 49407 183 By ReyarB
            '*********************************************************************
            
            If (MapData(X, Y).Graphic(3).grhindex = 49407) Or (MapData(X, Y).Graphic(4).grhindex = 49407) Then
                If chkFaroles.value = 1 Then
                    If MapData(X, Y).particle_Index <> 239 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y & " :::: Se puso la Particula = 239"
                        MapData(X, Y).particle_Index = 239
                        MapInfo.Changed = 1
                    End If

                    If MapData(X - 1, Y).particle_Index <> 240 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X - 1 & ", " & Y & " :::: Se puso la Particula = 240"
                        MapData(X - 1, Y).particle_Index = 240
                        MapInfo.Changed = 1
                    End If

                    If MapData(X + 1, Y - 1).particle_Index <> 241 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X + 1 & ", " & Y - 1 & " :::: Se puso la Particula = 241"
                        MapData(X + 1, Y - 1).particle_Index = 241
                        MapInfo.Changed = 1
                    End If

                    If MapData(X - 1, Y - 1).particle_Index <> 240 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X - 1 & ", " & Y - 1 & " :::: Se puso la Particula = 240"
                        MapData(X - 1, Y - 1).particle_Index = 240
                        MapInfo.Changed = 1
                    End If

                Else

                    If MapData(X, Y).particle_Index <> 239 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y & " :::: Falta la Particula = 239"
                    End If

                    If MapData(X - 1, Y).particle_Index <> 240 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X - 1 & ", " & Y & " :::: Falta la Particula = 240"
                    End If

                    If MapData(X + 1, Y - 1).particle_Index <> 241 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X + 1 & ", " & Y - 1 & " :::: Falta la Particula = 241"
                    End If

                    If MapData(X - 1, Y - 1).particle_Index <> 240 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X - 1 & ", " & Y - 1 & " :::: Falta la Particula = 240"
                    End If

                End If
            End If
            
            '*********************************************************************
            'Candelabros Iglesia 4242 4243 183 By ReyarB
            '*********************************************************************
            
            If (MapData(X, Y).Graphic(3).grhindex = 4243) Or (MapData(X, Y).Graphic(4).grhindex = 4243) Then
                If chkFaroles.value = 1 Then
                    If MapData(X, Y - 1).particle_Index <> 255 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y - 1 & " :::: Se puso la Particula = 255"
                        MapData(X, Y - 1).particle_Index = 255
                        MapInfo.Changed = 1
                    End If
                Else

                    If MapData(X, Y - 1).particle_Index <> 255 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y - 1 & " :::: Falta la Particula = 255"
                    End If
                End If
            End If
            
            If (MapData(X, Y).Graphic(3).grhindex = 4242) Or (MapData(X, Y).Graphic(4).grhindex = 4242) Then
                If chkFaroles.value = 1 Then
                    If MapData(X, Y - 1).particle_Index <> 256 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y - 1&; " :::: Se puso la Particula = 256"
                        MapData(X, Y - 1).particle_Index = 256
                        MapInfo.Changed = 1
                    End If
                Else

                    If MapData(X, Y - 1).particle_Index <> 256 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y - 1&; " :::: Falta la Particula = 256"
                    End If
                End If
            End If
            
            '*********************************************************************
            'Candelabros 3 velas Iglesia 49390 257 258 259 By ReyarB
            '*********************************************************************
            
            If (MapData(X, Y).Graphic(3).grhindex = 49390) Or (MapData(X, Y).Graphic(4).grhindex = 49390) Then
                If chkFaroles.value = 1 Then
                    If MapData(X, Y - 2).particle_Index <> 258 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y - 2 & " :::: Se puso la Particula = 258"
                        MapData(X, Y - 2).particle_Index = 258
                        MapInfo.Changed = 1
                    End If
                Else

                    If MapData(X, Y - 2).particle_Index <> 258 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y - 2 & " :::: Falta la Particula = 258"
                    End If
                End If
            
                If chkFaroles.value = 1 Then
                    If MapData(X, Y - 3).particle_Index <> 259 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y - 3 & " :::: Se puso la Particula = 259"
                        MapData(X, Y - 3).particle_Index = 259
                        MapInfo.Changed = 1
                    End If
                Else

                    If MapData(X, Y - 3).particle_Index <> 259 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y - 3 & " :::: Falta la Particula = 259"
                    End If
                End If
                
                If chkFaroles.value = 1 Then
                    If MapData(X, Y - 1).particle_Index <> 257 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y - 1 & " :::: Se puso la Particula = 257"
                        MapData(X, Y - 1).particle_Index = 257
                        MapInfo.Changed = 1
                    End If
                Else

                    If MapData(X, Y - 1).particle_Index <> 257 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y - 1 & " :::: Falta la Particula = 257"
                    End If
                End If
            End If
            '*********************************************************************
            'Candelabros 3 velas Iglesia 49390 257 258 259 By ReyarB
            '*********************************************************************
            
            If (MapData(X, Y).Graphic(3).grhindex = 50806) Then
                If chkFaroles.value = 1 Then
                    If MapData(X, Y - 2).particle_Index <> 260 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y - 2 & " :::: Se puso la Particula = 260"
                        MapData(X, Y - 2).particle_Index = 260
                        MapInfo.Changed = 1
                    End If
                Else

                    If MapData(X, Y - 2).particle_Index <> 260 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y - 2 & " :::: Falta la Particula = 260"
                    End If
                End If

                If chkFaroles.value = 1 Then
                    If MapData(X, Y - 3).particle_Index <> 261 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y - 3 & " :::: Se puso la Particula = 261"
                        MapData(X, Y - 3).particle_Index = 261
                        MapInfo.Changed = 1
                    End If
                Else

                    If MapData(X, Y - 3).particle_Index <> 261 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y - 3 & " :::: Falta la Particula = 261"
                    End If
                End If

                If chkFaroles.value = 1 Then
                    If MapData(X, Y - 1).particle_Index <> 262 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y - 1 & " :::: Se puso la Particula = 262"
                        MapData(X, Y - 1).particle_Index = 262
                        MapInfo.Changed = 1
                    End If
                Else

                    If MapData(X, Y - 1).particle_Index <> 262 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y - 1 & " :::: Falta la Particula = 262"
                    End If
                End If
            End If
            
            '*********************************************************************
            'Candelabros 3 velas Iglesia 50808 263 264 265 By ReyarB
            '*********************************************************************
            
            If (MapData(X, Y).Graphic(3).grhindex = 50808) Then
                If chkFaroles.value = 1 Then
                    If MapData(X, Y - 2).particle_Index <> 263 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y - 2 & " :::: Se puso la Particula = 263"
                        MapData(X, Y - 2).particle_Index = 263
                        MapInfo.Changed = 1
                    End If
                Else

                    If MapData(X, Y - 2).particle_Index <> 263 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y - 2 & " :::: Falta la Particula = 263"
                    End If
                End If

                If chkFaroles.value = 1 Then
                    If MapData(X, Y - 1).particle_Index <> 264 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y - 1 & " :::: Se puso la Particula = 264"
                        MapData(X, Y - 1).particle_Index = 264
                        MapInfo.Changed = 1
                    End If
                Else

                    If MapData(X, Y - 1).particle_Index <> 264 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y - 1 & " :::: Falta la Particula = 264"
                    End If
                End If

                If chkFaroles.value = 1 Then
                    If MapData(X, Y - 3).particle_Index <> 265 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y - 3 & " :::: Se puso la Particula = 265"
                        MapData(X, Y - 3).particle_Index = 265
                        MapInfo.Changed = 1
                    End If
                Else

                    If MapData(X, Y - 3).particle_Index <> 265 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y - 3 & " :::: Falta la Particula = 265"
                    End If
                End If
            End If
            
            '*********************************************************************
            'Candelabros 3 velas Iglesia 49404 49405  By ReyarB
            '*********************************************************************
            
            If (MapData(X, Y).Graphic(3).grhindex = 49405) Then
                If chkFaroles.value = 1 Then
                    If MapData(X, Y).particle_Index <> 266 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y & " :::: Se puso la Particula = 266"
                        MapData(X, Y).particle_Index = 266
                        MapInfo.Changed = 1
                    End If
                Else

                    If MapData(X, Y).particle_Index <> 266 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y & " :::: Falta la Particula = 266"
                    End If
                End If

            End If
            
            If (MapData(X, Y).Graphic(3).grhindex = 49404) Then
                If chkFaroles.value = 1 Then
                    If MapData(X, Y).particle_Index <> 268 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y & " :::: Se puso la Particula = 268"
                        MapData(X, Y).particle_Index = 268
                        MapInfo.Changed = 1
                    End If
                Else

                    If MapData(X, Y).particle_Index <> 268 Then
                        Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y & " :::: Falta la Particula = 268"
                    End If
                End If

            End If
            
            '************************************************************************************
            'Prueba de arboles que aparecen de golpe SUBO UN LUGAR
            '************************************************************************************
            
            If (MapData(X, Y).Graphic(3).grhindex = 12581) And Y < 33 Then
                If chkArboles.value = 1 Then
                
                    Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: Bajar el Árbol para que no desaparesca es muy ALTO. Grafico = " & MapData(X, Y).Graphic(3).grhindex
                Else

                    Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: Bajar el Árbol para que no desaparesca es muy ALTO. Grafico =  " & MapData(X, Y).Graphic(3).grhindex

                End If

            End If
            
            ' subo 2 lugares el arbol o lo borro si no puedo
            If MapData(X, Y).OBJInfo.ObjIndex Then
                If ObjData(MapData(X, Y).OBJInfo.ObjIndex).ObjType = 4 Then
                    
                    If X > (1 + BordeX) And X < (MapSize.Width - BordeX) And Y = 21 Then
                        If chkGraficosDe.value = 1 Then
                            If (MapData(X, Y - 1).OBJInfo.ObjIndex = 0 And MapData(X, Y - 2).OBJInfo.ObjIndex = 0 And Not MapData(X, Y - 1).Blocked > 0 And Not MapData(X, Y - 2).Blocked > 0 And Not MapData(X - 1, Y - 2).Blocked > 0 And Not MapData(X + 1, Y - 2).Blocked > 0) Then
                                MapData(X, Y - 2).OBJInfo.ObjIndex = MapData(X, Y).OBJInfo.ObjIndex
                                MapData(X, Y - 2).Blocked = 15
                                MapData(X, Y).Blocked = 0
                                MapData(X, Y).OBJInfo.ObjIndex = 0
                                Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: Subo el Árbol para que no desaparesca " & MapData(X, Y).OBJInfo.ObjIndex
                                MapInfo.Changed = 1
                            Else
                                MapData(X, Y).Blocked = 0
                                MapData(X, Y).OBJInfo.ObjIndex = 0
                                MapInfo.Changed = 1
                                Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: Borro el Árbol para que no desaparesca " & MapData(X, Y).OBJInfo.ObjIndex

                            End If
                        Else
                            Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: Mover o borrar este árbol desaparece al cambiar de mapa, árbol = " & MapData(X, Y).OBJInfo.ObjIndex
                        End If
                            
                    End If
                End If
            End If
            
            If MapData(X, Y).OBJInfo.ObjIndex Then
                If ObjData(MapData(X, Y).OBJInfo.ObjIndex).ObjType = 4 Then

                    If X > (1 + BordeX) And X < (MapSize.Width - BordeX) And Y = 22 Then
                        If chkGraficosDe.value = 1 Then
                            If (MapData(X, Y - 1).OBJInfo.ObjIndex = 0 And MapData(X, Y - 2).OBJInfo.ObjIndex = 0 And Not MapData(X, Y - 1).Blocked > 0 And Not MapData(X, Y - 2).Blocked > 0 And Not MapData(X + 1, Y - 1).Blocked > 0 And Not MapData(X - 1, Y - 2).Blocked > 0) Then
                                MapData(X, Y - 2).OBJInfo.ObjIndex = MapData(X, Y).OBJInfo.ObjIndex
                                MapData(X, Y - 2).Blocked = 15
                                MapData(X, Y).Blocked = 0
                                MapData(X, Y).OBJInfo.ObjIndex = 0
                                Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: subo el árbol para que no desaparesca " & MapData(X, Y).OBJInfo.ObjIndex
                                MapInfo.Changed = 1
                            Else
                                MapData(X, Y).Blocked = 0
                                MapData(X, Y).OBJInfo.ObjIndex = 0
                                MapInfo.Changed = 1
                                Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: Borro el árbol para que no desaparesca " & MapData(X, Y).OBJInfo.ObjIndex

                            End If
                        Else
                            Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: Mover o borrar este árbol desaparece al cambiar de mapa, árbol = " & MapData(X, Y).OBJInfo.ObjIndex
                        End If

                    End If
                End If
            End If
            
            If MapData(X, Y).OBJInfo.ObjIndex Then
                If ObjData(MapData(X, Y).OBJInfo.ObjIndex).ObjType = 4 Then
                    
                    If X > (1 + BordeX) And X < (MapSize.Width - BordeX) And (Y >= 23 And Y <= 26) Then
                        If chkGraficosDe.value = 1 Then
                            If (MapData(X, Y + 1).OBJInfo.ObjIndex = 0 And MapData(X, Y + 2).OBJInfo.ObjIndex = 0 And MapData(X, Y + 3).OBJInfo.ObjIndex = 0 And Not MapData(X, Y + 1).Blocked > 0 And Not MapData(X, Y + 2).Blocked > 0 And Not MapData(X, Y + 3).Blocked > 0 And Not MapData(X - 1, Y + 3).Blocked > 0 And Not MapData(X + 1, Y + 3).Blocked > 0) Then
                                MapData(X, Y + 3).OBJInfo.ObjIndex = MapData(X, Y).OBJInfo.ObjIndex
                                MapData(X, Y + 3).Blocked = 15
                                MapData(X, Y).Blocked = 0
                                MapData(X, Y).OBJInfo.ObjIndex = 0
                                Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: Bajo el árbol para que no desaparesca " & MapData(X, Y).OBJInfo.ObjIndex
                                MapInfo.Changed = 1
                            Else
                                MapData(X, Y).Blocked = 0
                                MapData(X, Y).OBJInfo.ObjIndex = 0
                                MapInfo.Changed = 1
                                Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: Borro el árbol para que no desaparesca " & MapData(X, Y).OBJInfo.ObjIndex

                            End If
                        Else
                            Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: Mover o borrar este árbol desaparece al cambiar de mapa, árbol = " & MapData(X, Y).OBJInfo.ObjIndex
                        End If
                            
                    End If
                End If
            End If
            
            '******************************************************************************************
            ' Arbol en Agua
            '******************************************************************************************
            If MapData(X, Y).OBJInfo.ObjIndex Then
                If (ObjData(MapData(X, Y).OBJInfo.ObjIndex).ObjType = 4 And (MapData(X, Y).Graphic(1).grhindex >= 1505 And MapData(X, Y).Graphic(1).grhindex <= 1520) And MapData(X, Y).Graphic(2).grhindex = 0) Then
                    If chkArboles.value = 1 Then
                        Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: Árbol en el agua lo saco :" & MapData(X, Y).OBJInfo.ObjIndex
                        MapData(X, Y).Blocked = 0
                        MapData(X, Y).OBJInfo.ObjIndex = 0
                        MapInfo.Changed = 1
                    Else
                        Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: Árbol en el agua " & MapData(X, Y).OBJInfo.ObjIndex
                    
                    End If
                End If
            End If
            
            '***********************************************************************************************************************************
            'Arbol en su lugar solo grafico para mover
            '***********************************************************************************************************************************
            If (MapData(X, Y).Graphic(3).grhindex >= 15108 And MapData(X, Y).Graphic(3).grhindex <= 15110) Or MapData(X, Y).Graphic(3).grhindex = 12731 Or MapData(X, Y).Graphic(3).grhindex = 304 Or MapData(X, Y).Graphic(3).grhindex = 305 Or MapData(X, Y).Graphic(3).grhindex = 641 Or MapData(X, Y).Graphic(3).grhindex = 644 Or _
               MapData(X, Y).Graphic(3).grhindex = 647 Or MapData(X, Y).Graphic(3).grhindex = 735 Or MapData(X, Y).Graphic(3).grhindex = 1121 Or MapData(X, Y).Graphic(3).grhindex = 1126 Or MapData(X, Y).Graphic(3).grhindex = 2931 Or (MapData(X, Y).Graphic(3).grhindex >= 1161 And MapData(X, Y).Graphic(3).grhindex <= 1168) Or (MapData(X, Y).Graphic(3).grhindex >= 7000 And MapData(X, Y).Graphic(3).grhindex <= 7002) Or _
               (MapData(X, Y).Graphic(3).grhindex >= 7222 And MapData(X, Y).Graphic(3).grhindex <= 7226) Or MapData(X, Y).Graphic(3).grhindex = 12309 Or MapData(X, Y).Graphic(3).grhindex = 12310 Or (MapData(X, Y).Graphic(3).grhindex >= 12582 And MapData(X, Y).Graphic(3).grhindex <= 12586) Or _
               (MapData(X, Y).Graphic(3).grhindex >= 12164 And MapData(X, Y).Graphic(3).grhindex <= 12173) Or (MapData(X, Y).Graphic(3).grhindex >= 12175 And MapData(X, Y).Graphic(3).grhindex <= 12179) Or (MapData(X, Y).Graphic(3).grhindex >= 14950 And MapData(X, Y).Graphic(3).grhindex <= 14959) Or (MapData(X, Y).Graphic(3).grhindex >= 14961 And MapData(X, Y).Graphic(3).grhindex <= 14980) Or _
               (MapData(X, Y).Graphic(3).grhindex >= 14982 And MapData(X, Y).Graphic(3).grhindex <= 14988) Or MapData(X, Y).Graphic(3).grhindex = 16833 Or MapData(X, Y).Graphic(3).grhindex = 16834 Or _
               (MapData(X, Y).Graphic(3).grhindex >= 26075 And MapData(X, Y).Graphic(3).grhindex <= 26081) Or MapData(X, Y).Graphic(3).grhindex = 26192 Or (MapData(X, Y).Graphic(3).grhindex >= 32142 And MapData(X, Y).Graphic(3).grhindex <= 32154) Or (MapData(X, Y).Graphic(3).grhindex >= 32159 And MapData(X, Y).Graphic(3).grhindex <= 32162) Or _
               (MapData(X, Y).Graphic(3).grhindex >= 32343 And MapData(X, Y).Graphic(3).grhindex <= 32352) Or (MapData(X, Y).Graphic(3).grhindex >= 55626 And MapData(X, Y).Graphic(3).grhindex <= 55634) Or (MapData(X, Y).Graphic(3).grhindex >= 55636 And MapData(X, Y).Graphic(3).grhindex <= 55640) Or MapData(X, Y).Graphic(3).grhindex = 55642 Or _
               (MapData(X, Y).Graphic(3).grhindex >= 50985 And MapData(X, Y).Graphic(3).grhindex <= 50991) Or (MapData(X, Y).Graphic(3).grhindex >= 2547 And MapData(X, Y).Graphic(3).grhindex <= 2549) Or (MapData(X, Y).Graphic(3).grhindex >= 6597 And MapData(X, Y).Graphic(3).grhindex <= 6598) Or MapData(X, Y).Graphic(3).grhindex = 50968 Then

                If X > (1 + BordeX) And X < (MapSize.Width - BordeX) And Y > (1 + BordeY) And Y < (MapSize.Height - BordeY) Then
                    If Not IsBlock(X, Y) Then
                        If chkArboles.value = 1 Then
                            MapData(X, Y).Blocked = 15
                            MapInfo.Changed = 1
                            'espejar
                            Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: Faltaba Bloqueo " & MapData(X, Y).Graphic(3).grhindex&; " ya fue puesto."
                        Else
                            Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: FALTA Bloqueo " & MapData(X, Y).Graphic(3).grhindex
                        End If

                    End If

                End If
                
                '**********************************************************
                'Árbol en zona de desaparecer
                '**********************************************************
                MapSup = 0
                Call VerMapaArriba
                
                If X > (1 + BordeX) And X < (MapSize.Width - BordeX) And (Y = 21 Or Y = 22) Then
                
                
                
                    If (chkGraficosDe.value = 1 And MapSup = 1) Then
                        If (MapData(X, Y - 2).Graphic(3).grhindex = 0 And Not MapData(X, Y - 1).Blocked > 0 And Not MapData(X, Y - 2).Blocked > 0 And Not MapData(X - 1, Y - 2).Blocked > 0 And Not MapData(X + 1, Y - 2).Blocked > 0) And Not (MapData(X, Y).Graphic(1).grhindex >= 1505 And MapData(X, Y).Graphic(1).grhindex <= 1520) Then
                            MapData(X, Y - 2).Graphic(3).grhindex = MapData(X, Y).Graphic(3).grhindex
                            MapData(X, Y - 2).Blocked = 15
                            MapData(X, Y).Blocked = 0
                            MapData(X, Y).Graphic(3).grhindex = 0
                            Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: Subo el Árbol para que no desaparesca " & MapData(X, Y).Graphic(3).grhindex
                            MapInfo.Changed = 1
                        Else
                            MapData(X, Y).Blocked = 0
                            MapData(X, Y).Graphic(3).grhindex = 0
                            MapInfo.Changed = 1
                            Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: Borro el Árbol para que no desaparesca " & MapData(X, Y).Graphic(3).grhindex

                        End If
                    Else
                        If MapSup = 1 Then
                            Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: Mover o borrar este árbol desaparece al cambiar de mapa, árbol = " & MapData(X, Y).Graphic(3).grhindex
                        End If
                    End If
                End If
                
                MapSup = 0
                Call VerMapaArriba
                
                If X > (1 + BordeX) And X < (MapSize.Width - BordeX) And (Y >= 23 And Y <= 26) Then
                
                    If (chkGraficosDe.value = 1 And MapSup = 1) Then
                        If (MapData(X, Y + 1).Graphic(3).grhindex = 0 And MapData(X, Y + 2).Graphic(3).grhindex = 0 And MapData(X, Y + 3).Graphic(3).grhindex = 0 And Not MapData(X, Y + 1).Blocked > 0 And Not MapData(X, Y + 2).Blocked > 0 And Not MapData(X, Y + 3).Blocked > 0 And Not MapData(X - 1, Y + 3).Blocked > 0 And Not MapData(X + 1, Y + 3).Blocked > 0 And Not (MapData(X, Y).Graphic(1).grhindex >= 1505 And MapData(X, Y).Graphic(1).grhindex <= 1520)) Then
                            MapData(X, Y + 3).Graphic(3).grhindex = MapData(X, Y).Graphic(3).grhindex
                            MapData(X, Y + 3).Blocked = 15
                            MapData(X, Y).Blocked = 0
                            MapData(X, Y).Graphic(3).grhindex = 0
                            Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: Bajo el árbol para que no desaparesca " & MapData(X, Y).Graphic(3).grhindex
                            MapInfo.Changed = 1
                        Else
                            MapData(X, Y).Blocked = 0
                            MapData(X, Y).Graphic(3).grhindex = 0
                            MapInfo.Changed = 1
                            Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: Borro el árbol para que no desaparesca " & MapData(X, Y).Graphic(3).grhindex

                        End If
                    Else
                        If MapSup = 1 Then
                            Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & Y & " :::: Mover o borrar este árbol desaparece al cambiar de mapa, árbol = " & MapData(X, Y).Graphic(3).grhindex
                        End If
                    End If
                            
                End If
                
            End If

            '************************************************************************************
            'Flores y plantas bloq parcial en X,Y y X,Y+1
            '************************************************************************************
            If (MapData(X, Y).Graphic(3).grhindex >= 5298 And MapData(X, Y).Graphic(3).grhindex <= 5304) Or (MapData(X, Y).Graphic(3).grhindex >= 14399 And MapData(X, Y).Graphic(3).grhindex <= 14401) Or MapData(X, Y).Graphic(3).grhindex = 14407 Or MapData(X, Y).Graphic(3).grhindex = 14430 Or MapData(X, Y).Graphic(3).grhindex = 14431 Or MapData(X, Y).Graphic(3).grhindex = 4459 Or MapData(X, Y).Graphic(3).grhindex = 4705 Or MapData(X, Y).Graphic(3).grhindex = 5204 Or MapData(X, Y).Graphic(3).grhindex = 12394 Or MapData(X, Y).Graphic(3).grhindex = 4704 Or MapData(X, Y).Graphic(3).grhindex = 12526 Or MapData(X, Y).Graphic(3).grhindex = 4703 Or MapData(X, Y).Graphic(3).grhindex = 4693 Then

                    If X > (1 + BordeX) And X < (MapSize.Width - BordeX) And Y > (1 + BordeY) And Y < (MapSize.Height - BordeY) Then

                        Call FloresPlantas(X, Y)

                    End If
                End If

                iz = 0
                de = 0
                ar = 0
                ab = 0

            
        Next
        
    Next
     

    
End Sub '

Private Sub ListadoNpcs_Timer()
    Dim X As Integer, Y As Integer, BordeX As Integer, BordeY As Integer

    For Y = 1 + BordeY To MapSize.Height - BordeY
        For X = 1 + BordeX To MapSize.Width - BordeX
    
            '*************************************************************************************************
            'Listado NPCs
            '*************************************************************************************************
            
'                If MapData(X, y).NPCIndex >= 700 And MapData(X, y).NPCIndex <= 720 Then
'                        Print #handle, MapName & ";Posición;" & X & ", " & y & ";NPC Nº; " & MapData(X, y).NPCIndex & ";Nombre;" & NpcData(MapData(X, y).NPCIndex).Name
'                End If
                
                If MapData(X, Y).NpcIndex > 0 Then
                        Print #handle, MapName & ";Posición;" & X & ", " & Y & ";NPC Nº; " & MapData(X, Y).NpcIndex & ";Nombre;" & NpcData(MapData(X, Y).NpcIndex).name
                End If
            'End If
    
        Next
    
    Next
    
 Print #handle, MapName & " ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"



    Call modMapIO.GuardarMapa(PATH_Save & MapName)
    
End Sub

Private Sub FarolBander(ByVal X As Integer, ByVal Y As Integer, ByVal grafico As Long, ByVal part As Long)
                '*********************************************************************
                'Farol Bander 2640 by ReyarB
                '*********************************************************************
                If MapData(X, Y).Graphic(3).grhindex = 2460 Then
                    If chkFaroles.value = 1 Then
                        If MapData(X, Y - 3).particle_Index <> 269 Then
                            Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y - 3 & " :::: Se puso la Particula = 269"
                            MapData(X, Y - 3).particle_Index = 269
                            MapInfo.Changed = 1
                        End If
    
                        If MapData(X - 1, Y - 3).particle_Index <> 270 Then
                            Print #handle, MapName & " ::: Posición del Particula: " & X - 1 & ", " & Y - 3 & " :::: Se puso la Particula = 270"
                            MapData(X - 1, Y - 3).particle_Index = 270
                            MapInfo.Changed = 1
                        End If
                    Else

                        If MapData(X, Y - 3).particle_Index <> 269 Then
                            Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y - 3 & " :::: Falta la Particula = 269"
                        End If
    
                        If MapData(X - 1, Y - 3).particle_Index <> 270 Then
                            Print #handle, MapName & " ::: Posición del Particula: " & X - 1 & ", " & Y - 3 & " :::: Falta la Particula = 270"
                        End If
'                        End If
                    End If
                End If
End Sub

Private Sub MagiaGas(ByVal X As Integer, ByVal Y As Integer, ByVal grafico As Long, ByVal part As Long)
                '*********************************************************************
                'Gas de Pociones 2640 by ReyarB
                '*********************************************************************
                If MapData(X, Y).Graphic(3).grhindex = 2093 Or MapData(X, Y).Graphic(3).grhindex = 28223 Then
                    If chkFaroles.value = 1 Then
                        If MapData(X, Y - 1).particle_Index <> 275 Then
                            Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y - 1 & " :::: Se puso la Particula = 275"
                            MapData(X, Y - 1).particle_Index = 275
                            MapInfo.Changed = 1
                        End If
                     If MapData(X, Y).particle_Index <> 276 Then
                            Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y - 1 & " :::: Se puso la Particula = 276"
                            MapData(X, Y).particle_Index = 276
                            MapInfo.Changed = 1
                     End If
    
                    Else
    
                        If MapData(X, Y - 1).particle_Index <> 275 Then
                            Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y - 1 & " :::: Falta la Particula = 275"
                        End If
'                        End If
                    End If
                End If
End Sub

Private Sub HogarLeña(ByVal X As Integer, ByVal Y As Integer, ByVal grafico As Long, ByVal part As Long)
                '*********************************************************************
                'Hogar a leña de casas grh 2407,2408 by ReyarB
                '*********************************************************************
                If ((MapData(X, Y).Graphic(3).grhindex = 2407 Or MapData(X, Y).Graphic(3).grhindex = 2408) And MapData(X, Y).Trigger < 50) Then
                    If chkCasas.value = 1 Then
                        If MapData(X, Y).particle_Index <> 250 Then
                            Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y & " :::: Se puso la Particula = 250"
                            MapData(X, Y).particle_Index = 250
                            MapInfo.Changed = 1
                        End If
    
                        If MapData(X, Y - 4).particle_Index <> 180 Then
                            Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y - 4 & " :::: Se puso la Particula = 180"
                            MapData(X, Y - 4).particle_Index = 180
                            MapInfo.Changed = 1
                        End If
                    Else

                        If MapData(X, Y).particle_Index <> 250 Then
                            Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y & " :::: FALTA la Particula = 250"
                        End If
    
                        If MapData(X, Y - 4).particle_Index <> 180 Then
                            Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y - 4 & " :::: FALTA la Particula = 180"
                            MapData(X, Y - 4).particle_Index = 180
                        End If
                    End If
                End If
End Sub

Private Sub FixLuces(ByVal X As Integer, ByVal Y As Integer, Rango, color, grafico, Particula)

    If chkFaroles.value = 1 Then

            If (MapData(X, Y).Graphic(3).grhindex = 12747 And MapData(X, Y - 3).particle_Index <> 271) Then
                If MapData(X, Y - 3).particle_Index <> 271 Then
                    MapData(X, Y - 3).particle_Index = 271
                    MapInfo.Changed = 1
                    Print #handle, MapName & " ::: Se coloca Perticula en: " & X & ", " & Y & " :::: : 271"
                End If
            End If
            
            If (MapData(X, Y).Graphic(3).grhindex = 12748 And MapData(X, Y - 1).particle_Index <> 272) Then
                If MapData(X, Y - 1).particle_Index <> 272 Then
                    MapData(X, Y - 1).particle_Index = 272
                    MapInfo.Changed = 1
                    Print #handle, MapName & " ::: Se coloca Perticula en: " & X & ", " & Y & " :::: : 272"
                End If
            End If

            If (MapData(X, Y).Graphic(3).grhindex = 12749 And MapData(X, Y - 2).particle_Index <> 273) Then
                If MapData(X, Y - 2).particle_Index <> 273 Then
                    MapData(X, Y - 2).particle_Index = 273
                    MapInfo.Changed = 1
                    Print #handle, MapName & " ::: Se coloca Perticula en: " & X - 1 & ", " & Y - 2 & " :::: : 273"
                End If
            End If

            If (MapData(X, Y).Graphic(3).grhindex = 12750 And MapData(X, Y - 2).particle_Index <> 274) Then
                If MapData(X, Y - 2).particle_Index <> 274 Then
                    MapData(X, Y - 2).particle_Index = 274
                    MapInfo.Changed = 1
                    Print #handle, MapName & " ::: Se coloca Perticula en: " & X - 1 & ", " & Y - 2 & " :::: : 274"
                End If
            End If
    Else
    
    
            If (MapData(X, Y).Graphic(3).grhindex = 12747 And MapData(X, Y - 3).particle_Index <> 271) Then
                If MapData(X, Y - 3).particle_Index <> 271 Then
                    Print #handle, MapName & " ::: Falta Perticula en: " & X & ", " & Y & " :::: : 271"
                End If
            End If
            
            If (MapData(X, Y).Graphic(3).grhindex = 12748 And MapData(X, Y - 1).particle_Index <> 272) Then
                If MapData(X, Y - 1).particle_Index <> 272 Then
                    Print #handle, MapName & " ::: Falta Perticula en: " & X & ", " & Y & " :::: : 272"
                End If
            End If

            If (MapData(X, Y).Graphic(3).grhindex = 12749 And MapData(X, Y - 2).particle_Index <> 273) Then
                If MapData(X, Y - 2).particle_Index <> 273 Then
                    Print #handle, MapName & " ::: Falta Perticula en: " & X - 1 & ", " & Y - 2 & " :::: : 273"
                End If
            End If

            If (MapData(X, Y).Graphic(3).grhindex = 12750 And MapData(X, Y - 2).particle_Index <> 274) Then
                If MapData(X, Y - 2).particle_Index <> 274 Then
                    Print #handle, MapName & " ::: Falta Perticula en: " & X - 1 & ", " & Y - 2 & " :::: : 274"
                End If
            End If
    End If
    
    If chkLuzFalor.value = 1 Then

            If MapData(X, Y).Graphic(3).grhindex = 12747 Then
                If MapData(X, Y - 1).luz.Rango = 0 Then
                    MapData(X, Y - 1).luz.Rango = 103
                    MapData(X, Y - 1).luz.color = 16777215
                    MapInfo.Changed = 1
                    Print #handle, MapName & " ::: Se coloca Luz en: " & X & ", " & Y & " :::: Rango de: 103"
                End If
            End If
            
            
            If MapData(X, Y).Graphic(3).grhindex = 12748 Then
                If MapData(X, Y + 1).luz.Rango = 0 Then
                    MapData(X, Y + 1).luz.Rango = 103
                    MapData(X, Y + 1).luz.color = 16777215
                    MapInfo.Changed = 1
                    Print #handle, MapName & " ::: Se coloca Luz en: " & X & ", " & Y & " :::: Rango de: 103"
                End If
            End If
            
            If MapData(X, Y).Graphic(3).grhindex = 12749 Then
                If MapData(X, Y).luz.Rango = 0 Then
                    MapData(X, Y).luz.Rango = 103
                    MapData(X, Y).luz.color = 16777215
                    MapInfo.Changed = 1
                    Print #handle, MapName & " ::: Se coloca Luz en: " & X & ", " & Y & " :::: Rango de: 103"
                End If
            End If
            
            If MapData(X, Y).Graphic(3).grhindex = 12750 Then
                If MapData(X, Y).luz.Rango = 0 Then
                    MapData(X, Y).luz.Rango = 103
                    MapData(X, Y).luz.color = 16777215
                    MapInfo.Changed = 1
                    Print #handle, MapName & " ::: Se coloca Luz en: " & X & ", " & Y & " :::: Rango de: 103"
                End If
            End If
            
            If MapData(X, Y).Graphic(3).grhindex = 5626 Then
                If MapData(X, Y - 1).luz.Rango = 0 Then
                    MapData(X, Y - 1).luz.Rango = 103
                    MapData(X, Y - 1).luz.color = 16777215
                    MapInfo.Changed = 1
                    Print #handle, MapName & " ::: Se coloca Luz en: " & X & ", " & Y & " :::: Rango de: 103"
                End If
            End If

            If MapData(X, Y).Graphic(3).grhindex = 5625 Or MapData(X, Y).Graphic(3).grhindex = 2460 Then
                If MapData(X + 1, Y).luz.Rango = 0 Then
                    MapData(X + 1, Y).luz.Rango = 103
                    MapData(X + 1, Y).luz.color = 16777215
                    MapInfo.Changed = 1
                    Print #handle, MapName & " ::: Se coloca Luz en: " & X & ", " & Y & " :::: Rango de: 103"
                End If
            End If

            If MapData(X, Y).Graphic(3).grhindex = 5624 Then
                If MapData(X, Y).luz.Rango = 0 Then
                    MapData(X, Y).luz.Rango = 103
                    MapData(X, Y).luz.color = 16777215
                    MapInfo.Changed = 1
                    Print #handle, MapName & " ::: Se coloca Luz en: " & X & ", " & Y & " :::: Rango de:103 "
                End If
            End If

            If MapData(X, Y).Graphic(3).grhindex = 5627 Then
                If MapData(X, Y + 1).luz.Rango = 0 Then
                    MapData(X, Y + 1).luz.Rango = 103
                    MapData(X, Y + 1).luz.color = 16777215
                    MapInfo.Changed = 1
                    Print #handle, MapName & " ::: Se coloca Luz en: " & X & ", " & Y & " :::: Rango de: 103"
                End If

            End If
 
    Else
            If MapData(X, Y).Graphic(3).grhindex = 12747 Then
                If MapData(X, Y - 1).luz.Rango = 0 Then
                    Print #handle, MapName & " ::: Falta Luz en: " & X & ", " & Y - 1 & " :::: Rango de: 103"
                End If
            End If
            
            
            If MapData(X, Y).Graphic(3).grhindex = 12748 Then
                If MapData(X, Y + 1).luz.Rango = 0 Then
                    Print #handle, MapName & " ::: Falta Luz en: " & X & ", " & Y + 1 & " :::: Rango de: 103"
                End If
            End If
            
            If MapData(X, Y).Graphic(3).grhindex = 12749 Then
                If MapData(X, Y).luz.Rango = 0 Then
                    Print #handle, MapName & " ::: Falta Luz en: " & X & ", " & Y & " :::: Rango de: 103"
                End If
            End If
            
            If MapData(X, Y).Graphic(3).grhindex = 12750 Then
                If MapData(X, Y).luz.Rango = 0 Then
                    Print #handle, MapName & " ::: Faltaa Luz en: " & X & ", " & Y & " :::: Rango de: 103"
                End If
            End If
            
            If MapData(X, Y).Graphic(3).grhindex = 5626 Then
                If MapData(X, Y - 1).luz.Rango = 0 Then
                    Print #handle, MapName & " ::: Falta Luz en: " & X & ", " & Y & " :::: Rango de: 103"
                End If
            End If

            If MapData(X, Y).Graphic(3).grhindex = 5625 Or MapData(X, Y).Graphic(3).grhindex = 2460 Then
                If MapData(X + 1, Y).luz.Rango = 0 Then
                    Print #handle, MapName & " ::: Falta Luz en: " & X & ", " & Y & " :::: Rango de: 103"
                End If
            End If

            If MapData(X, Y).Graphic(3).grhindex = 5624 Then
                If MapData(X, Y).luz.Rango = 0 Then
                    Print #handle, MapName & " ::: Falta Luz en: " & X & ", " & Y & " :::: Rango de:103 "
                End If
            End If

            If MapData(X, Y).Graphic(3).grhindex = 5627 Then
                If MapData(X, Y + 1).luz.Rango = 0 Then
                    Print #handle, MapName & " ::: Falta Luz en: " & X & ", " & Y & " :::: Rango de: 103"
                End If

            End If
         
    End If

'                    MapData(X, Y).luz.Rango = 0
'                    MapData(X, Y).luz.color = 0
            

End Sub

Private Sub VerMapaArriba()

    Dim xmap As Integer
    Dim ymap As Integer
    
    ymap = 10
    
    For xmap = (14) To (87)

        If MapData(xmap, ymap).TileExit.Map > 0 Then
            MapSup = 1
            Exit For
        Else
        'Debug.Print "NO hay traslados"
        End If

    Next
End Sub

Private Sub Puertatipo2(ByVal X As Integer, ByVal Y As Integer, ByVal ObjetoType As Integer, ByVal SuptipoObj As Integer)
    '*********************************************************************************************************
    ' Falta la IA
    '*********************************************************************************************************
    If ObjData(MapData(X, Y).OBJInfo.ObjIndex).ObjType = 6 And ObjData(MapData(X, Y).OBJInfo.ObjIndex).Subtipo = 2 Then
    
        If X > (1) And X < (MapSize.Width) And Y > (1) And Y < (MapSize.Height) Then
    
            If Not IsBlock(X + 2, Y - 1) And Not ((MapData(X + 2, Y - 1).Blocked And 1) <> 0 And (MapData(X + 2, Y + 2).Blocked And 4) <> 0) Then
                Print #handle, MapName & " ::: Posición: " & X + 2 & ", " & Y - 1 & " :::: FALTA BLOQUEO TOTAL"
            End If
    
            If Not IsBlock(X - 2, Y - 1) And Not ((MapData(X - 2, Y - 1).Blocked And 1) <> 0 And (MapData(X - 2, Y + 2).Blocked And 4) <> 0) Then
                Print #handle, MapName & " ::: Posición: " & X - 2 & ", " & Y - 1 & " :::: FALTA BLOQUEO TOTAL"
            End If
    
            If ObjData(MapData(X, Y).OBJInfo.ObjIndex).Cerrada = 1 Then
    
                If (MapData(X - 1, Y - 1).Blocked And 1) = 0 Then
                    Print #handle, MapName & " ::: Posición: " & X - 1 & ", " & Y - 1 & " :::: FALTA BLOQUEO PARCIAL"
                End If
    
                If (MapData(X, Y - 1).Blocked And 1) = 0 Then
                    Print #handle, MapName & " ::: Posición: " & X & ", " & Y - 1 & " :::: FALTA BLOQUEO PARCIAL"
                End If
    
                If (MapData(X + 1, Y - 1).Blocked And 1) = 0 Then
                    Print #handle, MapName & " ::: Posición: " & X + 1 & ", " & Y - 1 & " :::: FALTA BLOQUEO PARCIAL"
                End If
    
                If (MapData(X - 1, Y).Blocked And 4) = 0 Then
                    Print #handle, MapName & " ::: Posición: " & X - 1 & ", " & Y & " :::: FALTA BLOQUEO PARCIAL"
                End If
    
                If (MapData(X, Y).Blocked And 4) = 0 Then
                    Print #handle, MapName & " ::: Posición: " & X & ", " & Y & " :::: FALTA BLOQUEO PARCIAL"
                End If
    
                If (MapData(X + 1, Y).Blocked And 4) = 0 Then
                    Print #handle, MapName & " ::: Posición: " & X + 1 & ", " & Y & " :::: FALTA BLOQUEO PARCIAL"
                End If
    
            Else
    
                If (MapData(X, Y - 1).Blocked And 1) <> 0 Then
                    Print #handle, MapName & " ::: Posición: " & X & ", " & Y - 1 & " :::: HAY BLOQUEO PARCIAL"
                End If
    
                If (MapData(X - 1, Y - 1).Blocked And 1) <> 0 Then
                    Print #handle, MapName & " ::: Posición: " & X - 1 & ", " & Y - 1 & " :::: HAY BLOQUEO PARCIAL"
                End If
    
                If (MapData(X - 1, Y).Blocked And 4) <> 0 Then
                    Print #handle, MapName & " ::: Posición: " & X - 1 & ", " & Y & " :::: HAY BLOQUEO PARCIAL"
                End If
    
                If (MapData(X + 1, Y - 1).Blocked And 1) <> 0 Then
                    Print #handle, MapName & " ::: Posición: " & X + 1 & ", " & Y - 1 & " :::: FALTA BLOQUEO PARCIAL"
                End If
    
                If (MapData(X + 1, Y).Blocked And 4) <> 0 Then
                    Print #handle, MapName & " ::: Posición: " & X + 1 & ", " & Y & " :::: FALTA BLOQUEO PARCIAL"
                End If
    
                If (MapData(X, Y).Blocked And 4) <> 0 Then
                    Print #handle, MapName & " ::: Posición: " & X & ", " & Y & " :::: HAY BLOQUEO PARCIAL"
                End If
            End If
        End If
    End If
End Sub

Private Sub FloresPlantas(ByVal X As Integer, ByVal Y As Integer)

    If chkArboles.value = 1 Then
    
        iz = 0
    
        If IsNorte(X, Y - 1) Then iz = iz + Norte
        If IsEste(X + 1, Y) Then iz = iz + Este
        If IsOeste(X - 1, Y) Then iz = iz + Oeste
        
        If (MapData(X, Y).Blocked <> 1 + iz) Then
            MapData(X, Y).Blocked = 1 + iz
            MapInfo.Changed = 1
            Print #handle, MapName & " ::: Posición del Cartel: " & X & ", " & Y & " :::: Faltaba Bloqueo " & MapData(X, Y).Graphic(3).grhindex & " ya fue puesto."
        End If
                    
        iz = 0
    
        If IsSur(X, Y + 2) Then iz = iz + Sur
        If IsOeste(X - 1, Y + 1) Then iz = iz + Oeste
        If IsEste(X + 1, Y + 1) Then iz = iz + Este
    
        If (MapData(X, Y + 1).Blocked <> 4 + iz) Then
            Print #handle, MapName & " ::: Posición del Cartel: " & X & ", " & Y + 1 & " :::: Faltaba Bloqueo " & MapData(X, Y).Graphic(3).grhindex & " ya fue puesto."
            MapData(X, Y + 1).Blocked = 4 + iz
            MapInfo.Changed = 1
        End If
    
    Else
        iz = 0
    
        If IsNorte(X, Y - 1) Then iz = iz + Norte
        If IsEste(X + 1, Y) Then iz = iz + Este
        If IsOeste(X - 1, Y) Then iz = iz + Oeste
        If (MapData(X, Y).Blocked <> 1 + iz) Then
            Print #handle, MapName & " ::: Posición del Cartel: " & X & ", " & Y & " :::: Falta Bloqueo " & MapData(X, Y).Graphic(3).grhindex
        End If
                    
        iz = 0
    
        If IsSur(X, Y + 2) Then iz = iz + Sur
        If IsOeste(X - 1, Y + 1) Then iz = iz + Oeste
        If IsEste(X + 1, Y + 1) Then iz = iz + Este
        If (MapData(X, Y + 1).Blocked <> 4 + iz) Then
            Print #handle, MapName & " ::: Posición del Cartel: " & X & ", " & Y + 1 & " :::: Falta Bloqueo " & MapData(X, Y).Graphic(3).grhindex
        End If
    End If
End Sub

Private Sub Faroles(ByVal X As Integer, ByVal Y As Integer)

                    If chkFaroles.value = 1 Then

                        'grafico 5624
                        If (MapData(X, Y).Graphic(3).grhindex = 5624) Then
                            If IsNorte(X, Y - 1) Then iz = iz + Norte
                            If IsSur(X, Y + 1) Then iz = iz + 1
                            If IsOeste(X - 1, Y) Then iz = iz + 2
                            If MapData(X, Y).Blocked <> (8 + iz) Then
                                Print #handle, MapName & " ::: Posición del Farol: " & X & ", " & Y & " :::: FALTA Bloqueo o diferente a Bloq Oeste y Este  .Grafico;" & MapData(X, Y).Graphic(3).grhindex & " Se inserto Bloq : " & (8 + iz)
                                MapData(X, Y).Blocked = (8 + iz)
                                MapInfo.Changed = 1
                            End If
                            
                        End If

                        If (MapData(X, Y).Graphic(3).grhindex = 5624) Then
                            If IsNorte(X + 1, Y - 1) Then de = de + 4
                            If IsSur(X + 1, Y + 1) Then de = de + 1
                            If IsEste(X + 2, Y) Then de = de + 8
                            If MapData(X + 1, Y).Blocked <> (2 + de) Then
                                Print #handle, MapName & " ::: Posición del Farol: " & X + 1 & ", " & Y & " :::: FALTA Bloqueo o diferente a Bloq Oeste y Este  .Grafico;" & MapData(X, Y).Graphic(3).grhindex & " Se inserto Bloq : " & (2 + de)
                                MapData(X + 1, Y).Blocked = (2 + de)
                                MapInfo.Changed = 1
                            End If
                        End If

                        If (MapData(X, Y).Graphic(3).grhindex = 5624) Then
                            If MapData(X - 1, Y - 2).particle_Index <> 235 Then
                                Print #handle, MapName & " ::: Posición del Particula: " & X - 1 & ", " & Y - 2 & " :::: Se puso la Particula = 235"
                                MapData(X - 1, Y - 2).particle_Index = 235
                                MapInfo.Changed = 1
                            End If

                        End If
                        'grafico 5624 final
                        
                        'grafico 5625
                        If (MapData(X, Y).Graphic(3).grhindex = 5625) Then
                        
                            If IsNorte(X, Y - 1) Then iz = iz + 4
                            If IsSur(X, Y + 1) Then iz = iz + 1
                            If IsEste(X + 1, Y) Then iz = iz + 8
                            If MapData(X, Y).Blocked <> (2 + iz) Then
                                Print #handle, MapName & " ::: Posición del Farol: " & X & ", " & Y & " :::: FALTA Bloqueo o diferente a Bloq Oeste y Este  .Grafico;" & MapData(X, Y).Graphic(3).grhindex & " Se inserto Bloq : " & (2 + iz)
                                MapData(X, Y).Blocked = (2 + iz)
                                MapInfo.Changed = 1
                            End If
                        End If

                        If (MapData(X, Y).Graphic(3).grhindex = 5625) Then
                        
                            If IsNorte(X - 1, Y - 1) Then de = de + 4
                            If IsSur(X - 1, Y + 1) Then de = de + 1
                            If IsOeste(X - 2, Y) Then de = de + 2
                            If MapData(X - 1, Y).Blocked <> (8 + de) Then
                                Print #handle, MapName & " ::: Posición del Farol: " & X - 1 & ", " & Y & " :::: FALTA Bloqueo o diferente a Bloq Oeste y Este  .Grafico;" & MapData(X, Y).Graphic(3).grhindex & " Se inserto Bloq : " & (8 + de)
                                MapData(X - 1, Y).Blocked = (8 + de)
                                MapInfo.Changed = 1
                            End If
                        End If

                        If (MapData(X, Y).Graphic(3).grhindex = 5625) Then
                            If MapData(X + 1, Y - 2).particle_Index <> 234 Then
                                Print #handle, MapName & " ::: Posición del Particula: " & X + 1 & ", " & Y - 2 & " :::: Se puso la Particula = 234"
                                MapData(X + 1, Y - 2).particle_Index = 234
                                MapInfo.Changed = 1
                            End If
                        End If

                        'grafico 5625 final
                        'grafico 5627
                        If (MapData(X, Y).Graphic(3).grhindex = 5627 Or MapData(X, Y).Graphic(3).grhindex = 12747 Or MapData(X, Y).Graphic(3).grhindex = 12748) And (MapData(X, Y).Blocked <> 15) <> 0 Then
                            Print #handle, MapName & " ::: Posición del Farol: " & X & ", " & Y & " :::: FALTA Bloqueo o diferente a Bloq Total  .Grafico;" & MapData(X, Y).Graphic(3).grhindex & " Se inserto Bloq : 15 "
                            MapData(X, Y).Blocked = 15
                            MapInfo.Changed = 1
                        End If
                        
                        If MapData(X, Y).Graphic(3).grhindex = 12749 And (MapData(X - 1, Y).Blocked <> 15) <> 0 Then
                            Print #handle, MapName & " ::: Posición del Farol: " & X & ", " & Y & " :::: FALTA Bloqueo o diferente a Bloq Total  .Grafico;" & MapData(X - 1, Y).Graphic(3).grhindex & " Se inserto Bloq : 15 "
                            MapData(X - 1, Y).Blocked = 15
                            MapInfo.Changed = 1
                        End If
                        
                        If MapData(X, Y).Graphic(3).grhindex = 12750 And (MapData(X + 1, Y).Blocked <> 15) <> 0 Then
                            Print #handle, MapName & " ::: Posición del Farol: " & X & ", " & Y & " :::: FALTA Bloqueo o diferente a Bloq Total  .Grafico;" & MapData(X + 1, Y).Graphic(3).grhindex & " Se inserto Bloq : 15 "
                            MapData(X + 1, Y).Blocked = 15
                            MapInfo.Changed = 1
                        End If

                        If (MapData(X, Y).Graphic(3).grhindex = 5627) Then
                            If MapData(X, Y - 1).particle_Index <> 236 Then
                                Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y - 1 & " :::: Se puso la Particula = 236"
                                MapData(X, Y - 1).particle_Index = 236
                                MapInfo.Changed = 1
                            End If
                        End If
                        'grafico 5627 final

                        'grafico 5626
                        If (MapData(X, Y).Graphic(3).grhindex = 5626) And (MapData(X, Y).Blocked <> 15) <> 0 Then
                            Print #handle, MapName & " ::: Posición del Farol: " & X & ", " & Y & " :::: FALTA Bloqueo o diferente a Bloq Total  .Grafico;" & MapData(X, Y).Graphic(3).grhindex & " Se inserto Bloq : 15 "
                            MapData(X, Y).Blocked = 15
                            MapInfo.Changed = 1
                        End If

                        If (MapData(X, Y).Graphic(3).grhindex = 5626) Then
                            If MapData(X, Y - 2).particle_Index <> 237 Then
                                Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y - 2 & " :::: Se puso la Particula = 237  -" & MapData(X, Y - 2).particle_Index
                                MapData(X, Y - 2).particle_Index = 237
                                MapInfo.Changed = 1
                            End If
                        End If
                        'grafico 5626 final
                        
                    Else
                       
                        'grafico 5624
                        If (MapData(X, Y).Graphic(3).grhindex = 5624) Then
                            If IsNorte(X, Y - 1) Then iz = iz + 4
                            If IsSur(X, Y + 1) Then iz = iz + 1
                            If IsOeste(X - 1, Y) Then iz = iz + 2
                            If MapData(X, Y).Blocked <> (8 + iz) Then
                                Print #handle, MapName & " ::: Posición del Farol: " & X & ", " & Y & " :::: FALTA Bloqueo en el farol o diferente a Bloq Oeste  .Grafico;" & MapData(X, Y).Graphic(3).grhindex
                            End If
                            
                        End If

                        If (MapData(X, Y).Graphic(3).grhindex = 5624) Then
                            If IsNorte(X + 1, Y - 1) Then de = de + 4
                            If IsSur(X + 1, Y + 1) Then de = de + 1
                            If IsEste(X + 2, Y) Then de = de + 8
                            If MapData(X + 1, Y).Blocked <> (2 + de) Then
                                Print #handle, MapName & " ::: Posición del Farol: " & X + 1 & ", " & Y & " :::: FALTA Bloqueo en el farol o diferente a Bloq Este .Grafico;" & MapData(X, Y).Graphic(3).grhindex
                            End If
                        End If

                        If (MapData(X, Y).Graphic(3).grhindex = 5624) Then
                            If MapData(X - 1, Y - 2).particle_Index <> 235 Then
                                Print #handle, MapName & " ::: Posición del Particula: " & X - 1 & ", " & Y - 2 & " :::: Se puso la Particula = 235"
                                MapData(X - 1, Y - 2).particle_Index = 235
                            End If

                        End If
                        'grafico 5624 final
                        
                        'grafico 5625
                        If (MapData(X, Y).Graphic(3).grhindex = 5625) Then
                            If IsNorte(X, Y - 1) Then iz = iz + 4
                            If IsSur(X, Y + 1) Then iz = iz + 1
                            If IsEste(X + 1, Y) Then iz = iz + 8
                            If MapData(X, Y).Blocked <> (2 + iz) Then
                                Print #handle, MapName & " ::: Posición del Farol: " & X & ", " & Y & " :::: FALTA Bloqueo en el farol o diferente a Bloq Este .Grafico;" & MapData(X, Y).Graphic(3).grhindex
                            End If
                        End If

                        If (MapData(X, Y).Graphic(3).grhindex = 5625) Then
                            If IsNorte(X - 1, Y - 1) Then de = de + 4
                            If IsSur(X - 1, Y + 1) Then de = de + 1
                            If IsOeste(X - 2, Y) Then de = de + 2
                            If MapData(X - 1, Y).Blocked <> (8 + de) Then
                                Print #handle, MapName & " ::: Posición del Farol: " & X - 1 & ", " & Y & " :::: FALTA Bloqueo en el farol o diferente a Bloq  Oeste  .Grafico;" & MapData(X, Y).Graphic(3).grhindex
                            End If
                        End If

                        If (MapData(X, Y).Graphic(3).grhindex = 5625) Then
                            If MapData(X + 1, Y - 2).particle_Index <> 234 Then
                                Print #handle, MapName & " ::: Posición del Particula: " & X + 1 & ", " & Y - 2 & " :::: Falta la Particula = 234"
                            End If
                        End If

                        'grafico 5625 final
                        'grafico 5627
                        If (MapData(X, Y).Graphic(3).grhindex = 5627) And (MapData(X, Y).Blocked <> 15) <> 0 Then
                            Print #handle, MapName & " ::: Posición del Farol: " & X & ", " & Y & " :::: FALTA Bloqueo en el farol; " & MapData(X, Y).Graphic(3).grhindex
                        End If

                        If (MapData(X, Y).Graphic(3).grhindex = 5627) Then
                            If MapData(X, Y - 1).particle_Index <> 236 Then
                                Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y - 1 & " :::: Falta la Particula = 236"
                            End If
                        End If
                        'grafico 5627 final

                        If MapData(X, Y).Graphic(3).grhindex = 12749 And (MapData(X - 1, Y).Blocked <> 15) <> 0 Then
                            Print #handle, MapName & " ::: Posición del Farol: " & X - 1 & ", " & Y & " :::: FALTA Bloqueo o diferente a Bloq Total  .Grafico;" & MapData(X, Y).Graphic(3).grhindex & " Se inserto Bloq : 15 "

                        End If
                        
                        If MapData(X, Y).Graphic(3).grhindex = 12750 And (MapData(X + 1, Y).Blocked <> 15) <> 0 Then
                            Print #handle, MapName & " ::: Posición del Farol: " & X + 1 & ", " & Y & " :::: FALTA Bloqueo o diferente a Bloq Total  .Grafico;" & MapData(X, Y).Graphic(3).grhindex & " Se inserto Bloq : 15 "

                        End If
                        'grafico 5626
                        If (MapData(X, Y).Graphic(3).grhindex = 5626) And (MapData(X, Y).Blocked <> 15) <> 0 Then
                            Print #handle, MapName & " ::: Posición del Farol: " & X & ", " & Y & " :::: FALTA Bloqueo en el farol; " & MapData(X, Y).Graphic(3).grhindex
                        End If

                        If (MapData(X, Y).Graphic(3).grhindex = 5626) Then
                            If MapData(X, Y - 2).particle_Index <> 237 Then
                                Print #handle, MapName & " ::: Posición del Particula: " & X & ", " & Y - 2 & " :::: Falta la Particula = 237"
                            End If
                        End If
                        
                    End If

End Sub

Private Sub NPCsBordes(ByVal X As Integer, ByVal Y As Integer)

    For Y = 1 To MapSize.Height
        For X = 1 To MapSize.Width
        
        
'           **********************************************************************
'             pongo grafico doble a arboles

'            If MapData(X, y).OBJInfo.objindex Then
'                If (ObjData(MapData(X, y).OBJInfo.objindex).ObjType = 4 And (MapData(X, y).Graphic(3).grhindex = 0)) Then
'                    If chkArboles.Value = 1 Then
'                        Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & y & " :::: Árbol con doble Grafico :" & MapData(X, y).Graphic(3).grhindex
'                        MapData(X, y).Graphic(3).grhindex = MapData(X, y).ObjGrh.grhindex
'                        MapInfo.Changed = 1
'                    Else
'                        Print #handle, MapName & " ::: Posición del Árbol: " & X & ", " & y & " :::: Árbol Falta el doble Grafico " & MapData(X, y).Graphic(3).grhindex
'
'                    End If
'                End If
'            End If
'            ********************************************************************
                'Remplaso Bloques  4*4
                
'               Dim igraf As Integer
'
'               Dim PrimerGraficoOLD As Long
'               Dim PrimerGraficoNEW As Long
'
'               PrimerGraficoOLD = 7912
'               PrimerGraficoNEW = 7912
'
'               For igraf = 0 To 15
'
'                           If MapData(X, y).Graphic(1).grhindex = (PrimerGraficoOLD + igraf) Then
'                                   Print #handle, MapName & " ::: Posición camino : " & X & ", " & y & " ; " & MapData(X, y).Graphic(1).grhindex
'                                   Debug.Print igraf
'                                   MapData(X, y).Graphic(1).grhindex = 0
'                                   MapData(X, y).Graphic(3).grhindex = (PrimerGraficoNEW + igraf)
'
'                                   MapInfo.Changed = 1
'
'                           End If
'                Next igraf



            ' ** Quitar NPCs, Objetos y Translados en los Bordes Exteriores
            If (X < 12 Or X > 88 Or Y < 10 Or Y > 91) Then
                If chkQuitarNPCs.value = 1 Then
    
                    'Quitar NPCs
                    If MapData(X, Y).NpcIndex > 0 Then
                        Print #handle, MapName & " ::: Posición del NPC: " & X & ", " & Y & " :::: Saco el NPC fuera del Mapa." & MapData(X, Y).NpcIndex
                        EraseChar MapData(X, Y).CharIndex
                        MapData(X, Y).NpcIndex = 0
                        MapInfo.Changed = 1
                    End If
    
                    ' Quitar Objetos
                    '                MapData(X, Y).OBJInfo.objindex = 0
                    '                MapData(X, Y).OBJInfo.Amount = 0
                    '                MapData(X, Y).ObjGrh.grhindex = 0
                    ' Quitar Translado
                    If MapData(X, Y).TileExit.Map > 0 Then
                        Print #handle, MapName & " ::: Posición del transaldo fuera del mapa: " & X & ", " & Y & " :::: Translado a :" & MapData(X, Y).TileExit.Map & ", " & MapData(X, Y).TileExit.X & ", " & MapData(X, Y).TileExit.Y & " Eliminado"
                        MapData(X, Y).TileExit.Map = 0
                        MapData(X, Y).TileExit.X = 0
                        MapData(X, Y).TileExit.Y = 0
                        MapInfo.Changed = 1
                    End If

                    ' Quitar Triggers
                    If MapData(X, Y).Trigger > 0 Then
                        Print #handle, MapName & " ::: Posición del fuera del mapa: " & X & ", " & Y & " :::: Trigger's :" & MapData(X, Y).Trigger & " Eliminado"
                        MapData(X, Y).Trigger = 0
                        MapInfo.Changed = 1
                    End If
                    
                Else

                    If MapData(X, Y).NpcIndex > 0 Then
                        Print #handle, MapName & " ::: Posición del NPC: " & X + 1 & ", " & Y & " :::: NPC fuera del Mapa." & MapData(X, Y).NpcIndex
                    End If
    
                    If MapData(X, Y).TileExit.Map > 0 Then
                        Print #handle, MapName & " ::: Posición del transaldo fuera del mapa: " & X & ", " & Y & " :::: Translado a : & MapData(X, Y).TileExit.Map & ", " & MapData(X, Y).TileExit.X & ", " & MapData(X, Y).TileExit.Y"
                    End If
                    
                    If MapData(X, Y).Trigger > 0 Then
                        Print #handle, MapName & " ::: Posición del Trigger fuera del mapa: " & X & ", " & Y & " :::: Trigger's :" & MapData(X, Y).Trigger
                    End If
                    
                
                End If
            End If
        Next X
    Next Y

End Sub

Private Sub cabañas(ByVal X As Integer, ByVal Y As Integer, ByVal Grh As Long)

    '***************************************************************************************
    'Bloqeos Cabañas by ReyarB
    '***************************************************************************************
    If chkCasas.value = 1 Then
        If MapData(X, Y).Graphic(3).grhindex = 5307 Then
            If MapData(X - 1, Y + 1).Blocked <> 12 Or MapData(X - 1, Y).Blocked <> 1 And Not MapData(X, Y).Blocked = 15 Then
                Print #handle, MapName & " ::: Faltan bloqueos en : " & X - 1 & ", " & Y + 1 & " ::::  SE podia entrar faltaba bloqueo en la Pared de la casa."
                Print #handle, MapName & " ::: Faltan bloqueos en : " & X - 1 & ", " & Y & " ::::  SE podia entrar faltaba bloqueo en la Pared de la casa."
                MapData(X - 1, Y + 1).Blocked = 12
                MapData(X - 1, Y).Blocked = 1
                MapInfo.Changed = 1
            End If
        End If
    Else

        If MapData(X - 1, Y + 1).Blocked <> 12 Or MapData(X - 1, Y).Blocked <> 1 And Not MapData(X, Y).Blocked = 15 Then
            Print #handle, MapName & " ::: Faltan bloqueos en : " & X & ", " & Y & " ::::  Se puede entrar falta bloqueo en la Pared de la casa."
    
        End If
    End If

    If chkCasas.value = 1 Then
        If MapData(X, Y).Graphic(3).grhindex = 5309 Then
            If MapData(X, Y).Blocked <> 2 Or MapData(X - 1, Y).Blocked <> 8 Or MapData(X, Y - 1).Blocked <> 2 Or MapData(X - 1, Y - 1).Blocked <> 8 Or MapData(X, Y - 2).Blocked <> 2 Or MapData(X - 1, Y - 2).Blocked <> 8 And Not MapData(X, Y).Blocked = 15 Then
                Print #handle, MapName & " ::: Faltan bloqueos en : " & X & ", " & Y & " ::::  SE podia entrar faltaba bloqueo en la Pared de la casa."

                If IsNorte(X, Y - 3) Then iz = iz + Norte

                MapData(X, Y).Blocked = 2
                MapData(X - 1, Y).Blocked = 8
                MapData(X, Y - 1).Blocked = 2
                MapData(X - 1, Y - 1).Blocked = 8
                MapData(X, Y - 2).Blocked = 2 + iz
                MapData(X - 1, Y - 2).Blocked = 8

                MapInfo.Changed = 1
            End If
        End If
    Else

        If MapData(X, Y).Blocked <> 2 Or MapData(X - 1, Y).Blocked <> 8 Or MapData(X, Y - 1).Blocked <> 2 Or MapData(X - 1, Y - 1).Blocked <> 8 Or MapData(X, Y - 2).Blocked <> 2 Or MapData(X - 1, Y - 2).Blocked <> 8 And Not MapData(X, Y).Blocked = 15 Then
            Print #handle, MapName & " ::: Faltan bloqueos en : " & X & ", " & Y & " ::::  SE podia entrar faltaba bloqueo en la Pared de la casa."
        End If
    End If

    If chkCasas.value = 1 Then
        If MapData(X, Y).Graphic(3).grhindex = 5645 Then
            If MapData(X, Y).Blocked <> 2 Or MapData(X - 1, Y).Blocked <> 8 Or MapData(X, Y - 1).Blocked <> 2 Or MapData(X - 1, Y - 1).Blocked <> 8 Or MapData(X, Y - 2).Blocked <> 2 Or MapData(X - 1, Y - 2).Blocked <> 8 And Not MapData(X, Y).Blocked = 15 Then
                Print #handle, MapName & " ::: Faltan bloqueos en : " & X & ", " & Y & " ::::  SE podia entrar faltaba bloqueo en la Pared de la casa."

                If IsNorte(X, Y - 3) Then iz = iz + Norte

                MapData(X, Y).Blocked = 2
                MapData(X - 1, Y).Blocked = 8
                MapData(X, Y - 1).Blocked = 2
                MapData(X - 1, Y - 1).Blocked = 8
                MapData(X, Y - 2).Blocked = 2 + iz
                MapData(X - 1, Y - 2).Blocked = 8
            
                MapInfo.Changed = 1
            End If
        End If
    Else

        If MapData(X, Y).Blocked <> 2 Or MapData(X - 1, Y).Blocked <> 8 Or MapData(X, Y - 1).Blocked <> 2 Or MapData(X - 1, Y - 1).Blocked <> 8 Or MapData(X, Y - 2).Blocked <> 2 Or MapData(X - 1, Y - 2).Blocked <> 8 And Not MapData(X, Y).Blocked = 15 Then
            Print #handle, MapName & " ::: Faltan bloqueos en : " & X & ", " & Y & " ::::  SE podia entrar faltaba bloqueo en la Pared de la casa."
        End If
    End If

    If chkCasas.value = 1 Then
        If MapData(X, Y).Graphic(3).grhindex = 5643 Or MapData(X, Y).Graphic(3).grhindex = 5847 Then
            If MapData(X, Y).Blocked <> 2 Or MapData(X - 1, Y).Blocked <> 8 Or MapData(X, Y - 1).Blocked <> 2 Or MapData(X - 1, Y - 1).Blocked <> 8 Or MapData(X, Y - 2).Blocked <> 2 Or MapData(X - 1, Y - 2).Blocked <> 8 And Not MapData(X, Y).Blocked = 15 Then
                Print #handle, MapName & " ::: Faltan bloqueos en : " & X & ", " & Y & " ::::  SE podia entrar faltaba bloqueo en la Pared de la casa."

                If IsNorte(X - 1, Y - 3) Then iz = iz + Norte

                MapData(X, Y).Blocked = 2
                MapData(X - 1, Y).Blocked = 8
                MapData(X, Y - 1).Blocked = 2
                MapData(X - 1, Y - 1).Blocked = 8
                MapData(X, Y - 2).Blocked = 2
                MapData(X - 1, Y - 2).Blocked = 8 + iz

                MapInfo.Changed = 1
            End If
        End If
    Else

        If MapData(X, Y).Blocked <> 2 Or MapData(X - 1, Y).Blocked <> 8 Or MapData(X, Y - 1).Blocked <> 2 Or MapData(X - 1, Y - 1).Blocked <> 8 Or MapData(X, Y - 2).Blocked <> 2 Or MapData(X - 1, Y - 2).Blocked <> 8 And Not MapData(X, Y).Blocked = 15 Then
            Print #handle, MapName & " ::: Faltan bloqueos en : " & X & ", " & Y & " ::::  SE podia entrar faltaba bloqueo en la Pared de la casa."
        End If
    End If

    If chkCasas.value = 1 Then
        If MapData(X, Y).Graphic(3).grhindex = 5682 Then
            If MapData(X, Y).Blocked <> 1 Or MapData(X - 1, Y + 1).Blocked <> 4 And Not MapData(X, Y).Blocked = 15 Then
                Print #handle, MapName & " ::: Faltan bloqueos en : " & X & ", " & Y & " ::::  SE podia entrar faltaba bloqueo en la Pared de la casa."
                MapData(X, Y).Blocked = 1
                MapData(X, Y + 1).Blocked = 4
                MapData(X - 1, Y).Blocked = 1
                MapData(X - 1, Y + 1).Blocked = 4

                MapInfo.Changed = 1
            End If
        End If
    Else

        If MapData(X, Y).Blocked <> 1 Or MapData(X - 1, Y + 1).Blocked <> 4 And Not MapData(X, Y).Blocked = 15 Then
            Print #handle, MapName & " ::: Faltan bloqueos en : " & X & ", " & Y & " ::::  SE podia entrar faltaba bloqueo en la Pared de la casa."
        End If
    End If

    If chkCasas.value = 1 Then
        If MapData(X, Y).Graphic(3).grhindex = 5678 Or MapData(X, Y).Graphic(3).grhindex = 5683 Then
            If MapData(X + 1, Y).Blocked <> 1 Or MapData(X + 1, Y + 1).Blocked <> 4 And Not MapData(X, Y).Blocked = 15 Then
                Print #handle, MapName & " ::: Faltan bloqueos en : " & X + 1 & ", " & Y & " ::::  SE podia entrar faltaba bloqueo en la Pared de la casa."
                MapData(X + 1, Y).Blocked = 1
                MapData(X + 1, Y + 1).Blocked = 4

                MapInfo.Changed = 1
            End If
        End If
    Else

        If MapData(X + 1, Y).Blocked <> 1 Or MapData(X + 1, Y + 1).Blocked <> 4 And Not MapData(X, Y).Blocked = 15 Then
            Print #handle, MapName & " ::: Faltan bloqueos en : " & X + 1 & ", " & Y & " ::::  SE podia entrar faltaba bloqueo en la Pared de la casa."
        End If
    End If

    If chkCasas.value = 1 Then
        If MapData(X, Y).Graphic(3).grhindex = 5676 Or MapData(X, Y).Graphic(3).grhindex = 5681 Then
            If MapData(X, Y).Blocked <> 3 Or MapData(X - 1, Y).Blocked <> 8 Or MapData(X + 1, Y).Blocked <> 1 Or MapData(X + 1, Y + 1).Blocked <> 4 Or MapData(X, Y + 1).Blocked <> 4 Or MapData(X, Y - 1).Blocked <> 2 Or MapData(X - 1, Y - 1).Blocked <> 8 Or MapData(X, Y - 2).Blocked <> 2 Or MapData(X - 1, Y - 2).Blocked <> 8 And Not MapData(X, Y).Blocked = 15 Then
                Print #handle, MapName & " ::: Faltan bloqueos en : " & X & ", " & Y & " ::::  SE podia entrar faltaba bloqueo en la Pared de la casa."
                MapData(X, Y).Blocked = 3
                MapData(X - 1, Y).Blocked = 8
                MapData(X, Y + 1).Blocked = 4
                MapData(X, Y - 1).Blocked = 2
                MapData(X - 1, Y - 1).Blocked = 8
                MapData(X, Y - 2).Blocked = 2
                MapData(X - 1, Y - 2).Blocked = 8
                MapData(X + 1, Y).Blocked = 1
                MapData(X + 1, Y + 1).Blocked = 4

                MapInfo.Changed = 1
            End If
        End If
    Else

        If MapData(X, Y).Blocked <> 3 Or MapData(X - 1, Y).Blocked <> 8 Or MapData(X + 1, Y).Blocked <> 1 Or MapData(X + 1, Y + 1).Blocked <> 4 Or MapData(X, Y + 1).Blocked <> 4 Or MapData(X, Y - 1).Blocked <> 2 Or MapData(X - 1, Y - 1).Blocked <> 8 Or MapData(X, Y - 2).Blocked <> 2 Or MapData(X - 1, Y - 2).Blocked <> 8 And Not MapData(X, Y).Blocked = 15 Then
            Print #handle, MapName & " ::: Faltan bloqueos en : " & X & ", " & Y & " ::::  SE podia entrar faltaba bloqueo en la Pared de la casa."
        End If
    End If

    If chkCasas.value = 1 Then
        If MapData(X, Y).Graphic(3).grhindex = 5680 Then
            If MapData(X - 1, Y).Blocked <> 9 Or MapData(X - 1, Y + 1).Blocked <> 4 Or MapData(X, Y).Blocked <> 2 Or MapData(X, Y - 1).Blocked <> 2 Or MapData(X, Y - 2).Blocked <> 2 Or MapData(X, Y - 1).Blocked <> 8 Or MapData(X, Y - 2).Blocked <> 8 And Not MapData(X - 1, Y).Blocked = 15 Then
                Print #handle, MapName & " ::: Faltan bloqueos en : " & X - 1 & ", " & Y & " ::::  SE podia entrar faltaba bloqueo en la Pared de la casa."
                MapData(X - 1, Y).Blocked = 9
                MapData(X, Y).Blocked = 2
                MapData(X - 1, Y + 1).Blocked = 4
                MapData(X - 1, Y - 1).Blocked = 8
                MapData(X, Y - 1).Blocked = 2
                MapData(X - 1, Y - 2).Blocked = 8
                MapData(X, Y - 2).Blocked = 2

                MapInfo.Changed = 1
            End If
        End If
    Else

        If MapData(X - 1, Y).Blocked <> 9 Or MapData(X - 1, Y + 1).Blocked <> 4 Or MapData(X, Y).Blocked <> 2 Or MapData(X, Y - 1).Blocked <> 2 Or MapData(X, Y - 2).Blocked <> 2 Or MapData(X, Y - 1).Blocked <> 8 Or MapData(X, Y - 2).Blocked <> 8 And Not MapData(X - 1, Y).Blocked = 15 Then
            Print #handle, MapName & " ::: Faltan bloqueos en : " & X - 1 & ", " & Y & " ::::  SE podia entrar faltaba bloqueo en la Pared de la casa."
        End If
    End If

    If chkCasas.value = 1 Then
        If MapData(X, Y).Graphic(3).grhindex = 5684 Then
            If MapData(X + 1, Y).Blocked <> 9 Or MapData(X + 1, Y + 1).Blocked <> 4 Or MapData(X + 2, Y).Blocked <> 2 Or MapData(X + 2, Y - 1).Blocked <> 2 Or MapData(X + 2, Y - 2).Blocked <> 2 Or MapData(X, Y - 1).Blocked <> 8 Or MapData(X, Y - 2).Blocked <> 8 And Not MapData(X + 1, Y).Blocked = 15 Then
                Print #handle, MapName & " ::: Faltan bloqueos en : " & X + 1 & ", " & Y & " ::::  SE podia entrar faltaba bloqueo en la Pared de la casa."
                MapData(X + 1, Y).Blocked = 9
                MapData(X + 2, Y).Blocked = 2
                MapData(X + 1, Y + 1).Blocked = 4
                MapData(X + 1, Y - 1).Blocked = 8
                MapData(X + 2, Y - 1).Blocked = 2
                MapData(X + 1, Y - 2).Blocked = 8
                MapData(X + 2, Y - 2).Blocked = 2

                MapData(X, Y).Blocked = 1
                MapData(X, Y + 1).Blocked = 4
                MapData(X - 1, Y).Blocked = 1
                MapData(X - 1, Y + 1).Blocked = 4

                MapInfo.Changed = 1
            End If
        End If
    Else

        If MapData(X + 1, Y).Blocked <> 9 Or MapData(X + 1, Y + 1).Blocked <> 4 Or MapData(X + 2, Y).Blocked <> 2 Or MapData(X + 2, Y - 1).Blocked <> 2 Or MapData(X + 2, Y - 2).Blocked <> 2 Or MapData(X, Y - 1).Blocked <> 8 Or MapData(X, Y - 2).Blocked <> 8 And Not MapData(X + 1, Y).Blocked = 15 Then
            Print #handle, MapName & " ::: Faltan bloqueos en : " & X + 1 & ", " & Y & " ::::  SE podia entrar faltaba bloqueo en la Pared de la casa."
        End If
    End If

    If MapData(X, Y).Graphic(3).grhindex = 5687 Or MapData(X, Y).Graphic(3).grhindex = 5688 Or MapData(X, Y).Graphic(3).grhindex = 5689 Or MapData(X, Y).Graphic(3).grhindex = 5257 Or MapData(X, Y).Graphic(3).grhindex = 5305 Or MapData(X, Y).Graphic(3).grhindex = 5306 Or MapData(X, Y).Graphic(3).grhindex = 5677 Or MapData(X, Y).Graphic(3).grhindex = 5679 Then
        Dim Costado As Integer
        Dim i       As Integer

        If chkCasas.value = 1 Then
            iz = 0

            For i = 1 To 3

                If IsNorte(X + 2 - i, Y - 1) Then iz = iz + Norte

                'If IsSur(X + 2 - i, Y + 1) Then iz = iz + Sur
                If IsEste(X + 2 - i + 1, Y) Then iz = iz + Este
                If IsOeste(X - 1 - i + 1, Y) Then iz = iz + Oeste
                If iz = 16 Then iz = 0
                If MapData(X + 2 - i, Y).Blocked <> (1 + iz) And Not MapData(X + 2 - i, Y).Blocked = 15 Then
                    Print #handle, MapName & " ::: Faltan bloqueos en : " & X + 2 - i & ", " & Y & " ::::  SE podia entrar faltaba bloqueo en la Pared de la casa."
                    MapData(X + 2 - i, Y).Blocked = (1 + iz)
                    MapInfo.Changed = 1
                End If
                iz = 0
            Next

            For i = 1 To 3

                'If IsNorte(X + 1 - i, Y) Then iz = iz + Norte
                If IsSur(X + 2 - i, Y + 2) Then iz = iz + Sur
                If IsEste(X + 2 - i + 1, Y + 1) Then iz = iz + Este
                If IsOeste(X - 1 - i + 1, Y + 1) Then iz = iz + Oeste
                If iz = 16 Then iz = 0
                If MapData(X + 2 - i, Y + 1).Blocked <> (4 + iz) And Not MapData(X + 2 - i, Y + 1).Blocked = 15 Then
                    Print #handle, MapName & " ::: Faltan bloqueos en : " & X + 1 - i & ", " & Y + 1 & " ::::  SE podia entrar faltaba bloqueo en la Pared de la casa."
                    MapData(X + 2 - i, Y + 1).Blocked = (4 + iz)
                    MapInfo.Changed = 1
                End If
                iz = 0
            Next

        Else

            For i = 1 To 3

                If IsNorte(X + 2 - i, Y - 1) Then iz = iz + Norte

                'If IsSur(X + 2 - i, Y + 1) Then iz = iz + Sur
                If IsEste(X + 2 - i + 1, Y) Then iz = iz + Este
                If IsOeste(X - 2 - i + 1, Y) Then iz = iz + Oeste
                If iz = 16 Then iz = 0
                If MapData(X + 2 - i, Y).Blocked <> (1 + iz) And Not MapData(X + 2 - i, Y).Blocked = 15 Then
                    Print #handle, MapName & " ::: Faltan bloqueos en : " & X + 2 - i & ", " & Y & " ::::  SE puede entrar falta bloqueo en la Pared de la casa."
                End If
                iz = 0
            Next

            For i = 1 To 3

                'If IsNorte(X + 2 - i, Y) Then iz = iz + Norte
                If IsSur(X + 2 - i, Y + 2) Then iz = iz + Sur
                If IsEste(X + 2 - i + 1, Y + 1) Then iz = iz + Este
                If IsOeste(X - 2 - i + 1, Y + 1) Then iz = iz + Oeste
                If iz = 16 Then iz = 0
                If MapData(X + 2 - i, Y + 1).Blocked <> (4 + iz) And Not MapData(X + 2 - i, Y + 1).Blocked = 15 Then
                    Print #handle, MapName & " ::: Faltan bloqueos en : " & X + 2 - i & ", " & Y + 1 & " ::::  SE puede entrar falta bloqueo en la Pared de la casa."
                End If
                iz = 0
            Next

        End If

    End If

    If MapData(X, Y).Graphic(3).grhindex = 5256 Or MapData(X, Y).Graphic(3).grhindex = 5576 Or MapData(X, Y).Graphic(3).grhindex = 5686 Then

        If chkCasas.value = 1 Then
            iz = 0

            For i = 1 To 2

                '                If IsNorte(X + 2 - i, Y - 1) Then iz = iz + Norte
                '                'If IsSur(X + 2 - i, Y + 1) Then iz = iz + Sur
                '                If IsEste(X + 2 - i + 1, Y) Then iz = iz + Este
                '                If IsOeste(X - 1 - i + 1, Y) Then iz = iz + Oeste
                '                If iz = 16 Then iz = 0
                If MapData(X + 2 - i, Y).Blocked <> (1 + iz) And Not MapData(X + 2 - i, Y).Blocked = 15 Then
                    Print #handle, MapName & " ::: Faltan bloqueos en : " & X + 2 - i & ", " & Y & " ::::  SE podia entrar faltaba bloqueo en la Pared de la casa."
                    MapData(X + 2 - i, Y).Blocked = (1 + iz)
                    MapInfo.Changed = 1
                End If
                iz = 0
            Next

            For i = 1 To 2

                '                'If IsNorte(X + 2 - i, Y) Then iz = iz + Norte
                '                If IsSur(X + 2 - i, Y + 2) Then iz = iz + Sur
                '                If IsEste(X + 2 - i + 1, Y + 1) Then iz = iz + Este
                '                If IsOeste(X - 2 + 1, Y + 1) Then iz = iz + Oeste
                '                If iz = 16 Then iz = 0
                If MapData(X + 2 - i, Y + 1).Blocked <> (4 + iz) And Not MapData(X + 2 - i, Y + 1).Blocked = 15 Then
                    Print #handle, MapName & " ::: Faltan bloqueos en : " & X + 2 - i & ", " & Y + 1 & " ::::  SE podia entrar faltaba bloqueo en la Pared de la casa."
                    MapData(X + 2 - i, Y + 1).Blocked = (4 + iz)
                    MapInfo.Changed = 1
                End If
                iz = 0
            Next

        Else

            For i = 1 To 2

                '                If IsNorte(X + 2 - i, Y - 1) Then iz = iz + Norte
                '                'If IsSur(X + 2 - i, Y + 1) Then iz = iz + Sur
                '                If IsEste(X + 2 - i + 1, Y) Then iz = iz + Este
                '                If IsOeste(X - 2 + 1, Y) Then iz = iz + Oeste
                '                If iz = 16 Then iz = 0
                If MapData(X + 2 - i, Y).Blocked <> (1 + iz) And Not MapData(X + 2 - i, Y).Blocked = 15 Then
                    Print #handle, MapName & " ::: Faltan bloqueos en : " & X + 2 - i & ", " & Y & " ::::  SE puede entrar falta bloqueo en la Pared de la casa."
                End If
                iz = 0
            Next

            For i = 1 To 2

                '                'If IsNorte(X + 2 - i, Y) Then iz = iz + Norte
                '                If IsSur(X + 2 - i, Y + 2) Then iz = iz + Sur
                '                If IsEste(X + 2 - i + 1, Y + 1) Then iz = iz + Este
                '                If IsOeste(X - 2 - i + 1, Y + 1) Then iz = iz + Oeste
                '                If iz = 16 Then iz = 0
                If MapData(X + 2 - i, Y + 1).Blocked <> (4 + iz) And Not MapData(X + 2 - i, Y + 1).Blocked = 15 Then
                    Print #handle, MapName & " ::: Faltan bloqueos en : " & X + 1 - i & ", " & Y + 1 & " ::::  SE puede entrar falta bloqueo en la Pared de la casa."
                End If
                iz = 0
            Next
    
        End If
     
    End If
    
End Sub

