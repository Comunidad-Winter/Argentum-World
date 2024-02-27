VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConsoleDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const p_CANTIDADDIALOGOS As Byte = 200
Public RENDER_DIALOGS As Byte
Public RENDER_DIALOGS_MAX As Byte
Private Type t_consoleDlg
    Texto As String
    red As Byte
    green As Byte
    blue As Byte
    status As Integer
    priv As Integer
End Type

Private Type console_modes
    Private As Boolean
    Guild As Boolean
    Normal As Boolean
    Combat As Boolean
    Party As Boolean
    Global As Boolean
End Type

Private console_mode As console_modes


Private p_Dialogos() As t_consoleDlg
'Private p_Dialogos_filtered() As t_consoleDlg
Private offset       As Single
Private usedFont     As StdFont

Private Const DIALOG_HEIGHT As Integer = 14
Public dialogIndex As Integer
Private dialogCount As Long
Private chatColor(3) As RGBA
Private dialog_offset As Double
Public Alpha As Single
Public scroll_offset As Integer
Private GRH_CONSOLE_DESPLEGAR As grh
Public console_height As Integer
Private consoleAlpha As Single
Public consoleAlpha_min As Single
Public consoleAlpha_max As Single
Public consoleAlpha_min_pos As Single
Public consoleAlpha_max_pos As Single
Public consola_transparente As Boolean
Public menu_consola_visible As Boolean
Public cambiando_intensidad As Boolean
Private alphaConsole As Integer
Private SelectedDialog As Integer

Private ConsoleTexture As Direct3DTexture8
Private ConsoleSurface As Direct3DSurface8
Private ConsoleProjection As D3DMATRIX


''
' Removes all dialogs from screen.

Public Sub RemoveDialogs()
    
    On Error GoTo RemoveDialogs_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/02/07
    'Removes all dialogs
    '***************************************************
    Dim i As Long
    
    For i = 1 To p_CANTIDADDIALOGOS
        p_Dialogos(i).Texto = vbNullString
        p_Dialogos(i).red = 255
        p_Dialogos(i).green = 255
        p_Dialogos(i).blue = 255
        p_Dialogos(i).priv = -1
        p_Dialogos(i).status = -1
    Next i

    
    Exit Sub

RemoveDialogs_Err:
    Call RegistrarError(err.Number, err.Description, "clsConsoleDlg.RemoveDialogs", Erl)
    Resume Next
    
End Sub

''
' Retrieves the number of dialogs to be displayed on screen.
'
' @return   The number of dialogs to be displayed on screen.

Public Property Get CantidadDialogos() As Byte
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/01/07
    'Retrieves the number of dialogs to be displayed on screen
    '***************************************************
    
    On Error GoTo CantidadDialogos_Err
    
    CantidadDialogos = p_CANTIDADDIALOGOS

    
    Exit Property

CantidadDialogos_Err:
    Call RegistrarError(err.Number, err.Description, "clsConsoleDlg.CantidadDialogos", Erl)
    Resume Next
    
End Property


Private Sub Class_Initialize()
    
    On Error GoTo Class_Initialize_Err
    
    ReDim p_Dialogos(0 To p_CANTIDADDIALOGOS - 1) As t_consoleDlg
    'ReDim p_Dialogos_filtered(0 To p_CANTIDADDIALOGOS - 1) As t_consoleDlg
    dialogIndex = 0
    dialogCount = 0
    RENDER_DIALOGS = 8
    console_height = 14 * RENDER_DIALOGS + 14
    Call InitGrh(GRH_CONSOLE_DESPLEGAR, 30139)
    GRH_CONSOLE_DESPLEGAR.Angle = 90
    
    Alpha = 255
    'Esto hay que meterlo en archivo.
    consoleAlpha_min_pos = 555
    consoleAlpha_max_pos = 620
    consoleAlpha_min = 30
    consoleAlpha_max = 200
    SelectedDialog = -1
    
    Exit Sub

Class_Initialize_Err:
    Call RegistrarError(err.Number, err.Description, "clsConsoleDlg.Class_Initialize", Erl)
    Resume Next
    
End Sub

Public Sub SetupConsoleTexture()
    ' Inicializo la textura
    Set ConsoleTexture = DirectDevice.CreateTexture(MainViewWidth, console_height, 0, D3DUSAGE_RENDERTARGET, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT)
    Set ConsoleSurface = ConsoleTexture.GetSurfaceLevel(0)
    Call D3DXMatrixOrthoOffCenterLH(ConsoleProjection, 0, MainViewWidth, console_height, 0, -1#, 1#)
    Call RedrawTexture
End Sub

Public Function MouseClick(ByVal x As Integer, ByVal y As Integer) As Boolean
    
    If X >= 700 And X <= 736 And Y > console_height And Y <= console_height + 20 Then
        If GRH_CONSOLE_DESPLEGAR.Angle = 90 Then
            GRH_CONSOLE_DESPLEGAR.Angle = 270
            RENDER_DIALOGS = 3
            console_height = 14 * RENDER_DIALOGS + 14
        Else
            GRH_CONSOLE_DESPLEGAR.Angle = 90
            RENDER_DIALOGS = RENDER_DIALOGS_MAX
            console_height = 14 * RENDER_DIALOGS + 14
        End If
        MouseClick = True
        
        Call SetupConsoleTexture
    
    End If
    
    If menu_consola_visible Then
        '587,20, 22,22
        'Global
        If X >= 591 And X <= 608 And Y >= 32 And Y <= 48 Then
            console_mode.Global = Not console_mode.Global
        End If
        
        'Party
        If X >= 654 And X <= 671 And Y >= 32 And Y <= 48 Then
            console_mode.Party = Not console_mode.Party
        End If
        
        'Clan
        If X >= 709 And X <= 731 And Y >= 32 And Y <= 48 Then
            console_mode.Guild = Not console_mode.Guild
        End If
        
        'Privado
        If X >= 558 And X <= 580 And Y >= 58 And Y <= 80 Then
            console_mode.Private = Not console_mode.Private
        End If
        
        'Combate
        If X >= 632 And X <= 654 And Y >= 58 And Y <= 80 Then
            console_mode.Combat = Not console_mode.Combat
        End If
        
        'Normal
        If X >= 707 And X <= 729 And Y >= 58 And Y <= 80 Then
            console_mode.Normal = Not console_mode.Normal
        End If
    End If
    
    Call CopiarLinea(X, Y)
End Function

Public Sub MouseMove(ByVal x As Integer, ByVal y As Integer)
    Dim NewSelectedValue As Integer
    If MouseY < console_height - 10 Then
        NewSelectedValue = (MouseY - 4) \ DIALOG_HEIGHT
    Else
        NewSelectedValue = -1
    End If
    
    If NewSelectedValue <> SelectedDialog Then
        SelectedDialog = NewSelectedValue
        Call RedrawTexture
    End If
End Sub

Private Sub CopiarLinea(ByVal X As Integer, ByVal Y As Integer)
    If x > 0 And y > 0 And x < MainViewWidth And y <= console_height Then
        Dim Index As Long
        Index = RENDER_DIALOGS - ((Y - 4) \ DIALOG_HEIGHT) + scroll_offset
        Dim line As String
        
       ' line = p_Dialogos((dialogIndex - Index) Mod p_CANTIDADDIALOGOS).Texto
        
        If InStr(1, line, "http://") > 0 Or InStr(1, line, "https://") > 0 Then

            Dim index2 As Integer
            Index = InStr(1, line, "http")
            index2 = InStr(Index, line, " ")
            If index2 = 0 Then index2 = Len(line)
            line = mid(line, Index, index2 - Index + 1)
            'Call ShellExecute(0, "open", url, 0, 0, 1)
        Else
        
        End If
        
        Call Clipboard.SetText(line)
    End If
End Sub

Public Sub RedrawTexture()
    Call DirectDevice.SetRenderTarget(ConsoleSurface, pZBuffer, 0)
    Call DirectDevice.SetTransform(D3DTS_PROJECTION, ConsoleProjection)
    Call Engine_BeginScene

    Dim i As Long
        
    
    Dim dy As Single
    Dim alphaOffset As Integer
    Dim Texto(2) As String
    alphaOffset = DIALOG_HEIGHT * (p_CANTIDADDIALOGOS - 2)
    
    Dim color(3) As RGBA
    
    Call RGBAList(color, 0, 0, 0, 125)
    
    Call Engine_Draw_Box(0, 0, MainViewWidth, console_height, RGBA_From_Comp(0, 0, 0, 255))
    
    
    If SelectedDialog >= 0 Then
        Call Engine_Draw_Box(3, 4 + SelectedDialog * DIALOG_HEIGHT, 733, 14, RGBA_From_Comp(125, 125, 125, 6))
    End If

    Dim Index As Integer
    For i = 1 To min(RENDER_DIALOGS, dialogCount) + IIf(offset > 0, 1, 0)
    
        Index = ((dialogIndex - i) - scroll_offset) Mod p_CANTIDADDIALOGOS
        If Index < 0 Then Index = Index + p_CANTIDADDIALOGOS
        If Index >= 0 Then
            If Trim$(p_Dialogos(Index).Texto) <> vbNullString Then
                If i = 1 And offset > 0 Then
                    Call RGBAList(chatColor, p_Dialogos(Index).red, p_Dialogos(Index).green, p_Dialogos(Index).blue, 255 - (offset * 255) / 14)
                ElseIf i = RENDER_DIALOGS + 1 Then
                    Call RGBAList(chatColor, p_Dialogos(Index).red, p_Dialogos(Index).green, p_Dialogos(Index).blue, (offset * 255) / 14)
                Else
                    Call RGBAList(chatColor, p_Dialogos(Index).red, p_Dialogos(Index).green, p_Dialogos(Index).blue, 255)
                End If
                
                If dialogCount < RENDER_DIALOGS Then
                    Call Engine_Text_Render_In_Console(p_Dialogos(Index).Texto, 5, (dialogCount - i) * DIALOG_HEIGHT - 6 + offset, chatColor, p_Dialogos(Index).status, 1, False, 0, 255, p_Dialogos(Index).priv)
                Else
                    Call Engine_Text_Render_In_Console(p_Dialogos(Index).Texto, 5, -i * DIALOG_HEIGHT + DIALOG_HEIGHT * RENDER_DIALOGS - 6 + offset, chatColor, p_Dialogos(Index).status, 1, False, 0, 255, p_Dialogos(Index).priv)
                End If
            End If
        End If
    Next i
    
    ' Termino de dibujar a la texture
    Call Engine_PauseScene
    
    ' Reseteo el target e inicio de nuevo la escena
    Call ResetRenderTarget
    Call DirectDevice.SetTransform(D3DTS_PROJECTION, Projection)
End Sub

Private Sub PrepareAndRedraw()
    Call Engine_PauseScene
    Call RedrawTexture
    Call Engine_ResumeScene
End Sub

Public Sub Draw()
    
    On Error GoTo Draw_Err
    
    If QueRender <> 0 Then Exit Sub
    
    Dim NeedsRedraw As Boolean
    
    If MouseY < console_height + 20 And MouseX < 740 Then
        If alphaConsole < consoleAlpha Then
            alphaConsole = alphaConsole + timerTicksPerFrame * 50
            If alphaConsole > consoleAlpha Then alphaConsole = consoleAlpha
            NeedsRedraw = True
        End If
    ElseIf alphaConsole <> 0 Then
        alphaConsole = 0
        NeedsRedraw = True
    End If
    
    Dim Color(3) As RGBA
    Call RGBAList(Color, 255, 255, 255, consoleAlpha)
    
    With SpriteBatch
        Call .SetTexture(ConsoleTexture)
        Call .SetAlpha(False)
        Call .Draw(1, 0, MainViewWidth, console_height - 1, Color)
    End With
    
    If consola_transparente Then
        If consoleAlpha < consoleAlpha_min Then
            consoleAlpha = consoleAlpha + timerTicksPerFrame * 30
        ElseIf consoleAlpha > consoleAlpha_min Then
            consoleAlpha = consoleAlpha - timerTicksPerFrame * 30
        End If
    Else
        If consoleAlpha < consoleAlpha_max Then
            consoleAlpha = consoleAlpha + timerTicksPerFrame * 30
        End If
    End If
    
    If consoleAlpha > consoleAlpha_max Then consoleAlpha = consoleAlpha_max
    If consoleAlpha <= consoleAlpha_min Then consoleAlpha = consoleAlpha_min

    If offset > 0 Then
        offset = offset - timerTicksPerFrame * 3
        If offset <= 0 Then
            offset = 0
        End If
        NeedsRedraw = True
    End If
    
    If NeedsRedraw Then
        Call PrepareAndRedraw
    End If

    Call RGBAList(Color(), 255, 255, 255, alphaConsole)
    'Box desplegable
    Call Engine_Draw_Box(700, console_height - 1, 56, 20, RGBA_From_Comp(0, 0, 0, alphaConsole / 2))
    'Icono desplegable
    Call Grh_Render_Advance(GRH_CONSOLE_DESPLEGAR, 711, console_height + 16, 30, 10, Color)
    
    If menu_consola_visible Then
        Dim MenuX As Integer, MenuY As Integer
        MenuX = 500
        MenuY = 0
        Dim title_color(3) As RGBA
        Dim text_color(3) As RGBA
        Call RGBAList(title_color, 54, 126, 127)
        Call RGBAList(text_color, 220, 220, 220)
        'Box principal
        Call Engine_Draw_Box(MenuX, MenuY, 236, 155, RGBA_From_Comp(15, 15, 15, 255))
        
        'Boxs internos
        Call Engine_Draw_Box_Border(MenuX + 6, MenuY + 28, 224, 55, RGBA_From_Comp(15, 15, 15, 255), RGBA_From_Comp(70, 70, 70, 255))
        Call Engine_Draw_Box_Border(MenuX + 6, MenuY + 94, 224, 40, RGBA_From_Comp(15, 15, 15, 255), RGBA_From_Comp(70, 70, 70, 255))
        
        'Título
        Call Engine_Text_Render("Configuracion de consola", MenuX + 47, MenuY + 8, title_color)
        
        'Texto
        Call Engine_Text_Render("Modo:", MenuX + 12, MenuY + 35, text_color)
        Call Engine_Text_Render("Global", MenuX + 50, MenuY + 35, text_color)
        Call Engine_Text_Render("Party", MenuX + 117, MenuY + 35, text_color)
        Call Engine_Text_Render("Clan", MenuX + 180, MenuY + 35, text_color)
        Call Engine_Text_Render("Private", MenuX + 12, MenuY + 60, text_color)
        Call Engine_Text_Render("Combat", MenuX + 84, MenuY + 60, text_color)
        Call Engine_Text_Render("Normal", MenuX + 158, MenuY + 60, text_color)
        Call Engine_Text_Render("Intensidad de color:", MenuX + 12, MenuY + 101, text_color)
        '32117-32118
        Dim grhOn As grh, grhOff As grh
        
        Call InitGrh(grhOn, 32118)
        Call InitGrh(grhOff, 32117)
        Dim color_grh(3) As RGBA
        
        If console_mode.Global Then
            Call Grh_Render(grhOn, MenuX + 85, MenuY + 20, text_color)
        Else
            Call Grh_Render(grhOff, MenuX + 85, MenuY + 20, text_color)
        End If
        
        If console_mode.Party Then
            Call Grh_Render(grhOn, MenuX + 148, MenuY + 20, text_color)
        Else
            Call Grh_Render(grhOff, MenuX + 148, MenuY + 20, text_color)
        End If
        
        If console_mode.Guild Then
            Call Grh_Render(grhOn, MenuX + 202, MenuY + 20, text_color)
        Else
            Call Grh_Render(grhOff, MenuX + 202, MenuY + 20, text_color)
        End If
        
        If console_mode.Private Then
            Call Grh_Render(grhOn, MenuX + 51, MenuY + 46, text_color)
        Else
            Call Grh_Render(grhOff, MenuX + 51, MenuY + 46, text_color)
        End If
        
        If console_mode.Combat Then
            Call Grh_Render(grhOn, MenuX + 125, MenuY + 46, text_color)
        Else
            Call Grh_Render(grhOff, MenuX + 125, MenuY + 46, text_color)
        End If
        
        If console_mode.Normal Then
            Call Grh_Render(grhOn, MenuX + 202, MenuY + 46, text_color)
        Else
            Call Grh_Render(grhOff, MenuX + 202, MenuY + 46, text_color)
        End If
                
        Call Engine_Draw_Box_Border(MenuX + 12, MenuY + 115, 212, 15, RGBA_From_Comp(15, 15, 15, 255), RGBA_From_Comp(70, 70, 70, 255))
        Call Engine_Draw_Box(consoleAlpha_min_pos, MenuY + 116, consoleAlpha_max_pos - consoleAlpha_min_pos, 14, RGBA_From_Comp(11, 120, 120, 255))
        Call Engine_Draw_Box(MenuX + 63, MenuY + 116, 2, 14, RGBA_From_Comp(220, 220, 220, 255))
        
        
        'Call Engine_Text_Render("Global", MenuX + 8, MenuY + 30, text_color)
    End If
    
    Exit Sub

Draw_Err:
    Call RegistrarError(err.Number, err.Description, "clsConsoleDlg.Draw", Erl)
    Resume Next
    
End Sub


''
' Splits dialogs into lines fitting properly the render area and inserts them

Public Sub PushBackText(ByVal s As String, ByVal red As Byte, ByVal green As Byte, ByVal blue As Byte, Optional ByVal status As Integer = -1, Optional ByVal priv As Integer = -1)
    
    On Error GoTo PushBackText_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/07/08
    'Splits dialogs into lines fitting properly the render area and inserts them
    '***************************************************
    Dim outputText() As String
    Dim i As Long
    Dim separatorPos As Long
    
    outputText = Prepare_Multiline_Text(s, 600)
    
    separatorPos = InStr(1, outputText(0), ">")
    If separatorPos > 0 Then
        outputText(0) = Chr(1) & Left(outputText(0), separatorPos) & Chr(1) & Right(outputText(0), Len(outputText(0)) - separatorPos)
    End If
    
    'Show message and continue with the rest
    For i = 0 To UBound(outputText)
        Call RealPushBackText(outputText(i), red, green, blue, status, priv)
    Next i
    
    Call RedrawTexture
        
    Exit Sub

PushBackText_Err:
    Call RegistrarError(err.Number, err.Description, "clsConsoleDlg.PushBackText", Erl)
    Resume Next
    
End Sub

Private Sub RealPushBackText(ByVal s As String, ByVal red As Byte, ByVal green As Byte, ByVal blue As Byte, Optional ByVal status As Integer = -1, Optional ByVal priv As Integer = -1)
    
    On Error GoTo RealPushBackText_Err

    Dim i As Integer

    p_Dialogos(dialogIndex Mod p_CANTIDADDIALOGOS).Texto = s
    p_Dialogos(dialogIndex Mod p_CANTIDADDIALOGOS).red = red
    p_Dialogos(dialogIndex Mod p_CANTIDADDIALOGOS).green = green
    p_Dialogos(dialogIndex Mod p_CANTIDADDIALOGOS).blue = blue
    p_Dialogos(dialogIndex Mod p_CANTIDADDIALOGOS).status = status
    p_Dialogos(dialogIndex Mod p_CANTIDADDIALOGOS).priv = priv

    'Select Case type_message
        
    '    Case 1 'Private
    '        If console_mode.Private Then
                'p_Dialogos_filtered
    '        End If
    '    Case 2 'Guild
        
    '    Case 3 'Normal
        
    '    Case 4 'Combat
        
    '    Case 5 'Party
        
    '    Case 6 'Global
    
    '    Case Else
    
    'End Select
    
    
    offset = IIf(dialogCount < RENDER_DIALOGS, 0, DIALOG_HEIGHT)
    dialogIndex = dialogIndex + 1
    dialogIndex = dialogIndex Mod p_CANTIDADDIALOGOS
    dialogCount = dialogCount + 1
    If MouseY > console_height Then
        scroll_offset = 0
    Else
        If scroll_offset > 0 Then
            scroll_offset = scroll_offset + 1
            offset = 0
        End If
    End If
    
    Exit Sub

RealPushBackText_Err:
    Call RegistrarError(err.Number, err.Description, "clsConsoleDlg.RealPushBackText", Erl)
    Resume Next
    
End Sub

''
' Retrieves the font being used to render dialogs.
'
' @return   The font being used to render dialogs.

Public Property Get font() As StdFont
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 04/27/06
    'Retrieves the font being used to render dialogs.
    '**************************************************************
    
    On Error GoTo font_Err
    
    font = usedFont

    
    Exit Property

font_Err:
    Call RegistrarError(err.Number, err.Description, "clsConsoleDlg.font", Erl)
    Resume Next
    
End Property

''
' Sets the font being used to render dialogs.
'
' @param    nValue  The new font to be used to render dialogs.

Public Property Let font(ByRef nValue As StdFont)
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 04/27/06
    'Sets the font being used to render dialogs.
    '**************************************************************
    
    On Error GoTo font_Err
    
    Set usedFont = nValue

    
    Exit Property

font_Err:
    Call RegistrarError(err.Number, err.Description, "clsConsoleDlg.font", Erl)
    Resume Next
End Property


' Limpia la consola
Public Sub Clear()
    dialogIndex = 0
    dialogCount = 0
End Sub
