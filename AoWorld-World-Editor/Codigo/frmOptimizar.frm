VERSION 5.00
Begin VB.Form frmOptimizar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Optimizar Mapa"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3600
   Icon            =   "frmOptimizar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   3600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkBloquearArbolesEtc 
      Caption         =   "Bloquear Arboles, Carteles, Foros y Yacimientos"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin VB.CheckBox chkMapearArbolesEtc 
      Caption         =   "Mapear Arboles, Carteles, Foros y Yacimientos que no esten en la 3ra Capa"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   3375
   End
   Begin VB.CheckBox chkQuitarTodoBordes 
      Caption         =   "Quitar NPCs, Objetos y Translados en los Bordes Exteriores"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   3375
   End
   Begin VB.CheckBox chkQuitarTrigTrans 
      Caption         =   "Quitar Trigger's en Translados"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3375
   End
   Begin VB.CheckBox chkQuitarTrigBloq 
      Caption         =   "Quitar Trigger's Bloqueados"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3375
   End
   Begin VB.CheckBox chkQuitarTrans 
      Caption         =   "Quitar Translados Bloqueados"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin WorldEditor.lvButtons_H cOptimizar 
      Default         =   -1  'True
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1296
      Caption         =   "&Optimizar"
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
      cBack           =   12648384
   End
   Begin WorldEditor.lvButtons_H cCancelar 
      Height          =   735
      Left            =   1920
      TabIndex        =   6
      Top             =   2640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      Caption         =   "&Cancelar"
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
      Mode            =   1
      Value           =   0   'False
      cBack           =   -2147483633
   End
End
Attribute VB_Name = "frmOptimizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Optimizar()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 16/10/06
    '*************************************************
    
    On Error GoTo Optimizar_Err
    
    Dim Y As Integer
    Dim X As Integer

    If Not MapaCargado Then
        Exit Sub

    End If

    ' Quita Translados Bloqueados
    ' Quita Trigger's Bloqueados
    ' Quita Trigger's en Translados
    ' Quita NPCs, Objetos y Translados en los Bordes Exteriores
    ' Mapea Arboles, Carteles, Foros y Yacimientos que no esten en la 3ra Capa

     Call modEdicion.Deshacer_Add(x, y, 1, 1) ' Hago deshacer

    For Y = 1 To MapSize.Height
        For X = 1 To MapSize.Width

            ' ** Quitar NPCs, Objetos y Translados en los Bordes Exteriores
            If (X < 12 Or X > 88 Or Y < 10 Or Y > 91) And chkQuitarTodoBordes.Value = 1 Then

                'Quitar NPCs
                If MapData(X, Y).NPCIndex > 0 Then
                    EraseChar MapData(X, Y).CharIndex
                    MapData(X, Y).NPCIndex = 0

                End If

                ' Quitar Objetos
                MapData(X, Y).OBJInfo.ObjIndex = 0
                MapData(X, Y).OBJInfo.Amount = 0
                MapData(X, Y).ObjGrh.grhindex = 0
                ' Quitar Translados
                MapData(X, Y).TileExit.Map = 0
                MapData(X, Y).TileExit.X = 0
                MapData(X, Y).TileExit.Y = 0
                ' Quitar Triggers
                MapData(X, Y).Trigger = 0

            End If

            ' ** Quitar Translados y Triggers en Bloqueo
            If MapData(X, Y).Blocked = &HF Then
                If MapData(X, Y).TileExit.Map <> 0 And chkQuitarTrans.Value = 1 Then ' Quita Translado Bloqueado
                    MapData(X, Y).TileExit.Map = 0
                    MapData(X, Y).TileExit.Y = 0
                    MapData(X, Y).TileExit.X = 0
                ElseIf MapData(X, Y).Trigger > 0 And chkQuitarTrigBloq.Value = 1 Then ' Quita Trigger Bloqueado
                    MapData(X, Y).Trigger = 0

                End If

            End If

            ' ** Quitar Triggers en Translado
            If MapData(X, Y).TileExit.Map <> 0 And chkQuitarTrigTrans.Value = 1 Then
                If MapData(X, Y).Trigger > 0 Then ' Quita Trigger en Translado
                    MapData(X, Y).Trigger = 0

                End If

            End If

            ' ** Mapea Arboles, Carteles, Foros y Yacimientos que no esten en la 3ra Capa
            If MapData(X, Y).OBJInfo.ObjIndex > 0 And (chkMapearArbolesEtc.Value = 1 Or chkBloquearArbolesEtc.Value = 1) Then

                Select Case ObjData(MapData(X, Y).OBJInfo.ObjIndex).ObjType

                    Case 4, 10, 22  ' Arboles,  Foros, Yacimientos

                        If MapData(X, Y).Graphic(3).grhindex <> MapData(X, Y).ObjGrh.grhindex And chkMapearArbolesEtc.Value = 1 Then
                            MapData(X, Y).Graphic(3) = MapData(X, Y).ObjGrh
                        End If
                        If chkBloquearArbolesEtc.Value = 1 And MapData(X, Y).Blocked = 0 Then
                        MapData(X, Y).Blocked = 15
                        End If

                    Case 8 '  Carteles

                        If MapData(X, Y).Graphic(3).grhindex <> MapData(X, Y).ObjGrh.grhindex And chkMapearArbolesEtc.Value = 1 Then
                            MapData(X, Y).Graphic(3) = MapData(X, Y).ObjGrh
                        End If
                        If chkBloquearArbolesEtc.Value = 1 Then

                            
                            If (MapData(X, Y).Blocked <> 1) Then
                                MapData(X, Y).Blocked = 1
                                MapInfo.Changed = 1
                            End If
                            
                            
                           If (MapData(X, Y + 1).Blocked <> 4) Then
                                MapData(X, Y + 1).Blocked = 4
                                MapInfo.Changed = 1
                            End If
                        
                        
                        End If
                End Select


            End If

            ' ** Mapea Arboles, Carteles, Foros y Yacimientos que no esten en la 3ra Capa
        Next X
    Next Y

    'Set changed flag
    

    
    Exit Sub

Optimizar_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmOptimizar.Optimizar", Erl)
    Resume Next
    
End Sub

Private Sub cCancelar_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 22/09/06
    '*************************************************
    
    On Error GoTo cCancelar_Click_Err
    
    Unload Me

    
    Exit Sub

cCancelar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmOptimizar.cCancelar_Click", Erl)
    Resume Next
    
End Sub

Public Sub cOptimizar_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 22/09/06
    '*************************************************
    
    On Error GoTo cOptimizar_Click_Err
    
    Call Optimizar
    MapInfo.Changed = 1

    
    Exit Sub

cOptimizar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmOptimizar.cOptimizar_Click", Erl)
    Resume Next
    
End Sub
