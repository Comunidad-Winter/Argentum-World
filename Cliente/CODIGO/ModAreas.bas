Attribute VB_Name = "ModAreas"
Option Explicit

'LAS GUARDAMOS PARA PROCESAR LOS MPs y sabes si borrar personajes
Public MinLimiteX      As Integer

Public MaxLimiteX      As Integer

Public MinLimiteY      As Integer

Public MaxLimiteY      As Integer

Public Const AREA_DIM As Byte = 13
Public Const FULL_AREA_DIM As Byte = AREA_DIM * 5

Public Const MAP_MAX_X As Integer = 65
Public Const MAP_MAX_Y As Integer = 65

Public Const NUM_MAPAS As Integer = 100

Public LBoundRoof As Integer, UBoundRoof As Integer


Public Sub CambioDeArea(ByVal x As Integer, ByVal y As Integer, Heading As Byte)
    
    On Error GoTo CambioDeArea_Err
    

    Dim loopX As Long, loopY As Long
    
    MinLimiteX = (x \ AREA_DIM) * AREA_DIM + 1
    MinLimiteY = (y \ AREA_DIM) * AREA_DIM + 1
    
    If Heading = E_Heading.south Then
        
        MaxLimiteX = MinLimiteX + AREA_DIM * 2 - 1
        MinLimiteX = MinLimiteX - AREA_DIM
        
        
        MinLimiteY = MinLimiteY - AREA_DIM * 2 - 1
        MaxLimiteY = MinLimiteY + AREA_DIM

    ElseIf Heading = E_Heading.NORTH Then


        MaxLimiteX = MinLimiteX + AREA_DIM * 2 - 1
        MinLimiteX = MinLimiteX - AREA_DIM
        
        
        MinLimiteY = MinLimiteY + AREA_DIM * 2 - 1
        MaxLimiteY = MinLimiteY + AREA_DIM

    ElseIf Heading = E_Heading.EAST Then

        MaxLimiteY = MinLimiteY + AREA_DIM * 2 - 1
        MinLimiteY = MinLimiteY - AREA_DIM
        
        
        MinLimiteX = MinLimiteX - AREA_DIM * 2 - 1
        MaxLimiteX = MinLimiteX + AREA_DIM

    ElseIf Heading = E_Heading.WEST Then


        MaxLimiteY = MinLimiteY + AREA_DIM * 2 - 1
        MinLimiteY = MinLimiteY - AREA_DIM
        
        
        MinLimiteX = MinLimiteX + AREA_DIM * 2 - 1
        MaxLimiteX = MinLimiteX + AREA_DIM

   
    ElseIf Heading = 255 Or Heading = 5 Then
        'Esto pasa por cuando cambiamos de mapa o logeamos...
        
        MaxLimiteX = MinLimiteX + AREA_DIM * 2 - 1
        MaxLimiteY = MinLimiteY + AREA_DIM * 2 - 1
        MinLimiteY = MinLimiteY - AREA_DIM
        MinLimiteX = MinLimiteX - AREA_DIM
    End If
    
       
    For loopX = MinLimiteX To MaxLimiteX
        For loopY = MinLimiteY To MaxLimiteY
            If loopX > 0 And loopY > 0 And loopX <= MapSize.Width And loopY <= MapSize.Height Then
            
                With MapData(rrX(loopX), rrY(loopY))
                
                    If .charindex > 0 Then
                        If .charindex <> UserCharIndex Then
                            Call EraseChar(.charindex)
    
                        End If
                    End If
                    
                    'Erase OBJs
                    If Not EsObjetoFijo(loopX, loopY) Then
                        .ObjGrh.GrhIndex = 0
                        .OBJInfo.ObjIndex = 0
                    End If

                End With
            End If
        
        Next loopY
    Next loopX
    
    Call RefreshAllChars

    
    Exit Sub

CambioDeArea_Err:
    Call RegistrarError(err.Number, err.Description, "ModAreas.CambioDeArea", Erl)
    Resume Next
    
End Sub
