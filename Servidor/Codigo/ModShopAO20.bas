Attribute VB_Name = "ModShopAO"
'********************* COPYRIGHT NOTICE*********************
' Copyright (c) 2021-22 Martin Trionfetti, Pablo Marquez
' www.argentumunited.com.ar
' All rights reserved.
' Refer to licence for conditions of use.
' This copyright notice must always be left intact.
'****************** END OF COPYRIGHT NOTICE*****************
'
Option Explicit


Public Sub init_transaction(ByVal obj_num As Long, ByVal UserIndex As Integer)
    
    Dim obj As t_ObjData
    
    obj.ObjNum = obj_num
    With UserList(UserIndex)
        
        'Me fijo si es un item de shop
        If Not is_purchaseable_item(obj) Then
            Call WriteConsoleMsg(UserIndex, "Error al realizar la transacci�n", e_FontTypeNames.FONTTYPE_INFO)
            Call LogShopErrors("El usuario " & .Name & " intent� comprar un objeto que no es de shop (REVISAR) | " & obj.Name)
            Exit Sub
        End If
        
        If obj.Valor > .Stats.Creditos Then
            Call WriteConsoleMsg(UserIndex, "Error al realizar la transacci�n.", e_FontTypeNames.FONTTYPE_INFO)
            Call LogShopErrors("El usuario " & .Name & " intent� editar el valor del objeto (REVISAR) | " & obj.Name)
            Exit Sub
        End If
        
        'Me fijo si tiene espacio en el inventario
        Dim objInventario As t_Obj
        
        objInventario.Amount = 1
        objInventario.ObjIndex = obj.ObjNum
        
        If Not MeterItemEnInventario(UserIndex, objInventario) Then
            Call WriteConsoleMsg(UserIndex, "Asegurate de tener espacio suficiente en tu inventario.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        Else
            'Descuento los cr�ditos
            .Stats.Creditos = .Stats.Creditos - obj.Valor
            
            'Genero un log de los cr�ditos que gast� y cuantos le quedan luego de la transacci�n.
            Call LogShopTransactions(.Name & " | Compr� -> " & ObjData(obj.ObjNum).Name & " | Valor -> " & obj.Valor)
            Call Execute("update user set credits = ? where id = ?;", .Stats.Creditos, .ID)
            Call writeUpdateShopClienteCredits(UserIndex)
        End If
                
    End With
        
End Sub

Private Function is_purchaseable_item(ByRef obj As t_ObjData) As Boolean
    Dim i As Long
    
    For i = 1 To UBound(ObjShop)
        If ObjShop(i).ObjNum = obj.ObjNum Then
            'Si es un item de shop, aparte le agrego el valor (por ref)
            obj.Valor = ObjShop(i).Valor
            is_purchaseable_item = True
            Exit Function
        End If
    Next i
    
    is_purchaseable_item = False
    
End Function
