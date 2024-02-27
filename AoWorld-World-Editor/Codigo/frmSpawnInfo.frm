VERSION 5.00
Begin VB.Form frmSpawnInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Editar Npc Spawn"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6750
   Icon            =   "frmSpawnInfo.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   463
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox tBuscar 
      Height          =   285
      Left            =   2880
      TabIndex        =   24
      Top             =   960
      Width           =   3735
   End
   Begin VB.ListBox lstBuscar 
      Height          =   1425
      Left            =   2880
      TabIndex        =   22
      Top             =   1320
      Width           =   3735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Eliminar selec"
      Height          =   375
      Left            =   5040
      TabIndex        =   21
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Agregar nuevo"
      Height          =   375
      Left            =   5040
      TabIndex        =   20
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox tCant 
      Height          =   375
      Left            =   5040
      TabIndex        =   18
      Text            =   "0"
      Top             =   4080
      Width           =   1575
   End
   Begin VB.TextBox tNpc 
      Height          =   375
      Left            =   5040
      TabIndex        =   16
      Text            =   "0"
      Top             =   3360
      Width           =   1575
   End
   Begin VB.ListBox lstNpcs 
      Height          =   2985
      Left            =   120
      TabIndex        =   15
      Top             =   3000
      Width           =   4695
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Guardar Cambios"
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton btnCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5040
      TabIndex        =   13
      Top             =   6480
      Width           =   1455
   End
   Begin VB.TextBox tZX1 
      Height          =   285
      Left            =   600
      TabIndex        =   6
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox tZY1 
      Height          =   285
      Left            =   600
      TabIndex        =   5
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox tZX2 
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox tZY2 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox txtIrZona 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Text            =   "1"
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmdIraZona 
      Caption         =   "Ir a Spawn"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtMap 
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Buscar npc:"
      Height          =   255
      Left            =   2880
      TabIndex        =   23
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Cantidad:"
      Height          =   255
      Left            =   5040
      TabIndex        =   19
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "NPC Index:"
      Height          =   255
      Left            =   5040
      TabIndex        =   17
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   16
      X2              =   432
      Y1              =   424
      Y2              =   424
   End
   Begin VB.Label tNumZona 
      BackColor       =   &H80000004&
      Caption         =   "Spawn Nº:"
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
      TabIndex        =   12
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "X1:"
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   11
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Y1:"
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   10
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "X2:"
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   6
      Left            =   1680
      TabIndex        =   9
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Y2:"
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   5
      Left            =   1680
      TabIndex        =   8
      Top             =   2400
      Width           =   375
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      X1              =   8
      X2              =   184
      Y1              =   80
      Y2              =   80
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      X1              =   184
      X2              =   184
      Y1              =   80
      Y2              =   184
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      X1              =   184
      X2              =   8
      Y1              =   184
      Y2              =   184
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      X1              =   8
      X2              =   8
      Y1              =   80
      Y2              =   184
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Map:"
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   375
   End
End
Attribute VB_Name = "frmSpawnInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nSpawn As t_NpcSpawn
Dim OpenedSpawn As Integer

Dim BorrarNpcs As Boolean
Dim loadingText As Boolean
Private Sub btnCancelar_Click()
Unload Me
End Sub


Public Sub OpenSpawn(id As Integer)
    Dim i As Integer
    If id > 0 Then
        nSpawn = NpcSpawn(id)
    Else
        Dim newSpawn As t_NpcSpawn
        nSpawn = newSpawn
        nSpawn.Map = UserMap
        If RectanguloX > RectanguloX2 Then
            nSpawn.X2 = RectanguloX
            nSpawn.X = RectanguloX2
        Else
            nSpawn.X = RectanguloX
            nSpawn.X2 = RectanguloX2
        End If
        
        If RectanguloY > RectanguloY2 Then
            nSpawn.Y2 = RectanguloY
            nSpawn.Y = RectanguloY2
        Else
            nSpawn.Y = RectanguloY
            nSpawn.Y2 = RectanguloY2
        End If
        
        If nSpawn.X > 0 Then
            Dim X As Integer
            Dim Y As Integer
            For Y = nSpawn.Y To nSpawn.Y2
                For X = nSpawn.X To nSpawn.X2
                    If MapData(X, Y).NpcIndex > 0 Then
                        If NpcData(MapData(X, Y).NpcIndex).Hostile Then
                            AddNpc (MapData(X, Y).NpcIndex)
                        End If
                    End If
                Next X
            Next Y
            BorrarNpcs = nSpawn.CantNpcs > 0
        End If
        
    End If
    
    OpenedSpawn = id
    
    With nSpawn
        txtMap.Text = .Map
        tZX1.Text = .X
        tZY1.Text = .Y
        tZX2.Text = .X2
        tZY2.Text = .Y2
        
        lstNpcs.Clear
        For i = 1 To .CantNpcs
            lstNpcs.AddItem .NPCs(i).Cantidad & " x " & .NPCs(i).NpcIndex & " | " & NpcData(.NPCs(i).NpcIndex).name
        Next i

    End With
    If id > 0 Then
        tNumZona.Caption = "Spawn N°: " & id
    Else
        tNumZona.Caption = "Spawn Nuevo"
    End If
    txtIrZona.Text = id
End Sub

Private Sub cmdIraZona_Click()
    Dim i As Integer
    Dim e As Integer

    If txtIrZona.Text > NumSpawns Then
        Exit Sub
    Else
        i = Val(txtIrZona.Text)
    End If

    If i < 0 Then i = 0

    OpenSpawn (i)
End Sub

Private Sub UpdateNpc()
If loadingText Or lstNpcs.ListIndex = -1 Then Exit Sub
Dim i As Integer
i = lstNpcs.ListIndex + 1
nSpawn.NPCs(i).NpcIndex = tNpc.Text
nSpawn.NPCs(i).Cantidad = tCant.Text
lstNpcs.List(i - 1) = nSpawn.NPCs(i).Cantidad & " x " & nSpawn.NPCs(i).NpcIndex & " | " & NpcData(nSpawn.NPCs(i).NpcIndex).name
End Sub

Private Sub Command2_Click()
nSpawn.CantNpcs = nSpawn.CantNpcs + 1
Dim i As Integer
i = nSpawn.CantNpcs
ReDim Preserve nSpawn.NPCs(1 To i)
nSpawn.NPCs(i).NpcIndex = tNpc.Text
nSpawn.NPCs(i).Cantidad = tCant.Text

lstNpcs.AddItem nSpawn.NPCs(i).Cantidad & " x " & nSpawn.NPCs(i).NpcIndex & " | " & NpcData(nSpawn.NPCs(i).NpcIndex).name
End Sub

Sub AddNpc(Index As Integer)
Dim i As Integer
For i = 1 To nSpawn.CantNpcs
    If nSpawn.NPCs(i).NpcIndex = Index Then
        nSpawn.NPCs(i).Cantidad = nSpawn.NPCs(i).Cantidad + 1
        Exit For
    End If
Next i
If i > nSpawn.CantNpcs Then
    nSpawn.CantNpcs = i
    ReDim Preserve nSpawn.NPCs(1 To i)
    nSpawn.NPCs(i).NpcIndex = Index
    nSpawn.NPCs(i).Cantidad = 1
End If

End Sub

Private Sub Command26_Click()

    With nSpawn
        .Map = Val(txtMap.Text)
        .X = Val(tZX1.Text)
        .Y = Val(tZY1.Text)
        .X2 = Val(tZX2.Text)
        .Y2 = Val(tZY2.Text)

    

        OpenedSpawn = SaveSpawn(OpenedSpawn, nSpawn)
        Call FrmMain.DibujarSpawns
        
        If BorrarNpcs Then
            Dim X As Integer
            Dim Y As Integer
            For Y = .Y To .Y2
                For X = .X To .X2
                    If MapData(X, Y).NpcIndex > 0 Then
                        If NpcData(MapData(X, Y).NpcIndex).Hostile Then
                            EraseChar (MapData(X, Y).CharIndex)
                            MapData(X, Y).NpcIndex = 0
                        End If
                    End If
                Next X
            Next Y
        End If

    End With
    Unload Me
End Sub

Private Sub Command3_Click()
Dim i As Integer
Dim X As Integer
i = lstNpcs.ListIndex + 1

If i < nSpawn.CantNpcs Then
    For X = i To nSpawn.CantNpcs - 1
        nSpawn.NPCs(X) = nSpawn.NPCs(X + 1)
    Next X
End If
nSpawn.CantNpcs = nSpawn.CantNpcs - 1
ReDim Preserve nSpawn.NPCs(1 To nSpawn.CantNpcs)

lstNpcs.RemoveItem (lstNpcs.ListIndex)
End Sub

Private Sub lstBuscar_Click()
tNpc.Text = Left$(lstBuscar.Text, InStr(1, lstBuscar.Text, " - ") - 1)
End Sub

Private Sub lstNpcs_Click()
loadingText = True
tNpc.Text = nSpawn.NPCs(lstNpcs.ListIndex + 1).NpcIndex
tCant.Text = nSpawn.NPCs(lstNpcs.ListIndex + 1).Cantidad
loadingText = False
End Sub

Private Sub tBuscar_Change()
Dim i As Integer
lstBuscar.Clear
For i = 1 To UBound(NpcData)
    If NpcData(i).Hostile = 1 Then
        If InStr(1, UCase$(NpcData(i).name), UCase$(tBuscar.Text), vbTextCompare) > 0 Then
            lstBuscar.AddItem (i & " - " & NpcData(i).name)
        End If
    End If
Next i
End Sub

Private Sub tCant_Change()
UpdateNpc
End Sub

Private Sub tNpc_Change()
UpdateNpc
End Sub
