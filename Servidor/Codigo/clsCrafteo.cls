VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCrafteo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Items() As Integer

Public tipo As Byte

Public Precio As Long

Public Probabilidad As Byte

Public Resultado As Integer

Public Sub SetItems(data() As Integer)
    m_Items = data
End Sub

Public Property Get Tamaño() As Integer
    Tamaño = UBound(m_Items)
End Property

Public Property Get Items(ByVal index As Integer) As Integer
    Items = m_Items(index)
End Property
