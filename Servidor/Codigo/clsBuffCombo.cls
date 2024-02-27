VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBuffCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public maxCombos As Byte
Public magia As Boolean
Public golpe As Boolean
Private buff() As Byte


Public Sub setBuffs(ByVal index As Byte, ByVal value As Byte)
    ReDim Preserve buff(1 To index) As Byte
    buff(index) = value
End Sub

Public Function getBuffs(ByVal index As Byte) As Byte
    getBuffs = buff(index)
End Function

