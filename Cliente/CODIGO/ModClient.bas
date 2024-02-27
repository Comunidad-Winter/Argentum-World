Attribute VB_Name = "ModClient"

'RevolucionAo 1.0
'Pablo Mercavides
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_EXSTYLE As Long = (-20)

Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const WS_EX_TRANSPARENT As Long = &H20&

Private Const WM_NCLBUTTONDOWN = &HA1

Private Const HTCAPTION = 2

Private Const WS_EX_LAYERED = &H80000

Private Const LWA_ALPHA = &H2&

Public Sub Make_Transparent_Richtext(ByVal hwnd As Long)
    'If Win2kXP Then
    
    On Error GoTo Make_Transparent_Richtext_Err
    
    Call SetWindowLong(hwnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)

    
    Exit Sub

Make_Transparent_Richtext_Err:
    Call RegistrarError(Err.number, Err.Description, "ModClient.Make_Transparent_Richtext", Erl)
    Resume Next
    
End Sub
