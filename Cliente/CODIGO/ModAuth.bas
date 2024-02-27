Attribute VB_Name = "ModAuth"
Option Explicit

Public Enum e_state
    Idle = 0
    RequestAccountLogin
    AccountLogged
    RequestCharList
    RequestLogout
    RequestSignUp
    RequestValidateAccount
    RequestForgotPassword
    RequestResetPassword
    RequestDeleteChar
    ConfirmDeleteChar
    RequestVerificationCode
End Enum

Public Enum e_operation
    Authenticate = 0
    SignUp
    ValidateAccount
    ForgotPassword
    ResetPassword
    deletechar
    ConfirmDeleteChar
    RequestVerificationCode
End Enum


Public SessionOpened As Boolean

Public Auth_state As e_state
Public LoginOperation As e_operation
Public public_key() As Byte
Public encrypted_session_token As String
Public delete_char_validate_code As String



Public Sub DebugPrint(ByVal str As String, Optional ByVal int1 As Integer = 0, Optional ByVal int2 As Integer = 0, Optional ByVal int3 As Integer = 0, Optional ByVal asd As Boolean = False)

    Debug.Print (str)
    
End Sub

Function FileToString(strFileName As String) As String
  Open strFileName For Input As #1
    FileToString = StrConv(InputB(LOF(1), 1), vbUnicode)
  Close #1
End Function
