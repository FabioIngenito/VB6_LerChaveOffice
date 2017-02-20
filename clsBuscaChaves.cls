VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBuscaChaves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim lngReturn           As Long
Dim strReturnWinReg     As String

Public Function BuscaVersao(strSoftware As String) As String

    ' Ler chave do Software/Versão
    lngReturn = objWinReg.QueryValueEx(HKEY_CLASSES_ROOT, strSoftware & ".Application\CurVer", "", 0, strReturnWinReg)

    If lngReturn Then
        BuscaVersao = strReturnWinReg
    Else
        'Não possui o software Instalado
        BuscaVersao = "Não foi encontrada a chave do Aplicativo " & strSoftware & " neste equipamento."
    End If
    
End Function

'-------------------------------------------------------------------
' - API VIEWER -
'-------------------------------------------------------------------
' C:\Program Files (x86)\Microsoft Visual Studio\COMMON\Tools\Winapi\WIN32API.TXT
'-------------------------------------------------------------------
' API Type: Declares / Avaliable itens: "RegQueryValueEx"
'-------------------------------------------------------------------
'Public Declare Function RegQueryValueEx Lib "advapi32.dll"
'Alias "RegQueryValueExA"
'(ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long,
' lpType As Long, lpData As Any, lpcbData As Long)
'As Long

' Note that if you declare the lpData parameter as String, you must pass it By Value.
