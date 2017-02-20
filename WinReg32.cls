VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WinReg32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
' Nome            : WinReg32.cls
' Descrição       : Classe utilitária de manipulação do banco de dados
'                   de registro do Windows
' Autor           :
' Início          : 21/03/1999
' Ultima Alteração: 10/08/2001
' Dependencias    :
'--------------------------------------------------------------------------------

Option Explicit

' Fechar uma chave previamente aberta no banco de dados
' de registro do Windows.
Private Declare Function RegCloseKey Lib "advapi32.dll" _
        (ByVal hKey As Long) As Long

' Criar uma chave no banco de dados de registro do Windows. Se a
' chave já existir, o método irá abri-la.
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias _
        "RegCreateKeyExA" (ByVal hKey As Long, _
        ByVal lpSubKey As String, ByVal Reserved As Long, _
        ByVal lpClass As String, ByVal dwOptions As Long, _
        ByVal samDesired As Long, _
        lpSecurityAttributes As SECURITY_ATTRIBUTES, _
        phkResult As Long, lpdwDisposition As Long) As Long

' Conectar ao banco de dados de registro do Windows de
' uma máquina remota.
Private Declare Function RegConnectRegistry Lib "advapi32.dll" Alias _
        "RegConnectRegistryA" (ByVal lpMachineName As String, _
        ByVal hKey As Long, phkResult As Long) As Long

' Deletar uma chave do banco de dados de registro do Windows.
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias _
        "RegDeleteKeyA" (ByVal hKey As Long, _
        ByVal lpSubKey As String) As Long

' Remover uma valor do banco de dados de registro do Windows.
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias _
        "RegDeleteValueA" (ByVal hKey As Long, _
        ByVal lpValueName As String) As Long

' Abrir uma chave no banco de dados de registro do Windows.
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
        "RegOpenKeyExA" (ByVal hKey As Long, _
        ByVal lpSubKey As String, ByVal ulOptions As Long, _
        ByVal samDesired As Long, phkResult As Long) As Long

' Ler um valor no banco de dados de registro do Windows.
' Note que se você declarar o parâmetro lpData como String,
' você precisa passa-lo por valor.
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias _
        "RegQueryValueExA" (ByVal hKey As Long, _
        ByVal lpValueName As String, ByVal lpReserved As Long, _
        lpType As Long, ByVal lpData As String, lpcbData As Long) As Long

' Gravar um valor no banco de dados de registro do Windows.
' Note que se você declarar o parâmetro lpData como String,
' você precisa passa-lo por valor.
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias _
        "RegSetValueExA" (ByVal hKey As Long, _
        ByVal lpValueName As String, ByVal Reserved As Long, _
        ByVal dwType As Long, ByVal lpData As String, _
        ByVal cbData As Long) As Long

' Prover permissões para as chaves
' do banco de dados de registro do Windows.
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const SYNCHRONIZE = &H100000
Private Const KEY_NOTIFY = &H10
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))

' Prover informação se um chave foi criada ou aberta pela
' chamada a API RegCreateKeyEx.
Private Const REG_CREATED_NEW_KEY = &H1     ' New Registry Key created
Private Const REG_OPENED_EXISTING_KEY = &H2 ' Existing Key opened

' Tipo definido para atributos de segurança.
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

' Prover data types dos valores do banco de dados
' de registro do Windows.
Public Enum REG_TYPE_ENUM
    REG_NONE = 0                ' No value type
    REG_SZ = 1                  ' Unicode nul terminated string
    REG_EXPAND_SZ = 2           ' Unicode nul terminated string
    REG_BINARY = 3              ' Free form binary
    REG_DWORD = 4               ' 32-bit number
    REG_DWORD_BIG_ENDIAN = 5    ' 32-bit number
    REG_LINK = 6                ' Symbolic Link (unicode)
    REG_MULTI_SZ = 7            ' Multiple Unicode strings
End Enum

' Prover opções de criação de chaves no
' banco de dados de registro do Windows.
Public Enum REG_OPTION_ENUM
    REG_OPTION_NON_VOLATILE = 0 ' Key is preserved when system is rebooted
    REG_OPTION_VOLATILE = 1     ' Key is not preserved when system is rebooted
End Enum

' Prover lista de valores indentificadores das chaves
' do banco de dados de registro d Windows.
Public Enum HKEY_ENUM
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum

Public Function CreateKeyEx(ByVal hKey As HKEY_ENUM, ByVal lpSubKey As String, _
                            ByVal lpClass As String, ByVal dwOptions As REG_OPTION_ENUM, _
                            phkResult As Long, lpdwDisposition As Long, _
                            lpDispDescription As String, _
                            Optional lpMachineName As Variant) As Long

' Área para declaração de variáveis e objetos de
' escopo local.
Dim lngReturn As Long
Dim lpSecurityAttributes As SECURITY_ATTRIBUTES
    
    On Error GoTo ErrorHandler
    
    ' Avaliar se o parâmetro lpMachineName foi informado.
    If Not IsMissing(lpMachineName) Then
        
        ' Ativar API para conectar-se a um banco de dados de
        ' registro do Windows de uma máquina remota.
        lngReturn = RegConnectRegistry(CStr(lpMachineName), _
                                       hKey, phkResult)
        
        ' Avaliar se a API retornou sucesso.
        If lngReturn <> 0 Then
            GoTo ErrorHandler
        End If
        
        ' Atribuir o handler para a chave do registro remoto
        ' ao parâmetro hKey.
        hKey = phkResult
        
    End If
    
    ' Ativar API para ler o valor.
    lngReturn = RegCreateKeyEx(hKey, lpSubKey, 0&, _
                               lpClass, dwOptions, KEY_ALL_ACCESS, _
                               lpSecurityAttributes, phkResult, lpdwDisposition)
    
    ' Avaliar se a API retornou sucesso.
    If lngReturn <> 0 Then
        GoTo ErrorHandler
    End If
    
    ' Avalair se a chave foi criada ou aberta.
    If lpdwDisposition = REG_CREATED_NEW_KEY Then
        lpDispDescription = "The key did not exist and was created."
    ElseIf lpdwDisposition = REG_OPENED_EXISTING_KEY Then
        lpDispDescription = "The key existed and was simply opened without being changed."
    End If
    
    ' Fechar a chave no banco de dados de registro
    ' do Windows.
    RegCloseKey (hKey)
    
    ' Retornar sucesso ao método.
    CreateKeyEx = -1
    
    Exit Function
    
ErrorHandler:
    
    ' Fechar a chave no banco de dados de registro
    ' do Windows.
    RegCloseKey (phkResult)
    
    ' Retornar código de erro ao método.
    CreateKeyEx = 0
    
End Function

Public Function DeleteKey(ByVal hKey As HKEY_ENUM, ByVal lpSubKey As String, _
                          Optional lpMachineName As Variant) As Long

' Área para declaração de variáveis e objetos de
' escopo local.
Dim lngReturn As Long
Dim phkResult As Long
    
    On Error GoTo ErrorHandler
    
    ' Avaliar se o parâmetro lpMachineName foi informado.
    If Not IsMissing(lpMachineName) Then
        
        ' Ativar API para conectar-se a um banco de dados de
        ' registro do Windows de uma máquina remota.
        lngReturn = RegConnectRegistry(CStr(lpMachineName), _
                                       hKey, phkResult)
        
        ' Avaliar se a API retornou sucesso.
        If lngReturn <> 0 Then
            GoTo ErrorHandler
        End If
        
        ' Atribuir o handler para a chave do registro remoto
        ' ao parâmetro hKey.
        hKey = phkResult
        
    End If
    
    ' Ativar API para deletar uma chave do banco de dados
    ' de registro do Windows.
    lngReturn = RegDeleteKey(hKey, lpSubKey)
    
    ' Avaliar se a API retornou sucesso.
    If lngReturn <> 0 Then
        GoTo ErrorHandler
    End If
    
    ' Fechar a chave no banco de dados de registro
    ' do Windows.
    RegCloseKey (hKey)
    
    ' Retornar sucesso ao método.
    DeleteKey = -1
    
    Exit Function
    
ErrorHandler:
    
    ' Fechar a chave no banco de dados de registro
    ' do Windows.
    RegCloseKey (phkResult)
    
    ' Retornar código de erro ao método.
    DeleteKey = 0
    
End Function

Public Function DeleteValue(ByVal hKey As HKEY_ENUM, ByVal lpSubKey As String, _
                            ByVal lpValueName As String, _
                            Optional lpMachineName As Variant) As Long

' Área para declaração de variáveis e objetos de
' escopo local.
Dim lngReturn As Long
Dim phkResult As Long
    
    On Error GoTo ErrorHandler
    
    ' Abrir a chave no banco de dados de
    ' registro do Windows.
    lngReturn = OpenKeyEx(hKey, lpSubKey, _
            phkResult, KEY_ALL_ACCESS, lpMachineName)
    
    ' Avaliar se a API retornou sucesso.
    If lngReturn <> 0 Then
        GoTo ErrorHandler
    End If
    
    ' Ativar API para deletar um valor do banco de dados
    ' de registro do Windows.
    lngReturn = RegDeleteValue(phkResult, lpValueName)
    
    ' Avaliar se a API retornou sucesso.
    If lngReturn <> 0 Then
        GoTo ErrorHandler
    End If
    
    ' Fechar a chave no banco de dados de registro
    ' do Windows.
    RegCloseKey (phkResult)
    
    ' Retornar sucesso ao método.
    DeleteValue = -1
    
    Exit Function
    
ErrorHandler:
    
    ' Fechar a chave no banco de dados de registro
    ' do Windows.
    RegCloseKey (phkResult)
    
    ' Retornar código de erro ao método.
    DeleteValue = 0
    
End Function

Private Function OpenKeyEx(ByVal hKey As Long, ByVal lpSubKey As String, _
                           phkResult As Long, ByVal samDesired As Long, _
                           Optional lpMachineName As Variant) As Long

' Área para declaração de variáveis e objetos de
' escopo local.
Dim lngReturn As Long
    
    On Error GoTo ErrorHandler
    
    ' Avaliar se o parâmetro lpMachineName foi informado.
    If Not IsMissing(lpMachineName) Then
        
        ' Ativar API para conectar-se a um banco de dados de
        ' registro do Windows de uma máquina remota.
        lngReturn = RegConnectRegistry(CStr(lpMachineName), _
                                       hKey, phkResult)
        
        ' Avaliar se a API retornou sucesso.
        If lngReturn <> 0 Then
            GoTo ErrorHandler
        End If
        
        ' Atribuir o handler para a chave do registro remoto
        ' ao parâmetro hKey.
        hKey = phkResult
        
    End If
    
    ' Ativar API para abrir a chave no banco de dados de
    ' registro do Windows.
    lngReturn = RegOpenKeyEx(hKey, lpSubKey, _
                             0&, samDesired, phkResult)
    
    ' Avaliar se a API retornou sucesso.
    If lngReturn <> 0 Then
        GoTo ErrorHandler
    End If
    
    ' Retornar sucesso ao método.
    OpenKeyEx = 0
    
    Exit Function
    
ErrorHandler:
    
    ' Retornar código de erro ao método.
    OpenKeyEx = lngReturn
    
End Function

Public Function QueryValueEx(ByVal hKey As HKEY_ENUM, ByVal lpSubKey As String, _
                             ByVal lpValueName As String, lpType As Long, _
                             lpData As String, _
                             Optional lpMachineName As Variant) As Long

' Área para declaração de variáveis e objetos de
' escopo local.
Dim lngReturn As Long
Dim phkResult As Long
Dim lpDataTmp As String * 255
Dim lpchData As Long
    
    On Error GoTo ErrorHandler
    
    ' Inicializar Buffer.
    lpchData = 255
    
    ' Abrir a chave no banco de dados de
    ' registro do Windows.
    lngReturn = OpenKeyEx(hKey, lpSubKey, _
            phkResult, KEY_READ, lpMachineName)
    
    ' Avaliar se a API retornou sucesso.
    If lngReturn <> 0 Then
        GoTo ErrorHandler
    End If
    
    ' Ativar API para ler o valor.
    lngReturn = RegQueryValueEx(phkResult, lpValueName, _
                                0&, lpType, lpDataTmp, lpchData)
    
    ' Avaliar se a API retornou sucesso.
    If lngReturn <> 0 Then
        GoTo ErrorHandler
    End If
    
    ' Avaliar qual o data type do valor que foi
    ' lido do banco de dados de registro do Windows.
    If lpType = REG_SZ Or lpType = REG_EXPAND_SZ Or _
            lpType = REG_MULTI_SZ Then
        ' Valor possui null terminate.
        lpData = Left$(lpDataTmp, lpchData - 1)
    Else
        lpData = Left$(lpDataTmp, lpchData)
    End If
    
    ' Fechar a chave no banco de dados de registro
    ' do Windows.
    RegCloseKey (phkResult)
    
    ' Retornar sucesso ao método.
    QueryValueEx = -1
    
    Exit Function
    
ErrorHandler:
    
    ' Fechar a chave no banco de dados de registro
    ' do Windows.
    RegCloseKey (phkResult)
    
    ' Retornar código de erro ao método.
    QueryValueEx = 0
    
End Function

Public Function SetValueEx(ByVal hKey As HKEY_ENUM, ByVal lpSubKey As String, _
                           ByVal lpValueName As String, ByVal dwType As REG_TYPE_ENUM, _
                           ByVal lpData As String, _
                           Optional lpMachineName As Variant) As Long

' Área para declaração de variáveis e objetos de
' escopo local.
Dim lngReturn As Long
Dim phkResult As Long
    
    On Error GoTo ErrorHandler
    
    ' Abrir a chave no banco de dados de
    ' registro do Windows.
    lngReturn = OpenKeyEx(hKey, lpSubKey, _
            phkResult, KEY_ALL_ACCESS, lpMachineName)
    
    ' Avaliar se a API retornou sucesso.
    If lngReturn <> 0 Then
        GoTo ErrorHandler
    End If
    
    ' Ativar API para gravar o valor no banco de dados
    ' de registro do Windows.
    lngReturn = RegSetValueEx(phkResult, lpValueName, 0&, _
                              dwType, lpData, Len(lpData))
    
    ' Avaliar se a API retornou sucesso.
    If lngReturn <> 0 Then
        GoTo ErrorHandler
    End If
    
    ' Fechar a chave no banco de dados de registro
    ' do Windows.
    RegCloseKey (phkResult)
    
    ' Retornar sucesso ao método.
    SetValueEx = -1
    
    Exit Function
    
ErrorHandler:
    
    ' Fechar a chave no banco de dados de registro
    ' do Windows.
    RegCloseKey (phkResult)
    
    ' Retornar código de erro ao método.
    SetValueEx = 0
    
End Function
