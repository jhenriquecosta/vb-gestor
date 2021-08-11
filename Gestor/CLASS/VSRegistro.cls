VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VSRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

    Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
    End Type
    
    Const KEY_ALL_ACCESS = &HF003F
    Const KEY_CREATE_LINK = &H20
    Const KEY_CREATE_SUB_KEY = &H4
    Const KEY_ENUMERATE_SUB_KEYS = &H8
    Const KEY_EXECUTE = &H20019
    Const KEY_NOTIFY = &H10
    Const KEY_QUERY_VALUE = &H1
    Const KEY_READ = &H20019
    Const KEY_SET_VALUE = &H2
    Const KEY_WRITE = &H20006

'FecharChave
    Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long
'CriarChave
    Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
'ApagarChave
    Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hkey As Long, ByVal lpSubKey As String) As Long
'ApagarValor
    Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hkey As Long, ByVal lpValueName As String) As Long
'EnumerarChaves
    'Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
'EnumerarValores
    Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
'AbrirChave
    Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
'ConsultarValor
    Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
'DefinirValor
    Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Enum eChave
    chvClasses = &H80000000
    chvUsuarioAtual = &H80000001
    chvMaquinaLocal = &H80000002
    chvUsuarios = &H80000003
    chvConfiguracaoAtual = &H80000005 'Somente Win95/98
    chvDadosDinamicos = &H80000006
    chvDadosPerformance = &H80000004 'Somente WinNT/2000
End Enum

Public Enum eTipoDado
    tipBinario = 3 'Sequencia de bytes "não-texto"
    tipDWord = 4 'Mesmo que DWordLittleEndian
    tipDWordBigEndian = 5 'Inteiro de 32 bits. Forma oposta ao armazenamento dos PCs intel. A ordem dos bits é inversa
    tipDWordLittleEndian = 4 'Inteiro de 32 bits. É como os PCs Intel armazenam numeros
    tipStringExpandida = 7 'String terminada em "nulo". Contém variáveis de ambiente não expandidas
    tipNenhum = 0 'Sem tipo de dado
    tipListaRecursos = 8 'Lista de num mapa de recursos
    tipString = 1 'String terminada em nulo
End Enum

Public Function Escrever(Chave As eChave, Subchave As String, Parametro As String, Valor As String, Tipo As eTipoDado) As Boolean
    If bRegistrado Then
        Dim hChave As Long
        
        hChave = Criar(Chave, Subchave)
        If hChave <> 0 Then
            Escrever = Definir(hChave, Parametro, Valor, Tipo)
        End If
        Fechar Chave
    End If
End Function

Public Function Ler(Chave As eChave, Subchave As String, Parametro As String) As Variant
    If bRegistrado Then
        Dim hChave As Long
        Dim Tipo As eTipoDado
        
        hChave = Abrir(Chave, Subchave)
        If hChave <> 0 Then
            Ler = Consultar(hChave, Parametro, Tipo)
        End If
        Fechar Chave
    End If
End Function

Public Function Apagar(Chave As eChave, Optional Subchave As String, Optional Parametro As String) As Boolean
    If bRegistrado Then
        Dim hChave As Long
    
        If Parametro <> "" Then
            hChave = Abrir(Chave, Subchave)
            If hChave <> 0 Then
                Apagar = ApagarParametro(hChave, Parametro)
            End If
            Fechar Chave
        Else
            If Subchave <> "" Then
                Apagar = ApagarChave(Chave, Subchave)
            End If
        End If
    End If
End Function

Public Function Existe(Chave As eChave, Subchave As String) As Boolean
    If bRegistrado Then
        Existe = IIf(Abrir(Chave, Subchave) <> 0, True, False)
        Fechar Chave
    End If
End Function

Private Function Fechar(Chave As eChave) As Boolean
'PLATAFORMAS
'--------------
'   Windows 95
'   Windows 98
'   Windows NT 3.1/+
'   Windows 2000
'   Windows CE 1.0/+

'DESCRICAO
'--------------
'   Fecha uma chave de registro que tenha sido previamente aberta. Esta prática libera recursos do computador
    
    If bRegistrado Then Fechar = IIf(RegCloseKey(Chave) = 0, True, False)
End Function

Private Function Criar(Chave As eChave, Subchave As String) As Long
'PLATAFORMAS
'--------------
'   Windows 95
'   Windows 98
'   Windows NT 3.1/+
'   Windows 2000
'   Windows CE 1.0/+

'DESCRICAO
'--------------
'   Cria uma nova chave de registro. Se a chave já existir, ela será aberta. O handle da chave aberta é o retorno da função

    If bRegistrado Then
        Dim hChave As Long
        Dim typSeguranca As SECURITY_ATTRIBUTES
        Dim lngNovoUsado As Long
        
        With typSeguranca
            .nLength = Len(typSeguranca)
            .lpSecurityDescriptor = 0
            .bInheritHandle = 1
        End With
        
        Criar = IIf(RegCreateKeyEx(Chave, Subchave, 0, "", 0, KEY_ALL_ACCESS, typSeguranca, hChave, lngNovoUsado) = 0, hChave, 0)
    End If
End Function

Private Function ApagarChave(Chave As eChave, Subchave As String) As Boolean
'PLATAFORMAS
'--------------
'   Windows 95
'   Windows 98
'   Windows NT 3.1/+
'   Windows 2000
'   Windows CE 1.0/+

'DESCRICAO
'--------------
'   Apaga uma chave de registro com todos os valores que ela contém.
'   No WinNT/2000 ocorrerá um erro se houver subchaves.
    
    If bRegistrado Then ApagarChave = IIf(RegDeleteKey(Chave, Subchave) = 0, True, False)
End Function

Private Function ApagarParametro(hChave As Long, Parametro As String) As Boolean
'PLATAFORMAS
'--------------
'   Windows 95
'   Windows 98
'   Windows NT 3.1/+

'DESCRICAO
'--------------
'   Apaga um Parametro guardado numa chave específica do registro. Esta funcao só apaga Parametroes; não consegue
'apagar subchaves.
    
    If bRegistrado Then ApagarParametro = IIf(RegDeleteValue(hChave, Parametro) = 0, True, False)
End Function

Private Function Abrir(Chave As eChave, Subchave As String) As Long
'PLATAFORMAS
'--------------
'   Windows 32's
'   Windows 98
'   Windows NT 3.1/+

'DESCRICAO
'--------------
'   Abre um chave do registro e retorna o handle da chave aberta.
'   Não consegue criar a chave, caso ela não exista.

    If bRegistrado Then
        Dim hChave As Long
    
        Abrir = IIf(RegOpenKeyEx(Chave, Subchave, 0, KEY_ALL_ACCESS, hChave) = 0, hChave, 0)
    End If
End Function

Private Function Consultar(hChave As Long, Parametro As String, ByRef Tipo As eTipoDado) As Variant
'PLATAFORMAS
'--------------
'   Windows 95
'   Windows 98
'   Windows NT 3.1/+
'   Windows 2000
'   Windows CE 1.0/+

'DESCRICAO
'--------------
'   Lê um Parametro de uma chave do registro.
    If bRegistrado Then
        Dim strBuffer As String
        Dim lngTamanho As Long
        Dim lngSucesso As Long
    
        strBuffer = Space(255)
        lngTamanho = 255
        lngSucesso = RegQueryValueEx(hChave, Parametro, 0, Tipo, ByVal strBuffer, lngTamanho)
        If lngSucesso = 0 Then
            If Tipo = tipString Or Tipo = tipStringExpandida Then
                Consultar = Left(strBuffer, lngTamanho - 1)
            Else
                Consultar = strBuffer
            End If
        End If
    End If
End Function

Private Function Definir(hChave As Long, Parametro As String, Valor As String, Tipo As eTipoDado) As Boolean
'PLATAFORMAS
'--------------
'   Windows 95
'   Windows 98
'   Windows NT 3.1/+
'   Windows 2000
'   Windows CE 1.0/+

'DESCRICAO
'--------------
'   Escreve um valor numa chave de registro. Se o valor não exisitir, ele será criado.
'   Ao escrever uma string ou um valor numérico simples, o parâmetro lpData deve ser passado
'por valor. Qualquer outro valor não necessita do ByVal
    If bRegistrado Then
        Dim strBuffer As String
        
        If Tipo = tipString Or Tipo = tipStringExpandida Then
            strBuffer = Valor & vbNullChar
            Definir = IIf(RegSetValueEx(hChave, Parametro, 0, Tipo, ByVal strBuffer, Len(strBuffer)) = 0, True, False)
        Else
            Definir = IIf(RegSetValueEx(hChave, Parametro, 0, Tipo, strBuffer, Len(strBuffer)) = 0, True, False)
        End If
    End If
End Function

Private Sub Class_Initialize()
    ValidaComponente "CLASS"
End Sub
