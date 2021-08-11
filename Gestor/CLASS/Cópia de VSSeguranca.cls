VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VSSeguranca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function apiSerialNumber Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Util As New VSUtil

Public Function Mascara(Senha As String) As String
    On Error GoTo Trata
    If bRegistrado Then
        Dim i As Integer
        Dim Tam As Byte
        Tam = Len(Senha)
        For i = 1 To Tam
            If i Mod 2 = 0 Then
                Mascara = Mascara & Chr(Asc(Mid(Senha, i, 1)) + 6 + i)
            Else
                Mascara = Mascara & Chr(Asc(Mid(Senha, i, 1)) + 2 + i)
            End If
        Next
    End If
    Exit Function
Trata:
    If Err.Number <> 0 Then
        Util.Avisa Err.Description
        Screen.MousePointer = 0
    End If
End Function

Public Function Desmascara(Senha As String) As String
    On Error GoTo Trata
    If bRegistrado Then
        Dim Senha_Cripto As String
        Dim Tam As Byte
        Dim i As Integer
        Tam = Len(Senha)
        For i = 1 To Tam
            Senha_Cripto = Mid(Senha, i, 1)
            If i Mod 2 = 0 Then
                Desmascara = Desmascara & Chr(Asc(Senha_Cripto) - (6 + i))
            Else
                Desmascara = Desmascara & Chr(Asc(Senha_Cripto) - (2 + i))
            End If
        Next
    End If
    Exit Function
Trata:
    If Err.Number <> 0 Then
        Util.Avisa Err.Description
        Screen.MousePointer = 0
    End If

End Function

Public Function NumSerie() As String
    On Error GoTo Trata
    If bRegistrado Then
        Dim Diretorio_Raiz As String
        Dim LabelVol As String
        Dim TamVol As Long
        Dim Dec_SerNum As Long
        Dim Oct_SerNum As Long
        Dim MaxLen As Long
        Dim Flags As Long
        Dim Nome As String
        Dim TamNome As Long
        
        Diretorio_Raiz = Left(CurDir, 3)
        If apiSerialNumber(Diretorio_Raiz, LabelVol, TamVol, Dec_SerNum, MaxLen, Flags, Nome, TamNome) Then
            'Retorna um monte de informações do sistema
            NumSerie = Format(Hex(Dec_SerNum), "00000000") 'Formata o numero de serie
        Else ' Erro na função. Retorna "0000-0000"
            NumSerie = "00000000"
        End If
    End If
    Exit Function
Trata:
    If Err.Number <> 0 Then
        Util.Avisa Err.Description
        Screen.MousePointer = 0
    End If
    
End Function

Public Function Criptografa(Texto As String) As String
    On Error GoTo Trata
    
    If bRegistrado Then
        Dim i As Integer
        Dim Tam As Integer
        Dim Temp As String
        Tam = Len(Texto)
        Temp = ""
        
        'Debug.Print "Texto: " & Texto
        If Tam Mod 2 = 0 Then
            For i = 1 To Tam
                If i Mod 2 = 0 Then
                    Temp = Temp & Chr(255 - Asc(Mid(Texto, Tam + 2 - i, 1)))
                Else
                    Temp = Temp & Chr(255 - Asc(Mid(Texto, i, 1)))
                End If
                'Debug.Print Chr(255 - Asc(Mid(Temp, I, 1)))
            Next
        Else
            For i = 1 To Tam
                If i Mod 2 = 0 Then
                    Temp = Temp & Chr(255 - Asc(Mid(Texto, i, 1)))
                Else
                    Temp = Temp & Chr(255 - Asc(Mid(Texto, Tam + 1 - i, 1)))
                End If
            Next
        End If
        
        'Debug.Print "Enbaralhado + Complementar: " & Temp
        
        If Tam <= 255 Then
            For i = 1 To Tam
                Mid(Temp, i, 1) = Chr(IIf(Asc(Mid(Temp, i, 1)) + i > 255, (Asc(Mid(Temp, i, 1)) + i) - 255, 255 - (Asc(Mid(Temp, i, 1)) + i)))
            Next
        End If
        
        'Debug.Print "Incrementação: " & Temp
        
        Criptografa = Temp
    End If
    Exit Function
    
Trata:
    If Err.Number <> 0 Then
        Util.Avisa Err.Description
        Screen.MousePointer = 0
    End If
End Function


Public Function DesCriptografa(Texto As String) As String
    On Error GoTo Trata
    
    If bRegistrado Then
        Dim i As Integer
        Dim Tam As Integer
        Dim Temp As String
    
        Tam = Len(Texto)
        
        'Debug.Print "Texto: " & Texto
        
        If Tam <= 255 Then
            For i = 1 To Tam
                Mid(Texto, i, 1) = Chr(IIf(Asc(Mid(Texto, i, 1)) + i > 255, (Asc(Mid(Texto, i, 1)) + i) - 255, 255 - (Asc(Mid(Texto, i, 1)) + i)))
                'Debug.Print Chr(255 - Asc(Mid(Texto, I, 1)))
            Next
        End If
        
        'Debug.Print "Desincrementação: " & Texto
        
        If Tam Mod 2 = 0 Then
            For i = 1 To Tam
                If i Mod 2 = 0 Then
                    Temp = Temp & Chr(255 - Asc(Mid(Texto, Tam + 2 - i, 1)))
                Else
                    Temp = Temp & Chr(255 - Asc(Mid(Texto, i, 1)))
                End If
                'Debug.Print Chr(255 - Asc(Mid(Temp, I, 1)))
            Next
        Else
            For i = 1 To Tam
                If i Mod 2 = 0 Then
                    Temp = Temp & Chr(255 - Asc(Mid(Texto, i, 1)))
                Else
                    Temp = Temp & Chr(255 - Asc(Mid(Texto, Tam + 1 - i, 1)))
                End If
            Next
        End If
    
        'Debug.Print "Desenbaralhado + Complementar (Resposta): " & Temp
        
        DesCriptografa = Temp
    End If
    Exit Function
    
Trata:
    If Err.Number <> 0 Then
        Util.Avisa Err.Description
        Screen.MousePointer = 0
    End If
End Function

Public Function ExisteUsuario(Bdados As Object, User As String) As String
    On Error GoTo Trata
    If bRegistrado Then
        Dim Sql As String
        Dim RS As VSRecordset
        
        ExisteUsuario = ""
        Sql = "SELECT * FROM TAB_USUARIO WHERE TUS_COD_USUARIO = '" & (User) & "'"
        If Bdados.AbreTabela(Sql, RS) Then
            If RS!TUS_ATIVO = True Or RS!TUS_ATIVO = 1 Then ExisteUsuario = (RS!TUS_NOME)
        End If
        Bdados.FechaTabela RS
    End If
    Exit Function
Trata:
    If Err.Number <> 0 Then
        Util.Erro Err.Description
        Screen.MousePointer = 0
    End If
End Function

Private Sub Class_Initialize()
    ValidaComponente "CLASS"
End Sub
