Attribute VB_Name = "m_Database"
Option Explicit

' Desenvolvedor: Jairo Milagre da Fonseca Jr

' BIBLIOTECAS necess�rias:
' ---> Microsoft ActiveX Data Objects 2.8 Library
' ---> Microsoft ADO Ext. 2.8 for DDL and Security

' Declara��o de vari�veis p�blicas
Public Enum eCrud
    Create = 1
    Read = 2
    Update = 3
    Delete = 4
End Enum

Public cnn  As ADODB.Connection
Public rst  As ADODB.Recordset
Public cat  As ADOX.Catalog
Public sSQL As String
Public Function Conecta() As Boolean

    Dim oConfig As New c__Config
    
    ' Cria objeto de conex�o com o banco de dados
    Set cnn = New ADODB.Connection
    Set cat = New ADOX.Catalog
    
    ' Inicia a fun��o com o valor Falso, pois a conex�o ainda n�o aconteceu
    Conecta = False
    
    With cnn
        .Provider = oConfig.GetProvedorDB   ' Escolhe o provedor da conex�o
        On Error GoTo Erro                  ' Se a conex�o der problema, desvia para o r�tulo Erro
        .Open oConfig.GetCaminhoBD          ' Abre a conex�o com o banco de dados
        Set cat.ActiveConnection = cnn      ' Seta cat�logo
    End With
    
    ' Se a conex�o for um sucesso, retorna Verdadeiro
    Conecta = True
    
    ' Sai da fun��o
    Exit Function
    
Erro:
    ' Mensagem caso a conex�o com o banco de dados der problema
    MsgBox "Banco de dados n�o existe ou n�o est� acess�vel.", vbInformation
    
End Function
Public Sub Desconecta()
    
    cnn.Close           ' Fecha o objeto de conex�o
    Set cat = Nothing

End Sub



