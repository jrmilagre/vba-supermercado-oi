VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_dfc_Lancamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Classe

Private Type TClass
    ID                  As Long
    Data                As Date
    Valor               As Currency
    ContaContabilID     As Long
    ContaCorrenteID     As Long
    LojaID              As Long
    Historico           As String
End Type

Private This As TClass

Private Const sTabela As String = "tbl_dfc_lancamentos"

' Propriedades

Public Property Get ID() As Long: ID = This.ID: End Property
Public Property Let ID(ByVal Value As Long): This.ID = CLng(Value): End Property

Public Property Get Data() As Variant: Data = This.Data: End Property
Public Property Let Data(ByVal Value As Variant): This.Data = Value: End Property

Public Property Get Valor() As Variant: Valor = This.Valor: End Property
Public Property Let Valor(ByVal Value As Variant): This.Valor = Value: End Property

Public Property Get ContaContabilID() As Long: ContaContabilID = This.ContaContabilID: End Property
Public Property Let ContaContabilID(ByVal Value As Long): This.ContaContabilID = Value: End Property

Public Property Get ContaCorrenteID() As Long: ContaCorrenteID = This.ContaCorrenteID: End Property
Public Property Let ContaCorrenteID(ByVal Value As Long): This.ContaCorrenteID = Value: End Property

Public Property Get LojaID() As Long: LojaID = This.LojaID: End Property
Public Property Let LojaID(ByVal Value As Long): This.LojaID = Value: End Property

Public Property Get Historico() As String: Historico = This.Historico: End Property
Public Property Let Historico(ByVal Value As String): This.Historico = Value: End Property

' Métodos

Public Sub CRUD(Acao As Integer, Optional ID As Long)
    
    If ID > 0 Then
    
        sSQL = "SELECT * FROM " & sTabela & " WHERE id = " & ID
        
    End If

    Set rst = New ADODB.Recordset
    
    If Acao = Read Then ' READ
        
        With rst
        
            .Open sSQL, cnn, adOpenStatic, adLockReadOnly
        
            Me.ID = .Fields("id").Value
            Me.Data = .Fields("data").Value
            Me.Valor = .Fields("valor").Value
            Me.ContaContabilID = .Fields("conta_contabil_id").Value
            Me.ContaCorrenteID = .Fields("conta_corrente_id").Value
            Me.LojaID = .Fields("loja_id").Value
            Me.Historico = .Fields("historico").Value
            
        End With
        
    ElseIf Acao = Create Or Acao = Update Then
        
        With rst
        
            If Acao = Create Then
            
                .Open sTabela, cnn, adOpenDynamic, adLockOptimistic, adCmdTable
                
                .AddNew
                
                Me.ID = .Fields("id").Value
                
            ElseIf Acao = Update Then
            
                .Open sSQL, cnn, adOpenDynamic, adLockOptimistic
                
            End If
            
            .Fields("data").Value = Me.Data
            .Fields("valor").Value = Me.Valor
            .Fields("conta_contabil_id").Value = Me.ContaContabilID
            .Fields("conta_corrente_id").Value = Me.ContaCorrenteID
            .Fields("loja_id").Value = Me.LojaID
            .Fields("historico").Value = Me.Historico
            
            .Update
            
        End With
        
    ElseIf Acao = Delete Then
    
        With rst
        
            .Open sSQL, cnn, adOpenDynamic, adLockOptimistic
            
            .Delete
            
        End With
        
    End If
    
    Set rst = Nothing

End Sub
Public Function Todos(Optional Ordem As String) As ADODB.Recordset

    sSQL = "SELECT * "
    sSQL = sSQL & "FROM " & sTabela & " "
    
'    sSQL = "SELECT nome, id, nascimento, salario "
'    sSQL = sSQL & "FROM tbl_contatos "
'    sSQL = sSQL & "WHERE nome LIKE ('%" & txbPesquisa.Text & "%') "
    If Ordem <> "" Then
        sSQL = sSQL & "ORDER BY " & Ordem
    End If
    
    Set Todos = New ADODB.Recordset
    
    With Todos
        .CursorLocation = adUseServer
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .PageSize = 24
        .Open sSQL, cnn
    End With
    
End Function
Public Function ContasContabeis() As ADODB.Recordset

    sSQL = "SELECT tbl_dfc_mascara.grupo, tbl_dfc_contas_contabeis.subgrupo, "
    sSQL = sSQL & "tbl_dfc_contas_contabeis.conta, tbl_dfc_contas_contabeis.id "
    sSQL = sSQL & "FROM tbl_dfc_mascara INNER JOIN tbl_dfc_contas_contabeis ON "
    sSQL = sSQL & "tbl_dfc_mascara.id = tbl_dfc_contas_contabeis.dfc_id "
    sSQL = sSQL & "ORDER BY tbl_dfc_mascara.ordem, tbl_dfc_contas_contabeis.subgrupo, tbl_dfc_contas_contabeis.conta "
    
    Set ContasContabeis = New ADODB.Recordset
    
    With ContasContabeis
        .CursorLocation = adUseServer
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .Open sSQL, cnn
    End With

End Function
