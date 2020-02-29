VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cConta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Classe

Private Type TClass
    ID         As Long
    Conta      As String
    Subgrupo   As String
    DfcID      As Integer
End Type

Private This As TClass

Private Const sTabela As String = "tbl_dfc_contas"

' Propriedades

Public Property Get ID() As Long: ID = This.ID: End Property
Public Property Let ID(ByVal Value As Long): This.ID = CLng(Value): End Property

Public Property Get Conta() As String: Conta = This.Conta: End Property
Public Property Let Conta(ByVal Value As String): This.Conta = Value: End Property

Public Property Get Subgrupo() As String: Subgrupo = This.Subgrupo: End Property
Public Property Let Subgrupo(ByVal Value As String): This.Subgrupo = Value: End Property

Public Property Get DfcID() As Integer: DfcID = This.DfcID: End Property
Public Property Let DfcID(ByVal Value As Integer): This.DfcID = Value: End Property

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
            Me.Conta = .Fields("conta").Value
            Me.Subgrupo = .Fields("subgrupo").Value
            Me.DfcID = .Fields("dfc_id").Value
            
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
            
            .Fields("conta").Value = Me.Conta
            .Fields("subgrupo").Value = Me.Subgrupo
            .Fields("dfc_id").Value = Me.DfcID
            
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
Public Function GetGrupoDFC(ID As Integer) As String

    Dim r As New ADODB.Recordset

    sSQL = "SELECT grupo FROM tbl_dfc_mascara WHERE id = " & ID
    
    Set r = cnn.Execute(sSQL)
    
    GetGrupoDFC = r.Fields("grupo").Value
    
    Set r = Nothing

End Function

