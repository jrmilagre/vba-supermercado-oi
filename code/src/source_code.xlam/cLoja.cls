VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLoja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Classe

Private Type TClass
    ID         As Long
    Loja       As String
End Type

Private This As TClass

Private Const sTabela As String = "tbl_dfc_lojas"

' Propriedades

Public Property Get ID() As Long: ID = This.ID: End Property
Public Property Let ID(ByVal Value As Long): This.ID = CLng(Value): End Property

Public Property Get Loja() As String: Loja = This.Loja: End Property
Public Property Let Loja(ByVal Value As String): This.Loja = Value: End Property

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
            Me.Loja = .Fields("loja").Value
            
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
            
            .Fields("loja").Value = Me.Loja
            
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

