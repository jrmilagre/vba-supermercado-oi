VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_dfc_Lojas 
   Caption         =   ":: Cadastro de Lojas ::"
   ClientHeight    =   9015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9960
   OleObjectBlob   =   "f_dfc_Lojas.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "f_dfc_Lojas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oLoja               As New c_dfc_Loja
Private colControles        As New Collection       ' Para eventos de campos
Private myRst               As New ADODB.Recordset
Private bAtualizaScrool     As Boolean

Private Sub UserForm_Initialize()
    
    Call PopulaCombos
    
    Call EventosCampos
    
    Call BuscaRegistros

End Sub
Private Sub UserForm_Terminate()
    
    Set oLoja = Nothing
    Set myRst = Nothing
    
    Call Desconecta
    
End Sub
Private Sub btnIncluir_Click()
    
    Call PosDecisaoTomada("Inclusão")
    
End Sub
Private Sub btnAlterar_Click()
    
    Call PosDecisaoTomada("Alteração")

End Sub
Private Sub btnExcluir_Click()

    Call PosDecisaoTomada("Exclusão")
    
End Sub
Private Sub PosDecisaoTomada(Decisao As String)

    btnCancelar.Visible = True: btnConfirmar.Visible = True
    btnConfirmar.Caption = "Confirmar " & Decisao
    btnCancelar.Caption = "Cancelar " & Decisao
    
    btnIncluir.Visible = False: btnAlterar.Visible = False: btnExcluir.Visible = False
    
    MultiPage1.Value = 1
    
    If Decisao <> "Exclusão" Then
    
        If Decisao = "Inclusão" Then
        
            Call Campos("Limpar")
            
        End If
        
        Call Campos("Habilitar")
        
        txbLoja.SetFocus
        
    End If
    
    MultiPage1.Pages(0).Enabled = False
    
End Sub
Private Sub btnConfirmar_Click()
    
    Call Gravar(Replace(btnConfirmar.Caption, "Confirmar ", ""))
    
End Sub
Private Sub btnCancelar_Click()
    
    btnIncluir.Visible = True: btnAlterar.Visible = True: btnExcluir.Visible = True
    btnConfirmar.Visible = False: btnCancelar.Visible = False
    
    Call Campos("Limpar")
    Call Campos("Desabilitar")
    
    btnAlterar.Enabled = False
    btnExcluir.Enabled = False
    btnIncluir.SetFocus
   
    MultiPage1.Value = 0
    
    lstPrincipal.ListIndex = -1 ' Tira a seleção
    
End Sub
Private Sub lstPrincipal_Change()

    Dim n As Long
    
    If lstPrincipal.ListIndex >= 0 Then
    
        btnAlterar.Enabled = True
        btnExcluir.Enabled = True
    
        With oLoja
    
            .CRUD eCrud.Read, (CLng(lstPrincipal.List(lstPrincipal.ListIndex, 1)))
    
            lblCabID.Caption = IIf(.ID = 0, "", Format(.ID, "000000"))
            lblCabLoja.Caption = .Loja
            txbLoja.Text = .Loja
            
        End With
        
    End If

End Sub
Private Sub Campos(Acao As String)
    
    Dim sDecisao    As String
    Dim b           As Boolean
    
    sDecisao = Replace(btnConfirmar.Caption, "Confirmar ", "")
    
    If Acao <> "Limpar" Then
    
        If Acao = "Desabilitar" Then
            b = False
        ElseIf Acao = "Habilitar" Then
            b = True
        End If
        
        MultiPage1.Pages(0).Enabled = Not b
        
        txbLoja.Enabled = b: lblLoja.Enabled = b
        
    Else
    
        lblCabID.Caption = ""
        lblCabLoja.Caption = ""
        txbLoja.Text = Empty
             
    End If

End Sub
Private Sub lstPrincipalPopular(Pagina As Long)

    Dim n           As Byte
    Dim vNascimento As Variant
    Dim vSalario    As Variant
    Dim oLegenda     As control
    
    ' Limpa cores da legenda
    For n = 1 To myRst.PageSize
        Set oLegenda = Controls("l" & Format(n, "00")): oLegenda.BackColor = &H8000000F
    Next n

    ' Define página que será exibida do Recordset
    myRst.AbsolutePage = Pagina
    
    With lstPrincipal
        .Clear                                      ' Limpa conteúdo
        .ColumnCount = 2                            ' Define número de colunas
        .ColumnWidths = "180 pt; 0pt;"              ' Configura largura das colunas
        .Font = "Consolas"                          ' Configura fonte
        
        n = 1
        
        While Not myRst.EOF = True And n <= myRst.PageSize
            
            ' Preenche ListBox
            .AddItem
            
            .List(.ListCount - 1, 0) = myRst.Fields("loja").Value
            .List(.ListCount - 1, 1) = myRst.Fields("id").Value
            
            ' Colore a legenda
'            Set oLegenda = Controls("l" & Format(n, "00"))
'
'            If myRst.Fields("sexo").Value = "F" Then
'                oLegenda.BackColor = &HFF80FF
'            ElseIf myRst.Fields("sexo").Value = "M" Then
'                oLegenda.BackColor = &HFF8080
'            Else
'                oLegenda.BackColor = &H8000000F
'            End If
            
            ' Próximo registro
            myRst.MoveNext: n = n + 1
            
        Wend
        
    End With
    
    ' Posiciona scroll de navegação em páginas
    lblPaginaAtual.Caption = Pagina
    lblNumeroPaginas.Caption = myRst.PageCount
    bAtualizaScrool = False: scrPagina.Value = CLng(lblPaginaAtual.Caption): bAtualizaScrool = True
    lblTotalRegistros.Caption = Format(myRst.RecordCount, "#,##0")
    
    ' Trata os botões de navegação
    Call TrataBotoesNavegacao

End Sub
Private Sub Gravar(Decisao As String)

    Dim vbResposta  As VbMsgBoxResult
    Dim e           As eCrud
    
    vbResposta = MsgBox("Deseja realmente fazer a " & Decisao & "?", vbYesNo + vbQuestion, "Pergunta")
    
    If vbResposta = vbYes Then
    
        If Decisao <> "Exclusão" Then
        
            If txbLoja.Text = Empty Then
                MsgBox "Campo 'Loja' é obrigatório", vbCritical: MultiPage1.Value = 1: txbLoja.SetFocus
            Else
                
                With oLoja
                    
                    .Loja = Replace(txbLoja.Text, "'", "`")
                    
                    If Decisao = "Inclusão" Then
                        .CRUD eCrud.Create
                    Else
                        .CRUD eCrud.Update, .ID
                    End If
                    
                End With
                
                MsgBox Decisao & " realizada com sucesso.", vbInformation, Decisao & " de registro"
                
                Call BuscaRegistros
                                    
            End If
        
        Else ' Se for exclusão
        
            oLoja.CRUD eCrud.Delete, oLoja.ID
                
            MsgBox Decisao & " realizada com sucesso.", vbInformation, Decisao & " de registro"
            
            Call BuscaRegistros
            
        End If
               
    ElseIf vbResposta = vbNo Then
        
        If Decisao = "Exclusão" Then
            
            Call btnCancelar_Click
            
        End If
        
    End If
    
End Sub
Private Sub EventosCampos()

    ' Declara variáveis
    Dim oControle   As MSForms.control
    Dim oEvento     As c__EventoCampo
    Dim sTag        As String
    Dim sField()    As String
    
    ' Laço para percorrer todos os TextBox e atribuir eventos
    ' de acordo com o tipo de cada campo
    For Each oControle In Me.Controls
    
        If Len(oControle.Tag) > 0 Then
        
            If TypeName(oControle) = "TextBox" Then
            
                Set oEvento = New c__EventoCampo
                
                With oEvento
                
                    sField() = Split(oControle.Tag, ".")
                    
                    oControle.ControlTipText = cat.Tables(sField(0)).Columns(sField(1)).Properties("Description").Value
                    
                    .FieldType = cat.Tables(sField(0)).Columns(sField(1)).Type
                    .MaxLength = cat.Tables(sField(0)).Columns(sField(1)).DefinedSize
                    .Nullable = cat.Tables(sField(0)).Columns(sField(1)).Properties("Nullable")
                    
                    Set .cGeneric = oControle
                    
                End With
                
                colControles.Add oEvento
                
            End If
            
        End If
    Next

End Sub
Private Sub btnFiltrar_Click()

    Call BuscaRegistros

End Sub
Private Sub BuscaRegistros(Optional Ordem As String)

    Dim n As Byte
    Dim o As control

    Set myRst = oLoja.Todos(Ordem, True)
    
    If myRst.PageCount > 0 Then
        
        bAtualizaScrool = False
        
        With scrPagina
            .Max = myRst.PageCount
            .Value = myRst.PageCount
        End With
        
        Call lstPrincipalPopular(myRst.PageCount)
        
    Else
    
        lstPrincipal.Clear
        
        For n = 1 To myRst.PageSize
            Set o = Controls("l" & Format(n, "00")): o.BackColor = &H8000000F
        Next n
        
    End If
    
    Call btnCancelar_Click
    

End Sub
Private Sub TrataBotoesNavegacao()

    If CLng(lblPaginaAtual.Caption) = myRst.PageCount And CLng(lblPaginaAtual.Caption) > 1 Then
    
        btnPaginaInicial.Enabled = True
        btnPaginaAnterior.Enabled = True
        btnPaginaFinal.Enabled = False
        btnPaginaSeguinte.Enabled = False
        
    ElseIf CLng(lblPaginaAtual.Caption) < myRst.PageCount And CLng(lblPaginaAtual.Caption) = 1 Then
    
        btnPaginaInicial.Enabled = False
        btnPaginaAnterior.Enabled = False
        btnPaginaFinal.Enabled = True
        btnPaginaSeguinte.Enabled = True
        
    ElseIf CLng(lblPaginaAtual.Caption) = myRst.PageCount And CLng(lblPaginaAtual.Caption) = 1 Then
    
        btnPaginaInicial.Enabled = False
        btnPaginaAnterior.Enabled = False
        btnPaginaFinal.Enabled = False
        btnPaginaSeguinte.Enabled = False
    
    Else
    
        btnPaginaInicial.Enabled = True
        btnPaginaAnterior.Enabled = True
        btnPaginaFinal.Enabled = True
        btnPaginaSeguinte.Enabled = True
        
    End If

End Sub
Private Sub btnPaginaInicial_Click()
    
    Call lstPrincipalPopular(1)
    
End Sub
Private Sub btnPaginaAnterior_Click()

    Call lstPrincipalPopular(CLng(lblPaginaAtual.Caption) - 1)
    
End Sub
Private Sub btnPaginaSeguinte_Click()

    Call lstPrincipalPopular(CLng(lblPaginaAtual.Caption) + 1)

End Sub
Private Sub btnPaginaFinal_Click()

    Call lstPrincipalPopular(myRst.PageCount)
    
End Sub
Private Sub btnRegistroAnterior_Click()

        If lstPrincipal.ListIndex > 0 Then
        
            lstPrincipal.ListIndex = lstPrincipal.ListIndex - 1
            
        ElseIf lstPrincipal.ListIndex = 0 And CLng(lblPaginaAtual.Caption) > 1 Then
            
            Call lstPrincipalPopular(CLng(lblPaginaAtual.Caption) - 1)
            
            lstPrincipal.ListIndex = myRst.PageSize - 1
            
        ElseIf CLng(lblPaginaAtual.Caption) = 1 And lstPrincipal.ListIndex = 0 Then
        
            MsgBox "Primeiro registro"
            Exit Sub
            
        Else
        
            lstPrincipal.ListIndex = -1
            
        End If
        
End Sub
Private Sub btnRegistroSeguinte_Click()

    If lstPrincipal.ListIndex = -1 Then
        
        lstPrincipal.ListIndex = 0
    
    ElseIf lstPrincipal.ListIndex = myRst.PageSize - 1 And CLng(lblPaginaAtual.Caption) < myRst.PageCount Then
        
        Call lstPrincipalPopular(CLng(lblPaginaAtual.Caption) + 1)
        
        lstPrincipal.ListIndex = 0
        
    ElseIf CLng(lblPaginaAtual.Caption) = myRst.PageCount And (lstPrincipal.ListIndex + 1) = lstPrincipal.ListCount Then
    
        MsgBox "Último registro"
        Exit Sub
        
    Else
    
        lstPrincipal.ListIndex = lstPrincipal.ListIndex + 1
    
    End If
    
End Sub
Private Sub scrPagina_Change()

    If bAtualizaScrool = True Then
        
        Call lstPrincipalPopular(scrPagina.Value)
        
    End If

End Sub
Private Sub PopulaCombos()

'    With cbbSexo
'        .Clear
'        .ColumnCount = 2
'        .ColumnWidths = "60pt; 0pt;"
'
'        .AddItem
'        .List(.ListCount - 1, 0) = "MASCULINO"
'        .List(.ListCount - 1, 1) = "M"
'
'        .AddItem
'        .List(.ListCount - 1, 0) = "FEMININO"
'        .List(.ListCount - 1, 1) = "F"
'    End With

End Sub
Private Sub lstPrincipal_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    MultiPage1.Value = 1
    
End Sub
Private Sub lblHdNome_Click()

    Call BuscaRegistros("loja")
    
End Sub
