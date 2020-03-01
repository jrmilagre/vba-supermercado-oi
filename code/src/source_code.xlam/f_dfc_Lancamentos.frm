VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_dfc_Lancamentos 
   Caption         =   ":: Lançamentos ::"
   ClientHeight    =   9015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9960
   OleObjectBlob   =   "f_dfc_Lancamentos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "f_dfc_Lancamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private oLancamento         As New c_dfc_Lancamento
Private oContaCorrente      As New c_dfc_ContaCorrente
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
    
    Set oLancamento = Nothing
    Set myRst = Nothing
    
    Call Desconecta
    
End Sub
Private Sub btnValor_Click()
    ccurVisor = IIf(txbValor.Text = "", 0, CCur(txbValor.Text))
    txbValor.Text = Format(GetCalculadora, "#,##0.00")
End Sub
Private Sub btnData_Click()
    dtDate = IIf(txbData.Text = Empty, Date, txbData.Text)
    txbData.Text = GetCalendario
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
        
        txbData.SetFocus
        
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
    
        With oLancamento
    
            .CRUD eCrud.Read, (CLng(lstPrincipal.List(lstPrincipal.ListIndex, 0)))
    
            lblCabID.Caption = IIf(.ID = 0, "", Format(.ID, "000000"))
            lblCabNome.Caption = .Data
            txbData.Text = .Data
            txbValor.Text = Format(.Valor, "#,##0.00")
            txbHistorico.Text = .Historico
            
            For n = 0 To cbbContaContabil.ListCount
                If cbbContaContabil.List(n, 1) = .ContaContabilID Then: cbbContaContabil.ListIndex = n: Exit For
            Next n
            
            For n = 0 To cbbContaCorrente.ListCount
                If cbbContaCorrente.List(n, 1) = .ContaCorrenteID Then: cbbContaCorrente.ListIndex = n: Exit For
            Next n
            
            For n = 0 To cbbLoja.ListCount
                If cbbLoja.List(n, 1) = .LojaID Then: cbbLoja.ListIndex = n: Exit For
            Next n
            
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
        
        txbData.Enabled = b: lblData.Enabled = b: btnData.Enabled = b
        txbValor.Enabled = b: lblValor.Enabled = b: btnValor.Enabled = b
        cbbContaContabil.Enabled = b: lblContaContabil.Enabled = b
        cbbContaCorrente.Enabled = b: lblContaCorrente.Enabled = b
        cbbLoja.Enabled = b: lblLoja.Enabled = b
        txbHistorico.Enabled = b: lblHistorico.Enabled = b
        
    Else
    
        lblCabID.Caption = ""
        lblCabNome.Caption = ""
        
        txbData.Text = IIf(sDecisao = "Inclusão", Date, Empty)
        txbValor.Text = IIf(sDecisao = "Inclusão", Format(0, "#,##0.00"), "")
        cbbContaContabil.ListIndex = -1
        cbbContaCorrente.ListIndex = -1
        cbbLoja.ListIndex = -1
        txbHistorico.Text = Empty
             
    End If

End Sub
Private Sub lstPrincipalPopular(Pagina As Long)

    Dim n           As Byte
    Dim vNascimento As Variant
    Dim vSalario    As Variant
    Dim oLegenda    As control
    Dim cValor      As Currency
    
    ' Limpa cores da legenda
    For n = 1 To myRst.PageSize
        Set oLegenda = Controls("l" & Format(n, "00")): oLegenda.BackColor = &H8000000F
    Next n

    ' Define página que será exibida do Recordset
    myRst.AbsolutePage = Pagina
    
    With lstPrincipal
        .Clear                                      ' Limpa conteúdo
        .ColumnCount = 4                            ' Define número de colunas
        .ColumnWidths = "180 pt; 0pt; 55pt; 60pt;"  ' Configura largura das colunas
        .Font = "Consolas"                          ' Configura fonte
        
        n = 1
        
        While Not myRst.EOF = True And n <= myRst.PageSize
            
            ' Preenche ListBox
            .AddItem
            
            .List(.ListCount - 1, 0) = myRst.Fields("id").Value
            .List(.ListCount - 1, 1) = myRst.Fields("data").Value
            
            cValor = myRst.Fields("valor").Value
            
            .List(.ListCount - 1, 2) = Space(12 - Len(Format(cValor, "#,##0.00"))) & Format(cValor, "#,##0.00")
            
            ' Colore a legenda
            Set oLegenda = Controls("l" & Format(n, "00"))
            
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
        
            If txbData.Text = Empty Then
                MsgBox "Campo 'Data' é obrigatório", vbCritical: MultiPage1.Value = 1: txbData.SetFocus
            ElseIf txbValor.Text = Empty Then
                MsgBox "Campo 'Valor' é obrigatório", vbCritical: MultiPage1.Value = 1: txbValor.SetFocus
            ElseIf cbbContaContabil.ListIndex = -1 Then
                MsgBox "Campo 'Conta contábil' é obrigatório", vbCritical: MultiPage1.Value = 1: cbbContaContabil.SetFocus
            ElseIf cbbContaCorrente.ListIndex = -1 Then
                MsgBox "Campo 'Conta corrente' é obrigatório", vbCritical: MultiPage1.Value = 1: cbbContaCorrente.SetFocus
            ElseIf cbbLoja.ListIndex = -1 Then
                MsgBox "Campo 'Loja' é obrigatório", vbCritical: MultiPage1.Value = 1: cbbLoja.SetFocus
            Else
                
                With oLancamento
                    
                    .Data = CDate(txbData.Text)
                    .Valor = CCur(txbValor.Text)
                    .ContaContabilID = CLng(cbbContaContabil.List(cbbContaContabil.ListIndex, 1))
                    .ContaCorrenteID = CLng(cbbContaCorrente.List(cbbContaCorrente.ListIndex, 1))
                    .LojaID = CLng(cbbLoja.List(cbbLoja.ListIndex, 1))
                    .Historico = txbHistorico.Text
                    
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
        
            oLancamento.CRUD eCrud.Delete, oLancamento.ID
                
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

    Set myRst = oLancamento.Todos(Ordem)
    
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

    Dim r As New ADODB.Recordset
    
    Set r = oLancamento.ContasContabeis
    
    With cbbContaContabil
    
        .Clear
        .ColumnCount = 2
        .ColumnWidths = "60pt; 0pt;"
        
        While Not r.EOF = True
        
            .AddItem
            .List(.ListCount - 1, 0) = r.Fields("grupo").Value & " : " & r.Fields("subgrupo").Value & " : " & r.Fields("conta").Value
            .List(.ListCount - 1, 1) = r.Fields("id").Value
            
            r.MoveNext
        
        Wend
        
    End With
    
    Set r = oContaCorrente.Todos(, False)
    
    With cbbContaCorrente
    
        .Clear
        .ColumnCount = 2
        .ColumnWidths = "60pt; 0pt;"
        
        While Not r.EOF = True
        
            .AddItem
            .List(.ListCount - 1, 0) = r.Fields("conta").Value
            .List(.ListCount - 1, 1) = r.Fields("id").Value
            
            r.MoveNext
        
        Wend
        
    End With
    
    Set r = oLoja.Todos(, False)
    
    With cbbLoja
    
        .Clear
        .ColumnCount = 2
        .ColumnWidths = "60pt; 0pt;"
        
        While Not r.EOF = True
        
            .AddItem
            .List(.ListCount - 1, 0) = r.Fields("loja").Value
            .List(.ListCount - 1, 1) = r.Fields("id").Value
            
            r.MoveNext
        
        Wend
        
    End With
    
    Set r = Nothing

End Sub
Private Sub lstPrincipal_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    MultiPage1.Value = 1
    
End Sub
'Private Sub lblHdNome_Click()
'
'    Call BuscaRegistros("nome")
'
'End Sub
'Private Sub lblHdNascimento_Click()
'
'    Call BuscaRegistros("nascimento")
'
'End Sub
'Private Sub lblHdSalario_Click()
'
'    Call BuscaRegistros("salario")
'
'End Sub
