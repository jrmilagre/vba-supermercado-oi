Attribute VB_Name = "m_Calculadora"
Option Private Module
Option Explicit

Public ccurVisor            As Currency
Dim aBotoes()               As New c_Calculadora  ' Vetor que armazena todos os bot�es de dia do Calend�rio

Public Function GetCalculadora() As Double
    
    ' Declara vari�veis
    Dim lTotalBotoes As Long   ' Total de r�tulos
    Dim Ctrl As control
    Dim frm As New f_Calculadora      ' Formul�rio
    
    ' Atribui cada um dos bot�es em um elemento do vetor da classe
    For Each Ctrl In frm.Controls
        If Ctrl.name Like "btnUsar" Then
            lTotalBotoes = lTotalBotoes + 1
            ReDim Preserve aBotoes(1 To lTotalBotoes)
            Set aBotoes(lTotalBotoes).btnGrupo = Ctrl
        End If
    Next Ctrl
    
    frm.Show
    
    ' Se a data escolhida for nula ou inv�lida, retorna-se a data atual:
    If IsNumeric(CDbl(frm.txbVisor.Text)) Then
        GetCalculadora = CDbl(frm.txbVisor.Text)
    Else
        GetCalculadora = ccurVisor
    End If
    
    Unload frm
End Function
