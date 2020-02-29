Attribute VB_Name = "m_Calendar"
Option Private Module
Option Explicit

Public Const sMascaraData   As String = "DD/MM/YYYY"   ' M�scara de formata��o de datas
Public dtDate               As Date
Dim aBotoes()               As New c_Calendario  ' Vetor que armazena todos os bot�es de dia do Calend�rio

Public Function GetCalendario() As Date
    ' Fun��o GetCalendario
    
    ' Declara vari�veis
    Dim lTotalBotoes As Long   ' Total de r�tulos
    Dim Ctrl As control
    Dim frm As New f_Calendario      ' Formul�rio
    
    'Set frm = New fCalendario ' Cria novo objeto setando formul�rio nele
    
    ' Atribui cada um dos bot�es em um elemento do vetor da classe
    For Each Ctrl In frm.Controls
        If Ctrl.name Like "l?c?" Then
            lTotalBotoes = lTotalBotoes + 1
            ReDim Preserve aBotoes(1 To lTotalBotoes)
            Set aBotoes(lTotalBotoes).btnGrupo = Ctrl
        End If
    Next Ctrl
    
    frm.Show
    
    ' Se a data escolhida for nula ou inv�lida, retorna-se a data atual:
    If IsDate(frm.Tag) Then
        GetCalendario = frm.Tag
    Else
        GetCalendario = dtDate
    End If
    
    Unload frm
End Function
    

