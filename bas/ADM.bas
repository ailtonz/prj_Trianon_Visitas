Attribute VB_Name = "ADM"
Option Compare Database
Option Explicit

Public strSQL As String

Public Function ValorCal()
On Error GoTo ValorCal_Err
   
      ' Testa se o form está aberto e em modo formulário
   If EstaAberto("Calendário") And IsFormView(Forms!Calendário) Then
      ' Captura o valor atual do calendário
      ValorCal = Forms!Calendário!Cal.value
   Else
      ValorCal = Now
   End If
   
   
ValorCal_Fim:
   Exit Function
ValorCal_Err:
   MsgBox Err.Description
   Resume ValorCal_Fim:
End Function


Public Function VerificarCadastro(codRoteiro, codObra, Viagem, codLigacao) As Boolean
Dim rDados As DAO.Recordset

Set rDados = CurrentDb.OpenRecordset("Select * from RoteirosItens where codRoteiro = " & codRoteiro & " and codObra = " & codObra & " and Viagem = " & Viagem & " and codLigacao = " & codLigacao & "")

If rDados.EOF Then 'Ñ achou o registro
    VerificarCadastro = False
Else
    VerificarCadastro = True
End If

rDados.Close

Set rDados = Nothing

End Function


Public Function CadastrarViagem(TipoCadastro As String, codRoteiro, Viagem, Ordem, Obra, Cliente, codObra, codCadastro, Trabalho, dtTrabalho, OBS, codLigacao)

Select Case TipoCadastro

    Case "Coloca"
    
        strSQL = "INSERT INTO RoteirosItens ( codRoteiro, Viagem, Ordem, Obra, Cliente, codObra, codCadastro, C, DT_C, OBS, codLigacao ) " & _
                 "values (" & codRoteiro & ",'" & Viagem & "','" & Ordem & "','" & Obra & "','" & Cliente & "'," & codObra & "," & codCadastro & "," & Trabalho & ",'" & dtTrabalho & "','" & OBS & "','" & codLigacao & "')"

    Case "Retira"
    
        strSQL = "INSERT INTO RoteirosItens ( codRoteiro, Viagem, Ordem, Obra, Cliente, codObra, codCadastro, R, DT_R, OBS, codLigacao ) " & _
                 "values (" & codRoteiro & ",'" & Viagem & "','" & Ordem & "','" & Obra & "','" & Cliente & "'," & codObra & "," & codCadastro & "," & Trabalho & ",'" & dtTrabalho & "','" & OBS & "','" & codLigacao & "')"
    
    Case "Troca"
    
        strSQL = "INSERT INTO RoteirosItens ( codRoteiro, Viagem, Ordem, Obra, Cliente, codObra, codCadastro, T, DT_T, OBS, codLigacao ) " & _
                 "values (" & codRoteiro & ",'" & Viagem & "','" & Ordem & "','" & Obra & "','" & Cliente & "'," & codObra & "," & codCadastro & "," & Trabalho & ",'" & dtTrabalho & "','" & OBS & "','" & codLigacao & "')"

End Select

ExecutarSQL strSQL

End Function


Public Function AtualizarViagem(TipoCadastro As String, codRoteiro, Viagem, codObra, Trabalho, dtTrabalho, OBS, codLigacao)
Dim strSQL As String

Select Case TipoCadastro

    Case "Coloca"

        strSQL = "UPDATE RoteirosItens SET RoteirosItens.C = " & Trabalho & ", RoteirosItens.DT_C = '" & dtTrabalho & "', RoteirosItens.OBS = '" & OBS & "', RoteirosItens.codLigacao = '" & codLigacao & "'" & _
                 " WHERE ((RoteirosItens.codRoteiro)=" & codRoteiro & ") AND ((RoteirosItens.Viagem)=" & Viagem & ") AND ((RoteirosItens.codObra)=" & codObra & ") AND ((RoteirosItens.codLigacao)=" & codLigacao & ")"
    
    Case "Retira"

        strSQL = "UPDATE RoteirosItens SET RoteirosItens.R = " & Trabalho & ", RoteirosItens.DT_R = '" & dtTrabalho & "', RoteirosItens.OBS = '" & OBS & "', RoteirosItens.codLigacao = '" & codLigacao & "'" & _
                 " WHERE ((RoteirosItens.codRoteiro)=" & codRoteiro & ") AND ((RoteirosItens.Viagem)=" & Viagem & ") AND ((RoteirosItens.codObra)=" & codObra & ") AND ((RoteirosItens.codLigacao)=" & codLigacao & ")"
    
    Case "Troca"

        strSQL = "UPDATE RoteirosItens SET RoteirosItens.T = " & Trabalho & ", RoteirosItens.DT_T = '" & dtTrabalho & "', RoteirosItens.OBS = '" & OBS & "', RoteirosItens.codLigacao = '" & codLigacao & "'" & _
                 " WHERE ((RoteirosItens.codRoteiro)=" & codRoteiro & ") AND ((RoteirosItens.Viagem)=" & Viagem & ") AND ((RoteirosItens.codObra)=" & codObra & ") AND ((RoteirosItens.codLigacao)=" & codLigacao & ")"
End Select

ExecutarSQL strSQL

End Function


