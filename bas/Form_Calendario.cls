VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Calendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim blnInicio As Boolean   ' indica inicialização do form
Public StringF2 As String

Private Sub Cal_AfterUpdate()

  ' Atualiza cx. texto DataLonga
  Me!DataAuxiliar.Requery
  ' Atualiza data na Agenda
  Me!lstLigacoes.Requery
  
End Sub

Private Sub Cal_NewMonth()
   ' Calendario não deve aceitar datas nulas (sem data)
   Cal.ValueIsNull = False
   
   ' Põe o foco na caixa de texto DataLonga.
   ' Veja o comentário no procedimento
   ' DataAuxiliar_GotFocus.
   If Not blnInicio Then
      Me!DataAuxiliar.SetFocus
   End If
End Sub

Private Sub Cal_NewYear()
   ' Calendario não deve aceitar datas nulas (sem data)
   Cal.ValueIsNull = False
   
   ' Põe o foco na caixa de texto DataAuxiliar.
   ' Veja o comentário no procedimento
   ' DataAuxiliar_GotFocus.
   
   ' O If evita um erro quando o form está sendo
   ' carregado. Nesse momento, o controle ainda
   ' não pode receber o foco. A variável blnInicio
   ' indica que é o momento de abertura do form.
   If Not blnInicio Then
      Me!DataAuxiliar.SetFocus
   End If
End Sub

Private Sub AtualizaAgenda()
On Error GoTo Atualiza_Err

  ' Atualiza cx. texto DataLonga
  Me!DataAuxiliar.Requery
  ' Atualiza data na Agenda
  Me!lstLigacoes.Requery


Atualiza_Fim:
    Exit Sub
Atualiza_Err:
    MsgBox Err.Description
    Resume Atualiza_Fim
End Sub


Private Sub cmdHoje_Click()
On Error GoTo Err_cmdHoje_Click
Dim sqlLigacoes As String: sqlLigacoes = "SELECT qryProximoContato.codLigacao, qryProximoContato.Cliente, qryProximoContato.['Primeiro_Contato'], qryProximoContato.['Data_Evento'], qryProximoContato.Status,qryProximoContato.eMail  FROM qryProximoContato; UNION SELECT qryLigacoesDia.codLigacao,  qryLigacoesDia.Cliente, qryLigacoesDia.['Primeiro_Contato'], qryLigacoesDia.['Data_Evento'], qryLigacoesDia.Status, qryLigacoesDia.eMail FROM qryLigacoesDia;"
    
    ' Calendario: hoje
    Cal.Today
    ' Agenda: hoje
'    AtualizaAgenda
    
    Me.lstLigacoes.RowSource = sqlLigacoes
    Me.lstLigacoes.ColumnWidths = "0cm;7cm;3cm;3cm;3cm;3cm"
    Me.lstLigacoes.Requery
    
Exit_cmdHoje_Click:
    Exit Sub
Err_cmdHoje_Click:
    MsgBox Err.Description
    Resume Exit_cmdHoje_Click
End Sub

Private Sub cmdImprimir_Click()
On Error GoTo cmdImprimir_Err

   ' Abre o form Imprimir como cx. de diálogo
   DoCmd.OpenForm "Imprimir", , , , , acDialog
   
cmdImprimir_Fim:
    Exit Sub
cmdImprimir_Err:
    MsgBox Err.Description
    Resume cmdImprimir_Fim
End Sub

Private Sub DataAuxiliar_GotFocus()
   ' Ao receber o foco, DataAuxiliar provoca a atualização
   ' do subform Agenda. Esta solução foi adotada porque o
   ' o objeto Calendario trava as caixas de combinação de
   ' mês e ano em janeiro e em 1900 se o método Agenda.Requery
   ' for chamado nos eventos Cal_NewMonth e Cal_NewYear,
   ' associados à escolha de mês ou ano nessas caixas.
   ' Esse método poderia ser aplicado à caixa DataLonga, mas
   ' para isso ela precisaria ficar como caixa de texto ativa,
   ' o que não faz sentido. Por isso criou-se o controle
   ' adicional DataAuxiliar, que fica ativo, mas tem tamanho
   ' bastante reduzido. Obs: DataAuxiliar não pode ser invisível,
   ' porque assim não tem condições de receber o foco.
   Me!lstLigacoes.Requery
'   Me!lstAtrasos.Requery
End Sub

Private Sub Form_Load()
Dim blRet As Boolean

    ' Since we are not passing a filename of a Bitmap file
    ' the standard Window File Dialog will popup.
'    blRet = SetMDIBackGroundImage()
    
    blnInicio = True
'    'Calendario: hoje
    Cal.Today
    Cal.ValueIsNull = False
    blnInicio = False
    
    strTabela = "4"
    StringF2 = ""
'    Filtro strTabela
    Me.KeyPreview = True
    Me.lstLigacoes.SetFocus
    Me.lstLigacoes.Selected(1) = True
'    DoCmd.Maximize

'Call cmdHoje_Click

End Sub

Private Sub cmdFiltrar_Click()

    Dim txtFiltro As String
    txtFiltro = InputBox("Digite uma palavra para fazer o filtro:", "Filtro", StringF2, 0, 0)
    StringF2 = txtFiltro
    Filtro "4", txtFiltro
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    
        Case vbKeyInsert
        
           cmdNovo_Click
           
        Case vbKeyReturn
        
            cmdAlterar_Click
            
        Case vbKeyDelete
           
            cmdExcluir_Click
        
        Case vbKeyF2
        
            cmdFiltrar_Click
            
    End Select
End Sub

Private Sub cmdNovo_Click()

    Manipulacao "4", "Novo"
    
End Sub

Private Sub cmdAlterar_Click()

    Manipulacao "4", "Alterar"
    
End Sub

Private Sub cmdExcluir_Click()

    Manipulacao "4", "Excluir"
    
End Sub

Private Sub lstEventos_DblClick(Cancel As Integer)
Dim rstFormularios As DAO.Recordset
Dim strSQL As String

strSQL = "SELECT First(Categorias.Descricao01) AS NomeDoFormulario, Categorias.Descricao02 AS Legenda, Categorias.Descricao03 AS Identificacao FROM Categorias GROUP BY Categorias.Descricao02, Categorias.Descricao03, Categorias.Titulo HAVING (((Categorias.Titulo)='TipoSistema')) ORDER BY First(Categorias.Descricao01)"

Set rstFormularios = CurrentDb.OpenRecordset(strSQL)

DoCmd.OpenForm rstFormularios.Fields("NomeDoFormulario"), , , rstFormularios.Fields("Identificacao") & " = " & Form_Calendario.lstEventos.value

Set rstFormularios = Nothing

End Sub

Private Sub lstLigacoes_DblClick(Cancel As Integer)

    cmdAlterar_Click
    
End Sub

Private Sub cmdFechar_Click()
On Error GoTo Err_cmdFechar_Click

    DoCmd.Close

Exit_cmdFechar_Click:
    Exit Sub

Err_cmdFechar_Click:
    MsgBox Err.Description
    Resume Exit_cmdFechar_Click
    
End Sub

Public Sub Manipulacao(Tabela As String, Operacao As String)
Dim rstFormularios As DAO.Recordset

Dim strSQL As String

strSQL = "SELECT First(Categorias.Descricao01) AS NomeDoFormulario, Categorias.Descricao02 AS Legenda, Categorias.Descricao03 AS Identificacao FROM Categorias GROUP BY Categorias.Descricao02, Categorias.Descricao03, Categorias.Titulo HAVING (((Categorias.Titulo)='TipoSistema')) ORDER BY First(Categorias.Descricao01)"


If IsNull(Form_Calendario.lstLigacoes.value) And Operacao <> "Novo" Then
   Exit Sub
End If


Set rstFormularios = _
    CurrentDb.OpenRecordset(strSQL)
        
    
Select Case Operacao

 Case "Novo"
        
    DoCmd.OpenForm rstFormularios.Fields("NomeDoFormulario"), , , , acFormAdd
    
 Case "Alterar"

    DoCmd.OpenForm rstFormularios.Fields("NomeDoFormulario"), , , rstFormularios.Fields("Identificacao") & " = " & Form_Calendario.lstLigacoes.value

 Case "Excluir"

    If MsgBox("ATENÇÃO: Você deseja realmente excluir este registro ???", vbQuestion + vbOKCancel) = vbOK Then
       DoCmd.SetWarnings False
       DoCmd.RunSQL ("Delete from Ligacoes where codLigacao = " & Form_Calendario.lstLigacoes.value)
       DoCmd.SetWarnings True
    End If
    
'    Me.lstAtrasos.Requery

End Select

Form_Calendario.lstLigacoes.Requery

Saida:

End Sub

Private Function Filtro(strTabela As String, Optional Procurar As String)

Dim rstFormularios As DAO.Recordset
Dim rstForm_Campos As DAO.Recordset
Dim rstForm_TabRelacionada As DAO.Recordset
Dim rstResultado As DAO.Recordset

Dim Sql As String
Dim SqlAux As String
Dim Contagem As Integer
Dim a, b, c As Integer
Dim Colunas As Integer

Dim Procuras(30) As String
Dim ProcurasAux As Integer
ProcurasAux = 1

For b = 1 To Len(Procurar)
   If Mid(Procurar, b, 1) = "+" Then
      ProcurasAux = ProcurasAux + 1
   Else
      Procuras(ProcurasAux) = Procuras(ProcurasAux) + Mid(Procurar, b, 1)
   End If
Next b

Set rstFormularios = _
    CurrentDb.OpenRecordset("Select * from Formularios " & _
                            " where codFormulario = " & _
                            strTabela & "")

Set rstForm_Campos = _
    CurrentDb.OpenRecordset("Select * from Formularios_Campos " & _
                            " where codFormulario = " & _
                            strTabela)

Set rstForm_TabRelacionada = _
    CurrentDb.OpenRecordset("Select * from Formularios_TabelaRelacionada " & _
                            " where codFormulario = " & _
                            strTabela)
Sql = "Select "

While Not rstForm_Campos.EOF
    If rstForm_Campos.Fields("Pesquisa") = True Then
        Sql = Sql & IIf(IsNull(rstForm_Campos.Fields("Nome")), _
                      rstForm_Campos.Fields("Campo"), _
                      rstForm_Campos.Fields("Campo") & _
                      " AS " & rstForm_Campos.Fields("Nome")) & ", "
    End If

    rstForm_Campos.MoveNext
Wend

Sql = Left(Sql, Len(Sql) - 2) & " "

Sql = Sql & " from "

If Not rstForm_TabRelacionada.EOF Then

    SqlAux = ""
    Contagem = 1
    rstForm_TabRelacionada.MoveFirst

    While Not rstForm_TabRelacionada.EOF

      SqlAux = "(" & SqlAux & IIf(Contagem <> 1, "", rstFormularios.Fields("TabelaPrincipal")) & " Left Join " & _
               rstForm_TabRelacionada.Fields("TabelaRelacionada") & " ON " & _
               rstFormularios.Fields("TabelaPrincipal") & "." & rstForm_TabRelacionada.Fields("CampoChave_Pai") & " = " & _
               rstForm_TabRelacionada.Fields("TabelaRelacionada") & "." & rstForm_TabRelacionada.Fields("CampoChave_Filho") & ")"

      rstForm_TabRelacionada.MoveNext
      Contagem = Contagem + 1

    Wend

    If SqlAux <> "" Then
       Sql = Sql & SqlAux
    End If

End If

If SqlAux = "" Then
   Sql = Sql & "" & rstFormularios.Fields("TabelaPrincipal") & " Where ( "
'Else
'   Sql = Sql & " Where ("
End If

rstForm_Campos.MoveFirst

For c = 1 To ProcurasAux

   rstForm_Campos.MoveFirst
   Sql = Sql & " ( "
   While Not rstForm_Campos.EOF
     If rstForm_Campos.Fields("Filtro") = True Then
        Sql = Sql & rstForm_Campos.Fields("Campo") & " Like '*" _
                  & LCase(Trim(Procuras(c))) & "*' OR "
     End If
     rstForm_Campos.MoveNext
   Wend
   Sql = Left(Sql, Len(Sql) - 3) & ") "
   If c <> ProcurasAux Then
      Sql = Sql + " And "
   End If

Next c

Sql = Sql + " ) "

Sql = Sql & "Order By "

rstForm_Campos.MoveFirst

While Not rstForm_Campos.EOF

  If rstForm_Campos.Fields("Ordem") <> "" Then
     Sql = Sql & rstForm_Campos.Fields("Campo") _
               & " " & rstForm_Campos.Fields("Ordem") & ", "
  End If

  rstForm_Campos.MoveNext

Wend

Sql = Left(Sql, Len(Sql) - 2) & " "

Sql = Sql & ";"


'GerarSaida Sql, "Sql.Log"

Me.lstLigacoes.RowSource = Sql
Me.lstLigacoes.ColumnHeads = True
Me.lstLigacoes.ColumnCount = rstForm_Campos.RecordCount
Me.Caption = rstFormularios.Fields("TituloDoFormulario")

Dim strTamanho As String

rstForm_Campos.MoveFirst
While Not rstForm_Campos.EOF
  If Not IsNull(rstForm_Campos.Fields("Tamanho")) Then
     strTamanho = strTamanho & Str(rstForm_Campos.Fields("Tamanho")) & "cm;"
  End If
  rstForm_Campos.MoveNext
Wend

Me.lstLigacoes.ColumnWidths = strTamanho

'If IsNull(rstFormularios.Fields("campodesoma")) = True Then
'
'   Me.lblSoma.Caption = "Qtd: " & Me.lstLigacoes.ListCount - 1
'
'Else
'
'   Dim Soma
'   Set rstResultado = CurrentDb.OpenRecordset(Sql)
'   Do While Not rstResultado.EOF
'      Soma = Soma + rstResultado.Fields(rstFormularios.Fields("campodesoma"))
'      rstResultado.MoveNext
'   Loop
'   Me.lblSoma.Caption = "Qtd: " & Me.lstLigacoes.ListCount - 1 & "   Soma: " & FormatNumber(Soma, 2)
'   rstResultado.Close
'
'End If


rstFormularios.Close
rstForm_Campos.Close
rstForm_TabRelacionada.Close



End Function




