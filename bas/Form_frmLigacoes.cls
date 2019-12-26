VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmLigacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cboeMail_Click()
    Me.Usuarios = Me.cboeMail.Column(1)
End Sub

Private Sub Cliente_Enter()
    Me.Cliente.Requery
End Sub

Private Sub Form_Load()
'    DoCmd.Maximize

    'Data de Agendamento
    If (IsNull(Me.Data)) Then Me.Data = Form_Calendario.Cal.value
    

End Sub

Private Sub cmdSalvar_Click()
On Error GoTo Err_cmdSalvar_Click

'        'Data de Agendamento
'        Me.Data = Form_Calendario.Cal.value
        
        DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
        Form_Calendario.lstLigacoes.Requery
        DoCmd.Close

Exit_cmdSalvar_Click:
    Exit Sub

Err_cmdSalvar_Click:
    If Not (Err.Number = 2046 Or Err.Number = 0) Then MsgBox Err.Description
    DoCmd.Close
    Resume Exit_cmdSalvar_Click
End Sub

Private Sub cmdFechar_Click()
On Error GoTo Err_cmdFechar_Click

    DoCmd.DoMenuItem acFormBar, acEditMenu, acUndo, , acMenuVer70
    DoCmd.CancelEvent
    DoCmd.Close

Exit_cmdFechar_Click:
    Exit Sub

Err_cmdFechar_Click:
    If Not (Err.Number = 2046 Or Err.Number = 0) Then MsgBox Err.Description
    DoCmd.Close
    Resume Exit_cmdFechar_Click

End Sub

Private Sub cmdNovoCliente_Click()
On Error GoTo Err_cmdNovoCliente_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "CadastroCliente"
    DoCmd.OpenForm stDocName, , , stLinkCriteria, acFormAdd

Exit_cmdNovoCliente_Click:
    Exit Sub

Err_cmdNovoCliente_Click:
    MsgBox Err.Description
    Resume Exit_cmdNovoCliente_Click
    
End Sub
Private Sub cmdOrcamento_Click()
On Error GoTo Err_cmdOrcamento_Click

    Dim stDocName As String

    stDocName = "rptAgendaCadastro"
    DoCmd.OpenReport stDocName, acPreview, , "codLigacao = " & Me.codLigacao

Exit_cmdOrcamento_Click:
    Exit Sub

Err_cmdOrcamento_Click:
    MsgBox Err.Description
    Resume Exit_cmdOrcamento_Click
    
End Sub
