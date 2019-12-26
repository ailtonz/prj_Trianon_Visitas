Attribute VB_Name = "modPeriodosDeTempos"
Option Compare Database

Public Function CalcularVencimento(Dia As Integer, Optional MES As Integer, Optional ANO As Integer) As Date

If Month(Now) = 2 Then
    If Dia = 29 Or Dia = 30 Or Dia = 31 Then
        Dia = 1
        MES = MES + 1
    End If
End If

If MES > 0 And ANO > 0 Then
    CalcularVencimento = Format((DateSerial(ANO, MES, Dia)), "dd/mm/yyyy")
ElseIf MES = 0 And ANO > 0 Then
    CalcularVencimento = Format((DateSerial(ANO, Month(Now), Dia)), "dd/mm/yyyy")
ElseIf MES = 0 And ANO = 0 Then
    CalcularVencimento = Format((DateSerial(Year(Now), Month(Now), Dia)), "dd/mm/yyyy")
End If

End Function
