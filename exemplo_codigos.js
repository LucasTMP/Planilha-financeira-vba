Private Sub Btcadastrar_Click()


'''''''''''''''''''''''''''''''' variaveis

Dim valorservico As Single

Dim i As Integer
'Dim DataNow As Date

i = 12
valorservico = 0

'DataNow = Sheets("Parâmetros").Range("D2").Value


'''''''''''''''''''''''''''''''' Tratamento de erros


If txtnome.Value = "" Or Len(Trim(txtnome.Value)) = 0 Then
MsgBox "Não deixe o campo NOME em branco!", vbOKOnly + vbInformation, "Dados Faltando"
txtnome.Value = ""
txtnome.SetFocus
Exit Sub
End If

If Range("Tabela6").Find(txtnome.Value, Lookat:=xlWhole) Is Nothing Then
Else
MsgBox "Nome já cadastrado!", vbOKOnly + vbInformation, "Dados duplicados"
Exit Sub
End If

If modbi = False And modse = False And modan = False Then
MsgBox "Selecione uma MODALIDADE", vbOKOnly + vbInformation, "Dados Faltando"
Exit Sub
End If

If aplica_desconto.Value = True Then
If txtdesconto.Text = "" Or txtdesconto.Value = 0 Then
MsgBox "O campo de desconto não pode ser Nulo", vbOKOnly + vbInformation, "Dados Incorretos"
Exit Sub
End If
End If

If txtdate.Text = "" Then
MsgBox "Preencha o campo de DATA DE ÍNICIO", vbOKOnly + vbInformation, "Dados Faltando"
txtdate.SetFocus
Exit Sub
End If

If txtdate.Text Like "??[/-]??[/-]????" Then
Else
MsgBox "Data invalida, formato de data: dd/mm/yyyy", vbOKOnly + vbInformation, "Dados incorretos"
Exit Sub
End If

If Not IsDate(txtdate) Then
MsgBox "Data Inválida!", vbOKOnly + vbInformation, "Dados Faltando"
txtdate.SetFocus
Exit Sub
End If


'If datarelatoriopri = False And datarelatorioseg = False And datarelatorioter = False And datarelatorioqua = False Then
'MsgBox "Selecione uma DATA DE RELATORIO!", vbOKOnly + vbInformation, "Dados Faltando"
'Exit Sub
'End If

If formapagamentocartao = False And formapagamentodinheiro = False And formapagamentocheque = False And formapagamentopix = False Then
MsgBox "Selecione uma FORMA DE PAGAMENTO!", vbOKOnly + vbInformation, "Dados Faltando"
Exit Sub
End If


If parcelas.Value = "" Then
MsgBox "Selecione uma quantidade de PARCELAS!", vbOKOnly + vbInformation, "Dados Faltando"
Exit Sub
End If

''''''''''''''''''''''''''''''''' posição

Do While Cells(i, 3) <> ""
  i = i + 1
Loop

'''''''''''''''''''''''''''''''' definindo valores

Worksheets("Cadastros").Cells(i, 3) = txtnome

Worksheets("Cadastros").Cells(i, 4) = txtdate

'Worksheets("Cadastros").Cells(i, 4) = Format(Date, "mm/dd")


If modbi = True Then
Worksheets("Cadastros").Cells(i, 6) = "Bienal"
valorservico = Worksheets("Parâmetros").Cells(4, 1)
End If

If modse = True Then
Worksheets("Cadastros").Cells(i, 6) = "Semestral"
valorservico = Worksheets("Parâmetros").Cells(2, 1)
End If

If modan = True Then
Worksheets("Cadastros").Cells(i, 6) = "Anual"
valorservico = Worksheets("Parâmetros").Cells(3, 1)
End If





'If datarelatoriopri = True Then
'Worksheets("Cadastros").Cells(i, 7) = "1» Semana"
'End If

'If datarelatorioseg = True Then
'Worksheets("Cadastros").Cells(i, 7) = "2» Semana"
'End If

'If datarelatorioter = True Then
'Worksheets("Cadastros").Cells(i, 7) = "3» Semana"
'End If

'If datarelatorioqua = True Then
'Worksheets("Cadastros").Cells(i, 7) = "4» Semana"
'End If



If formapagamentocartao = True Then
Worksheets("Financeiro").Cells(i, 9) = "Cartão"
End If

If formapagamentodinheiro = True Then
Worksheets("Financeiro").Cells(i, 9) = "Dinheiro"
End If

If formapagamentocheque = True Then
Worksheets("Financeiro").Cells(i, 9) = "Cheque"
End If

If formapagamentopix = True Then
Worksheets("Financeiro").Cells(i, 9) = "Pix"
End If


If aplica_desconto.Value = True Then
Worksheets("Financeiro").Cells(i, 6) = txtdesconto.Value / 100
Else
Worksheets("Financeiro").Cells(i, 6) = 0
End If


Worksheets("Financeiro").Cells(i, 10) = parcelas.Value
Worksheets("Financeiro").Cells(i, 3) = txtnome
Worksheets("Financeiro").Cells(i, 5) = valorservico
Worksheets("Financeiro").Cells(i, 11) = 1
Worksheets("Financeiro").Cells(i, 12) = "Pendente"
Worksheets("Financeiro").Cells(i, 13) = "Em Andamento"
' Worksheets("Financeiro").Cells(i, 17) = DataNow



'''''''''''''''''''''''''''''' limpando campos

txtnome.Text = ""
txtdate.Text = ""
modbi = False
modse = False
modan = False
datarelatoriopri = False
datarelatorioseg = False
datarelatorioter = False
datarelatorioqua = False
formapagamentocartao = False
formapagamentodinheiro = False
formapagamentocheque = False
formapagamentopix = False
parcelas.Value = ""


valorservico = 0
Filtro
Unload Me





End Sub


Private Sub aplica_desconto_Click()
If aplica_desconto.Value = True Then
Label6.Visible = True
txtdesconto.Visible = True
Else
Label6.Visible = False
txtdesconto.Visible = False
End If
End Sub


Private Sub txtdate_AfterUpdate()

If txtdate.Text Like "??[/-]??[/-]????" Then
Else
MsgBox "Data invalida, formato de data: dd/mm/yyyy", vbOKOnly + vbInformation, "Dados incorretos"
End If

End Sub


Private Sub txtdate_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

txtdate.Text = formatadata.DATA(KeyAscii, txtdate.Text)

End Sub

Private Sub txtdesconto_Change()

If Not IsNumeric(txtdesconto.Value) Then

MsgBox "Permitido apenas numeros", vbOKOnly + vbInformation, "Dados incorretos"

End If

End Sub


Private Sub UserForm_Initialize()

With parcelas

.AddItem 1
.AddItem 2
.AddItem 3
.AddItem 4
.AddItem 5
.AddItem 6
.AddItem 7
.AddItem 8
.AddItem 9
.AddItem 10
.AddItem 11
.AddItem 12
.AddItem 13
.AddItem 14
.AddItem 15
.AddItem 16
.AddItem 17
.AddItem 18
.AddItem 19
.AddItem 20
.AddItem 21
.AddItem 22
.AddItem 23
.AddItem 24

End With

parcelas.ListIndex = 0

txtdate.Text = Format(Date, "dd/mm/yyyy")

End Sub
