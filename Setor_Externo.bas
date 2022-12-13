Attribute VB_Name = "Setor_Externo"
Option Explicit
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" ( _
    ByVal pCaller As LongPtr, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As LongPtr, _
    ByVal lpfnCB As LongPtr) As LongPtr
Sub BP()

Dim Setor_Externo As Workbook
Dim Dado As Workbook
Dim FileURL As String
Dim DestinationFile As String
Dim coluna As Integer
Dim nome As String

'nome = VBA.Interaction.Environ$("UserName")

FileURL = "https://www.bcb.gov.br/content/estatisticas/Documents/Tabelas_especiais/BalPagM.xlsx"

DestinationFile = "C:\Users\e-vinicius.geronimo\CNI - Confederação Nacional da Indústria\ECON - 2 Informe Conjuntural\Setor Externo\Base de Dados - Setor Externo\0 - Antigo\BP mensal\BalPagM" & Format(Date, "ddmmmyy") & ".xlsx"

If URLDownloadToFile(0, FileURL, DestinationFile, 0, 0) = 0 Then
    Debug.Print "File download started"
Else
    Debug.Print "File download not started"
End If

Set Setor_Externo = ActiveWorkbook

Workbooks.Open (DestinationFile)

Set Dado = ActiveWorkbook

Dado.Activate
Dado.Sheets("Balanço").Select
coluna = Range("B7").End(xlToRight).Column
Range(Cells(5, 1), Cells(389, coluna)).Copy

Setor_Externo.Activate
Sheets("BP mensal").Select
Range("A5").PasteSpecial xlPasteAll


Dado.Activate
Application.CutCopyMode = False
Dado.Close

Setor_Externo.Activate
Sheets("BP mensal").Select
Range("A2").Value = "Última Atualização: " & Now
Range(Cells(5, 1), Cells(389, coluna)).Copy

Sheets("Balanço mensal (bilhões)").Select
Range("A5").PasteSpecial xlPasteAll
Range(Cells(9, 2), Cells(389, coluna)).Clear

Range("B9").Select
ActiveCell.FormulaR1C1 = "='BP mensal'!RC/1000"
Range("B10").Select
ActiveCell.FormulaR1C1 = "='BP mensal'!RC/1000"

Range("B9:B10").Select
    Selection.NumberFormat = "0"
    Selection.AutoFill Destination:=Range(Cells(9, 2), Cells(389, 2)), Type:=xlFillDefault
    
Range("B9:B389").Select
    Selection.AutoFill Destination:=Range(Cells(9, 2), Cells(389, coluna)), Type:=xlFillDefault
    
Range("A2").Value = "Última Atualização: " & Now
Range("A3").Select

    
End Sub


Sub Balança_Semanal()

Dim Setor_Externo As Workbook 'Define a planilha do setor externo
Dim Dado As Workbook ' Define a planilha com o dado atualizado
Dim FileURL As String 'Link para baixar o dado
Dim DestinationFile As String 'Caminho na nuvem que download será realizado
Dim linha_dado As Integer ' ultima linha com dado da planilha com o dado atualizado
Dim linha_SE As Integer 'ultima linha com dado da planilha do setor externo
Dim nome As String 'nome do usuário (parte do e-mail da cni antes do @)

'captura o nome do usuário que será usado para definir o caminho na nuvem para o download
'nome = VBA.Interaction.Environ$("UserName")

'Define o link de onde o arquivo será baixado
FileURL = "https://balanca.economia.gov.br/balanca/semanal/Tabela_Resumo.xlsx"

'Define o caminho que o download será realizado
DestinationFile = "C:\Users\e-vinicius.geronimo\CNI - Confederação Nacional da Indústria\ECON - 2 Informe Conjuntural\Setor Externo\Base de Dados - Setor Externo\0 - Antigo\BC semanal\Tabela_Resumo" & Format(Date, "ddmmmyy") & ".xlsx"

'Baixa o dado no caminho definido
If URLDownloadToFile(0, FileURL, DestinationFile, 0, 0) = 0 Then
    Debug.Print "File download started"
Else
    Debug.Print "File download not started"
End If

'Chama a planilha de Setor Externo de Setor_Externo
Set Setor_Externo = ActiveWorkbook

'Abre o dado baixado
Workbooks.Open (DestinationFile)

'Chama a planilha com o dado de Dado
Set Dado = ActiveWorkbook

'Ativa a planilha dado
Dado.Activate
'Seleciona a aba com o dado
Sheets("Sheet 1").Select
'Ajusta a altura das linhas
Rows("14:23").RowHeight = 15
'Captura o numero da ultima linha com os dados semanais
linha_dado = Range("A16").End(xlDown).Row


'Copia e cola semana e dias uteis
'Copia
Range(Cells(16, 1), Cells(linha_dado, 2)).Copy
'Seleciona a planilha setor externo
Setor_Externo.Activate
'Seleciona a aba BC Semanal
Sheets("BC semanal").Select
'Captura o numero da ultima linha com dados
linha_SE = Range("B10").End(xlDown).Row
'Cola os dados copiados na linha abaixo da última
Cells(linha_SE + 1, 2).PasteSpecial xlPasteValues


'Copia e cola Exportação média p/ dia útil
'Ativa a planilha dado
Dado.Activate
'Seleciona a aba sheet 1
Sheets("Sheet 1").Select
'Copia
Range(Cells(16, 4), Cells(linha_dado, 4)).Copy
Setor_Externo.Activate
'Seleciona a aba BC Semanal
Sheets("BC semanal").Select
'Cola os dados copiados na linha abaixo da última
Cells(linha_SE + 1, 4).PasteSpecial xlPasteValues


'Copia e cola Importação média p/ dia útil
'Ativa a planilha dado
Dado.Activate
'Seleciona a aba sheet 1
Sheets("Sheet 1").Select
'Copia
Range(Cells(16, 6), Cells(linha_dado, 6)).Copy
'Seleciona a planilha setor externo
Setor_Externo.Activate
'Seleciona a aba BC Semanal
Sheets("BC semanal").Select
'Cola os dados copiados na linha abaixo da última
Cells(linha_SE + 1, 5).PasteSpecial xlPasteValues


'Copia e cola Saldo média p/ dia útil
'Ativa a planilha dado
Dado.Activate
'Seleciona a aba sheet 1
Sheets("Sheet 1").Select
'Copia
Range(Cells(16, 10), Cells(linha_dado, 10)).Copy
'Seleciona a planilha setor externo
Setor_Externo.Activate
'Seleciona a aba BC Semanal
Sheets("BC semanal").Select
'Cola os dados copiados na linha abaixo da última
Cells(linha_SE + 1, 6).PasteSpecial xlPasteValues


'Copia e cola mês
'Ativa a planilha dado
Dado.Activate
'Seleciona a aba sheet 1
Sheets("Sheet 1").Select
'Copia
Range("A14").Copy
'Seleciona a planilha setor externo
Setor_Externo.Activate
'Seleciona a aba BC Semanal
Sheets("BC semanal").Select
'Cola os dados copiados na linha abaixo da última
Cells(linha_SE + 1, 1).PasteSpecial xlPasteValues


'Ativa a planilha dado
Dado.Activate
'Desseleciona a parte copiada
Application.CutCopyMode = False
'Fecha a aba dados sem salvar as alterações
Dado.Close SaveChanges:=False
'Seleciona a planilha setor externo
Setor_Externo.Activate
'Seleciona a aba BC Semanal
Sheets("BC semanal").Select
'Escreve a data e horário da atualização
Range("A3").Value = "Última Atualização: " & Now
End Sub

Sub Imp_Exp()

Dim Setor_Externo As Workbook 'Define a planilha do setor externo
Dim Dado As Workbook 'Define a planilha com o dado atualizado
Dim FileURL As String 'Link para baixar o dado
Dim DestinationFile As String 'Caminho na nuvem que download será realizado

Dim linha_dado As Single  'ultima linha com dado da planilha com o dado atualizado
Dim Ano_Corrente As Single 'captura o ano corrente
Dim linha As Single 'captura a linha com o ultimo(dezembro) dado do ano anterior
Dim ultima_linha_ano_corrente As Single 'captura a linha com o primeiro(janeiro) dado do ano corrente
Dim ultima_linha_dados As Single ' captura a ultima linha com dados
Dim ultima_linha_dados2 As Single
Dim linha_Calculo As Single 'define a linha em que o calculo será escrito
Dim Acumulado_Exp As Single 'define o Valor acumulado das exportações até o mês corrente
Dim Acumulado_Imp As Single 'define o Valor acumulado das importações até o mês corrente
Dim Acumulado_Saldo As Single 'define o Valor acumulado do saldo até o mês corrente
Dim Var_Acumulado_Exp As Single 'define a variação do Valor acumulado das exportações até o mês corrente
Dim Var_Acumulado_Imp As Single 'define a variação do Valor acumulado das importações até o mês corrente
Dim Var_Acumulado_Saldo As Single 'define a variação do Valor acumulado do saldo até o mês corrente
Dim Var_Anual_Exp As Single 'Define a variação do valor acuulado na aba anual
Dim Var_Anual_Imp As Single 'Define a variação do valor acuulado na aba anual
Dim Var_Anual_Saldo As Single 'Define a variação do valor acuulado na aba anual
Dim mes_corrente As String 'Define o mes do dado mais recente
Dim nome As String 'Define o nome de usuário

'captura o nome do usuário que será usado para definir o caminho na nuvem para o download
'nome = VBA.Interaction.Environ$("UserName")
'Define o link de onde o arquivo será baixado
FileURL = "https://balanca.economia.gov.br/balanca/SH/TOTAL.xlsx"
'Define o caminho que o download será realizado
DestinationFile = "C:\Users\e-vinicius.geronimo\CNI - Confederação Nacional da Indústria\ECON - 2 Informe Conjuntural\Setor Externo\Base de Dados - Setor Externo\0 - Antigo\Comexstat\TOTAL" & Format(Date, "ddmmmyy") & ".xlsx"
'Baixa o dado no caminho definido
If URLDownloadToFile(0, FileURL, DestinationFile, 0, 0) = 0 Then
    Debug.Print "File download started"
Else
    Debug.Print "File download not started"
End If
'Torna a aba mensal visível
Sheets("Exp Imp Saldo - Mensal").Visible = True
'Chama a planilha de Setor Externo de Setor_Externo
Set Setor_Externo = ActiveWorkbook
'Abre a planilha com o dado selecionado
Workbooks.Open (DestinationFile)
'Chama a planilha com o dado de Dado
Set Dado = ActiveWorkbook
'Ativa a planilha dado
Dado.Activate
'Seleciona a aba com o dado
Sheets("DADOS_SH").Select
'Deleta colunas que não interessam
Sheets("DADOS_SH").Range("I:K").Delete
Sheets("DADOS_SH").Range("G:G").Delete
Sheets("DADOS_SH").Range("E:E").Delete
Sheets("DADOS_SH").Range("C:C").Delete

 'Captura o numero da ultima linha com os dados
linha_dado = Sheets("DADOS_SH").Range("A1").End(xlDown).Row
 'Copia
Range(Cells(1, 1), Cells(linha_dado, 5)).Copy
 'Seleciona a planilha setor externo
Setor_Externo.Activate
'Seleciona a aba
Sheets("Exp Imp Saldo - Mensal").Select
'Cola os dados copiados
Cells(9, 1).PasteSpecial xlPasteValues

'Ativa a planilha dado
Dado.Activate
'Desseleciona a parte copiada
Application.CutCopyMode = False
'Fecha a aba dados sem salvar as alterações
Dado.Close SaveChanges:=False
 'Seleciona a planilha setor externo
Setor_Externo.Activate
 'Seleciona a aba
Sheets("Exp Imp Saldo - Mensal").Select

'captura o ano atual
Ano_Corrente = Sheets("Exp Imp Saldo - Mensal").Range("A10").Value
'captura a linha com o ultimo(dezembro) dado do ano anterior
linha = 10 + Sheets("Exp Imp Saldo - Mensal").Range("B10").Value
'captura a linha com o primeiro(janeiro) dado do ano corrente
ultima_linha_ano_corrente = linha - 1
' captura a ultima linha com dados
ultima_linha_dados = Sheets("Exp Imp Saldo - Mensal").Range("A10").End(xlDown).Row
'define a linha em que o calculo será escrito
linha_Calculo = 10

'Loop que realiza o calculo do valor acumudado exportação do mes corrente até janeiro do ano corrente para todas as celulas
Do Until linha_Calculo = ultima_linha_dados
    Acumulado_Exp = Application.WorksheetFunction.Sum(Range(Cells(linha_Calculo, 3), Cells(ultima_linha_ano_corrente, 3)))
    Sheets("Exp Imp Saldo - Mensal").Cells(linha_Calculo, 7).Value = Acumulado_Exp
    linha_Calculo = linha_Calculo + 1
    ultima_linha_ano_corrente = ultima_linha_ano_corrente + 1
Loop

'Redefine o valor das variaveis usadas no calculo
ultima_linha_ano_corrente = linha - 1
linha_Calculo = 10

'Loop para o calculo do valor acumudado importação do mes corrente até janeiro do ano corrente para todas as celulas
Do Until linha_Calculo = ultima_linha_dados

    Acumulado_Imp = Application.WorksheetFunction.Sum(Range(Cells(linha_Calculo, 4), Cells(ultima_linha_ano_corrente, 4)))

    Sheets("Exp Imp Saldo - Mensal").Cells(linha_Calculo, 8).Value = Acumulado_Imp

    linha_Calculo = linha_Calculo + 1
    ultima_linha_ano_corrente = ultima_linha_ano_corrente + 1
    
Loop

ultima_linha_ano_corrente = linha - 1
linha_Calculo = 10

'Loop para o calculo do valor acumudado do saldo do mes corrente até janeiro do ano corrente para todas as celulas
Do Until linha_Calculo = ultima_linha_dados

    Acumulado_Saldo = Application.WorksheetFunction.Sum(Range(Cells(linha_Calculo, 5), Cells(ultima_linha_ano_corrente, 5)))

    Sheets("Exp Imp Saldo - Mensal").Cells(linha_Calculo, 9).Value = Acumulado_Saldo

    linha_Calculo = linha_Calculo + 1
    ultima_linha_ano_corrente = ultima_linha_ano_corrente + 1
Loop


linha_Calculo = 10

Do Until linha_Calculo = ultima_linha_dados - 1

    Var_Acumulado_Exp = ((Cells(linha_Calculo, 7) / Cells(linha_Calculo + 1, 7))) - 1
    
    Sheets("Exp Imp Saldo - Mensal").Cells(linha_Calculo, 11).Value = Var_Acumulado_Exp

    linha_Calculo = linha_Calculo + 1
    
Loop


linha_Calculo = 10

Do Until linha_Calculo = ultima_linha_dados - 1

    Var_Acumulado_Imp = ((Cells(linha_Calculo, 8) / Cells(linha_Calculo + 1, 8))) - 1
    
    Sheets("Exp Imp Saldo - Mensal").Cells(linha_Calculo, 12).Value = Var_Acumulado_Imp

    linha_Calculo = linha_Calculo + 1
    
Loop


linha_Calculo = 10

Do Until linha_Calculo = ultima_linha_dados - 1

    Var_Acumulado_Saldo = ((Cells(linha_Calculo, 9) / Cells(linha_Calculo + 1, 9))) - 1
    
    Sheets("Exp Imp Saldo - Mensal").Cells(linha_Calculo, 13).Value = Var_Acumulado_Saldo

    linha_Calculo = linha_Calculo + 1
    
Loop

'*********************************************************      Anual         *********************************************************************************************

Sheets("Exp Imp Saldo - Mensal").Select
Range("G9:I9,G10:I10,G22:I22,G34:I34,G46:I46,G58:I58,G70:I70,G82:I82,G94:I94,G106:I106,G118:I118,G130:I130,G142:I142,G154:I154,G166:I166,G178:I178,G190:I190,G202:I202,G214:I214,G226:I226,G238:I238,G250:I250,G262:I262,G274:I274,G286:I286").Select
Selection.Copy
Sheets("Exp Imp Saldo - Anual").Select
Range("B9").PasteSpecial (xlPasteValues)

Sheets("Exp Imp Saldo - Mensal").Select
Range("G298:I298,G310:I310,G322:I322,G334:I334,G346:I346,G358:I358,G370:I370,G382:I382,G394:I394,G406:I406,G418:I418,G430:I430,G442:I442,G454:I454,G466:I466,G478:I478,G490:I490,G502:I502").Select
Selection.Copy
Sheets("Exp Imp Saldo - Anual").Select
Range("B34").PasteSpecial (xlPasteValues)


Sheets("Exp Imp Saldo - Mensal").Select
Range("A9,A10,A22,A34").Copy
Sheets("Exp Imp Saldo - Anual").Select
Range("A9").PasteSpecial (xlPasteValues)

Range("A10:A12").Select
Selection.AutoFill Destination:=Range("A10:A48"), Type:=xlFillDefault

Range("A9").Value = "CO_ANO"
Range("B8:D8").Merge

linha_Calculo = 10
ultima_linha_dados2 = Sheets("Exp Imp Saldo - Anual").Range("B10").End(xlDown).Row

Do Until linha_Calculo = ultima_linha_dados2

    Var_Anual_Exp = ((Cells(linha_Calculo, 2) / Cells(linha_Calculo + 1, 2))) - 1
    
    Sheets("Exp Imp Saldo - Anual").Cells(linha_Calculo, 6).Value = Var_Anual_Exp

    linha_Calculo = linha_Calculo + 1
    
Loop

linha_Calculo = 10

Do Until linha_Calculo = ultima_linha_dados2

    Var_Anual_Imp = ((Cells(linha_Calculo, 3) / Cells(linha_Calculo + 1, 3))) - 1
    
    Sheets("Exp Imp Saldo - Anual").Cells(linha_Calculo, 7).Value = Var_Anual_Imp

    linha_Calculo = linha_Calculo + 1
    
Loop

linha_Calculo = 10

Do Until linha_Calculo = ultima_linha_dados2

    Var_Anual_Saldo = ((Cells(linha_Calculo, 4) / Cells(linha_Calculo + 1, 4))) - 1
    
    Sheets("Exp Imp Saldo - Anual").Cells(linha_Calculo, 8).Value = Var_Anual_Saldo

    linha_Calculo = linha_Calculo + 1
    
Loop

Sheets("Exp Imp Saldo - Mensal").Select
mes_corrente = Range("B10").Value


Sheets("Exp Imp Saldo - Anual").Range("B8").Value = Application.WorksheetFunction.Concat("Acumulado até o mês corrente - 01 até ", mes_corrente)
Sheets("Exp Imp Saldo - Anual").Range("F8").Value = Application.WorksheetFunction.Concat("Acumulado até o mês corrente - 01 até ", mes_corrente)


Sheets("Exp Imp Saldo - Mensal").Select
Range("A4").Value = "Última Atualização: " & Now

Sheets("Exp Imp Saldo - Anual").Select
Range("A4").Value = "Última Atualização: " & Now

Sheets("Exp Imp Saldo - Mensal").Visible = False

Range("A1").Select

End Sub

Sub X_M_media_diaria()

Dim Setor_Externo As Workbook
Dim Dado As Workbook
Dim FileURL As String
Dim DestinationFile As String
Dim nome As String
Dim ultima_linha As Integer


'nome = VBA.Interaction.Environ$("UserName")


FileURL = "https://balanca.economia.gov.br/balanca/SH/TOTAL.xlsx"

DestinationFile = "C:\Users\e-vinicius.geronimo\CNI - Confederação Nacional da Indústria\ECON - 2 Informe Conjuntural\Setor Externo\Base de Dados - Setor Externo\0 - Antigo\Comexstat\TOTAL" & Format(Date, "ddmmmyy") & ".xlsx"

If URLDownloadToFile(0, FileURL, DestinationFile, 0, 0) = 0 Then
    Debug.Print "File download started"
Else
    Debug.Print "File download not started"
End If


Set Setor_Externo = ActiveWorkbook

Sheets("Médias diárias X e M").Select
Rows("7:7").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

Workbooks.Open (DestinationFile)

Set Dado = ActiveWorkbook

Dado.Activate

Sheets("DADOS_SH").Select
ultima_linha = Range("A1").End(xlDown).Row
Range(Cells(2, 1), Cells(ultima_linha, 11)).Copy

Setor_Externo.Activate
Sheets("Médias diárias X e M").Select
Range("A7").PasteSpecial (xlPasteAll)

Dado.Activate
Application.CutCopyMode = False
Dado.Close

Setor_Externo.Activate
Sheets("Médias diárias X e M").Select
Range("M8:T8").Select
Selection.AutoFill Destination:=Range("M7:T8"), Type:=xlFillDefault
Range("A3").Value = "Última Atualização: " & Now
Rows("7:7").EntireRow.AutoFit



End Sub


Sub IP_IQ_Total()

Dim FileURL As String
Dim DestinationFile As String
Dim Setor_Externo As Workbook
Dim Dado As Workbook
Dim Intervalo_Dados As Double
Dim Nome_aba As String
Dim ultima_linha As Integer
Dim linha_tabela As Integer
Dim nome As String

'nome = VBA.Interaction.Environ$("UserName")
    
FileURL = "https://balanca.economia.gov.br/balanca/IPQ/arquivos/Dados_totais_mensal.csv"

DestinationFile = "C:\Users\e-vinicius.geronimo\CNI - Confederação Nacional da Indústria\ECON - 2 Informe Conjuntural\Setor Externo\Base de Dados - Setor Externo\0 - Antigo\IP e IQ\TOTAL\totais" & Format(Date, "ddmmmyy") & ".xls"

If URLDownloadToFile(0, FileURL, DestinationFile, 0, 0) = 0 Then
    Debug.Print "File download started"
Else
    Debug.Print "File download not started"
End If

Set Setor_Externo = ActiveWorkbook

Workbooks.Open (DestinationFile)

Set Dado = ActiveWorkbook

Dado.Activate

'Texto para colunas
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :=";", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1), Array(6, 1), Array(7, 1)), TrailingMinusNumbers:=True
        


'Cria a tabela ****************************************************************************************************************************

   
    ultima_linha = Range("B1").End(xlDown).Row
          
    Nome_aba = ActiveSheet.Name & "!"
               
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        Nome_aba & Range(Cells(1, 2), Cells(ultima_linha, 7)).Address(ReferenceStyle:=xlR1C1), Version:=7).CreatePivotTable _
        TableDestination:="Planilha1!R3C1", TableName:="PivotTable1", DefaultVersion _
        :=7
        
    Sheets("Planilha1").Select
    With ActiveSheet.PivotTables("PivotTable1")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
       
    With ActiveSheet.PivotTables("PivotTable1").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With

'Seleciona os campos ****************************************************************************************************

    ActiveSheet.PivotTables("PivotTable1").RepeatAllLabels xlRepeatLabels
    
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE")
        .Orientation = xlPageField
        .Position = 1
    End With
    
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO")
        .Orientation = xlPageField
        .Position = 2
    End With
    
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CO_ANO")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CO_MES")
        .Orientation = xlRowField
        .Position = 2
    End With
   
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("INDICE"), "Sum of INDICE", xlSum
        
        
'Formatando para forma tabular e repetir em cada linha******************************************************************************************************************************
          
          
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CO_ANO")
        .LayoutForm = xlTabular
        .RepeatLabels = True
    End With
   
   
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CO_MES")
        .LayoutForm = xlTabular
        .RepeatLabels = True
    End With
    
     
    ActiveSheet.PivotTables("PivotTable1").PivotFields("CO_MES").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
   
    ActiveSheet.PivotTables("PivotTable1").PivotFields("CO_ANO").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    

    ActiveSheet.PivotTables("PivotTable1").ColumnGrand = False
    
'Filtro ********************************************************************************************************************************

    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").CurrentPage = "EXP"
    
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").CurrentPage = "PRECO"
    
'Copia e cola ***************************************************************************************************************************

linha_tabela = Range("A3").End(xlDown).Row
   
Range(Cells(4, 1), Cells(linha_tabela, 3)).Copy

Setor_Externo.Activate
Sheets("IP IQ").Select
Range("A12").PasteSpecial xlPasteValues

'Filtro ********************************************************************************************************************************

Dado.Activate

    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").CurrentPage = "EXP"
    
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").CurrentPage = "QUANTUM"
     
'Copia e cola ***************************************************************************************************************************

Range(Cells(4, 1), Cells(linha_tabela, 3)).Copy

Setor_Externo.Activate
Sheets("IP IQ").Select
Range("D12").PasteSpecial xlPasteValues

'Filtro ********************************************************************************************************************************
Dado.Activate

    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").CurrentPage = "IMP"
    
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").CurrentPage = "PRECO"
    
    
'Copia e cola ***************************************************************************************************************************

Range(Cells(4, 1), Cells(linha_tabela, 3)).Copy

Setor_Externo.Activate
Sheets("IP IQ").Select
Range("G12").PasteSpecial xlPasteValues


'Filtro ********************************************************************************************************************************
Dado.Activate
  
ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").ClearAllFilters
ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").CurrentPage = "IMP"
    
ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").ClearAllFilters
ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").CurrentPage = "QUANTUM"
    
    
'Copia e cola ***************************************************************************************************************************

Range(Cells(4, 1), Cells(linha_tabela, 3)).Copy

Setor_Externo.Activate
Sheets("IP IQ").Select
Range("J12").PasteSpecial xlPasteValues


'Filtro ********************************************************************************************************************************

Dado.Activate

ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").ClearAllFilters
ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").CurrentPage = "EXP_DESSAZONALIZADA"
    
ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").ClearAllFilters
ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").CurrentPage = "QUANTUM"
    
'Copia e cola ***************************************************************************************************************************

Range(Cells(4, 1), Cells(linha_tabela, 3)).Copy

Setor_Externo.Activate
Sheets("IP IQ").Select
Range("M12").PasteSpecial xlPasteAll

'Filtro ********************************************************************************************************************************

Dado.Activate


ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").ClearAllFilters
ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").CurrentPage = "IMP_DESSAZONALIZADA"
    
ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").ClearAllFilters
ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").CurrentPage = "QUANTUM"


'Copia e cola ***************************************************************************************************************************

Range(Cells(4, 1), Cells(linha_tabela, 3)).Copy

Setor_Externo.Activate
Sheets("IP IQ").Select
Range("P12").PasteSpecial xlPasteAll


Dado.Activate
Application.CutCopyMode = False
Dado.Close

Setor_Externo.Activate
Sheets("IP IQ").Select
Range("A5").Value = "Última Atualização: " & Now

End Sub

Sub IP_IQ_GCE()

Dim FileURL As String
Dim DestinationFile As String
Dim Setor_Externo As Workbook
Dim Dado As Workbook
Dim Intervalo_Dados As Double
Dim Nome_aba As String
Dim ultima_linha As Integer
Dim linha_tabela As Integer
Dim nome As String

'nome = VBA.Interaction.Environ$("UserName")
 
 
FileURL = "https://balanca.economia.gov.br/balanca/IPQ/arquivos/Dados_cgce_mensal.csv"

DestinationFile = "C:\Users\e-vinicius.geronimo\CNI - Confederação Nacional da Indústria\ECON - 2 Informe Conjuntural\Setor Externo\Base de Dados - Setor Externo\0 - Antigo\IP e IQ\GCE\GCE" & Format(Date, "ddmmmyy") & ".xls"

If URLDownloadToFile(0, FileURL, DestinationFile, 0, 0) = 0 Then
    Debug.Print "File download started"
Else
    Debug.Print "File download not started"
End If

Set Setor_Externo = ActiveWorkbook

Workbooks.Open (DestinationFile)

Set Dado = ActiveWorkbook

Dado.Activate

'Texto para colunas
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :=";", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1), Array(6, 1), Array(7, 1)), TrailingMinusNumbers:=True
        
        
'Cria a tabela ****************************************************************************************************************************
   
    ultima_linha = Range("B1").End(xlDown).Row
          
    Nome_aba = ActiveSheet.Name & "!"
               
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        Nome_aba & Range(Cells(1, 2), Cells(ultima_linha, 10)).Address(ReferenceStyle:=xlR1C1), Version:=7).CreatePivotTable _
        TableDestination:="Planilha1!R3C1", TableName:="PivotTable1", DefaultVersion _
        :=7
        
    Sheets("Planilha1").Select
    With ActiveSheet.PivotTables("PivotTable1")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
       
    With ActiveSheet.PivotTables("PivotTable1").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
           
'Seleciona os campos ****************************************************************************************************
           
           
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CO_ANO")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CO_MES")
        .Orientation = xlRowField
        .Position = 2
    End With
    
    
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("NO_CGCE")
        .Orientation = xlColumnField
        .Position = 1
    End With
    
        
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE")
        .Orientation = xlPageField
        .Position = 1
    End With
    
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO")
        .Orientation = xlPageField
        .Position = 2
    End With
    
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("INDICE"), "Sum of INDICE", xlSum
    
'Formatando para forma tabular e repetir em cada linha e tirar os subtotais e grand totals******************************************************************************************************************************
    
    
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CO_MES")
        .LayoutForm = xlTabular
        .RepeatLabels = True
    End With
    
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CO_ANO")
        .LayoutForm = xlTabular
        .RepeatLabels = True
    End With
    
    ActiveSheet.PivotTables("PivotTable1").PivotFields("CO_ANO").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    
    ActiveSheet.PivotTables("PivotTable1").PivotFields("CO_MES").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    
    
    ActiveSheet.PivotTables("PivotTable1").RowGrand = False
    ActiveSheet.PivotTables("PivotTable1").ColumnGrand = False
    
 'Filtro ********************************************************************************************************************************
       
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").CurrentPage = "EXP"
    
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").CurrentPage = "PRECO"
    
'Copia e cola ***************************************************************************************************************************

linha_tabela = Range("A3").End(xlDown).Row

Range(Cells(5, 1), Cells(linha_tabela, 5)).Copy

Setor_Externo.Activate
Sheets("IP IQ - GCE").Select
Range("A12").PasteSpecial xlPasteValues

'Filtro ********************************************************************************************************************************
Dado.Activate
       
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").CurrentPage = "EXP"
    
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").CurrentPage = "QUANTUM"
    
'Copia e cola ***************************************************************************************************************************

Range(Cells(5, 1), Cells(linha_tabela, 5)).Copy

Setor_Externo.Activate
Sheets("IP IQ - GCE").Select
Range("F12").PasteSpecial xlPasteValues

'Filtro ********************************************************************************************************************************
Dado.Activate
       
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").CurrentPage = "IMP"
    
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").CurrentPage = "PRECO"
    
'Copia e cola ***************************************************************************************************************************

Range(Cells(5, 1), Cells(linha_tabela, 6)).Copy

Setor_Externo.Activate
Sheets("IP IQ - GCE").Select
Range("K12").PasteSpecial xlPasteValues

'Filtro ********************************************************************************************************************************
Dado.Activate
       
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").CurrentPage = "IMP"
    
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").CurrentPage = "QUANTUM"
    
'Copia e cola ***************************************************************************************************************************

Range(Cells(5, 1), Cells(linha_tabela, 6)).Copy

Setor_Externo.Activate
Sheets("IP IQ - GCE").Select
Range("Q12").PasteSpecial xlPasteValues


'Filtro ********************************************************************************************************************************
Dado.Activate
       
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").CurrentPage = "EXP_DESSAZONALIZADA"
    
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").CurrentPage = "QUANTUM"
    
'Copia e cola ***************************************************************************************************************************

Range(Cells(5, 1), Cells(linha_tabela, 6)).Copy

Setor_Externo.Activate
Sheets("IP IQ - GCE").Select
Range("W12").PasteSpecial xlPasteValues

'Filtro ********************************************************************************************************************************
Dado.Activate
       
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").CurrentPage = "IMP_DESSAZONALIZADA"
    
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").CurrentPage = "QUANTUM"
    
'Copia e cola ***************************************************************************************************************************

Range(Cells(5, 1), Cells(linha_tabela, 6)).Copy

Setor_Externo.Activate
Sheets("IP IQ - GCE").Select
Range("AB12").PasteSpecial xlPasteValues

  
Dado.Activate
Application.CutCopyMode = False
Dado.Close

Range("A5").Value = "Última Atualização: " & Now

End Sub


Sub IP_IQ_ISIC()

Dim FileURL As String
Dim DestinationFile As String
Dim Setor_Externo As Workbook
Dim Dado As Workbook
Dim Intervalo_Dados As Double
Dim Nome_aba As String
Dim ultima_linha As Double
Dim linha_tabela As Integer
Dim nome As String

'nome = VBA.Interaction.Environ$("UserName")
 
 
FileURL = "https://balanca.economia.gov.br/balanca/IPQ/arquivos/Dados_isic_mensal.csv"

DestinationFile = "C:\Users\e-vinicius.geronimo\CNI - Confederação Nacional da Indústria\ECON - 2 Informe Conjuntural\Setor Externo\Base de Dados - Setor Externo\0 - Antigo\IP e IQ\ISIC\GCE" & Format(Date, "ddmmmyy") & ".xls"

If URLDownloadToFile(0, FileURL, DestinationFile, 0, 0) = 0 Then
    Debug.Print "File download started"
Else
    Debug.Print "File download not started"
End If

Set Setor_Externo = ActiveWorkbook

Workbooks.Open (DestinationFile)

Set Dado = ActiveWorkbook

Dado.Activate

'Texto para colunas
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :=";", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1), Array(6, 1), Array(7, 1)), TrailingMinusNumbers:=True

'Cria a tabela ****************************************************************************************************************************


    ultima_linha = Range("A2").End(xlDown).Row
          
    Nome_aba = ActiveSheet.Name & "!"
    
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        Nome_aba & Range(Cells(1, 2), Cells(ultima_linha, 10)).Address(ReferenceStyle:=xlR1C1), Version:=7).CreatePivotTable _
        TableDestination:="Planilha1!R3C1", TableName:="PivotTable1", DefaultVersion _
        :=7
        
    With ActiveSheet.PivotTables("PivotTable1")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    
    With ActiveSheet.PivotTables("PivotTable1").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable1").RepeatAllLabels xlRepeatLabels
    
'Seleciona os campos ****************************************************************************************************
    
    
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO")
        .Orientation = xlPageField
        .Position = 1
    End With
    
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE")
        .Orientation = xlPageField
        .Position = 1
    End With
    
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CO_ANO")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CO_MES")
        .Orientation = xlRowField
        .Position = 2
    End With
    
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("INDICE"), "Sum of INDICE", xlSum
    
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("NO_ISIC")
        .Orientation = xlColumnField
        .Position = 1
    End With
    
'Formatando para forma tabular e repetir em cada linha e tirar os subtotais e grand totals******************************************************************************************************************************
    
    
    ActiveSheet.PivotTables("PivotTable1").PivotFields("CO_MES").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CO_MES")
        .LayoutForm = xlTabular
        .RepeatLabels = True
    End With
    
    ActiveSheet.PivotTables("PivotTable1").PivotFields("CO_ANO").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CO_ANO")
        .LayoutForm = xlTabular
        .RepeatLabels = True
    End With
        
'Filtro ********************************************************************************************************************************
    
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("NO_ISIC")
        .PivotItems("Agricultura e Pecuária").Visible = False
        .PivotItems("Confecção de Artigos do Vestuário e Acessórios").Visible = False
        .PivotItems("Extração de Minerais Não-Metálicos").Visible = False
        .PivotItems("Fabricação de Bebidas").Visible = False
        .PivotItems("Fabricação de Celulose, Papel e Produtos de Papel").Visible = False
        .PivotItems("Fabricação de Coque, de Produtos Derivados do Petróleo e de Biocombustíveis").Visible = False
        .PivotItems("Fabricação de Equipamentos de Informática, Produtos Eletrônicos e Ópticos").Visible = False
        .PivotItems("Fabricação de Máquinas e Equipamentos").Visible = False
        .PivotItems("Fabricação de Máquinas, Aparelhos e Materiais Elétricos").Visible = False
        .PivotItems("Fabricação de Móveis").Visible = False
        .PivotItems("Fabricação de Outros Equipamentos de Transporte, Exceto Véiculos Automotores").Visible = False
        .PivotItems("Fabricação de Produtos Alimentícios").Visible = False
        .PivotItems("Fabricação de Produtos de Borracha e Material Plásticos").Visible = False
        .PivotItems("Fabricação de Produtos de Madeira").Visible = False
        .PivotItems("Fabricação de Produtos de Metal, Exceto Máquinas e Equipamentos").Visible = False
        .PivotItems("Fabricação de Produtos Diversos").Visible = False
        .PivotItems("Fabricação de Produtos Farmoquímicos e Farmacêuticos").Visible = False
        .PivotItems("Fabricação de Produtos Minerais Não-Metálicos").Visible = False
        .PivotItems("Fabricação de Produtos Químicos").Visible = False
        .PivotItems("Fabricação de Produtos Têxteis").Visible = False
        .PivotItems("Fabricação de Veículos Automotores, Reboques e Carrocerias").Visible = False
        .PivotItems("Metalurgia").Visible = False
        .PivotItems("Preparação de Couros e Fabricação de Artefatos de Couro, Artigos para Viagem e Calçados").Visible = False
    
    End With
    ActiveSheet.PivotTables("PivotTable1").RowGrand = False
    ActiveSheet.PivotTables("PivotTable1").ColumnGrand = False
    
'Filtro ********************************************************************************************************************************
       
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").CurrentPage = "EXP"
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").CurrentPage = "PRECO"
    
'Copia e cola ***************************************************************************************************************************
    
linha_tabela = Range("A3").End(xlDown).Row

Range(Cells(5, 1), Cells(linha_tabela, 5)).Copy

Setor_Externo.Activate
Sheets("IP IQ - ISIC").Select
Range("A12").PasteSpecial xlPasteValues

'Filtro ********************************************************************************************************************************
 Dado.Activate
 
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").CurrentPage = "EXP"
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").CurrentPage = "QUANTUM"
    
'Copia e cola ***************************************************************************************************************************
    
Range(Cells(5, 1), Cells(linha_tabela, 5)).Copy

Setor_Externo.Activate
Sheets("IP IQ - ISIC").Select
Range("F12").PasteSpecial xlPasteValues


'Filtro ********************************************************************************************************************************
 Dado.Activate
 
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").CurrentPage = "IMP"
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").CurrentPage = "PRECO"
    
'Copia e cola ***************************************************************************************************************************
   
Range(Cells(5, 1), Cells(linha_tabela, 5)).Copy

Setor_Externo.Activate
Sheets("IP IQ - ISIC").Select
Range("K12").PasteSpecial xlPasteValues

'Filtro ********************************************************************************************************************************
 Dado.Activate
 
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").CurrentPage = "IMP"
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").CurrentPage = "QUANTUM"
    
'Copia e cola ***************************************************************************************************************************
   
Range(Cells(5, 1), Cells(linha_tabela, 5)).Copy

Setor_Externo.Activate
Sheets("IP IQ - ISIC").Select
Range("P12").PasteSpecial xlPasteValues
    
'Filtro ********************************************************************************************************************************
 Dado.Activate
 
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").CurrentPage = "EXP_DESSAZONALIZADA"
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").CurrentPage = "QUANTUM"
    
'Copia e cola ***************************************************************************************************************************
   
Range(Cells(5, 1), Cells(linha_tabela, 5)).Copy

Setor_Externo.Activate
Sheets("IP IQ - ISIC").Select
Range("U12").PasteSpecial xlPasteValues

'Filtro ********************************************************************************************************************************
 Dado.Activate
 
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").CurrentPage = "IMP_DESSAZONALIZADA"
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").CurrentPage = "QUANTUM"
    
'Copia e cola ***************************************************************************************************************************
   
Range(Cells(5, 1), Cells(linha_tabela, 5)).Copy

Setor_Externo.Activate
Sheets("IP IQ - ISIC").Select
Range("Z12").PasteSpecial xlPasteValues


Dado.Activate
Application.CutCopyMode = False
Dado.Close

Range("A5").Value = "Última Atualização: " & Now
        
End Sub

Sub IP_IQ_ISIC_DIV()
        
Dim FileURL As String
Dim DestinationFile As String
Dim Setor_Externo As Workbook
Dim Dado As Workbook
Dim Intervalo_Dados As Double
Dim Nome_aba As String
Dim ultima_linha As Double
Dim linha_tabela As Integer
Dim linha_cola As Integer
Dim nome As String
Dim linha_c As Double


'nome = VBA.Interaction.Environ$("UserName")
 
 
FileURL = "https://balanca.economia.gov.br/balanca/IPQ/arquivos/Dados_isic_mensal.csv"

DestinationFile = "C:\Users\e-vinicius.geronimo\CNI - Confederação Nacional da Indústria\ECON - 2 Informe Conjuntural\Setor Externo\Base de Dados - Setor Externo\0 - Antigo\IP e IQ\ISIC\GCEd" & Format(Date, "ddmmmyy") & ".xls"

If URLDownloadToFile(0, FileURL, DestinationFile, 0, 0) = 0 Then
    Debug.Print "File download started"
Else
    Debug.Print "File download not started"
End If

Set Setor_Externo = ActiveWorkbook
Sheets("IP IQ - ISIC div").Select
Range("A7:Z2000").Clear

Workbooks.Open (DestinationFile)

Set Dado = ActiveWorkbook

Dado.Activate

'Texto para colunas
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :=";", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1), Array(6, 1), Array(7, 1)), TrailingMinusNumbers:=True

'Cria a tabela ****************************************************************************************************************************


    ultima_linha = Range("A2").End(xlDown).Row
          
    Nome_aba = ActiveSheet.Name & "!"
    
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        Nome_aba & Range(Cells(1, 2), Cells(ultima_linha, 10)).Address(ReferenceStyle:=xlR1C1), Version:=7).CreatePivotTable _
        TableDestination:="Planilha1!R3C1", TableName:="PivotTable1", DefaultVersion _
        :=7
        
           
    With ActiveSheet.PivotTables("PivotTable1")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    
    With ActiveSheet.PivotTables("PivotTable1").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    
    ActiveSheet.PivotTables("PivotTable1").RepeatAllLabels xlRepeatLabels
    
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CO_ANO")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CO_MES")
        .Orientation = xlRowField
        .Position = 2
    End With
        
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("NO_ISIC")
        .Orientation = xlColumnField
        .Position = 1
    End With
    
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO")
        .Orientation = xlPageField
        .Position = 1
    End With
    
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE")
        .Orientation = xlPageField
        .Position = 1
    End With
    
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("INDICE"), "Sum of INDICE", xlSum
    
    
    ActiveSheet.PivotTables("PivotTable1").PivotFields("CO_ANO").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CO_ANO")
        .LayoutForm = xlTabular
        .RepeatLabels = True
    End With
    
    ActiveSheet.PivotTables("PivotTable1").RowGrand = False
    ActiveSheet.PivotTables("PivotTable1").ColumnGrand = False
    
    ActiveSheet.PivotTables("PivotTable1").PivotFields("CO_MES").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CO_MES")
        .LayoutForm = xlTabular
        .RepeatLabels = True
    End With
    
        With ActiveSheet.PivotTables("PivotTable1").PivotFields("NO_ISIC")
        .PivotItems("Agropecuária").Visible = False
        .PivotItems("Indústrias de Transformação").Visible = False
        .PivotItems("Indústrias Extrativas").Visible = False
    End With
    
'Filtro ********************************************************************************************************************************
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").CurrentPage = "EXP"
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").CurrentPage = "PRECO"

'Copia e cola ***************************************************************************************************************************
    
linha_tabela = Range("A4").End(xlDown).Row

Range(Cells(4, 1), Cells(linha_tabela, 25)).Copy

Setor_Externo.Activate
Sheets("IP IQ - ISIC div").Select

Cells(7, 1) = "PX"
Range("A8").PasteSpecial xlPasteValues

'Filtro ********************************************************************************************************************************
 Dado.Activate
 
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").CurrentPage = "EXP"
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").CurrentPage = "QUANTUM"
    
'Copia e cola ***************************************************************************************************************************
   
Setor_Externo.Activate
Sheets("IP IQ - ISIC div").Select
linha_cola = Range("A8").End(xlDown).Row

Cells(linha_cola + 2, 1) = "QX"

Dado.Activate
Range(Cells(4, 1), Cells(linha_tabela, 25)).Copy

Setor_Externo.Activate
linha_c = linha_cola + 3
Range(Cells(linha_c, 1), Cells(linha_c, 1)).PasteSpecial xlPasteValues

'Filtro ********************************************************************************************************************************
 Dado.Activate
 
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").CurrentPage = "IMP"
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").CurrentPage = "PRECO"
    
'Copia e cola ***************************************************************************************************************************
   
Range(Cells(4, 1), Cells(linha_tabela, 25)).Copy

Setor_Externo.Activate
Sheets("IP IQ - ISIC div").Select
linha_cola = Range("A8").End(xlDown).End(xlDown).End(xlDown).Row


Cells(linha_cola + 2, 1) = "PM"

Range(Cells(linha_cola + 3, 1), Cells(linha_cola + 3, 1)).PasteSpecial xlPasteValues

'Filtro ********************************************************************************************************************************
 Dado.Activate
 
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").CurrentPage = "IMP"
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").CurrentPage = "QUANTUM"
    
'Copia e cola ***************************************************************************************************************************
   
Range(Cells(4, 1), Cells(linha_tabela, 25)).Copy

Setor_Externo.Activate
Sheets("IP IQ - ISIC div").Select
linha_cola = Range("A8").End(xlDown).End(xlDown).End(xlDown).End(xlDown).End(xlDown).Row

Cells(linha_cola + 2, 1) = "QM"

Range(Cells(linha_cola + 3, 1), Cells(linha_cola + 3, 1)).PasteSpecial xlPasteValues

Range("A5").Value = "Última Atualização: " & Now
Range("A7").Select


'Filtro ********************************************************************************************************************************
 Dado.Activate
 
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").CurrentPage = "EXP_Dessazonalizada"
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").CurrentPage = "QUANTUM"
    
'Copia e cola ***************************************************************************************************************************
   
Range(Cells(4, 1), Cells(linha_tabela, 25)).Copy

Setor_Externo.Activate
Sheets("IQ Dessaz - ISIC div").Select
Range("A7:Z1000").Clear

Cells(7, 1) = "QX Dessazonalizado"

Dado.Activate
Range(Cells(4, 1), Cells(linha_tabela, 25)).Copy

Setor_Externo.Activate
Range("A8").PasteSpecial xlPasteValues

'Filtro ********************************************************************************************************************************
 Dado.Activate
 
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").CurrentPage = "IMP_Dessazonalizada"
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").CurrentPage = "QUANTUM"
    
'Copia e cola ***************************************************************************************************************************
   
Range(Cells(4, 1), Cells(linha_tabela, 25)).Copy

Setor_Externo.Activate
Sheets("IQ Dessaz - ISIC div").Select
linha_cola = Range("A8").End(xlDown).Row

Cells(linha_cola + 2, 1) = "QM Dessazonalizado"

Range(Cells(linha_cola + 3, 1), Cells(linha_cola + 3, 1)).PasteSpecial xlPasteValues


Dado.Activate
Application.CutCopyMode = False
Dado.Close

Range("A5").Value = "Última Atualização: " & Now
Range("A7").Select
       
End Sub

Sub Mapa_Calor()

Dim Ultima_Coluna As Integer

Sheets("Mapa de Calor").Select
Ultima_Coluna = Range("A8").End(xlDown).Column
Range(Cells(8, Ultima_Coluna), Cells(104, Ultima_Coluna)).Select
Selection.AutoFill Destination:=Range(Cells(8, Ultima_Coluna), Cells(104, Ultima_Coluna + 1)), Type:=xlFillDefault

End Sub

Sub Tabelas_X()

Dim ultima_linha As Integer
Dim ultimo_mes As Integer

Dim Total1 As Single
Dim Total2 As Single
Dim Total3 As Single

Dim Agropecuaria1 As Single
Dim Agropecuaria2 As Single
Dim Agropecuaria3 As Single

Dim Transformação1 As Single
Dim Transformação2 As Single
Dim Transformação3 As Single

Dim Extrativa1 As Single
Dim Extrativa2 As Single
Dim Extrativa3 As Single

Dim Mes1 As Integer
Dim Mes2 As Integer
Dim Mes3 As Integer

Dim Ano1 As Integer
Dim Ano2 As Integer
Dim Ano3 As Integer

Dim Data1 As String
Dim Data2 As String
Dim Data3 As String

Dim media1 As Single
Dim media2 As Single
Dim media3 As Single

Dim sRange1 As Range
Dim sRange2 As Range
Dim sRange3 As Range


ultima_linha = Sheets("IP IQ").Range("A12").End(xlDown).Row
ultimo_mes = Sheets("IP IQ").Range("B12").End(xlDown).Value - 1


Mes1 = Sheets("IP IQ").Cells(ultima_linha, 2).Value
Ano1 = Sheets("IP IQ").Cells(ultima_linha, 1).Value
Data1 = Mes1 & " " & Ano1
Sheets("Tabelas IP IQ - Exportação").Cells(13, 1) = Data1

Mes2 = Sheets("IP IQ").Cells(ultima_linha - 1, 2).Value
Ano2 = Sheets("IP IQ").Cells(ultima_linha - 1, 1).Value
Data2 = Mes2 & " " & Ano2
Sheets("Tabelas IP IQ - Exportação").Cells(12, 1) = Data2

Mes3 = Sheets("IP IQ").Cells(ultima_linha - 12, 2).Value
Ano3 = Sheets("IP IQ").Cells(ultima_linha - 12, 1).Value
Data3 = Mes3 & " " & Ano3
Sheets("Tabelas IP IQ - Exportação").Cells(11, 1) = Data3

Sheets("Tabelas IP IQ - Exportação").Cells(15, 1) = Data1 & " - " & Data3
Sheets("Tabelas IP IQ - Exportação").Cells(16, 1) = Data1 & " - " & Data2

Range("A11:A16").Copy Range("A21")
Range("A11:A16").Copy Range("A31")


'Total meses
Total1 = Sheets("IP IQ").Cells(ultima_linha, 3).Value
Total2 = Sheets("IP IQ").Cells(ultima_linha - 1, 3).Value
Total3 = Sheets("IP IQ").Cells(ultima_linha - 12, 3).Value
Sheets("Tabelas IP IQ - Exportação").Cells(11, 2).Value = Total3
Sheets("Tabelas IP IQ - Exportação").Cells(12, 2).Value = Total2
Sheets("Tabelas IP IQ - Exportação").Cells(13, 2).Value = Total1

Total1 = Sheets("IP IQ").Cells(ultima_linha, 6).Value
Total2 = Sheets("IP IQ").Cells(ultima_linha - 1, 6).Value
Total3 = Sheets("IP IQ").Cells(ultima_linha - 12, 6).Value
Sheets("Tabelas IP IQ - Exportação").Cells(21, 2).Value = Total3
Sheets("Tabelas IP IQ - Exportação").Cells(22, 2).Value = Total2
Sheets("Tabelas IP IQ - Exportação").Cells(23, 2).Value = Total1

Total1 = Sheets("IP IQ").Cells(ultima_linha, 15).Value
Total2 = Sheets("IP IQ").Cells(ultima_linha - 1, 15).Value
Total3 = Sheets("IP IQ").Cells(ultima_linha - 12, 15).Value
Sheets("Tabelas IP IQ - Exportação").Cells(31, 2).Value = Total3
Sheets("Tabelas IP IQ - Exportação").Cells(32, 2).Value = Total2
Sheets("Tabelas IP IQ - Exportação").Cells(33, 2).Value = Total1

'Agro meses
Agropecuaria1 = Sheets("IP IQ - ISIC").Cells(ultima_linha, 3).Value
Agropecuaria2 = Sheets("IP IQ - ISIC").Cells(ultima_linha - 1, 3).Value
Agropecuaria3 = Sheets("IP IQ - ISIC").Cells(ultima_linha - 12, 3).Value
Sheets("Tabelas IP IQ - Exportação").Cells(11, 3).Value = Agropecuaria3
Sheets("Tabelas IP IQ - Exportação").Cells(12, 3).Value = Agropecuaria2
Sheets("Tabelas IP IQ - Exportação").Cells(13, 3).Value = Agropecuaria1

Agropecuaria1 = Sheets("IP IQ - ISIC").Cells(ultima_linha, 8).Value
Agropecuaria2 = Sheets("IP IQ - ISIC").Cells(ultima_linha - 1, 8).Value
Agropecuaria3 = Sheets("IP IQ - ISIC").Cells(ultima_linha - 12, 8).Value
Sheets("Tabelas IP IQ - Exportação").Cells(21, 3).Value = Agropecuaria3
Sheets("Tabelas IP IQ - Exportação").Cells(22, 3).Value = Agropecuaria2
Sheets("Tabelas IP IQ - Exportação").Cells(23, 3).Value = Agropecuaria1

Agropecuaria1 = Sheets("IP IQ - ISIC").Cells(ultima_linha, 23).Value
Agropecuaria2 = Sheets("IP IQ - ISIC").Cells(ultima_linha - 1, 23).Value
Agropecuaria3 = Sheets("IP IQ - ISIC").Cells(ultima_linha - 12, 23).Value
Sheets("Tabelas IP IQ - Exportação").Cells(31, 3).Value = Agropecuaria3
Sheets("Tabelas IP IQ - Exportação").Cells(32, 3).Value = Agropecuaria2
Sheets("Tabelas IP IQ - Exportação").Cells(33, 3).Value = Agropecuaria1

'Transformação meses
Transformação1 = Sheets("IP IQ - ISIC").Cells(ultima_linha, 4).Value
Transformação2 = Sheets("IP IQ - ISIC").Cells(ultima_linha - 1, 4).Value
Transformação3 = Sheets("IP IQ - ISIC").Cells(ultima_linha - 12, 4).Value
Sheets("Tabelas IP IQ - Exportação").Cells(11, 4).Value = Transformação3
Sheets("Tabelas IP IQ - Exportação").Cells(12, 4).Value = Transformação2
Sheets("Tabelas IP IQ - Exportação").Cells(13, 4).Value = Transformação1

Transformação1 = Sheets("IP IQ - ISIC").Cells(ultima_linha, 9).Value
Transformação2 = Sheets("IP IQ - ISIC").Cells(ultima_linha - 1, 9).Value
Transformação3 = Sheets("IP IQ - ISIC").Cells(ultima_linha - 12, 9).Value
Sheets("Tabelas IP IQ - Exportação").Cells(21, 4).Value = Transformação3
Sheets("Tabelas IP IQ - Exportação").Cells(22, 4).Value = Transformação2
Sheets("Tabelas IP IQ - Exportação").Cells(23, 4).Value = Transformação1

Transformação1 = Sheets("IP IQ - ISIC").Cells(ultima_linha, 24).Value
Transformação2 = Sheets("IP IQ - ISIC").Cells(ultima_linha - 1, 24).Value
Transformação3 = Sheets("IP IQ - ISIC").Cells(ultima_linha - 12, 24).Value
Sheets("Tabelas IP IQ - Exportação").Cells(31, 4).Value = Transformação3
Sheets("Tabelas IP IQ - Exportação").Cells(32, 4).Value = Transformação2
Sheets("Tabelas IP IQ - Exportação").Cells(33, 4).Value = Transformação1


'Extrativa meses
Extrativa1 = Sheets("IP IQ - ISIC").Cells(ultima_linha, 5).Value
Extrativa2 = Sheets("IP IQ - ISIC").Cells(ultima_linha - 1, 5).Value
Extrativa3 = Sheets("IP IQ - ISIC").Cells(ultima_linha - 12, 5).Value
Sheets("Tabelas IP IQ - Exportação").Cells(11, 5).Value = Extrativa3
Sheets("Tabelas IP IQ - Exportação").Cells(12, 5).Value = Extrativa2
Sheets("Tabelas IP IQ - Exportação").Cells(13, 5).Value = Extrativa1

Extrativa1 = Sheets("IP IQ - ISIC").Cells(ultima_linha, 10).Value
Extrativa2 = Sheets("IP IQ - ISIC").Cells(ultima_linha - 1, 10).Value
Extrativa3 = Sheets("IP IQ - ISIC").Cells(ultima_linha - 12, 10).Value
Sheets("Tabelas IP IQ - Exportação").Cells(21, 5).Value = Extrativa3
Sheets("Tabelas IP IQ - Exportação").Cells(22, 5).Value = Extrativa2
Sheets("Tabelas IP IQ - Exportação").Cells(23, 5).Value = Extrativa1

Extrativa1 = Sheets("IP IQ - ISIC").Cells(ultima_linha, 25).Value
Extrativa2 = Sheets("IP IQ - ISIC").Cells(ultima_linha - 1, 25).Value
Extrativa3 = Sheets("IP IQ - ISIC").Cells(ultima_linha - 12, 25).Value
Sheets("Tabelas IP IQ - Exportação").Cells(31, 5).Value = Extrativa3
Sheets("Tabelas IP IQ - Exportação").Cells(32, 5).Value = Extrativa2
Sheets("Tabelas IP IQ - Exportação").Cells(33, 5).Value = Extrativa1


'Total Anos IPX
Sheets("IP IQ").Select
'Total2022
Set sRange1 = Range("C300:C" & ultima_linha)
media1 = Application.WorksheetFunction.Average(sRange1)
'Total2021
Set sRange2 = Range("C288:C" & 288 + ultimo_mes)
media2 = Application.WorksheetFunction.Average(sRange2)
'Total2020
Set sRange3 = Range("C276:C" & 276 + ultimo_mes)
media3 = Application.WorksheetFunction.Average(sRange3)

Sheets("Tabelas IP IQ - Exportação").Cells(11, 8).Value = media3
Sheets("Tabelas IP IQ - Exportação").Cells(12, 8).Value = media2
Sheets("Tabelas IP IQ - Exportação").Cells(13, 8).Value = media1

'Total Anos IQX
'Total2022
Set sRange1 = Range("F300:F" & ultima_linha)
media1 = Application.WorksheetFunction.Average(sRange1)
'Total2021
Set sRange2 = Range("F288:F" & 288 + ultimo_mes)
media2 = Application.WorksheetFunction.Average(sRange2)
'Total2020
Set sRange3 = Range("F276:F" & 276 + ultimo_mes)
media3 = Application.WorksheetFunction.Average(sRange3)

Sheets("Tabelas IP IQ - Exportação").Cells(21, 8).Value = media3
Sheets("Tabelas IP IQ - Exportação").Cells(22, 8).Value = media2
Sheets("Tabelas IP IQ - Exportação").Cells(23, 8).Value = media1



'AgropecuáriaAnos IPX
Sheets("IP IQ - ISIC").Select
'Agropecuária2022
Set sRange1 = Range("C300:C" & ultima_linha)
media1 = Application.WorksheetFunction.Average(sRange1)
'Agropecuária2021
Set sRange2 = Range("C288:C" & 288 + ultimo_mes)
media2 = Application.WorksheetFunction.Average(sRange2)
'Agropecuária2020
Set sRange3 = Range("C276:C" & 276 + ultimo_mes)
media3 = Application.WorksheetFunction.Average(sRange3)

Sheets("Tabelas IP IQ - Exportação").Cells(11, 9).Value = media3
Sheets("Tabelas IP IQ - Exportação").Cells(12, 9).Value = media2
Sheets("Tabelas IP IQ - Exportação").Cells(13, 9).Value = media1

'AgropecuáriaAnos IQX
'Agropecuária2022
Set sRange1 = Range("H300:H" & ultima_linha)
media1 = Application.WorksheetFunction.Average(sRange1)
'Agropecuária2021
Set sRange2 = Range("H288:H" & 288 + ultimo_mes)
media2 = Application.WorksheetFunction.Average(sRange2)
'Agropecuária2020
Set sRange3 = Range("H276:H" & 276 + ultimo_mes)
media3 = Application.WorksheetFunction.Average(sRange3)

Sheets("Tabelas IP IQ - Exportação").Cells(21, 9).Value = media3
Sheets("Tabelas IP IQ - Exportação").Cells(22, 9).Value = media2
Sheets("Tabelas IP IQ - Exportação").Cells(23, 9).Value = media1


'TransformaçãoAnos IPX
'Transformação2022
Set sRange1 = Range("D300:D" & ultima_linha)
media1 = Application.WorksheetFunction.Average(sRange1)
'Transformação2021
Set sRange2 = Range("D288:D" & 288 + ultimo_mes)
media2 = Application.WorksheetFunction.Average(sRange2)
'Transformação2020
Set sRange3 = Range("D276:D" & 276 + ultimo_mes)
media3 = Application.WorksheetFunction.Average(sRange3)

Sheets("Tabelas IP IQ - Exportação").Cells(11, 10).Value = media3
Sheets("Tabelas IP IQ - Exportação").Cells(12, 10).Value = media2
Sheets("Tabelas IP IQ - Exportação").Cells(13, 10).Value = media1

'TransformaçãoAnos IQX
'Transformação2022
Set sRange1 = Range("I300:I" & ultima_linha)
media1 = Application.WorksheetFunction.Average(sRange1)
'Transformação2021
Set sRange2 = Range("I288:I" & 288 + ultimo_mes)
media2 = Application.WorksheetFunction.Average(sRange2)
'Transformação2020
Set sRange3 = Range("I276:I" & 276 + ultimo_mes)
media3 = Application.WorksheetFunction.Average(sRange3)

Sheets("Tabelas IP IQ - Exportação").Cells(21, 10).Value = media3
Sheets("Tabelas IP IQ - Exportação").Cells(22, 10).Value = media2
Sheets("Tabelas IP IQ - Exportação").Cells(23, 10).Value = media1



'ExtrativaAnos IPX
'Extrativa2022
Set sRange1 = Range("E300:E" & ultima_linha)
media1 = Application.WorksheetFunction.Average(sRange1)
'Extrativa2021
Set sRange2 = Range("E288:E" & 288 + ultimo_mes)
media2 = Application.WorksheetFunction.Average(sRange2)
'Extrativa2020
Set sRange3 = Range("E276:E" & 276 + ultimo_mes)
media3 = Application.WorksheetFunction.Average(sRange3)

Sheets("Tabelas IP IQ - Exportação").Cells(11, 11).Value = media3
Sheets("Tabelas IP IQ - Exportação").Cells(12, 11).Value = media2
Sheets("Tabelas IP IQ - Exportação").Cells(13, 11).Value = media1

'ExtrativaAnos IQX
'Extrativa2022
Set sRange1 = Range("J300:J" & ultima_linha)
media1 = Application.WorksheetFunction.Average(sRange1)
'Extrativa2021
Set sRange2 = Range("J288:J" & 288 + ultimo_mes)
media2 = Application.WorksheetFunction.Average(sRange2)
'Extrativa2020
Set sRange3 = Range("J276:J" & 276 + ultimo_mes)
media3 = Application.WorksheetFunction.Average(sRange3)

Sheets("Tabelas IP IQ - Exportação").Cells(21, 11).Value = media3
Sheets("Tabelas IP IQ - Exportação").Cells(22, 11).Value = media2
Sheets("Tabelas IP IQ - Exportação").Cells(23, 11).Value = media1

Sheets("Tabelas IP IQ - Exportação").Range("A2").Value = "Última Atualização: " & Now
Sheets("Tabelas IP IQ - Exportação").Select

End Sub

Sub Tabelas_M()

Dim ultima_linha As Integer
Dim ultimo_mes As Integer

Dim Total1 As Single
Dim Total2 As Single
Dim Total3 As Single

Dim Agropecuaria1 As Single
Dim Agropecuaria2 As Single
Dim Agropecuaria3 As Single

Dim Transformação1 As Single
Dim Transformação2 As Single
Dim Transformação3 As Single

Dim Extrativa1 As Single
Dim Extrativa2 As Single
Dim Extrativa3 As Single

Dim Capital1 As Single
Dim Capital2 As Single
Dim Capital3 As Single

Dim Consumo1 As Single
Dim Consumo2 As Single
Dim Consumo3 As Single

Dim Intermediarios1 As Single
Dim Intermediarios2 As Single
Dim Intermediarios3 As Single

Dim Combustiveis1 As Single
Dim Combustiveis2 As Single
Dim Combustiveis3 As Single

Dim Mes1 As Integer
Dim Mes2 As Integer
Dim Mes3 As Integer

Dim Ano1 As Integer
Dim Ano2 As Integer
Dim Ano3 As Integer

Dim Data1 As String
Dim Data2 As String
Dim Data3 As String

Dim media1 As Single
Dim media2 As Single
Dim media3 As Single

Dim sRange1 As Range
Dim sRange2 As Range
Dim sRange3 As Range

ultima_linha = Sheets("IP IQ").Range("A11").End(xlDown).Row
ultimo_mes = Sheets("IP IQ").Range("B11").End(xlDown).Value - 1


Mes1 = Sheets("IP IQ").Cells(ultima_linha, 2).Value
Ano1 = Sheets("IP IQ").Cells(ultima_linha, 1).Value
Data1 = Mes1 & " " & Ano1
Sheets("Tabelas IP IQ - Importação").Cells(10, 1) = Data1

Mes2 = Sheets("IP IQ").Cells(ultima_linha - 1, 2).Value
Ano2 = Sheets("IP IQ").Cells(ultima_linha - 1, 1).Value
Data2 = Mes2 & " " & Ano2
Sheets("Tabelas IP IQ - Importação").Cells(9, 1) = Data2

Mes3 = Sheets("IP IQ").Cells(ultima_linha - 12, 2).Value
Ano3 = Sheets("IP IQ").Cells(ultima_linha - 12, 1).Value
Data3 = Mes3 & " " & Ano3
Sheets("Tabelas IP IQ - Importação").Cells(8, 1) = Data3

Sheets("Tabelas IP IQ - Importação").Cells(12, 1) = Data1 & " - " & Data3
Sheets("Tabelas IP IQ - Importação").Cells(13, 1) = Data1 & " - " & Data2

Range("A8:A13").Copy Range("A18")
Range("A8:A13").Copy Range("A28")

'Total meses
Total1 = Sheets("IP IQ").Cells(ultima_linha, 9).Value
Total2 = Sheets("IP IQ").Cells(ultima_linha - 1, 9).Value
Total3 = Sheets("IP IQ").Cells(ultima_linha - 12, 9).Value
Sheets("Tabelas IP IQ - Importação").Cells(8, 2).Value = Total3
Sheets("Tabelas IP IQ - Importação").Cells(9, 2).Value = Total2
Sheets("Tabelas IP IQ - Importação").Cells(10, 2).Value = Total1

Total1 = Sheets("IP IQ").Cells(ultima_linha, 12).Value
Total2 = Sheets("IP IQ").Cells(ultima_linha - 1, 12).Value
Total3 = Sheets("IP IQ").Cells(ultima_linha - 12, 12).Value
Sheets("Tabelas IP IQ - Importação").Cells(18, 2).Value = Total3
Sheets("Tabelas IP IQ - Importação").Cells(19, 2).Value = Total2
Sheets("Tabelas IP IQ - Importação").Cells(20, 2).Value = Total1

Total1 = Sheets("IP IQ").Cells(ultima_linha, 18).Value
Total2 = Sheets("IP IQ").Cells(ultima_linha - 1, 18).Value
Total3 = Sheets("IP IQ").Cells(ultima_linha - 12, 18).Value
Sheets("Tabelas IP IQ - Importação").Cells(28, 2).Value = Total3
Sheets("Tabelas IP IQ - Importação").Cells(29, 2).Value = Total2
Sheets("Tabelas IP IQ - Importação").Cells(30, 2).Value = Total1


'Agro meses
Agropecuaria1 = Sheets("IP IQ - ISIC").Cells(ultima_linha, 13).Value
Agropecuaria2 = Sheets("IP IQ - ISIC").Cells(ultima_linha - 1, 13).Value
Agropecuaria3 = Sheets("IP IQ - ISIC").Cells(ultima_linha - 12, 13).Value
Sheets("Tabelas IP IQ - Importação").Cells(8, 3).Value = Agropecuaria3
Sheets("Tabelas IP IQ - Importação").Cells(9, 3).Value = Agropecuaria2
Sheets("Tabelas IP IQ - Importação").Cells(10, 3).Value = Agropecuaria1

Agropecuaria1 = Sheets("IP IQ - ISIC").Cells(ultima_linha, 18).Value
Agropecuaria2 = Sheets("IP IQ - ISIC").Cells(ultima_linha - 1, 18).Value
Agropecuaria3 = Sheets("IP IQ - ISIC").Cells(ultima_linha - 12, 18).Value
Sheets("Tabelas IP IQ - Importação").Cells(18, 3).Value = Agropecuaria3
Sheets("Tabelas IP IQ - Importação").Cells(19, 3).Value = Agropecuaria2
Sheets("Tabelas IP IQ - Importação").Cells(20, 3).Value = Agropecuaria1

Agropecuaria1 = Sheets("IP IQ - ISIC").Cells(ultima_linha, 28).Value
Agropecuaria2 = Sheets("IP IQ - ISIC").Cells(ultima_linha - 1, 28).Value
Agropecuaria3 = Sheets("IP IQ - ISIC").Cells(ultima_linha - 12, 28).Value
Sheets("Tabelas IP IQ - Importação").Cells(28, 3).Value = Agropecuaria3
Sheets("Tabelas IP IQ - Importação").Cells(29, 3).Value = Agropecuaria2
Sheets("Tabelas IP IQ - Importação").Cells(30, 3).Value = Agropecuaria1


'Transformação meses
Transformação1 = Sheets("IP IQ - ISIC").Cells(ultima_linha, 14).Value
Transformação2 = Sheets("IP IQ - ISIC").Cells(ultima_linha - 1, 14).Value
Transformação3 = Sheets("IP IQ - ISIC").Cells(ultima_linha - 12, 14).Value
Sheets("Tabelas IP IQ - Importação").Cells(8, 4).Value = Transformação3
Sheets("Tabelas IP IQ - Importação").Cells(9, 4).Value = Transformação2
Sheets("Tabelas IP IQ - Importação").Cells(10, 4).Value = Transformação1

Transformação1 = Sheets("IP IQ - ISIC").Cells(ultima_linha, 19).Value
Transformação2 = Sheets("IP IQ - ISIC").Cells(ultima_linha - 1, 19).Value
Transformação3 = Sheets("IP IQ - ISIC").Cells(ultima_linha - 12, 19).Value
Sheets("Tabelas IP IQ - Importação").Cells(18, 4).Value = Transformação3
Sheets("Tabelas IP IQ - Importação").Cells(19, 4).Value = Transformação2
Sheets("Tabelas IP IQ - Importação").Cells(20, 4).Value = Transformação1

Transformação1 = Sheets("IP IQ - ISIC").Cells(ultima_linha, 29).Value
Transformação2 = Sheets("IP IQ - ISIC").Cells(ultima_linha - 1, 29).Value
Transformação3 = Sheets("IP IQ - ISIC").Cells(ultima_linha - 12, 29).Value
Sheets("Tabelas IP IQ - Importação").Cells(28, 4).Value = Transformação3
Sheets("Tabelas IP IQ - Importação").Cells(29, 4).Value = Transformação2
Sheets("Tabelas IP IQ - Importação").Cells(30, 4).Value = Transformação1


'Extrativa meses
Extrativa1 = Sheets("IP IQ - ISIC").Cells(ultima_linha, 15).Value
Extrativa2 = Sheets("IP IQ - ISIC").Cells(ultima_linha - 1, 15).Value
Extrativa3 = Sheets("IP IQ - ISIC").Cells(ultima_linha - 12, 15).Value
Sheets("Tabelas IP IQ - Importação").Cells(8, 5).Value = Extrativa3
Sheets("Tabelas IP IQ - Importação").Cells(9, 5).Value = Extrativa2
Sheets("Tabelas IP IQ - Importação").Cells(10, 5).Value = Extrativa1

Extrativa1 = Sheets("IP IQ - ISIC").Cells(ultima_linha, 20).Value
Extrativa2 = Sheets("IP IQ - ISIC").Cells(ultima_linha - 1, 20).Value
Extrativa3 = Sheets("IP IQ - ISIC").Cells(ultima_linha - 12, 20).Value
Sheets("Tabelas IP IQ - Importação").Cells(18, 5).Value = Extrativa3
Sheets("Tabelas IP IQ - Importação").Cells(19, 5).Value = Extrativa2
Sheets("Tabelas IP IQ - Importação").Cells(20, 5).Value = Extrativa1

Extrativa1 = Sheets("IP IQ - ISIC").Cells(ultima_linha, 30).Value
Extrativa2 = Sheets("IP IQ - ISIC").Cells(ultima_linha - 1, 30).Value
Extrativa3 = Sheets("IP IQ - ISIC").Cells(ultima_linha - 12, 30).Value
Sheets("Tabelas IP IQ - Importação").Cells(28, 5).Value = Extrativa3
Sheets("Tabelas IP IQ - Importação").Cells(29, 5).Value = Extrativa2
Sheets("Tabelas IP IQ - Importação").Cells(30, 5).Value = Extrativa1


'Capital meses
Capital1 = Sheets("IP IQ - GCE").Cells(ultima_linha, 13).Value
Capital2 = Sheets("IP IQ - GCE").Cells(ultima_linha - 1, 13).Value
Capital3 = Sheets("IP IQ - GCE").Cells(ultima_linha - 12, 13).Value
Sheets("Tabelas IP IQ - Importação").Cells(8, 6).Value = Capital3
Sheets("Tabelas IP IQ - Importação").Cells(9, 6).Value = Capital2
Sheets("Tabelas IP IQ - Importação").Cells(10, 6).Value = Capital1

Capital1 = Sheets("IP IQ - GCE").Cells(ultima_linha, 19).Value
Capital2 = Sheets("IP IQ - GCE").Cells(ultima_linha - 1, 19).Value
Capital3 = Sheets("IP IQ - GCE").Cells(ultima_linha - 12, 19).Value
Sheets("Tabelas IP IQ - Importação").Cells(18, 6).Value = Capital3
Sheets("Tabelas IP IQ - Importação").Cells(19, 6).Value = Capital2
Sheets("Tabelas IP IQ - Importação").Cells(20, 6).Value = Capital1

Capital1 = Sheets("IP IQ - GCE").Cells(ultima_linha, 30).Value
Capital2 = Sheets("IP IQ - GCE").Cells(ultima_linha - 1, 30).Value
Capital3 = Sheets("IP IQ - GCE").Cells(ultima_linha - 12, 30).Value
Sheets("Tabelas IP IQ - Importação").Cells(28, 6).Value = Capital3
Sheets("Tabelas IP IQ - Importação").Cells(29, 6).Value = Capital2
Sheets("Tabelas IP IQ - Importação").Cells(30, 6).Value = Capital1


'Consumo meses
Consumo1 = Sheets("IP IQ - GCE").Cells(ultima_linha, 14).Value
Consumo2 = Sheets("IP IQ - GCE").Cells(ultima_linha - 1, 14).Value
Consumo3 = Sheets("IP IQ - GCE").Cells(ultima_linha - 12, 14).Value
Sheets("Tabelas IP IQ - Importação").Cells(8, 7).Value = Consumo3
Sheets("Tabelas IP IQ - Importação").Cells(9, 7).Value = Consumo2
Sheets("Tabelas IP IQ - Importação").Cells(10, 7).Value = Consumo1

Consumo1 = Sheets("IP IQ - GCE").Cells(ultima_linha, 20).Value
Consumo2 = Sheets("IP IQ - GCE").Cells(ultima_linha - 1, 20).Value
Consumo3 = Sheets("IP IQ - GCE").Cells(ultima_linha - 12, 20).Value
Sheets("Tabelas IP IQ - Importação").Cells(18, 7).Value = Consumo3
Sheets("Tabelas IP IQ - Importação").Cells(19, 7).Value = Consumo2
Sheets("Tabelas IP IQ - Importação").Cells(20, 7).Value = Consumo1

Consumo1 = Sheets("IP IQ - GCE").Cells(ultima_linha, 31).Value
Consumo2 = Sheets("IP IQ - GCE").Cells(ultima_linha - 1, 31).Value
Consumo3 = Sheets("IP IQ - GCE").Cells(ultima_linha - 12, 31).Value
Sheets("Tabelas IP IQ - Importação").Cells(28, 7).Value = Consumo3
Sheets("Tabelas IP IQ - Importação").Cells(29, 7).Value = Consumo2
Sheets("Tabelas IP IQ - Importação").Cells(30, 7).Value = Consumo1


'Intermediários
Intermediarios1 = Sheets("IP IQ - GCE").Cells(ultima_linha, 15).Value
Intermediarios2 = Sheets("IP IQ - GCE").Cells(ultima_linha - 1, 15).Value
Intermediarios3 = Sheets("IP IQ - GCE").Cells(ultima_linha - 12, 15).Value
Sheets("Tabelas IP IQ - Importação").Cells(8, 8).Value = Intermediarios3
Sheets("Tabelas IP IQ - Importação").Cells(9, 8).Value = Intermediarios2
Sheets("Tabelas IP IQ - Importação").Cells(10, 8).Value = Intermediarios1

Intermediarios1 = Sheets("IP IQ - GCE").Cells(ultima_linha, 21).Value
Intermediarios2 = Sheets("IP IQ - GCE").Cells(ultima_linha - 1, 21).Value
Intermediarios3 = Sheets("IP IQ - GCE").Cells(ultima_linha - 12, 21).Value
Sheets("Tabelas IP IQ - Importação").Cells(18, 8).Value = Intermediarios3
Sheets("Tabelas IP IQ - Importação").Cells(19, 8).Value = Intermediarios2
Sheets("Tabelas IP IQ - Importação").Cells(20, 8).Value = Intermediarios1

Intermediarios1 = Sheets("IP IQ - GCE").Cells(ultima_linha, 32).Value
Intermediarios2 = Sheets("IP IQ - GCE").Cells(ultima_linha - 1, 32).Value
Intermediarios3 = Sheets("IP IQ - GCE").Cells(ultima_linha - 12, 32).Value
Sheets("Tabelas IP IQ - Importação").Cells(28, 8).Value = Intermediarios3
Sheets("Tabelas IP IQ - Importação").Cells(29, 8).Value = Intermediarios2
Sheets("Tabelas IP IQ - Importação").Cells(30, 8).Value = Intermediarios1

'Combustível
Combustiveis1 = Sheets("IP IQ - GCE").Cells(ultima_linha, 16).Value
Combustiveis2 = Sheets("IP IQ - GCE").Cells(ultima_linha - 1, 16).Value
Combustiveis3 = Sheets("IP IQ - GCE").Cells(ultima_linha - 12, 16).Value
Sheets("Tabelas IP IQ - Importação").Cells(8, 9).Value = Combustiveis3
Sheets("Tabelas IP IQ - Importação").Cells(9, 9).Value = Combustiveis2
Sheets("Tabelas IP IQ - Importação").Cells(10, 9).Value = Combustiveis1

Combustiveis1 = Sheets("IP IQ - GCE").Cells(ultima_linha, 22).Value
Combustiveis2 = Sheets("IP IQ - GCE").Cells(ultima_linha - 1, 22).Value
Combustiveis3 = Sheets("IP IQ - GCE").Cells(ultima_linha - 12, 22).Value
Sheets("Tabelas IP IQ - Importação").Cells(18, 9).Value = Combustiveis3
Sheets("Tabelas IP IQ - Importação").Cells(19, 9).Value = Combustiveis2
Sheets("Tabelas IP IQ - Importação").Cells(20, 9).Value = Combustiveis1

Combustiveis1 = Sheets("IP IQ - GCE").Cells(ultima_linha, 33).Value
Combustiveis2 = Sheets("IP IQ - GCE").Cells(ultima_linha - 1, 33).Value
Combustiveis3 = Sheets("IP IQ - GCE").Cells(ultima_linha - 12, 33).Value
Sheets("Tabelas IP IQ - Importação").Cells(28, 9).Value = Combustiveis3
Sheets("Tabelas IP IQ - Importação").Cells(29, 9).Value = Combustiveis2
Sheets("Tabelas IP IQ - Importação").Cells(30, 9).Value = Combustiveis1

'Total Anos IPM
Sheets("IP IQ").Select
'Total2022
Set sRange1 = Range("I300:I" & ultima_linha)
media1 = Application.WorksheetFunction.Average(sRange1)
'Total2021
Set sRange2 = Range("I288:I" & 288 + ultimo_mes)
media2 = Application.WorksheetFunction.Average(sRange2)
'Total2020
Set sRange3 = Range("I276:I" & 276 + ultimo_mes)
media3 = Application.WorksheetFunction.Average(sRange3)

Sheets("Tabelas IP IQ - Importação").Cells(8, 12).Value = media3
Sheets("Tabelas IP IQ - Importação").Cells(9, 12).Value = media2
Sheets("Tabelas IP IQ - Importação").Cells(10, 12).Value = media1

'Total Anos IQM
'Total2022
Set sRange1 = Range("L300:L" & ultima_linha)
media1 = Application.WorksheetFunction.Average(sRange1)
'Total2021
Set sRange2 = Range("L288:L" & 288 + ultimo_mes)
media2 = Application.WorksheetFunction.Average(sRange2)
'Total2020
Set sRange3 = Range("L276:L" & 276 + ultimo_mes)
media3 = Application.WorksheetFunction.Average(sRange3)

Sheets("Tabelas IP IQ - Importação").Cells(18, 12).Value = media3
Sheets("Tabelas IP IQ - Importação").Cells(19, 12).Value = media2
Sheets("Tabelas IP IQ - Importação").Cells(20, 12).Value = media1


'AgropecuáriaAnos IPM
Sheets("IP IQ - ISIC").Select
'Agropecuária2022
Set sRange1 = Range("M300:M" & ultima_linha)
media1 = Application.WorksheetFunction.Average(sRange1)
'Agropecuária2021
Set sRange2 = Range("M288:M" & 288 + ultimo_mes)
media2 = Application.WorksheetFunction.Average(sRange2)
'Agropecuária2020
Set sRange3 = Range("M276:M" & 276 + ultimo_mes)
media3 = Application.WorksheetFunction.Average(sRange3)

Sheets("Tabelas IP IQ - Importação").Cells(8, 13).Value = media3
Sheets("Tabelas IP IQ - Importação").Cells(9, 13).Value = media2
Sheets("Tabelas IP IQ - Importação").Cells(10, 13).Value = media1

'AgropecuáriaAnos IQM
'Agropecuária2022
Set sRange1 = Range("R300:R" & ultima_linha)
media1 = Application.WorksheetFunction.Average(sRange1)
'Agropecuária2021
Set sRange2 = Range("R288:R" & 288 + ultimo_mes)
media2 = Application.WorksheetFunction.Average(sRange2)
'Agropecuária2020
Set sRange3 = Range("R276:R" & 276 + ultimo_mes)
media3 = Application.WorksheetFunction.Average(sRange3)

Sheets("Tabelas IP IQ - Importação").Cells(18, 13).Value = media3
Sheets("Tabelas IP IQ - Importação").Cells(19, 13).Value = media2
Sheets("Tabelas IP IQ - Importação").Cells(20, 13).Value = media1


'TransformaçãoAnos IPM
'Transformação2022
Set sRange1 = Range("N300:N" & ultima_linha)
media1 = Application.WorksheetFunction.Average(sRange1)
'Transformação2021
Set sRange2 = Range("N288:N" & 288 + ultimo_mes)
media2 = Application.WorksheetFunction.Average(sRange2)
'Transformação2020
Set sRange3 = Range("N276:N" & 276 + ultimo_mes)
media3 = Application.WorksheetFunction.Average(sRange3)

Sheets("Tabelas IP IQ - Importação").Cells(8, 14).Value = media3
Sheets("Tabelas IP IQ - Importação").Cells(9, 14).Value = media2
Sheets("Tabelas IP IQ - Importação").Cells(10, 14).Value = media1


'TransformaçãoAnos IQM
'Transformação2022
Set sRange1 = Range("S300:S" & ultima_linha)
media1 = Application.WorksheetFunction.Average(sRange1)
'Transformação2021
Set sRange2 = Range("S288:S" & 288 + ultimo_mes)
media2 = Application.WorksheetFunction.Average(sRange2)
'Transformação2020
Set sRange3 = Range("S276:S" & 276 + ultimo_mes)
media3 = Application.WorksheetFunction.Average(sRange3)

Sheets("Tabelas IP IQ - Importação").Cells(18, 14).Value = media3
Sheets("Tabelas IP IQ - Importação").Cells(19, 14).Value = media2
Sheets("Tabelas IP IQ - Importação").Cells(20, 14).Value = media1



'ExtrativaAnos IPM
'Extrativa2022
Set sRange1 = Range("O300:O" & ultima_linha)
media1 = Application.WorksheetFunction.Average(sRange1)
'Extrativa2021
Set sRange2 = Range("O288:O" & 288 + ultimo_mes)
media2 = Application.WorksheetFunction.Average(sRange2)
'Extrativa2020
Set sRange3 = Range("O276:O" & 276 + ultimo_mes)
media3 = Application.WorksheetFunction.Average(sRange3)

Sheets("Tabelas IP IQ - Importação").Cells(8, 15).Value = media3
Sheets("Tabelas IP IQ - Importação").Cells(9, 15).Value = media2
Sheets("Tabelas IP IQ - Importação").Cells(10, 15).Value = media1

'ExtrativaAnos IQM
'Extrativa2022
Set sRange1 = Range("T300:T" & ultima_linha)
media1 = Application.WorksheetFunction.Average(sRange1)
'Extrativa2021
Set sRange2 = Range("T288:T" & 288 + ultimo_mes)
media2 = Application.WorksheetFunction.Average(sRange2)
'Extrativa2020
Set sRange3 = Range("T276:T" & 276 + ultimo_mes)
media3 = Application.WorksheetFunction.Average(sRange3)

Sheets("Tabelas IP IQ - Importação").Cells(18, 15).Value = media3
Sheets("Tabelas IP IQ - Importação").Cells(19, 15).Value = media2
Sheets("Tabelas IP IQ - Importação").Cells(20, 15).Value = media1


'CapitalAnos IPM
Sheets("IP IQ - GCE").Select
'Capital2022
Set sRange1 = Range("M300:M" & ultima_linha)
media1 = Application.WorksheetFunction.Average(sRange1)
'Capital2021
Set sRange2 = Range("M288:M" & 288 + ultimo_mes)
media2 = Application.WorksheetFunction.Average(sRange2)
'Capital2020
Set sRange3 = Range("M276:M" & 276 + ultimo_mes)
media3 = Application.WorksheetFunction.Average(sRange3)

Sheets("Tabelas IP IQ - Importação").Cells(8, 16).Value = media3
Sheets("Tabelas IP IQ - Importação").Cells(9, 16).Value = media2
Sheets("Tabelas IP IQ - Importação").Cells(10, 16).Value = media1

'CapitalAnos IQM
'Capital2022
Set sRange1 = Range("S300:S" & ultima_linha)
media1 = Application.WorksheetFunction.Average(sRange1)
'Capital2021
Set sRange2 = Range("S288:S" & 288 + ultimo_mes)
media2 = Application.WorksheetFunction.Average(sRange2)
'Capital2020
Set sRange3 = Range("S276:S" & 276 + ultimo_mes)
media3 = Application.WorksheetFunction.Average(sRange3)

Sheets("Tabelas IP IQ - Importação").Cells(18, 16).Value = media3
Sheets("Tabelas IP IQ - Importação").Cells(19, 16).Value = media2
Sheets("Tabelas IP IQ - Importação").Cells(20, 16).Value = media1


'ConsumoAnos IPM
'Consumo2022
Set sRange1 = Range("N300:N" & ultima_linha)
media1 = Application.WorksheetFunction.Average(sRange1)
'Consumo2021
Set sRange2 = Range("N288:N" & 288 + ultimo_mes)
media2 = Application.WorksheetFunction.Average(sRange2)
'Consumo2020
Set sRange3 = Range("N276:N" & 276 + ultimo_mes)
media3 = Application.WorksheetFunction.Average(sRange3)

Sheets("Tabelas IP IQ - Importação").Cells(8, 17).Value = media3
Sheets("Tabelas IP IQ - Importação").Cells(9, 17).Value = media2
Sheets("Tabelas IP IQ - Importação").Cells(10, 17).Value = media1

'ConsumoAnos IQM
'Consumo2022
Set sRange1 = Range("T300:T" & ultima_linha)
media1 = Application.WorksheetFunction.Average(sRange1)
'Consumo2021
Set sRange2 = Range("T288:T" & 288 + ultimo_mes)
media2 = Application.WorksheetFunction.Average(sRange2)
'Consumo2020
Set sRange3 = Range("T276:T" & 276 + ultimo_mes)
media3 = Application.WorksheetFunction.Average(sRange3)

Sheets("Tabelas IP IQ - Importação").Cells(18, 17).Value = media3
Sheets("Tabelas IP IQ - Importação").Cells(19, 17).Value = media2
Sheets("Tabelas IP IQ - Importação").Cells(20, 17).Value = media1



'IntermediáriosAnos IPM
'Intermediários2022
Set sRange1 = Range("O300:O" & ultima_linha)
media1 = Application.WorksheetFunction.Average(sRange1)
'Intermediários2021
Set sRange2 = Range("O288:O" & 288 + ultimo_mes)
media2 = Application.WorksheetFunction.Average(sRange2)
'Intermediários2020
Set sRange3 = Range("O276:O" & 276 + ultimo_mes)
media3 = Application.WorksheetFunction.Average(sRange3)

Sheets("Tabelas IP IQ - Importação").Cells(8, 18).Value = media3
Sheets("Tabelas IP IQ - Importação").Cells(9, 18).Value = media2
Sheets("Tabelas IP IQ - Importação").Cells(10, 18).Value = media1

'IntermediáriosAnos IQM
'Intermediários2022
Set sRange1 = Range("U300:U" & ultima_linha)
media1 = Application.WorksheetFunction.Average(sRange1)
'Intermediários2021
Set sRange2 = Range("U288:U" & 288 + ultimo_mes)
media2 = Application.WorksheetFunction.Average(sRange2)
'Intermediários2020
Set sRange3 = Range("U276:U" & 276 + ultimo_mes)
media3 = Application.WorksheetFunction.Average(sRange3)

Sheets("Tabelas IP IQ - Importação").Cells(18, 18).Value = media3
Sheets("Tabelas IP IQ - Importação").Cells(19, 18).Value = media2
Sheets("Tabelas IP IQ - Importação").Cells(20, 18).Value = media1


'CombustiveisAnos IPM
'Combustiveis2022
Set sRange1 = Range("P300:P" & ultima_linha)
media1 = Application.WorksheetFunction.Average(sRange1)
'Combustiveis2021
Set sRange2 = Range("P288:P" & 288 + ultimo_mes)
media2 = Application.WorksheetFunction.Average(sRange2)
'Combustiveis2020
Set sRange3 = Range("P276:P" & 276 + ultimo_mes)
media3 = Application.WorksheetFunction.Average(sRange3)

Sheets("Tabelas IP IQ - Importação").Cells(8, 19).Value = media3
Sheets("Tabelas IP IQ - Importação").Cells(9, 19).Value = media2
Sheets("Tabelas IP IQ - Importação").Cells(10, 19).Value = media1

'CombustiveisAnos IQM
'Combustiveis2022
Set sRange1 = Range("V300:V" & ultima_linha)
media1 = Application.WorksheetFunction.Average(sRange1)
'Combustiveis2021
Set sRange2 = Range("V288:V" & 288 + ultimo_mes)
media2 = Application.WorksheetFunction.Average(sRange2)
'Combustiveis2020
Set sRange3 = Range("V276:V" & 276 + ultimo_mes)
media3 = Application.WorksheetFunction.Average(sRange3)

Sheets("Tabelas IP IQ - Importação").Cells(18, 19).Value = media3
Sheets("Tabelas IP IQ - Importação").Cells(19, 19).Value = media2
Sheets("Tabelas IP IQ - Importação").Cells(20, 19).Value = media1

Sheets("Tabelas IP IQ - Importação").Range("A2").Value = "Última Atualização: " & Now
Sheets("Tabelas IP IQ - Importação").Select
End Sub
