Attribute VB_Name = "Automa��o_Constru��o"
Sub De_Play_Imediatamente()
Call Produ��o
Call Emprego
Call UCO
Call ICEI_Constru��o
Call Exp_atividade_empreendimentos
Call Exp_insumos_empregados
Call problemas_ponderado
Call condi��es_financeiras
Call Ivestimento_Constru��o
Call Tabelas

Sheets("indicadores").Visible = False
Sheets("problemas_ponderado").Visible = False

End Sub
Sub Produ��o()

Dim coluna As Integer
Dim Gr�fico As Object
Dim Gr�fico_edit As Object
Dim m�dia As Integer

'Seleciona a aba Indicadores
Sheets("Indicadores").Select
'Pega o n�mero da pultima coluna
coluna = Range("C10").End(xlToRight).Column
'Calcula a m�dia hist�rica
m�dia = Application.WorksheetFunction.Average(Range(Cells(11, 3), Cells(11, coluna)))
'Adiciona a aba Desempenho
Sheets.Add(Before:=Sheets("Indicadores")).Name = "Produ��o"
'Seleciona a aba desempenho
Sheets("Produ��o").Select
'Escreve os dois prmeiros meses da s�rie
Range("B1").Value = "12/01/09"
Range("C1").Value = "01/01/10"
'Completa a linha dos meses at� o �ltimo m�s
Range("B1:C1").Select
    Selection.NumberFormat = "mmm-yy"
    Selection.AutoFill Destination:=Range(Cells(1, 2), Cells(1, coluna - 1)), Type:=xlFillDefault
    
'Nomeia as s�ries
Range("A2").Value = "Produ��o"
Range("A3").Value = "M�dia hist�rica"
Range("A4").Value = "Linha Divis�ria"

'Copia e cola as s�ries
Sheets("Indicadores").Select
Range(Cells(11, 3), Cells(11, coluna)).Copy
Sheets("Produ��o").Select
Range("B2").PasteSpecial xlPasteValues
'Atribui o valor da m�dia
Range(Cells(3, 2), Cells(3, coluna - 1)).Value = m�dia
'Atribui o valor da m�dia
Range(Cells(4, 2), Cells(4, coluna - 1)).Value = "50"

'Cria o gr�fico
Set Gr�fico = Sheets("Produ��o").Shapes.AddChart2

Gr�fico.Select ' Seleciona o Gr�fico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gr�fico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gr�fico
    ActiveChart.Parent.Top = Parent.Range("B5").Top 'reposiciona o grafico em rela��o ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("B5").Left ' reposiciona o gr�fico em rela��o � borda esquerda da planilha
    ActiveChart.FullSeriesCollection(1).Name = "='Produ��o'!" & Cells(2, 1).Address 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(1).Values = "='Produ��o'!" & Range(Cells(2, coluna - 13), Cells(2, coluna - 1)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(1).XValues = "='Produ��o'!" & Range(Cells(1, coluna - 13), Cells(1, coluna - 1)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(2).Name = "='Produ��o'!$A$4" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(2).Values = "='Produ��o'!" & Range(Cells(4, coluna - 13), Cells(4, coluna - 1)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(2).XValues = "='Produ��o'!" & Range(Cells(1, coluna - 13), Cells(1, coluna - 1)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(3).Name = "='Produ��o'!$A$3" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(3).Values = "='Produ��o'!" & Range(Cells(3, coluna - 13), Cells(3, coluna - 1)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(3).XValues = "='Produ��o'!" & Range(Cells(1, coluna - 13), Cells(1, coluna - 1)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Sondagem Ind�stria da Constru��o\Automa��o\Templates\Produ��o_Constru��o") ' Aplica o template do gr�fico
    
Set Gr�fico_edit = Sheets("Produ��o").Shapes.AddChart2

Gr�fico_edit.Select ' Seleciona o Gr�fico
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Sondagem Ind�stria da Constru��o\Automa��o\Templates\Produ��o_Constru��o_edit") ' Aplica o template do gr�fico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gr�fico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gr�fico
    ActiveChart.Parent.Top = Parent.Range("L5").Top 'reposiciona o grafico em rela��o ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("L5").Left ' reposiciona o gr�fico em rela��o � borda esquerda da planilha
    ActiveChart.FullSeriesCollection(1).Name = "='Produ��o'!" & Cells(2, 1).Address 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(1).Values = "='Produ��o'!" & Range(Cells(2, coluna - 13), Cells(2, coluna - 1)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(1).XValues = "='Produ��o'!" & Range(Cells(1, coluna - 13), Cells(1, coluna - 1)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
End Sub
Sub Emprego()

Dim coluna As Integer
Dim Gr�fico As Object
Dim Gr�fico_edit As Object
Dim m�dia As Integer

'Seleciona a aba Indicadores
Sheets("Indicadores").Select
'Pega o n�mero da pultima coluna
coluna = Range("C39").End(xlToRight).Column
'Calcula a m�dia hist�rica
m�dia = Application.WorksheetFunction.Average(Range(Cells(39, 3), Cells(39, coluna)))
'Adiciona a aba Desempenho
Sheets.Add(Before:=Sheets("Indicadores")).Name = "Emprego"
'Seleciona a aba desempenho
Sheets("Emprego").Select
'Escreve os dois prmeiros meses da s�rie
Range("B1").Value = "01/01/11"
Range("C1").Value = "02/01/11"
'Completa a linha dos meses at� o �ltimo m�s
Range("B1:C1").Select
    Selection.NumberFormat = "mmm-yy"
    Selection.AutoFill Destination:=Range(Cells(1, 2), Cells(1, coluna - 1)), Type:=xlFillDefault
    
'Nomeia as s�ries
Range("A2").Value = "Emprego"
Range("A3").Value = "M�dia hist�rica"
Range("A4").Value = "Linha Divis�ria"

'Copia e cola as s�ries
Sheets("Indicadores").Select
Range(Cells(39, 3), Cells(39, coluna)).Copy
Sheets("Emprego").Select
Range("B2").PasteSpecial xlPasteValues
'Atribui o valor da m�dia
Range(Cells(3, 2), Cells(3, coluna - 1)).Value = m�dia
'Atribui o valor da m�dia
Range(Cells(4, 2), Cells(4, coluna - 1)).Value = "50"

'Cria o gr�fico
Set Gr�fico = Sheets("Emprego").Shapes.AddChart2

Gr�fico.Select ' Seleciona o Gr�fico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gr�fico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gr�fico
    ActiveChart.Parent.Top = Parent.Range("B5").Top 'reposiciona o grafico em rela��o ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("B5").Left ' reposiciona o gr�fico em rela��o � borda esquerda da planilha
    ActiveChart.FullSeriesCollection(1).Name = "='Emprego'!" & Cells(2, 1).Address 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(1).Values = "='Emprego'!" & Range(Cells(2, coluna - 13), Cells(2, coluna - 1)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(1).XValues = "='Emprego'!" & Range(Cells(1, coluna - 13), Cells(1, coluna - 1)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(2).Name = "='Emprego'!$A$4" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(2).Values = "='Emprego'!" & Range(Cells(4, coluna - 13), Cells(4, coluna - 1)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(2).XValues = "='Emprego'!" & Range(Cells(1, coluna - 13), Cells(1, coluna - 1)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(3).Name = "='Emprego'!$A$3" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(3).Values = "='Emprego'!" & Range(Cells(3, coluna - 13), Cells(3, coluna - 1)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(3).XValues = "='Emprego'!" & Range(Cells(1, coluna - 13), Cells(1, coluna - 1)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Sondagem Ind�stria da Constru��o\Automa��o\Templates\Emprego_Constru��o") ' Aplica o template do gr�fico
    
Set Gr�fico_edit = Sheets("Emprego").Shapes.AddChart2

Gr�fico_edit.Select ' Seleciona o Gr�fico
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Sondagem Ind�stria da Constru��o\Automa��o\Templates\Emprego_Constru��o_edit") ' Aplica o template do gr�fico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gr�fico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gr�fico
    ActiveChart.Parent.Top = Parent.Range("L5").Top 'reposiciona o grafico em rela��o ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("L5").Left ' reposiciona o gr�fico em rela��o � borda esquerda da planilha
    ActiveChart.FullSeriesCollection(1).Name = "='Emprego'!" & Cells(2, 1).Address 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(1).Values = "='Emprego'!" & Range(Cells(2, coluna - 13), Cells(2, coluna - 1)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(1).XValues = "='Emprego'!" & Range(Cells(1, coluna - 13), Cells(1, coluna - 1)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
End Sub

Sub UCO()

Dim C As Integer 'N�mero da �ltima Coluna
Dim Gr�fico As Object 'Gr�fico
Dim Gr�fico_edit As Object 'Gr�fico


Sheets("Indicadores").Select

C = Sheets("Indicadores").Range("C161").End(xlToRight).Column 'Define o n�mero da �ltima coluna

ActiveSheet.Range("C300").Value = "Jan"
ActiveSheet.Range("D300").Value = "Fev"
ActiveSheet.Range("E300").Value = "Mar"
ActiveSheet.Range("F300").Value = "Abr"
ActiveSheet.Range("G300").Value = "Mai"
ActiveSheet.Range("H300").Value = "Jun"
ActiveSheet.Range("I300").Value = "Jul"
ActiveSheet.Range("J300").Value = "Ago"
ActiveSheet.Range("K300").Value = "Set"
ActiveSheet.Range("L300").Value = "Out"
ActiveSheet.Range("M300").Value = "Nov"
ActiveSheet.Range("N300").Value = "Dez"

ActiveSheet.Range("B301").Value = "2012"
ActiveSheet.Range("B302").Value = "2013"
ActiveSheet.Range("B303").Value = "2014"
ActiveSheet.Range("B304").Value = "2015"
ActiveSheet.Range("B305").Value = "2016"
ActiveSheet.Range("B306").Value = "2017"
ActiveSheet.Range("B307").Value = "2018"
ActiveSheet.Range("B308").Value = "2019"
ActiveSheet.Range("B309").Value = "2020"
ActiveSheet.Range("B310").Value = "2021"
ActiveSheet.Range("B311").Value = "2022"
ActiveSheet.Range("B313").Value = "m�dia 2012-2014"
ActiveSheet.Range("B314").Value = "m�dia 2015-2019"

'2012
Sheets("Indicadores").Range("C161:N161").Copy (Sheets("Indicadores").Range("C301"))
'2013
Sheets("Indicadores").Range("O161:Z161").Copy (Sheets("Indicadores").Range("C302"))
'2014
Sheets("Indicadores").Range("AA161:AL161").Copy (Sheets("Indicadores").Range("C303"))
'2015
Sheets("Indicadores").Range("AM161:AX161").Copy (Sheets("Indicadores").Range("C304"))
'2016
Sheets("Indicadores").Range("AY161:BJ161").Copy (Sheets("Indicadores").Range("C305"))
'2017
Sheets("Indicadores").Range("BK161:BV161").Copy (Sheets("Indicadores").Range("C306"))
'2018
Sheets("Indicadores").Range("BW161:CH161").Copy (Sheets("Indicadores").Range("C307"))
'2019
Sheets("Indicadores").Range("CI161:CT161").Copy (Sheets("Indicadores").Range("C308"))
'2020
Sheets("Indicadores").Range("CU161:DF161").Copy (Sheets("Indicadores").Range("C309"))
'2021
Sheets("Indicadores").Range("DG161:DR161").Copy (Sheets("Indicadores").Range("C310"))
'2022
Sheets("Indicadores").Range(Cells(161, 123), Cells(161, C)).Copy (Sheets("Indicadores").Range("C311"))

ActiveSheet.Range("C313").Value = Application.Average(Range("C301:C303"))
ActiveSheet.Range("D313").Value = Application.Average(Range("D301:D303"))
ActiveSheet.Range("E313").Value = Application.Average(Range("E301:E303"))
ActiveSheet.Range("F313").Value = Application.Average(Range("F301:F303"))
ActiveSheet.Range("G313").Value = Application.Average(Range("G301:G303"))
ActiveSheet.Range("H313").Value = Application.Average(Range("H301:H303"))
ActiveSheet.Range("I313").Value = Application.Average(Range("I301:I303"))
ActiveSheet.Range("J313").Value = Application.Average(Range("J301:J303"))
ActiveSheet.Range("K313").Value = Application.Average(Range("K301:K303"))
ActiveSheet.Range("L313").Value = Application.Average(Range("L301:L303"))
ActiveSheet.Range("M313").Value = Application.Average(Range("M301:M303"))
ActiveSheet.Range("N313").Value = Application.Average(Range("N301:N303"))

ActiveSheet.Range("C314").Value = Application.Average(Range("C304:C308"))
ActiveSheet.Range("D314").Value = Application.Average(Range("D304:D308"))
ActiveSheet.Range("E314").Value = Application.Average(Range("E304:E308"))
ActiveSheet.Range("F314").Value = Application.Average(Range("F304:F308"))
ActiveSheet.Range("G314").Value = Application.Average(Range("G304:G308"))
ActiveSheet.Range("H314").Value = Application.Average(Range("H304:H308"))
ActiveSheet.Range("I314").Value = Application.Average(Range("I304:I308"))
ActiveSheet.Range("J314").Value = Application.Average(Range("J304:J308"))
ActiveSheet.Range("K314").Value = Application.Average(Range("K304:K308"))
ActiveSheet.Range("L314").Value = Application.Average(Range("L304:L308"))
ActiveSheet.Range("M314").Value = Application.Average(Range("M304:M308"))
ActiveSheet.Range("N314").Value = Application.Average(Range("N304:N308"))

'Adiciona a aba Desempenho
Sheets.Add(Before:=Sheets("Indicadores")).Name = "UCO (%)"
Sheets("Indicadores").Select
'Data
ActiveSheet.Range("C300:N300").Copy (Sheets("UCO (%)").Range("B3"))
'2020
ActiveSheet.Range("B309:N309").Copy (Sheets("UCO (%)").Range("A6"))
'2021
ActiveSheet.Range("B310:N310").Copy (Sheets("UCO (%)").Range("A7"))
'2022
ActiveSheet.Range("B311:N311").Copy (Sheets("UCO (%)").Range("A8"))
'M�dia1
ActiveSheet.Range("B313:N313").Copy (Sheets("UCO (%)").Range("A4"))
'm�dia2
ActiveSheet.Range("B314:N314").Copy (Sheets("UCO (%)").Range("A5"))


Sheets("UCO (%)").Select 'Seleciona a aba UCI (%)
Range("A1").Value = "Utiliza��o da capacidade de opera��o"

Set Gr�fico = Sheets("UCO (%)").Shapes.AddChart2 'Adiciona o gr�fico

Gr�fico.Select ' Seleciona o Gr�fico
ActiveChart.Parent.Height = 300 'ajusta a altura do gr�fico
ActiveChart.Parent.Width = 425 ' ajusta a largura do gr�fico
ActiveChart.Parent.Top = Parent.Range("A10").Top 'reposiciona o grafico em rela��o ao topo da planilha
ActiveChart.Parent.Left = Parent.Range("A10").Left ' reposiciona o gr�fico em rela��o � borda esquerda da planilha
ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
ActiveChart.FullSeriesCollection(1).Name = "='UCO (%)'!$A$4" 'Determina o nome da s�rie
ActiveChart.FullSeriesCollection(1).Values = "='UCO (%)'!$B$4:$M$4" 'determina os valores da s�rie
ActiveChart.FullSeriesCollection(1).XValues = "='UCO (%)'!$B$3:$M$3" 'determina os valores referentes ao eixo x da s�rie adicionada
ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
ActiveChart.FullSeriesCollection(2).Name = "='UCO (%)'!$A$5" 'Determina o nome da s�rie
ActiveChart.FullSeriesCollection(2).Values = "='UCO (%)'!$B$5:$M$5" 'determina os valores da s�rie
ActiveChart.FullSeriesCollection(2).XValues = "='UCO (%)'!$B$3:$M$3" 'determina os valores referentes ao eixo x da s�rie adicionada
ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
ActiveChart.FullSeriesCollection(3).Name = "='UCO (%)'!$A$6" 'Determina o nome da s�rie
ActiveChart.FullSeriesCollection(3).Values = "='UCO (%)'!$B$6:$M$6" 'determina os valores da s�rie
ActiveChart.FullSeriesCollection(3).XValues = "='UCO (%)'!$B$3:$M$3" 'determina os valores referentes ao eixo x da s�rie adicionada
ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
ActiveChart.FullSeriesCollection(4).Name = "='UCO (%)'!$A$7" 'Determina o nome da s�rie
ActiveChart.FullSeriesCollection(4).Values = "='UCO (%)'!$B$7:$M$7" 'determina os valores da s�rie
ActiveChart.FullSeriesCollection(4).XValues = "='UCO (%)'!$B$3:$M$3" 'determina os valores referentes ao eixo x da s�rie adicionada
ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
ActiveChart.FullSeriesCollection(5).Name = "='UCO (%)'!$A$8" 'Determina o nome da s�rie
ActiveChart.FullSeriesCollection(5).Values = "='UCO (%)'!$B$8:$M$8" 'determina os valores da s�rie
ActiveChart.FullSeriesCollection(5).XValues = "='UCO (%)'!$B$3:$M$3" 'determina os valores referentes ao eixo x da s�rie adicionada
ActiveChart.SetElement (msoElementLegendRight)
ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Sondagem Ind�stria da Constru��o\Automa��o\Templates\UCO_Constru��o") ' Aplica o template do gr�fico

Set Gr�fico_edit = Sheets("UCO (%)").Shapes.AddChart2

Gr�fico_edit.Select ' Seleciona o Gr�fico
ActiveChart.Parent.Height = 300 'ajusta a altura do gr�fico
ActiveChart.Parent.Width = 425 ' ajusta a largura do gr�fico
ActiveChart.Parent.Top = Parent.Range("K10").Top 'reposiciona o grafico em rela��o ao topo da planilha
ActiveChart.Parent.Left = Parent.Range("K10").Left ' reposiciona o gr�fico em rela��o � borda esquerda da planilha
ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
ActiveChart.FullSeriesCollection(1).Name = "='UCO (%)'!$A$4" 'Determina o nome da s�rie
ActiveChart.FullSeriesCollection(1).Values = "='UCO (%)'!$B$4:$M$4" 'determina os valores da s�rie
ActiveChart.FullSeriesCollection(1).XValues = "='UCO (%)'!$B$3:$M$3" 'determina os valores referentes ao eixo x da s�rie adicionada
ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
ActiveChart.FullSeriesCollection(2).Name = "='UCO (%)'!$A$5" 'Determina o nome da s�rie
ActiveChart.FullSeriesCollection(2).Values = "='UCO (%)'!$B$5:$M$5" 'determina os valores da s�rie
ActiveChart.FullSeriesCollection(2).XValues = "='UCO (%)'!$B$3:$M$3" 'determina os valores referentes ao eixo x da s�rie adicionada
ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
ActiveChart.FullSeriesCollection(3).Name = "='UCO (%)'!$A$6" 'Determina o nome da s�rie
ActiveChart.FullSeriesCollection(3).Values = "='UCO (%)'!$B$6:$M$6" 'determina os valores da s�rie
ActiveChart.FullSeriesCollection(3).XValues = "='UCO (%)'!$B$3:$M$3" 'determina os valores referentes ao eixo x da s�rie adicionada
ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
ActiveChart.FullSeriesCollection(4).Name = "='UCO (%)'!$A$7" 'Determina o nome da s�rie
ActiveChart.FullSeriesCollection(4).Values = "='UCO (%)'!$B$7:$M$7" 'determina os valores da s�rie
ActiveChart.FullSeriesCollection(4).XValues = "='UCO (%)'!$B$3:$M$3" 'determina os valores referentes ao eixo x da s�rie adicionada
ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
ActiveChart.FullSeriesCollection(5).Name = "='UCO (%)'!$A$8" 'Determina o nome da s�rie
ActiveChart.FullSeriesCollection(5).Values = "='UCO (%)'!$B$8:$M$8" 'determina os valores da s�rie
ActiveChart.FullSeriesCollection(5).XValues = "='UCO (%)'!$B$3:$M$3" 'determina os valores referentes ao eixo x da s�rie adicionada
ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Sondagem Ind�stria da Constru��o\Automa��o\Templates\UCO_Constru��o_Edit") ' Aplica o template do gr�fico
End Sub
Sub ICEI_Constru��o()

Dim coluna As Integer
Dim Gr�fico As Object
Dim Gr�fico_edit As Object
Dim m�dia As Integer

'Seleciona a aba Indicadores
Sheets("Indicadores").Select
'Pega o n�mero da pultima coluna
coluna = Range("C203").End(xlToRight).Column
'Calcula a m�dia hist�rica
m�dia = Application.WorksheetFunction.Average(Range(Cells(203, 3), Cells(203, coluna)))
'Adiciona a aba Desempenho
Sheets.Add(Before:=Sheets("Indicadores")).Name = "ICEI_Constru��o"
'Seleciona a aba desempenho
Sheets("ICEI_Constru��o").Select
'Escreve os dois prmeiros meses da s�rie
Range("B1").Value = "01/01/10"
Range("C1").Value = "02/01/10"
'Completa a linha dos meses at� o �ltimo m�s
Range("B1:C1").Select
    Selection.NumberFormat = "mmm-yy"
    Selection.AutoFill Destination:=Range(Cells(1, 2), Cells(1, coluna - 1)), Type:=xlFillDefault
    
'Nomeia as s�ries
Range("A2").Value = "ICEI"
Range("A3").Value = "Linha Divis�ria"
Range("A4").Value = "M�dia hist�rica"
Range("A5").Value = "M�s destaque"

'Copia e cola as s�ries
Sheets("Indicadores").Select
Range(Cells(203, 3), Cells(203, coluna)).Copy
Sheets("ICEI_Constru��o").Select
Range("B2").PasteSpecial xlPasteValues
'Atribui o valor da m�dia
Range(Cells(3, 2), Cells(3, coluna - 1)).Value = "50"
'Atribui o valor da m�dia
Range(Cells(4, 2), Cells(4, coluna - 1)).Value = m�dia

coluna_ICEI = Range("B2").End(xlToRight).Column
Do Until coluna_ICEI <= 2
Sheets("ICEI_Constru��o").Cells(5, coluna_ICEI) = Sheets("ICEI_Constru��o").Cells(2, coluna_ICEI)
coluna_ICEI = coluna_ICEI - 12
Loop

'Cria o gr�fico
Set Gr�fico = Sheets("ICEI_Constru��o").Shapes.AddChart2

Gr�fico.Select ' Seleciona o Gr�fico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gr�fico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gr�fico
    ActiveChart.Parent.Top = Parent.Range("A7").Top 'reposiciona o grafico em rela��o ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("A7").Left ' reposiciona o gr�fico em rela��o � borda esquerda da planilha
    ActiveChart.FullSeriesCollection(1).Name = "='ICEI_Constru��o'!" & Cells(2, 1).Address 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(1).Values = "='ICEI_Constru��o'!" & Range(Cells(2, coluna - 133), Cells(2, coluna - 1)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(1).XValues = "='ICEI_Constru��o'!" & Range(Cells(1, coluna - 133), Cells(1, coluna - 1)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(2).Name = "='ICEI_Constru��o'!$A$3" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(2).Values = "='ICEI_Constru��o'!" & Range(Cells(3, coluna - 133), Cells(3, coluna - 1)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(2).XValues = "='ICEI_Constru��o'!" & Range(Cells(1, coluna - 133), Cells(1, coluna - 1)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(3).Name = "='ICEI_Constru��o'!$A$4" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(3).Values = "='ICEI_Constru��o'!" & Range(Cells(4, coluna - 133), Cells(4, coluna - 1)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(3).XValues = "='ICEI_Constru��o'!" & Range(Cells(1, coluna - 133), Cells(1, coluna - 1)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(4).Name = "='ICEI_Constru��o'!$A$5" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(4).Values = "='ICEI_Constru��o'!" & Range(Cells(5, coluna - 133), Cells(5, coluna - 1)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(4).XValues = "='ICEI_Constru��o'!" & Range(Cells(1, coluna - 133), Cells(1, coluna - 1)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Sondagem Ind�stria da Constru��o\Automa��o\Templates\ICEI_Constru��o") ' Aplica o template do gr�fico
    ActiveChart.FullSeriesCollection(4).DataLabels.Select
    Selection.Position = xlLabelPositionAbove
    
Set Gr�fico_edit = Sheets("ICEI_Constru��o").Shapes.AddChart2

Gr�fico_edit.Select ' Seleciona o Gr�fico
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Sondagem Ind�stria da Constru��o\Automa��o\Templates\ICEI_Constru��o_edit") ' Aplica o template do gr�fico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gr�fico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gr�fico
    ActiveChart.Parent.Top = Parent.Range("K7").Top 'reposiciona o grafico em rela��o ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("K7").Left ' reposiciona o gr�fico em rela��o � borda esquerda da planilha
    ActiveChart.FullSeriesCollection(1).Name = "='ICEI_Constru��o'!" & Cells(2, 1).Address 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(1).Values = "='ICEI_Constru��o'!" & Range(Cells(2, coluna - 133), Cells(2, coluna - 1)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(1).XValues = "='ICEI_Constru��o'!" & Range(Cells(1, coluna - 133), Cells(1, coluna - 1)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(2).Name = "='ICEI_Constru��o'!$A$4" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(2).Values = "='ICEI_Constru��o'!" & Range(Cells(4, coluna - 133), Cells(4, coluna - 1)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(2).XValues = "='ICEI_Constru��o'!" & Range(Cells(1, coluna - 133), Cells(1, coluna - 1)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(3).Name = "='ICEI_Constru��o'!$A$5" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(3).Values = "='ICEI_Constru��o'!" & Range(Cells(5, coluna - 133), Cells(5, coluna - 1)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(3).XValues = "='ICEI_Constru��o'!" & Range(Cells(1, coluna - 133), Cells(1, coluna - 1)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.FullSeriesCollection(2).Select
    
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = -0.5
        .Transparency = 0
        .Visible = msoTrue
        .DashStyle = msoLineSysDot
    End With
    ActiveChart.FullSeriesCollection(3).Select
    With Selection
        .MarkerStyle = 1
        .MarkerSize = 5
    End With
    Selection.MarkerStyle = 8
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(3, 103, 173)
        .Transparency = 0
        .Solid
    End With
    ActiveChart.SetElement (msoElementDataLabelTop)
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Format.Line.Visible = msoFalse
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    With Selection.Format.TextFrame2.TextRange.Font.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(3, 103, 173)
        .Transparency = 0
        .Solid
    End With
    ActiveChart.FullSeriesCollection(3).Select
    Selection.Format.Line.Visible = msoFalse
    Range("A1").Select
End Sub

Sub Exp_atividade_empreendimentos()

Dim coluna As Integer
Dim Gr�fico As Object
Dim Gr�fico_edit As Object

'Seleciona a aba Indicadores
Sheets("Indicadores").Select
'Pega o n�mero da pultima coluna
coluna = Range("C105").End(xlToRight).Column
'Adiciona a aba Desempenho
Sheets.Add(Before:=Sheets("Indicadores")).Name = "Exp_atividade e empreendimentos"
'Seleciona a aba desempenho
Sheets("Exp_atividade e empreendimentos").Select
'Escreve os dois prmeiros meses da s�rie
Range("B1").Value = "01/01/10"
Range("C1").Value = "02/01/10"
'Completa a linha dos meses at� o �ltimo m�s
Range("B1:C1").Select
    Selection.NumberFormat = "mmm-yy"
    Selection.AutoFill Destination:=Range(Cells(1, 2), Cells(1, coluna - 1)), Type:=xlFillDefault
    
'Nomeia as s�ries
Range("A2").Value = "Expectativa do n�vel de atividade"
Range("A3").Value = "Expectativa de novos empreendimentos e servi�os"
Range("A4").Value = "Linha Divis�ria"

'Copia e cola as s�ries
Sheets("Indicadores").Select
Range(Cells(105, 3), Cells(105, coluna)).Copy
Sheets("Exp_atividade e empreendimentos").Select
Range("B2").PasteSpecial xlPasteValues
'Copia e cola as s�ries
Sheets("Indicadores").Select
Range(Cells(119, 3), Cells(119, coluna)).Copy
Sheets("Exp_atividade e empreendimentos").Select
Range("B3").PasteSpecial xlPasteValues
'Atribui o valor da m�dia
Range(Cells(4, 2), Cells(4, coluna - 1)).Value = "50"

'Cria o gr�fico
Set Gr�fico = Sheets("Exp_atividade e empreendimentos").Shapes.AddChart2

Gr�fico.Select ' Seleciona o Gr�fico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gr�fico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gr�fico
    ActiveChart.Parent.Top = Parent.Range("A6").Top 'reposiciona o grafico em rela��o ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("A65").Left ' reposiciona o gr�fico em rela��o � borda esquerda da planilha
    ActiveChart.FullSeriesCollection(1).Name = "='Exp_atividade e empreendimentos'!" & Cells(2, 1).Address 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(1).Values = "='Exp_atividade e empreendimentos'!" & Range(Cells(2, coluna - 121), Cells(2, coluna - 1)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(1).XValues = "='Exp_atividade e empreendimentos'!" & Range(Cells(1, coluna - 121), Cells(1, coluna - 1)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(2).Name = "='Exp_atividade e empreendimentos'!" & Cells(3, 1).Address 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(2).Values = "='Exp_atividade e empreendimentos'!" & Range(Cells(3, coluna - 121), Cells(3, coluna - 1)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(2).XValues = "='Exp_atividade e empreendimentos'!" & Range(Cells(1, coluna - 121), Cells(1, coluna - 1)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(3).Name = "='Exp_atividade e empreendimentos'!" & Cells(4, 1).Address 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(3).Values = "='Exp_atividade e empreendimentos'!" & Range(Cells(4, coluna - 121), Cells(4, coluna - 1)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(3).XValues = "='Exp_atividade e empreendimentos'!" & Range(Cells(1, coluna - 121), Cells(1, coluna - 1)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Sondagem Ind�stria da Constru��o\Automa��o\Templates\Expectativa - Atividade Empreendimentos - Constru��o") ' Aplica o template do gr�fico
    
Set Gr�fico_edit = Sheets("Exp_atividade e empreendimentos").Shapes.AddChart2

Gr�fico_edit.Select ' Seleciona o Gr�fico
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Sondagem Ind�stria da Constru��o\Automa��o\Templates\Expectativa - Atividade Empreendimentos - Constru��o_Edit") ' Aplica o template do gr�fico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gr�fico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gr�fico
    ActiveChart.Parent.Top = Parent.Range("K6").Top 'reposiciona o grafico em rela��o ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("K6").Left ' reposiciona o gr�fico em rela��o � borda esquerda da planilha
    ActiveChart.FullSeriesCollection(1).Name = "='Exp_atividade e empreendimentos'!" & Cells(2, 1).Address 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(1).Values = "='Exp_atividade e empreendimentos'!" & Range(Cells(2, coluna - 121), Cells(2, coluna - 1)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(1).XValues = "='Exp_atividade e empreendimentos'!" & Range(Cells(1, coluna - 121), Cells(1, coluna - 1)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(2).Name = "='Exp_atividade e empreendimentos'!" & Cells(3, 1).Address 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(2).Values = "='Exp_atividade e empreendimentos'!" & Range(Cells(3, coluna - 121), Cells(3, coluna - 1)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(2).XValues = "='Exp_atividade e empreendimentos'!" & Range(Cells(1, coluna - 121), Cells(1, coluna - 1)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
End Sub

Sub Exp_insumos_empregados()

Dim coluna As Integer
Dim Gr�fico As Object
Dim Gr�fico_edit As Object

'Seleciona a aba Indicadores
Sheets("Indicadores").Select
'Pega o n�mero da pultima coluna
coluna = Range("C133").End(xlToRight).Column
'Adiciona a aba Desempenho
Sheets.Add(Before:=Sheets("Indicadores")).Name = "Exp_insumos e empregados"
'Seleciona a aba desempenho
Sheets("Exp_insumos e empregados").Select
'Escreve os dois prmeiros meses da s�rie
Range("B1").Value = "01/01/10"
Range("C1").Value = "02/01/10"
'Completa a linha dos meses at� o �ltimo m�s
Range("B1:C1").Select
    Selection.NumberFormat = "mmm-yy"
    Selection.AutoFill Destination:=Range(Cells(1, 2), Cells(1, coluna - 1)), Type:=xlFillDefault
    
'Nomeia as s�ries
Range("A2").Value = "Expectativa de compras de insumos e mat�rias-primas"
Range("A3").Value = "Expectativa do n�mero de empregados"
Range("A4").Value = "Linha Divis�ria"

'Copia e cola as s�ries
Sheets("Indicadores").Select
Range(Cells(133, 3), Cells(133, coluna)).Copy
Sheets("Exp_insumos e empregados").Select
Range("B2").PasteSpecial xlPasteValues
'Copia e cola as s�ries
Sheets("Indicadores").Select
Range(Cells(147, 3), Cells(147, coluna)).Copy
Sheets("Exp_insumos e empregados").Select
Range("B3").PasteSpecial xlPasteValues
'Atribui o valor da m�dia
Range(Cells(4, 2), Cells(4, coluna - 1)).Value = "50"

'Cria o gr�fico
Set Gr�fico = Sheets("Exp_insumos e empregados").Shapes.AddChart2

Gr�fico.Select ' Seleciona o Gr�fico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gr�fico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gr�fico
    ActiveChart.Parent.Top = Parent.Range("A6").Top 'reposiciona o grafico em rela��o ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("A65").Left ' reposiciona o gr�fico em rela��o � borda esquerda da planilha
    ActiveChart.FullSeriesCollection(1).Name = "='Exp_insumos e empregados'!" & Cells(2, 1).Address 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(1).Values = "='Exp_insumos e empregados'!" & Range(Cells(2, coluna - 121), Cells(2, coluna - 1)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(1).XValues = "='Exp_insumos e empregados'!" & Range(Cells(1, coluna - 121), Cells(1, coluna - 1)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(2).Name = "='Exp_insumos e empregados'!" & Cells(3, 1).Address 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(2).Values = "='Exp_insumos e empregados'!" & Range(Cells(3, coluna - 121), Cells(3, coluna - 1)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(2).XValues = "='Exp_insumos e empregados'!" & Range(Cells(1, coluna - 121), Cells(1, coluna - 1)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(3).Name = "='Exp_insumos e empregados'!" & Cells(4, 1).Address 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(3).Values = "='Exp_insumos e empregados'!" & Range(Cells(4, coluna - 121), Cells(4, coluna - 1)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(3).XValues = "='Exp_insumos e empregados'!" & Range(Cells(1, coluna - 121), Cells(1, coluna - 1)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Sondagem Ind�stria da Constru��o\Automa��o\Templates\Expectativa - Atividade Empreendimentos - Constru��o") ' Aplica o template do gr�fico
    
Set Gr�fico_edit = Sheets("Exp_insumos e empregados").Shapes.AddChart2

Gr�fico_edit.Select ' Seleciona o Gr�fico
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Sondagem Ind�stria da Constru��o\Automa��o\Templates\Expectativa - Atividade Empreendimentos - Constru��o_Edit") ' Aplica o template do gr�fico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gr�fico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gr�fico
    ActiveChart.Parent.Top = Parent.Range("K6").Top 'reposiciona o grafico em rela��o ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("K6").Left ' reposiciona o gr�fico em rela��o � borda esquerda da planilha
    ActiveChart.FullSeriesCollection(1).Name = "='Exp_insumos e empregados'!" & Cells(2, 1).Address 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(1).Values = "='Exp_insumos e empregados'!" & Range(Cells(2, coluna - 121), Cells(2, coluna - 1)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(1).XValues = "='Exp_insumos e empregados'!" & Range(Cells(1, coluna - 121), Cells(1, coluna - 1)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(2).Name = "='Exp_insumos e empregados'!" & Cells(3, 1).Address 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(2).Values = "='Exp_insumos e empregados'!" & Range(Cells(3, coluna - 121), Cells(3, coluna - 1)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(2).XValues = "='Exp_insumos e empregados'!" & Range(Cells(1, coluna - 121), Cells(1, coluna - 1)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
End Sub

Sub problemas_ponderado()

Dim V As Integer 'Numero do trimestre mais recente
Dim X As Integer 'N�mero do trimestre anterior
Dim GrafProblemas As Object ' Gr�fico

Sheets("problemas_ponderado").Select ' Seleciona a aba Principais_Problemas
ActiveSheet.Range("B11:B28").Copy ActiveSheet.Range("B165") ' Copia e cola o nome das categorias menos outros e nehum.

V = Sheets("problemas_ponderado").Range("C8").End(xlToRight).Column 'Define o n�mero da �ltima coluna
X = V - 1 'Define o n�mero da primeira coluna

ActiveSheet.Range(Cells(11, X), Cells(28, V)).Copy ActiveSheet.Range("C165") 'Copia os valores para formar a tabela
ActiveSheet.Range(Cells(8, X), Cells(8, V)).Copy ActiveSheet.Range("C164") ' copia o nome dos trimestres

'Filtra os valores na tabela de forma decrescente de acordo com o trimestre mais recente
ActiveSheet.Range("B164:D164").Select
Selection.AutoFilter
ActiveSheet.AutoFilter.Sort.SortFields. _
        Add2 Key:=Range("D164"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
        
ActiveSheet.Range("B10").Copy ActiveSheet.Range("B183") 'Copia o nome das categorias outros e nenhum na tabela
ActiveSheet.Range(Cells(10, X), Cells(10, V)).Copy ActiveSheet.Range("C183") ' copia os valores das categorias outros e nenhum na tabela
ActiveSheet.Range("B29").Copy ActiveSheet.Range("B184") 'Copia o nome das categorias outros e nenhum na tabela
ActiveSheet.Range(Cells(29, X), Cells(29, V)).Copy ActiveSheet.Range("C184") ' copia os valores das categorias outros e nenhum na tabela



'Adiciona a aba Desempenho
Sheets.Add(Before:=Sheets("Indicadores")).Name = "Problemas"

'Copia e cola as s�ries
Sheets("problemas_ponderado").Select
Range("B164:D184").Copy
Sheets("Problemas").Select
Range("B2").PasteSpecial

Set GrafProblemas = Sheets("Problemas").Shapes.AddChart2 'Adiciona o gr�fico

GrafProblemas.Select ' Seleciona o Gr�fico
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Sondagem Industrial\Automa��o\Templates\Principais Problemas.crtx") ' Aplica o template do gr�fico
    ActiveChart.Parent.Height = 630 'ajusta a altura do gr�fico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gr�fico
    ActiveChart.Parent.Top = Parent.Range("F2").Top 'reposiciona o grafico em rela��o ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("F2").Left ' reposiciona o gr�fico em rela��o � borda esquerda da planilha
    
    ActiveChart.FullSeriesCollection(1).Name = "='Problemas'!$D$2" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(1).Values = "='Problemas'!$D$3:$D$22" 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(1).XValues = "='Problemas'!$B$3:$B$22" 'determina os valores referentes ao eixo x da s�rie adicionada
    
    ActiveChart.FullSeriesCollection(2).Name = "='Problemas'!$C$2" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(2).Values = "='Problemas'!$C$3:$C$22" 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(2).XValues = "='Problemas'!$B$3:$B$22" 'determina os valores referentes ao eixo x da s�rie adicionada
End Sub

Sub condi��es_financeiras()
'Define as vari�veis
Dim coluna As Integer
Dim coluna_Pre�o As Integer
Dim Gr�fico_Lucro_Situa��o As Object
Dim Gr�fico_Cr�dito As Object
Dim Gr�fico_Pre�o As Object
'Seleciona a aba indicadores
Sheets("Indicadores").Select
'pega o n�mero da �ltima coluna
coluna = Range("C66").End(xlToRight).Column
coluna_Pre�o = Range("C175").End(xlToRight).Column
'Adiciona a aba Situa��o Financiera
Sheets.Add(Before:=Sheets("Indicadores")).Name = "condicoes financeiras"
'Seleciona a aba situa��o fincanceira
Sheets("condicoes financeiras").Select
'Escreve os primeiros trimestres da serie
Range("B1").Value = "IV-09"
Range("C1").Value = "I-10"
Range("D1").Value = "II-10"
Range("E1").Value = "III-10"
Range("F1").Value = "IV-10"
'Completa a s�rie dos trimestres
Range("C1:F1").Select
    Selection.NumberFormat = "mmm-yy"
    Selection.AutoFill Destination:=Range(Cells(1, 3), Cells(1, coluna - 1)), Type:=xlFillDefault
'Nomeia as s�ries
Range("A2").Value = "Lucro operacional"
Range("A3").Value = "Situa��o financeira"
Range("A4").Value = "Facilidade de acesso ao cr�dito"
Range("A5").Value = "Pre�o m�dio dos insumos e mat�rias-primas"
Range("A6").Value = "Linha divis�ria"
'Copia e cola os valores da s�rie
Sheets("Indicadores").Select
Range(Cells(66, 3), Cells(66, coluna)).Copy
Sheets("condicoes financeiras").Select
Range("B2").PasteSpecial xlPasteValues

Sheets("Indicadores").Select
Range(Cells(79, 3), Cells(79, coluna)).Copy
Sheets("condicoes financeiras").Select
Range("B3").PasteSpecial xlPasteValues

Sheets("Indicadores").Select
Range(Cells(92, 3), Cells(92, coluna)).Copy
Sheets("condicoes financeiras").Select
Range("B4").PasteSpecial xlPasteValues

Sheets("Indicadores").Select
Range(Cells(175, 3), Cells(175, coluna_Pre�o)).Copy
Sheets("condicoes financeiras").Select
Range("K5").PasteSpecial xlPasteValues
'Atribui os valores da m�dia
Range(Cells(6, 2), Cells(6, coluna - 1)).Value = "50"


Range("A1").Value = Data
Range("C300").Select

'Adiciona o gr�fico
Set Gr�fico_Lucro_Situa��o = Sheets("condicoes financeiras").Shapes.AddChart2

Gr�fico_Lucro_Situa��o.Select ' Seleciona o Gr�fico
        ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Sondagem Ind�stria da Constru��o\Automa��o\Templates\Lucro e situa��o financeira_Constru��o") ' Aplica o template do gr�fico

    ActiveChart.Parent.Height = 300 'ajusta a altura do gr�fico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gr�fico
    ActiveChart.Parent.Top = Parent.Range("A8").Top 'reposiciona o grafico em rela��o ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("A8").Left ' reposiciona o gr�fico em rela��o � borda esquerda da planilha
     'adiciona uma nova s�rie ao gr�fico
    ActiveChart.FullSeriesCollection(1).Name = "='condicoes financeiras'!$A$2" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(1).Values = "='condicoes financeiras'!" & Range(Cells(2, coluna - 40), Cells(2, coluna - 1)).Address
    ActiveChart.FullSeriesCollection(1).XValues = "='condicoes financeiras'!" & Range(Cells(1, coluna - 40), Cells(1, coluna - 1)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
     'adiciona uma nova s�rie ao gr�fico
    ActiveChart.FullSeriesCollection(2).Name = "='condicoes financeiras'!$A$3" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(2).Values = "='condicoes financeiras'!" & Range(Cells(3, coluna - 40), Cells(3, coluna - 1)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(2).XValues = "='condicoes financeiras'!" & Range(Cells(1, coluna - 40), Cells(1, coluna - 1)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
     'adiciona uma nova s�rie ao gr�fico
    ActiveChart.FullSeriesCollection(3).Name = "='condicoes financeiras'!$A$6" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(3).Values = "='condicoes financeiras'!" & Range(Cells(6, coluna - 40), Cells(6, coluna - 1)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(3).XValues = "='condicoes financeiras'!" & Range(Cells(1, coluna - 40), Cells(1, coluna - 1)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
 

Set Gr�fico_Cr�dito = Sheets("condicoes financeiras").Shapes.AddChart2

Gr�fico_Cr�dito.Select ' Seleciona o Gr�fico
ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Sondagem Ind�stria da Constru��o\Automa��o\Templates\Lucro e situa��o financeira_Constru��o") ' Aplica o template do gr�fico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gr�fico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gr�fico
    ActiveChart.Parent.Top = Parent.Range("J8").Top 'reposiciona o grafico em rela��o ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("J8").Left ' reposiciona o gr�fico em rela��o � borda esquerda da planilha
     'adiciona uma nova s�rie ao gr�fico
    ActiveChart.FullSeriesCollection(1).Name = "='condicoes financeiras'!$A$4" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(1).Values = "='condicoes financeiras'!" & Range(Cells(4, coluna - 40), Cells(4, coluna - 1)).Address
    ActiveChart.FullSeriesCollection(1).XValues = "='condicoes financeiras'!" & Range(Cells(1, coluna - 40), Cells(1, coluna - 1)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
     'adiciona uma nova s�rie ao gr�fico
    ActiveChart.FullSeriesCollection(2).Name = "='condicoes financeiras'!$A$6" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(2).Values = "='condicoes financeiras'!" & Range(Cells(6, coluna - 40), Cells(6, coluna - 1)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(2).XValues = "='condicoes financeiras'!" & Range(Cells(1, coluna - 40), Cells(1, coluna - 1)).Address 'determina os valores referentes ao eixo x da s�rie adicionada

Set Gr�fico_Pre�o = Sheets("condicoes financeiras").Shapes.AddChart2

Gr�fico_Pre�o.Select ' Seleciona o Gr�fico
ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Sondagem Ind�stria da Constru��o\Automa��o\Templates\Cr�dito_Constru��o") ' Aplica o template do gr�fico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gr�fico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gr�fico
    ActiveChart.Parent.Top = Parent.Range("S8").Top 'reposiciona o grafico em rela��o ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("S8").Left ' reposiciona o gr�fico em rela��o � borda esquerda da planilha
     'adiciona uma nova s�rie ao gr�fico
    ActiveChart.FullSeriesCollection(1).Name = "='condicoes financeiras'!$A$5" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(1).Values = "='condicoes financeiras'!" & Range(Cells(5, coluna - 40), Cells(5, coluna - 1)).Address
    ActiveChart.FullSeriesCollection(1).XValues = "='condicoes financeiras'!" & Range(Cells(1, coluna - 40), Cells(1, coluna - 1)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
     'adiciona uma nova s�rie ao gr�fico
    ActiveChart.FullSeriesCollection(2).Name = "='condicoes financeiras'!$A$6" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(2).Values = "='condicoes financeiras'!" & Range(Cells(6, coluna - 40), Cells(6, coluna - 1)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(2).XValues = "='condicoes financeiras'!" & Range(Cells(1, coluna - 40), Cells(1, coluna - 1)).Address 'determina os valores referentes ao eixo x da s�rie adicionada


Range("A1").Select
End Sub
Sub Ivestimento_Constru��o()

Dim coluna As Integer
Dim Gr�fico As Object
Dim Gr�fico_edit As Object
Dim m�dia As Integer

'Seleciona a aba Indicadores
Sheets("Indicadores").Select
'Pega o n�mero da pultima coluna
coluna = Range("C189").End(xlToRight).Column
'Calcula a m�dia hist�rica
m�dia = Application.WorksheetFunction.Average(Range(Cells(189, 3), Cells(189, coluna)))
'Adiciona a aba Desempenho
Sheets.Add(Before:=Sheets("Indicadores")).Name = "Inten��o de investimento"
'Seleciona a aba desempenho
Sheets("Inten��o de investimento").Select
'Escreve os dois prmeiros meses da s�rie
Range("B1").Value = "11/01/13"
Range("C1").Value = "12/01/13"
Range("D1").Value = "01/01/14"
Range("E1").Value = "02/01/14"
'Completa a linha dos meses at� o �ltimo m�s
Range("D1:E1").Select
    Selection.NumberFormat = "mmm-yy"
    Selection.AutoFill Destination:=Range(Cells(1, 4), Cells(1, coluna - 1)), Type:=xlFillDefault
    
'Nomeia as s�ries
Range("A2").Value = "Inten��o de investimento"
Range("A3").Value = "M�dia hist�rica"

'Copia e cola as s�ries
Sheets("Indicadores").Select
Range(Cells(189, 3), Cells(189, coluna)).Copy
Sheets("Inten��o de investimento").Select
Range("B2").PasteSpecial xlPasteValues
'Atribui o valor da m�dia
Range(Cells(3, 2), Cells(3, coluna - 1)).Value = m�dia

'Cria o gr�fico
Set Gr�fico = Sheets("Inten��o de investimento").Shapes.AddChart2

Gr�fico.Select ' Seleciona o Gr�fico
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Sondagem Ind�stria da Constru��o\Automa��o\Templates\Investimento_Constru��o") ' Aplica o template do gr�fico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gr�fico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gr�fico
    ActiveChart.Parent.Top = Parent.Range("A7").Top 'reposiciona o grafico em rela��o ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("A7").Left ' reposiciona o gr�fico em rela��o � borda esquerda da planilha
    ActiveChart.FullSeriesCollection(1).Name = "='Inten��o de investimento'!" & Cells(2, 1).Address 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(1).Values = "='Inten��o de investimento'!" & Range(Cells(2, coluna - 85), Cells(2, coluna - 1)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(1).XValues = "='Inten��o de investimento'!" & Range(Cells(1, coluna - 85), Cells(1, coluna - 1)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(2).Name = "='Inten��o de investimento'!$A$3" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(2).Values = "='Inten��o de investimento'!" & Range(Cells(3, coluna - 85), Cells(3, coluna - 1)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(2).XValues = "='Inten��o de investimento'!" & Range(Cells(1, coluna - 85), Cells(1, coluna - 1)).Address 'determina os valores referentes ao eixo x da s�rie adicionada

    
Set Gr�fico_edit = Sheets("Inten��o de investimento").Shapes.AddChart2

Gr�fico_edit.Select ' Seleciona o Gr�fico
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Sondagem Ind�stria da Constru��o\Automa��o\Templates\Investimento_Constru��o_edit") ' Aplica o template do gr�fico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gr�fico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gr�fico
    ActiveChart.Parent.Top = Parent.Range("K7").Top 'reposiciona o grafico em rela��o ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("K7").Left ' reposiciona o gr�fico em rela��o � borda esquerda da planilha
    ActiveChart.FullSeriesCollection(1).Name = "='Inten��o de investimento'!" & Cells(2, 1).Address 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(1).Values = "='Inten��o de investimento'!" & Range(Cells(2, coluna - 85), Cells(2, coluna - 1)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(1).XValues = "='Inten��o de investimento'!" & Range(Cells(1, coluna - 85), Cells(1, coluna - 1)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
  
    Range("A1").Select
End Sub

Sub Tabelas()

Sheets.Add(Before:=Sheets("Indicadores")).Name = "TABELAS"

Sheets("TABELAS").Select

Range("A1").Value = "Resultados por porte e setor"


Range("A2").Value = "Desempenho da ind�stria da constru��o"

Range("B3").Value = "UCO (%)1"
Range("E3").Value = "�NDICE DE EVOLU��O DO N�VEL DE ATIVIDADE2"
Range("H3").Value = "�NDICE DE N�VEL DE ATIVIDADE EFETIVO EM RELA��O AO USUAL3"
Range("K3").Value = "�NDICE DE EVOLU��O DO N�MERO DE EMPREGADOS2"
Range("B3:D3").Merge
Range("E3:G3").Merge
Range("H3:J3").Merge
Range("M3:K3").Merge


Range("A10").Value = "Expectativas da ind�stria da constru��o"

Range("B11").Value = "�NDICES DE EXPECTATIVAS4"
Range("B12").Value = "N�VEL DE ATIVIDADE"
Range("E12").Value = "NOVOS EMPREENDIMENTOS E SERVI�OS"
Range("H12").Value = "COMPRA DE INSUMOS E MAT�RIAS-PRIMAS"
Range("K12").Value = "N�MERO DE EMPREGADOS"
Range("N11").Value = "INTEN��O DE INVESTIMENTO5"
Range("B11:M11").Merge
Range("B11:D11").Merge
Range("E11:G11").Merge
Range("H11:J11").Merge
Range("K11:M11").Merge
Range("N11:P11").Merge


Range("A19").Value = "�ndice de Confian�a do Empres�rio da Ind�stria da Constru��o e seus componentes"

Range("B20").Value = "ICEI - CONSTRU��O6"
Range("E20").Value = "�NDICE DE CONDI��ES ATUAIS7"
Range("H20").Value = "�NDICE DE EXPECTATIVAS8"
Range("B20:D20").Merge
Range("E20:G20").Merge
Range("H20:J20").Merge


Range("A28").Value = "Condi��es financeiras no trimestre"

Range("B29").Value = "MARGEM DE LUCRO OPERACIONAL"
Range("E29").Value = "PRE�O M�DIO DAS MAT�RIAS-PRIMAS"
Range("H29").Value = "SITUA��O FINANCEIRA"
Range("K29").Value = "ACESSO AO CR�DITO"
Range("B29:D29").Merge
Range("E29:G29").Merge
Range("H29:J29").Merge
Range("K29:M29").Merge


Range("A5").Value = "CONSTRU��O"
Range("A6").Value = "PEQUENA"
Range("A7").Value = "M�DIA"
Range("A8").Value = "GRANDE"
Sheets("TABELAS").Range("A5:A8").Copy (Sheets("TABELAS").Range("A14"))
Sheets("TABELAS").Range("A5:A8").Copy (Sheets("TABELAS").Range("A22"))
Sheets("TABELAS").Range("A5:A8").Copy (Sheets("TABELAS").Range("A31"))


Range("V1").Value = "Problemas Principais"

Range("V3").Value = "Itens"
Range("W3").Value = "Geral"
Range("Z3").Value = "Pequenas"
Range("AC3").Value = "M�dias"
Range("AF3").Value = "Grandes"

Range("Y4").Value = "Posi��o"
Range("AB4").Value = "Posi��o"
Range("AE4").Value = "Posi��o"
Range("AH4").Value = "Posi��o"

Sheets("condicoes financeiras").Select
ultimaC = Range("B1").End(xlToRight).Column
Range(Cells(1, ultimaC - 1), Cells(1, ultimaC)).Copy
Sheets("TABELAS").Select
Range("W4").PasteSpecial (xlPasteValues)
Sheets("TABELAS").Range("W4:X4").Copy (Sheets("TABELAS").Range("Z4"))
Sheets("TABELAS").Range("W4:X4").Copy (Sheets("TABELAS").Range("AC4"))
Sheets("TABELAS").Range("W4:X4").Copy (Sheets("TABELAS").Range("AF4"))

'Define as variavies que ser�o usadas para preencher as celulas
Coluna_Atividade_1 = Sheets("Produ��o").Range("B1").End(xlToRight).Column
Coluna_Atividade_2 = Coluna_Atividade_1 - 1
Coluna_Atividade_3 = Coluna_Atividade_1 - 12

'Define, atribui e copia e cola as datas
Datas_1 = Sheets("Produ��o").Cells(1, Coluna_Atividade_1).Value
Datas_2 = Sheets("Produ��o").Cells(1, Coluna_Atividade_2).Value
Datas_3 = Sheets("Produ��o").Cells(1, Coluna_Atividade_3).Value

Sheets("TABELAS").Cells(4, 2).Value = Datas_3
Sheets("TABELAS").Cells(4, 3).Value = Datas_2
Sheets("TABELAS").Cells(4, 4).Value = Datas_1

Sheets("TABELAS").Range("B4:D4").Copy (Sheets("TABELAS").Range("E4:G4"))
Sheets("TABELAS").Range("B4:D4").Copy (Sheets("TABELAS").Range("H4:J4"))
Sheets("TABELAS").Range("B4:D4").Copy (Sheets("TABELAS").Range("K4:M4"))

ColunaUCO_1 = Sheets("Indicadores").Range("C161").End(xlToRight).Column
ColunaUCO_2 = ColunaUCO_1 - 1
ColunaUCO_3 = ColunaUCO_1 - 12

'Atribui os valores da coluna UCO
'Constru��o
ValoresC_1 = Sheets("Indicadores").Cells(161, ColunaUCO_1).Value
ValoresC_2 = Sheets("Indicadores").Cells(161, ColunaUCO_2).Value
ValoresC_3 = Sheets("Indicadores").Cells(161, ColunaUCO_3).Value
Sheets("TABELAS").Cells(5, 2).Value = ValoresC_3
Sheets("TABELAS").Cells(5, 3).Value = ValoresC_2
Sheets("TABELAS").Cells(5, 4).Value = ValoresC_1
'Pequena
ValoresP_1 = Sheets("Indicadores").Cells(163, ColunaUCO_1).Value
ValoresP_2 = Sheets("Indicadores").Cells(163, ColunaUCO_2).Value
ValoresP_3 = Sheets("Indicadores").Cells(163, ColunaUCO_3).Value
Sheets("TABELAS").Cells(6, 2).Value = ValoresP_3
Sheets("TABELAS").Cells(6, 3).Value = ValoresP_2
Sheets("TABELAS").Cells(6, 4).Value = ValoresP_1
'M�dia
ValoresM_1 = Sheets("Indicadores").Cells(164, ColunaUCO_1).Value
ValoresM_2 = Sheets("Indicadores").Cells(164, ColunaUCO_2).Value
ValoresM_3 = Sheets("Indicadores").Cells(164, ColunaUCO_3).Value
Sheets("TABELAS").Cells(7, 2).Value = ValoresM_3
Sheets("TABELAS").Cells(7, 3).Value = ValoresM_2
Sheets("TABELAS").Cells(7, 4).Value = ValoresM_1
'Grande
ValoresG_1 = Sheets("Indicadores").Cells(165, ColunaUCO_1).Value
ValoresG_2 = Sheets("Indicadores").Cells(165, ColunaUCO_2).Value
ValoresG_3 = Sheets("Indicadores").Cells(165, ColunaUCO_3).Value
Sheets("TABELAS").Cells(8, 2).Value = ValoresG_3
Sheets("TABELAS").Cells(8, 3).Value = ValoresG_2
Sheets("TABELAS").Cells(8, 4).Value = ValoresG_1


'Atribui os valores da coluna Atividade
Coluna_Atividade_1 = Sheets("Indicadores").Range("C10").End(xlToRight).Column
Coluna_Atividade_2 = Coluna_Atividade_1 - 1
Coluna_Atividade_3 = Coluna_Atividade_1 - 12

'Constru��o
ValoresC_1 = Sheets("Indicadores").Cells(11, Coluna_Atividade_1).Value
ValoresC_2 = Sheets("Indicadores").Cells(11, Coluna_Atividade_2).Value
ValoresC_3 = Sheets("Indicadores").Cells(11, Coluna_Atividade_3).Value
Sheets("TABELAS").Cells(5, 5).Value = ValoresC_3
Sheets("TABELAS").Cells(5, 6).Value = ValoresC_2
Sheets("TABELAS").Cells(5, 7).Value = ValoresC_1
'Pequena
ValoresP_1 = Sheets("Indicadores").Cells(13, Coluna_Atividade_1).Value
ValoresP_2 = Sheets("Indicadores").Cells(13, Coluna_Atividade_2).Value
ValoresP_3 = Sheets("Indicadores").Cells(13, Coluna_Atividade_3).Value
Sheets("TABELAS").Cells(6, 5).Value = ValoresP_3
Sheets("TABELAS").Cells(6, 6).Value = ValoresP_2
Sheets("TABELAS").Cells(6, 7).Value = ValoresP_1
'M�dia
ValoresM_1 = Sheets("Indicadores").Cells(14, Coluna_Atividade_1).Value
ValoresM_2 = Sheets("Indicadores").Cells(14, Coluna_Atividade_2).Value
ValoresM_3 = Sheets("Indicadores").Cells(14, Coluna_Atividade_3).Value
Sheets("TABELAS").Cells(7, 5).Value = ValoresM_3
Sheets("TABELAS").Cells(7, 6).Value = ValoresM_2
Sheets("TABELAS").Cells(7, 7).Value = ValoresM_1
'Grande
ValoresG_1 = Sheets("Indicadores").Cells(15, Coluna_Atividade_1).Value
ValoresG_2 = Sheets("Indicadores").Cells(15, Coluna_Atividade_2).Value
ValoresG_3 = Sheets("Indicadores").Cells(15, Coluna_Atividade_3).Value
Sheets("TABELAS").Cells(8, 5).Value = ValoresG_3
Sheets("TABELAS").Cells(8, 6).Value = ValoresG_2
Sheets("TABELAS").Cells(8, 7).Value = ValoresG_1

'Atribui os valores da coluna Atividade Efetiva Ususal
Coluna_AtividadeEU_1 = Sheets("Indicadores").Range("C25").End(xlToRight).Column
Coluna_AtividadeEU_2 = Coluna_AtividadeEU_1 - 1
Coluna_AtividadeEU_3 = Coluna_AtividadeEU_1 - 12

'Constru��o
ValoresC_1 = Sheets("Indicadores").Cells(25, Coluna_AtividadeEU_1).Value
ValoresC_2 = Sheets("Indicadores").Cells(25, Coluna_AtividadeEU_2).Value
ValoresC_3 = Sheets("Indicadores").Cells(25, Coluna_AtividadeEU_3).Value
Sheets("TABELAS").Cells(5, 8).Value = ValoresC_3
Sheets("TABELAS").Cells(5, 9).Value = ValoresC_2
Sheets("TABELAS").Cells(5, 10).Value = ValoresC_1
'Pequena
ValoresP_1 = Sheets("Indicadores").Cells(27, Coluna_AtividadeEU_1).Value
ValoresP_2 = Sheets("Indicadores").Cells(27, Coluna_AtividadeEU_2).Value
ValoresP_3 = Sheets("Indicadores").Cells(27, Coluna_AtividadeEU_3).Value
Sheets("TABELAS").Cells(6, 8).Value = ValoresP_3
Sheets("TABELAS").Cells(6, 9).Value = ValoresP_2
Sheets("TABELAS").Cells(6, 10).Value = ValoresP_1
'M�dia
ValoresM_1 = Sheets("Indicadores").Cells(28, Coluna_AtividadeEU_1).Value
ValoresM_2 = Sheets("Indicadores").Cells(28, Coluna_AtividadeEU_2).Value
ValoresM_3 = Sheets("Indicadores").Cells(28, Coluna_AtividadeEU_3).Value
Sheets("TABELAS").Cells(7, 8).Value = ValoresM_3
Sheets("TABELAS").Cells(7, 9).Value = ValoresM_2
Sheets("TABELAS").Cells(7, 10).Value = ValoresM_1
'Grande
ValoresG_1 = Sheets("Indicadores").Cells(29, Coluna_AtividadeEU_1).Value
ValoresG_2 = Sheets("Indicadores").Cells(29, Coluna_AtividadeEU_2).Value
ValoresG_3 = Sheets("Indicadores").Cells(29, Coluna_AtividadeEU_3).Value
Sheets("TABELAS").Cells(8, 8).Value = ValoresG_3
Sheets("TABELAS").Cells(8, 9).Value = ValoresG_2
Sheets("TABELAS").Cells(8, 10).Value = ValoresG_1

'Atribui os valores da coluna Empregados
Coluna_Empregados_1 = Sheets("Indicadores").Range("C39").End(xlToRight).Column
Coluna_Empregados_2 = Coluna_Empregados_1 - 1
Coluna_Empregados_3 = Coluna_Empregados_1 - 12

'Constru��o
ValoresC_1 = Sheets("Indicadores").Cells(39, Coluna_Empregados_1).Value
ValoresC_2 = Sheets("Indicadores").Cells(39, Coluna_Empregados_2).Value
ValoresC_3 = Sheets("Indicadores").Cells(39, Coluna_Empregados_3).Value
Sheets("TABELAS").Cells(5, 11).Value = ValoresC_3
Sheets("TABELAS").Cells(5, 12).Value = ValoresC_2
Sheets("TABELAS").Cells(5, 13).Value = ValoresC_1
'Pequena
ValoresP_1 = Sheets("Indicadores").Cells(41, Coluna_Empregados_1).Value
ValoresP_2 = Sheets("Indicadores").Cells(41, Coluna_Empregados_2).Value
ValoresP_3 = Sheets("Indicadores").Cells(41, Coluna_Empregados_3).Value
Sheets("TABELAS").Cells(6, 11).Value = ValoresP_3
Sheets("TABELAS").Cells(6, 12).Value = ValoresP_2
Sheets("TABELAS").Cells(6, 13).Value = ValoresP_1
'M�dia
ValoresM_1 = Sheets("Indicadores").Cells(42, Coluna_Empregados_1).Value
ValoresM_2 = Sheets("Indicadores").Cells(42, Coluna_Empregados_2).Value
ValoresM_3 = Sheets("Indicadores").Cells(42, Coluna_Empregados_3).Value
Sheets("TABELAS").Cells(7, 11).Value = ValoresM_3
Sheets("TABELAS").Cells(7, 12).Value = ValoresM_2
Sheets("TABELAS").Cells(7, 13).Value = ValoresM_1
'Grande
ValoresG_1 = Sheets("Indicadores").Cells(43, Coluna_Empregados_1).Value
ValoresG_2 = Sheets("Indicadores").Cells(43, Coluna_Empregados_2).Value
ValoresG_3 = Sheets("Indicadores").Cells(43, Coluna_Empregados_3).Value
Sheets("TABELAS").Cells(8, 11).Value = ValoresG_3
Sheets("TABELAS").Cells(8, 12).Value = ValoresG_2
Sheets("TABELAS").Cells(8, 13).Value = ValoresG_1

'*************************************************** C�digo da parte de Expectativas **********************************************************

'Define as variavies que ser�o usadas para preencher as celulas
Coluna_NAtividade_1 = Sheets("Exp_atividade e empreendimentos").Range("B1").End(xlToRight).Column
Coluna_NAtividade_2 = Coluna_NAtividade_1 - 1
Coluna_NAtividade_3 = Coluna_NAtividade_1 - 12

'Define, atribui e copia e cola as datas
Data_1 = Sheets("Exp_atividade e empreendimentos").Cells(1, Coluna_NAtividade_1).Value
Data_2 = Sheets("Exp_atividade e empreendimentos").Cells(1, Coluna_NAtividade_2).Value
Data_3 = Sheets("Exp_atividade e empreendimentos").Cells(1, Coluna_NAtividade_3).Value

Sheets("TABELAS").Cells(13, 2).Value = Data_3
Sheets("TABELAS").Cells(13, 3).Value = Data_2
Sheets("TABELAS").Cells(13, 4).Value = Data_1

Sheets("TABELAS").Range("B13:D13").Copy (Sheets("TABELAS").Range("E13:G13"))
Sheets("TABELAS").Range("B13:D13").Copy (Sheets("TABELAS").Range("H13:J13"))
Sheets("TABELAS").Range("B13:D13").Copy (Sheets("TABELAS").Range("K13:M13"))
Sheets("TABELAS").Range("B13:D13").Copy (Sheets("TABELAS").Range("N13:P13"))

Coluna_NAtividade_1 = Sheets("Indicadores").Range("C105").End(xlToRight).Column
Coluna_NAtividade_2 = Coluna_NAtividade_1 - 1
Coluna_NAtividade_3 = Coluna_NAtividade_1 - 12

'Atribui os valores da coluna N�vel de Atividade
'Constru��o
ValoresC_1 = Sheets("Indicadores").Cells(105, Coluna_NAtividade_1).Value
ValoresC_2 = Sheets("Indicadores").Cells(105, Coluna_NAtividade_2).Value
ValoresC_3 = Sheets("Indicadores").Cells(105, Coluna_NAtividade_3).Value
Sheets("TABELAS").Cells(14, 2).Value = ValoresC_3
Sheets("TABELAS").Cells(14, 3).Value = ValoresC_2
Sheets("TABELAS").Cells(14, 4).Value = ValoresC_1
'Pequena
ValoresP_1 = Sheets("Indicadores").Cells(107, Coluna_NAtividade_1).Value
ValoresP_2 = Sheets("Indicadores").Cells(107, Coluna_NAtividade_2).Value
ValoresP_3 = Sheets("Indicadores").Cells(107, Coluna_NAtividade_3).Value
Sheets("TABELAS").Cells(15, 2).Value = ValoresP_3
Sheets("TABELAS").Cells(15, 3).Value = ValoresP_2
Sheets("TABELAS").Cells(15, 4).Value = ValoresP_1
'M�dia
ValoresM_1 = Sheets("Indicadores").Cells(108, Coluna_NAtividade_1).Value
ValoresM_2 = Sheets("Indicadores").Cells(108, Coluna_NAtividade_2).Value
ValoresM_3 = Sheets("Indicadores").Cells(108, Coluna_NAtividade_3).Value
Sheets("TABELAS").Cells(16, 2).Value = ValoresM_3
Sheets("TABELAS").Cells(16, 3).Value = ValoresM_2
Sheets("TABELAS").Cells(16, 4).Value = ValoresM_1
'Grande
ValoresG_1 = Sheets("Indicadores").Cells(109, Coluna_NAtividade_1).Value
ValoresG_2 = Sheets("Indicadores").Cells(109, Coluna_NAtividade_2).Value
ValoresG_3 = Sheets("Indicadores").Cells(109, Coluna_NAtividade_3).Value
Sheets("TABELAS").Cells(17, 2).Value = ValoresG_3
Sheets("TABELAS").Cells(17, 3).Value = ValoresG_2
Sheets("TABELAS").Cells(17, 4).Value = ValoresG_1


'Atribui os valores da coluna Novos empreendimentos
Coluna_Novos_1 = Sheets("Indicadores").Range("C119").End(xlToRight).Column
Coluna_Novos_2 = Coluna_Novos_1 - 1
Coluna_Novos_3 = Coluna_Novos_1 - 12

'Constru��o
ValoresC_1 = Sheets("Indicadores").Cells(119, Coluna_Novos_1).Value
ValoresC_2 = Sheets("Indicadores").Cells(119, Coluna_Novos_2).Value
ValoresC_3 = Sheets("Indicadores").Cells(119, Coluna_Novos_3).Value
Sheets("TABELAS").Cells(14, 5).Value = ValoresC_3
Sheets("TABELAS").Cells(14, 6).Value = ValoresC_2
Sheets("TABELAS").Cells(14, 7).Value = ValoresC_1
'Pequena
ValoresP_1 = Sheets("Indicadores").Cells(121, Coluna_Novos_1).Value
ValoresP_2 = Sheets("Indicadores").Cells(121, Coluna_Novos_2).Value
ValoresP_3 = Sheets("Indicadores").Cells(121, Coluna_Novos_3).Value
Sheets("TABELAS").Cells(15, 5).Value = ValoresP_3
Sheets("TABELAS").Cells(15, 6).Value = ValoresP_2
Sheets("TABELAS").Cells(15, 7).Value = ValoresP_1
'M�dia
ValoresM_1 = Sheets("Indicadores").Cells(122, Coluna_Novos_1).Value
ValoresM_2 = Sheets("Indicadores").Cells(122, Coluna_Novos_2).Value
ValoresM_3 = Sheets("Indicadores").Cells(122, Coluna_Novos_3).Value
Sheets("TABELAS").Cells(16, 5).Value = ValoresM_3
Sheets("TABELAS").Cells(16, 6).Value = ValoresM_2
Sheets("TABELAS").Cells(16, 7).Value = ValoresM_1
'Grande
ValoresG_1 = Sheets("Indicadores").Cells(123, Coluna_Novos_1).Value
ValoresG_2 = Sheets("Indicadores").Cells(123, Coluna_Novos_2).Value
ValoresG_3 = Sheets("Indicadores").Cells(123, Coluna_Novos_3).Value
Sheets("TABELAS").Cells(17, 5).Value = ValoresG_3
Sheets("TABELAS").Cells(17, 6).Value = ValoresG_2
Sheets("TABELAS").Cells(17, 7).Value = ValoresG_1

'Atribui os valores da coluna compra de insumos
Coluna_Compra_1 = Sheets("Indicadores").Range("C133").End(xlToRight).Column
Coluna_Compra_2 = Coluna_Compra_1 - 1
Coluna_Compra_3 = Coluna_Compra_1 - 12

'Constru��o
ValoresC_1 = Sheets("Indicadores").Cells(133, Coluna_Compra_1).Value
ValoresC_2 = Sheets("Indicadores").Cells(133, Coluna_Compra_2).Value
ValoresC_3 = Sheets("Indicadores").Cells(133, Coluna_Compra_3).Value
Sheets("TABELAS").Cells(14, 8).Value = ValoresC_3
Sheets("TABELAS").Cells(14, 9).Value = ValoresC_2
Sheets("TABELAS").Cells(14, 10).Value = ValoresC_1
'Pequena
ValoresP_1 = Sheets("Indicadores").Cells(135, Coluna_Compra_1).Value
ValoresP_2 = Sheets("Indicadores").Cells(135, Coluna_Compra_2).Value
ValoresP_3 = Sheets("Indicadores").Cells(135, Coluna_Compra_3).Value
Sheets("TABELAS").Cells(15, 8).Value = ValoresP_3
Sheets("TABELAS").Cells(15, 9).Value = ValoresP_2
Sheets("TABELAS").Cells(15, 10).Value = ValoresP_1
'M�dia
ValoresM_1 = Sheets("Indicadores").Cells(136, Coluna_Compra_1).Value
ValoresM_2 = Sheets("Indicadores").Cells(136, Coluna_Compra_2).Value
ValoresM_3 = Sheets("Indicadores").Cells(136, Coluna_Compra_3).Value
Sheets("TABELAS").Cells(16, 8).Value = ValoresM_3
Sheets("TABELAS").Cells(16, 9).Value = ValoresM_2
Sheets("TABELAS").Cells(16, 10).Value = ValoresM_1
'Grande
ValoresG_1 = Sheets("Indicadores").Cells(137, Coluna_Compra_1).Value
ValoresG_2 = Sheets("Indicadores").Cells(137, Coluna_Compra_2).Value
ValoresG_3 = Sheets("Indicadores").Cells(137, Coluna_Compra_3).Value
Sheets("TABELAS").Cells(17, 8).Value = ValoresG_3
Sheets("TABELAS").Cells(17, 9).Value = ValoresG_2
Sheets("TABELAS").Cells(17, 10).Value = ValoresG_1


'Atribui os valores da coluna empregados
Coluna_Empregados_1 = Sheets("Indicadores").Range("C147").End(xlToRight).Column
Coluna_Empregados_2 = Coluna_Empregados_1 - 1
Coluna_Empregados_3 = Coluna_Empregados_1 - 12

'Constru��o
ValoresC_1 = Sheets("Indicadores").Cells(147, Coluna_Empregados_1).Value
ValoresC_2 = Sheets("Indicadores").Cells(147, Coluna_Empregados_2).Value
ValoresC_3 = Sheets("Indicadores").Cells(147, Coluna_Empregados_3).Value
Sheets("TABELAS").Cells(14, 11).Value = ValoresC_3
Sheets("TABELAS").Cells(14, 12).Value = ValoresC_2
Sheets("TABELAS").Cells(14, 13).Value = ValoresC_1
'Pequena
ValoresP_1 = Sheets("Indicadores").Cells(149, Coluna_Empregados_1).Value
ValoresP_2 = Sheets("Indicadores").Cells(149, Coluna_Empregados_2).Value
ValoresP_3 = Sheets("Indicadores").Cells(149, Coluna_Empregados_3).Value
Sheets("TABELAS").Cells(15, 11).Value = ValoresP_3
Sheets("TABELAS").Cells(15, 12).Value = ValoresP_2
Sheets("TABELAS").Cells(15, 13).Value = ValoresP_1
'M�dia
ValoresM_1 = Sheets("Indicadores").Cells(150, Coluna_Empregados_1).Value
ValoresM_2 = Sheets("Indicadores").Cells(150, Coluna_Empregados_2).Value
ValoresM_3 = Sheets("Indicadores").Cells(150, Coluna_Empregados_3).Value
Sheets("TABELAS").Cells(16, 11).Value = ValoresM_3
Sheets("TABELAS").Cells(16, 12).Value = ValoresM_2
Sheets("TABELAS").Cells(16, 13).Value = ValoresM_1
'Grande
ValoresG_1 = Sheets("Indicadores").Cells(151, Coluna_Empregados_1).Value
ValoresG_2 = Sheets("Indicadores").Cells(151, Coluna_Empregados_2).Value
ValoresG_3 = Sheets("Indicadores").Cells(151, Coluna_Empregados_3).Value
Sheets("TABELAS").Cells(17, 11).Value = ValoresG_3
Sheets("TABELAS").Cells(17, 12).Value = ValoresG_2
Sheets("TABELAS").Cells(17, 13).Value = ValoresG_1

'Atribui os valores da coluna investimento
Coluna_Investimentos_1 = Sheets("Indicadores").Range("C189").End(xlToRight).Column
Coluna_Investimentos_2 = Coluna_Investimentos_1 - 1
Coluna_Investimentos_3 = Coluna_Investimentos_1 - 12

'Constru��o
ValoresC_1 = Sheets("Indicadores").Cells(189, Coluna_Investimentos_1).Value
ValoresC_2 = Sheets("Indicadores").Cells(189, Coluna_Investimentos_2).Value
ValoresC_3 = Sheets("Indicadores").Cells(189, Coluna_Investimentos_3).Value
Sheets("TABELAS").Cells(14, 14).Value = ValoresC_3
Sheets("TABELAS").Cells(14, 15).Value = ValoresC_2
Sheets("TABELAS").Cells(14, 16).Value = ValoresC_1
'Pequena
ValoresP_1 = Sheets("Indicadores").Cells(191, Coluna_Investimentos_1).Value
ValoresP_2 = Sheets("Indicadores").Cells(191, Coluna_Investimentos_2).Value
ValoresP_3 = Sheets("Indicadores").Cells(191, Coluna_Investimentos_3).Value
Sheets("TABELAS").Cells(15, 14).Value = ValoresP_3
Sheets("TABELAS").Cells(15, 15).Value = ValoresP_2
Sheets("TABELAS").Cells(15, 16).Value = ValoresP_1
'M�dia
ValoresM_1 = Sheets("Indicadores").Cells(192, Coluna_Investimentos_1).Value
ValoresM_2 = Sheets("Indicadores").Cells(192, Coluna_Investimentos_2).Value
ValoresM_3 = Sheets("Indicadores").Cells(192, Coluna_Investimentos_3).Value
Sheets("TABELAS").Cells(16, 14).Value = ValoresM_3
Sheets("TABELAS").Cells(16, 15).Value = ValoresM_2
Sheets("TABELAS").Cells(16, 16).Value = ValoresM_1
'Grande
ValoresG_1 = Sheets("Indicadores").Cells(193, Coluna_Investimentos_1).Value
ValoresG_2 = Sheets("Indicadores").Cells(193, Coluna_Investimentos_2).Value
ValoresG_3 = Sheets("Indicadores").Cells(193, Coluna_Investimentos_3).Value
Sheets("TABELAS").Cells(17, 14).Value = ValoresG_3
Sheets("TABELAS").Cells(17, 15).Value = ValoresG_2
Sheets("TABELAS").Cells(17, 16).Value = ValoresG_1

'*************************************************** C�digo da parte de ICEI **********************************************************
Coluna_ICEI_1 = Sheets("ICEI_Constru��o").Range("B1").End(xlToRight).Column
Coluna_ICEI_2 = Coluna_ICEI_1 - 1
Coluna_ICEI_3 = Coluna_ICEI_1 - 12

'Define, atribui e copia e cola as datas
Data_1 = Sheets("ICEI_Constru��o").Cells(1, Coluna_ICEI_1).Value
Data_2 = Sheets("ICEI_Constru��o").Cells(1, Coluna_ICEI_2).Value
Data_3 = Sheets("ICEI_Constru��o").Cells(1, Coluna_ICEI_3).Value

Sheets("TABELAS").Cells(21, 2).Value = Data_3
Sheets("TABELAS").Cells(21, 3).Value = Data_2
Sheets("TABELAS").Cells(21, 4).Value = Data_1

Sheets("TABELAS").Range("B21:D21").Copy (Sheets("TABELAS").Range("E21"))
Sheets("TABELAS").Range("B21:D21").Copy (Sheets("TABELAS").Range("H21"))

'Atribui os valores da coluna ICEI
Coluna_ICEI_1 = Sheets("Indicadores").Range("C203").End(xlToRight).Column
Coluna_ICEI_2 = Coluna_ICEI_1 - 1
Coluna_ICEI_3 = Coluna_ICEI_1 - 12

'Constru��o
ValoresC_1 = Sheets("Indicadores").Cells(203, Coluna_ICEI_1).Value
ValoresC_2 = Sheets("Indicadores").Cells(203, Coluna_ICEI_2).Value
ValoresC_3 = Sheets("Indicadores").Cells(203, Coluna_ICEI_3).Value
Sheets("TABELAS").Cells(22, 2).Value = ValoresC_3
Sheets("TABELAS").Cells(22, 3).Value = ValoresC_2
Sheets("TABELAS").Cells(22, 4).Value = ValoresC_1
'Pequena
ValoresP_1 = Sheets("Indicadores").Cells(205, Coluna_ICEI_1).Value
ValoresP_2 = Sheets("Indicadores").Cells(205, Coluna_ICEI_2).Value
ValoresP_3 = Sheets("Indicadores").Cells(205, Coluna_ICEI_3).Value
Sheets("TABELAS").Cells(23, 2).Value = ValoresP_3
Sheets("TABELAS").Cells(23, 3).Value = ValoresP_2
Sheets("TABELAS").Cells(23, 4).Value = ValoresP_1
'M�dia
ValoresM_1 = Sheets("Indicadores").Cells(206, Coluna_ICEI_1).Value
ValoresM_2 = Sheets("Indicadores").Cells(206, Coluna_ICEI_2).Value
ValoresM_3 = Sheets("Indicadores").Cells(206, Coluna_ICEI_3).Value
Sheets("TABELAS").Cells(24, 2).Value = ValoresM_3
Sheets("TABELAS").Cells(24, 3).Value = ValoresM_2
Sheets("TABELAS").Cells(24, 4).Value = ValoresM_1
'Grande
ValoresG_1 = Sheets("Indicadores").Cells(207, Coluna_ICEI_1).Value
ValoresG_2 = Sheets("Indicadores").Cells(207, Coluna_ICEI_2).Value
ValoresG_3 = Sheets("Indicadores").Cells(207, Coluna_ICEI_3).Value
Sheets("TABELAS").Cells(25, 2).Value = ValoresG_3
Sheets("TABELAS").Cells(25, 3).Value = ValoresG_2
Sheets("TABELAS").Cells(25, 4).Value = ValoresG_1


'Atribui os valores da coluna Condi��es
Coluna_Condi��es_1 = Sheets("Indicadores").Range("C217").End(xlToRight).Column
Coluna_Condi��es_2 = Coluna_Condi��es_1 - 1
Coluna_Condi��es_3 = Coluna_Condi��es_1 - 12

'Constru��o
ValoresC_1 = Sheets("Indicadores").Cells(217, Coluna_Condi��es_1).Value
ValoresC_2 = Sheets("Indicadores").Cells(217, Coluna_Condi��es_2).Value
ValoresC_3 = Sheets("Indicadores").Cells(217, Coluna_Condi��es_3).Value
Sheets("TABELAS").Cells(22, 5).Value = ValoresC_3
Sheets("TABELAS").Cells(22, 6).Value = ValoresC_2
Sheets("TABELAS").Cells(22, 7).Value = ValoresC_1
'Pequena
ValoresP_1 = Sheets("Indicadores").Cells(219, Coluna_Condi��es_1).Value
ValoresP_2 = Sheets("Indicadores").Cells(219, Coluna_Condi��es_2).Value
ValoresP_3 = Sheets("Indicadores").Cells(219, Coluna_Condi��es_3).Value
Sheets("TABELAS").Cells(23, 5).Value = ValoresP_3
Sheets("TABELAS").Cells(23, 6).Value = ValoresP_2
Sheets("TABELAS").Cells(23, 7).Value = ValoresP_1
'M�dia
ValoresM_1 = Sheets("Indicadores").Cells(220, Coluna_Condi��es_1).Value
ValoresM_2 = Sheets("Indicadores").Cells(220, Coluna_Condi��es_2).Value
ValoresM_3 = Sheets("Indicadores").Cells(220, Coluna_Condi��es_3).Value
Sheets("TABELAS").Cells(24, 5).Value = ValoresM_3
Sheets("TABELAS").Cells(24, 6).Value = ValoresM_2
Sheets("TABELAS").Cells(24, 7).Value = ValoresM_1
'Grande
ValoresG_1 = Sheets("Indicadores").Cells(221, Coluna_Condi��es_1).Value
ValoresG_2 = Sheets("Indicadores").Cells(221, Coluna_Condi��es_2).Value
ValoresG_3 = Sheets("Indicadores").Cells(221, Coluna_Condi��es_3).Value
Sheets("TABELAS").Cells(25, 5).Value = ValoresG_3
Sheets("TABELAS").Cells(25, 6).Value = ValoresG_2
Sheets("TABELAS").Cells(25, 7).Value = ValoresG_1



'Atribui os valores da coluna Expectativa
Coluna_Expectativa_1 = Sheets("Indicadores").Range("C260").End(xlToRight).Column
Coluna_Expectativa_2 = Coluna_Expectativa_1 - 1
Coluna_Expectativa_3 = Coluna_Expectativa_1 - 12

'Constru��o
ValoresC_1 = Sheets("Indicadores").Cells(260, Coluna_Expectativa_1).Value
ValoresC_2 = Sheets("Indicadores").Cells(260, Coluna_Expectativa_2).Value
ValoresC_3 = Sheets("Indicadores").Cells(260, Coluna_Expectativa_3).Value
Sheets("TABELAS").Cells(22, 8).Value = ValoresC_3
Sheets("TABELAS").Cells(22, 9).Value = ValoresC_2
Sheets("TABELAS").Cells(22, 10).Value = ValoresC_1
'Pequena
ValoresP_1 = Sheets("Indicadores").Cells(262, Coluna_Expectativa_1).Value
ValoresP_2 = Sheets("Indicadores").Cells(262, Coluna_Expectativa_2).Value
ValoresP_3 = Sheets("Indicadores").Cells(262, Coluna_Expectativa_3).Value
Sheets("TABELAS").Cells(23, 8).Value = ValoresP_3
Sheets("TABELAS").Cells(23, 9).Value = ValoresP_2
Sheets("TABELAS").Cells(23, 10).Value = ValoresP_1
'M�dia
ValoresM_1 = Sheets("Indicadores").Cells(263, Coluna_Expectativa_1).Value
ValoresM_2 = Sheets("Indicadores").Cells(263, Coluna_Expectativa_2).Value
ValoresM_3 = Sheets("Indicadores").Cells(263, Coluna_Expectativa_3).Value
Sheets("TABELAS").Cells(24, 8).Value = ValoresM_3
Sheets("TABELAS").Cells(24, 9).Value = ValoresM_2
Sheets("TABELAS").Cells(24, 10).Value = ValoresM_1
'Grande
ValoresG_1 = Sheets("Indicadores").Cells(264, Coluna_Expectativa_1).Value
ValoresG_2 = Sheets("Indicadores").Cells(264, Coluna_Expectativa_2).Value
ValoresG_3 = Sheets("Indicadores").Cells(264, Coluna_Expectativa_3).Value
Sheets("TABELAS").Cells(25, 8).Value = ValoresG_3
Sheets("TABELAS").Cells(25, 9).Value = ValoresG_2
Sheets("TABELAS").Cells(25, 10).Value = ValoresG_1


'*************************************************** C�digo da parte de Condi��es Financeiras **********************************************************

'Define as variavies que ser�o usadas para preencher as celulas
Coluna_Lucro_1 = Sheets("condicoes financeiras").Range("B1").End(xlToRight).Column
Coluna_Lucro_2 = Coluna_Lucro_1 - 1
Coluna_Lucro_3 = Coluna_Lucro_1 - 4

'Define, atribui e copia e cola as datas
Data_1 = Sheets("condicoes financeiras").Cells(1, Coluna_Lucro_1).Value
Data_2 = Sheets("condicoes financeiras").Cells(1, Coluna_Lucro_2).Value
Data_3 = Sheets("condicoes financeiras").Cells(1, Coluna_Lucro_3).Value

Sheets("TABELAS").Cells(30, 2).Value = Data_3
Sheets("TABELAS").Cells(30, 3).Value = Data_2
Sheets("TABELAS").Cells(30, 4).Value = Data_1

Sheets("TABELAS").Range("B30:D30").Copy (Sheets("TABELAS").Range("E30"))
Sheets("TABELAS").Range("B30:D30").Copy (Sheets("TABELAS").Range("H30"))
Sheets("TABELAS").Range("B30:D30").Copy (Sheets("TABELAS").Range("K30"))

'Atribui os valores da coluna margem de lucro operacional
Coluna_Lucro_1 = Sheets("Indicadores").Range("C66").End(xlToRight).Column
Coluna_Lucro_2 = Coluna_Lucro_1 - 1
Coluna_Lucro_3 = Coluna_Lucro_1 - 12

'Constru��o
ValoresC_1 = Sheets("Indicadores").Cells(66, Coluna_Lucro_1).Value
ValoresC_2 = Sheets("Indicadores").Cells(66, Coluna_Lucro_2).Value
ValoresC_3 = Sheets("Indicadores").Cells(66, Coluna_Lucro_3).Value
Sheets("TABELAS").Cells(31, 2).Value = ValoresC_3
Sheets("TABELAS").Cells(31, 3).Value = ValoresC_2
Sheets("TABELAS").Cells(31, 4).Value = ValoresC_1
'Pequena
ValoresP_1 = Sheets("Indicadores").Cells(68, Coluna_Lucro_1).Value
ValoresP_2 = Sheets("Indicadores").Cells(68, Coluna_Lucro_2).Value
ValoresP_3 = Sheets("Indicadores").Cells(68, Coluna_Lucro_3).Value
Sheets("TABELAS").Cells(32, 2).Value = ValoresP_3
Sheets("TABELAS").Cells(32, 3).Value = ValoresP_2
Sheets("TABELAS").Cells(32, 4).Value = ValoresP_1
'M�dia
ValoresM_1 = Sheets("Indicadores").Cells(69, Coluna_Lucro_1).Value
ValoresM_2 = Sheets("Indicadores").Cells(69, Coluna_Lucro_2).Value
ValoresM_3 = Sheets("Indicadores").Cells(69, Coluna_Lucro_3).Value
Sheets("TABELAS").Cells(33, 2).Value = ValoresM_3
Sheets("TABELAS").Cells(33, 3).Value = ValoresM_2
Sheets("TABELAS").Cells(33, 4).Value = ValoresM_1
'Grande
ValoresG_1 = Sheets("Indicadores").Cells(70, Coluna_Lucro_1).Value
ValoresG_2 = Sheets("Indicadores").Cells(70, Coluna_Lucro_2).Value
ValoresG_3 = Sheets("Indicadores").Cells(70, Coluna_Lucro_3).Value
Sheets("TABELAS").Cells(34, 2).Value = ValoresG_3
Sheets("TABELAS").Cells(34, 3).Value = ValoresG_2
Sheets("TABELAS").Cells(34, 4).Value = ValoresG_1

'Atribui os valores da coluna Pre�o
Coluna_Pre�o_1 = Sheets("Indicadores").Range("C175").End(xlToRight).Column
Coluna_Pre�o_2 = Coluna_Pre�o_1 - 1
Coluna_Pre�o_3 = Coluna_Pre�o_1 - 12

'Constru��o
ValoresC_1 = Sheets("Indicadores").Cells(175, Coluna_Pre�o_1).Value
ValoresC_2 = Sheets("Indicadores").Cells(175, Coluna_Pre�o_2).Value
ValoresC_3 = Sheets("Indicadores").Cells(175, Coluna_Pre�o_3).Value
Sheets("TABELAS").Cells(31, 5).Value = ValoresC_3
Sheets("TABELAS").Cells(31, 6).Value = ValoresC_2
Sheets("TABELAS").Cells(31, 7).Value = ValoresC_1
'Pequena
ValoresP_1 = Sheets("Indicadores").Cells(177, Coluna_Pre�o_1).Value
ValoresP_2 = Sheets("Indicadores").Cells(177, Coluna_Pre�o_2).Value
ValoresP_3 = Sheets("Indicadores").Cells(177, Coluna_Pre�o_3).Value
Sheets("TABELAS").Cells(32, 5).Value = ValoresP_3
Sheets("TABELAS").Cells(32, 6).Value = ValoresP_2
Sheets("TABELAS").Cells(32, 7).Value = ValoresP_1
'M�dia
ValoresM_1 = Sheets("Indicadores").Cells(178, Coluna_Pre�o_1).Value
ValoresM_2 = Sheets("Indicadores").Cells(178, Coluna_Pre�o_2).Value
ValoresM_3 = Sheets("Indicadores").Cells(178, Coluna_Pre�o_3).Value
Sheets("TABELAS").Cells(33, 5).Value = ValoresM_3
Sheets("TABELAS").Cells(33, 6).Value = ValoresM_2
Sheets("TABELAS").Cells(33, 7).Value = ValoresM_1
'Grande
ValoresG_1 = Sheets("Indicadores").Cells(179, Coluna_Pre�o_1).Value
ValoresG_2 = Sheets("Indicadores").Cells(179, Coluna_Pre�o_2).Value
ValoresG_3 = Sheets("Indicadores").Cells(179, Coluna_Pre�o_3).Value
Sheets("TABELAS").Cells(34, 5).Value = ValoresG_3
Sheets("TABELAS").Cells(34, 6).Value = ValoresG_2
Sheets("TABELAS").Cells(34, 7).Value = ValoresG_1

'Atribui os valores da coluna Situa��o
Coluna_Situa��o_1 = Sheets("Indicadores").Range("C79").End(xlToRight).Column
Coluna_Situa��o_2 = Coluna_Situa��o_1 - 1
Coluna_Situa��o_3 = Coluna_Situa��o_1 - 12

'Constru��o
ValoresC_1 = Sheets("Indicadores").Cells(79, Coluna_Situa��o_1).Value
ValoresC_2 = Sheets("Indicadores").Cells(79, Coluna_Situa��o_2).Value
ValoresC_3 = Sheets("Indicadores").Cells(79, Coluna_Situa��o_3).Value
Sheets("TABELAS").Cells(31, 8).Value = ValoresC_3
Sheets("TABELAS").Cells(31, 9).Value = ValoresC_2
Sheets("TABELAS").Cells(31, 10).Value = ValoresC_1
'Pequena
ValoresP_1 = Sheets("Indicadores").Cells(81, Coluna_Situa��o_1).Value
ValoresP_2 = Sheets("Indicadores").Cells(81, Coluna_Situa��o_2).Value
ValoresP_3 = Sheets("Indicadores").Cells(81, Coluna_Situa��o_3).Value
Sheets("TABELAS").Cells(32, 8).Value = ValoresP_3
Sheets("TABELAS").Cells(32, 9).Value = ValoresP_2
Sheets("TABELAS").Cells(32, 10).Value = ValoresP_1
'M�dia
ValoresM_1 = Sheets("Indicadores").Cells(82, Coluna_Situa��o_1).Value
ValoresM_2 = Sheets("Indicadores").Cells(82, Coluna_Situa��o_2).Value
ValoresM_3 = Sheets("Indicadores").Cells(82, Coluna_Situa��o_3).Value
Sheets("TABELAS").Cells(33, 8).Value = ValoresM_3
Sheets("TABELAS").Cells(33, 9).Value = ValoresM_2
Sheets("TABELAS").Cells(33, 10).Value = ValoresM_1
'Grande
ValoresG_1 = Sheets("Indicadores").Cells(83, Coluna_Situa��o_1).Value
ValoresG_2 = Sheets("Indicadores").Cells(83, Coluna_Situa��o_2).Value
ValoresG_3 = Sheets("Indicadores").Cells(83, Coluna_Situa��o_3).Value
Sheets("TABELAS").Cells(34, 8).Value = ValoresG_3
Sheets("TABELAS").Cells(34, 9).Value = ValoresG_2
Sheets("TABELAS").Cells(34, 10).Value = ValoresG_1

'Atribui os valores da coluna cr�dito
Coluna_Cr�dito_1 = Sheets("Indicadores").Range("C92").End(xlToRight).Column
Coluna_Cr�dito_2 = Coluna_Cr�dito_1 - 1
Coluna_Cr�dito_3 = Coluna_Cr�dito_1 - 12

'Constru��o
ValoresC_1 = Sheets("Indicadores").Cells(92, Coluna_Cr�dito_1).Value
ValoresC_2 = Sheets("Indicadores").Cells(92, Coluna_Cr�dito_2).Value
ValoresC_3 = Sheets("Indicadores").Cells(92, Coluna_Cr�dito_3).Value
Sheets("TABELAS").Cells(31, 11).Value = ValoresC_3
Sheets("TABELAS").Cells(31, 12).Value = ValoresC_2
Sheets("TABELAS").Cells(31, 13).Value = ValoresC_1
'Pequena
ValoresP_1 = Sheets("Indicadores").Cells(94, Coluna_Cr�dito_1).Value
ValoresP_2 = Sheets("Indicadores").Cells(94, Coluna_Cr�dito_2).Value
ValoresP_3 = Sheets("Indicadores").Cells(94, Coluna_Cr�dito_3).Value
Sheets("TABELAS").Cells(32, 11).Value = ValoresP_3
Sheets("TABELAS").Cells(32, 12).Value = ValoresP_2
Sheets("TABELAS").Cells(32, 13).Value = ValoresP_1
'M�dia
ValoresM_1 = Sheets("Indicadores").Cells(95, Coluna_Cr�dito_1).Value
ValoresM_2 = Sheets("Indicadores").Cells(95, Coluna_Cr�dito_2).Value
ValoresM_3 = Sheets("Indicadores").Cells(95, Coluna_Cr�dito_3).Value
Sheets("TABELAS").Cells(33, 11).Value = ValoresM_3
Sheets("TABELAS").Cells(33, 12).Value = ValoresM_2
Sheets("TABELAS").Cells(33, 13).Value = ValoresM_1
'Grande
ValoresG_1 = Sheets("Indicadores").Cells(96, Coluna_Cr�dito_1).Value
ValoresG_2 = Sheets("Indicadores").Cells(96, Coluna_Cr�dito_2).Value
ValoresG_3 = Sheets("Indicadores").Cells(96, Coluna_Cr�dito_3).Value
Sheets("TABELAS").Cells(34, 11).Value = ValoresG_3
Sheets("TABELAS").Cells(34, 12).Value = ValoresG_2
Sheets("TABELAS").Cells(34, 13).Value = ValoresG_1



'*******************************************************Princiapais Problemas******************************************************
Sheets("problemas_ponderado").Select

Coluna_Ultimo_Tri = Sheets("problemas_ponderado").Range("C10").End(xlToRight).Column
Coluna_Tri_Anterior = Coluna_Ultimo_Tri - 1
linha = 165

'Rank geral
Do Until linha = 183
posi��oG = Application.WorksheetFunction.Rank_Eq(Cells(linha, 4), Range("D165:D182").Cells, 0)
Cells(linha, 5).Value = posi��oG
linha = linha + 1
Loop

'Proc v Pequenas 1
linha = 165
Do Until linha = 185
Valor = Application.WorksheetFunction.VLookup(Cells(linha, 2), Range(Cells(32, 2), Cells(51, Coluna_Ultimo_Tri)), Coluna_Tri_Anterior - 1, 0)
Cells(linha, 6).Value = Valor
linha = linha + 1
Loop

'Proc v Pequenas 2
linha = 165
Do Until linha = 185
Valor = Application.WorksheetFunction.VLookup(Cells(linha, 2), Range(Cells(32, 2), Cells(51, Coluna_Ultimo_Tri)), Coluna_Ultimo_Tri - 1, 0)
Cells(linha, 7).Value = Valor
linha = linha + 1
Loop

'Rank Pequenas
linha = 165
Do Until linha = 183
posi��oP = Application.WorksheetFunction.Rank_Eq(Cells(linha, 7), Range("G165:G182").Cells, 0)
Cells(linha, 8).Value = posi��oP
linha = linha + 1
Loop

'Proc v medias 1
linha = 165
Do Until linha = 185
Valor = Application.WorksheetFunction.VLookup(Cells(linha, 2), Range(Cells(53, 2), Cells(72, Coluna_Ultimo_Tri)), Coluna_Tri_Anterior - 1, 0)
Cells(linha, 9).Value = Valor
linha = linha + 1
Loop

'Proc v medias 2
linha = 165
Do Until linha = 185
Valor = Application.WorksheetFunction.VLookup(Cells(linha, 2), Range(Cells(53, 2), Cells(72, Coluna_Ultimo_Tri)), Coluna_Ultimo_Tri - 1, 0)
Cells(linha, 10).Value = Valor
linha = linha + 1
Loop

'Rank medias
linha = 165
Do Until linha = 183
posi��oM = Application.WorksheetFunction.Rank_Eq(Cells(linha, 10), Range("J165:J182").Cells, 0)
Cells(linha, 11).Value = posi��oM
linha = linha + 1
Loop

'Proc v Grandes 1
linha = 165
Do Until linha = 185
Valor = Application.WorksheetFunction.VLookup(Cells(linha, 2), Range(Cells(74, 2), Cells(93, Coluna_Ultimo_Tri)), Coluna_Tri_Anterior - 1, 0)
Cells(linha, 12).Value = Valor
linha = linha + 1
Loop

'Proc v Grandes 2
linha = 165
Do Until linha = 185
Valor = Application.WorksheetFunction.VLookup(Cells(linha, 2), Range(Cells(74, 2), Cells(93, Coluna_Ultimo_Tri)), Coluna_Ultimo_Tri - 1, 0)
Cells(linha, 13).Value = Valor
linha = linha + 1
Loop

'Rank Grandes
linha = 165
Do Until linha = 183
posi��oGr = Application.WorksheetFunction.Rank_Eq(Cells(linha, 13), Range("M165:M182").Cells, 0)
Cells(linha, 14).Value = posi��oGr
linha = linha + 1
Loop

Sheets("problemas_ponderado").Select
Range("B165:N184").Copy
Sheets("TABELAS").Select
Range("V5").PasteSpecial xlPasteValues


Dim Sondagem As Workbook
Dim Modelo As Workbook
    
'   Capture current workbook
    Set Sondagem = ActiveWorkbook
    
'   Open new workbook
    Workbooks.Open ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Sondagem Ind�stria da Constru��o\Automa��o\Templates\Formata��o_das_ tabelas.xlsx")

'   Capture new workbook
    Set Modelo = ActiveWorkbook
    
Modelo.Activate
Sheets("Tabela principal").Select
Range("A1:AH44").Copy

' Go back to original workbook
Sondagem.Activate
Sheets("TABELAS").Select
Range("A1:AH44").PasteSpecial (xlPasteFormats)

Modelo.Activate
Range("A1:AH44").Select
    Application.CutCopyMode = False
Modelo.Close

End Sub





















