Attribute VB_Name = "C�digo_Definitivo"
Sub PraDarPlay()
Call Aba_Gr�fico
Call Tabelas
Call An�lise_Vermelho
Call An�lise_Azul
Call An�lise_Verde
Call Formata��o
End Sub


Sub Aba_Gr�fico()

Sheets.Add(Before:=Sheets("PRODU��O")).Name = "GR�FICO" 'Adiciona a aba gr�ficos

'Adiciona o tit�lo dos gr�ficos, que ser�o alocados de acordo com a posi�ao desses t�tulos
ActiveSheet.Range("A1").Value = "Evolu��o da Produ��o"
ActiveSheet.Range("A2").Value = "�ndice de difus�o (0 a 100 pontos)*"
ActiveSheet.Range("J1").Value = "Evolu��o do n�mero de empregados"
ActiveSheet.Range("J2").Value = "�ndice de difus�o (0 a 100 pontos)*"
ActiveSheet.Range("S1").Value = "Evolu��o do n�vel de estoques e do estoque efetivo em rela��o ao planejado"
ActiveSheet.Range("S2").Value = "�ndice de difus�o (0 a 100 pontos)*"
ActiveSheet.Range("AC1").Value = "Utiliza��o m�dia da capacidade instalada"
ActiveSheet.Range("AC2").Value = "Percentual (%)"
ActiveSheet.Range("AM1").Value = "Utiliza��o da capacidade instalada efetiva em rela��o ao usual"
ActiveSheet.Range("AM2").Value = "�ndice de difus�o (0 a 100 pontos)*"
ActiveSheet.Range("A27").Value = "�ndice de expectativa (Compra de Mat�rias-primas e N�mero de empregados)"
ActiveSheet.Range("A28").Value = "�ndice de difus�o (0 a 100 pontos)*"
ActiveSheet.Range("J27").Value = "�ndice de expectativa (Demanda e Expora��o)"
ActiveSheet.Range("J28").Value = "�ndice de difus�o (0 a 100 pontos)*"
ActiveSheet.Range("S27").Value = "Inten��o de investimento"
ActiveSheet.Range("AC27").Value = "Principais problemas enfrentados pela ind�stria no trimestre"
ActiveSheet.Range("AC28").Value = "Percentual (%)"
ActiveSheet.Range("A53").Value = "Facilidade de acesso ao cr�dito"
ActiveSheet.Range("A54").Value = "�ndice de difus�o (0 a 100 pontos)*"
ActiveSheet.Range("J53").Value = "Pre�o m�dio das mat�rias-primas"
ActiveSheet.Range("J54").Value = "�ndice de difus�o (0 a 100 pontos)*"
ActiveSheet.Range("S53").Value = "Satisfa��o com o lucro operacional e com a situa��o financeira"
ActiveSheet.Range("S54").Value = "�ndice de difus�o (0 a 100 pontos)*"

'********************************************************  Gr�fico Produ��o     ***************************************************************************

Dim U As Integer 'N�mero da �ltima Coluna
Dim P As Integer 'N�mero da primeira coluna
Dim cht As Object 'Gr�fico

U = Sheets("PRODU��O").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
P = U - 12 'Define o n�mero da primeira coluna

Sheets("PRODU��O").Select 'Seleciona a aba Produ��o
Sheets("PRODU��O").Range(Cells(55, P), Cells(55, U)).Value = "50" ' Insere a s�rie da linha divis�ria
Sheets("PRODU��O").Cells(55, 1).Value = "Linha divis�ria" 'Nomeia a celula que ser� usada como referencia para o t�tulo da s�rie
Sheets("PRODU��O").Cells(7, 2).Value = "Produ��o" 'Nomeia a celula que ser� usada como referencia para o t�tulo da s�rie


Set cht = Sheets("GR�FICO").Shapes.AddChart2 'Adiciona o gr�fico

Sheets("GR�FICO").Select 'Seleciona a aba gr�fico

cht.Select ' Seleciona o Gr�fico
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Sondagem Industrial\Automa��o\Templates\Emprego.crtx") ' Aplica o template do gr�fico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gr�fico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gr�fico
    ActiveChart.Parent.Top = Parent.Range("A3").Top 'reposiciona o grafico em rela��o ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("A3").Left ' reposiciona o gr�fico em rela��o � borda esquerda da planilha
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
    ActiveChart.FullSeriesCollection(1).Name = "=PRODU��O!" & Cells(7, 2).Address 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(1).Values = "=PRODU��O!" & Range(Cells(9, P), Cells(9, U)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(1).XValues = "=PRODU��O!" & Range(Cells(8, P), Cells(8, U)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
    ActiveChart.FullSeriesCollection(2).Name = "=PRODU��O!$A$55" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(2).Values = "=PRODU��O!" & Range(Cells(55, P), Cells(55, U)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(2).XValues = "=PRODU��O!" & Range(Cells(8, P), Cells(8, U)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.FullSeriesCollection(3).Delete ' Deleta os lixos importados do template

'********************************************************  Gr�fico Emprego    ********************************************************************

Dim A As Integer 'N�mero da �ltima Coluna
Dim B As Integer 'N�mero da primeira coluna
Dim GrafEmp As Object 'Gr�fico

A = Sheets("EMPREGADOS").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
B = A - 12 'Define o n�mero da primeira coluna

Sheets("EMPREGADOS").Select 'Seleciona a aba EMPREGADOS
Sheets("EMPREGADOS").Range(Cells(55, B), Cells(55, A)).Value = "50" ' Insere a s�rie da linha divis�ria
Sheets("EMPREGADOS").Cells(55, 1).Value = "Linha divis�ria" 'Nomeia a celula que ser� usada como referencia para o t�tulo da s�rie
Sheets("EMPREGADOS").Cells(7, 2).Value = "Emprego" 'Nomeia a celula que ser� usada como referencia para o t�tulo da s�rie


Set GrafEmp = Sheets("GR�FICO").Shapes.AddChart2 'Adiciona o gr�fico

Sheets("GR�FICO").Select 'Seleciona a aba gr�fico

GrafEmp.Select ' Seleciona o Gr�fico
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Sondagem Industrial\Automa��o\Templates\Emprego.crtx") ' Aplica o template do gr�fico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gr�fico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gr�fico
    ActiveChart.Parent.Top = Parent.Range("J3").Top 'reposiciona o grafico em rela��o ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("J3").Left ' reposiciona o gr�fico em rela��o � borda esquerda da planilha
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
    ActiveChart.FullSeriesCollection(1).Name = "=EMPREGADOS!" & Cells(7, 2).Address 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(1).Values = "=EMPREGADOS!" & Range(Cells(9, B), Cells(9, A)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(1).XValues = "=EMPREGADOS!" & Range(Cells(8, B), Cells(8, A)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
    ActiveChart.FullSeriesCollection(2).Name = "=EMPREGADOS!$A$55" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(2).Values = "=EMPREGADOS!" & Range(Cells(55, B), Cells(55, A)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(2).XValues = "=EMPREGADOS!" & Range(Cells(8, B), Cells(8, A)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.FullSeriesCollection(3).Delete ' Deleta os lixos importados do template

'********************************************************  Gr�fico Estoques   ********************************************************************

Dim F As Integer 'N�mero da �ltima Coluna
Dim G As Integer 'N�mero da primeira coluna
Dim GrafEst As Object 'Gr�fico

F = Sheets("ESTOQUES (evolu��o)").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
G = F - 12 'Define o n�mero da primeira coluna

Sheets("ESTOQUES (evolu��o)").Select 'Seleciona a aba ESTOQUES (evolu��o)
Sheets("ESTOQUES (evolu��o)").Range(Cells(55, G), Cells(55, F)).Value = "50" ' Insere a s�rie da linha divis�ria
Sheets("ESTOQUES (evolu��o)").Cells(55, 1).Value = "Linha divis�ria" 'Nomeia a celula que ser� usada como referencia para o t�tulo da s�rie
Sheets("ESTOQUES (evolu��o)").Cells(7, 2).Value = "Evolu��o" 'Nomeia a celula que ser� usada como referencia para o t�tulo da s�rie
Sheets("ESTOQUES (efetivo-planejado)").Cells(7, 2).Value = "Efetivo-planejado"

Set GrafEst = Sheets("GR�FICO").Shapes.AddChart2 'Adiciona o gr�fico

Sheets("GR�FICO").Select 'Seleciona a aba gr�fico

GrafEst.Select ' Seleciona o Gr�fico
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Sondagem Industrial\Automa��o\Templates\Estoque.crtx") ' Aplica o template do gr�fico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gr�fico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gr�fico
    ActiveChart.Parent.Top = Parent.Range("S3").Top 'reposiciona o grafico em rela��o ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("S3").Left ' reposiciona o gr�fico em rela��o � borda esquerda da planilha
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
    ActiveChart.FullSeriesCollection(1).Name = "='ESTOQUES (evolu��o)'!" & Cells(7, 2).Address 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(1).Values = "='ESTOQUES (evolu��o)'!" & Range(Cells(9, G), Cells(9, F)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(1).XValues = "='ESTOQUES (evolu��o)'!" & Range(Cells(8, G), Cells(8, F)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.NumberFormat = "#.##0,0"
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
    ActiveChart.FullSeriesCollection(2).Name = "='ESTOQUES (efetivo-planejado)'!" & Cells(7, 2).Address 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(2).Values = "='ESTOQUES (efetivo-planejado)'!" & Range(Cells(9, G + 12), Cells(9, F + 12)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(2).XValues = "='ESTOQUES (efetivo-planejado'!" & Range(Cells(8, G + 12), Cells(8, F + 12)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
    ActiveChart.FullSeriesCollection(3).Name = "='ESTOQUES (evolu��o)'!$A$55" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(3).Values = "='ESTOQUES (evolu��o)'!" & Range(Cells(55, G), Cells(55, F)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(3).XValues = "='ESTOQUES (evolu��o)'!" & Range(Cells(8, G), Cells(8, F)).Address 'determina os valores referentes ao eixo x da s�rie adicionada

   

'********************************************************  Gr�fico UCI    ********************************************************************

Dim C As Integer 'N�mero da �ltima Coluna
Dim GrafUCI As Object 'Gr�fico

C = Sheets("UCI (%)").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna

Sheets("UCI (%)").Select 'Seleciona a aba UCI (%)

'Cria a tabela com os anos nas linhas e os meses nas colunas
ActiveSheet.Range("K58").Value = "2011"
ActiveSheet.Range("K59").Value = "2012"
ActiveSheet.Range("K60").Value = "2013"
ActiveSheet.Range("K61").Value = "2014"
ActiveSheet.Range("K62").Value = "2015"
ActiveSheet.Range("K63").Value = "2016"
ActiveSheet.Range("K64").Value = "2017"
ActiveSheet.Range("K65").Value = "2018"
ActiveSheet.Range("K66").Value = "2019"
ActiveSheet.Range("K67").Value = "2020"
ActiveSheet.Range("K68").Value = "2021"
ActiveSheet.Range("K69").Value = "2022"

ActiveSheet.Range("L57").Value = "Jan"
ActiveSheet.Range("M57").Value = "Fev"
ActiveSheet.Range("N57").Value = "Mar"
ActiveSheet.Range("O57").Value = "Abr"
ActiveSheet.Range("P57").Value = "Mai"
ActiveSheet.Range("Q57").Value = "Jun"
ActiveSheet.Range("R57").Value = "Jul"
ActiveSheet.Range("S57").Value = "Ago"
ActiveSheet.Range("T57").Value = "Set"
ActiveSheet.Range("U57").Value = "Out"
ActiveSheet.Range("V57").Value = "Nov"
ActiveSheet.Range("W57").Value = "Dez"

'Copia os dados nos de acordo com a tabela criada no c�digo anterior
ActiveSheet.Range("B9:M9").Copy ActiveSheet.Range("L58")
ActiveSheet.Range("N9:Y9").Copy ActiveSheet.Range("L59")
ActiveSheet.Range("Z9:AK9").Copy ActiveSheet.Range("L60")
ActiveSheet.Range("AL9:AW9").Copy ActiveSheet.Range("L61")
ActiveSheet.Range("AX9:BI9").Copy ActiveSheet.Range("L62")
ActiveSheet.Range("BJ9:BU9").Copy ActiveSheet.Range("L63")
ActiveSheet.Range("BV9:CG9").Copy ActiveSheet.Range("L64")
ActiveSheet.Range("CH9:CS9").Copy ActiveSheet.Range("L65")
ActiveSheet.Range("CT9:DE9").Copy ActiveSheet.Range("L66")
ActiveSheet.Range("DF9:DQ9").Copy ActiveSheet.Range("L67")
ActiveSheet.Range("DR9:EC9").Copy ActiveSheet.Range("L68")
ActiveSheet.Range(Cells(9, 134), Cells(9, C)).Copy ActiveSheet.Range("L69")

'Calcula e nomeia a m�dia dos meses com os valores de 2011 a 2019
ActiveSheet.Range("K56").Value = "M�dia 2011 - 2019"
ActiveSheet.Range("L56").Value = Application.Average(Range("L58:L66"))
ActiveSheet.Range("M56").Value = Application.Average(Range("M58:M66"))
ActiveSheet.Range("N56").Value = Application.Average(Range("N58:N66"))
ActiveSheet.Range("O56").Value = Application.Average(Range("O58:O66"))
ActiveSheet.Range("P56").Value = Application.Average(Range("P58:P66"))
ActiveSheet.Range("Q56").Value = Application.Average(Range("Q58:Q66"))
ActiveSheet.Range("R56").Value = Application.Average(Range("R58:R66"))
ActiveSheet.Range("S56").Value = Application.Average(Range("S58:S66"))
ActiveSheet.Range("T56").Value = Application.Average(Range("T58:T66"))
ActiveSheet.Range("U56").Value = Application.Average(Range("U58:U66"))
ActiveSheet.Range("V56").Value = Application.Average(Range("V58:V66"))
ActiveSheet.Range("W56").Value = Application.Average(Range("W58:W66"))
ActiveSheet.Range("L56:W56").Select
Selection.NumberFormat = "0.0"

Set GrafUCI = Sheets("GR�FICO").Shapes.AddChart2 'Adiciona o gr�fico

Sheets("GR�FICO").Select 'Seleciona a aba gr�fico

GrafUCI.Select ' Seleciona o Gr�fico
ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Sondagem Industrial\Automa��o\Templates\UCI.crtx") ' Aplica o template do gr�fico
ActiveChart.Parent.Height = 300 'ajusta a altura do gr�fico
ActiveChart.Parent.Width = 425 ' ajusta a largura do gr�fico
ActiveChart.Parent.Top = Parent.Range("AC3").Top 'reposiciona o grafico em rela��o ao topo da planilha
ActiveChart.Parent.Left = Parent.Range("AC3").Left ' reposiciona o gr�fico em rela��o � borda esquerda da planilha
ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
ActiveChart.FullSeriesCollection(1).Name = "='UCI (%)'!$K$56" 'Determina o nome da s�rie
ActiveChart.FullSeriesCollection(1).Values = "='UCI (%)'!$L$56:$W$56" 'determina os valores da s�rie
ActiveChart.FullSeriesCollection(1).XValues = "='UCI (%)'!$L$57:$W$57" 'determina os valores referentes ao eixo x da s�rie adicionada
ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
ActiveChart.FullSeriesCollection(2).Name = "='UCI (%)'!$K$67" 'Determina o nome da s�rie
ActiveChart.FullSeriesCollection(2).Values = "='UCI (%)'!$L$67:$W$67" 'determina os valores da s�rie
ActiveChart.FullSeriesCollection(2).XValues = "='UCI (%)'!$L$57:$W$57" 'determina os valores referentes ao eixo x da s�rie adicionada
ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
ActiveChart.FullSeriesCollection(3).Name = "='UCI (%)'!$K$68" 'Determina o nome da s�rie
ActiveChart.FullSeriesCollection(3).Values = "='UCI (%)'!$L$68:$W$68" 'determina os valores da s�rie
ActiveChart.FullSeriesCollection(3).XValues = "='UCI (%)'!$L$57:$W$57" 'determina os valores referentes ao eixo x da s�rie adicionada
ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
ActiveChart.FullSeriesCollection(4).Name = "='UCI (%)'!$K$69" 'Determina o nome da s�rie
ActiveChart.FullSeriesCollection(4).Values = "='UCI (%)'!$L$69:$W$69" 'determina os valores da s�rie
ActiveChart.FullSeriesCollection(4).XValues = "='UCI (%)'!$B$57:$M$57" 'determina os valores referentes ao eixo x da s�rie adicionada
ActiveChart.SetElement (msoElementLegendBottom)



'                   * Os c�digos abaixo cont�m um exemplo do que deve ser feito para o ano de 2022.
    



'Sempre que um ano novo se iniciar � necess�rio ajustar este c�digo, a come�ar pela consolida��o do ano passado e a adi��o
'do novo ano na parte do 'c�digo descrita por "'Copia os dados nos de acordo com a tabela criada no c�digo anterior" a partir da seguencia abaixo:

'   ActiveSheet.Range("B9:M9").Copy ActiveSheet.Range("B58:M58")
'   ActiveSheet.Range("N9:Y9").Copy ActiveSheet.Range("B59:M59")
'   ActiveSheet.Range("Z9:AK9").Copy ActiveSheet.Range("B60:M60")
'   ActiveSheet.Range("AL9:AW9").Copy ActiveSheet.Range("B61:M61")
'   ActiveSheet.Range("AX9:BI9").Copy ActiveSheet.Range("B62:M62")
'   ActiveSheet.Range("BJ9:BU9").Copy ActiveSheet.Range("B63:M63")
'   ActiveSheet.Range("BV9:CG9").Copy ActiveSheet.Range("B64:M64")
'   ActiveSheet.Range("CH9:CS9").Copy ActiveSheet.Range("B65:M65")
'   ActiveSheet.Range("CT9:DE9").Copy ActiveSheet.Range("B66:M66")
'   ActiveSheet.Range("DF9:DQ9").Copy ActiveSheet.Range("B67:M67")
'   ActiveSheet.Range("DR9:EC9").Copy ActiveSheet.Range("B68:M68")
'   ActiveSheet.Range(Cells(9, 134), Cells(9, C)).Copy ActiveSheet.Range("B69")

'� necess�rio tamb�m  a adi��o da linha com o ano novo com na tabela de anos e meses com o c�digo que segue na sess�o
'descrita por "'Cria a tabela com os anos nas linhas e os meses nas colunas"

'                      ActiveSheet.Range("A69").Value = "2022"


'H� duas maneiras de prossegui a partir deste momento 1 adicionando o ano passado � m�dia para manter a estrutura de 3 linhas
'ou 2 adicionar uma nova s�rie com o novo ano.

'1) Para adicionar o ano passado � m�dia basta fazer os seguintes ajustes na se��o "'Calcula e nomeia a m�dia dos meses com os valores de 2011 a 2019" :


'   ActiveSheet.Range("A56").Value = "M�dia 2011 - 2020"
'   ActiveSheet.Range("B56").Value = Application.Average(Range("B58:B67"))
'   ActiveSheet.Range("C56").Value = Application.Average(Range("C58:C67"))
'   ActiveSheet.Range("D56").Value = Application.Average(Range("D58:D67"))
'   ActiveSheet.Range("E56").Value = Application.Average(Range("E58:E67"))
'   ActiveSheet.Range("F56").Value = Application.Average(Range("F58:F67"))
'   ActiveSheet.Range("G56").Value = Application.Average(Range("G58:G67"))
'   ActiveSheet.Range("H56").Value = Application.Average(Range("H58:H67"))
'   ActiveSheet.Range("I56").Value = Application.Average(Range("I58:I67"))
'   ActiveSheet.Range("J56").Value = Application.Average(Range("J58:J67"))
'   ActiveSheet.Range("K56").Value = Application.Average(Range("K58:K67"))
'   ActiveSheet.Range("L56").Value = Application.Average(Range("L58:L67"))
'   ActiveSheet.Range("M56").Value = Application.Average(Range("M58:M67"))
'   ActiveSheet.Range("B56:M56").Select
'   Selection.NumberFormat = "0.0"

'1.1) Para manter a estrutura de 3 linhas basta ajustar a se��o ""'Seleciona o Gr�fico" da fprma que segue abaixo:

'   ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
'   ActiveChart.FullSeriesCollection(1).Name = "='UCI (%)'!$A$56" 'Determina o nome da s�rie
'   ActiveChart.FullSeriesCollection(1).Values = "='UCI (%)'!$B$56:$M$56" 'determina os valores da s�rie
'   ActiveChart.FullSeriesCollection(1).XValues = "='UCI (%)'!$B$57:$M$57" 'determina os valores referentes ao eixo x da s�rie adicionada
'   ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
'   ActiveChart.FullSeriesCollection(2).Name = "='UCI (%)'!$A$67" 'Determina o nome da s�rie
'   ActiveChart.FullSeriesCollection(2).Values = "='UCI (%)'!$B$68:$M$68" 'determina os valores da s�rie
'   ActiveChart.FullSeriesCollection(2).XValues = "='UCI (%)'!$B$57:$M$57" 'determina os valores referentes ao eixo x da s�rie adicionada
'   ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
'   ActiveChart.FullSeriesCollection(3).Name = "='UCI (%)'!$A$68" 'Determina o nome da s�rie
'   ActiveChart.FullSeriesCollection(3).Values = "='UCI (%)'!$B$69:$M$69" 'determina os valores da s�rie
'   ActiveChart.FullSeriesCollection(3).XValues = "='UCI (%)'!$B$57:$M$57" 'determina os valores referentes ao eixo x da s�rie adicionada
'   ActiveChart.FullSeriesCollection(6).Delete ' Deleta os lixos importados do template
'   ActiveChart.FullSeriesCollection(4).Delete ' Deleta os lixos importados do template
'   ActiveChart.FullSeriesCollection(4).Delete ' Deleta os lixos importados do template
    
    
'2) para adicionarar uma nova s�rie o novo ano a basta fazer os seguintes ajustes na se��o "'Seleciona o Gr�fico":

  
'   ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
'   ActiveChart.FullSeriesCollection(1).Name = "='UCI (%)'!$A$56" 'Determina o nome da s�rie
'   ActiveChart.FullSeriesCollection(1).Values = "='UCI (%)'!$B$56:$M$56" 'determina os valores da s�rie
'   ActiveChart.FullSeriesCollection(1).XValues = "='UCI (%)'!$B$57:$M$57" 'determina os valores referentes ao eixo x da s�rie adicionada
'   ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
'   ActiveChart.FullSeriesCollection(2).Name = "='UCI (%)'!$A$67" 'Determina o nome da s�rie
'   ActiveChart.FullSeriesCollection(2).Values = "='UCI (%)'!$B$67:$M$67" 'determina os valores da s�rie
'   ActiveChart.FullSeriesCollection(2).XValues = "='UCI (%)'!$B$57:$M$57" 'determina os valores referentes ao eixo x da s�rie adicionada
'   ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
'   ActiveChart.FullSeriesCollection(3).Name = "='UCI (%)'!$A$68" 'Determina o nome da s�rie
'   ActiveChart.FullSeriesCollection(3).Values = "='UCI (%)'!$B$68:$M$68" 'determina os valores da s�rie
'   ActiveChart.FullSeriesCollection(3).XValues = "='UCI (%)'!$B$57:$M$57" 'determina os valores referentes ao eixo x da s�rie adicionada
'   ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
'   ActiveChart.FullSeriesCollection(4).Name = "='UCI (%)'!$A$69" 'Determina o nome da s�rie
'   ActiveChart.FullSeriesCollection(4).Values = "='UCI (%)'!$B$69:$M$69" 'determina os valores da s�rie
'   ActiveChart.FullSeriesCollection(4).XValues = "='UCI (%)'!$B$57:$M$57" 'determina os valores referentes ao eixo x da s�rie adicionada
'   ActiveChart.FullSeriesCollection(5).Delete ' Deleta os lixos importados do template'
'   ActiveChart.FullSeriesCollection(5).Delete ' Deleta os lixos importados do template

'********************************************************  Gr�fico UCI Efetivo Usual    ********************************************************************

Dim D As Integer 'N�mero da �ltima Coluna
Dim E As Integer
Dim GrafUCIEU As Object 'Gr�fico

D = Sheets("UCI (efetiva-usual)").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
E = D - 132

Sheets("UCI (efetiva-usual)").Select 'Seleciona a aba UCI (efetiva-usual)
Sheets("UCI (efetiva-usual)").Range(Cells(55, E), Cells(55, D)).Value = "50" ' Insere a s�rie da linha divis�ria
Sheets("UCI (efetiva-usual)").Cells(55, 1).Value = "Linha divis�ria" 'Nomeia a celula que ser� usada como referencia para o t�tulo da s�rie
Sheets("UCI (efetiva-usual)").Cells(7, 2).Value = "UCI (efetiva-usual)" 'Nomeia a celula que ser� usada como referencia para o t�tulo da s�rie



Set GrafUCIEU = Sheets("GR�FICO").Shapes.AddChart2 'Adiciona o gr�fico

Sheets("GR�FICO").Select 'Seleciona a aba gr�fico

GrafUCIEU.Select ' Seleciona o Gr�fico
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Sondagem Industrial\Automa��o\Templates\UCI(Efetiva Usual).crtx") ' Aplica o template do gr�fico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gr�fico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gr�fico
    ActiveChart.Parent.Top = Parent.Range("AM3").Top 'reposiciona o grafico em rela��o ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("AM3").Left ' reposiciona o gr�fico em rela��o � borda esquerda da planilha
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
    ActiveChart.FullSeriesCollection(1).Name = "='UCI (efetiva-usual)'!" & Cells(7, 2).Address 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(1).Values = "='UCI (efetiva-usual)'!" & Range(Cells(9, E), Cells(9, D)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(1).XValues = "='UCI (efetiva-usual)'!" & Range(Cells(8, E), Cells(8, D)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
    ActiveChart.FullSeriesCollection(2).Name = "='UCI (efetiva-usual)'!$A$55" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(2).Values = "='UCI (efetiva-usual)'!" & Range(Cells(55, E), Cells(55, D)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(2).XValues = "='UCI (efetiva-usual)'!" & Range(Cells(8, E), Cells(8, D)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
   

'********************************************************  Gr�fico Expectativa Compras e Empregados    ********************************************************************

Dim J As Integer
Dim K As Integer
Dim GrafComEmp As Object

J = Sheets("EXPECTATIVA - COMPRAS").Range("A9").End(xlToRight).Column
K = J - 120

Sheets("EXPECTATIVA - COMPRAS").Select 'Seleciona a aba EXPECTATIVA - COMPRAS
Sheets("EXPECTATIVA - COMPRAS").Range(Cells(55, K), Cells(55, J)).Value = "50" ' Insere a s�rie da linha divis�ria
Sheets("EXPECTATIVA - COMPRAS").Cells(55, 1).Value = "Linha divis�ria" 'Nomeia a celula que ser� usada como referencia para o t�tulo da s�rie
Sheets("EXPECTATIVA - COMPRAS").Cells(7, 2).Value = "Expectativa de compras de mat�rias-primas" 'Nomeia a celula que ser� usada como referencia para o t�tulo da s�rie
Sheets("EXPECTATIVA - EMPREGADOS").Cells(7, 2).Value = "Expectativa de n�mero de empregados"


Set GrafGrafComEmp = Sheets("GR�FICO").Shapes.AddChart2 'Adiciona o gr�fico

Sheets("GR�FICO").Select 'Seleciona a aba gr�fico

GrafGrafComEmp.Select ' Seleciona o Gr�fico
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Sondagem Industrial\Automa��o\Templates\Expectativa - Demanda e Exporta��o.crtx") ' Aplica o template do gr�fico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gr�fico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gr�fico
    ActiveChart.Parent.Top = Parent.Range("A29").Top 'reposiciona o grafico em rela��o ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("A29").Left ' reposiciona o gr�fico em rela��o � borda esquerda da planilha
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
    ActiveChart.FullSeriesCollection(1).Name = "='EXPECTATIVA - COMPRAS'!" & Cells(7, 2).Address 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(1).Values = "='EXPECTATIVA - COMPRAS'!" & Range(Cells(9, K), Cells(9, J)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(1).XValues = "='EXPECTATIVA - COMPRAS'!" & Range(Cells(8, K), Cells(8, J)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
    ActiveChart.FullSeriesCollection(2).Name = "='EXPECTATIVA - EMPREGADOS'!" & Cells(7, 2).Address 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(2).Values = "='EXPECTATIVA - EMPREGADOS'!" & Range(Cells(9, K - 8), Cells(9, J - 8)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(2).XValues = "='EXPECTATIVA - EMPREGADOS'!" & Range(Cells(8, K - 8), Cells(8, J - 8)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
    ActiveChart.FullSeriesCollection(3).Name = "='EXPECTATIVA - COMPRAS'!$A$55" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(3).Values = "='EXPECTATIVA - COMPRAS'!" & Range(Cells(55, K), Cells(55, J)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(3).XValues = "='EXPECTATIVA - COMPRAS'!" & Range(Cells(8, K), Cells(8, J)).Address 'determina os valores referentes ao eixo x da s�rie adicionada


'********************************************************  Gr�fico Expectativa Demanda e Exporta��o    ********************************************************************

Dim H As Integer
Dim I As Integer
Dim GrafDemExt As Object

H = Sheets("EXPECTATIVAS - DEMANDA").Range("A9").End(xlToRight).Column
I = H - 132

Sheets("EXPECTATIVAS - DEMANDA").Select 'Seleciona a aba EXPECTATIVAS - DEMANDA
Sheets("EXPECTATIVAS - DEMANDA").Range(Cells(55, I), Cells(55, H)).Value = "50" ' Insere a s�rie da linha divis�ria
Sheets("EXPECTATIVAS - DEMANDA").Cells(55, 1).Value = "Linha divis�ria" 'Nomeia a celula que ser� usada como referencia para o t�tulo da s�rie
Sheets("EXPECTATIVAS - DEMANDA").Cells(7, 2).Value = "Expectativa de demanda" 'Nomeia a celula que ser� usada como referencia para o t�tulo da s�rie
Sheets("EXPECTATIVA - EXPORTA��O").Cells(7, 2).Value = "Expectativa de exporta��o"

Set GrafDemExt = Sheets("GR�FICO").Shapes.AddChart2 'Adiciona o gr�fico

Sheets("GR�FICO").Select 'Seleciona a aba gr�fico

GrafDemExt.Select ' Seleciona o Gr�fico
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Sondagem Industrial\Automa��o\Templates\Expectativa - Compra e empregados.crtx") ' Aplica o template do gr�fico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gr�fico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gr�fico
    ActiveChart.Parent.Top = Parent.Range("J29").Top 'reposiciona o grafico em rela��o ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("J29").Left ' reposiciona o gr�fico em rela��o � borda esquerda da planilha
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
    ActiveChart.FullSeriesCollection(1).Name = "='EXPECTATIVAS - DEMANDA'!" & Cells(7, 2).Address 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(1).Values = "='EXPECTATIVAS - DEMANDA'!" & Range(Cells(9, I), Cells(9, H)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(1).XValues = "='EXPECTATIVAS - DEMANDA'!" & Range(Cells(8, I), Cells(8, H)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
    ActiveChart.FullSeriesCollection(2).Name = "='EXPECTATIVA - EXPORTA��O'!" & Cells(7, 2).Address 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(2).Values = "='EXPECTATIVA - EXPORTA��O'!" & Range(Cells(9, I - 12), Cells(9, H - 12)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(2).XValues = "='EXPECTATIVA - EXPORTA��O'!" & Range(Cells(8, I - 12), Cells(8, H - 12)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
    ActiveChart.FullSeriesCollection(3).Name = "='EXPECTATIVAS - DEMANDA'!$A$55" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(3).Values = "='EXPECTATIVAS - DEMANDA'!" & Range(Cells(55, I), Cells(55, H)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(3).XValues = "='EXPECTATIVAS - DEMANDA'!" & Range(Cells(8, I), Cells(8, H)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
   

'********************************************************  Gr�fico Inten��o de investimento    ********************************************************************

Dim L As Integer 'N�mero da �ltima Coluna
Dim M As Integer 'N�mero da primeira Coluna
Dim GrafIntInv As Object 'Gr�fico

L = Sheets("EXPECTATIVA - INVESTIMENTO").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
M = L - 84 'Define o n�mero da primeira coluna

Sheets("EXPECTATIVA - INVESTIMENTO").Select 'Seleciona a aba EXPECTATIVA - INVESTIMENTO
Sheets("EXPECTATIVA - INVESTIMENTO").Cells(7, 2).Value = "Inten��o de investimento" 'Nomeia a celula que ser� usada como referencia para o t�tulo da s�rie

Set GrafIntInv = Sheets("GR�FICO").Shapes.AddChart2 'Adiciona o gr�fico

Sheets("GR�FICO").Select 'Seleciona a aba gr�fico

GrafIntInv.Select ' Seleciona o Gr�fico
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Sondagem Industrial\Automa��o\Templates\Investimento.crtx") ' Aplica o template do gr�fico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gr�fico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gr�fico
    ActiveChart.Parent.Top = Parent.Range("S29").Top 'reposiciona o gr�fico em rela��o ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("S29").Left ' reposiciona o gr�fico em rela��o � borda esquerda da planilha
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
    ActiveChart.FullSeriesCollection(1).Name = "='EXPECTATIVA - INVESTIMENTO'!" & Cells(7, 2).Address 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(1).Values = "='EXPECTATIVA - INVESTIMENTO'!" & Range(Cells(9, M), Cells(9, L)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(1).XValues = "='EXPECTATIVA - INVESTIMENTO'!" & Range(Cells(8, M), Cells(8, L)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    

'********************************************************  Gr�fico Credito    ********************************************************************

Dim Q As Integer 'N�mero da �ltima Coluna
Dim R As Integer 'N�mero da primeira coluna
Dim GrafCredito As Object 'Gr�fico

Q = Sheets("SITUACAO FINANCEIRA CREDITO").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
R = Q - 36 'Define o n�mero da primeira coluna

Sheets("SITUACAO FINANCEIRA CREDITO").Select 'Seleciona a aba SITUACAO FINANCEIRA CREDITO
Sheets("SITUACAO FINANCEIRA CREDITO").Range(Cells(55, R), Cells(55, Q)).Value = "50" ' Insere a s�rie da linha divis�ria
Sheets("SITUACAO FINANCEIRA CREDITO").Cells(55, 1).Value = "Linha divis�ria" 'Nomeia a celula que ser� usada como referencia para o t�tulo da s�rie
Sheets("SITUACAO FINANCEIRA CREDITO").Cells(7, 2).Value = "Facilidade de acesso ao cr�dito" 'Nomeia a celula que ser� usada como referencia para o t�tulo da s�rie

Set GrafCredito = Sheets("GR�FICO").Shapes.AddChart2 'Adiciona o gr�fico
Sheets("GR�FICO").Select 'Seleciona a aba gr�fico
GrafCredito.Select ' Seleciona o Gr�fico
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Sondagem Industrial\Automa��o\Templates\Cr�dito.crtx") ' Aplica o template do gr�fico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gr�fico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gr�fico
    ActiveChart.Parent.Top = Parent.Range("A55").Top 'reposiciona o grafico em rela��o ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("A55").Left ' reposiciona o gr�fico em rela��o � borda esquerda da planilha
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
    ActiveChart.FullSeriesCollection(1).Name = "='SITUACAO FINANCEIRA CREDITO'!" & Cells(7, 2).Address   'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(1).Values = "='SITUACAO FINANCEIRA CREDITO'!" & Range(Cells(9, R), Cells(9, Q)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(1).XValues = "='SITUACAO FINANCEIRA CREDITO'!" & Range(Cells(8, R), Cells(8, Q)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
    ActiveChart.FullSeriesCollection(2).Name = "='SITUACAO FINANCEIRA CREDITO'!$A$55" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(2).Values = "='SITUACAO FINANCEIRA CREDITO'!" & Range(Cells(55, R), Cells(55, Q)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(2).XValues = "='SITUACAO FINANCEIRA CREDITO'!" & Range(Cells(8, R), Cells(8, Q)).Address 'determina os valores referentes ao eixo x da s�rie adicionada

'********************************************************  Gr�fico Pre�o M�dio    ********************************************************************

Dim S As Integer 'N�mero da �ltima Coluna
Dim T As Integer 'N�mero da primeira coluna
Dim GrafPM As Object 'Gr�fico

S = Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
T = S - 36 'Define o n�mero da primeira coluna

Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Select 'Seleciona a aba SITUACAO FINANCEIRA PRE�O MEDIO
Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Range(Cells(55, T), Cells(55, S)).Value = "50" ' Insere a s�rie da linha divis�ria
Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(55, 1).Value = "Linha divis�ria" 'Nomeia a celula que ser� usada como referencia para o t�tulo da s�rie
Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(7, 2).Value = "Pre�o m�dio das mat�rias-primas" 'Nomeia a celula que ser� usada como referencia para o t�tulo da s�rie

Set GrafPM = Sheets("GR�FICO").Shapes.AddChart2 'Adiciona o gr�fico
Sheets("GR�FICO").Select 'Seleciona a aba gr�fico
GrafPM.Select ' Seleciona o Gr�fico
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Sondagem Industrial\Automa��o\Templates\Pre�o.crtx") ' Aplica o template do gr�fico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gr�fico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gr�fico
    ActiveChart.Parent.Top = Parent.Range("J55").Top 'reposiciona o grafico em rela��o ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("J55").Left ' reposiciona o gr�fico em rela��o � borda esquerda da planilha
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
    ActiveChart.FullSeriesCollection(1).Name = "='SITUACAO FINANCEIRA PRE�O MEDIO'!" & Cells(7, 2).Address  'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(1).Values = "='SITUACAO FINANCEIRA PRE�O MEDIO'!" & Range(Cells(9, T), Cells(9, S)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(1).XValues = "='SITUACAO FINANCEIRA PRE�O MEDIO'!" & Range(Cells(8, T), Cells(8, S)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
    ActiveChart.FullSeriesCollection(2).Name = "='SITUACAO FINANCEIRA PRE�O MEDIO'!$A$55" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(2).Values = "='SITUACAO FINANCEIRA PRE�O MEDIO'!" & Range(Cells(55, T), Cells(55, S)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(2).XValues = "='SITUACAO FINANCEIRA PRE�O MEDIO'!" & Range(Cells(8, T), Cells(8, S)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.Axes(xlValue).MinimumScale = 40
    ActiveChart.Axes(xlValue).MaximumScale = 85


'********************************************************  Gr�fico Lucro    ********************************************************************

Dim N As Integer 'N�mero da �ltima Coluna
Dim O As Integer 'N�mero da primeira coluna
Dim GrafSFL As Object 'Gr�fico

N = Sheets("SITUACAO FINANCEIRA LUCRO").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
O = N - 36 'Define o n�mero da primeira coluna

Sheets("SITUACAO FINANCEIRA LUCRO").Select 'Seleciona a aba Produ��o
Sheets("SITUACAO FINANCEIRA LUCRO").Range(Cells(55, O), Cells(55, N)).Value = "50" ' Insere a s�rie da linha divis�ria
Sheets("SITUACAO FINANCEIRA LUCRO").Cells(55, 1).Value = "Linha divis�ria" 'Nomeia a celula que ser� usada como referencia para o t�tulo da s�rie
Sheets("SITUACAO FINANCEIRA LUCRO").Cells(7, 2).Value = "Lucro Operacional" 'Nomeia a celula que ser� usada como referencia para o t�tulo da s�rie
Sheets("SITUACAO FINANCEIRA").Cells(7, 2).Value = "Situa��o financeira" 'Nomeia a celula que ser� usada como referencia para o t�tulo da s�rie

Set GrafSFL = Sheets("GR�FICO").Shapes.AddChart2 'Adiciona o gr�fico
Sheets("GR�FICO").Select 'Seleciona a aba gr�fico
GrafSFL.Select ' Seleciona o Gr�fico
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Sondagem Industrial\Automa��o\Templates\Lucro e situa��o financeira.crtx") ' Aplica o template do gr�fico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gr�fico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gr�fico
    ActiveChart.Parent.Top = Parent.Range("S55").Top 'reposiciona o grafico em rela��o ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("S55").Left ' reposiciona o gr�fico em rela��o � borda esquerda da planilha
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
    ActiveChart.FullSeriesCollection(1).Name = "='SITUACAO FINANCEIRA LUCRO'!" & Cells(7, 2).Address 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(1).Values = "='SITUACAO FINANCEIRA LUCRO'!" & Range(Cells(9, O), Cells(9, N)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(1).XValues = "='SITUACAO FINANCEIRA LUCRO'!" & Range(Cells(8, O), Cells(8, N)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
    ActiveChart.FullSeriesCollection(2).Name = "='SITUACAO FINANCEIRA LUCRO'!$A$55" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(2).Values = "='SITUACAO FINANCEIRA LUCRO'!" & Range(Cells(55, O), Cells(55, N)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(2).XValues = "='SITUACAO FINANCEIRA LUCRO'!" & Range(Cells(8, O), Cells(8, N)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
    ActiveChart.FullSeriesCollection(3).Name = "='SITUACAO FINANCEIRA'!" & Cells(7, 2).Address 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(3).Values = "='SITUACAO FINANCEIRA'!" & Range(Cells(9, O), Cells(9, N)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(3).XValues = "='SITUACAO FINANCEIRA'!" & Range(Cells(8, O), Cells(8, N)).Address 'determina os valores referentes ao eixo x da s�rie adicionada


'********************************************************  Gr�fico Principais problemas    ********************************************************************
 
Sheets("Principais_Problemas").Select ' Seleciona a aba Principais_Problemas
ActiveSheet.Range("B12:B28").Copy ActiveSheet.Range("B110") ' Copia e cola o nome das categorias menos outros e nehum.

Dim V As Integer 'Numero do trimestre mais recente
Dim X As Integer 'N�mero do trimestre anterior
Dim GrafProblemas As Object ' Gr�fico
 
V = Sheets("Principais_Problemas").Range("B13").End(xlToRight).Column 'Define o n�mero da �ltima coluna
X = V - 1 'Define o n�mero da primeira coluna

ActiveSheet.Range(Cells(13, X), Cells(28, V)).Copy ActiveSheet.Range("C111") 'Copia os valores para formar a tabela
ActiveSheet.Range(Cells(10, X), Cells(10, V)).Copy ActiveSheet.Range("C110") ' copia o nome dos trimestres

'Filtra os valores na tabela de forma decrescente de acordo com o trimestre mais recente
ActiveSheet.Range("B110:D110").Select
Selection.AutoFilter
ActiveSheet.AutoFilter.Sort.SortFields. _
        Add2 Key:=Range("D110"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
        
ActiveSheet.Range("B29:B30").Copy ActiveSheet.Range("B127") 'Copia o nome das categorias outros e nenhum na tabela
ActiveSheet.Range(Cells(29, X), Cells(30, V)).Copy ActiveSheet.Range("C127") ' copia os valores das categorias outros e nenhum na tabela

Set GrafProblemas = Sheets("GR�FICO").Shapes.AddChart2 'Adiciona o gr�fico
Sheets("GR�FICO").Select 'Seleciona a aba gr�fico
GrafProblemas.Select ' Seleciona o Gr�fico
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Sondagem Industrial\Automa��o\Templates\Principais Problemas.crtx") ' Aplica o template do gr�fico
    ActiveChart.Parent.Height = 630 'ajusta a altura do gr�fico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gr�fico
    ActiveChart.Parent.Top = Parent.Range("AC29").Top 'reposiciona o grafico em rela��o ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("AC29").Left ' reposiciona o gr�fico em rela��o � borda esquerda da planilha
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
    ActiveChart.FullSeriesCollection(1).Name = "='PRINCIPAIS_PROBLEMAS'!$D$110" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(1).Values = "='PRINCIPAIS_PROBLEMAS'!$D$111:$D$128" 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(1).XValues = "='PRINCIPAIS_PROBLEMAS'!$B$111:$B$128" 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova s�rie ao gr�fico
    ActiveChart.FullSeriesCollection(2).Name = "='PRINCIPAIS_PROBLEMAS'!$C$110" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(2).Values = "='PRINCIPAIS_PROBLEMAS'!$C$111:$C$128" 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(2).XValues = "='PRINCIPAIS_PROBLEMAS'!$B$111:$B$128" 'determina os valores referentes ao eixo x da s�rie adicionada

End Sub

Sub Tabelas()

'Adiciona a aba tabela
Sheets.Add(Before:=Sheets("PRODU��O")).Name = "TABELAS"

'Nomeia os titulos das colunas e mescla as celulas
Sheets("TABELAS").Cells(1, 1).Value = "Desempenho da ind�stria"

Sheets("TABELAS").Range(Cells(2, 2), Cells(3, 4)).Merge
Sheets("TABELAS").Cells(2, 2).Value = "EVOLU��O DA PRODU��O"

Sheets("TABELAS").Range(Cells(2, 5), Cells(3, 7)).Merge
Sheets("TABELAS").Cells(2, 5).Value = "EVOLU��O DO NO DE EMPREGADOS"

Sheets("TABELAS").Range(Cells(2, 8), Cells(3, 10)).Merge
Sheets("TABELAS").Cells(2, 8).Value = "UCI (%)"

Sheets("TABELAS").Range(Cells(2, 11), Cells(3, 13)).Merge
Sheets("TABELAS").Cells(2, 11).Value = " UCI EFETIVA-USUAL"

Sheets("TABELAS").Range(Cells(2, 14), Cells(3, 16)).Merge
Sheets("TABELAS").Cells(2, 14).Value = "EVOLU��O DOS ESTOQUES"

Sheets("TABELAS").Range(Cells(2, 17), Cells(3, 19)).Merge
Sheets("TABELAS").Cells(2, 17).Value = "ESTOQUE EFETIVO-PLANEJADO"

Sheets("TABELAS").Range(Cells(6, 1), Cells(7, 19)).Merge
Sheets("TABELAS").Cells(6, 1).Value = "POR SEGMENTO INDUSTRIAL"

Sheets("TABELAS").Range(Cells(10, 1), Cells(11, 19)).Merge
Sheets("TABELAS").Cells(10, 1).Value = "POR PORTE"

'Centraliza e alinha os titulos
Sheets("TABELAS").Range("A2:Q2").VerticalAlignment = xlCenter
Sheets("TABELAS").Range("A2:Q2").HorizontalAlignment = xlCenter

Sheets("TABELAS").Range("A6").VerticalAlignment = xlCenter
Sheets("TABELAS").Range("A6").HorizontalAlignment = xlCenter

Sheets("TABELAS").Range("A10").VerticalAlignment = xlCenter
Sheets("TABELAS").Range("A10").HorizontalAlignment = xlCenter

'Nomeia as linhas
Sheets("TABELAS").Range("A5").Value = "Ind�stria Geral"
Sheets("TABELAS").Range("A8").Value = "Ind�stria extrativa"
Sheets("TABELAS").Range("A9").Value = "Ind�stria de transforma��o"
Sheets("TABELAS").Range("A12").Value = "Pequena"
Sheets("TABELAS").Range("A13").Value = "M�dia"
Sheets("TABELAS").Range("A14").Value = "Grande"

'Define as variavies que ser�o usadas para preencher as celulas
Coluna_Produ��o_1 = Sheets("PRODU��O").Range("B8").End(xlToRight).Column
Coluna_Produ��o_2 = Coluna_Produ��o_1 - 1
Coluna_Produ��o_3 = Coluna_Produ��o_1 - 12

'Define, atribui e copia e cola as datas
Datas_1 = Sheets("PRODU��O").Cells(8, Coluna_Produ��o_1).Value
Datas_2 = Sheets("PRODU��O").Cells(8, Coluna_Produ��o_2).Value
Datas_3 = Sheets("PRODU��O").Cells(8, Coluna_Produ��o_3).Value

Sheets("TABELAS").Cells(4, 2).Value = Datas_3
Sheets("TABELAS").Cells(4, 3).Value = Datas_2
Sheets("TABELAS").Cells(4, 4).Value = Datas_1

Sheets("TABELAS").Range("B4:D4").Copy (Sheets("TABELAS").Range("E4:G4"))
Sheets("TABELAS").Range("B4:D4").Copy (Sheets("TABELAS").Range("H4:J4"))
Sheets("TABELAS").Range("B4:D4").Copy (Sheets("TABELAS").Range("K4:M4"))
Sheets("TABELAS").Range("B4:D4").Copy (Sheets("TABELAS").Range("N4:P4"))
Sheets("TABELAS").Range("B4:D4").Copy (Sheets("TABELAS").Range("Q4:S4"))

'Atribui os valores da coluna Evolu��o da produ��o
'Ind�stria Geral
ValoresIGP_1 = Sheets("PRODU��O").Cells(9, Coluna_Produ��o_1).Value
ValoresIGP_2 = Sheets("PRODU��O").Cells(9, Coluna_Produ��o_2).Value
ValoresIGP_3 = Sheets("PRODU��O").Cells(9, Coluna_Produ��o_3).Value
Sheets("TABELAS").Cells(5, 2).Value = ValoresIGP_3
Sheets("TABELAS").Cells(5, 3).Value = ValoresIGP_2
Sheets("TABELAS").Cells(5, 4).Value = ValoresIGP_1
'Ind�stria Extrativa
ValoresIEP_1 = Sheets("PRODU��O").Cells(21, Coluna_Produ��o_1).Value
ValoresIEP_2 = Sheets("PRODU��O").Cells(21, Coluna_Produ��o_2).Value
ValoresIEP_3 = Sheets("PRODU��O").Cells(21, Coluna_Produ��o_3).Value
Sheets("TABELAS").Cells(8, 2).Value = ValoresIEP_3
Sheets("TABELAS").Cells(8, 3).Value = ValoresIEP_2
Sheets("TABELAS").Cells(8, 4).Value = ValoresIEP_1
'Ind�stria da Transforma��o
ValoresITP_1 = Sheets("PRODU��O").Cells(26, Coluna_Produ��o_1).Value
ValoresITP_2 = Sheets("PRODU��O").Cells(26, Coluna_Produ��o_2).Value
ValoresITP_3 = Sheets("PRODU��O").Cells(26, Coluna_Produ��o_3).Value
Sheets("TABELAS").Cells(9, 2).Value = ValoresITP_3
Sheets("TABELAS").Cells(9, 3).Value = ValoresITP_2
Sheets("TABELAS").Cells(9, 4).Value = ValoresITP_1
'Pequena
ValoresPP_1 = Sheets("PRODU��O").Cells(17, Coluna_Produ��o_1).Value
ValoresPP_2 = Sheets("PRODU��O").Cells(17, Coluna_Produ��o_2).Value
ValoresPP_3 = Sheets("PRODU��O").Cells(17, Coluna_Produ��o_3).Value
Sheets("TABELAS").Cells(12, 2).Value = ValoresPP_3
Sheets("TABELAS").Cells(12, 3).Value = ValoresPP_2
Sheets("TABELAS").Cells(12, 4).Value = ValoresPP_1
'M�dia
ValoresMP_1 = Sheets("PRODU��O").Cells(18, Coluna_Produ��o_1).Value
ValoresMP_2 = Sheets("PRODU��O").Cells(18, Coluna_Produ��o_2).Value
ValoresMP_3 = Sheets("PRODU��O").Cells(18, Coluna_Produ��o_3).Value
Sheets("TABELAS").Cells(13, 2).Value = ValoresMP_3
Sheets("TABELAS").Cells(13, 3).Value = ValoresMP_2
Sheets("TABELAS").Cells(13, 4).Value = ValoresMP_1
'Grande
ValoresGP_1 = Sheets("PRODU��O").Cells(19, Coluna_Produ��o_1).Value
ValoresGP_2 = Sheets("PRODU��O").Cells(19, Coluna_Produ��o_2).Value
ValoresGP_3 = Sheets("PRODU��O").Cells(19, Coluna_Produ��o_3).Value
Sheets("TABELAS").Cells(14, 2).Value = ValoresGP_3
Sheets("TABELAS").Cells(14, 3).Value = ValoresGP_2
Sheets("TABELAS").Cells(14, 4).Value = ValoresGP_1

'Atribui os valores da coluna Evolu��o do N� de Empregoados
Coluna_Emprego_1 = Sheets("EMPREGADOS").Range("B8").End(xlToRight).Column
Coluna_Emprego_2 = Coluna_Emprego_1 - 1
Coluna_Emprego_3 = Coluna_Emprego_1 - 12

'Ind�stria Geral
ValoresIGE_1 = Sheets("EMPREGADOS").Cells(9, Coluna_Emprego_1).Value
ValoresIGE_2 = Sheets("EMPREGADOS").Cells(9, Coluna_Emprego_2).Value
ValoresIGE_3 = Sheets("EMPREGADOS").Cells(9, Coluna_Emprego_3).Value
Sheets("TABELAS").Cells(5, 5).Value = ValoresIGE_3
Sheets("TABELAS").Cells(5, 6).Value = ValoresIGE_2
Sheets("TABELAS").Cells(5, 7).Value = ValoresIGE_1
'Ind�stria Extrativa
ValoresIEE_1 = Sheets("EMPREGADOS").Cells(21, Coluna_Emprego_1).Value
ValoresIEE_2 = Sheets("EMPREGADOS").Cells(21, Coluna_Emprego_2).Value
ValoresIEE_3 = Sheets("EMPREGADOS").Cells(21, Coluna_Emprego_3).Value
Sheets("TABELAS").Cells(8, 5).Value = ValoresIEE_3
Sheets("TABELAS").Cells(8, 6).Value = ValoresIEE_2
Sheets("TABELAS").Cells(8, 7).Value = ValoresIEE_1
'Ind�stria Transforma��o
ValoresITE_1 = Sheets("EMPREGADOS").Cells(26, Coluna_Emprego_1).Value
ValoresITE_2 = Sheets("EMPREGADOS").Cells(26, Coluna_Emprego_2).Value
ValoresITE_3 = Sheets("EMPREGADOS").Cells(26, Coluna_Emprego_3).Value
Sheets("TABELAS").Cells(9, 5).Value = ValoresITE_3
Sheets("TABELAS").Cells(9, 6).Value = ValoresITE_2
Sheets("TABELAS").Cells(9, 7).Value = ValoresITE_1
'Pequena
ValoresPE_1 = Sheets("EMPREGADOS").Cells(17, Coluna_Emprego_1).Value
ValoresPE_2 = Sheets("EMPREGADOS").Cells(17, Coluna_Emprego_2).Value
ValoresPE_3 = Sheets("EMPREGADOS").Cells(17, Coluna_Emprego_3).Value
Sheets("TABELAS").Cells(12, 5).Value = ValoresPE_3
Sheets("TABELAS").Cells(12, 6).Value = ValoresPE_2
Sheets("TABELAS").Cells(12, 7).Value = ValoresPE_1
'M�dia
ValoresME_1 = Sheets("EMPREGADOS").Cells(18, Coluna_Emprego_1).Value
ValoresME_2 = Sheets("EMPREGADOS").Cells(18, Coluna_Emprego_2).Value
ValoresME_3 = Sheets("EMPREGADOS").Cells(18, Coluna_Emprego_3).Value
Sheets("TABELAS").Cells(13, 5).Value = ValoresME_3
Sheets("TABELAS").Cells(13, 6).Value = ValoresME_2
Sheets("TABELAS").Cells(13, 7).Value = ValoresME_1
'Grande
ValoresGE_1 = Sheets("EMPREGADOS").Cells(19, Coluna_Emprego_1).Value
ValoresGE_2 = Sheets("EMPREGADOS").Cells(19, Coluna_Emprego_2).Value
ValoresGE_3 = Sheets("EMPREGADOS").Cells(19, Coluna_Emprego_3).Value
Sheets("TABELAS").Cells(14, 5).Value = ValoresGE_3
Sheets("TABELAS").Cells(14, 6).Value = ValoresGE_2
Sheets("TABELAS").Cells(14, 7).Value = ValoresGE_1

'Atribui os valores da coluna UCI(%)
Coluna_UCI_1 = Sheets("UCI (%)").Range("B8").End(xlToRight).Column
Coluna_UCI_2 = Coluna_UCI_1 - 1
Coluna_UCI_3 = Coluna_UCI_1 - 12

'Ind�stria Geral
ValoresIG_UCI_1 = Sheets("UCI (%)").Cells(9, Coluna_UCI_1).Value
ValoresIG_UCI_2 = Sheets("UCI (%)").Cells(9, Coluna_UCI_2).Value
ValoresIG_UCI_3 = Sheets("UCI (%)").Cells(9, Coluna_UCI_3).Value
Sheets("TABELAS").Cells(5, 8).Value = ValoresIG_UCI_3
Sheets("TABELAS").Cells(5, 9).Value = ValoresIG_UCI_2
Sheets("TABELAS").Cells(5, 10).Value = ValoresIG_UCI_1
'Ind�stria extrativa
ValoresIE_UCI_1 = Sheets("UCI (%)").Cells(21, Coluna_UCI_1).Value
ValoresIE_UCI_2 = Sheets("UCI (%)").Cells(21, Coluna_UCI_2).Value
ValoresIE_UCI_3 = Sheets("UCI (%)").Cells(21, Coluna_UCI_3).Value
Sheets("TABELAS").Cells(8, 8).Value = ValoresIE_UCI_3
Sheets("TABELAS").Cells(8, 9).Value = ValoresIE_UCI_2
Sheets("TABELAS").Cells(8, 10).Value = ValoresIE_UCI_1
'Ind�stria Transforma��o
ValoresIT_UCI_1 = Sheets("UCI (%)").Cells(26, Coluna_UCI_1).Value
ValoresIT_UCI_2 = Sheets("UCI (%)").Cells(26, Coluna_UCI_2).Value
ValoresIT_UCI_3 = Sheets("UCI (%)").Cells(26, Coluna_UCI_3).Value
Sheets("TABELAS").Cells(9, 8).Value = ValoresIT_UCI_3
Sheets("TABELAS").Cells(9, 9).Value = ValoresIT_UCI_2
Sheets("TABELAS").Cells(9, 10).Value = ValoresIT_UCI_1
'Pequena
ValoresP_UCI_1 = Sheets("UCI (%)").Cells(17, Coluna_UCI_1).Value
ValoresP_UCI_2 = Sheets("UCI (%)").Cells(17, Coluna_UCI_2).Value
ValoresP_UCI_3 = Sheets("UCI (%)").Cells(17, Coluna_UCI_3).Value
Sheets("TABELAS").Cells(12, 8).Value = ValoresP_UCI_3
Sheets("TABELAS").Cells(12, 9).Value = ValoresP_UCI_2
Sheets("TABELAS").Cells(12, 10).Value = ValoresP_UCI_1
'M�dia
ValoresM_UCI_1 = Sheets("UCI (%)").Cells(18, Coluna_UCI_1).Value
ValoresM_UCI_2 = Sheets("UCI (%)").Cells(18, Coluna_UCI_2).Value
ValoresM_UCI_3 = Sheets("UCI (%)").Cells(18, Coluna_UCI_3).Value
Sheets("TABELAS").Cells(13, 8).Value = ValoresM_UCI_3
Sheets("TABELAS").Cells(13, 9).Value = ValoresM_UCI_2
Sheets("TABELAS").Cells(13, 10).Value = ValoresM_UCI_1
'Grande
ValoresG_UCI_1 = Sheets("UCI (%)").Cells(19, Coluna_UCI_1).Value
ValoresG_UCI_2 = Sheets("UCI (%)").Cells(19, Coluna_UCI_2).Value
ValoresG_UCI_3 = Sheets("UCI (%)").Cells(19, Coluna_UCI_3).Value
Sheets("TABELAS").Cells(14, 8).Value = ValoresG_UCI_3
Sheets("TABELAS").Cells(14, 9).Value = ValoresG_UCI_2
Sheets("TABELAS").Cells(14, 10).Value = ValoresG_UCI_1

'Atribui os valores da coluna UCI efetiva usual
Coluna_UCI_EU_1 = Sheets("UCI (efetiva-usual)").Range("B8").End(xlToRight).Column
Coluna_UCI_EU_2 = Coluna_UCI_EU_1 - 1
Coluna_UCI_EU_3 = Coluna_UCI_EU_1 - 12

'Ind�stria Geral
ValoresIG_UCI_EU_1 = Sheets("UCI (efetiva-usual)").Cells(9, Coluna_UCI_EU_1).Value
ValoresIG_UCI_EU_2 = Sheets("UCI (efetiva-usual)").Cells(9, Coluna_UCI_EU_2).Value
ValoresIG_UCI_EU_3 = Sheets("UCI (efetiva-usual)").Cells(9, Coluna_UCI_EU_3).Value
Sheets("TABELAS").Cells(5, 11).Value = ValoresIG_UCI_EU_3
Sheets("TABELAS").Cells(5, 12).Value = ValoresIG_UCI_EU_2
Sheets("TABELAS").Cells(5, 13).Value = ValoresIG_UCI_EU_1
'Ind�stria Extrativa
ValoresIE_UCI_EU_1 = Sheets("UCI (efetiva-usual)").Cells(21, Coluna_UCI_EU_1).Value
ValoresIE_UCI_EU_2 = Sheets("UCI (efetiva-usual)").Cells(21, Coluna_UCI_EU_2).Value
ValoresIE_UCI_EU_3 = Sheets("UCI (efetiva-usual)").Cells(21, Coluna_UCI_EU_3).Value
Sheets("TABELAS").Cells(8, 11).Value = ValoresIE_UCI_EU_3
Sheets("TABELAS").Cells(8, 12).Value = ValoresIE_UCI_EU_2
Sheets("TABELAS").Cells(8, 13).Value = ValoresIE_UCI_EU_1
'Ind�stria transforma��o
ValoresIT_UCI_EU_1 = Sheets("UCI (efetiva-usual)").Cells(26, Coluna_UCI_EU_1).Value
ValoresIT_UCI_EU_2 = Sheets("UCI (efetiva-usual)").Cells(26, Coluna_UCI_EU_2).Value
ValoresIT_UCI_EU_3 = Sheets("UCI (efetiva-usual)").Cells(26, Coluna_UCI_EU_3).Value
Sheets("TABELAS").Cells(9, 11).Value = ValoresIT_UCI_EU_3
Sheets("TABELAS").Cells(9, 12).Value = ValoresIT_UCI_EU_2
Sheets("TABELAS").Cells(9, 13).Value = ValoresIT_UCI_EU_1
'Pequena
ValoresP_UCI_EU_1 = Sheets("UCI (efetiva-usual)").Cells(17, Coluna_UCI_EU_1).Value
ValoresP_UCI_EU_2 = Sheets("UCI (efetiva-usual)").Cells(17, Coluna_UCI_EU_2).Value
ValoresP_UCI_EU_3 = Sheets("UCI (efetiva-usual)").Cells(17, Coluna_UCI_EU_3).Value
Sheets("TABELAS").Cells(12, 11).Value = ValoresP_UCI_EU_3
Sheets("TABELAS").Cells(12, 12).Value = ValoresP_UCI_EU_2
Sheets("TABELAS").Cells(12, 13).Value = ValoresP_UCI_EU_1
'M�dia
ValoresM_UCI_EU_1 = Sheets("UCI (efetiva-usual)").Cells(18, Coluna_UCI_EU_1).Value
ValoresM_UCI_EU_2 = Sheets("UCI (efetiva-usual)").Cells(18, Coluna_UCI_EU_2).Value
ValoresM_UCI_EU_3 = Sheets("UCI (efetiva-usual)").Cells(18, Coluna_UCI_EU_3).Value
Sheets("TABELAS").Cells(13, 11).Value = ValoresM_UCI_EU_3
Sheets("TABELAS").Cells(13, 12).Value = ValoresM_UCI_EU_2
Sheets("TABELAS").Cells(13, 13).Value = ValoresM_UCI_EU_1
'Grande
ValoresG_UCI_EU_1 = Sheets("UCI (efetiva-usual)").Cells(19, Coluna_UCI_EU_1).Value
ValoresG_UCI_EU_2 = Sheets("UCI (efetiva-usual)").Cells(19, Coluna_UCI_EU_2).Value
ValoresG_UCI_EU_3 = Sheets("UCI (efetiva-usual)").Cells(19, Coluna_UCI_EU_3).Value
Sheets("TABELAS").Cells(14, 11).Value = ValoresG_UCI_EU_3
Sheets("TABELAS").Cells(14, 12).Value = ValoresG_UCI_EU_2
Sheets("TABELAS").Cells(14, 13).Value = ValoresG_UCI_EU_1

'Atribui os valores da coluna Evolu��o dos estoques
Coluna_Estoques_1 = Sheets("ESTOQUES (evolu��o)").Range("B8").End(xlToRight).Column
Coluna_Estoques_2 = Coluna_Estoques_1 - 1
Coluna_Estoques_3 = Coluna_Estoques_1 - 12

'Ind�stria Geral
ValoresIG_Estoques_1 = Sheets("ESTOQUES (evolu��o)").Cells(9, Coluna_Estoques_1).Value
ValoresIG_Estoques_2 = Sheets("ESTOQUES (evolu��o)").Cells(9, Coluna_Estoques_2).Value
ValoresIG_Estoques_3 = Sheets("ESTOQUES (evolu��o)").Cells(9, Coluna_Estoques_3).Value
Sheets("TABELAS").Cells(5, 14).Value = ValoresIG_Estoques_3
Sheets("TABELAS").Cells(5, 15).Value = ValoresIG_Estoques_2
Sheets("TABELAS").Cells(5, 16).Value = ValoresIG_Estoques_1
'Ind�stria Extrativa
ValoresIE_Estoques_1 = Sheets("ESTOQUES (evolu��o)").Cells(21, Coluna_Estoques_1).Value
ValoresIE_Estoques_2 = Sheets("ESTOQUES (evolu��o)").Cells(21, Coluna_Estoques_2).Value
ValoresIE_Estoques_3 = Sheets("ESTOQUES (evolu��o)").Cells(21, Coluna_Estoques_3).Value
Sheets("TABELAS").Cells(8, 14).Value = ValoresIE_Estoques_3
Sheets("TABELAS").Cells(8, 15).Value = ValoresIE_Estoques_2
Sheets("TABELAS").Cells(8, 16).Value = ValoresIE_Estoques_1
'Ind�stria Transforma��o
ValoresIT_Estoques_1 = Sheets("ESTOQUES (evolu��o)").Cells(26, Coluna_Estoques_1).Value
ValoresIT_Estoques_2 = Sheets("ESTOQUES (evolu��o)").Cells(26, Coluna_Estoques_2).Value
ValoresIT_Estoques_3 = Sheets("ESTOQUES (evolu��o)").Cells(26, Coluna_Estoques_3).Value
Sheets("TABELAS").Cells(9, 14).Value = ValoresIT_Estoques_3
Sheets("TABELAS").Cells(9, 15).Value = ValoresIT_Estoques_2
Sheets("TABELAS").Cells(9, 16).Value = ValoresIT_Estoques_1
'Pequena
ValoresP_Estoques_1 = Sheets("ESTOQUES (evolu��o)").Cells(17, Coluna_Estoques_1).Value
ValoresP_Estoques_2 = Sheets("ESTOQUES (evolu��o)").Cells(17, Coluna_Estoques_2).Value
ValoresP_Estoques_3 = Sheets("ESTOQUES (evolu��o)").Cells(17, Coluna_Estoques_3).Value
Sheets("TABELAS").Cells(12, 14).Value = ValoresP_Estoques_3
Sheets("TABELAS").Cells(12, 15).Value = ValoresP_Estoques_2
Sheets("TABELAS").Cells(12, 16).Value = ValoresP_Estoques_1
'M�dia
ValoresM_Estoques_1 = Sheets("ESTOQUES (evolu��o)").Cells(18, Coluna_Estoques_1).Value
ValoresM_Estoques_2 = Sheets("ESTOQUES (evolu��o)").Cells(18, Coluna_Estoques_2).Value
ValoresM_Estoques_3 = Sheets("ESTOQUES (evolu��o)").Cells(18, Coluna_Estoques_3).Value
Sheets("TABELAS").Cells(13, 14).Value = ValoresM_Estoques_3
Sheets("TABELAS").Cells(13, 15).Value = ValoresM_Estoques_2
Sheets("TABELAS").Cells(13, 16).Value = ValoresM_Estoques_1
'Grande
ValoresG_Estoques_1 = Sheets("ESTOQUES (evolu��o)").Cells(19, Coluna_Estoques_1).Value
ValoresG_Estoques_2 = Sheets("ESTOQUES (evolu��o)").Cells(19, Coluna_Estoques_2).Value
ValoresG_Estoques_3 = Sheets("ESTOQUES (evolu��o)").Cells(19, Coluna_Estoques_3).Value
Sheets("TABELAS").Cells(14, 14).Value = ValoresG_Estoques_3
Sheets("TABELAS").Cells(14, 15).Value = ValoresG_Estoques_2
Sheets("TABELAS").Cells(14, 16).Value = ValoresG_Estoques_1

'Atribui os valores da coluna Estoque efetivo-planejado
Coluna_Estoques_EP_1 = Sheets("ESTOQUES (efetivo-planejado)").Range("B8").End(xlToRight).Column
Coluna_Estoques_EP_2 = Coluna_Estoques_EP_1 - 1
Coluna_Estoques_EP_3 = Coluna_Estoques_EP_1 - 12

'Ind�stria Geral
ValoresIG_Estoques_EP_1 = Sheets("ESTOQUES (efetivo-planejado)").Cells(9, Coluna_Estoques_EP_1).Value
ValoresIG_Estoques_EP_2 = Sheets("ESTOQUES (efetivo-planejado)").Cells(9, Coluna_Estoques_EP_2).Value
ValoresIG_Estoques_EP_3 = Sheets("ESTOQUES (efetivo-planejado)").Cells(9, Coluna_Estoques_EP_3).Value
Sheets("TABELAS").Cells(5, 17).Value = ValoresIG_Estoques_EP_3
Sheets("TABELAS").Cells(5, 18).Value = ValoresIG_Estoques_EP_2
Sheets("TABELAS").Cells(5, 19).Value = ValoresIG_Estoques_EP_1
'Ind�stria extrativa
ValoresIE_Estoques_EP_1 = Sheets("ESTOQUES (efetivo-planejado)").Cells(21, Coluna_Estoques_EP_1).Value
ValoresIE_Estoques_EP_2 = Sheets("ESTOQUES (efetivo-planejado)").Cells(21, Coluna_Estoques_EP_2).Value
ValoresIE_Estoques_EP_3 = Sheets("ESTOQUES (efetivo-planejado)").Cells(21, Coluna_Estoques_EP_3).Value
Sheets("TABELAS").Cells(8, 17).Value = ValoresIE_Estoques_EP_3
Sheets("TABELAS").Cells(8, 18).Value = ValoresIE_Estoques_EP_2
Sheets("TABELAS").Cells(8, 19).Value = ValoresIE_Estoques_EP_1
'Ind�stria Transforma��o
ValoresIT_Estoques_EP_1 = Sheets("ESTOQUES (efetivo-planejado)").Cells(26, Coluna_Estoques_EP_1).Value
ValoresIT_Estoques_EP_2 = Sheets("ESTOQUES (efetivo-planejado)").Cells(26, Coluna_Estoques_EP_2).Value
ValoresIT_Estoques_EP_3 = Sheets("ESTOQUES (efetivo-planejado)").Cells(26, Coluna_Estoques_EP_3).Value
Sheets("TABELAS").Cells(9, 17).Value = ValoresIT_Estoques_EP_3
Sheets("TABELAS").Cells(9, 18).Value = ValoresIT_Estoques_EP_2
Sheets("TABELAS").Cells(9, 19).Value = ValoresIT_Estoques_EP_1
'Pequena
ValoresP_Estoques_EP_1 = Sheets("ESTOQUES (efetivo-planejado)").Cells(17, Coluna_Estoques_EP_1).Value
ValoresP_Estoques_EP_2 = Sheets("ESTOQUES (efetivo-planejado)").Cells(17, Coluna_Estoques_EP_2).Value
ValoresP_Estoques_EP_3 = Sheets("ESTOQUES (efetivo-planejado)").Cells(17, Coluna_Estoques_EP_3).Value
Sheets("TABELAS").Cells(12, 17).Value = ValoresP_Estoques_EP_3
Sheets("TABELAS").Cells(12, 18).Value = ValoresP_Estoques_EP_2
Sheets("TABELAS").Cells(12, 19).Value = ValoresP_Estoques_EP_1
'M�dia
ValoresM_Estoques_EP_1 = Sheets("ESTOQUES (efetivo-planejado)").Cells(18, Coluna_Estoques_EP_1).Value
ValoresM_Estoques_EP_2 = Sheets("ESTOQUES (efetivo-planejado)").Cells(18, Coluna_Estoques_EP_2).Value
ValoresM_Estoques_EP_3 = Sheets("ESTOQUES (efetivo-planejado)").Cells(18, Coluna_Estoques_EP_3).Value
Sheets("TABELAS").Cells(13, 17).Value = ValoresM_Estoques_EP_3
Sheets("TABELAS").Cells(13, 18).Value = ValoresM_Estoques_EP_2
Sheets("TABELAS").Cells(13, 19).Value = ValoresM_Estoques_EP_1
'Grande
ValoresG_Estoques_EP_1 = Sheets("ESTOQUES (efetivo-planejado)").Cells(19, Coluna_Estoques_EP_1).Value
ValoresG_Estoques_EP_2 = Sheets("ESTOQUES (efetivo-planejado)").Cells(19, Coluna_Estoques_EP_2).Value
ValoresG_Estoques_EP_3 = Sheets("ESTOQUES (efetivo-planejado)").Cells(19, Coluna_Estoques_EP_3).Value
Sheets("TABELAS").Cells(14, 17).Value = ValoresG_Estoques_EP_3
Sheets("TABELAS").Cells(14, 18).Value = ValoresG_Estoques_EP_2
Sheets("TABELAS").Cells(14, 19).Value = ValoresG_Estoques_EP_1

'*************************************************** C�digo da parte de Expectativas **********************************************************

'Nomeia os titulos das colunas e mescla as celulas
Sheets("TABELAS").Cells(16, 1).Value = "Expectativas da ind�stria"

Sheets("TABELAS").Range(Cells(17, 2), Cells(18, 4)).Merge
Sheets("TABELAS").Cells(17, 2).Value = "DEMANDA"

Sheets("TABELAS").Range(Cells(17, 5), Cells(18, 7)).Merge
Sheets("TABELAS").Cells(17, 5).Value = "QUANTIDADE EXPORTADA"

Sheets("TABELAS").Range(Cells(17, 8), Cells(18, 10)).Merge
Sheets("TABELAS").Cells(17, 8).Value = "COMPRAS DE MAT�RIA-PRIMA"

Sheets("TABELAS").Range(Cells(17, 11), Cells(18, 13)).Merge
Sheets("TABELAS").Cells(17, 11).Value = "N� DE EMPREGADOS"

Sheets("TABELAS").Range(Cells(17, 14), Cells(18, 16)).Merge
Sheets("TABELAS").Cells(17, 14).Value = "INTEN��O DE INVESTIMENTO"

Sheets("TABELAS").Range(Cells(21, 1), Cells(22, 16)).Merge
Sheets("TABELAS").Cells(21, 1).Value = "POR SEGMENTO INDUSTRIAL"

Sheets("TABELAS").Range(Cells(25, 1), Cells(26, 16)).Merge
Sheets("TABELAS").Cells(25, 1).Value = "POR PORTE"

'Centraliza e alinha os titulos
Sheets("TABELAS").Range("A17:N17").VerticalAlignment = xlCenter
Sheets("TABELAS").Range("A17:N17").HorizontalAlignment = xlCenter

Sheets("TABELAS").Range("A21").VerticalAlignment = xlCenter
Sheets("TABELAS").Range("A21").HorizontalAlignment = xlCenter

Sheets("TABELAS").Range("A25").VerticalAlignment = xlCenter
Sheets("TABELAS").Range("A25").HorizontalAlignment = xlCenter

'Nomeia as linhas
Sheets("TABELAS").Range("A20").Value = "Ind�stria Geral"
Sheets("TABELAS").Range("A23").Value = "Ind�stria extrativa"
Sheets("TABELAS").Range("A24").Value = "Ind�stria de transforma��o"
Sheets("TABELAS").Range("A27").Value = "Pequena"
Sheets("TABELAS").Range("A28").Value = "M�dia"
Sheets("TABELAS").Range("A29").Value = "Grande"

'Define as variavies que ser�o usadas para preencher as celulas
Coluna_Demanda_1 = Sheets("EXPECTATIVAS - DEMANDA").Range("B8").End(xlToRight).Column
Coluna_Demanda_2 = Coluna_Demanda_1 - 1
Coluna_Demanda_3 = Coluna_Demanda_1 - 12

'Define, atribui e copia e cola as datas
Data_1 = Sheets("EXPECTATIVAS - DEMANDA").Cells(8, Coluna_Demanda_1).Value
Data_2 = Sheets("EXPECTATIVAS - DEMANDA").Cells(8, Coluna_Demanda_2).Value
Data_3 = Sheets("EXPECTATIVAS - DEMANDA").Cells(8, Coluna_Demanda_3).Value

Sheets("TABELAS").Cells(19, 2).Value = Data_3
Sheets("TABELAS").Cells(19, 3).Value = Data_2
Sheets("TABELAS").Cells(19, 4).Value = Data_1

Sheets("TABELAS").Range("B19:D19").Copy (Sheets("TABELAS").Range("E19:G19"))
Sheets("TABELAS").Range("B19:D19").Copy (Sheets("TABELAS").Range("H19:J19"))
Sheets("TABELAS").Range("B19:D19").Copy (Sheets("TABELAS").Range("K19:M19"))
Sheets("TABELAS").Range("B19:D19").Copy (Sheets("TABELAS").Range("N19:P19"))

'Atribui os valores da coluna Demanda
'Ind�stria Geral
ValoresIG_Demanda_1 = Sheets("EXPECTATIVAS - DEMANDA").Cells(9, Coluna_Demanda_1).Value
ValoresIG_Demanda_2 = Sheets("EXPECTATIVAS - DEMANDA").Cells(9, Coluna_Demanda_2).Value
ValoresIG_Demanda_3 = Sheets("EXPECTATIVAS - DEMANDA").Cells(9, Coluna_Demanda_3).Value
Sheets("TABELAS").Cells(20, 2).Value = ValoresIG_Demanda_3
Sheets("TABELAS").Cells(20, 3).Value = ValoresIG_Demanda_2
Sheets("TABELAS").Cells(20, 4).Value = ValoresIG_Demanda_1
'Ind�stria Extrativa
ValoresIE_Demanda_1 = Sheets("EXPECTATIVAS - DEMANDA").Cells(21, Coluna_Demanda_1).Value
ValoresIE_Demanda_2 = Sheets("EXPECTATIVAS - DEMANDA").Cells(21, Coluna_Demanda_2).Value
ValoresIE_Demanda_3 = Sheets("EXPECTATIVAS - DEMANDA").Cells(21, Coluna_Demanda_3).Value
Sheets("TABELAS").Cells(23, 2).Value = ValoresIE_Demanda_3
Sheets("TABELAS").Cells(23, 3).Value = ValoresIE_Demanda_2
Sheets("TABELAS").Cells(23, 4).Value = ValoresIE_Demanda_1
'Ind�stria Transforma��o
ValoresIT_Demanda_1 = Sheets("EXPECTATIVAS - DEMANDA").Cells(26, Coluna_Demanda_1).Value
ValoresIT_Demanda_2 = Sheets("EXPECTATIVAS - DEMANDA").Cells(26, Coluna_Demanda_2).Value
ValoresIT_Demanda_3 = Sheets("EXPECTATIVAS - DEMANDA").Cells(26, Coluna_Demanda_3).Value
Sheets("TABELAS").Cells(24, 2).Value = ValoresIT_Demanda_3
Sheets("TABELAS").Cells(24, 3).Value = ValoresIT_Demanda_2
Sheets("TABELAS").Cells(24, 4).Value = ValoresIT_Demanda_1
'Pequena
ValoresP_Demanda_1 = Sheets("EXPECTATIVAS - DEMANDA").Cells(17, Coluna_Demanda_1).Value
ValoresP_Demanda_2 = Sheets("EXPECTATIVAS - DEMANDA").Cells(17, Coluna_Demanda_2).Value
ValoresP_Demanda_3 = Sheets("EXPECTATIVAS - DEMANDA").Cells(17, Coluna_Demanda_3).Value
Sheets("TABELAS").Cells(27, 2).Value = ValoresP_Demanda_3
Sheets("TABELAS").Cells(27, 3).Value = ValoresP_Demanda_2
Sheets("TABELAS").Cells(27, 4).Value = ValoresP_Demanda_1
'M�dia
ValoresM_Demanda_1 = Sheets("EXPECTATIVAS - DEMANDA").Cells(18, Coluna_Demanda_1).Value
ValoresM_Demanda_2 = Sheets("EXPECTATIVAS - DEMANDA").Cells(18, Coluna_Demanda_2).Value
ValoresM_Demanda_3 = Sheets("EXPECTATIVAS - DEMANDA").Cells(18, Coluna_Demanda_3).Value
Sheets("TABELAS").Cells(28, 2).Value = ValoresM_Demanda_3
Sheets("TABELAS").Cells(28, 3).Value = ValoresM_Demanda_2
Sheets("TABELAS").Cells(28, 4).Value = ValoresM_Demanda_1
'Grande
ValoresG_Demanda_1 = Sheets("EXPECTATIVAS - DEMANDA").Cells(19, Coluna_Demanda_1).Value
ValoresG_Demanda_2 = Sheets("EXPECTATIVAS - DEMANDA").Cells(19, Coluna_Demanda_2).Value
ValoresG_Demanda_3 = Sheets("EXPECTATIVAS - DEMANDA").Cells(19, Coluna_Demanda_3).Value
Sheets("TABELAS").Cells(29, 2).Value = ValoresG_Demanda_3
Sheets("TABELAS").Cells(29, 3).Value = ValoresG_Demanda_2
Sheets("TABELAS").Cells(29, 4).Value = ValoresG_Demanda_1

'Atribui os valores da coluna Quantidade exportada
Coluna_Exporta��o_1 = Sheets("EXPECTATIVA - EXPORTA��O").Range("B8").End(xlToRight).Column
Coluna_Exporta��o_2 = Coluna_Exporta��o_1 - 1
Coluna_Exporta��o_3 = Coluna_Exporta��o_1 - 12

'Ind�stria Geral
ValoresIG_Exporta��o_1 = Sheets("EXPECTATIVA - EXPORTA��O").Cells(9, Coluna_Exporta��o_1).Value
ValoresIG_Exporta��o_2 = Sheets("EXPECTATIVA - EXPORTA��O").Cells(9, Coluna_Exporta��o_2).Value
ValoresIG_Exporta��o_3 = Sheets("EXPECTATIVA - EXPORTA��O").Cells(9, Coluna_Exporta��o_3).Value
Sheets("TABELAS").Cells(20, 5).Value = ValoresIG_Exporta��o_3
Sheets("TABELAS").Cells(20, 6).Value = ValoresIG_Exporta��o_2
Sheets("TABELAS").Cells(20, 7).Value = ValoresIG_Exporta��o_1
'Ind�stria Extrativa
ValoresIE_Exporta��o_1 = Sheets("EXPECTATIVA - EXPORTA��O").Cells(21, Coluna_Exporta��o_1).Value
ValoresIE_Exporta��o_2 = Sheets("EXPECTATIVA - EXPORTA��O").Cells(21, Coluna_Exporta��o_2).Value
ValoresIE_Exporta��o_3 = Sheets("EXPECTATIVA - EXPORTA��O").Cells(21, Coluna_Exporta��o_3).Value
Sheets("TABELAS").Cells(23, 5).Value = ValoresIE_Exporta��o_3
Sheets("TABELAS").Cells(23, 6).Value = ValoresIE_Exporta��o_2
Sheets("TABELAS").Cells(23, 7).Value = ValoresIE_Exporta��o_1
'Ind�stria Tansforma��o
ValoresIT_Exporta��o_1 = Sheets("EXPECTATIVA - EXPORTA��O").Cells(26, Coluna_Exporta��o_1).Value
ValoresIT_Exporta��o_2 = Sheets("EXPECTATIVA - EXPORTA��O").Cells(26, Coluna_Exporta��o_2).Value
ValoresIT_Exporta��o_3 = Sheets("EXPECTATIVA - EXPORTA��O").Cells(26, Coluna_Exporta��o_3).Value
Sheets("TABELAS").Cells(24, 5).Value = ValoresIT_Exporta��o_3
Sheets("TABELAS").Cells(24, 6).Value = ValoresIT_Exporta��o_2
Sheets("TABELAS").Cells(24, 7).Value = ValoresIT_Exporta��o_1
'Pequena
ValoresP_Exporta��o_1 = Sheets("EXPECTATIVA - EXPORTA��O").Cells(17, Coluna_Exporta��o_1).Value
ValoresP_Exporta��o_2 = Sheets("EXPECTATIVA - EXPORTA��O").Cells(17, Coluna_Exporta��o_2).Value
ValoresP_Exporta��o_3 = Sheets("EXPECTATIVA - EXPORTA��O").Cells(17, Coluna_Exporta��o_3).Value
Sheets("TABELAS").Cells(27, 5).Value = ValoresP_Exporta��o_3
Sheets("TABELAS").Cells(27, 6).Value = ValoresP_Exporta��o_2
Sheets("TABELAS").Cells(27, 7).Value = ValoresP_Exporta��o_1
'M�dia
ValoresM_Exporta��o_1 = Sheets("EXPECTATIVA - EXPORTA��O").Cells(18, Coluna_Exporta��o_1).Value
ValoresM_Exporta��o_2 = Sheets("EXPECTATIVA - EXPORTA��O").Cells(18, Coluna_Exporta��o_2).Value
ValoresM_Exporta��o_3 = Sheets("EXPECTATIVA - EXPORTA��O").Cells(18, Coluna_Exporta��o_3).Value
Sheets("TABELAS").Cells(28, 5).Value = ValoresM_Exporta��o_3
Sheets("TABELAS").Cells(28, 6).Value = ValoresM_Exporta��o_2
Sheets("TABELAS").Cells(28, 7).Value = ValoresM_Exporta��o_1
'Grande
ValoresG_Exporta��o_1 = Sheets("EXPECTATIVA - EXPORTA��O").Cells(19, Coluna_Exporta��o_1).Value
ValoresG_Exporta��o_2 = Sheets("EXPECTATIVA - EXPORTA��O").Cells(19, Coluna_Exporta��o_2).Value
ValoresG_Exporta��o_3 = Sheets("EXPECTATIVA - EXPORTA��O").Cells(19, Coluna_Exporta��o_3).Value
Sheets("TABELAS").Cells(29, 5).Value = ValoresG_Exporta��o_3
Sheets("TABELAS").Cells(29, 6).Value = ValoresG_Exporta��o_2
Sheets("TABELAS").Cells(29, 7).Value = ValoresG_Exporta��o_1

'Atribui os valores da coluna Compras de mat�ria prima
Coluna_Compras_1 = Sheets("EXPECTATIVA - COMPRAS").Range("B8").End(xlToRight).Column
Coluna_Compras_2 = Coluna_Compras_1 - 1
Coluna_Compras_3 = Coluna_Compras_1 - 12

'Ind�stria Geral
ValoresIG_Compras_1 = Sheets("EXPECTATIVA - COMPRAS").Cells(9, Coluna_Compras_1).Value
ValoresIG_Compras_2 = Sheets("EXPECTATIVA - COMPRAS").Cells(9, Coluna_Compras_2).Value
ValoresIG_Compras_3 = Sheets("EXPECTATIVA - COMPRAS").Cells(9, Coluna_Compras_3).Value
Sheets("TABELAS").Cells(20, 8).Value = ValoresIG_Compras_3
Sheets("TABELAS").Cells(20, 9).Value = ValoresIG_Compras_2
Sheets("TABELAS").Cells(20, 10).Value = ValoresIG_Compras_1
'Ind�stria Extrativa
ValoresIE_Compras_1 = Sheets("EXPECTATIVA - COMPRAS").Cells(21, Coluna_Compras_1).Value
ValoresIE_Compras_2 = Sheets("EXPECTATIVA - COMPRAS").Cells(21, Coluna_Compras_2).Value
ValoresIE_Compras_3 = Sheets("EXPECTATIVA - COMPRAS").Cells(21, Coluna_Compras_3).Value
Sheets("TABELAS").Cells(23, 8).Value = ValoresIE_Compras_3
Sheets("TABELAS").Cells(23, 9).Value = ValoresIE_Compras_2
Sheets("TABELAS").Cells(23, 10).Value = ValoresIE_Compras_1
'Ind�stria Tranforma��o
ValoresIT_Compras_1 = Sheets("EXPECTATIVA - COMPRAS").Cells(26, Coluna_Compras_1).Value
ValoresIT_Compras_2 = Sheets("EXPECTATIVA - COMPRAS").Cells(26, Coluna_Compras_2).Value
ValoresIT_Compras_3 = Sheets("EXPECTATIVA - COMPRAS").Cells(26, Coluna_Compras_3).Value
Sheets("TABELAS").Cells(24, 8).Value = ValoresIT_Compras_3
Sheets("TABELAS").Cells(24, 9).Value = ValoresIT_Compras_2
Sheets("TABELAS").Cells(24, 10).Value = ValoresIT_Compras_1
'Pequena
ValoresP_Compras_1 = Sheets("EXPECTATIVA - COMPRAS").Cells(17, Coluna_Compras_1).Value
ValoresP_Compras_2 = Sheets("EXPECTATIVA - COMPRAS").Cells(17, Coluna_Compras_2).Value
ValoresP_Compras_3 = Sheets("EXPECTATIVA - COMPRAS").Cells(17, Coluna_Compras_3).Value
Sheets("TABELAS").Cells(27, 8).Value = ValoresP_Compras_3
Sheets("TABELAS").Cells(27, 9).Value = ValoresP_Compras_2
Sheets("TABELAS").Cells(27, 10).Value = ValoresP_Compras_1
'M�dia
ValoresM_Compras_1 = Sheets("EXPECTATIVA - COMPRAS").Cells(18, Coluna_Compras_1).Value
ValoresM_Compras_2 = Sheets("EXPECTATIVA - COMPRAS").Cells(18, Coluna_Compras_2).Value
ValoresM_Compras_3 = Sheets("EXPECTATIVA - COMPRAS").Cells(18, Coluna_Compras_3).Value
Sheets("TABELAS").Cells(28, 8).Value = ValoresM_Compras_3
Sheets("TABELAS").Cells(28, 9).Value = ValoresM_Compras_2
Sheets("TABELAS").Cells(28, 10).Value = ValoresM_Compras_1
'Grande
ValoresG_Compras_1 = Sheets("EXPECTATIVA - COMPRAS").Cells(19, Coluna_Compras_1).Value
ValoresG_Compras_2 = Sheets("EXPECTATIVA - COMPRAS").Cells(19, Coluna_Compras_2).Value
ValoresG_Compras_3 = Sheets("EXPECTATIVA - COMPRAS").Cells(19, Coluna_Compras_3).Value
Sheets("TABELAS").Cells(29, 8).Value = ValoresG_Compras_3
Sheets("TABELAS").Cells(29, 9).Value = ValoresG_Compras_2
Sheets("TABELAS").Cells(29, 10).Value = ValoresG_Compras_1

'Atribui os valores da coluna N� de empregados
Coluna_EXEmpregados_1 = Sheets("EXPECTATIVA - EMPREGADOS").Range("B8").End(xlToRight).Column
Coluna_EXEmpregados_2 = Coluna_EXEmpregados_1 - 1
Coluna_EXEmpregados_3 = Coluna_EXEmpregados_1 - 12

'Ind�stria Geral
ValoresIG_EXEmpregados_1 = Sheets("EXPECTATIVA - EMPREGADOS").Cells(9, Coluna_EXEmpregados_1).Value
ValoresIG_EXEmpregados_2 = Sheets("EXPECTATIVA - EMPREGADOS").Cells(9, Coluna_EXEmpregados_2).Value
ValoresIG_EXEmpregados_3 = Sheets("EXPECTATIVA - EMPREGADOS").Cells(9, Coluna_EXEmpregados_3).Value
Sheets("TABELAS").Cells(20, 11).Value = ValoresIG_EXEmpregados_3
Sheets("TABELAS").Cells(20, 12).Value = ValoresIG_EXEmpregados_2
Sheets("TABELAS").Cells(20, 13).Value = ValoresIG_EXEmpregados_1
'Ind�stria Extrativa
ValoresIE_EXEmpregados_1 = Sheets("EXPECTATIVA - EMPREGADOS").Cells(21, Coluna_EXEmpregados_1).Value
ValoresIE_EXEmpregados_2 = Sheets("EXPECTATIVA - EMPREGADOS").Cells(21, Coluna_EXEmpregados_2).Value
ValoresIE_EXEmpregados_3 = Sheets("EXPECTATIVA - EMPREGADOS").Cells(21, Coluna_EXEmpregados_3).Value
Sheets("TABELAS").Cells(23, 11).Value = ValoresIE_EXEmpregados_3
Sheets("TABELAS").Cells(23, 12).Value = ValoresIE_EXEmpregados_2
Sheets("TABELAS").Cells(23, 13).Value = ValoresIE_EXEmpregados_1
'Ind�stria Transforma��o
ValoresIT_EXEmpregados_1 = Sheets("EXPECTATIVA - EMPREGADOS").Cells(26, Coluna_EXEmpregados_1).Value
ValoresIT_EXEmpregados_2 = Sheets("EXPECTATIVA - EMPREGADOS").Cells(26, Coluna_EXEmpregados_2).Value
ValoresIT_EXEmpregados_3 = Sheets("EXPECTATIVA - EMPREGADOS").Cells(26, Coluna_EXEmpregados_3).Value
Sheets("TABELAS").Cells(24, 11).Value = ValoresIT_EXEmpregados_3
Sheets("TABELAS").Cells(24, 12).Value = ValoresIT_EXEmpregados_2
Sheets("TABELAS").Cells(24, 13).Value = ValoresIT_EXEmpregados_1
'Pequena
ValoresP_EXEmpregados_1 = Sheets("EXPECTATIVA - EMPREGADOS").Cells(17, Coluna_EXEmpregados_1).Value
ValoresP_EXEmpregados_2 = Sheets("EXPECTATIVA - EMPREGADOS").Cells(17, Coluna_EXEmpregados_2).Value
ValoresP_EXEmpregados_3 = Sheets("EXPECTATIVA - EMPREGADOS").Cells(17, Coluna_EXEmpregados_3).Value
Sheets("TABELAS").Cells(27, 11).Value = ValoresP_EXEmpregados_3
Sheets("TABELAS").Cells(27, 12).Value = ValoresP_EXEmpregados_2
Sheets("TABELAS").Cells(27, 13).Value = ValoresP_EXEmpregados_1
'M�dia
ValoresM_EXEmpregados_1 = Sheets("EXPECTATIVA - EMPREGADOS").Cells(18, Coluna_EXEmpregados_1).Value
ValoresM_EXEmpregados_2 = Sheets("EXPECTATIVA - EMPREGADOS").Cells(18, Coluna_EXEmpregados_2).Value
ValoresM_EXEmpregados_3 = Sheets("EXPECTATIVA - EMPREGADOS").Cells(18, Coluna_EXEmpregados_3).Value
Sheets("TABELAS").Cells(28, 11).Value = ValoresM_EXEmpregados_3
Sheets("TABELAS").Cells(28, 12).Value = ValoresM_EXEmpregados_2
Sheets("TABELAS").Cells(28, 13).Value = ValoresM_EXEmpregados_1
'Grande
ValoresG_EXEmpregados_1 = Sheets("EXPECTATIVA - EMPREGADOS").Cells(19, Coluna_EXEmpregados_1).Value
ValoresG_EXEmpregados_2 = Sheets("EXPECTATIVA - EMPREGADOS").Cells(19, Coluna_EXEmpregados_2).Value
ValoresG_EXEmpregados_3 = Sheets("EXPECTATIVA - EMPREGADOS").Cells(19, Coluna_EXEmpregados_3).Value
Sheets("TABELAS").Cells(29, 11).Value = ValoresG_EXEmpregados_3
Sheets("TABELAS").Cells(29, 12).Value = ValoresG_EXEmpregados_2
Sheets("TABELAS").Cells(29, 13).Value = ValoresG_EXEmpregados_1

'Atribui os valores da coluna Inten��o de investimento
Coluna_Investimento_1 = Sheets("EXPECTATIVA - INVESTIMENTO").Range("B8").End(xlToRight).Column
Coluna_Investimento_2 = Coluna_Investimento_1 - 1
Coluna_Investimento_3 = Coluna_Investimento_1 - 12

'Ind�stria Geral
ValoresIG_Investimento_1 = Sheets("EXPECTATIVA - INVESTIMENTO").Cells(9, Coluna_Investimento_1).Value
ValoresIG_Investimento_2 = Sheets("EXPECTATIVA - INVESTIMENTO").Cells(9, Coluna_Investimento_2).Value
ValoresIG_Investimento_3 = Sheets("EXPECTATIVA - INVESTIMENTO").Cells(9, Coluna_Investimento_3).Value
Sheets("TABELAS").Cells(20, 14).Value = ValoresIG_Investimento_3
Sheets("TABELAS").Cells(20, 15).Value = ValoresIG_Investimento_2
Sheets("TABELAS").Cells(20, 16).Value = ValoresIG_Investimento_1
'Ind�stria Extrativa
ValoresIE_Investimento_1 = Sheets("EXPECTATIVA - INVESTIMENTO").Cells(21, Coluna_Investimento_1).Value
ValoresIE_Investimento_2 = Sheets("EXPECTATIVA - INVESTIMENTO").Cells(21, Coluna_Investimento_2).Value
ValoresIE_Investimento_3 = Sheets("EXPECTATIVA - INVESTIMENTO").Cells(21, Coluna_Investimento_3).Value
Sheets("TABELAS").Cells(23, 14).Value = ValoresIE_Investimento_3
Sheets("TABELAS").Cells(23, 15).Value = ValoresIE_Investimento_2
Sheets("TABELAS").Cells(23, 16).Value = ValoresIE_Investimento_1
'Ind�stria Transforma��o
ValoresIT_Investimento_1 = Sheets("EXPECTATIVA - INVESTIMENTO").Cells(26, Coluna_Investimento_1).Value
ValoresIT_Investimento_2 = Sheets("EXPECTATIVA - INVESTIMENTO").Cells(26, Coluna_Investimento_2).Value
ValoresIT_Investimento_3 = Sheets("EXPECTATIVA - INVESTIMENTO").Cells(26, Coluna_Investimento_3).Value
Sheets("TABELAS").Cells(24, 14).Value = ValoresIT_Investimento_3
Sheets("TABELAS").Cells(24, 15).Value = ValoresIT_Investimento_2
Sheets("TABELAS").Cells(24, 16).Value = ValoresIT_Investimento_1
'Pequena
ValoresP_Investimento_1 = Sheets("EXPECTATIVA - INVESTIMENTO").Cells(17, Coluna_Investimento_1).Value
ValoresP_Investimento_2 = Sheets("EXPECTATIVA - INVESTIMENTO").Cells(17, Coluna_Investimento_2).Value
ValoresP_Investimento_3 = Sheets("EXPECTATIVA - INVESTIMENTO").Cells(17, Coluna_Investimento_3).Value
Sheets("TABELAS").Cells(27, 14).Value = ValoresP_Investimento_3
Sheets("TABELAS").Cells(27, 15).Value = ValoresP_Investimento_2
Sheets("TABELAS").Cells(27, 16).Value = ValoresP_Investimento_1
'M�dia
ValoresM_Investimento_1 = Sheets("EXPECTATIVA - INVESTIMENTO").Cells(18, Coluna_Investimento_1).Value
ValoresM_Investimento_2 = Sheets("EXPECTATIVA - INVESTIMENTO").Cells(18, Coluna_Investimento_2).Value
ValoresM_Investimento_3 = Sheets("EXPECTATIVA - INVESTIMENTO").Cells(18, Coluna_Investimento_3).Value
Sheets("TABELAS").Cells(28, 14).Value = ValoresM_Investimento_3
Sheets("TABELAS").Cells(28, 15).Value = ValoresM_Investimento_2
Sheets("TABELAS").Cells(28, 16).Value = ValoresM_Investimento_1
'Grande
ValoresG_Investimento_1 = Sheets("EXPECTATIVA - INVESTIMENTO").Cells(19, Coluna_Investimento_1).Value
ValoresG_Investimento_2 = Sheets("EXPECTATIVA - INVESTIMENTO").Cells(19, Coluna_Investimento_2).Value
ValoresG_Investimento_3 = Sheets("EXPECTATIVA - INVESTIMENTO").Cells(19, Coluna_Investimento_3).Value
Sheets("TABELAS").Cells(29, 14).Value = ValoresG_Investimento_3
Sheets("TABELAS").Cells(29, 15).Value = ValoresG_Investimento_2
Sheets("TABELAS").Cells(29, 16).Value = ValoresG_Investimento_1


'*************************************************** C�digo da parte de Condi��es Financeiras **********************************************************

'Nomeia os titulos das colunas e mescla as celulas
Sheets("TABELAS").Cells(31, 1).Value = "Condi��es Financeiras no trimestre"

Sheets("TABELAS").Range(Cells(32, 2), Cells(33, 4)).Merge
Sheets("TABELAS").Cells(32, 2).Value = "MARGEM DE LUCRO OPERACIONAL"

Sheets("TABELAS").Range(Cells(32, 5), Cells(33, 7)).Merge
Sheets("TABELAS").Cells(32, 5).Value = "PRE�O M�DIO DAS MAT�RIAS-PRIMAS"

Sheets("TABELAS").Range(Cells(32, 8), Cells(33, 10)).Merge
Sheets("TABELAS").Cells(32, 8).Value = "SITUA��O FINANCEIRA"

Sheets("TABELAS").Range(Cells(32, 11), Cells(33, 13)).Merge
Sheets("TABELAS").Cells(32, 11).Value = "ACESSO AO CR�DITO"

Sheets("TABELAS").Range(Cells(36, 1), Cells(37, 13)).Merge
Sheets("TABELAS").Cells(36, 1).Value = "POR SEGMENTO INDUSTRIAL"

Sheets("TABELAS").Range(Cells(40, 1), Cells(41, 13)).Merge
Sheets("TABELAS").Cells(41, 1).Value = "POR PORTE"

'Centraliza e alinha os titulos
Sheets("TABELAS").Range("A32:M32").VerticalAlignment = xlCenter
Sheets("TABELAS").Range("A32:M32").HorizontalAlignment = xlCenter

Sheets("TABELAS").Range("A36").VerticalAlignment = xlCenter
Sheets("TABELAS").Range("A36").HorizontalAlignment = xlCenter

Sheets("TABELAS").Range("A40").VerticalAlignment = xlCenter
Sheets("TABELAS").Range("A40").HorizontalAlignment = xlCenter

'Nomeia as linhas
Sheets("TABELAS").Range("A35").Value = "Ind�stria Geral"
Sheets("TABELAS").Range("A38").Value = "Ind�stria extrativa"
Sheets("TABELAS").Range("A39").Value = "Ind�stria de transforma��o"
Sheets("TABELAS").Range("A42").Value = "Pequena"
Sheets("TABELAS").Range("A43").Value = "M�dia"
Sheets("TABELAS").Range("A44").Value = "Grande"

'Define as variavies que ser�o usadas para preencher as celulas
Coluna_Lucro_1 = Sheets("SITUACAO FINANCEIRA LUCRO").Range("B8").End(xlToRight).Column
Coluna_Lucro_2 = Coluna_Lucro_1 - 1
Coluna_Lucro_3 = Coluna_Lucro_1 - 4

'Define, atribui e copia e cola as datas
Data_1 = Sheets("SITUACAO FINANCEIRA LUCRO").Cells(8, Coluna_Lucro_1).Value
Data_2 = Sheets("SITUACAO FINANCEIRA LUCRO").Cells(8, Coluna_Lucro_2).Value
Data_3 = Sheets("SITUACAO FINANCEIRA LUCRO").Cells(8, Coluna_Lucro_3).Value

Sheets("TABELAS").Cells(34, 2).Value = Data_3
Sheets("TABELAS").Cells(34, 3).Value = Data_2
Sheets("TABELAS").Cells(34, 4).Value = Data_1

Sheets("TABELAS").Range("B34:D34").Copy (Sheets("TABELAS").Range("E34"))
Sheets("TABELAS").Range("B34:D34").Copy (Sheets("TABELAS").Range("H34"))
Sheets("TABELAS").Range("B34:D34").Copy (Sheets("TABELAS").Range("K34"))

'Atribui os valores da coluna margem de lucro operacional
'Ind�stria Geral
ValoresIG_Lucro_1 = Sheets("SITUACAO FINANCEIRA LUCRO").Cells(9, Coluna_Lucro_1).Value
ValoresIG_Lucro_2 = Sheets("SITUACAO FINANCEIRA LUCRO").Cells(9, Coluna_Lucro_2).Value
ValoresIG_Lucro_3 = Sheets("SITUACAO FINANCEIRA LUCRO").Cells(9, Coluna_Lucro_3).Value
Sheets("TABELAS").Cells(35, 2).Value = ValoresIG_Lucro_3
Sheets("TABELAS").Cells(35, 3).Value = ValoresIG_Lucro_2
Sheets("TABELAS").Cells(35, 4).Value = ValoresIG_Lucro_1
'Ind�stria Extrativa
ValoresIE_Lucro_1 = Sheets("SITUACAO FINANCEIRA LUCRO").Cells(21, Coluna_Lucro_1).Value
ValoresIE_Lucro_2 = Sheets("SITUACAO FINANCEIRA LUCRO").Cells(21, Coluna_Lucro_2).Value
ValoresIE_Lucro_3 = Sheets("SITUACAO FINANCEIRA LUCRO").Cells(21, Coluna_Lucro_3).Value
Sheets("TABELAS").Cells(38, 2).Value = ValoresIE_Lucro_3
Sheets("TABELAS").Cells(38, 3).Value = ValoresIE_Lucro_2
Sheets("TABELAS").Cells(38, 4).Value = ValoresIE_Lucro_1
'Ind�stria Transforma��o
ValoresIT_Lucro_1 = Sheets("SITUACAO FINANCEIRA LUCRO").Cells(26, Coluna_Lucro_1).Value
ValoresIT_Lucro_2 = Sheets("SITUACAO FINANCEIRA LUCRO").Cells(26, Coluna_Lucro_2).Value
ValoresIT_Lucro_3 = Sheets("SITUACAO FINANCEIRA LUCRO").Cells(26, Coluna_Lucro_3).Value
Sheets("TABELAS").Cells(39, 2).Value = ValoresIT_Lucro_3
Sheets("TABELAS").Cells(39, 3).Value = ValoresIT_Lucro_2
Sheets("TABELAS").Cells(39, 4).Value = ValoresIT_Lucro_1
'Pequena
ValoresP_Lucro_1 = Sheets("SITUACAO FINANCEIRA LUCRO").Cells(17, Coluna_Lucro_1).Value
ValoresP_Lucro_2 = Sheets("SITUACAO FINANCEIRA LUCRO").Cells(17, Coluna_Lucro_2).Value
ValoresP_Lucro_3 = Sheets("SITUACAO FINANCEIRA LUCRO").Cells(17, Coluna_Lucro_3).Value
Sheets("TABELAS").Cells(42, 2).Value = ValoresP_Lucro_3
Sheets("TABELAS").Cells(42, 3).Value = ValoresP_Lucro_2
Sheets("TABELAS").Cells(42, 4).Value = ValoresP_Lucro_1
'M�dia
ValoresM_Lucro_1 = Sheets("SITUACAO FINANCEIRA LUCRO").Cells(18, Coluna_Lucro_1).Value
ValoresM_Lucro_2 = Sheets("SITUACAO FINANCEIRA LUCRO").Cells(18, Coluna_Lucro_2).Value
ValoresM_Lucro_3 = Sheets("SITUACAO FINANCEIRA LUCRO").Cells(18, Coluna_Lucro_3).Value
Sheets("TABELAS").Cells(43, 2).Value = ValoresM_Lucro_3
Sheets("TABELAS").Cells(43, 3).Value = ValoresM_Lucro_2
Sheets("TABELAS").Cells(43, 4).Value = ValoresM_Lucro_1
'Grande
ValoresG_Lucro_1 = Sheets("SITUACAO FINANCEIRA LUCRO").Cells(19, Coluna_Lucro_1).Value
ValoresG_Lucro_2 = Sheets("SITUACAO FINANCEIRA LUCRO").Cells(19, Coluna_Lucro_2).Value
ValoresG_Lucro_3 = Sheets("SITUACAO FINANCEIRA LUCRO").Cells(19, Coluna_Lucro_3).Value
Sheets("TABELAS").Cells(44, 2).Value = ValoresG_Lucro_3
Sheets("TABELAS").Cells(44, 3).Value = ValoresG_Lucro_2
Sheets("TABELAS").Cells(44, 4).Value = ValoresG_Lucro_1

'Atribui os valores da coluna Pre�o m�dio de mat�rias primas
Coluna_Pre�o_1 = Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Range("B8").End(xlToRight).Column
Coluna_Pre�o_2 = Coluna_Pre�o_1 - 1
Coluna_Pre�o_3 = Coluna_Pre�o_1 - 12

'Ind�stria Geral
ValoresIG_Pre�o_1 = Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(9, Coluna_Pre�o_1).Value
ValoresIG_Pre�o_2 = Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(9, Coluna_Pre�o_2).Value
ValoresIG_Pre�o_3 = Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(9, Coluna_Pre�o_3).Value
Sheets("TABELAS").Cells(35, 5).Value = ValoresIG_Pre�o_3
Sheets("TABELAS").Cells(35, 6).Value = ValoresIG_Pre�o_2
Sheets("TABELAS").Cells(35, 7).Value = ValoresIG_Pre�o_1
'Ind�stria Extrativa
ValoresIE_Pre�o_1 = Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(21, Coluna_Pre�o_1).Value
ValoresIE_Pre�o_2 = Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(21, Coluna_Pre�o_2).Value
ValoresIE_Pre�o_3 = Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(21, Coluna_Pre�o_3).Value
Sheets("TABELAS").Cells(38, 5).Value = ValoresIE_Pre�o_3
Sheets("TABELAS").Cells(38, 6).Value = ValoresIE_Pre�o_2
Sheets("TABELAS").Cells(38, 7).Value = ValoresIE_Pre�o_1
'Ind�stria Tansforma��o
ValoresIT_Pre�o_1 = Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(26, Coluna_Pre�o_1).Value
ValoresIT_Pre�o_2 = Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(26, Coluna_Pre�o_2).Value
ValoresIT_Pre�o_3 = Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(26, Coluna_Pre�o_3).Value
Sheets("TABELAS").Cells(39, 5).Value = ValoresIT_Pre�o_3
Sheets("TABELAS").Cells(39, 6).Value = ValoresIT_Pre�o_2
Sheets("TABELAS").Cells(39, 7).Value = ValoresIT_Pre�o_1
'Pequena
ValoresP_Pre�o_1 = Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(17, Coluna_Pre�o_1).Value
ValoresP_Pre�o_2 = Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(17, Coluna_Pre�o_2).Value
ValoresP_Pre�o_3 = Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(17, Coluna_Pre�o_3).Value
Sheets("TABELAS").Cells(42, 5).Value = ValoresP_Pre�o_3
Sheets("TABELAS").Cells(42, 6).Value = ValoresP_Pre�o_2
Sheets("TABELAS").Cells(42, 7).Value = ValoresP_Pre�o_1
'M�dia
ValoresM_Pre�o_1 = Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(18, Coluna_Pre�o_1).Value
ValoresM_Pre�o_2 = Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(18, Coluna_Pre�o_2).Value
ValoresM_Pre�o_3 = Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(18, Coluna_Pre�o_3).Value
Sheets("TABELAS").Cells(43, 5).Value = ValoresM_Pre�o_3
Sheets("TABELAS").Cells(43, 6).Value = ValoresM_Pre�o_2
Sheets("TABELAS").Cells(43, 7).Value = ValoresM_Pre�o_1
'Grande
ValoresG_Pre�o_1 = Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(19, Coluna_Pre�o_1).Value
ValoresG_Pre�o_2 = Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(19, Coluna_Pre�o_2).Value
ValoresG_Pre�o_3 = Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(19, Coluna_Pre�o_3).Value
Sheets("TABELAS").Cells(44, 5).Value = ValoresG_Pre�o_3
Sheets("TABELAS").Cells(44, 6).Value = ValoresG_Pre�o_2
Sheets("TABELAS").Cells(44, 7).Value = ValoresG_Pre�o_1

'Atribui os valores da coluna Situa��o FInanceira
Coluna_Situa��o_1 = Sheets("SITUACAO FINANCEIRA").Range("B8").End(xlToRight).Column
Coluna_Situa��o_2 = Coluna_Situa��o_1 - 1
Coluna_Situa��o_3 = Coluna_Situa��o_1 - 12

'Ind�stria Geral
ValoresIG_Situa��o_1 = Sheets("SITUACAO FINANCEIRA").Cells(9, Coluna_Situa��o_1).Value
ValoresIG_Situa��o_2 = Sheets("SITUACAO FINANCEIRA").Cells(9, Coluna_Situa��o_2).Value
ValoresIG_Situa��o_3 = Sheets("SITUACAO FINANCEIRA").Cells(9, Coluna_Situa��o_3).Value
Sheets("TABELAS").Cells(35, 8).Value = ValoresIG_Situa��o_3
Sheets("TABELAS").Cells(35, 9).Value = ValoresIG_Situa��o_2
Sheets("TABELAS").Cells(35, 10).Value = ValoresIG_Situa��o_1
'Ind�stria Extrativa
ValoresIE_Situa��o_1 = Sheets("SITUACAO FINANCEIRA").Cells(21, Coluna_Situa��o_1).Value
ValoresIE_Situa��o_2 = Sheets("SITUACAO FINANCEIRA").Cells(21, Coluna_Situa��o_2).Value
ValoresIE_Situa��o_3 = Sheets("SITUACAO FINANCEIRA").Cells(21, Coluna_Situa��o_3).Value
Sheets("TABELAS").Cells(38, 8).Value = ValoresIE_Situa��o_3
Sheets("TABELAS").Cells(38, 9).Value = ValoresIE_Situa��o_2
Sheets("TABELAS").Cells(38, 10).Value = ValoresIE_Situa��o_1
'Ind�stria Tranforma��o
ValoresIT_Situa��o_1 = Sheets("SITUACAO FINANCEIRA").Cells(26, Coluna_Situa��o_1).Value
ValoresIT_Situa��o_2 = Sheets("SITUACAO FINANCEIRA").Cells(26, Coluna_Situa��o_2).Value
ValoresIT_Situa��o_3 = Sheets("SITUACAO FINANCEIRA").Cells(26, Coluna_Situa��o_3).Value
Sheets("TABELAS").Cells(39, 8).Value = ValoresIT_Situa��o_3
Sheets("TABELAS").Cells(39, 9).Value = ValoresIT_Situa��o_2
Sheets("TABELAS").Cells(39, 10).Value = ValoresIT_Situa��o_1
'Pequena
ValoresP_Situa��o_1 = Sheets("SITUACAO FINANCEIRA").Cells(17, Coluna_Situa��o_1).Value
ValoresP_Situa��o_2 = Sheets("SITUACAO FINANCEIRA").Cells(17, Coluna_Situa��o_2).Value
ValoresP_Situa��o_3 = Sheets("SITUACAO FINANCEIRA").Cells(17, Coluna_Situa��o_3).Value
Sheets("TABELAS").Cells(42, 8).Value = ValoresP_Situa��o_3
Sheets("TABELAS").Cells(42, 9).Value = ValoresP_Situa��o_2
Sheets("TABELAS").Cells(42, 10).Value = ValoresP_Situa��o_1
'M�dia
ValoresM_Situa��o_1 = Sheets("SITUACAO FINANCEIRA").Cells(18, Coluna_Situa��o_1).Value
ValoresM_Situa��o_2 = Sheets("SITUACAO FINANCEIRA").Cells(18, Coluna_Situa��o_2).Value
ValoresM_Situa��o_3 = Sheets("SITUACAO FINANCEIRA").Cells(18, Coluna_Situa��o_3).Value
Sheets("TABELAS").Cells(43, 8).Value = ValoresM_Situa��o_3
Sheets("TABELAS").Cells(43, 9).Value = ValoresM_Situa��o_2
Sheets("TABELAS").Cells(43, 10).Value = ValoresM_Situa��o_1
'Grande
ValoresG_Situa��o_1 = Sheets("SITUACAO FINANCEIRA").Cells(19, Coluna_Situa��o_1).Value
ValoresG_Situa��o_2 = Sheets("SITUACAO FINANCEIRA").Cells(19, Coluna_Situa��o_2).Value
ValoresG_Situa��o_3 = Sheets("SITUACAO FINANCEIRA").Cells(19, Coluna_Situa��o_3).Value
Sheets("TABELAS").Cells(44, 8).Value = ValoresG_Situa��o_3
Sheets("TABELAS").Cells(44, 9).Value = ValoresG_Situa��o_2
Sheets("TABELAS").Cells(44, 10).Value = ValoresG_Situa��o_1

'Atribui os valores da coluna Acesso ao cr�dito
Coluna_Cr�dito_1 = Sheets("SITUACAO FINANCEIRA CREDITO").Range("B8").End(xlToRight).Column
Coluna_Cr�dito_2 = Coluna_Cr�dito_1 - 1
Coluna_Cr�dito_3 = Coluna_Cr�dito_1 - 12

'Ind�stria Geral
ValoresIG_Cr�dito_1 = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(9, Coluna_Cr�dito_1).Value
ValoresIG_Cr�dito_2 = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(9, Coluna_Cr�dito_2).Value
ValoresIG_Cr�dito_3 = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(9, Coluna_Cr�dito_3).Value
Sheets("TABELAS").Cells(35, 11).Value = ValoresIG_Cr�dito_3
Sheets("TABELAS").Cells(35, 12).Value = ValoresIG_Cr�dito_2
Sheets("TABELAS").Cells(35, 13).Value = ValoresIG_Cr�dito_1
'Ind�stria Extrativa
ValoresIE_Cr�dito_1 = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(21, Coluna_Cr�dito_1).Value
ValoresIE_Cr�dito_2 = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(21, Coluna_Cr�dito_2).Value
ValoresIE_Cr�dito_3 = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(21, Coluna_Cr�dito_3).Value
Sheets("TABELAS").Cells(38, 11).Value = ValoresIE_Cr�dito_3
Sheets("TABELAS").Cells(38, 12).Value = ValoresIE_Cr�dito_2
Sheets("TABELAS").Cells(38, 13).Value = ValoresIE_Cr�dito_1
'Ind�stria Transforma��o
ValoresIT_Cr�dito_1 = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(26, Coluna_Cr�dito_1).Value
ValoresIT_Cr�dito_2 = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(26, Coluna_Cr�dito_2).Value
ValoresIT_Cr�dito_3 = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(26, Coluna_Cr�dito_3).Value
Sheets("TABELAS").Cells(39, 11).Value = ValoresIT_Cr�dito_3
Sheets("TABELAS").Cells(39, 12).Value = ValoresIT_Cr�dito_2
Sheets("TABELAS").Cells(39, 13).Value = ValoresIT_Cr�dito_1
'Pequena
ValoresP_Cr�dito_1 = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(17, Coluna_Cr�dito_1).Value
ValoresP_Cr�dito_2 = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(17, Coluna_Cr�dito_2).Value
ValoresP_Cr�dito_3 = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(17, Coluna_Cr�dito_3).Value
Sheets("TABELAS").Cells(42, 11).Value = ValoresP_Cr�dito_3
Sheets("TABELAS").Cells(42, 12).Value = ValoresP_Cr�dito_2
Sheets("TABELAS").Cells(42, 13).Value = ValoresP_Cr�dito_1
'M�dia
ValoresM_Cr�dito_1 = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(18, Coluna_Cr�dito_1).Value
ValoresM_Cr�dito_2 = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(18, Coluna_Cr�dito_2).Value
ValoresM_Cr�dito_3 = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(18, Coluna_Cr�dito_3).Value
Sheets("TABELAS").Cells(43, 11).Value = ValoresM_Cr�dito_3
Sheets("TABELAS").Cells(43, 12).Value = ValoresM_Cr�dito_2
Sheets("TABELAS").Cells(43, 13).Value = ValoresM_Cr�dito_1
'Grande
ValoresG_Cr�dito_1 = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(19, Coluna_Cr�dito_1).Value
ValoresG_Cr�dito_2 = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(19, Coluna_Cr�dito_2).Value
ValoresG_Cr�dito_3 = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(19, Coluna_Cr�dito_3).Value
Sheets("TABELAS").Cells(44, 11).Value = ValoresG_Cr�dito_3
Sheets("TABELAS").Cells(44, 12).Value = ValoresG_Cr�dito_2
Sheets("TABELAS").Cells(44, 13).Value = ValoresG_Cr�dito_1

'*******************************************************Princiapais Problemas******************************************************


Sheets("PRINCIPAIS_PROBLEMAS").Select
Range("C109").Value = "Geral"
Range("C109:E109").Merge
Range("F109").Value = "Pequenas"
Range("F109:H109").Merge
Range("F109").Value = "Pequenas"
Range("F109:H109").Merge
Range("I109").Value = "M�dia"
Range("I109:K109").Merge
Range("L109").Value = "Grande"
Range("L109:N109").Merge
Range("E110").Value = "Rank"

Range("C110:E110").Copy
Range("F110").PasteSpecial
Range("I110").PasteSpecial
Range("L110").PasteSpecial

Range("E127").Value = "-"
Range("E128").Value = "-"

Range("E127:E128").Copy
Range("H127:H128").PasteSpecial
Range("N127:N128").PasteSpecial
Range("K127:K128").PasteSpecial

Rows("111:111").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
     
     
Range("B110").Clear
Range("B111").Value = "Itens"
 
Range("C111").Value = "%"
Range("D111").Value = "%"
Range("E111").Value = "Posi��o"

Range("C111:E111").Copy
Range("F111:H111").PasteSpecial
Range("I111:K111").PasteSpecial
Range("L111:M111").PasteSpecial


Coluna_Ultimo_Tri = Sheets("PRINCIPAIS_PROBLEMAS").Range("C10").End(xlToRight).Column
Coluna_Tri_Anterior = Coluna_Ultimo_Tri - 1
linha = 112

'Rank geral
Do Until linha = 128
posi��oG = Application.WorksheetFunction.Rank_Eq(Cells(linha, 4), Range("D112:D127").Cells, 0)
Cells(linha, 5).Value = posi��oG
linha = linha + 1
Loop

'Proc v Pequenas 1
linha = 112
Do Until linha = 130
Valor = Application.WorksheetFunction.VLookup(Cells(linha, 2), Range(Cells(37, 2), Cells(54, Coluna_Ultimo_Tri)), Coluna_Tri_Anterior - 1, 0)
Cells(linha, 6).Value = Valor
linha = linha + 1
Loop

'Proc v Pequenas 2
linha = 112
Do Until linha = 130
Valor = Application.WorksheetFunction.VLookup(Cells(linha, 2), Range(Cells(37, 2), Cells(54, Coluna_Ultimo_Tri)), Coluna_Ultimo_Tri - 1, 0)
Cells(linha, 7).Value = Valor
linha = linha + 1
Loop

'Rank Pequenas
linha = 112
Do Until linha = 128
posi��oP = Application.WorksheetFunction.Rank_Eq(Cells(linha, 7), Range("G112:G127").Cells, 0)
Cells(linha, 8).Value = posi��oP
linha = linha + 1
Loop

'Proc v medias 1
linha = 112
Do Until linha = 130
Valor = Application.WorksheetFunction.VLookup(Cells(linha, 2), Range(Cells(61, 2), Cells(78, Coluna_Ultimo_Tri)), Coluna_Tri_Anterior - 1, 0)
Cells(linha, 9).Value = Valor
linha = linha + 1
Loop

'Proc v medias 2
linha = 112
Do Until linha = 130
Valor = Application.WorksheetFunction.VLookup(Cells(linha, 2), Range(Cells(61, 2), Cells(78, Coluna_Ultimo_Tri)), Coluna_Ultimo_Tri - 1, 0)
Cells(linha, 10).Value = Valor
linha = linha + 1
Loop

'Rank medias
linha = 112
Do Until linha = 128
posi��oM = Application.WorksheetFunction.Rank_Eq(Cells(linha, 10), Range("J112:J127").Cells, 0)
Cells(linha, 11).Value = posi��oM
linha = linha + 1
Loop

'Proc v Grandes 1
linha = 112
Do Until linha = 130
Valor = Application.WorksheetFunction.VLookup(Cells(linha, 2), Range(Cells(85, 2), Cells(102, Coluna_Ultimo_Tri)), Coluna_Tri_Anterior - 1, 0)
Cells(linha, 12).Value = Valor
linha = linha + 1
Loop

'Proc v Grandes 2
linha = 112
Do Until linha = 130
Valor = Application.WorksheetFunction.VLookup(Cells(linha, 2), Range(Cells(85, 2), Cells(102, Coluna_Ultimo_Tri)), Coluna_Ultimo_Tri - 1, 0)
Cells(linha, 13).Value = Valor
linha = linha + 1
Loop

'Rank Grandes
linha = 112
Do Until linha = 128
posi��oGr = Application.WorksheetFunction.Rank_Eq(Cells(linha, 13), Range("M112:M127").Cells, 0)
Cells(linha, 14).Value = posi��oGr
linha = linha + 1
Loop

Sheets("PRINCIPAIS_PROBLEMAS").Select
Range("B109:N129").Copy
Sheets("TABELAS").Select
Range("V2").PasteSpecial
Range("V1").Value = "Principais Problemas"

End Sub


Sub An�lise_Vermelho()

Dim Linha_Dados As Integer 'Define a linha que cont�m o dado a ser usado
Dim Coluna_Dados1 As Integer ' Define a coluna com o dado mais recente
Dim Coluna_Dados2 As Integer ' Define a coluna com o dado do m�s anterior
Dim Coluna_Dados3 As Integer ' Defie a coluna com o dado do mesmo m�s do ano anterior
Dim Linha_An�lise As Integer ' Define a linha que ser� feita a an�lise
Dim Coluna_An�lise As Integer 'Define a coluna que ser� feita a an�lise

Sheets("PRODU��O").Select


'Copia os t�tulos das categorias e cola onde ser� formada a tabela de an�lise
Range(Cells(9, 1), Cells(54, 9)).Copy (Cells(59, 1))
'Limpa os n�meros que foram colados mas mant�m a formata��o
Range(Cells(59, 2), Cells(105, 9)).ClearContents



'Nomeia as colunas de acordo com o dado que ser� calculado nelas
Cells(58, 2).Value = "Diferen�a para o m�s anterior"
Cells(58, 3).Value = "Diferen�a para ao mesmo m�s do ano anterior"
Cells(58, 4).Value = "Diferen�a para a m�dia hist�rica"
Cells(58, 5).Value = "Posi��o Decrescente (Maior valor 1�, menor valor �ltimo)"
Cells(58, 6).Value = "Posi��o Crescente  (Menor valor 1�, maior valor �ltimo)"
Cells(58, 7).Value = "Posi��o Crescente - Mesmo m�s  (Menor valor 1�, maior valor �ltimo)"
Cells(58, 8).Value = "Posi��o Decrescente -Mesmo m�s  (Maior valor 1�, menor valor �ltimo)"
Cells(58, 9).Value = "cruzou a linha de 50?"

'Atribui valores as variaveis definidas acima
Linha_Dados = 9 'Define o n�mero da primeira linha de dados
Coluna_Dados1 = Sheets("PRODU��O").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Coluna_Dados2 = Coluna_Dados1 - 1 'Define o n�mero da coluna do m�s anterior
Coluna_Dados3 = Coluna_Dados1 - 12
Linha_An�lise = 59 'Define a primeira linhas de an�lises
Coluna_An�lise = 2 'Define a coluna de an�lises

'Inserindo valores nas celulas vazias para fugir de bugs
Range(Cells(10, Coluna_Dados3), Cells(10, Coluna_Dados1)).Value = "0"
Range(Cells(16, Coluna_Dados3), Cells(16, Coluna_Dados1)).Value = "0"
Range(Cells(20, Coluna_Dados3), Cells(20, Coluna_Dados1)).Value = "0"
Range(Cells(22, Coluna_Dados3), Cells(23, Coluna_Dados1)).Value = "0"
Range(Cells(25, Coluna_Dados3), Cells(25, Coluna_Dados1)).Value = "0"
Range(Cells(29, Coluna_Dados3), Cells(29, Coluna_Dados1)).Value = "0"
Range(Cells(37, Coluna_Dados3), Cells(37, Coluna_Dados1)).Value = "0"


'Calculo da difern�a em pontos do valor mais recente em rela��o ao valor do m�s anterior
Do Until Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Calculo da diferen�a em si: o valor da celula de analise � igual ao valor mais recente menos o valor do m�s anterior
   Sheets("PRODU��O").Cells(Linha_An�lise, Coluna_An�lise).Value = Sheets("PRODU��O").Cells(Linha_Dados, Coluna_Dados1).Value - Sheets("PRODU��O").Cells(Linha_Dados, Coluna_Dados2).Value
    'Vai para a pr�xima linha de dados e de an�lise
   Linha_Dados = Linha_Dados + 1
   Linha_An�lise = Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Linha_Dados = 9
Coluna_Dados1 = Sheets("PRODU��O").Range("A9").End(xlToRight).Column
Coluna_Dados3 = Coluna_Dados1 - 12
Linha_An�lise = 59
Coluna_An�lise = 3

'C�lculo da diferen�a em pontos do valor mais recente em rela��o ao valor do mesmo m�s do ano anterior
Do Until Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Calculo da diferen�a em si: o valor da celula de analise � igual ao valor mais recente menos o valor do mesmo m�s do ano anterior
    Sheets("PRODU��O").Cells(Linha_An�lise, Coluna_An�lise).Value = Sheets("PRODU��O").Cells(Linha_Dados, Coluna_Dados1).Value - Sheets("PRODU��O").Cells(Linha_Dados, Coluna_Dados3).Value
    'Vai para a pr�xima linha de dados e de an�lise
    Linha_Dados = Linha_Dados + 1
    Linha_An�lise = Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Linha_Dados = 9
Coluna_Dados1 = Sheets("PRODU��O").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Linha_An�lise = 59
Coluna_An�lise = 4

'C�lculo da diferen�a em pontos do valor mais recente em rela��o ao valor da m�dia hist�rica
Do Until Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a vari�vel media como a m�dia do intervalo entre a coluna com o dado mais recente e o primeiro
    media = Application.Average(Range(Cells(Linha_Dados, Coluna_Dados1), Cells(Linha_Dados, 2)))
    'Calculo em si: o valor da celula de analise � igual ao valor mais recente menos o valor da m�dia
    Sheets("PRODU��O").Cells(Linha_An�lise, Coluna_An�lise).Value = Sheets("PRODU��O").Cells(Linha_Dados, Coluna_Dados1).Value - media
    'Vai para a pr�xima linha de dados e de an�lise
    Linha_Dados = Linha_Dados + 1
    Linha_An�lise = Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Linha_Dados = 9
Coluna_Dados1 = Sheets("PRODU��O").Range("A9").End(xlToRight).Column
Linha_An�lise = 59
Coluna_An�lise = 5

'Ordena��o decrescente da s�rie hist�rica completa
Do Until Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posi��o = WorksheetFunction.Rank_Eq(Cells(Linha_Dados, Coluna_Dados1), Range(Cells(Linha_Dados, Coluna_Dados1), Cells(Linha_Dados, 2)), 0)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("PRODU��O").Cells(Linha_An�lise, Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Linha_Dados = Linha_Dados + 1
    Linha_An�lise = Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Linha_Dados = 9
Coluna_Dados1 = Sheets("PRODU��O").Range("A9").End(xlToRight).Column
Linha_An�lise = 59
Coluna_An�lise = 6

'Ordena��o Crescente da s�rie hist�rica completa
Do Until Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posi��o = WorksheetFunction.Rank_Eq(Cells(Linha_Dados, Coluna_Dados1), Range(Cells(Linha_Dados, Coluna_Dados1), Cells(Linha_Dados, 2)), 1)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("PRODU��O").Cells(Linha_An�lise, Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Linha_Dados = Linha_Dados + 1
    Linha_An�lise = Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior e define a vari�vel Coluna_DadosP que representa a primeira coluna de dados
Coluna_Dados1 = Sheets("PRODU��O").Range("A9").End(xlToRight).Column
Coluna_DadosP = 2

Do Until Coluna_DadosP = Coluna_Dados1 + 1 ' Faz at� a variavel Coluna_DadosP ser igual a variavel Coluna_Dados1 mais uma unidade
    'Confere se o m�s da coluna em quest�o � igual ao m�s do dado mais recente
    If Month(Cells(8, Coluna_DadosP)) = Month(Cells(8, Coluna_Dados1)) Then
    'Caso seja igual, copia a coluna com os dados mais abaixo, a partir da linha 110
        Range(Cells(9, Coluna_DadosP), (Cells(54, Coluna_DadosP))).Copy (Cells(110, Coluna_DadosP))
    End If
    'Vai para a pr�xima coluna
    Coluna_DadosP = Coluna_DadosP + 1
'Repete a conferencia com a pr�xima coluna
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Linha_Dados = 110
Coluna_Dados1 = Sheets("PRODU��O").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Linha_An�lise = 59
Coluna_An�lise = 7

'Ordena��o decrescente da s�rie hist�rica dos meses do dado mais recente
Do Until Linha_Dados = 156 'Faz o calculo at� a vari�vel Linha_Dados ser 156
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo m�s do mais recente
    posi��o = Application.WorksheetFunction.Rank_Eq(Cells(Linha_Dados, Coluna_Dados1), Range(Cells(Linha_Dados, Coluna_Dados1), Cells(Linha_Dados, 2)))
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("PRODU��O").Cells(Linha_An�lise, Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Linha_Dados = Linha_Dados + 1
    Linha_An�lise = Linha_An�lise + 1
'Repete a ordena��o com a pr�xima linha
Loop

'Atribui os valores originais das variaveis ap�s o loop anterior
Linha_Dados = 110
Coluna_Dados1 = Sheets("PRODU��O").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Linha_An�lise = 59
Coluna_An�lise = 8
'Ordena��o crescente da s�rie hist�rica dos meses do dado mais recente
Do Until Linha_Dados = 156 'Faz o calculo at� a vari�vel Linha_Dados ser 156
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo m�s do mais recente
    posi��o = Application.WorksheetFunction.Rank_Eq(Cells(Linha_Dados, Coluna_Dados1), Range(Cells(Linha_Dados, Coluna_Dados1), Cells(Linha_Dados, 2)), 1)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("PRODU��O").Cells(Linha_An�lise, Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Linha_Dados = Linha_Dados + 1
    Linha_An�lise = Linha_An�lise + 1
'Repete a ordena��o com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Linha_Dados = 9
Coluna_Dados1 = Sheets("PRODU��O").Range("A9").End(xlToRight).Column
Coluna_Dados2 = Coluna_Dados1 - 1
Linha_An�lise = 59
Coluna_An�lise = 9

'Avalia��o se cruzou ou n�o a linha de 50 e o sentido
Do Until Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'se o dado mais recente for menor que 50 e o dado do m�s anterior for maior ou igual a 50 ent�o...
    If Cells(Linha_Dados, Coluna_Dados1) < 50 And Cells(Linha_Dados, Coluna_Dados2) >= 50 Then
    'a c�lula de an�lise recebe cruzou para baixo
    Cells(Linha_An�lise, Coluna_An�lise).Value = "Cruzou para baixo"
    'Caso n�o seja..
    Else
        'se o dado mais recente for maior ou igual a 50 e o dado do m~es anterior for menor ou igual a 50 ent�o...
        If Cells(Linha_Dados, Coluna_Dados1) >= 50 And Cells(Linha_Dados, Coluna_Dados2) <= 50 Then
        'a c�lula de an�lise recebe cruzou para cima
        Cells(Linha_An�lise, Coluna_An�lise).Value = "Cruzou para cima"
        'Caso n�o seja..
        Else
        'a c�lula de an�lise recebe n�o cruzou
        Cells(Linha_An�lise, Coluna_An�lise).Value = "N�o Cruzou"
        End If
    End If
    'Vai para a pr�xima linha de dados e de an�lise
    Linha_Dados = Linha_Dados + 1
    Linha_An�lise = Linha_An�lise + 1
'repete o processo com a nova linha
Loop

'Apaga as linhas com erros/dados faltantes/t�tulos e subt�tulos
Range(Cells(60, 2), Cells(60, 9)).ClearContents
Range(Cells(66, 2), Cells(66, 9)).ClearContents
Range(Cells(70, 2), Cells(70, 9)).ClearContents
Range(Cells(72, 2), Cells(73, 9)).Value = "-"
Range(Cells(75, 2), Cells(75, 9)).Value = "-"
Range(Cells(79, 2), Cells(79, 9)).Value = "-"
Range(Cells(87, 2), Cells(87, 9)).Value = "-"

'Inserindo valores nas celulas vazias para fugir de bugs
Range(Cells(10, Coluna_Dados3), Cells(10, Coluna_Dados1)).ClearContents
Range(Cells(16, Coluna_Dados3), Cells(16, Coluna_Dados1)).ClearContents
Range(Cells(20, Coluna_Dados3), Cells(20, Coluna_Dados1)).ClearContents
Range(Cells(22, Coluna_Dados3), Cells(23, Coluna_Dados1)).Value = "-"
Range(Cells(25, Coluna_Dados3), Cells(25, Coluna_Dados1)).Value = "-"
Range(Cells(29, Coluna_Dados3), Cells(29, Coluna_Dados1)).Value = "-"
Range(Cells(37, Coluna_Dados3), Cells(37, Coluna_Dados1)).Value = "-"

Range("E59:H104").NumberFormat = "0"



'**********************************************              An�lise_Empregados                ***********************************************************************


Dim Empregados_Linha_Dados As Integer 'Define a linha que cont�m o dado a ser usado
Dim Empregados_Coluna_Dados1 As Integer ' Define a coluna com o dado mais recente
Dim Empregados_Coluna_Dados2 As Integer ' Define a coluna com o dado do m�s anterior
Dim Empregados_Coluna_Dados3 As Integer ' Defie a coluna com o dado do mesmo m�s do ano anterior
Dim Empregados_Linha_An�lise As Integer ' Define a linha que ser� feita a an�lise
Dim Empregados_Coluna_An�lise As Integer 'Define a coluna que ser� feita a an�lise


Sheets("EMPREGADOS").Select


'Copia os t�tulos das categorias e cola onde ser� formada a tabela de an�lise
Sheets("EMPREGADOS").Range(Cells(9, 1), Cells(54, 9)).Copy (Sheets("EMPREGADOS").Cells(59, 1))
'Limpa os n�meros que foram colados mas mant�m a formata��o
Sheets("EMPREGADOS").Range(Cells(59, 2), Cells(105, 9)).ClearContents



'Nomeia as colunas de acordo com o dado que ser� calculado nelas
Sheets("EMPREGADOS").Cells(58, 2).Value = "Diferen�a para o m�s anterior"
Sheets("EMPREGADOS").Cells(58, 3).Value = "Diferen�a para ao mesmo m�s do ano anterior"
Sheets("EMPREGADOS").Cells(58, 4).Value = "Diferen�a para a m�dia hist�rica"
Sheets("EMPREGADOS").Cells(58, 5).Value = "Posi��o Decrescente (Maior valor 1�, menor valor �ltimo)"
Sheets("EMPREGADOS").Cells(58, 6).Value = "Posi��o Crescente  (Menor valor 1�, maior valor �ltimo)"
Sheets("EMPREGADOS").Cells(58, 7).Value = "Posi��o Crescente - Mesmo m�s  (Menor valor 1�, maior valor �ltimo)"
Sheets("EMPREGADOS").Cells(58, 8).Value = "Posi��o Decrescente -Mesmo m�s  (Maior valor 1�, menor valor �ltimo)"
Sheets("EMPREGADOS").Cells(58, 9).Value = "cruzou a linha de 50?"

'Atribui valores as variaveis definidas acima
Empregados_Linha_Dados = 9 'Define o n�mero da primeira linha de dados
Empregados_Coluna_Dados1 = Sheets("EMPREGADOS").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Empregados_Coluna_Dados2 = Empregados_Coluna_Dados1 - 1 'Define o n�mero da coluna do m�s anterior
Empregados_Coluna_Dados3 = Empregados_Coluna_Dados1 - 12
Empregados_Linha_An�lise = 59 'Define a primeira linhas de an�lises
Empregados_Coluna_An�lise = 2 'Define a coluna de an�lises

'Inserindo valores nas celulas vazias para fugir de bugs
Sheets("EMPREGADOS").Range(Cells(10, Empregados_Coluna_Dados3), Cells(10, Empregados_Coluna_Dados1)).Value = "0"
Sheets("EMPREGADOS").Range(Cells(16, Empregados_Coluna_Dados3), Cells(16, Empregados_Coluna_Dados1)).Value = "0"
Sheets("EMPREGADOS").Range(Cells(20, Empregados_Coluna_Dados3), Cells(20, Empregados_Coluna_Dados1)).Value = "0"
Sheets("EMPREGADOS").Range(Cells(22, Empregados_Coluna_Dados3), Cells(23, Empregados_Coluna_Dados1)).Value = "0"
Sheets("EMPREGADOS").Range(Cells(25, Empregados_Coluna_Dados3), Cells(25, Empregados_Coluna_Dados1)).Value = "0"
Sheets("EMPREGADOS").Range(Cells(29, Empregados_Coluna_Dados3), Cells(29, Empregados_Coluna_Dados1)).Value = "0"
Sheets("EMPREGADOS").Range(Cells(37, Empregados_Coluna_Dados3), Cells(37, Empregados_Coluna_Dados1)).Value = "0"


'Calculo da difern�a em pontos do valor mais recente em rela��o ao valor do m�s anterior
Do Until Empregados_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Calculo da diferen�a em si: o valor da celula de analise � igual ao valor mais recente menos o valor do m�s anterior
   Sheets("EMPREGADOS").Cells(Empregados_Linha_An�lise, Empregados_Coluna_An�lise).Value = Sheets("EMPREGADOS").Cells(Empregados_Linha_Dados, Empregados_Coluna_Dados1).Value - Sheets("EMPREGADOS").Cells(Empregados_Linha_Dados, Empregados_Coluna_Dados2).Value
    'Vai para a pr�xima linha de dados e de an�lise
   Empregados_Linha_Dados = Empregados_Linha_Dados + 1
   Empregados_Linha_An�lise = Empregados_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Empregados_Linha_Dados = 9
Empregados_Coluna_Dados1 = Sheets("EMPREGADOS").Range("A9").End(xlToRight).Column
Empregados_Coluna_Dados3 = Empregados_Coluna_Dados1 - 12
Empregados_Linha_An�lise = 59
Empregados_Coluna_An�lise = 3

'C�lculo da diferen�a em pontos do valor mais recente em rela��o ao valor do mesmo m�s do ano anterior
Do Until Empregados_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Calculo da diferen�a em si: o valor da celula de analise � igual ao valor mais recente menos o valor do mesmo m�s do ano anterior
    Sheets("EMPREGADOS").Cells(Empregados_Linha_An�lise, Empregados_Coluna_An�lise).Value = Sheets("EMPREGADOS").Cells(Empregados_Linha_Dados, Empregados_Coluna_Dados1).Value - Sheets("EMPREGADOS").Cells(Empregados_Linha_Dados, Empregados_Coluna_Dados3).Value
    'Vai para a pr�xima linha de dados e de an�lise
    Empregados_Linha_Dados = Empregados_Linha_Dados + 1
    Empregados_Linha_An�lise = Empregados_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Empregados_Linha_Dados = 9
Empregados_Coluna_Dados1 = Sheets("EMPREGADOS").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Empregados_Linha_An�lise = 59
Empregados_Coluna_An�lise = 4

'C�lculo da diferen�a em pontos do valor mais recente em rela��o ao valor da m�dia hist�rica
Do Until Empregados_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a vari�vel media como a m�dia do intervalo entre a coluna com o dado mais recente e o primeiro
    media = Application.Average(Sheets("EMPREGADOS").Range(Cells(Empregados_Linha_Dados, Empregados_Coluna_Dados1), Cells(Empregados_Linha_Dados, 2)))
    'Calculo em si: o valor da celula de analise � igual ao valor mais recente menos o valor da m�dia
    Sheets("EMPREGADOS").Cells(Empregados_Linha_An�lise, Empregados_Coluna_An�lise).Value = Sheets("EMPREGADOS").Cells(Empregados_Linha_Dados, Empregados_Coluna_Dados1).Value - media
    'Vai para a pr�xima linha de dados e de an�lise
    Empregados_Linha_Dados = Empregados_Linha_Dados + 1
    Empregados_Linha_An�lise = Empregados_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Empregados_Linha_Dados = 9
Empregados_Coluna_Dados1 = Sheets("EMPREGADOS").Range("A9").End(xlToRight).Column
Empregados_Linha_An�lise = 59
Empregados_Coluna_An�lise = 5

'Ordena��o decrescente da s�rie hist�rica completa
Do Until Empregados_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posi��o = WorksheetFunction.Rank_Eq(Sheets("EMPREGADOS").Cells(Empregados_Linha_Dados, Empregados_Coluna_Dados1), Sheets("EMPREGADOS").Range(Cells(Empregados_Linha_Dados, Empregados_Coluna_Dados1), Cells(Empregados_Linha_Dados, 2)), 0)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("EMPREGADOS").Cells(Empregados_Linha_An�lise, Empregados_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Empregados_Linha_Dados = Empregados_Linha_Dados + 1
    Empregados_Linha_An�lise = Empregados_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Empregados_Linha_Dados = 9
Empregados_Coluna_Dados1 = Sheets("EMPREGADOS").Range("A9").End(xlToRight).Column
Empregados_Linha_An�lise = 59
Empregados_Coluna_An�lise = 6

'Ordena��o Crescente da s�rie hist�rica completa
Do Until Empregados_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posi��o = WorksheetFunction.Rank_Eq(Sheets("EMPREGADOS").Cells(Empregados_Linha_Dados, Empregados_Coluna_Dados1), Sheets("EMPREGADOS").Range(Cells(Empregados_Linha_Dados, Empregados_Coluna_Dados1), Cells(Empregados_Linha_Dados, 2)), 1)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("EMPREGADOS").Cells(Empregados_Linha_An�lise, Empregados_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Empregados_Linha_Dados = Empregados_Linha_Dados + 1
    Empregados_Linha_An�lise = Empregados_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior e define a vari�vel Coluna_DadosP que representa a primeira coluna de dados
Empregados_Coluna_Dados1 = Sheets("EMPREGADOS").Range("A9").End(xlToRight).Column
Empregados_Coluna_DadosP = 2

Do Until Empregados_Coluna_DadosP = Empregados_Coluna_Dados1 + 1 ' Faz at� a variavel Coluna_DadosP ser igual a variavel Coluna_Dados1 mais uma unidade
    'Confere se o m�s da coluna em quest�o � igual ao m�s do dado mais recente
    If Month(Sheets("EMPREGADOS").Cells(8, Empregados_Coluna_DadosP)) = Month(Sheets("EMPREGADOS").Cells(8, Empregados_Coluna_Dados1)) Then
    'Caso seja igual, copia a coluna com os dados mais abaixo, a partir da linha 110
        Sheets("EMPREGADOS").Range(Cells(9, Empregados_Coluna_DadosP), (Cells(54, Empregados_Coluna_DadosP))).Copy (Sheets("EMPREGADOS").Cells(110, Empregados_Coluna_DadosP))
    End If
    'Vai para a pr�xima coluna
    Empregados_Coluna_DadosP = Empregados_Coluna_DadosP + 1
'Repete a conferencia com a pr�xima coluna
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Empregados_Linha_Dados = 110
Empregados_Coluna_Dados1 = Sheets("EMPREGADOS").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Empregados_Linha_An�lise = 59
Empregados_Coluna_An�lise = 7

'Ordena��o decrescente da s�rie hist�rica dos meses do dado mais recente
Do Until Empregados_Linha_Dados = 156 'Faz o calculo at� a vari�vel Linha_Dados ser 156
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo m�s do mais recente
    posi��o = Application.WorksheetFunction.Rank_Eq(Sheets("EMPREGADOS").Cells(Empregados_Linha_Dados, Empregados_Coluna_Dados1), Sheets("EMPREGADOS").Range(Cells(Empregados_Linha_Dados, Empregados_Coluna_Dados1), Cells(Empregados_Linha_Dados, 2)))
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("EMPREGADOS").Cells(Empregados_Linha_An�lise, Empregados_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Empregados_Linha_Dados = Empregados_Linha_Dados + 1
    Empregados_Linha_An�lise = Empregados_Linha_An�lise + 1
'Repete a ordena��o com a pr�xima linha
Loop

'Atribui os valores originais das variaveis ap�s o loop anterior
Empregados_Linha_Dados = 110
Empregados_Coluna_Dados1 = Sheets("EMPREGADOS").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Empregados_Linha_An�lise = 59
Empregados_Coluna_An�lise = 8
'Ordena��o crescente da s�rie hist�rica dos meses do dado mais recente
Do Until Empregados_Linha_Dados = 156 'Faz o calculo at� a vari�vel Linha_Dados ser 156
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo m�s do mais recente
    posi��o = Application.WorksheetFunction.Rank_Eq(Sheets("EMPREGADOS").Cells(Empregados_Linha_Dados, Empregados_Coluna_Dados1), Sheets("EMPREGADOS").Range(Cells(Empregados_Linha_Dados, Empregados_Coluna_Dados1), Cells(Empregados_Linha_Dados, 2)), 1)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("EMPREGADOS").Cells(Empregados_Linha_An�lise, Empregados_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Empregados_Linha_Dados = Empregados_Linha_Dados + 1
    Empregados_Linha_An�lise = Empregados_Linha_An�lise + 1
'Repete a ordena��o com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Empregados_Linha_Dados = 9
Empregados_Coluna_Dados1 = Sheets("EMPREGADOS").Range("A9").End(xlToRight).Column
Empregados_Coluna_Dados2 = Empregados_Coluna_Dados1 - 1
Empregados_Linha_An�lise = 59
Empregados_Coluna_An�lise = 9

'Avalia��o se cruzou ou n�o a linha de 50 e o sentido
Do Until Empregados_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'se o dado mais recente for menor que 50 e o dado do m�s anterior for maior ou igual a 50 ent�o...
    If Cells(Empregados_Linha_Dados, Empregados_Coluna_Dados1) < 50 And Cells(Empregados_Linha_Dados, Empregados_Coluna_Dados2) >= 50 Then
    
    'a c�lula de an�lise recebe cruzou para baixo
    Cells(Empregados_Linha_An�lise, Empregados_Coluna_An�lise).Value = "Cruzou para baixo"
    'Caso n�o seja..
    Else
        'se o dado mais recente for maior ou igual a 50 e o dado do m~es anterior for menor ou igual a 50 ent�o...
        If Cells(Empregados_Linha_Dados, Empregados_Coluna_Dados1) >= 50 And Cells(Empregados_Linha_Dados, Empregados_Coluna_Dados2) <= 50 Then
        'a c�lula de an�lise recebe cruzou para cima
        Cells(Empregados_Linha_An�lise, Empregados_Coluna_An�lise).Value = "Cruzou para cima"
        'Caso n�o seja..
        Else
        'a c�lula de an�lise recebe n�o cruzou
        Cells(Empregados_Linha_An�lise, Empregados_Coluna_An�lise).Value = "N�o Cruzou"
        End If
    End If
    'Vai para a pr�xima linha de dados e de an�lise
    Empregados_Linha_Dados = Empregados_Linha_Dados + 1
    Empregados_Linha_An�lise = Empregados_Linha_An�lise + 1
'repete o processo com a nova linha
Loop

'Apaga as linhas com erros/dados faltantes/t�tulos e subt�tulos
Range(Cells(60, 2), Cells(60, 9)).ClearContents
Range(Cells(66, 2), Cells(66, 9)).ClearContents
Range(Cells(70, 2), Cells(70, 9)).ClearContents
Range(Cells(72, 2), Cells(73, 9)).Value = "-"
Range(Cells(75, 2), Cells(75, 9)).Value = "-"
Range(Cells(79, 2), Cells(79, 9)).Value = "-"
Range(Cells(87, 2), Cells(87, 9)).Value = "-"

'Inserindo valores nas celulas vazias para fugir de bugs
Range(Cells(10, Empregados_Coluna_Dados3), Cells(10, Empregados_Coluna_Dados1)).ClearContents
Range(Cells(16, Empregados_Coluna_Dados3), Cells(16, Empregados_Coluna_Dados1)).ClearContents
Range(Cells(20, Empregados_Coluna_Dados3), Cells(20, Empregados_Coluna_Dados1)).ClearContents
Range(Cells(22, Empregados_Coluna_Dados3), Cells(23, Empregados_Coluna_Dados1)).Value = "-"
Range(Cells(25, Empregados_Coluna_Dados3), Cells(25, Empregados_Coluna_Dados1)).Value = "-"
Range(Cells(29, Empregados_Coluna_Dados3), Cells(29, Empregados_Coluna_Dados1)).Value = "-"
Range(Cells(37, Empregados_Coluna_Dados3), Cells(37, Empregados_Coluna_Dados1)).Value = "-"

Range("E59:H104").NumberFormat = "0"




'**************************************************            An�lise_UCI                        ********************************************************



Dim UCI_Linha_Dados As Integer 'Define a linha que cont�m o dado a ser usado
Dim UCI_Coluna_Dados1 As Integer ' Define a coluna com o dado mais recente
Dim UCI_Coluna_Dados2 As Integer ' Define a coluna com o dado do m�s anterior
Dim UCI_Coluna_Dados3 As Integer ' Defie a coluna com o dado do mesmo m�s do ano anterior
Dim UCI_Linha_An�lise As Integer ' Define a linha que ser� feita a an�lise
Dim UCI_Coluna_An�lise As Integer 'Define a coluna que ser� feita a an�lise

Sheets("UCI (%)").Select


'Copia os t�tulos das categorias e cola onde ser� formada a tabela de an�lise
Sheets("UCI (%)").Range(Cells(9, 1), Cells(54, 9)).Copy (Sheets("UCI (%)").Cells(59, 1))
'Limpa os n�meros que foram colados mas mant�m a formata��o
Sheets("UCI (%)").Range(Cells(59, 2), Cells(105, 9)).ClearContents



'Nomeia as colunas de acordo com o dado que ser� calculado nelas
Sheets("UCI (%)").Cells(58, 2).Value = "Diferen�a para o m�s anterior"
Sheets("UCI (%)").Cells(58, 3).Value = "Diferen�a para ao mesmo m�s do ano anterior"
Sheets("UCI (%)").Cells(58, 4).Value = "Diferen�a para a m�dia hist�rica"
Sheets("UCI (%)").Cells(58, 5).Value = "Posi��o Decrescente (Maior valor 1�, menor valor �ltimo)"
Sheets("UCI (%)").Cells(58, 6).Value = "Posi��o Crescente  (Menor valor 1�, maior valor �ltimo)"
Sheets("UCI (%)").Cells(58, 7).Value = "Posi��o Crescente - Mesmo m�s  (Menor valor 1�, maior valor �ltimo)"
Sheets("UCI (%)").Cells(58, 8).Value = "Posi��o Decrescente -Mesmo m�s  (Maior valor 1�, menor valor �ltimo)"
Sheets("UCI (%)").Cells(58, 9).Value = "cruzou a linha de 50?"

'Atribui valores as variaveis definidas acima
UCI_Linha_Dados = 9 'Define o n�mero da primeira linha de dados
UCI_Coluna_Dados1 = Sheets("UCI (%)").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
UCI_Coluna_Dados2 = UCI_Coluna_Dados1 - 1 'Define o n�mero da coluna do m�s anterior
UCI_Coluna_Dados3 = UCI_Coluna_Dados1 - 12
UCI_Linha_An�lise = 59 'Define a primeira linhas de an�lises
UCI_Coluna_An�lise = 2 'Define a coluna de an�lises

'Inserindo valores nas celulas vazias para fugir de bugs
Sheets("UCI (%)").Range(Cells(10, UCI_Coluna_Dados3), Cells(10, UCI_Coluna_Dados1)).Value = "0"
Sheets("UCI (%)").Range(Cells(16, UCI_Coluna_Dados3), Cells(16, UCI_Coluna_Dados1)).Value = "0"
Sheets("UCI (%)").Range(Cells(20, UCI_Coluna_Dados3), Cells(20, UCI_Coluna_Dados1)).Value = "0"
Sheets("UCI (%)").Range(Cells(22, UCI_Coluna_Dados3), Cells(23, UCI_Coluna_Dados1)).Value = "0"
Sheets("UCI (%)").Range(Cells(25, UCI_Coluna_Dados3), Cells(25, UCI_Coluna_Dados1)).Value = "0"
Sheets("UCI (%)").Range(Cells(29, UCI_Coluna_Dados3), Cells(29, UCI_Coluna_Dados1)).Value = "0"
Sheets("UCI (%)").Range(Cells(37, UCI_Coluna_Dados3), Cells(37, UCI_Coluna_Dados1)).Value = "0"


'Calculo da difern�a em pontos do valor mais recente em rela��o ao valor do m�s anterior
Do Until UCI_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Calculo da diferen�a em si: o valor da celula de analise � igual ao valor mais recente menos o valor do m�s anterior
   Sheets("UCI (%)").Cells(UCI_Linha_An�lise, UCI_Coluna_An�lise).Value = Sheets("UCI (%)").Cells(UCI_Linha_Dados, UCI_Coluna_Dados1).Value - Sheets("UCI (%)").Cells(UCI_Linha_Dados, UCI_Coluna_Dados2).Value
    'Vai para a pr�xima linha de dados e de an�lise
   UCI_Linha_Dados = UCI_Linha_Dados + 1
   UCI_Linha_An�lise = UCI_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
UCI_Linha_Dados = 9
UCI_Coluna_Dados1 = Sheets("UCI (%)").Range("A9").End(xlToRight).Column
UCI_Coluna_Dados3 = UCI_Coluna_Dados1 - 12
UCI_Linha_An�lise = 59
UCI_Coluna_An�lise = 3

'C�lculo da diferen�a em pontos do valor mais recente em rela��o ao valor do mesmo m�s do ano anterior
Do Until UCI_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Calculo da diferen�a em si: o valor da celula de analise � igual ao valor mais recente menos o valor do mesmo m�s do ano anterior
    Sheets("UCI (%)").Cells(UCI_Linha_An�lise, UCI_Coluna_An�lise).Value = Sheets("UCI (%)").Cells(UCI_Linha_Dados, UCI_Coluna_Dados1).Value - Sheets("UCI (%)").Cells(UCI_Linha_Dados, UCI_Coluna_Dados3).Value
    'Vai para a pr�xima linha de dados e de an�lise
    UCI_Linha_Dados = UCI_Linha_Dados + 1
    UCI_Linha_An�lise = UCI_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
UCI_Linha_Dados = 9
UCI_Coluna_Dados1 = Sheets("UCI (%)").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
UCI_Linha_An�lise = 59
UCI_Coluna_An�lise = 4

'C�lculo da diferen�a em pontos do valor mais recente em rela��o ao valor da m�dia hist�rica
Do Until UCI_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a vari�vel media como a m�dia do intervalo entre a coluna com o dado mais recente e o primeiro
    media = Application.Average(Sheets("UCI (%)").Range(Cells(UCI_Linha_Dados, UCI_Coluna_Dados1), Cells(UCI_Linha_Dados, 2)))
    'Calculo em si: o valor da celula de analise � igual ao valor mais recente menos o valor da m�dia
    Sheets("UCI (%)").Cells(UCI_Linha_An�lise, UCI_Coluna_An�lise).Value = Sheets("UCI (%)").Cells(UCI_Linha_Dados, UCI_Coluna_Dados1).Value - media
    'Vai para a pr�xima linha de dados e de an�lise
    UCI_Linha_Dados = UCI_Linha_Dados + 1
    UCI_Linha_An�lise = UCI_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
UCI_Linha_Dados = 9
UCI_Coluna_Dados1 = Sheets("UCI (%)").Range("A9").End(xlToRight).Column
UCI_Linha_An�lise = 59
UCI_Coluna_An�lise = 5

'Ordena��o decrescente da s�rie hist�rica completa
Do Until UCI_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posi��o = WorksheetFunction.Rank_Eq(Sheets("UCI (%)").Cells(UCI_Linha_Dados, UCI_Coluna_Dados1), Sheets("UCI (%)").Range(Cells(UCI_Linha_Dados, UCI_Coluna_Dados1), Cells(UCI_Linha_Dados, 2)), 0)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("UCI (%)").Cells(UCI_Linha_An�lise, UCI_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    UCI_Linha_Dados = UCI_Linha_Dados + 1
    UCI_Linha_An�lise = UCI_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
UCI_Linha_Dados = 9
UCI_Coluna_Dados1 = Sheets("UCI (%)").Range("A9").End(xlToRight).Column
UCI_Linha_An�lise = 59
UCI_Coluna_An�lise = 6

'Ordena��o Crescente da s�rie hist�rica completa
Do Until UCI_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posi��o = WorksheetFunction.Rank_Eq(Sheets("UCI (%)").Cells(UCI_Linha_Dados, UCI_Coluna_Dados1), Sheets("UCI (%)").Range(Cells(UCI_Linha_Dados, UCI_Coluna_Dados1), Cells(UCI_Linha_Dados, 2)), 1)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("UCI (%)").Cells(UCI_Linha_An�lise, UCI_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    UCI_Linha_Dados = UCI_Linha_Dados + 1
    UCI_Linha_An�lise = UCI_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior e define a vari�vel Coluna_DadosP que representa a primeira coluna de dados
UCI_Coluna_Dados1 = Sheets("UCI (%)").Range("A9").End(xlToRight).Column
UCI_Coluna_DadosP = 2

Do Until UCI_Coluna_DadosP = UCI_Coluna_Dados1 + 1 ' Faz at� a variavel Coluna_DadosP ser igual a variavel Coluna_Dados1 mais uma unidade
    'Confere se o m�s da coluna em quest�o � igual ao m�s do dado mais recente
    If Month(Sheets("UCI (%)").Cells(8, UCI_Coluna_DadosP)) = Month(Sheets("UCI (%)").Cells(8, UCI_Coluna_Dados1)) Then
    'Caso seja igual, copia a coluna com os dados mais abaixo, a partir da linha 110
        Sheets("UCI (%)").Range(Cells(9, UCI_Coluna_DadosP), (Cells(54, UCI_Coluna_DadosP))).Copy (Sheets("UCI (%)").Cells(110, UCI_Coluna_DadosP))
    End If
    'Vai para a pr�xima coluna
    UCI_Coluna_DadosP = UCI_Coluna_DadosP + 1
'Repete a conferencia com a pr�xima coluna
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
UCI_Linha_Dados = 110
UCI_Coluna_Dados1 = Sheets("UCI (%)").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
UCI_Linha_An�lise = 59
UCI_Coluna_An�lise = 7

'Ordena��o decrescente da s�rie hist�rica dos meses do dado mais recente
Do Until UCI_Linha_Dados = 156 'Faz o calculo at� a vari�vel Linha_Dados ser 156
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo m�s do mais recente
    posi��o = Application.WorksheetFunction.Rank_Eq(Sheets("UCI (%)").Cells(UCI_Linha_Dados, UCI_Coluna_Dados1), Sheets("UCI (%)").Range(Cells(UCI_Linha_Dados, UCI_Coluna_Dados1), Cells(UCI_Linha_Dados, 2)))
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("UCI (%)").Cells(UCI_Linha_An�lise, UCI_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    UCI_Linha_Dados = UCI_Linha_Dados + 1
    UCI_Linha_An�lise = UCI_Linha_An�lise + 1
'Repete a ordena��o com a pr�xima linha
Loop

'Atribui os valores originais das variaveis ap�s o loop anterior
UCI_Linha_Dados = 110
UCI_Coluna_Dados1 = Sheets("UCI (%)").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
UCI_Linha_An�lise = 59
UCI_Coluna_An�lise = 8
'Ordena��o crescente da s�rie hist�rica dos meses do dado mais recente
Do Until UCI_Linha_Dados = 156 'Faz o calculo at� a vari�vel Linha_Dados ser 156
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo m�s do mais recente
    posi��o = Application.WorksheetFunction.Rank_Eq(Sheets("UCI (%)").Cells(UCI_Linha_Dados, UCI_Coluna_Dados1), Sheets("UCI (%)").Range(Cells(UCI_Linha_Dados, UCI_Coluna_Dados1), Cells(UCI_Linha_Dados, 2)), 1)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("UCI (%)").Cells(UCI_Linha_An�lise, UCI_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    UCI_Linha_Dados = UCI_Linha_Dados + 1
    UCI_Linha_An�lise = UCI_Linha_An�lise + 1
'Repete a ordena��o com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
UCI_Linha_Dados = 9
UCI_Coluna_Dados1 = Sheets("UCI (%)").Range("A9").End(xlToRight).Column
UCI_Coluna_Dados2 = UCI_Coluna_Dados1 - 1
UCI_Linha_An�lise = 59
UCI_Coluna_An�lise = 9

'Avalia��o se cruzou ou n�o a linha de 50 e o sentido
Do Until UCI_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'se o dado mais recente for menor que 50 e o dado do m�s anterior for maior ou igual a 50 ent�o...
    If Cells(UCI_Linha_Dados, UCI_Coluna_Dados1) < 50 And Cells(UCI_Linha_Dados, UCI_Coluna_Dados2) >= 50 Then
    
    'a c�lula de an�lise recebe cruzou para baixo
    Cells(UCI_Linha_An�lise, UCI_Coluna_An�lise).Value = "Cruzou para baixo"
    'Caso n�o seja..
    Else
        'se o dado mais recente for maior ou igual a 50 e o dado do m~es anterior for menor ou igual a 50 ent�o...
        If Cells(UCI_Linha_Dados, UCI_Coluna_Dados1) >= 50 And Cells(UCI_Linha_Dados, UCI_Coluna_Dados2) <= 50 Then
        'a c�lula de an�lise recebe cruzou para cima
        Cells(UCI_Linha_An�lise, UCI_Coluna_An�lise).Value = "Cruzou para cima"
        'Caso n�o seja..
        Else
        'a c�lula de an�lise recebe n�o cruzou
        Cells(UCI_Linha_An�lise, UCI_Coluna_An�lise).Value = "N�o Cruzou"
        End If
    End If
    'Vai para a pr�xima linha de dados e de an�lise
    UCI_Linha_Dados = UCI_Linha_Dados + 1
    UCI_Linha_An�lise = UCI_Linha_An�lise + 1
'repete o processo com a nova linha
Loop

'Apaga as linhas com erros/dados faltantes/t�tulos e subt�tulos
Range(Cells(60, 2), Cells(60, 9)).ClearContents
Range(Cells(66, 2), Cells(66, 9)).ClearContents
Range(Cells(70, 2), Cells(70, 9)).ClearContents
Range(Cells(72, 2), Cells(73, 9)).Value = "-"
Range(Cells(75, 2), Cells(75, 9)).Value = "-"
Range(Cells(79, 2), Cells(79, 9)).Value = "-"
Range(Cells(87, 2), Cells(87, 9)).Value = "-"

'Inserindo valores nas celulas vazias para fugir de bugs
Range(Cells(10, UCI_Coluna_Dados3), Cells(10, UCI_Coluna_Dados1)).ClearContents
Range(Cells(16, UCI_Coluna_Dados3), Cells(16, UCI_Coluna_Dados1)).ClearContents
Range(Cells(20, UCI_Coluna_Dados3), Cells(20, UCI_Coluna_Dados1)).ClearContents
Range(Cells(22, UCI_Coluna_Dados3), Cells(23, UCI_Coluna_Dados1)).Value = "-"
Range(Cells(25, UCI_Coluna_Dados3), Cells(25, UCI_Coluna_Dados1)).Value = "-"
Range(Cells(29, UCI_Coluna_Dados3), Cells(29, UCI_Coluna_Dados1)).Value = "-"
Range(Cells(37, UCI_Coluna_Dados3), Cells(37, UCI_Coluna_Dados1)).Value = "-"

Range("E59:H104").NumberFormat = "0"





'*************************************       An�lise_UCI_Efetiva_Usual                  **************************************************************



Dim UCI_Efetiva_Usual_Linha_Dados As Integer 'Define a linha que cont�m o dado a ser usado
Dim UCI_Efetiva_Usual_Coluna_Dados1 As Integer ' Define a coluna com o dado mais recente
Dim UCI_Efetiva_Usual_Coluna_Dados2 As Integer ' Define a coluna com o dado do m�s anterior
Dim UCI_Efetiva_Usual_Coluna_Dados3 As Integer ' Defie a coluna com o dado do mesmo m�s do ano anterior
Dim UCI_Efetiva_Usual_Linha_An�lise As Integer ' Define a linha que ser� feita a an�lise
Dim UCI_Efetiva_Usual_Coluna_An�lise As Integer 'Define a coluna que ser� feita a an�lise

Sheets("UCI (efetiva-usual)").Select


'Copia os t�tulos das categorias e cola onde ser� formada a tabela de an�lise
Sheets("UCI (efetiva-usual)").Range(Cells(9, 1), Cells(54, 9)).Copy (Sheets("UCI (efetiva-usual)").Cells(59, 1))
'Limpa os n�meros que foram colados mas mant�m a formata��o
Sheets("UCI (efetiva-usual)").Range(Cells(59, 2), Cells(105, 9)).ClearContents



'Nomeia as colunas de acordo com o dado que ser� calculado nelas
Sheets("UCI (efetiva-usual)").Cells(58, 2).Value = "Diferen�a para o m�s anterior"
Sheets("UCI (efetiva-usual)").Cells(58, 3).Value = "Diferen�a para ao mesmo m�s do ano anterior"
Sheets("UCI (efetiva-usual)").Cells(58, 4).Value = "Diferen�a para a m�dia hist�rica"
Sheets("UCI (efetiva-usual)").Cells(58, 5).Value = "Posi��o Decrescente (Maior valor 1�, menor valor �ltimo)"
Sheets("UCI (efetiva-usual)").Cells(58, 6).Value = "Posi��o Crescente  (Menor valor 1�, maior valor �ltimo)"
Sheets("UCI (efetiva-usual)").Cells(58, 7).Value = "Posi��o Crescente - Mesmo m�s  (Menor valor 1�, maior valor �ltimo)"
Sheets("UCI (efetiva-usual)").Cells(58, 8).Value = "Posi��o Decrescente -Mesmo m�s  (Maior valor 1�, menor valor �ltimo)"
Sheets("UCI (efetiva-usual)").Cells(58, 9).Value = "cruzou a linha de 50?"

'Atribui valores as variaveis definidas acima
UCI_Efetiva_Usual_Linha_Dados = 9 'Define o n�mero da primeira linha de dados
UCI_Efetiva_Usual_Coluna_Dados1 = Sheets("UCI (efetiva-usual)").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
UCI_Efetiva_Usual_Coluna_Dados2 = UCI_Efetiva_Usual_Coluna_Dados1 - 1 'Define o n�mero da coluna do m�s anterior
UCI_Efetiva_Usual_Coluna_Dados3 = UCI_Efetiva_Usual_Coluna_Dados1 - 12
UCI_Efetiva_Usual_Linha_An�lise = 59 'Define a primeira linhas de an�lises
UCI_Efetiva_Usual_Coluna_An�lise = 2 'Define a coluna de an�lises

'Inserindo valores nas celulas vazias para fugir de bugs
Sheets("UCI (efetiva-usual)").Range(Cells(10, UCI_Efetiva_Usual_Coluna_Dados3), Cells(10, UCI_Efetiva_Usual_Coluna_Dados1)).Value = "0"
Sheets("UCI (efetiva-usual)").Range(Cells(16, UCI_Efetiva_Usual_Coluna_Dados3), Cells(16, UCI_Efetiva_Usual_Coluna_Dados1)).Value = "0"
Sheets("UCI (efetiva-usual)").Range(Cells(20, UCI_Efetiva_Usual_Coluna_Dados3), Cells(20, UCI_Efetiva_Usual_Coluna_Dados1)).Value = "0"
Sheets("UCI (efetiva-usual)").Range(Cells(22, UCI_Efetiva_Usual_Coluna_Dados3), Cells(23, UCI_Efetiva_Usual_Coluna_Dados1)).Value = "0"
Sheets("UCI (efetiva-usual)").Range(Cells(25, UCI_Efetiva_Usual_Coluna_Dados3), Cells(25, UCI_Efetiva_Usual_Coluna_Dados1)).Value = "0"
Sheets("UCI (efetiva-usual)").Range(Cells(29, UCI_Efetiva_Usual_Coluna_Dados3), Cells(29, UCI_Efetiva_Usual_Coluna_Dados1)).Value = "0"
Sheets("UCI (efetiva-usual)").Range(Cells(37, UCI_Efetiva_Usual_Coluna_Dados3), Cells(37, UCI_Efetiva_Usual_Coluna_Dados1)).Value = "0"


'Calculo da difern�a em pontos do valor mais recente em rela��o ao valor do m�s anterior
Do Until UCI_Efetiva_Usual_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Calculo da diferen�a em si: o valor da celula de analise � igual ao valor mais recente menos o valor do m�s anterior
   Sheets("UCI (efetiva-usual)").Cells(UCI_Efetiva_Usual_Linha_An�lise, UCI_Efetiva_Usual_Coluna_An�lise).Value = Sheets("UCI (efetiva-usual)").Cells(UCI_Efetiva_Usual_Linha_Dados, UCI_Efetiva_Usual_Coluna_Dados1).Value - Sheets("UCI (efetiva-usual)").Cells(UCI_Efetiva_Usual_Linha_Dados, UCI_Efetiva_Usual_Coluna_Dados2).Value
    'Vai para a pr�xima linha de dados e de an�lise
   UCI_Efetiva_Usual_Linha_Dados = UCI_Efetiva_Usual_Linha_Dados + 1
   UCI_Efetiva_Usual_Linha_An�lise = UCI_Efetiva_Usual_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
UCI_Efetiva_Usual_Linha_Dados = 9
UCI_Efetiva_Usual_Coluna_Dados1 = Sheets("UCI (efetiva-usual)").Range("A9").End(xlToRight).Column
UCI_Efetiva_Usual_Coluna_Dados3 = UCI_Efetiva_Usual_Coluna_Dados1 - 12
UCI_Efetiva_Usual_Linha_An�lise = 59
UCI_Efetiva_Usual_Coluna_An�lise = 3

'C�lculo da diferen�a em pontos do valor mais recente em rela��o ao valor do mesmo m�s do ano anterior
Do Until UCI_Efetiva_Usual_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Calculo da diferen�a em si: o valor da celula de analise � igual ao valor mais recente menos o valor do mesmo m�s do ano anterior
    Sheets("UCI (efetiva-usual)").Cells(UCI_Efetiva_Usual_Linha_An�lise, UCI_Efetiva_Usual_Coluna_An�lise).Value = Sheets("UCI (efetiva-usual)").Cells(UCI_Efetiva_Usual_Linha_Dados, UCI_Efetiva_Usual_Coluna_Dados1).Value - Sheets("UCI (efetiva-usual)").Cells(UCI_Efetiva_Usual_Linha_Dados, UCI_Efetiva_Usual_Coluna_Dados3).Value
    'Vai para a pr�xima linha de dados e de an�lise
    UCI_Efetiva_Usual_Linha_Dados = UCI_Efetiva_Usual_Linha_Dados + 1
    UCI_Efetiva_Usual_Linha_An�lise = UCI_Efetiva_Usual_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
UCI_Efetiva_Usual_Linha_Dados = 9
UCI_Efetiva_Usual_Coluna_Dados1 = Sheets("UCI (efetiva-usual)").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
UCI_Efetiva_Usual_Linha_An�lise = 59
UCI_Efetiva_Usual_Coluna_An�lise = 4

'C�lculo da diferen�a em pontos do valor mais recente em rela��o ao valor da m�dia hist�rica
Do Until UCI_Efetiva_Usual_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a vari�vel media como a m�dia do intervalo entre a coluna com o dado mais recente e o primeiro
    media = Application.Average(Sheets("UCI (efetiva-usual)").Range(Cells(UCI_Efetiva_Usual_Linha_Dados, UCI_Efetiva_Usual_Coluna_Dados1), Cells(UCI_Efetiva_Usual_Linha_Dados, 2)))
    'Calculo em si: o valor da celula de analise � igual ao valor mais recente menos o valor da m�dia
    Sheets("UCI (efetiva-usual)").Cells(UCI_Efetiva_Usual_Linha_An�lise, UCI_Efetiva_Usual_Coluna_An�lise).Value = Sheets("UCI (efetiva-usual)").Cells(UCI_Efetiva_Usual_Linha_Dados, UCI_Efetiva_Usual_Coluna_Dados1).Value - media
    'Vai para a pr�xima linha de dados e de an�lise
    UCI_Efetiva_Usual_Linha_Dados = UCI_Efetiva_Usual_Linha_Dados + 1
    UCI_Efetiva_Usual_Linha_An�lise = UCI_Efetiva_Usual_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
UCI_Efetiva_Usual_Linha_Dados = 9
UCI_Efetiva_Usual_Coluna_Dados1 = Sheets("UCI (efetiva-usual)").Range("A9").End(xlToRight).Column
UCI_Efetiva_Usual_Linha_An�lise = 59
UCI_Efetiva_Usual_Coluna_An�lise = 5

'Ordena��o decrescente da s�rie hist�rica completa
Do Until UCI_Efetiva_Usual_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posi��o = WorksheetFunction.Rank_Eq(Sheets("UCI (efetiva-usual)").Cells(UCI_Efetiva_Usual_Linha_Dados, UCI_Efetiva_Usual_Coluna_Dados1), Sheets("UCI (efetiva-usual)").Range(Cells(UCI_Efetiva_Usual_Linha_Dados, UCI_Efetiva_Usual_Coluna_Dados1), Cells(UCI_Efetiva_Usual_Linha_Dados, 2)), 0)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("UCI (efetiva-usual)").Cells(UCI_Efetiva_Usual_Linha_An�lise, UCI_Efetiva_Usual_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    UCI_Efetiva_Usual_Linha_Dados = UCI_Efetiva_Usual_Linha_Dados + 1
    UCI_Efetiva_Usual_Linha_An�lise = UCI_Efetiva_Usual_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
UCI_Efetiva_Usual_Linha_Dados = 9
UCI_Efetiva_Usual_Coluna_Dados1 = Sheets("UCI (efetiva-usual)").Range("A9").End(xlToRight).Column
UCI_Efetiva_Usual_Linha_An�lise = 59
UCI_Efetiva_Usual_Coluna_An�lise = 6

'Ordena��o Crescente da s�rie hist�rica completa
Do Until UCI_Efetiva_Usual_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posi��o = WorksheetFunction.Rank_Eq(Sheets("UCI (efetiva-usual)").Cells(UCI_Efetiva_Usual_Linha_Dados, UCI_Efetiva_Usual_Coluna_Dados1), Sheets("UCI (efetiva-usual)").Range(Cells(UCI_Efetiva_Usual_Linha_Dados, UCI_Efetiva_Usual_Coluna_Dados1), Cells(UCI_Efetiva_Usual_Linha_Dados, 2)), 1)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("UCI (efetiva-usual)").Cells(UCI_Efetiva_Usual_Linha_An�lise, UCI_Efetiva_Usual_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    UCI_Efetiva_Usual_Linha_Dados = UCI_Efetiva_Usual_Linha_Dados + 1
    UCI_Efetiva_Usual_Linha_An�lise = UCI_Efetiva_Usual_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior e define a vari�vel Coluna_DadosP que representa a primeira coluna de dados
UCI_Efetiva_Usual_Coluna_Dados1 = Sheets("UCI (efetiva-usual)").Range("A9").End(xlToRight).Column
UCI_Efetiva_Usual_Coluna_DadosP = 2

Do Until UCI_Efetiva_Usual_Coluna_DadosP = UCI_Efetiva_Usual_Coluna_Dados1 + 1 ' Faz at� a variavel Coluna_DadosP ser igual a variavel Coluna_Dados1 mais uma unidade
    'Confere se o m�s da coluna em quest�o � igual ao m�s do dado mais recente
    If Month(Sheets("UCI (efetiva-usual)").Cells(8, UCI_Efetiva_Usual_Coluna_DadosP)) = Month(Sheets("UCI (efetiva-usual)").Cells(8, UCI_Efetiva_Usual_Coluna_Dados1)) Then
    'Caso seja igual, copia a coluna com os dados mais abaixo, a partir da linha 110
        Sheets("UCI (efetiva-usual)").Range(Cells(9, UCI_Efetiva_Usual_Coluna_DadosP), (Cells(54, UCI_Efetiva_Usual_Coluna_DadosP))).Copy (Sheets("UCI (efetiva-usual)").Cells(110, UCI_Efetiva_Usual_Coluna_DadosP))
    End If
    'Vai para a pr�xima coluna
    UCI_Efetiva_Usual_Coluna_DadosP = UCI_Efetiva_Usual_Coluna_DadosP + 1
'Repete a conferencia com a pr�xima coluna
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
UCI_Efetiva_Usual_Linha_Dados = 110
UCI_Efetiva_Usual_Coluna_Dados1 = Sheets("UCI (efetiva-usual)").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
UCI_Efetiva_Usual_Linha_An�lise = 59
UCI_Efetiva_Usual_Coluna_An�lise = 7

'Ordena��o decrescente da s�rie hist�rica dos meses do dado mais recente
Do Until UCI_Efetiva_Usual_Linha_Dados = 156 'Faz o calculo at� a vari�vel Linha_Dados ser 156
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo m�s do mais recente
    posi��o = Application.WorksheetFunction.Rank_Eq(Sheets("UCI (efetiva-usual)").Cells(UCI_Efetiva_Usual_Linha_Dados, UCI_Efetiva_Usual_Coluna_Dados1), Sheets("UCI (efetiva-usual)").Range(Cells(UCI_Efetiva_Usual_Linha_Dados, UCI_Efetiva_Usual_Coluna_Dados1), Cells(UCI_Efetiva_Usual_Linha_Dados, 2)))
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("UCI (efetiva-usual)").Cells(UCI_Efetiva_Usual_Linha_An�lise, UCI_Efetiva_Usual_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    UCI_Efetiva_Usual_Linha_Dados = UCI_Efetiva_Usual_Linha_Dados + 1
    UCI_Efetiva_Usual_Linha_An�lise = UCI_Efetiva_Usual_Linha_An�lise + 1
'Repete a ordena��o com a pr�xima linha
Loop

'Atribui os valores originais das variaveis ap�s o loop anterior
UCI_Efetiva_Usual_Linha_Dados = 110
UCI_Efetiva_Usual_Coluna_Dados1 = Sheets("UCI (efetiva-usual)").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
UCI_Efetiva_Usual_Linha_An�lise = 59
UCI_Efetiva_Usual_Coluna_An�lise = 8
'Ordena��o crescente da s�rie hist�rica dos meses do dado mais recente
Do Until UCI_Efetiva_Usual_Linha_Dados = 156 'Faz o calculo at� a vari�vel Linha_Dados ser 156
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo m�s do mais recente
    posi��o = Application.WorksheetFunction.Rank_Eq(Sheets("UCI (efetiva-usual)").Cells(UCI_Efetiva_Usual_Linha_Dados, UCI_Efetiva_Usual_Coluna_Dados1), Sheets("UCI (efetiva-usual)").Range(Cells(UCI_Efetiva_Usual_Linha_Dados, UCI_Efetiva_Usual_Coluna_Dados1), Cells(UCI_Efetiva_Usual_Linha_Dados, 2)), 1)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("UCI (efetiva-usual)").Cells(UCI_Efetiva_Usual_Linha_An�lise, UCI_Efetiva_Usual_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    UCI_Efetiva_Usual_Linha_Dados = UCI_Efetiva_Usual_Linha_Dados + 1
    UCI_Efetiva_Usual_Linha_An�lise = UCI_Efetiva_Usual_Linha_An�lise + 1
'Repete a ordena��o com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
UCI_Efetiva_Usual_Linha_Dados = 9
UCI_Efetiva_Usual_Coluna_Dados1 = Sheets("UCI (efetiva-usual)").Range("A9").End(xlToRight).Column
UCI_Efetiva_Usual_Coluna_Dados2 = UCI_Efetiva_Usual_Coluna_Dados1 - 1
UCI_Efetiva_Usual_Linha_An�lise = 59
UCI_Efetiva_Usual_Coluna_An�lise = 9

'Avalia��o se cruzou ou n�o a linha de 50 e o sentido
Do Until UCI_Efetiva_Usual_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'se o dado mais recente for menor que 50 e o dado do m�s anterior for maior ou igual a 50 ent�o...
    If Cells(UCI_Efetiva_Usual_Linha_Dados, UCI_Efetiva_Usual_Coluna_Dados1) < 50 And Cells(UCI_Efetiva_Usual_Linha_Dados, UCI_Efetiva_Usual_Coluna_Dados2) >= 50 Then
    
    'a c�lula de an�lise recebe cruzou para baixo
    Cells(UCI_Efetiva_Usual_Linha_An�lise, UCI_Efetiva_Usual_Coluna_An�lise).Value = "Cruzou para baixo"
    'Caso n�o seja..
    Else
        'se o dado mais recente for maior ou igual a 50 e o dado do m~es anterior for menor ou igual a 50 ent�o...
        If Cells(UCI_Efetiva_Usual_Linha_Dados, UCI_Efetiva_Usual_Coluna_Dados1) >= 50 And Cells(UCI_Efetiva_Usual_Linha_Dados, UCI_Efetiva_Usual_Coluna_Dados2) <= 50 Then
        'a c�lula de an�lise recebe cruzou para cima
        Cells(UCI_Efetiva_Usual_Linha_An�lise, UCI_Efetiva_Usual_Coluna_An�lise).Value = "Cruzou para cima"
        'Caso n�o seja..
        Else
        'a c�lula de an�lise recebe n�o cruzou
        Cells(UCI_Efetiva_Usual_Linha_An�lise, UCI_Efetiva_Usual_Coluna_An�lise).Value = "N�o Cruzou"
        End If
    End If
    'Vai para a pr�xima linha de dados e de an�lise
    UCI_Efetiva_Usual_Linha_Dados = UCI_Efetiva_Usual_Linha_Dados + 1
    UCI_Efetiva_Usual_Linha_An�lise = UCI_Efetiva_Usual_Linha_An�lise + 1
'repete o processo com a nova linha
Loop

'Apaga as linhas com erros/dados faltantes/t�tulos e subt�tulos
Range(Cells(60, 2), Cells(60, 9)).ClearContents
Range(Cells(66, 2), Cells(66, 9)).ClearContents
Range(Cells(70, 2), Cells(70, 9)).ClearContents
Range(Cells(72, 2), Cells(73, 9)).Value = "-"
Range(Cells(75, 2), Cells(75, 9)).Value = "-"
Range(Cells(79, 2), Cells(79, 9)).Value = "-"
Range(Cells(87, 2), Cells(87, 9)).Value = "-"

'Inserindo valores nas celulas vazias para fugir de bugs
Range(Cells(10, UCI_Efetiva_Usual_Coluna_Dados3), Cells(10, UCI_Efetiva_Usual_Coluna_Dados1)).ClearContents
Range(Cells(16, UCI_Efetiva_Usual_Coluna_Dados3), Cells(16, UCI_Efetiva_Usual_Coluna_Dados1)).ClearContents
Range(Cells(20, UCI_Efetiva_Usual_Coluna_Dados3), Cells(20, UCI_Efetiva_Usual_Coluna_Dados1)).ClearContents
Range(Cells(22, UCI_Efetiva_Usual_Coluna_Dados3), Cells(23, UCI_Efetiva_Usual_Coluna_Dados1)).Value = "-"
Range(Cells(25, UCI_Efetiva_Usual_Coluna_Dados3), Cells(25, UCI_Efetiva_Usual_Coluna_Dados1)).Value = "-"
Range(Cells(29, UCI_Efetiva_Usual_Coluna_Dados3), Cells(29, UCI_Efetiva_Usual_Coluna_Dados1)).Value = "-"
Range(Cells(37, UCI_Efetiva_Usual_Coluna_Dados3), Cells(37, UCI_Efetiva_Usual_Coluna_Dados1)).Value = "-"

Range("E59:H104").NumberFormat = "0"




'*************************************          ESTOQUES_evolu��o            *********************************************************************




Dim Estoques_Evolu��o_Linha_Dados As Integer 'Define a linha que cont�m o dado a ser usado
Dim Estoques_Evolu��o_Coluna_Dados1 As Integer ' Define a coluna com o dado mais recente
Dim Estoques_Evolu��o_Coluna_Dados2 As Integer ' Define a coluna com o dado do m�s anterior
Dim Estoques_Evolu��o_Coluna_Dados3 As Integer ' Defie a coluna com o dado do mesmo m�s do ano anterior
Dim Estoques_Evolu��o_Linha_An�lise As Integer ' Define a linha que ser� feita a an�lise
Dim Estoques_Evolu��o_Coluna_An�lise As Integer 'Define a coluna que ser� feita a an�lise

Sheets("ESTOQUES (evolu��o)").Select

'Copia os t�tulos das categorias e cola onde ser� formada a tabela de an�lise
Sheets("ESTOQUES (evolu��o)").Range(Cells(9, 1), Cells(54, 9)).Copy (Sheets("ESTOQUES (evolu��o)").Cells(59, 1))
'Limpa os n�meros que foram colados mas mant�m a formata��o
Sheets("ESTOQUES (evolu��o)").Range(Cells(59, 2), Cells(105, 9)).ClearContents



'Nomeia as colunas de acordo com o dado que ser� calculado nelas
Sheets("ESTOQUES (evolu��o)").Cells(58, 2).Value = "Diferen�a para o m�s anterior"
Sheets("ESTOQUES (evolu��o)").Cells(58, 3).Value = "Diferen�a para ao mesmo m�s do ano anterior"
Sheets("ESTOQUES (evolu��o)").Cells(58, 4).Value = "Diferen�a para a m�dia hist�rica"
Sheets("ESTOQUES (evolu��o)").Cells(58, 5).Value = "Posi��o Decrescente (Maior valor 1�, menor valor �ltimo)"
Sheets("ESTOQUES (evolu��o)").Cells(58, 6).Value = "Posi��o Crescente  (Menor valor 1�, maior valor �ltimo)"
Sheets("ESTOQUES (evolu��o)").Cells(58, 7).Value = "Posi��o Crescente - Mesmo m�s  (Menor valor 1�, maior valor �ltimo)"
Sheets("ESTOQUES (evolu��o)").Cells(58, 8).Value = "Posi��o Decrescente -Mesmo m�s  (Maior valor 1�, menor valor �ltimo)"
Sheets("ESTOQUES (evolu��o)").Cells(58, 9).Value = "cruzou a linha de 50?"

'Atribui valores as variaveis definidas acima
Estoques_Evolu��o_Linha_Dados = 9 'Define o n�mero da primeira linha de dados
Estoques_Evolu��o_Coluna_Dados1 = Sheets("ESTOQUES (evolu��o)").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Estoques_Evolu��o_Coluna_Dados2 = Estoques_Evolu��o_Coluna_Dados1 - 1 'Define o n�mero da coluna do m�s anterior
Estoques_Evolu��o_Coluna_Dados3 = Estoques_Evolu��o_Coluna_Dados1 - 12
Estoques_Evolu��o_Linha_An�lise = 59 'Define a primeira linhas de an�lises
Estoques_Evolu��o_Coluna_An�lise = 2 'Define a coluna de an�lises

'Inserindo valores nas celulas vazias para fugir de bugs
Sheets("ESTOQUES (evolu��o)").Range(Cells(10, Estoques_Evolu��o_Coluna_Dados3), Cells(10, Estoques_Evolu��o_Coluna_Dados1)).Value = "0"
Sheets("ESTOQUES (evolu��o)").Range(Cells(16, Estoques_Evolu��o_Coluna_Dados3), Cells(16, Estoques_Evolu��o_Coluna_Dados1)).Value = "0"
Sheets("ESTOQUES (evolu��o)").Range(Cells(20, Estoques_Evolu��o_Coluna_Dados3), Cells(20, Estoques_Evolu��o_Coluna_Dados1)).Value = "0"
Sheets("ESTOQUES (evolu��o)").Range(Cells(22, Estoques_Evolu��o_Coluna_Dados3), Cells(23, Estoques_Evolu��o_Coluna_Dados1)).Value = "0"
Sheets("ESTOQUES (evolu��o)").Range(Cells(25, Estoques_Evolu��o_Coluna_Dados3), Cells(25, Estoques_Evolu��o_Coluna_Dados1)).Value = "0"
Sheets("ESTOQUES (evolu��o)").Range(Cells(29, Estoques_Evolu��o_Coluna_Dados3), Cells(29, Estoques_Evolu��o_Coluna_Dados1)).Value = "0"
Sheets("ESTOQUES (evolu��o)").Range(Cells(37, Estoques_Evolu��o_Coluna_Dados3), Cells(37, Estoques_Evolu��o_Coluna_Dados1)).Value = "0"


'Calculo da difern�a em pontos do valor mais recente em rela��o ao valor do m�s anterior
Do Until Estoques_Evolu��o_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Calculo da diferen�a em si: o valor da celula de analise � igual ao valor mais recente menos o valor do m�s anterior
   Sheets("ESTOQUES (evolu��o)").Cells(Estoques_Evolu��o_Linha_An�lise, Estoques_Evolu��o_Coluna_An�lise).Value = Sheets("ESTOQUES (evolu��o)").Cells(Estoques_Evolu��o_Linha_Dados, Estoques_Evolu��o_Coluna_Dados1).Value - Sheets("ESTOQUES (evolu��o)").Cells(Estoques_Evolu��o_Linha_Dados, Estoques_Evolu��o_Coluna_Dados2).Value
    'Vai para a pr�xima linha de dados e de an�lise
   Estoques_Evolu��o_Linha_Dados = Estoques_Evolu��o_Linha_Dados + 1
   Estoques_Evolu��o_Linha_An�lise = Estoques_Evolu��o_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Estoques_Evolu��o_Linha_Dados = 9
Estoques_Evolu��o_Coluna_Dados1 = Sheets("ESTOQUES (evolu��o)").Range("A9").End(xlToRight).Column
Estoques_Evolu��o_Coluna_Dados3 = Estoques_Evolu��o_Coluna_Dados1 - 12
Estoques_Evolu��o_Linha_An�lise = 59
Estoques_Evolu��o_Coluna_An�lise = 3

'C�lculo da diferen�a em pontos do valor mais recente em rela��o ao valor do mesmo m�s do ano anterior
Do Until Estoques_Evolu��o_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Calculo da diferen�a em si: o valor da celula de analise � igual ao valor mais recente menos o valor do mesmo m�s do ano anterior
    Sheets("ESTOQUES (evolu��o)").Cells(Estoques_Evolu��o_Linha_An�lise, Estoques_Evolu��o_Coluna_An�lise).Value = Sheets("ESTOQUES (evolu��o)").Cells(Estoques_Evolu��o_Linha_Dados, Estoques_Evolu��o_Coluna_Dados1).Value - Sheets("ESTOQUES (evolu��o)").Cells(Estoques_Evolu��o_Linha_Dados, Estoques_Evolu��o_Coluna_Dados3).Value
    'Vai para a pr�xima linha de dados e de an�lise
    Estoques_Evolu��o_Linha_Dados = Estoques_Evolu��o_Linha_Dados + 1
    Estoques_Evolu��o_Linha_An�lise = Estoques_Evolu��o_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Estoques_Evolu��o_Linha_Dados = 9
Estoques_Evolu��o_Coluna_Dados1 = Sheets("ESTOQUES (evolu��o)").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Estoques_Evolu��o_Linha_An�lise = 59
Estoques_Evolu��o_Coluna_An�lise = 4

'C�lculo da diferen�a em pontos do valor mais recente em rela��o ao valor da m�dia hist�rica
Do Until Estoques_Evolu��o_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a vari�vel media como a m�dia do intervalo entre a coluna com o dado mais recente e o primeiro
    media = Application.Average(Sheets("ESTOQUES (evolu��o)").Range(Cells(Estoques_Evolu��o_Linha_Dados, Estoques_Evolu��o_Coluna_Dados1), Cells(Estoques_Evolu��o_Linha_Dados, 2)))
    'Calculo em si: o valor da celula de analise � igual ao valor mais recente menos o valor da m�dia
    Sheets("ESTOQUES (evolu��o)").Cells(Estoques_Evolu��o_Linha_An�lise, Estoques_Evolu��o_Coluna_An�lise).Value = Sheets("ESTOQUES (evolu��o)").Cells(Estoques_Evolu��o_Linha_Dados, Estoques_Evolu��o_Coluna_Dados1).Value - media
    'Vai para a pr�xima linha de dados e de an�lise
    Estoques_Evolu��o_Linha_Dados = Estoques_Evolu��o_Linha_Dados + 1
    Estoques_Evolu��o_Linha_An�lise = Estoques_Evolu��o_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Estoques_Evolu��o_Linha_Dados = 9
Estoques_Evolu��o_Coluna_Dados1 = Sheets("ESTOQUES (evolu��o)").Range("A9").End(xlToRight).Column
Estoques_Evolu��o_Linha_An�lise = 59
Estoques_Evolu��o_Coluna_An�lise = 5

'Ordena��o decrescente da s�rie hist�rica completa
Do Until Estoques_Evolu��o_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posi��o = WorksheetFunction.Rank_Eq(Sheets("ESTOQUES (evolu��o)").Cells(Estoques_Evolu��o_Linha_Dados, Estoques_Evolu��o_Coluna_Dados1), Sheets("ESTOQUES (evolu��o)").Range(Cells(Estoques_Evolu��o_Linha_Dados, Estoques_Evolu��o_Coluna_Dados1), Cells(Estoques_Evolu��o_Linha_Dados, 2)), 0)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("ESTOQUES (evolu��o)").Cells(Estoques_Evolu��o_Linha_An�lise, Estoques_Evolu��o_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Estoques_Evolu��o_Linha_Dados = Estoques_Evolu��o_Linha_Dados + 1
    Estoques_Evolu��o_Linha_An�lise = Estoques_Evolu��o_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Estoques_Evolu��o_Linha_Dados = 9
Estoques_Evolu��o_Coluna_Dados1 = Sheets("ESTOQUES (evolu��o)").Range("A9").End(xlToRight).Column
Estoques_Evolu��o_Linha_An�lise = 59
Estoques_Evolu��o_Coluna_An�lise = 6

'Ordena��o Crescente da s�rie hist�rica completa
Do Until Estoques_Evolu��o_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posi��o = WorksheetFunction.Rank_Eq(Sheets("ESTOQUES (evolu��o)").Cells(Estoques_Evolu��o_Linha_Dados, Estoques_Evolu��o_Coluna_Dados1), Sheets("ESTOQUES (evolu��o)").Range(Cells(Estoques_Evolu��o_Linha_Dados, Estoques_Evolu��o_Coluna_Dados1), Cells(Estoques_Evolu��o_Linha_Dados, 2)), 1)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("ESTOQUES (evolu��o)").Cells(Estoques_Evolu��o_Linha_An�lise, Estoques_Evolu��o_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Estoques_Evolu��o_Linha_Dados = Estoques_Evolu��o_Linha_Dados + 1
    Estoques_Evolu��o_Linha_An�lise = Estoques_Evolu��o_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior e define a vari�vel Coluna_DadosP que representa a primeira coluna de dados
Estoques_Evolu��o_Coluna_Dados1 = Sheets("ESTOQUES (evolu��o)").Range("A9").End(xlToRight).Column
Estoques_Evolu��o_Coluna_DadosP = 2

Do Until Estoques_Evolu��o_Coluna_DadosP = Estoques_Evolu��o_Coluna_Dados1 + 1 ' Faz at� a variavel Coluna_DadosP ser igual a variavel Coluna_Dados1 mais uma unidade
    'Confere se o m�s da coluna em quest�o � igual ao m�s do dado mais recente
    If Month(Sheets("ESTOQUES (evolu��o)").Cells(8, Estoques_Evolu��o_Coluna_DadosP)) = Month(Sheets("ESTOQUES (evolu��o)").Cells(8, Estoques_Evolu��o_Coluna_Dados1)) Then
    'Caso seja igual, copia a coluna com os dados mais abaixo, a partir da linha 110
        Sheets("ESTOQUES (evolu��o)").Range(Cells(9, Estoques_Evolu��o_Coluna_DadosP), (Cells(54, Estoques_Evolu��o_Coluna_DadosP))).Copy (Sheets("ESTOQUES (evolu��o)").Cells(110, Estoques_Evolu��o_Coluna_DadosP))
    End If
    'Vai para a pr�xima coluna
    Estoques_Evolu��o_Coluna_DadosP = Estoques_Evolu��o_Coluna_DadosP + 1
'Repete a conferencia com a pr�xima coluna
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Estoques_Evolu��o_Linha_Dados = 110
Estoques_Evolu��o_Coluna_Dados1 = Sheets("ESTOQUES (evolu��o)").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Estoques_Evolu��o_Linha_An�lise = 59
Estoques_Evolu��o_Coluna_An�lise = 7

'Ordena��o decrescente da s�rie hist�rica dos meses do dado mais recente
Do Until Estoques_Evolu��o_Linha_Dados = 156 'Faz o calculo at� a vari�vel Linha_Dados ser 156
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo m�s do mais recente
    posi��o = Application.WorksheetFunction.Rank_Eq(Sheets("ESTOQUES (evolu��o)").Cells(Estoques_Evolu��o_Linha_Dados, Estoques_Evolu��o_Coluna_Dados1), Sheets("ESTOQUES (evolu��o)").Range(Cells(Estoques_Evolu��o_Linha_Dados, Estoques_Evolu��o_Coluna_Dados1), Cells(Estoques_Evolu��o_Linha_Dados, 2)))
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("ESTOQUES (evolu��o)").Cells(Estoques_Evolu��o_Linha_An�lise, Estoques_Evolu��o_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Estoques_Evolu��o_Linha_Dados = Estoques_Evolu��o_Linha_Dados + 1
    Estoques_Evolu��o_Linha_An�lise = Estoques_Evolu��o_Linha_An�lise + 1
'Repete a ordena��o com a pr�xima linha
Loop

'Atribui os valores originais das variaveis ap�s o loop anterior
Estoques_Evolu��o_Linha_Dados = 110
Estoques_Evolu��o_Coluna_Dados1 = Sheets("ESTOQUES (evolu��o)").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Estoques_Evolu��o_Linha_An�lise = 59
Estoques_Evolu��o_Coluna_An�lise = 8
'Ordena��o crescente da s�rie hist�rica dos meses do dado mais recente
Do Until Estoques_Evolu��o_Linha_Dados = 156 'Faz o calculo at� a vari�vel Linha_Dados ser 156
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo m�s do mais recente
    posi��o = Application.WorksheetFunction.Rank_Eq(Sheets("ESTOQUES (evolu��o)").Cells(Estoques_Evolu��o_Linha_Dados, Estoques_Evolu��o_Coluna_Dados1), Sheets("ESTOQUES (evolu��o)").Range(Cells(Estoques_Evolu��o_Linha_Dados, Estoques_Evolu��o_Coluna_Dados1), Cells(Estoques_Evolu��o_Linha_Dados, 2)), 1)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("ESTOQUES (evolu��o)").Cells(Estoques_Evolu��o_Linha_An�lise, Estoques_Evolu��o_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Estoques_Evolu��o_Linha_Dados = Estoques_Evolu��o_Linha_Dados + 1
    Estoques_Evolu��o_Linha_An�lise = Estoques_Evolu��o_Linha_An�lise + 1
'Repete a ordena��o com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Estoques_Evolu��o_Linha_Dados = 9
Estoques_Evolu��o_Coluna_Dados1 = Sheets("ESTOQUES (evolu��o)").Range("A9").End(xlToRight).Column
Estoques_Evolu��o_Coluna_Dados2 = Estoques_Evolu��o_Coluna_Dados1 - 1
Estoques_Evolu��o_Linha_An�lise = 59
Estoques_Evolu��o_Coluna_An�lise = 9

'Avalia��o se cruzou ou n�o a linha de 50 e o sentido
Do Until Estoques_Evolu��o_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'se o dado mais recente for menor que 50 e o dado do m�s anterior for maior ou igual a 50 ent�o...
    If Cells(Estoques_Evolu��o_Linha_Dados, Estoques_Evolu��o_Coluna_Dados1) < 50 And Cells(Estoques_Evolu��o_Linha_Dados, Estoques_Evolu��o_Coluna_Dados2) >= 50 Then
    
    'a c�lula de an�lise recebe cruzou para baixo
    Cells(Estoques_Evolu��o_Linha_An�lise, Estoques_Evolu��o_Coluna_An�lise).Value = "Cruzou para baixo"
    'Caso n�o seja..
    Else
        'se o dado mais recente for maior ou igual a 50 e o dado do m~es anterior for menor ou igual a 50 ent�o...
        If Cells(Estoques_Evolu��o_Linha_Dados, Estoques_Evolu��o_Coluna_Dados1) >= 50 And Cells(Estoques_Evolu��o_Linha_Dados, Estoques_Evolu��o_Coluna_Dados2) <= 50 Then
        'a c�lula de an�lise recebe cruzou para cima
        Cells(Estoques_Evolu��o_Linha_An�lise, Estoques_Evolu��o_Coluna_An�lise).Value = "Cruzou para cima"
        'Caso n�o seja..
        Else
        'a c�lula de an�lise recebe n�o cruzou
        Cells(Estoques_Evolu��o_Linha_An�lise, Estoques_Evolu��o_Coluna_An�lise).Value = "N�o Cruzou"
        End If
    End If
    'Vai para a pr�xima linha de dados e de an�lise
    Estoques_Evolu��o_Linha_Dados = Estoques_Evolu��o_Linha_Dados + 1
    Estoques_Evolu��o_Linha_An�lise = Estoques_Evolu��o_Linha_An�lise + 1
'repete o processo com a nova linha
Loop

'Apaga as linhas com erros/dados faltantes/t�tulos e subt�tulos
Range(Cells(60, 2), Cells(60, 9)).ClearContents
Range(Cells(66, 2), Cells(66, 9)).ClearContents
Range(Cells(70, 2), Cells(70, 9)).ClearContents
Range(Cells(72, 2), Cells(73, 9)).Value = "-"
Range(Cells(75, 2), Cells(75, 9)).Value = "-"
Range(Cells(79, 2), Cells(79, 9)).Value = "-"
Range(Cells(87, 2), Cells(87, 9)).Value = "-"

'Inserindo valores nas celulas vazias para fugir de bugs
Range(Cells(10, Estoques_Evolu��o_Coluna_Dados3), Cells(10, Estoques_Evolu��o_Coluna_Dados1)).ClearContents
Range(Cells(16, Estoques_Evolu��o_Coluna_Dados3), Cells(16, Estoques_Evolu��o_Coluna_Dados1)).ClearContents
Range(Cells(20, Estoques_Evolu��o_Coluna_Dados3), Cells(20, Estoques_Evolu��o_Coluna_Dados1)).ClearContents
Range(Cells(22, Estoques_Evolu��o_Coluna_Dados3), Cells(23, Estoques_Evolu��o_Coluna_Dados1)).Value = "-"
Range(Cells(25, Estoques_Evolu��o_Coluna_Dados3), Cells(25, Estoques_Evolu��o_Coluna_Dados1)).Value = "-"
Range(Cells(29, Estoques_Evolu��o_Coluna_Dados3), Cells(29, Estoques_Evolu��o_Coluna_Dados1)).Value = "-"
Range(Cells(37, Estoques_Evolu��o_Coluna_Dados3), Cells(37, Estoques_Evolu��o_Coluna_Dados1)).Value = "-"

Range("E59:H104").NumberFormat = "0"





'*******************************************           Estoques_Efetivo_Planejado           *******************************************************
  
 
 

Dim Estoques_Efetivo_Planejado_Linha_Dados As Integer 'Define a linha que cont�m o dado a ser usado
Dim Estoques_Efetivo_Planejado_Coluna_Dados1 As Integer ' Define a coluna com o dado mais recente
Dim Estoques_Efetivo_Planejado_Coluna_Dados2 As Integer ' Define a coluna com o dado do m�s anterior
Dim Estoques_Efetivo_Planejado_Coluna_Dados3 As Integer ' Defie a coluna com o dado do mesmo m�s do ano anterior
Dim Estoques_Efetivo_Planejado_Linha_An�lise As Integer ' Define a linha que ser� feita a an�lise
Dim Estoques_Efetivo_Planejado_Coluna_An�lise As Integer 'Define a coluna que ser� feita a an�lise
Sheets("Estoques (efetivo-planejado)").Select


'Copia os t�tulos das categorias e cola onde ser� formada a tabela de an�lise
Sheets("Estoques (efetivo-planejado)").Range(Cells(9, 1), Cells(54, 9)).Copy (Sheets("Estoques (efetivo-planejado)").Cells(59, 1))
'Limpa os n�meros que foram colados mas mant�m a formata��o
Sheets("Estoques (efetivo-planejado)").Range(Cells(59, 2), Cells(105, 9)).ClearContents



'Nomeia as colunas de acordo com o dado que ser� calculado nelas
Sheets("Estoques (efetivo-planejado)").Cells(58, 2).Value = "Diferen�a para o m�s anterior"
Sheets("Estoques (efetivo-planejado)").Cells(58, 3).Value = "Diferen�a para ao mesmo m�s do ano anterior"
Sheets("Estoques (efetivo-planejado)").Cells(58, 4).Value = "Diferen�a para a m�dia hist�rica"
Sheets("Estoques (efetivo-planejado)").Cells(58, 5).Value = "Posi��o Decrescente (Maior valor 1�, menor valor �ltimo)"
Sheets("Estoques (efetivo-planejado)").Cells(58, 6).Value = "Posi��o Crescente  (Menor valor 1�, maior valor �ltimo)"
Sheets("Estoques (efetivo-planejado)").Cells(58, 7).Value = "Posi��o Crescente - Mesmo m�s  (Menor valor 1�, maior valor �ltimo)"
Sheets("Estoques (efetivo-planejado)").Cells(58, 8).Value = "Posi��o Decrescente -Mesmo m�s  (Maior valor 1�, menor valor �ltimo)"
Sheets("Estoques (efetivo-planejado)").Cells(58, 9).Value = "cruzou a linha de 50?"

'Atribui valores as variaveis definidas acima
Estoques_Efetivo_Planejado_Linha_Dados = 9 'Define o n�mero da primeira linha de dados
Estoques_Efetivo_Planejado_Coluna_Dados1 = Sheets("Estoques (efetivo-planejado)").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Estoques_Efetivo_Planejado_Coluna_Dados2 = Estoques_Efetivo_Planejado_Coluna_Dados1 - 1 'Define o n�mero da coluna do m�s anterior
Estoques_Efetivo_Planejado_Coluna_Dados3 = Estoques_Efetivo_Planejado_Coluna_Dados1 - 12
Estoques_Efetivo_Planejado_Linha_An�lise = 59 'Define a primeira linhas de an�lises
Estoques_Efetivo_Planejado_Coluna_An�lise = 2 'Define a coluna de an�lises

'Inserindo valores nas celulas vazias para fugir de bugs
Sheets("Estoques (efetivo-planejado)").Range(Cells(10, Estoques_Efetivo_Planejado_Coluna_Dados3), Cells(10, Estoques_Efetivo_Planejado_Coluna_Dados1)).Value = "0"
Sheets("Estoques (efetivo-planejado)").Range(Cells(16, Estoques_Efetivo_Planejado_Coluna_Dados3), Cells(16, Estoques_Efetivo_Planejado_Coluna_Dados1)).Value = "0"
Sheets("Estoques (efetivo-planejado)").Range(Cells(20, Estoques_Efetivo_Planejado_Coluna_Dados3), Cells(20, Estoques_Efetivo_Planejado_Coluna_Dados1)).Value = "0"
Sheets("Estoques (efetivo-planejado)").Range(Cells(22, Estoques_Efetivo_Planejado_Coluna_Dados3), Cells(23, Estoques_Efetivo_Planejado_Coluna_Dados1)).Value = "0"
Sheets("Estoques (efetivo-planejado)").Range(Cells(25, Estoques_Efetivo_Planejado_Coluna_Dados3), Cells(25, Estoques_Efetivo_Planejado_Coluna_Dados1)).Value = "0"
Sheets("Estoques (efetivo-planejado)").Range(Cells(29, Estoques_Efetivo_Planejado_Coluna_Dados3), Cells(29, Estoques_Efetivo_Planejado_Coluna_Dados1)).Value = "0"
Sheets("Estoques (efetivo-planejado)").Range(Cells(37, Estoques_Efetivo_Planejado_Coluna_Dados3), Cells(37, Estoques_Efetivo_Planejado_Coluna_Dados1)).Value = "0"


'Calculo da difern�a em pontos do valor mais recente em rela��o ao valor do m�s anterior
Do Until Estoques_Efetivo_Planejado_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Calculo da diferen�a em si: o valor da celula de analise � igual ao valor mais recente menos o valor do m�s anterior
   Sheets("Estoques (efetivo-planejado)").Cells(Estoques_Efetivo_Planejado_Linha_An�lise, Estoques_Efetivo_Planejado_Coluna_An�lise).Value = Sheets("Estoques (efetivo-planejado)").Cells(Estoques_Efetivo_Planejado_Linha_Dados, Estoques_Efetivo_Planejado_Coluna_Dados1).Value - Sheets("Estoques (efetivo-planejado)").Cells(Estoques_Efetivo_Planejado_Linha_Dados, Estoques_Efetivo_Planejado_Coluna_Dados2).Value
    'Vai para a pr�xima linha de dados e de an�lise
   Estoques_Efetivo_Planejado_Linha_Dados = Estoques_Efetivo_Planejado_Linha_Dados + 1
   Estoques_Efetivo_Planejado_Linha_An�lise = Estoques_Efetivo_Planejado_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Estoques_Efetivo_Planejado_Linha_Dados = 9
Estoques_Efetivo_Planejado_Coluna_Dados1 = Sheets("Estoques (efetivo-planejado)").Range("A9").End(xlToRight).Column
Estoques_Efetivo_Planejado_Coluna_Dados3 = Estoques_Efetivo_Planejado_Coluna_Dados1 - 12
Estoques_Efetivo_Planejado_Linha_An�lise = 59
Estoques_Efetivo_Planejado_Coluna_An�lise = 3

'C�lculo da diferen�a em pontos do valor mais recente em rela��o ao valor do mesmo m�s do ano anterior
Do Until Estoques_Efetivo_Planejado_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Calculo da diferen�a em si: o valor da celula de analise � igual ao valor mais recente menos o valor do mesmo m�s do ano anterior
    Sheets("Estoques (efetivo-planejado)").Cells(Estoques_Efetivo_Planejado_Linha_An�lise, Estoques_Efetivo_Planejado_Coluna_An�lise).Value = Sheets("Estoques (efetivo-planejado)").Cells(Estoques_Efetivo_Planejado_Linha_Dados, Estoques_Efetivo_Planejado_Coluna_Dados1).Value - Sheets("Estoques (efetivo-planejado)").Cells(Estoques_Efetivo_Planejado_Linha_Dados, Estoques_Efetivo_Planejado_Coluna_Dados3).Value
    'Vai para a pr�xima linha de dados e de an�lise
    Estoques_Efetivo_Planejado_Linha_Dados = Estoques_Efetivo_Planejado_Linha_Dados + 1
    Estoques_Efetivo_Planejado_Linha_An�lise = Estoques_Efetivo_Planejado_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Estoques_Efetivo_Planejado_Linha_Dados = 9
Estoques_Efetivo_Planejado_Coluna_Dados1 = Sheets("Estoques (efetivo-planejado)").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Estoques_Efetivo_Planejado_Linha_An�lise = 59
Estoques_Efetivo_Planejado_Coluna_An�lise = 4

'C�lculo da diferen�a em pontos do valor mais recente em rela��o ao valor da m�dia hist�rica
Do Until Estoques_Efetivo_Planejado_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a vari�vel media como a m�dia do intervalo entre a coluna com o dado mais recente e o primeiro
    media = Application.Average(Sheets("Estoques (efetivo-planejado)").Range(Cells(Estoques_Efetivo_Planejado_Linha_Dados, Estoques_Efetivo_Planejado_Coluna_Dados1), Cells(Estoques_Efetivo_Planejado_Linha_Dados, 2)))
    'Calculo em si: o valor da celula de analise � igual ao valor mais recente menos o valor da m�dia
    Sheets("Estoques (efetivo-planejado)").Cells(Estoques_Efetivo_Planejado_Linha_An�lise, Estoques_Efetivo_Planejado_Coluna_An�lise).Value = Sheets("Estoques (efetivo-planejado)").Cells(Estoques_Efetivo_Planejado_Linha_Dados, Estoques_Efetivo_Planejado_Coluna_Dados1).Value - media
    'Vai para a pr�xima linha de dados e de an�lise
    Estoques_Efetivo_Planejado_Linha_Dados = Estoques_Efetivo_Planejado_Linha_Dados + 1
    Estoques_Efetivo_Planejado_Linha_An�lise = Estoques_Efetivo_Planejado_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Estoques_Efetivo_Planejado_Linha_Dados = 9
Estoques_Efetivo_Planejado_Coluna_Dados1 = Sheets("Estoques (efetivo-planejado)").Range("A9").End(xlToRight).Column
Estoques_Efetivo_Planejado_Linha_An�lise = 59
Estoques_Efetivo_Planejado_Coluna_An�lise = 5

'Ordena��o decrescente da s�rie hist�rica completa
Do Until Estoques_Efetivo_Planejado_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posi��o = WorksheetFunction.Rank_Eq(Sheets("Estoques (efetivo-planejado)").Cells(Estoques_Efetivo_Planejado_Linha_Dados, Estoques_Efetivo_Planejado_Coluna_Dados1), Sheets("Estoques (efetivo-planejado)").Range(Cells(Estoques_Efetivo_Planejado_Linha_Dados, Estoques_Efetivo_Planejado_Coluna_Dados1), Cells(Estoques_Efetivo_Planejado_Linha_Dados, 2)), 0)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("Estoques (efetivo-planejado)").Cells(Estoques_Efetivo_Planejado_Linha_An�lise, Estoques_Efetivo_Planejado_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Estoques_Efetivo_Planejado_Linha_Dados = Estoques_Efetivo_Planejado_Linha_Dados + 1
    Estoques_Efetivo_Planejado_Linha_An�lise = Estoques_Efetivo_Planejado_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Estoques_Efetivo_Planejado_Linha_Dados = 9
Estoques_Efetivo_Planejado_Coluna_Dados1 = Sheets("Estoques (efetivo-planejado)").Range("A9").End(xlToRight).Column
Estoques_Efetivo_Planejado_Linha_An�lise = 59
Estoques_Efetivo_Planejado_Coluna_An�lise = 6

'Ordena��o Crescente da s�rie hist�rica completa
Do Until Estoques_Efetivo_Planejado_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posi��o = WorksheetFunction.Rank_Eq(Sheets("Estoques (efetivo-planejado)").Cells(Estoques_Efetivo_Planejado_Linha_Dados, Estoques_Efetivo_Planejado_Coluna_Dados1), Sheets("Estoques (efetivo-planejado)").Range(Cells(Estoques_Efetivo_Planejado_Linha_Dados, Estoques_Efetivo_Planejado_Coluna_Dados1), Cells(Estoques_Efetivo_Planejado_Linha_Dados, 2)), 1)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("Estoques (efetivo-planejado)").Cells(Estoques_Efetivo_Planejado_Linha_An�lise, Estoques_Efetivo_Planejado_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Estoques_Efetivo_Planejado_Linha_Dados = Estoques_Efetivo_Planejado_Linha_Dados + 1
    Estoques_Efetivo_Planejado_Linha_An�lise = Estoques_Efetivo_Planejado_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior e define a vari�vel Coluna_DadosP que representa a primeira coluna de dados
Estoques_Efetivo_Planejado_Coluna_Dados1 = Sheets("Estoques (efetivo-planejado)").Range("A9").End(xlToRight).Column
Estoques_Efetivo_Planejado_Coluna_DadosP = 2

Do Until Estoques_Efetivo_Planejado_Coluna_DadosP = Estoques_Efetivo_Planejado_Coluna_Dados1 + 1 ' Faz at� a variavel Coluna_DadosP ser igual a variavel Coluna_Dados1 mais uma unidade
    'Confere se o m�s da coluna em quest�o � igual ao m�s do dado mais recente
    If Month(Sheets("Estoques (efetivo-planejado)").Cells(8, Estoques_Efetivo_Planejado_Coluna_DadosP)) = Month(Sheets("Estoques (efetivo-planejado)").Cells(8, Estoques_Efetivo_Planejado_Coluna_Dados1)) Then
    'Caso seja igual, copia a coluna com os dados mais abaixo, a partir da linha 110
        Sheets("Estoques (efetivo-planejado)").Range(Cells(9, Estoques_Efetivo_Planejado_Coluna_DadosP), (Cells(54, Estoques_Efetivo_Planejado_Coluna_DadosP))).Copy (Sheets("Estoques (efetivo-planejado)").Cells(110, Estoques_Efetivo_Planejado_Coluna_DadosP))
    End If
    'Vai para a pr�xima coluna
    Estoques_Efetivo_Planejado_Coluna_DadosP = Estoques_Efetivo_Planejado_Coluna_DadosP + 1
'Repete a conferencia com a pr�xima coluna
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Estoques_Efetivo_Planejado_Linha_Dados = 110
Estoques_Efetivo_Planejado_Coluna_Dados1 = Sheets("Estoques (efetivo-planejado)").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Estoques_Efetivo_Planejado_Linha_An�lise = 59
Estoques_Efetivo_Planejado_Coluna_An�lise = 7

'Ordena��o decrescente da s�rie hist�rica dos meses do dado mais recente
Do Until Estoques_Efetivo_Planejado_Linha_Dados = 156 'Faz o calculo at� a vari�vel Linha_Dados ser 156
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo m�s do mais recente
    posi��o = Application.WorksheetFunction.Rank_Eq(Sheets("Estoques (efetivo-planejado)").Cells(Estoques_Efetivo_Planejado_Linha_Dados, Estoques_Efetivo_Planejado_Coluna_Dados1), Sheets("Estoques (efetivo-planejado)").Range(Cells(Estoques_Efetivo_Planejado_Linha_Dados, Estoques_Efetivo_Planejado_Coluna_Dados1), Cells(Estoques_Efetivo_Planejado_Linha_Dados, 2)))
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("Estoques (efetivo-planejado)").Cells(Estoques_Efetivo_Planejado_Linha_An�lise, Estoques_Efetivo_Planejado_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Estoques_Efetivo_Planejado_Linha_Dados = Estoques_Efetivo_Planejado_Linha_Dados + 1
    Estoques_Efetivo_Planejado_Linha_An�lise = Estoques_Efetivo_Planejado_Linha_An�lise + 1
'Repete a ordena��o com a pr�xima linha
Loop

'Atribui os valores originais das variaveis ap�s o loop anterior
Estoques_Efetivo_Planejado_Linha_Dados = 110
Estoques_Efetivo_Planejado_Coluna_Dados1 = Sheets("Estoques (efetivo-planejado)").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Estoques_Efetivo_Planejado_Linha_An�lise = 59
Estoques_Efetivo_Planejado_Coluna_An�lise = 8
'Ordena��o crescente da s�rie hist�rica dos meses do dado mais recente
Do Until Estoques_Efetivo_Planejado_Linha_Dados = 156 'Faz o calculo at� a vari�vel Linha_Dados ser 156
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo m�s do mais recente
    posi��o = Application.WorksheetFunction.Rank_Eq(Sheets("Estoques (efetivo-planejado)").Cells(Estoques_Efetivo_Planejado_Linha_Dados, Estoques_Efetivo_Planejado_Coluna_Dados1), Sheets("Estoques (efetivo-planejado)").Range(Cells(Estoques_Efetivo_Planejado_Linha_Dados, Estoques_Efetivo_Planejado_Coluna_Dados1), Cells(Estoques_Efetivo_Planejado_Linha_Dados, 2)), 1)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("Estoques (efetivo-planejado)").Cells(Estoques_Efetivo_Planejado_Linha_An�lise, Estoques_Efetivo_Planejado_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Estoques_Efetivo_Planejado_Linha_Dados = Estoques_Efetivo_Planejado_Linha_Dados + 1
    Estoques_Efetivo_Planejado_Linha_An�lise = Estoques_Efetivo_Planejado_Linha_An�lise + 1
'Repete a ordena��o com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Estoques_Efetivo_Planejado_Linha_Dados = 9
Estoques_Efetivo_Planejado_Coluna_Dados1 = Sheets("Estoques (efetivo-planejado)").Range("A9").End(xlToRight).Column
Estoques_Efetivo_Planejado_Coluna_Dados2 = Estoques_Efetivo_Planejado_Coluna_Dados1 - 1
Estoques_Efetivo_Planejado_Linha_An�lise = 59
Estoques_Efetivo_Planejado_Coluna_An�lise = 9

'Avalia��o se cruzou ou n�o a linha de 50 e o sentido
Do Until Estoques_Efetivo_Planejado_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'se o dado mais recente for menor que 50 e o dado do m�s anterior for maior ou igual a 50 ent�o...
    If Cells(Estoques_Efetivo_Planejado_Linha_Dados, Estoques_Efetivo_Planejado_Coluna_Dados1) < 50 And Cells(Estoques_Efetivo_Planejado_Linha_Dados, Estoques_Efetivo_Planejado_Coluna_Dados2) >= 50 Then
    
    'a c�lula de an�lise recebe cruzou para baixo
    Cells(Estoques_Efetivo_Planejado_Linha_An�lise, Estoques_Efetivo_Planejado_Coluna_An�lise).Value = "Cruzou para baixo"
    'Caso n�o seja..
    Else
        'se o dado mais recente for maior ou igual a 50 e o dado do m~es anterior for menor ou igual a 50 ent�o...
        If Cells(Estoques_Efetivo_Planejado_Linha_Dados, Estoques_Efetivo_Planejado_Coluna_Dados1) >= 50 And Cells(Estoques_Efetivo_Planejado_Linha_Dados, Estoques_Efetivo_Planejado_Coluna_Dados2) <= 50 Then
        'a c�lula de an�lise recebe cruzou para cima
        Cells(Estoques_Efetivo_Planejado_Linha_An�lise, Estoques_Efetivo_Planejado_Coluna_An�lise).Value = "Cruzou para cima"
        'Caso n�o seja..
        Else
        'a c�lula de an�lise recebe n�o cruzou
        Cells(Estoques_Efetivo_Planejado_Linha_An�lise, Estoques_Efetivo_Planejado_Coluna_An�lise).Value = "N�o Cruzou"
        End If
    End If
    'Vai para a pr�xima linha de dados e de an�lise
    Estoques_Efetivo_Planejado_Linha_Dados = Estoques_Efetivo_Planejado_Linha_Dados + 1
    Estoques_Efetivo_Planejado_Linha_An�lise = Estoques_Efetivo_Planejado_Linha_An�lise + 1
'repete o processo com a nova linha
Loop

'Apaga as linhas com erros/dados faltantes/t�tulos e subt�tulos
Range(Cells(60, 2), Cells(60, 9)).ClearContents
Range(Cells(66, 2), Cells(66, 9)).ClearContents
Range(Cells(70, 2), Cells(70, 9)).ClearContents
Range(Cells(72, 2), Cells(73, 9)).Value = "-"
Range(Cells(75, 2), Cells(75, 9)).Value = "-"
Range(Cells(79, 2), Cells(79, 9)).Value = "-"
Range(Cells(87, 2), Cells(87, 9)).Value = "-"

'Inserindo valores nas celulas vazias para fugir de bugs
Range(Cells(10, Estoques_Efetivo_Planejado_Coluna_Dados3), Cells(10, Estoques_Efetivo_Planejado_Coluna_Dados1)).ClearContents
Range(Cells(16, Estoques_Efetivo_Planejado_Coluna_Dados3), Cells(16, Estoques_Efetivo_Planejado_Coluna_Dados1)).ClearContents
Range(Cells(20, Estoques_Efetivo_Planejado_Coluna_Dados3), Cells(20, Estoques_Efetivo_Planejado_Coluna_Dados1)).ClearContents
Range(Cells(22, Estoques_Efetivo_Planejado_Coluna_Dados3), Cells(23, Estoques_Efetivo_Planejado_Coluna_Dados1)).Value = "-"
Range(Cells(25, Estoques_Efetivo_Planejado_Coluna_Dados3), Cells(25, Estoques_Efetivo_Planejado_Coluna_Dados1)).Value = "-"
Range(Cells(29, Estoques_Efetivo_Planejado_Coluna_Dados3), Cells(29, Estoques_Efetivo_Planejado_Coluna_Dados1)).Value = "-"
Range(Cells(37, Estoques_Efetivo_Planejado_Coluna_Dados3), Cells(37, Estoques_Efetivo_Planejado_Coluna_Dados1)).Value = "-"

Range("E59:H104").NumberFormat = "0"

End Sub


Sub An�lise_Azul()

 
Dim Expectativas_Demanda_Linha_Dados As Integer 'Define a linha que cont�m o dado a ser usado
Dim Expectativas_Demanda_Coluna_Dados1 As Integer ' Define a coluna com o dado mais recente
Dim Expectativas_Demanda_Coluna_Dados2 As Integer ' Define a coluna com o dado do m�s anterior
Dim Expectativas_Demanda_Coluna_Dados3 As Integer ' Defie a coluna com o dado do mesmo m�s do ano anterior
Dim Expectativas_Demanda_Linha_An�lise As Integer ' Define a linha que ser� feita a an�lise
Dim Expectativas_Demanda_Coluna_An�lise As Integer 'Define a coluna que ser� feita a an�lise

Sheets("EXPECTATIVAS - DEMANDA").Select

'Copia os t�tulos das categorias e cola onde ser� formada a tabela de an�lise
Sheets("EXPECTATIVAS - DEMANDA").Range(Cells(9, 1), Cells(54, 9)).Copy (Sheets("EXPECTATIVAS - DEMANDA").Cells(59, 1))
'Limpa os n�meros que foram colados mas mant�m a formata��o
Sheets("EXPECTATIVAS - DEMANDA").Range(Cells(59, 2), Cells(105, 9)).ClearContents



'Nomeia as colunas de acordo com o dado que ser� calculado nelas
Sheets("EXPECTATIVAS - DEMANDA").Cells(58, 2).Value = "Diferen�a para o m�s anterior"
Sheets("EXPECTATIVAS - DEMANDA").Cells(58, 3).Value = "Diferen�a para ao mesmo m�s do ano anterior"
Sheets("EXPECTATIVAS - DEMANDA").Cells(58, 4).Value = "Diferen�a para a m�dia hist�rica"
Sheets("EXPECTATIVAS - DEMANDA").Cells(58, 5).Value = "Posi��o Decrescente (Maior valor 1�, menor valor �ltimo)"
Sheets("EXPECTATIVAS - DEMANDA").Cells(58, 6).Value = "Posi��o Crescente  (Menor valor 1�, maior valor �ltimo)"
Sheets("EXPECTATIVAS - DEMANDA").Cells(58, 7).Value = "Posi��o Crescente - Mesmo m�s  (Menor valor 1�, maior valor �ltimo)"
Sheets("EXPECTATIVAS - DEMANDA").Cells(58, 8).Value = "Posi��o Decrescente -Mesmo m�s  (Maior valor 1�, menor valor �ltimo)"
Sheets("EXPECTATIVAS - DEMANDA").Cells(58, 9).Value = "cruzou a linha de 50?"

'Atribui valores as variaveis definidas acima
Expectativas_Demanda_Linha_Dados = 9 'Define o n�mero da primeira linha de dados
Expectativas_Demanda_Coluna_Dados1 = Sheets("EXPECTATIVAS - DEMANDA").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Expectativas_Demanda_Coluna_Dados2 = Expectativas_Demanda_Coluna_Dados1 - 1 'Define o n�mero da coluna do m�s anterior
Expectativas_Demanda_Coluna_Dados3 = Expectativas_Demanda_Coluna_Dados1 - 12
Expectativas_Demanda_Linha_An�lise = 59 'Define a primeira linhas de an�lises
Expectativas_Demanda_Coluna_An�lise = 2 'Define a coluna de an�lises

'Inserindo valores nas celulas vazias para fugir de bugs
Sheets("EXPECTATIVAS - DEMANDA").Range(Cells(10, Expectativas_Demanda_Coluna_Dados3), Cells(10, Expectativas_Demanda_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVAS - DEMANDA").Range(Cells(16, Expectativas_Demanda_Coluna_Dados3), Cells(16, Expectativas_Demanda_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVAS - DEMANDA").Range(Cells(20, Expectativas_Demanda_Coluna_Dados3), Cells(20, Expectativas_Demanda_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVAS - DEMANDA").Range(Cells(22, Expectativas_Demanda_Coluna_Dados3), Cells(23, Expectativas_Demanda_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVAS - DEMANDA").Range(Cells(25, Expectativas_Demanda_Coluna_Dados3), Cells(25, Expectativas_Demanda_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVAS - DEMANDA").Range(Cells(29, Expectativas_Demanda_Coluna_Dados3), Cells(29, Expectativas_Demanda_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVAS - DEMANDA").Range(Cells(37, Expectativas_Demanda_Coluna_Dados3), Cells(37, Expectativas_Demanda_Coluna_Dados1)).Value = "0"


'Calculo da difern�a em pontos do valor mais recente em rela��o ao valor do m�s anterior
Do Until Expectativas_Demanda_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Calculo da diferen�a em si: o valor da celula de analise � igual ao valor mais recente menos o valor do m�s anterior
   Sheets("EXPECTATIVAS - DEMANDA").Cells(Expectativas_Demanda_Linha_An�lise, Expectativas_Demanda_Coluna_An�lise).Value = Sheets("EXPECTATIVAS - DEMANDA").Cells(Expectativas_Demanda_Linha_Dados, Expectativas_Demanda_Coluna_Dados1).Value - Sheets("EXPECTATIVAS - DEMANDA").Cells(Expectativas_Demanda_Linha_Dados, Expectativas_Demanda_Coluna_Dados2).Value
    'Vai para a pr�xima linha de dados e de an�lise
   Expectativas_Demanda_Linha_Dados = Expectativas_Demanda_Linha_Dados + 1
   Expectativas_Demanda_Linha_An�lise = Expectativas_Demanda_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Expectativas_Demanda_Linha_Dados = 9
Expectativas_Demanda_Coluna_Dados1 = Sheets("EXPECTATIVAS - DEMANDA").Range("A9").End(xlToRight).Column
Expectativas_Demanda_Coluna_Dados3 = Expectativas_Demanda_Coluna_Dados1 - 12
Expectativas_Demanda_Linha_An�lise = 59
Expectativas_Demanda_Coluna_An�lise = 3

'C�lculo da diferen�a em pontos do valor mais recente em rela��o ao valor do mesmo m�s do ano anterior
Do Until Expectativas_Demanda_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Calculo da diferen�a em si: o valor da celula de analise � igual ao valor mais recente menos o valor do mesmo m�s do ano anterior
    Sheets("EXPECTATIVAS - DEMANDA").Cells(Expectativas_Demanda_Linha_An�lise, Expectativas_Demanda_Coluna_An�lise).Value = Sheets("EXPECTATIVAS - DEMANDA").Cells(Expectativas_Demanda_Linha_Dados, Expectativas_Demanda_Coluna_Dados1).Value - Sheets("EXPECTATIVAS - DEMANDA").Cells(Expectativas_Demanda_Linha_Dados, Expectativas_Demanda_Coluna_Dados3).Value
    'Vai para a pr�xima linha de dados e de an�lise
    Expectativas_Demanda_Linha_Dados = Expectativas_Demanda_Linha_Dados + 1
    Expectativas_Demanda_Linha_An�lise = Expectativas_Demanda_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Expectativas_Demanda_Linha_Dados = 9
Expectativas_Demanda_Coluna_Dados1 = Sheets("EXPECTATIVAS - DEMANDA").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Expectativas_Demanda_Linha_An�lise = 59
Expectativas_Demanda_Coluna_An�lise = 4

'C�lculo da diferen�a em pontos do valor mais recente em rela��o ao valor da m�dia hist�rica
Do Until Expectativas_Demanda_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a vari�vel media como a m�dia do intervalo entre a coluna com o dado mais recente e o primeiro
    media = Application.Average(Sheets("EXPECTATIVAS - DEMANDA").Range(Cells(Expectativas_Demanda_Linha_Dados, Expectativas_Demanda_Coluna_Dados1), Cells(Expectativas_Demanda_Linha_Dados, 2)))
    'Calculo em si: o valor da celula de analise � igual ao valor mais recente menos o valor da m�dia
    Sheets("EXPECTATIVAS - DEMANDA").Cells(Expectativas_Demanda_Linha_An�lise, Expectativas_Demanda_Coluna_An�lise).Value = Sheets("EXPECTATIVAS - DEMANDA").Cells(Expectativas_Demanda_Linha_Dados, Expectativas_Demanda_Coluna_Dados1).Value - media
    'Vai para a pr�xima linha de dados e de an�lise
    Expectativas_Demanda_Linha_Dados = Expectativas_Demanda_Linha_Dados + 1
    Expectativas_Demanda_Linha_An�lise = Expectativas_Demanda_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Expectativas_Demanda_Linha_Dados = 9
Expectativas_Demanda_Coluna_Dados1 = Sheets("EXPECTATIVAS - DEMANDA").Range("A9").End(xlToRight).Column
Expectativas_Demanda_Linha_An�lise = 59
Expectativas_Demanda_Coluna_An�lise = 5

'Ordena��o decrescente da s�rie hist�rica completa
Do Until Expectativas_Demanda_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posi��o = WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVAS - DEMANDA").Cells(Expectativas_Demanda_Linha_Dados, Expectativas_Demanda_Coluna_Dados1), Sheets("EXPECTATIVAS - DEMANDA").Range(Cells(Expectativas_Demanda_Linha_Dados, Expectativas_Demanda_Coluna_Dados1), Cells(Expectativas_Demanda_Linha_Dados, 2)), 0)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("EXPECTATIVAS - DEMANDA").Cells(Expectativas_Demanda_Linha_An�lise, Expectativas_Demanda_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Expectativas_Demanda_Linha_Dados = Expectativas_Demanda_Linha_Dados + 1
    Expectativas_Demanda_Linha_An�lise = Expectativas_Demanda_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Expectativas_Demanda_Linha_Dados = 9
Expectativas_Demanda_Coluna_Dados1 = Sheets("EXPECTATIVAS - DEMANDA").Range("A9").End(xlToRight).Column
Expectativas_Demanda_Linha_An�lise = 59
Expectativas_Demanda_Coluna_An�lise = 6

'Ordena��o Crescente da s�rie hist�rica completa
Do Until Expectativas_Demanda_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posi��o = WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVAS - DEMANDA").Cells(Expectativas_Demanda_Linha_Dados, Expectativas_Demanda_Coluna_Dados1), Sheets("EXPECTATIVAS - DEMANDA").Range(Cells(Expectativas_Demanda_Linha_Dados, Expectativas_Demanda_Coluna_Dados1), Cells(Expectativas_Demanda_Linha_Dados, 2)), 1)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("EXPECTATIVAS - DEMANDA").Cells(Expectativas_Demanda_Linha_An�lise, Expectativas_Demanda_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Expectativas_Demanda_Linha_Dados = Expectativas_Demanda_Linha_Dados + 1
    Expectativas_Demanda_Linha_An�lise = Expectativas_Demanda_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior e define a vari�vel Coluna_DadosP que representa a primeira coluna de dados
Expectativas_Demanda_Coluna_Dados1 = Sheets("EXPECTATIVAS - DEMANDA").Range("A9").End(xlToRight).Column
Expectativas_Demanda_Coluna_DadosP = 2

Do Until Expectativas_Demanda_Coluna_DadosP = Expectativas_Demanda_Coluna_Dados1 + 1 ' Faz at� a variavel Coluna_DadosP ser igual a variavel Coluna_Dados1 mais uma unidade
    'Confere se o m�s da coluna em quest�o � igual ao m�s do dado mais recente
    If Month(Sheets("EXPECTATIVAS - DEMANDA").Cells(8, Expectativas_Demanda_Coluna_DadosP)) = Month(Sheets("EXPECTATIVAS - DEMANDA").Cells(8, Expectativas_Demanda_Coluna_Dados1)) Then
    'Caso seja igual, copia a coluna com os dados mais abaixo, a partir da linha 110
        Sheets("EXPECTATIVAS - DEMANDA").Range(Cells(9, Expectativas_Demanda_Coluna_DadosP), (Cells(54, Expectativas_Demanda_Coluna_DadosP))).Copy (Sheets("EXPECTATIVAS - DEMANDA").Cells(110, Expectativas_Demanda_Coluna_DadosP))
    End If
    'Vai para a pr�xima coluna
    Expectativas_Demanda_Coluna_DadosP = Expectativas_Demanda_Coluna_DadosP + 1
'Repete a conferencia com a pr�xima coluna
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Expectativas_Demanda_Linha_Dados = 110
Expectativas_Demanda_Coluna_Dados1 = Sheets("EXPECTATIVAS - DEMANDA").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Expectativas_Demanda_Linha_An�lise = 59
Expectativas_Demanda_Coluna_An�lise = 7

'Ordena��o decrescente da s�rie hist�rica dos meses do dado mais recente
Do Until Expectativas_Demanda_Linha_Dados = 156 'Faz o calculo at� a vari�vel Linha_Dados ser 156
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo m�s do mais recente
    posi��o = Application.WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVAS - DEMANDA").Cells(Expectativas_Demanda_Linha_Dados, Expectativas_Demanda_Coluna_Dados1), Sheets("EXPECTATIVAS - DEMANDA").Range(Cells(Expectativas_Demanda_Linha_Dados, Expectativas_Demanda_Coluna_Dados1), Cells(Expectativas_Demanda_Linha_Dados, 2)))
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("EXPECTATIVAS - DEMANDA").Cells(Expectativas_Demanda_Linha_An�lise, Expectativas_Demanda_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Expectativas_Demanda_Linha_Dados = Expectativas_Demanda_Linha_Dados + 1
    Expectativas_Demanda_Linha_An�lise = Expectativas_Demanda_Linha_An�lise + 1
'Repete a ordena��o com a pr�xima linha
Loop

'Atribui os valores originais das variaveis ap�s o loop anterior
Expectativas_Demanda_Linha_Dados = 110
Expectativas_Demanda_Coluna_Dados1 = Sheets("EXPECTATIVAS - DEMANDA").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Expectativas_Demanda_Linha_An�lise = 59
Expectativas_Demanda_Coluna_An�lise = 8
'Ordena��o crescente da s�rie hist�rica dos meses do dado mais recente
Do Until Expectativas_Demanda_Linha_Dados = 156 'Faz o calculo at� a vari�vel Linha_Dados ser 156
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo m�s do mais recente
    posi��o = Application.WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVAS - DEMANDA").Cells(Expectativas_Demanda_Linha_Dados, Expectativas_Demanda_Coluna_Dados1), Sheets("EXPECTATIVAS - DEMANDA").Range(Cells(Expectativas_Demanda_Linha_Dados, Expectativas_Demanda_Coluna_Dados1), Cells(Expectativas_Demanda_Linha_Dados, 2)), 1)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("EXPECTATIVAS - DEMANDA").Cells(Expectativas_Demanda_Linha_An�lise, Expectativas_Demanda_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Expectativas_Demanda_Linha_Dados = Expectativas_Demanda_Linha_Dados + 1
    Expectativas_Demanda_Linha_An�lise = Expectativas_Demanda_Linha_An�lise + 1
'Repete a ordena��o com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Expectativas_Demanda_Linha_Dados = 9
Expectativas_Demanda_Coluna_Dados1 = Sheets("EXPECTATIVAS - DEMANDA").Range("A9").End(xlToRight).Column
Expectativas_Demanda_Coluna_Dados2 = Expectativas_Demanda_Coluna_Dados1 - 1
Expectativas_Demanda_Linha_An�lise = 59
Expectativas_Demanda_Coluna_An�lise = 9

'Avalia��o se cruzou ou n�o a linha de 50 e o sentido
Do Until Expectativas_Demanda_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'se o dado mais recente for menor que 50 e o dado do m�s anterior for maior ou igual a 50 ent�o...
    If Cells(Expectativas_Demanda_Linha_Dados, Expectativas_Demanda_Coluna_Dados1) < 50 And Cells(Expectativas_Demanda_Linha_Dados, Expectativas_Demanda_Coluna_Dados2) >= 50 Then
    
    'a c�lula de an�lise recebe cruzou para baixo
    Cells(Expectativas_Demanda_Linha_An�lise, Expectativas_Demanda_Coluna_An�lise).Value = "Cruzou para baixo"
    'Caso n�o seja..
    Else
        'se o dado mais recente for maior ou igual a 50 e o dado do m~es anterior for menor ou igual a 50 ent�o...
        If Cells(Expectativas_Demanda_Linha_Dados, Expectativas_Demanda_Coluna_Dados1) >= 50 And Cells(Expectativas_Demanda_Linha_Dados, Expectativas_Demanda_Coluna_Dados2) <= 50 Then
        'a c�lula de an�lise recebe cruzou para cima
        Cells(Expectativas_Demanda_Linha_An�lise, Expectativas_Demanda_Coluna_An�lise).Value = "Cruzou para cima"
        'Caso n�o seja..
        Else
        'a c�lula de an�lise recebe n�o cruzou
        Cells(Expectativas_Demanda_Linha_An�lise, Expectativas_Demanda_Coluna_An�lise).Value = "N�o Cruzou"
        End If
    End If
    'Vai para a pr�xima linha de dados e de an�lise
    Expectativas_Demanda_Linha_Dados = Expectativas_Demanda_Linha_Dados + 1
    Expectativas_Demanda_Linha_An�lise = Expectativas_Demanda_Linha_An�lise + 1
'repete o processo com a nova linha
Loop

'Apaga as linhas com erros/dados faltantes/t�tulos e subt�tulos
Range(Cells(60, 2), Cells(60, 9)).ClearContents
Range(Cells(66, 2), Cells(66, 9)).ClearContents
Range(Cells(70, 2), Cells(70, 9)).ClearContents
Range(Cells(72, 2), Cells(73, 9)).Value = "-"
Range(Cells(75, 2), Cells(75, 9)).Value = "-"
Range(Cells(79, 2), Cells(79, 9)).Value = "-"
Range(Cells(87, 2), Cells(87, 9)).Value = "-"

'Inserindo valores nas celulas vazias para fugir de bugs
Range(Cells(10, Expectativas_Demanda_Coluna_Dados3), Cells(10, Expectativas_Demanda_Coluna_Dados1)).ClearContents
Range(Cells(16, Expectativas_Demanda_Coluna_Dados3), Cells(16, Expectativas_Demanda_Coluna_Dados1)).ClearContents
Range(Cells(20, Expectativas_Demanda_Coluna_Dados3), Cells(20, Expectativas_Demanda_Coluna_Dados1)).ClearContents
Range(Cells(22, Expectativas_Demanda_Coluna_Dados3), Cells(23, Expectativas_Demanda_Coluna_Dados1)).Value = "-"
Range(Cells(25, Expectativas_Demanda_Coluna_Dados3), Cells(25, Expectativas_Demanda_Coluna_Dados1)).Value = "-"
Range(Cells(29, Expectativas_Demanda_Coluna_Dados3), Cells(29, Expectativas_Demanda_Coluna_Dados1)).Value = "-"
Range(Cells(37, Expectativas_Demanda_Coluna_Dados3), Cells(37, Expectativas_Demanda_Coluna_Dados1)).Value = "-"

Range("E59:H104").NumberFormat = "0"


'******************************************             Expectativas_Exporta��o            *********************************************************

Dim Expectativas_Exporta��o_Linha_Dados As Integer 'Define a linha que cont�m o dado a ser usado
Dim Expectativas_Exporta��o_Coluna_Dados1 As Integer ' Define a coluna com o dado mais recente
Dim Expectativas_Exporta��o_Coluna_Dados2 As Integer ' Define a coluna com o dado do m�s anterior
Dim Expectativas_Exporta��o_Coluna_Dados3 As Integer ' Defie a coluna com o dado do mesmo m�s do ano anterior
Dim Expectativas_Exporta��o_Linha_An�lise As Integer ' Define a linha que ser� feita a an�lise
Dim Expectativas_Exporta��o_Coluna_An�lise As Integer 'Define a coluna que ser� feita a an�lise

Sheets("EXPECTATIVA - EXPORTA��O").Select


'Copia os t�tulos das categorias e cola onde ser� formada a tabela de an�lise
Sheets("EXPECTATIVA - EXPORTA��O").Range(Cells(9, 1), Cells(54, 9)).Copy (Sheets("EXPECTATIVA - EXPORTA��O").Cells(59, 1))
'Limpa os n�meros que foram colados mas mant�m a formata��o
Sheets("EXPECTATIVA - EXPORTA��O").Range(Cells(59, 2), Cells(105, 9)).ClearContents



'Nomeia as colunas de acordo com o dado que ser� calculado nelas
Sheets("EXPECTATIVA - EXPORTA��O").Cells(58, 2).Value = "Diferen�a para o m�s anterior"
Sheets("EXPECTATIVA - EXPORTA��O").Cells(58, 3).Value = "Diferen�a para ao mesmo m�s do ano anterior"
Sheets("EXPECTATIVA - EXPORTA��O").Cells(58, 4).Value = "Diferen�a para a m�dia hist�rica"
Sheets("EXPECTATIVA - EXPORTA��O").Cells(58, 5).Value = "Posi��o Decrescente (Maior valor 1�, menor valor �ltimo)"
Sheets("EXPECTATIVA - EXPORTA��O").Cells(58, 6).Value = "Posi��o Crescente  (Menor valor 1�, maior valor �ltimo)"
Sheets("EXPECTATIVA - EXPORTA��O").Cells(58, 7).Value = "Posi��o Crescente - Mesmo m�s  (Menor valor 1�, maior valor �ltimo)"
Sheets("EXPECTATIVA - EXPORTA��O").Cells(58, 8).Value = "Posi��o Decrescente -Mesmo m�s  (Maior valor 1�, menor valor �ltimo)"
Sheets("EXPECTATIVA - EXPORTA��O").Cells(58, 9).Value = "cruzou a linha de 50?"

'Atribui valores as variaveis definidas acima
Expectativas_Exporta��o_Linha_Dados = 9 'Define o n�mero da primeira linha de dados
Expectativas_Exporta��o_Coluna_Dados1 = Sheets("EXPECTATIVA - EXPORTA��O").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Expectativas_Exporta��o_Coluna_Dados2 = Expectativas_Exporta��o_Coluna_Dados1 - 1 'Define o n�mero da coluna do m�s anterior
Expectativas_Exporta��o_Coluna_Dados3 = Expectativas_Exporta��o_Coluna_Dados1 - 12
Expectativas_Exporta��o_Linha_An�lise = 59 'Define a primeira linhas de an�lises
Expectativas_Exporta��o_Coluna_An�lise = 2 'Define a coluna de an�lises

'Inserindo valores nas celulas vazias para fugir de bugs
Sheets("EXPECTATIVA - EXPORTA��O").Range(Cells(10, Expectativas_Exporta��o_Coluna_Dados3), Cells(10, Expectativas_Exporta��o_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - EXPORTA��O").Range(Cells(16, Expectativas_Exporta��o_Coluna_Dados3), Cells(16, Expectativas_Exporta��o_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - EXPORTA��O").Range(Cells(20, Expectativas_Exporta��o_Coluna_Dados3), Cells(20, Expectativas_Exporta��o_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - EXPORTA��O").Range(Cells(22, Expectativas_Exporta��o_Coluna_Dados3), Cells(23, Expectativas_Exporta��o_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - EXPORTA��O").Range(Cells(25, Expectativas_Exporta��o_Coluna_Dados3), Cells(25, Expectativas_Exporta��o_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - EXPORTA��O").Range(Cells(29, Expectativas_Exporta��o_Coluna_Dados3), Cells(29, Expectativas_Exporta��o_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - EXPORTA��O").Range(Cells(37, Expectativas_Exporta��o_Coluna_Dados3), Cells(37, Expectativas_Exporta��o_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - EXPORTA��O").Range(Cells(54, Expectativas_Exporta��o_Coluna_Dados3), Cells(54, Expectativas_Exporta��o_Coluna_Dados1)).Value = "0"

'Calculo da difern�a em pontos do valor mais recente em rela��o ao valor do m�s anterior
Do Until Expectativas_Exporta��o_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Calculo da diferen�a em si: o valor da celula de analise � igual ao valor mais recente menos o valor do m�s anterior
   Sheets("EXPECTATIVA - EXPORTA��O").Cells(Expectativas_Exporta��o_Linha_An�lise, Expectativas_Exporta��o_Coluna_An�lise).Value = Sheets("EXPECTATIVA - EXPORTA��O").Cells(Expectativas_Exporta��o_Linha_Dados, Expectativas_Exporta��o_Coluna_Dados1).Value - Sheets("EXPECTATIVA - EXPORTA��O").Cells(Expectativas_Exporta��o_Linha_Dados, Expectativas_Exporta��o_Coluna_Dados2).Value
    'Vai para a pr�xima linha de dados e de an�lise
   Expectativas_Exporta��o_Linha_Dados = Expectativas_Exporta��o_Linha_Dados + 1
   Expectativas_Exporta��o_Linha_An�lise = Expectativas_Exporta��o_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Expectativas_Exporta��o_Linha_Dados = 9
Expectativas_Exporta��o_Coluna_Dados1 = Sheets("EXPECTATIVA - EXPORTA��O").Range("A9").End(xlToRight).Column
Expectativas_Exporta��o_Coluna_Dados3 = Expectativas_Exporta��o_Coluna_Dados1 - 12
Expectativas_Exporta��o_Linha_An�lise = 59
Expectativas_Exporta��o_Coluna_An�lise = 3

'C�lculo da diferen�a em pontos do valor mais recente em rela��o ao valor do mesmo m�s do ano anterior
Do Until Expectativas_Exporta��o_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Calculo da diferen�a em si: o valor da celula de analise � igual ao valor mais recente menos o valor do mesmo m�s do ano anterior
    Sheets("EXPECTATIVA - EXPORTA��O").Cells(Expectativas_Exporta��o_Linha_An�lise, Expectativas_Exporta��o_Coluna_An�lise).Value = Sheets("EXPECTATIVA - EXPORTA��O").Cells(Expectativas_Exporta��o_Linha_Dados, Expectativas_Exporta��o_Coluna_Dados1).Value - Sheets("EXPECTATIVA - EXPORTA��O").Cells(Expectativas_Exporta��o_Linha_Dados, Expectativas_Exporta��o_Coluna_Dados3).Value
    'Vai para a pr�xima linha de dados e de an�lise
    Expectativas_Exporta��o_Linha_Dados = Expectativas_Exporta��o_Linha_Dados + 1
    Expectativas_Exporta��o_Linha_An�lise = Expectativas_Exporta��o_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Expectativas_Exporta��o_Linha_Dados = 9
Expectativas_Exporta��o_Coluna_Dados1 = Sheets("EXPECTATIVA - EXPORTA��O").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Expectativas_Exporta��o_Linha_An�lise = 59
Expectativas_Exporta��o_Coluna_An�lise = 4

'C�lculo da diferen�a em pontos do valor mais recente em rela��o ao valor da m�dia hist�rica
Do Until Expectativas_Exporta��o_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a vari�vel media como a m�dia do intervalo entre a coluna com o dado mais recente e o primeiro
    media = Application.Average(Sheets("EXPECTATIVA - EXPORTA��O").Range(Cells(Expectativas_Exporta��o_Linha_Dados, Expectativas_Exporta��o_Coluna_Dados1), Cells(Expectativas_Exporta��o_Linha_Dados, 2)))
    'Calculo em si: o valor da celula de analise � igual ao valor mais recente menos o valor da m�dia
    Sheets("EXPECTATIVA - EXPORTA��O").Cells(Expectativas_Exporta��o_Linha_An�lise, Expectativas_Exporta��o_Coluna_An�lise).Value = Sheets("EXPECTATIVA - EXPORTA��O").Cells(Expectativas_Exporta��o_Linha_Dados, Expectativas_Exporta��o_Coluna_Dados1).Value - media
    'Vai para a pr�xima linha de dados e de an�lise
    Expectativas_Exporta��o_Linha_Dados = Expectativas_Exporta��o_Linha_Dados + 1
    Expectativas_Exporta��o_Linha_An�lise = Expectativas_Exporta��o_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Expectativas_Exporta��o_Linha_Dados = 9
Expectativas_Exporta��o_Coluna_Dados1 = Sheets("EXPECTATIVA - EXPORTA��O").Range("A9").End(xlToRight).Column
Expectativas_Exporta��o_Linha_An�lise = 59
Expectativas_Exporta��o_Coluna_An�lise = 5

'Ordena��o decrescente da s�rie hist�rica completa
Do Until Expectativas_Exporta��o_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posi��o = WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVA - EXPORTA��O").Cells(Expectativas_Exporta��o_Linha_Dados, Expectativas_Exporta��o_Coluna_Dados1), Sheets("EXPECTATIVA - EXPORTA��O").Range(Cells(Expectativas_Exporta��o_Linha_Dados, Expectativas_Exporta��o_Coluna_Dados1), Cells(Expectativas_Exporta��o_Linha_Dados, 2)), 0)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("EXPECTATIVA - EXPORTA��O").Cells(Expectativas_Exporta��o_Linha_An�lise, Expectativas_Exporta��o_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Expectativas_Exporta��o_Linha_Dados = Expectativas_Exporta��o_Linha_Dados + 1
    Expectativas_Exporta��o_Linha_An�lise = Expectativas_Exporta��o_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Expectativas_Exporta��o_Linha_Dados = 9
Expectativas_Exporta��o_Coluna_Dados1 = Sheets("EXPECTATIVA - EXPORTA��O").Range("A9").End(xlToRight).Column
Expectativas_Exporta��o_Linha_An�lise = 59
Expectativas_Exporta��o_Coluna_An�lise = 6

'Ordena��o Crescente da s�rie hist�rica completa
Do Until Expectativas_Exporta��o_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posi��o = WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVA - EXPORTA��O").Cells(Expectativas_Exporta��o_Linha_Dados, Expectativas_Exporta��o_Coluna_Dados1), Sheets("EXPECTATIVA - EXPORTA��O").Range(Cells(Expectativas_Exporta��o_Linha_Dados, Expectativas_Exporta��o_Coluna_Dados1), Cells(Expectativas_Exporta��o_Linha_Dados, 2)), 1)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("EXPECTATIVA - EXPORTA��O").Cells(Expectativas_Exporta��o_Linha_An�lise, Expectativas_Exporta��o_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Expectativas_Exporta��o_Linha_Dados = Expectativas_Exporta��o_Linha_Dados + 1
    Expectativas_Exporta��o_Linha_An�lise = Expectativas_Exporta��o_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior e define a vari�vel Coluna_DadosP que representa a primeira coluna de dados
Expectativas_Exporta��o_Coluna_Dados1 = Sheets("EXPECTATIVA - EXPORTA��O").Range("A9").End(xlToRight).Column
Expectativas_Exporta��o_Coluna_DadosP = 2

Do Until Expectativas_Exporta��o_Coluna_DadosP = Expectativas_Exporta��o_Coluna_Dados1 + 1 ' Faz at� a variavel Coluna_DadosP ser igual a variavel Coluna_Dados1 mais uma unidade
    'Confere se o m�s da coluna em quest�o � igual ao m�s do dado mais recente
    If Month(Sheets("EXPECTATIVA - EXPORTA��O").Cells(8, Expectativas_Exporta��o_Coluna_DadosP)) = Month(Sheets("EXPECTATIVA - EXPORTA��O").Cells(8, Expectativas_Exporta��o_Coluna_Dados1)) Then
    'Caso seja igual, copia a coluna com os dados mais abaixo, a partir da linha 110
        Sheets("EXPECTATIVA - EXPORTA��O").Range(Cells(9, Expectativas_Exporta��o_Coluna_DadosP), (Cells(54, Expectativas_Exporta��o_Coluna_DadosP))).Copy (Sheets("EXPECTATIVA - EXPORTA��O").Cells(110, Expectativas_Exporta��o_Coluna_DadosP))
    End If
    'Vai para a pr�xima coluna
    Expectativas_Exporta��o_Coluna_DadosP = Expectativas_Exporta��o_Coluna_DadosP + 1
'Repete a conferencia com a pr�xima coluna
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Expectativas_Exporta��o_Linha_Dados = 110
Expectativas_Exporta��o_Coluna_Dados1 = Sheets("EXPECTATIVA - EXPORTA��O").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Expectativas_Exporta��o_Linha_An�lise = 59
Expectativas_Exporta��o_Coluna_An�lise = 7

'Ordena��o decrescente da s�rie hist�rica dos meses do dado mais recente
Do Until Expectativas_Exporta��o_Linha_Dados = 156 'Faz o calculo at� a vari�vel Linha_Dados ser 156
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo m�s do mais recente
    posi��o = Application.WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVA - EXPORTA��O").Cells(Expectativas_Exporta��o_Linha_Dados, Expectativas_Exporta��o_Coluna_Dados1), Sheets("EXPECTATIVA - EXPORTA��O").Range(Cells(Expectativas_Exporta��o_Linha_Dados, Expectativas_Exporta��o_Coluna_Dados1), Cells(Expectativas_Exporta��o_Linha_Dados, 2)))
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("EXPECTATIVA - EXPORTA��O").Cells(Expectativas_Exporta��o_Linha_An�lise, Expectativas_Exporta��o_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Expectativas_Exporta��o_Linha_Dados = Expectativas_Exporta��o_Linha_Dados + 1
    Expectativas_Exporta��o_Linha_An�lise = Expectativas_Exporta��o_Linha_An�lise + 1
'Repete a ordena��o com a pr�xima linha
Loop

'Atribui os valores originais das variaveis ap�s o loop anterior
Expectativas_Exporta��o_Linha_Dados = 110
Expectativas_Exporta��o_Coluna_Dados1 = Sheets("EXPECTATIVA - EXPORTA��O").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Expectativas_Exporta��o_Linha_An�lise = 59
Expectativas_Exporta��o_Coluna_An�lise = 8
'Ordena��o crescente da s�rie hist�rica dos meses do dado mais recente
Do Until Expectativas_Exporta��o_Linha_Dados = 156 'Faz o calculo at� a vari�vel Linha_Dados ser 156
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo m�s do mais recente
    posi��o = Application.WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVA - EXPORTA��O").Cells(Expectativas_Exporta��o_Linha_Dados, Expectativas_Exporta��o_Coluna_Dados1), Sheets("EXPECTATIVA - EXPORTA��O").Range(Cells(Expectativas_Exporta��o_Linha_Dados, Expectativas_Exporta��o_Coluna_Dados1), Cells(Expectativas_Exporta��o_Linha_Dados, 2)), 1)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("EXPECTATIVA - EXPORTA��O").Cells(Expectativas_Exporta��o_Linha_An�lise, Expectativas_Exporta��o_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Expectativas_Exporta��o_Linha_Dados = Expectativas_Exporta��o_Linha_Dados + 1
    Expectativas_Exporta��o_Linha_An�lise = Expectativas_Exporta��o_Linha_An�lise + 1
'Repete a ordena��o com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Expectativas_Exporta��o_Linha_Dados = 9
Expectativas_Exporta��o_Coluna_Dados1 = Sheets("EXPECTATIVA - EXPORTA��O").Range("A9").End(xlToRight).Column
Expectativas_Exporta��o_Coluna_Dados2 = Expectativas_Exporta��o_Coluna_Dados1 - 1
Expectativas_Exporta��o_Linha_An�lise = 59
Expectativas_Exporta��o_Coluna_An�lise = 9

'Avalia��o se cruzou ou n�o a linha de 50 e o sentido
Do Until Expectativas_Exporta��o_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'se o dado mais recente for menor que 50 e o dado do m�s anterior for maior ou igual a 50 ent�o...
    If Cells(Expectativas_Exporta��o_Linha_Dados, Expectativas_Exporta��o_Coluna_Dados1) < 50 And Cells(Expectativas_Exporta��o_Linha_Dados, Expectativas_Exporta��o_Coluna_Dados2) >= 50 Then
    
    'a c�lula de an�lise recebe cruzou para baixo
    Cells(Expectativas_Exporta��o_Linha_An�lise, Expectativas_Exporta��o_Coluna_An�lise).Value = "Cruzou para baixo"
    'Caso n�o seja..
    Else
        'se o dado mais recente for maior ou igual a 50 e o dado do m~es anterior for menor ou igual a 50 ent�o...
        If Cells(Expectativas_Exporta��o_Linha_Dados, Expectativas_Exporta��o_Coluna_Dados1) >= 50 And Cells(Expectativas_Exporta��o_Linha_Dados, Expectativas_Exporta��o_Coluna_Dados2) <= 50 Then
        'a c�lula de an�lise recebe cruzou para cima
        Cells(Expectativas_Exporta��o_Linha_An�lise, Expectativas_Exporta��o_Coluna_An�lise).Value = "Cruzou para cima"
        'Caso n�o seja..
        Else
        'a c�lula de an�lise recebe n�o cruzou
        Cells(Expectativas_Exporta��o_Linha_An�lise, Expectativas_Exporta��o_Coluna_An�lise).Value = "N�o Cruzou"
        End If
    End If
    'Vai para a pr�xima linha de dados e de an�lise
    Expectativas_Exporta��o_Linha_Dados = Expectativas_Exporta��o_Linha_Dados + 1
    Expectativas_Exporta��o_Linha_An�lise = Expectativas_Exporta��o_Linha_An�lise + 1
'repete o processo com a nova linha
Loop

'Apaga as linhas com erros/dados faltantes/t�tulos e subt�tulos
Range(Cells(60, 2), Cells(60, 9)).ClearContents
Range(Cells(66, 2), Cells(66, 9)).ClearContents
Range(Cells(70, 2), Cells(70, 9)).ClearContents
Range(Cells(72, 2), Cells(73, 9)).Value = "-"
Range(Cells(75, 2), Cells(75, 9)).Value = "-"
Range(Cells(79, 2), Cells(79, 9)).Value = "-"
Range(Cells(87, 2), Cells(87, 9)).Value = "-"
Range(Cells(104, 2), Cells(104, 9)).Value = "-"

'Inserindo valores nas celulas vazias para fugir de bugs
Range(Cells(10, Expectativas_Exporta��o_Coluna_Dados3), Cells(10, Expectativas_Exporta��o_Coluna_Dados1)).ClearContents
Range(Cells(16, Expectativas_Exporta��o_Coluna_Dados3), Cells(16, Expectativas_Exporta��o_Coluna_Dados1)).ClearContents
Range(Cells(20, Expectativas_Exporta��o_Coluna_Dados3), Cells(20, Expectativas_Exporta��o_Coluna_Dados1)).ClearContents
Range(Cells(22, Expectativas_Exporta��o_Coluna_Dados3), Cells(23, Expectativas_Exporta��o_Coluna_Dados1)).Value = "-"
Range(Cells(25, Expectativas_Exporta��o_Coluna_Dados3), Cells(25, Expectativas_Exporta��o_Coluna_Dados1)).Value = "-"
Range(Cells(29, Expectativas_Exporta��o_Coluna_Dados3), Cells(29, Expectativas_Exporta��o_Coluna_Dados1)).Value = "-"
Range(Cells(37, Expectativas_Exporta��o_Coluna_Dados3), Cells(37, Expectativas_Exporta��o_Coluna_Dados1)).Value = "-"
Range(Cells(54, Expectativas_Exporta��o_Coluna_Dados3), Cells(54, Expectativas_Exporta��o_Coluna_Dados1)).Value = "-"


Range("E59:H104").NumberFormat = "0"


'****************************************      Expectativa_Compras                ***********************************************************/


Dim Expectativas_Compras_Linha_Dados As Integer 'Define a linha que cont�m o dado a ser usado
Dim Expectativas_Compras_Coluna_Dados1 As Integer ' Define a coluna com o dado mais recente
Dim Expectativas_Compras_Coluna_Dados2 As Integer ' Define a coluna com o dado do m�s anterior
Dim Expectativas_Compras_Coluna_Dados3 As Integer ' Defie a coluna com o dado do mesmo m�s do ano anterior
Dim Expectativas_Compras_Linha_An�lise As Integer ' Define a linha que ser� feita a an�lise
Dim Expectativas_Compras_Coluna_An�lise As Integer 'Define a coluna que ser� feita a an�lise


Sheets("EXPECTATIVA - COMPRAS").Select



'Copia os t�tulos das categorias e cola onde ser� formada a tabela de an�lise
Sheets("EXPECTATIVA - COMPRAS").Range(Cells(9, 1), Cells(54, 9)).Copy (Sheets("EXPECTATIVA - COMPRAS").Cells(59, 1))
'Limpa os n�meros que foram colados mas mant�m a formata��o
Sheets("EXPECTATIVA - COMPRAS").Range(Cells(59, 2), Cells(105, 9)).ClearContents



'Nomeia as colunas de acordo com o dado que ser� calculado nelas
Sheets("EXPECTATIVA - COMPRAS").Cells(58, 2).Value = "Diferen�a para o m�s anterior"
Sheets("EXPECTATIVA - COMPRAS").Cells(58, 3).Value = "Diferen�a para ao mesmo m�s do ano anterior"
Sheets("EXPECTATIVA - COMPRAS").Cells(58, 4).Value = "Diferen�a para a m�dia hist�rica"
Sheets("EXPECTATIVA - COMPRAS").Cells(58, 5).Value = "Posi��o Decrescente (Maior valor 1�, menor valor �ltimo)"
Sheets("EXPECTATIVA - COMPRAS").Cells(58, 6).Value = "Posi��o Crescente  (Menor valor 1�, maior valor �ltimo)"
Sheets("EXPECTATIVA - COMPRAS").Cells(58, 7).Value = "Posi��o Crescente - Mesmo m�s  (Menor valor 1�, maior valor �ltimo)"
Sheets("EXPECTATIVA - COMPRAS").Cells(58, 8).Value = "Posi��o Decrescente -Mesmo m�s  (Maior valor 1�, menor valor �ltimo)"
Sheets("EXPECTATIVA - COMPRAS").Cells(58, 9).Value = "cruzou a linha de 50?"

'Atribui valores as variaveis definidas acima
Expectativas_Compras_Linha_Dados = 9 'Define o n�mero da primeira linha de dados
Expectativas_Compras_Coluna_Dados1 = Sheets("EXPECTATIVA - COMPRAS").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Expectativas_Compras_Coluna_Dados2 = Expectativas_Compras_Coluna_Dados1 - 1 'Define o n�mero da coluna do m�s anterior
Expectativas_Compras_Coluna_Dados3 = Expectativas_Compras_Coluna_Dados1 - 12
Expectativas_Compras_Linha_An�lise = 59 'Define a primeira linhas de an�lises
Expectativas_Compras_Coluna_An�lise = 2 'Define a coluna de an�lises

'Inserindo valores nas celulas vazias para fugir de bugs
Sheets("EXPECTATIVA - COMPRAS").Range(Cells(10, Expectativas_Compras_Coluna_Dados3), Cells(10, Expectativas_Compras_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - COMPRAS").Range(Cells(16, Expectativas_Compras_Coluna_Dados3), Cells(16, Expectativas_Compras_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - COMPRAS").Range(Cells(20, Expectativas_Compras_Coluna_Dados3), Cells(20, Expectativas_Compras_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - COMPRAS").Range(Cells(22, Expectativas_Compras_Coluna_Dados3), Cells(23, Expectativas_Compras_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - COMPRAS").Range(Cells(25, Expectativas_Compras_Coluna_Dados3), Cells(25, Expectativas_Compras_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - COMPRAS").Range(Cells(29, Expectativas_Compras_Coluna_Dados3), Cells(29, Expectativas_Compras_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - COMPRAS").Range(Cells(37, Expectativas_Compras_Coluna_Dados3), Cells(37, Expectativas_Compras_Coluna_Dados1)).Value = "0"


'Calculo da difern�a em pontos do valor mais recente em rela��o ao valor do m�s anterior
Do Until Expectativas_Compras_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Calculo da diferen�a em si: o valor da celula de analise � igual ao valor mais recente menos o valor do m�s anterior
   Sheets("EXPECTATIVA - COMPRAS").Cells(Expectativas_Compras_Linha_An�lise, Expectativas_Compras_Coluna_An�lise).Value = Sheets("EXPECTATIVA - COMPRAS").Cells(Expectativas_Compras_Linha_Dados, Expectativas_Compras_Coluna_Dados1).Value - Sheets("EXPECTATIVA - COMPRAS").Cells(Expectativas_Compras_Linha_Dados, Expectativas_Compras_Coluna_Dados2).Value
    'Vai para a pr�xima linha de dados e de an�lise
   Expectativas_Compras_Linha_Dados = Expectativas_Compras_Linha_Dados + 1
   Expectativas_Compras_Linha_An�lise = Expectativas_Compras_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Expectativas_Compras_Linha_Dados = 9
Expectativas_Compras_Coluna_Dados1 = Sheets("EXPECTATIVA - COMPRAS").Range("A9").End(xlToRight).Column
Expectativas_Compras_Coluna_Dados3 = Expectativas_Compras_Coluna_Dados1 - 12
Expectativas_Compras_Linha_An�lise = 59
Expectativas_Compras_Coluna_An�lise = 3

'C�lculo da diferen�a em pontos do valor mais recente em rela��o ao valor do mesmo m�s do ano anterior
Do Until Expectativas_Compras_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Calculo da diferen�a em si: o valor da celula de analise � igual ao valor mais recente menos o valor do mesmo m�s do ano anterior
    Sheets("EXPECTATIVA - COMPRAS").Cells(Expectativas_Compras_Linha_An�lise, Expectativas_Compras_Coluna_An�lise).Value = Sheets("EXPECTATIVA - COMPRAS").Cells(Expectativas_Compras_Linha_Dados, Expectativas_Compras_Coluna_Dados1).Value - Sheets("EXPECTATIVA - COMPRAS").Cells(Expectativas_Compras_Linha_Dados, Expectativas_Compras_Coluna_Dados3).Value
    'Vai para a pr�xima linha de dados e de an�lise
    Expectativas_Compras_Linha_Dados = Expectativas_Compras_Linha_Dados + 1
    Expectativas_Compras_Linha_An�lise = Expectativas_Compras_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Expectativas_Compras_Linha_Dados = 9
Expectativas_Compras_Coluna_Dados1 = Sheets("EXPECTATIVA - COMPRAS").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Expectativas_Compras_Linha_An�lise = 59
Expectativas_Compras_Coluna_An�lise = 4

'C�lculo da diferen�a em pontos do valor mais recente em rela��o ao valor da m�dia hist�rica
Do Until Expectativas_Compras_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a vari�vel media como a m�dia do intervalo entre a coluna com o dado mais recente e o primeiro
    media = Application.Average(Sheets("EXPECTATIVA - COMPRAS").Range(Cells(Expectativas_Compras_Linha_Dados, Expectativas_Compras_Coluna_Dados1), Cells(Expectativas_Compras_Linha_Dados, 2)))
    'Calculo em si: o valor da celula de analise � igual ao valor mais recente menos o valor da m�dia
    Sheets("EXPECTATIVA - COMPRAS").Cells(Expectativas_Compras_Linha_An�lise, Expectativas_Compras_Coluna_An�lise).Value = Sheets("EXPECTATIVA - COMPRAS").Cells(Expectativas_Compras_Linha_Dados, Expectativas_Compras_Coluna_Dados1).Value - media
    'Vai para a pr�xima linha de dados e de an�lise
    Expectativas_Compras_Linha_Dados = Expectativas_Compras_Linha_Dados + 1
    Expectativas_Compras_Linha_An�lise = Expectativas_Compras_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Expectativas_Compras_Linha_Dados = 9
Expectativas_Compras_Coluna_Dados1 = Sheets("EXPECTATIVA - COMPRAS").Range("A9").End(xlToRight).Column
Expectativas_Compras_Linha_An�lise = 59
Expectativas_Compras_Coluna_An�lise = 5

'Ordena��o decrescente da s�rie hist�rica completa
Do Until Expectativas_Compras_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posi��o = WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVA - COMPRAS").Cells(Expectativas_Compras_Linha_Dados, Expectativas_Compras_Coluna_Dados1), Sheets("EXPECTATIVA - COMPRAS").Range(Cells(Expectativas_Compras_Linha_Dados, Expectativas_Compras_Coluna_Dados1), Cells(Expectativas_Compras_Linha_Dados, 2)), 0)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("EXPECTATIVA - COMPRAS").Cells(Expectativas_Compras_Linha_An�lise, Expectativas_Compras_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Expectativas_Compras_Linha_Dados = Expectativas_Compras_Linha_Dados + 1
    Expectativas_Compras_Linha_An�lise = Expectativas_Compras_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Expectativas_Compras_Linha_Dados = 9
Expectativas_Compras_Coluna_Dados1 = Sheets("EXPECTATIVA - COMPRAS").Range("A9").End(xlToRight).Column
Expectativas_Compras_Linha_An�lise = 59
Expectativas_Compras_Coluna_An�lise = 6

'Ordena��o Crescente da s�rie hist�rica completa
Do Until Expectativas_Compras_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posi��o = WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVA - COMPRAS").Cells(Expectativas_Compras_Linha_Dados, Expectativas_Compras_Coluna_Dados1), Sheets("EXPECTATIVA - COMPRAS").Range(Cells(Expectativas_Compras_Linha_Dados, Expectativas_Compras_Coluna_Dados1), Cells(Expectativas_Compras_Linha_Dados, 2)), 1)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("EXPECTATIVA - COMPRAS").Cells(Expectativas_Compras_Linha_An�lise, Expectativas_Compras_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Expectativas_Compras_Linha_Dados = Expectativas_Compras_Linha_Dados + 1
    Expectativas_Compras_Linha_An�lise = Expectativas_Compras_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior e define a vari�vel Coluna_DadosP que representa a primeira coluna de dados
Expectativas_Compras_Coluna_Dados1 = Sheets("EXPECTATIVA - COMPRAS").Range("A9").End(xlToRight).Column
Expectativas_Compras_Coluna_DadosP = 2

Do Until Expectativas_Compras_Coluna_DadosP = Expectativas_Compras_Coluna_Dados1 + 1 ' Faz at� a variavel Coluna_DadosP ser igual a variavel Coluna_Dados1 mais uma unidade
    'Confere se o m�s da coluna em quest�o � igual ao m�s do dado mais recente
    If Month(Sheets("EXPECTATIVA - COMPRAS").Cells(8, Expectativas_Compras_Coluna_DadosP)) = Month(Sheets("EXPECTATIVA - COMPRAS").Cells(8, Expectativas_Compras_Coluna_Dados1)) Then
    'Caso seja igual, copia a coluna com os dados mais abaixo, a partir da linha 110
        Sheets("EXPECTATIVA - COMPRAS").Range(Cells(9, Expectativas_Compras_Coluna_DadosP), (Cells(54, Expectativas_Compras_Coluna_DadosP))).Copy (Sheets("EXPECTATIVA - COMPRAS").Cells(110, Expectativas_Compras_Coluna_DadosP))
    End If
    'Vai para a pr�xima coluna
    Expectativas_Compras_Coluna_DadosP = Expectativas_Compras_Coluna_DadosP + 1
'Repete a conferencia com a pr�xima coluna
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Expectativas_Compras_Linha_Dados = 110
Expectativas_Compras_Coluna_Dados1 = Sheets("EXPECTATIVA - COMPRAS").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Expectativas_Compras_Linha_An�lise = 59
Expectativas_Compras_Coluna_An�lise = 7

'Ordena��o decrescente da s�rie hist�rica dos meses do dado mais recente
Do Until Expectativas_Compras_Linha_Dados = 156 'Faz o calculo at� a vari�vel Linha_Dados ser 156
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo m�s do mais recente
    posi��o = Application.WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVA - COMPRAS").Cells(Expectativas_Compras_Linha_Dados, Expectativas_Compras_Coluna_Dados1), Sheets("EXPECTATIVA - COMPRAS").Range(Cells(Expectativas_Compras_Linha_Dados, Expectativas_Compras_Coluna_Dados1), Cells(Expectativas_Compras_Linha_Dados, 2)))
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("EXPECTATIVA - COMPRAS").Cells(Expectativas_Compras_Linha_An�lise, Expectativas_Compras_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Expectativas_Compras_Linha_Dados = Expectativas_Compras_Linha_Dados + 1
    Expectativas_Compras_Linha_An�lise = Expectativas_Compras_Linha_An�lise + 1
'Repete a ordena��o com a pr�xima linha
Loop

'Atribui os valores originais das variaveis ap�s o loop anterior
Expectativas_Compras_Linha_Dados = 110
Expectativas_Compras_Coluna_Dados1 = Sheets("EXPECTATIVA - COMPRAS").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Expectativas_Compras_Linha_An�lise = 59
Expectativas_Compras_Coluna_An�lise = 8
'Ordena��o crescente da s�rie hist�rica dos meses do dado mais recente
Do Until Expectativas_Compras_Linha_Dados = 156 'Faz o calculo at� a vari�vel Linha_Dados ser 156
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo m�s do mais recente
    posi��o = Application.WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVA - COMPRAS").Cells(Expectativas_Compras_Linha_Dados, Expectativas_Compras_Coluna_Dados1), Sheets("EXPECTATIVA - COMPRAS").Range(Cells(Expectativas_Compras_Linha_Dados, Expectativas_Compras_Coluna_Dados1), Cells(Expectativas_Compras_Linha_Dados, 2)), 1)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("EXPECTATIVA - COMPRAS").Cells(Expectativas_Compras_Linha_An�lise, Expectativas_Compras_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Expectativas_Compras_Linha_Dados = Expectativas_Compras_Linha_Dados + 1
    Expectativas_Compras_Linha_An�lise = Expectativas_Compras_Linha_An�lise + 1
'Repete a ordena��o com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Expectativas_Compras_Linha_Dados = 9
Expectativas_Compras_Coluna_Dados1 = Sheets("EXPECTATIVA - COMPRAS").Range("A9").End(xlToRight).Column
Expectativas_Compras_Coluna_Dados2 = Expectativas_Compras_Coluna_Dados1 - 1
Expectativas_Compras_Linha_An�lise = 59
Expectativas_Compras_Coluna_An�lise = 9

'Avalia��o se cruzou ou n�o a linha de 50 e o sentido
Do Until Expectativas_Compras_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'se o dado mais recente for menor que 50 e o dado do m�s anterior for maior ou igual a 50 ent�o...
    If Cells(Expectativas_Compras_Linha_Dados, Expectativas_Compras_Coluna_Dados1) < 50 And Cells(Expectativas_Compras_Linha_Dados, Expectativas_Compras_Coluna_Dados2) >= 50 Then
    
    'a c�lula de an�lise recebe cruzou para baixo
    Cells(Expectativas_Compras_Linha_An�lise, Expectativas_Compras_Coluna_An�lise).Value = "Cruzou para baixo"
    'Caso n�o seja..
    Else
        'se o dado mais recente for maior ou igual a 50 e o dado do m~es anterior for menor ou igual a 50 ent�o...
        If Cells(Expectativas_Compras_Linha_Dados, Expectativas_Compras_Coluna_Dados1) >= 50 And Cells(Expectativas_Compras_Linha_Dados, Expectativas_Compras_Coluna_Dados2) <= 50 Then
        'a c�lula de an�lise recebe cruzou para cima
        Cells(Expectativas_Compras_Linha_An�lise, Expectativas_Compras_Coluna_An�lise).Value = "Cruzou para cima"
        'Caso n�o seja..
        Else
        'a c�lula de an�lise recebe n�o cruzou
        Cells(Expectativas_Compras_Linha_An�lise, Expectativas_Compras_Coluna_An�lise).Value = "N�o Cruzou"
        End If
    End If
    'Vai para a pr�xima linha de dados e de an�lise
    Expectativas_Compras_Linha_Dados = Expectativas_Compras_Linha_Dados + 1
    Expectativas_Compras_Linha_An�lise = Expectativas_Compras_Linha_An�lise + 1
'repete o processo com a nova linha
Loop

'Apaga as linhas com erros/dados faltantes/t�tulos e subt�tulos
Range(Cells(60, 2), Cells(60, 9)).ClearContents
Range(Cells(66, 2), Cells(66, 9)).ClearContents
Range(Cells(70, 2), Cells(70, 9)).ClearContents
Range(Cells(72, 2), Cells(73, 9)).Value = "-"
Range(Cells(75, 2), Cells(75, 9)).Value = "-"
Range(Cells(79, 2), Cells(79, 9)).Value = "-"
Range(Cells(87, 2), Cells(87, 9)).Value = "-"

'Inserindo valores nas celulas vazias para fugir de bugs
Range(Cells(10, Expectativas_Compras_Coluna_Dados3), Cells(10, Expectativas_Compras_Coluna_Dados1)).ClearContents
Range(Cells(16, Expectativas_Compras_Coluna_Dados3), Cells(16, Expectativas_Compras_Coluna_Dados1)).ClearContents
Range(Cells(20, Expectativas_Compras_Coluna_Dados3), Cells(20, Expectativas_Compras_Coluna_Dados1)).ClearContents
Range(Cells(22, Expectativas_Compras_Coluna_Dados3), Cells(23, Expectativas_Compras_Coluna_Dados1)).Value = "-"
Range(Cells(25, Expectativas_Compras_Coluna_Dados3), Cells(25, Expectativas_Compras_Coluna_Dados1)).Value = "-"
Range(Cells(29, Expectativas_Compras_Coluna_Dados3), Cells(29, Expectativas_Compras_Coluna_Dados1)).Value = "-"
Range(Cells(37, Expectativas_Compras_Coluna_Dados3), Cells(37, Expectativas_Compras_Coluna_Dados1)).Value = "-"

Range("E59:H104").NumberFormat = "0"


'*********************************         Expectativa_Empregados              *********************************************************

Dim Expectativas_Empregados_Linha_Dados As Integer 'Define a linha que cont�m o dado a ser usado
Dim Expectativas_Empregados_Coluna_Dados1 As Integer ' Define a coluna com o dado mais recente
Dim Expectativas_Empregados_Coluna_Dados2 As Integer ' Define a coluna com o dado do m�s anterior
Dim Expectativas_Empregados_Coluna_Dados3 As Integer ' Defie a coluna com o dado do mesmo m�s do ano anterior
Dim Expectativas_Empregados_Linha_An�lise As Integer ' Define a linha que ser� feita a an�lise
Dim Expectativas_Empregados_Coluna_An�lise As Integer 'Define a coluna que ser� feita a an�lise

Sheets("EXPECTATIVA - EMPREGADOS").Select


'Copia os t�tulos das categorias e cola onde ser� formada a tabela de an�lise
Sheets("EXPECTATIVA - EMPREGADOS").Range(Cells(9, 1), Cells(54, 9)).Copy (Sheets("EXPECTATIVA - EMPREGADOS").Cells(59, 1))
'Limpa os n�meros que foram colados mas mant�m a formata��o
Sheets("EXPECTATIVA - EMPREGADOS").Range(Cells(59, 2), Cells(105, 9)).ClearContents



'Nomeia as colunas de acordo com o dado que ser� calculado nelas
Sheets("EXPECTATIVA - EMPREGADOS").Cells(58, 2).Value = "Diferen�a para o m�s anterior"
Sheets("EXPECTATIVA - EMPREGADOS").Cells(58, 3).Value = "Diferen�a para ao mesmo m�s do ano anterior"
Sheets("EXPECTATIVA - EMPREGADOS").Cells(58, 4).Value = "Diferen�a para a m�dia hist�rica"
Sheets("EXPECTATIVA - EMPREGADOS").Cells(58, 5).Value = "Posi��o Decrescente (Maior valor 1�, menor valor �ltimo)"
Sheets("EXPECTATIVA - EMPREGADOS").Cells(58, 6).Value = "Posi��o Crescente  (Menor valor 1�, maior valor �ltimo)"
Sheets("EXPECTATIVA - EMPREGADOS").Cells(58, 7).Value = "Posi��o Crescente - Mesmo m�s  (Menor valor 1�, maior valor �ltimo)"
Sheets("EXPECTATIVA - EMPREGADOS").Cells(58, 8).Value = "Posi��o Decrescente -Mesmo m�s  (Maior valor 1�, menor valor �ltimo)"
Sheets("EXPECTATIVA - EMPREGADOS").Cells(58, 9).Value = "cruzou a linha de 50?"

'Atribui valores as variaveis definidas acima
Expectativas_Empregados_Linha_Dados = 9 'Define o n�mero da primeira linha de dados
Expectativas_Empregados_Coluna_Dados1 = Sheets("EXPECTATIVA - EMPREGADOS").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Expectativas_Empregados_Coluna_Dados2 = Expectativas_Empregados_Coluna_Dados1 - 1 'Define o n�mero da coluna do m�s anterior
Expectativas_Empregados_Coluna_Dados3 = Expectativas_Empregados_Coluna_Dados1 - 12
Expectativas_Empregados_Linha_An�lise = 59 'Define a primeira linhas de an�lises
Expectativas_Empregados_Coluna_An�lise = 2 'Define a coluna de an�lises

'Inserindo valores nas celulas vazias para fugir de bugs
Sheets("EXPECTATIVA - EMPREGADOS").Range(Cells(10, Expectativas_Empregados_Coluna_Dados3), Cells(10, Expectativas_Empregados_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - EMPREGADOS").Range(Cells(16, Expectativas_Empregados_Coluna_Dados3), Cells(16, Expectativas_Empregados_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - EMPREGADOS").Range(Cells(20, Expectativas_Empregados_Coluna_Dados3), Cells(20, Expectativas_Empregados_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - EMPREGADOS").Range(Cells(22, Expectativas_Empregados_Coluna_Dados3), Cells(23, Expectativas_Empregados_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - EMPREGADOS").Range(Cells(25, Expectativas_Empregados_Coluna_Dados3), Cells(25, Expectativas_Empregados_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - EMPREGADOS").Range(Cells(29, Expectativas_Empregados_Coluna_Dados3), Cells(29, Expectativas_Empregados_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - EMPREGADOS").Range(Cells(37, Expectativas_Empregados_Coluna_Dados3), Cells(37, Expectativas_Empregados_Coluna_Dados1)).Value = "0"


'Calculo da difern�a em pontos do valor mais recente em rela��o ao valor do m�s anterior
Do Until Expectativas_Empregados_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Calculo da diferen�a em si: o valor da celula de analise � igual ao valor mais recente menos o valor do m�s anterior
   Sheets("EXPECTATIVA - EMPREGADOS").Cells(Expectativas_Empregados_Linha_An�lise, Expectativas_Empregados_Coluna_An�lise).Value = Sheets("EXPECTATIVA - EMPREGADOS").Cells(Expectativas_Empregados_Linha_Dados, Expectativas_Empregados_Coluna_Dados1).Value - Sheets("EXPECTATIVA - EMPREGADOS").Cells(Expectativas_Empregados_Linha_Dados, Expectativas_Empregados_Coluna_Dados2).Value
    'Vai para a pr�xima linha de dados e de an�lise
   Expectativas_Empregados_Linha_Dados = Expectativas_Empregados_Linha_Dados + 1
   Expectativas_Empregados_Linha_An�lise = Expectativas_Empregados_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Expectativas_Empregados_Linha_Dados = 9
Expectativas_Empregados_Coluna_Dados1 = Sheets("EXPECTATIVA - EMPREGADOS").Range("A9").End(xlToRight).Column
Expectativas_Empregados_Coluna_Dados3 = Expectativas_Empregados_Coluna_Dados1 - 12
Expectativas_Empregados_Linha_An�lise = 59
Expectativas_Empregados_Coluna_An�lise = 3

'C�lculo da diferen�a em pontos do valor mais recente em rela��o ao valor do mesmo m�s do ano anterior
Do Until Expectativas_Empregados_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Calculo da diferen�a em si: o valor da celula de analise � igual ao valor mais recente menos o valor do mesmo m�s do ano anterior
    Sheets("EXPECTATIVA - EMPREGADOS").Cells(Expectativas_Empregados_Linha_An�lise, Expectativas_Empregados_Coluna_An�lise).Value = Sheets("EXPECTATIVA - EMPREGADOS").Cells(Expectativas_Empregados_Linha_Dados, Expectativas_Empregados_Coluna_Dados1).Value - Sheets("EXPECTATIVA - EMPREGADOS").Cells(Expectativas_Empregados_Linha_Dados, Expectativas_Empregados_Coluna_Dados3).Value
    'Vai para a pr�xima linha de dados e de an�lise
    Expectativas_Empregados_Linha_Dados = Expectativas_Empregados_Linha_Dados + 1
    Expectativas_Empregados_Linha_An�lise = Expectativas_Empregados_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Expectativas_Empregados_Linha_Dados = 9
Expectativas_Empregados_Coluna_Dados1 = Sheets("EXPECTATIVA - EMPREGADOS").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Expectativas_Empregados_Linha_An�lise = 59
Expectativas_Empregados_Coluna_An�lise = 4

'C�lculo da diferen�a em pontos do valor mais recente em rela��o ao valor da m�dia hist�rica
Do Until Expectativas_Empregados_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a vari�vel media como a m�dia do intervalo entre a coluna com o dado mais recente e o primeiro
    media = Application.Average(Sheets("EXPECTATIVA - EMPREGADOS").Range(Cells(Expectativas_Empregados_Linha_Dados, Expectativas_Empregados_Coluna_Dados1), Cells(Expectativas_Empregados_Linha_Dados, 2)))
    'Calculo em si: o valor da celula de analise � igual ao valor mais recente menos o valor da m�dia
    Sheets("EXPECTATIVA - EMPREGADOS").Cells(Expectativas_Empregados_Linha_An�lise, Expectativas_Empregados_Coluna_An�lise).Value = Sheets("EXPECTATIVA - EMPREGADOS").Cells(Expectativas_Empregados_Linha_Dados, Expectativas_Empregados_Coluna_Dados1).Value - media
    'Vai para a pr�xima linha de dados e de an�lise
    Expectativas_Empregados_Linha_Dados = Expectativas_Empregados_Linha_Dados + 1
    Expectativas_Empregados_Linha_An�lise = Expectativas_Empregados_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Expectativas_Empregados_Linha_Dados = 9
Expectativas_Empregados_Coluna_Dados1 = Sheets("EXPECTATIVA - EMPREGADOS").Range("A9").End(xlToRight).Column
Expectativas_Empregados_Linha_An�lise = 59
Expectativas_Empregados_Coluna_An�lise = 5

'Ordena��o decrescente da s�rie hist�rica completa
Do Until Expectativas_Empregados_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posi��o = WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVA - EMPREGADOS").Cells(Expectativas_Empregados_Linha_Dados, Expectativas_Empregados_Coluna_Dados1), Sheets("EXPECTATIVA - EMPREGADOS").Range(Cells(Expectativas_Empregados_Linha_Dados, Expectativas_Empregados_Coluna_Dados1), Cells(Expectativas_Empregados_Linha_Dados, 2)), 0)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("EXPECTATIVA - EMPREGADOS").Cells(Expectativas_Empregados_Linha_An�lise, Expectativas_Empregados_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Expectativas_Empregados_Linha_Dados = Expectativas_Empregados_Linha_Dados + 1
    Expectativas_Empregados_Linha_An�lise = Expectativas_Empregados_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Expectativas_Empregados_Linha_Dados = 9
Expectativas_Empregados_Coluna_Dados1 = Sheets("EXPECTATIVA - EMPREGADOS").Range("A9").End(xlToRight).Column
Expectativas_Empregados_Linha_An�lise = 59
Expectativas_Empregados_Coluna_An�lise = 6

'Ordena��o Crescente da s�rie hist�rica completa
Do Until Expectativas_Empregados_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posi��o = WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVA - EMPREGADOS").Cells(Expectativas_Empregados_Linha_Dados, Expectativas_Empregados_Coluna_Dados1), Sheets("EXPECTATIVA - EMPREGADOS").Range(Cells(Expectativas_Empregados_Linha_Dados, Expectativas_Empregados_Coluna_Dados1), Cells(Expectativas_Empregados_Linha_Dados, 2)), 1)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("EXPECTATIVA - EMPREGADOS").Cells(Expectativas_Empregados_Linha_An�lise, Expectativas_Empregados_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Expectativas_Empregados_Linha_Dados = Expectativas_Empregados_Linha_Dados + 1
    Expectativas_Empregados_Linha_An�lise = Expectativas_Empregados_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior e define a vari�vel Coluna_DadosP que representa a primeira coluna de dados
Expectativas_Empregados_Coluna_Dados1 = Sheets("EXPECTATIVA - EMPREGADOS").Range("A9").End(xlToRight).Column
Expectativas_Empregados_Coluna_DadosP = 2

Do Until Expectativas_Empregados_Coluna_DadosP = Expectativas_Empregados_Coluna_Dados1 + 1 ' Faz at� a variavel Coluna_DadosP ser igual a variavel Coluna_Dados1 mais uma unidade
    'Confere se o m�s da coluna em quest�o � igual ao m�s do dado mais recente
    If Month(Sheets("EXPECTATIVA - EMPREGADOS").Cells(8, Expectativas_Empregados_Coluna_DadosP)) = Month(Sheets("EXPECTATIVA - EMPREGADOS").Cells(8, Expectativas_Empregados_Coluna_Dados1)) Then
    'Caso seja igual, copia a coluna com os dados mais abaixo, a partir da linha 110
        Sheets("EXPECTATIVA - EMPREGADOS").Range(Cells(9, Expectativas_Empregados_Coluna_DadosP), (Cells(54, Expectativas_Empregados_Coluna_DadosP))).Copy (Sheets("EXPECTATIVA - EMPREGADOS").Cells(110, Expectativas_Empregados_Coluna_DadosP))
    End If
    'Vai para a pr�xima coluna
    Expectativas_Empregados_Coluna_DadosP = Expectativas_Empregados_Coluna_DadosP + 1
'Repete a conferencia com a pr�xima coluna
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Expectativas_Empregados_Linha_Dados = 110
Expectativas_Empregados_Coluna_Dados1 = Sheets("EXPECTATIVA - EMPREGADOS").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Expectativas_Empregados_Linha_An�lise = 59
Expectativas_Empregados_Coluna_An�lise = 7

'Ordena��o decrescente da s�rie hist�rica dos meses do dado mais recente
Do Until Expectativas_Empregados_Linha_Dados = 156 'Faz o calculo at� a vari�vel Linha_Dados ser 156
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo m�s do mais recente
    posi��o = Application.WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVA - EMPREGADOS").Cells(Expectativas_Empregados_Linha_Dados, Expectativas_Empregados_Coluna_Dados1), Sheets("EXPECTATIVA - EMPREGADOS").Range(Cells(Expectativas_Empregados_Linha_Dados, Expectativas_Empregados_Coluna_Dados1), Cells(Expectativas_Empregados_Linha_Dados, 2)))
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("EXPECTATIVA - EMPREGADOS").Cells(Expectativas_Empregados_Linha_An�lise, Expectativas_Empregados_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Expectativas_Empregados_Linha_Dados = Expectativas_Empregados_Linha_Dados + 1
    Expectativas_Empregados_Linha_An�lise = Expectativas_Empregados_Linha_An�lise + 1
'Repete a ordena��o com a pr�xima linha
Loop

'Atribui os valores originais das variaveis ap�s o loop anterior
Expectativas_Empregados_Linha_Dados = 110
Expectativas_Empregados_Coluna_Dados1 = Sheets("EXPECTATIVA - EMPREGADOS").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Expectativas_Empregados_Linha_An�lise = 59
Expectativas_Empregados_Coluna_An�lise = 8
'Ordena��o crescente da s�rie hist�rica dos meses do dado mais recente
Do Until Expectativas_Empregados_Linha_Dados = 156 'Faz o calculo at� a vari�vel Linha_Dados ser 156
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo m�s do mais recente
    posi��o = Application.WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVA - EMPREGADOS").Cells(Expectativas_Empregados_Linha_Dados, Expectativas_Empregados_Coluna_Dados1), Sheets("EXPECTATIVA - EMPREGADOS").Range(Cells(Expectativas_Empregados_Linha_Dados, Expectativas_Empregados_Coluna_Dados1), Cells(Expectativas_Empregados_Linha_Dados, 2)), 1)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("EXPECTATIVA - EMPREGADOS").Cells(Expectativas_Empregados_Linha_An�lise, Expectativas_Empregados_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Expectativas_Empregados_Linha_Dados = Expectativas_Empregados_Linha_Dados + 1
    Expectativas_Empregados_Linha_An�lise = Expectativas_Empregados_Linha_An�lise + 1
'Repete a ordena��o com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Expectativas_Empregados_Linha_Dados = 9
Expectativas_Empregados_Coluna_Dados1 = Sheets("EXPECTATIVA - EMPREGADOS").Range("A9").End(xlToRight).Column
Expectativas_Empregados_Coluna_Dados2 = Expectativas_Empregados_Coluna_Dados1 - 1
Expectativas_Empregados_Linha_An�lise = 59
Expectativas_Empregados_Coluna_An�lise = 9

'Avalia��o se cruzou ou n�o a linha de 50 e o sentido
Do Until Expectativas_Empregados_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'se o dado mais recente for menor que 50 e o dado do m�s anterior for maior ou igual a 50 ent�o...
    If Cells(Expectativas_Empregados_Linha_Dados, Expectativas_Empregados_Coluna_Dados1) < 50 And Cells(Expectativas_Empregados_Linha_Dados, Expectativas_Empregados_Coluna_Dados2) >= 50 Then
    
    'a c�lula de an�lise recebe cruzou para baixo
    Cells(Expectativas_Empregados_Linha_An�lise, Expectativas_Empregados_Coluna_An�lise).Value = "Cruzou para baixo"
    'Caso n�o seja..
    Else
        'se o dado mais recente for maior ou igual a 50 e o dado do m~es anterior for menor ou igual a 50 ent�o...
        If Cells(Expectativas_Empregados_Linha_Dados, Expectativas_Empregados_Coluna_Dados1) >= 50 And Cells(Expectativas_Empregados_Linha_Dados, Expectativas_Empregados_Coluna_Dados2) <= 50 Then
        'a c�lula de an�lise recebe cruzou para cima
        Cells(Expectativas_Empregados_Linha_An�lise, Expectativas_Empregados_Coluna_An�lise).Value = "Cruzou para cima"
        'Caso n�o seja..
        Else
        'a c�lula de an�lise recebe n�o cruzou
        Cells(Expectativas_Empregados_Linha_An�lise, Expectativas_Empregados_Coluna_An�lise).Value = "N�o Cruzou"
        End If
    End If
    'Vai para a pr�xima linha de dados e de an�lise
    Expectativas_Empregados_Linha_Dados = Expectativas_Empregados_Linha_Dados + 1
    Expectativas_Empregados_Linha_An�lise = Expectativas_Empregados_Linha_An�lise + 1
'repete o processo com a nova linha
Loop

'Apaga as linhas com erros/dados faltantes/t�tulos e subt�tulos
Range(Cells(60, 2), Cells(60, 9)).ClearContents
Range(Cells(66, 2), Cells(66, 9)).ClearContents
Range(Cells(70, 2), Cells(70, 9)).ClearContents
Range(Cells(72, 2), Cells(73, 9)).Value = "-"
Range(Cells(75, 2), Cells(75, 9)).Value = "-"
Range(Cells(79, 2), Cells(79, 9)).Value = "-"
Range(Cells(87, 2), Cells(87, 9)).Value = "-"

'Inserindo valores nas celulas vazias para fugir de bugs
Range(Cells(10, Expectativas_Empregados_Coluna_Dados3), Cells(10, Expectativas_Empregados_Coluna_Dados1)).ClearContents
Range(Cells(16, Expectativas_Empregados_Coluna_Dados3), Cells(16, Expectativas_Empregados_Coluna_Dados1)).ClearContents
Range(Cells(20, Expectativas_Empregados_Coluna_Dados3), Cells(20, Expectativas_Empregados_Coluna_Dados1)).ClearContents
Range(Cells(22, Expectativas_Empregados_Coluna_Dados3), Cells(23, Expectativas_Empregados_Coluna_Dados1)).Value = "-"
Range(Cells(25, Expectativas_Empregados_Coluna_Dados3), Cells(25, Expectativas_Empregados_Coluna_Dados1)).Value = "-"
Range(Cells(29, Expectativas_Empregados_Coluna_Dados3), Cells(29, Expectativas_Empregados_Coluna_Dados1)).Value = "-"
Range(Cells(37, Expectativas_Empregados_Coluna_Dados3), Cells(37, Expectativas_Empregados_Coluna_Dados1)).Value = "-"

Range("E59:H104").NumberFormat = "0"




'***********************************         Expectativa_Investimento      ************************************************************



Dim Expectativas_Investimentos_Linha_Dados As Integer 'Define a linha que cont�m o dado a ser usado
Dim Expectativas_Investimentos_Coluna_Dados1 As Integer ' Define a coluna com o dado mais recente
Dim Expectativas_Investimentos_Coluna_Dados2 As Integer ' Define a coluna com o dado do m�s anterior
Dim Expectativas_Investimentos_Coluna_Dados3 As Integer ' Defie a coluna com o dado do mesmo m�s do ano anterior
Dim Expectativas_Investimentos_Linha_An�lise As Integer ' Define a linha que ser� feita a an�lise
Dim Expectativas_Investimentos_Coluna_An�lise As Integer 'Define a coluna que ser� feita a an�lise

Sheets("EXPECTATIVA - INVESTIMENTO").Select


'Copia os t�tulos das categorias e cola onde ser� formada a tabela de an�lise
Sheets("EXPECTATIVA - INVESTIMENTO").Range(Cells(9, 1), Cells(54, 9)).Copy (Sheets("EXPECTATIVA - INVESTIMENTO").Cells(59, 1))
'Limpa os n�meros que foram colados mas mant�m a formata��o
Sheets("EXPECTATIVA - INVESTIMENTO").Range(Cells(59, 2), Cells(105, 9)).ClearContents



'Nomeia as colunas de acordo com o dado que ser� calculado nelas
Sheets("EXPECTATIVA - INVESTIMENTO").Cells(58, 2).Value = "Diferen�a para o m�s anterior"
Sheets("EXPECTATIVA - INVESTIMENTO").Cells(58, 3).Value = "Diferen�a para ao mesmo m�s do ano anterior"
Sheets("EXPECTATIVA - INVESTIMENTO").Cells(58, 4).Value = "Diferen�a para a m�dia hist�rica"
Sheets("EXPECTATIVA - INVESTIMENTO").Cells(58, 5).Value = "Posi��o Decrescente (Maior valor 1�, menor valor �ltimo)"
Sheets("EXPECTATIVA - INVESTIMENTO").Cells(58, 6).Value = "Posi��o Crescente  (Menor valor 1�, maior valor �ltimo)"
Sheets("EXPECTATIVA - INVESTIMENTO").Cells(58, 7).Value = "Posi��o Crescente - Mesmo m�s  (Menor valor 1�, maior valor �ltimo)"
Sheets("EXPECTATIVA - INVESTIMENTO").Cells(58, 8).Value = "Posi��o Decrescente -Mesmo m�s  (Maior valor 1�, menor valor �ltimo)"
Sheets("EXPECTATIVA - INVESTIMENTO").Cells(58, 9).Value = "cruzou a linha de 50?"

'Atribui valores as variaveis definidas acima
Expectativas_Investimentos_Linha_Dados = 9 'Define o n�mero da primeira linha de dados
Expectativas_Investimentos_Coluna_Dados1 = Sheets("EXPECTATIVA - INVESTIMENTO").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Expectativas_Investimentos_Coluna_Dados2 = Expectativas_Investimentos_Coluna_Dados1 - 1 'Define o n�mero da coluna do m�s anterior
Expectativas_Investimentos_Coluna_Dados3 = Expectativas_Investimentos_Coluna_Dados1 - 12
Expectativas_Investimentos_Linha_An�lise = 59 'Define a primeira linhas de an�lises
Expectativas_Investimentos_Coluna_An�lise = 2 'Define a coluna de an�lises

'Inserindo valores nas celulas vazias para fugir de bugs
Sheets("EXPECTATIVA - INVESTIMENTO").Range(Cells(10, Expectativas_Investimentos_Coluna_Dados3), Cells(10, Expectativas_Investimentos_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - INVESTIMENTO").Range(Cells(16, Expectativas_Investimentos_Coluna_Dados3), Cells(16, Expectativas_Investimentos_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - INVESTIMENTO").Range(Cells(20, Expectativas_Investimentos_Coluna_Dados3), Cells(20, Expectativas_Investimentos_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - INVESTIMENTO").Range(Cells(22, Expectativas_Investimentos_Coluna_Dados3), Cells(23, Expectativas_Investimentos_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - INVESTIMENTO").Range(Cells(25, Expectativas_Investimentos_Coluna_Dados3), Cells(25, Expectativas_Investimentos_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - INVESTIMENTO").Range(Cells(29, Expectativas_Investimentos_Coluna_Dados3), Cells(29, Expectativas_Investimentos_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - INVESTIMENTO").Range(Cells(37, Expectativas_Investimentos_Coluna_Dados3), Cells(37, Expectativas_Investimentos_Coluna_Dados1)).Value = "0"


'Calculo da difern�a em pontos do valor mais recente em rela��o ao valor do m�s anterior
Do Until Expectativas_Investimentos_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Calculo da diferen�a em si: o valor da celula de analise � igual ao valor mais recente menos o valor do m�s anterior
   Sheets("EXPECTATIVA - INVESTIMENTO").Cells(Expectativas_Investimentos_Linha_An�lise, Expectativas_Investimentos_Coluna_An�lise).Value = Sheets("EXPECTATIVA - INVESTIMENTO").Cells(Expectativas_Investimentos_Linha_Dados, Expectativas_Investimentos_Coluna_Dados1).Value - Sheets("EXPECTATIVA - INVESTIMENTO").Cells(Expectativas_Investimentos_Linha_Dados, Expectativas_Investimentos_Coluna_Dados2).Value
    'Vai para a pr�xima linha de dados e de an�lise
   Expectativas_Investimentos_Linha_Dados = Expectativas_Investimentos_Linha_Dados + 1
   Expectativas_Investimentos_Linha_An�lise = Expectativas_Investimentos_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Expectativas_Investimentos_Linha_Dados = 9
Expectativas_Investimentos_Coluna_Dados1 = Sheets("EXPECTATIVA - INVESTIMENTO").Range("A9").End(xlToRight).Column
Expectativas_Investimentos_Coluna_Dados3 = Expectativas_Investimentos_Coluna_Dados1 - 12
Expectativas_Investimentos_Linha_An�lise = 59
Expectativas_Investimentos_Coluna_An�lise = 3

'C�lculo da diferen�a em pontos do valor mais recente em rela��o ao valor do mesmo m�s do ano anterior
Do Until Expectativas_Investimentos_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Calculo da diferen�a em si: o valor da celula de analise � igual ao valor mais recente menos o valor do mesmo m�s do ano anterior
    Sheets("EXPECTATIVA - INVESTIMENTO").Cells(Expectativas_Investimentos_Linha_An�lise, Expectativas_Investimentos_Coluna_An�lise).Value = Sheets("EXPECTATIVA - INVESTIMENTO").Cells(Expectativas_Investimentos_Linha_Dados, Expectativas_Investimentos_Coluna_Dados1).Value - Sheets("EXPECTATIVA - INVESTIMENTO").Cells(Expectativas_Investimentos_Linha_Dados, Expectativas_Investimentos_Coluna_Dados3).Value
    'Vai para a pr�xima linha de dados e de an�lise
    Expectativas_Investimentos_Linha_Dados = Expectativas_Investimentos_Linha_Dados + 1
    Expectativas_Investimentos_Linha_An�lise = Expectativas_Investimentos_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Expectativas_Investimentos_Linha_Dados = 9
Expectativas_Investimentos_Coluna_Dados1 = Sheets("EXPECTATIVA - INVESTIMENTO").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Expectativas_Investimentos_Linha_An�lise = 59
Expectativas_Investimentos_Coluna_An�lise = 4

'C�lculo da diferen�a em pontos do valor mais recente em rela��o ao valor da m�dia hist�rica
Do Until Expectativas_Investimentos_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a vari�vel media como a m�dia do intervalo entre a coluna com o dado mais recente e o primeiro
    media = Application.Average(Sheets("EXPECTATIVA - INVESTIMENTO").Range(Cells(Expectativas_Investimentos_Linha_Dados, Expectativas_Investimentos_Coluna_Dados1), Cells(Expectativas_Investimentos_Linha_Dados, 2)))
    'Calculo em si: o valor da celula de analise � igual ao valor mais recente menos o valor da m�dia
    Sheets("EXPECTATIVA - INVESTIMENTO").Cells(Expectativas_Investimentos_Linha_An�lise, Expectativas_Investimentos_Coluna_An�lise).Value = Sheets("EXPECTATIVA - INVESTIMENTO").Cells(Expectativas_Investimentos_Linha_Dados, Expectativas_Investimentos_Coluna_Dados1).Value - media
    'Vai para a pr�xima linha de dados e de an�lise
    Expectativas_Investimentos_Linha_Dados = Expectativas_Investimentos_Linha_Dados + 1
    Expectativas_Investimentos_Linha_An�lise = Expectativas_Investimentos_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Expectativas_Investimentos_Linha_Dados = 9
Expectativas_Investimentos_Coluna_Dados1 = Sheets("EXPECTATIVA - INVESTIMENTO").Range("A9").End(xlToRight).Column
Expectativas_Investimentos_Linha_An�lise = 59
Expectativas_Investimentos_Coluna_An�lise = 5

'Ordena��o decrescente da s�rie hist�rica completa
Do Until Expectativas_Investimentos_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posi��o = WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVA - INVESTIMENTO").Cells(Expectativas_Investimentos_Linha_Dados, Expectativas_Investimentos_Coluna_Dados1), Sheets("EXPECTATIVA - INVESTIMENTO").Range(Cells(Expectativas_Investimentos_Linha_Dados, Expectativas_Investimentos_Coluna_Dados1), Cells(Expectativas_Investimentos_Linha_Dados, 2)), 0)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("EXPECTATIVA - INVESTIMENTO").Cells(Expectativas_Investimentos_Linha_An�lise, Expectativas_Investimentos_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Expectativas_Investimentos_Linha_Dados = Expectativas_Investimentos_Linha_Dados + 1
    Expectativas_Investimentos_Linha_An�lise = Expectativas_Investimentos_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Expectativas_Investimentos_Linha_Dados = 9
Expectativas_Investimentos_Coluna_Dados1 = Sheets("EXPECTATIVA - INVESTIMENTO").Range("A9").End(xlToRight).Column
Expectativas_Investimentos_Linha_An�lise = 59
Expectativas_Investimentos_Coluna_An�lise = 6

'Ordena��o Crescente da s�rie hist�rica completa
Do Until Expectativas_Investimentos_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posi��o = WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVA - INVESTIMENTO").Cells(Expectativas_Investimentos_Linha_Dados, Expectativas_Investimentos_Coluna_Dados1), Sheets("EXPECTATIVA - INVESTIMENTO").Range(Cells(Expectativas_Investimentos_Linha_Dados, Expectativas_Investimentos_Coluna_Dados1), Cells(Expectativas_Investimentos_Linha_Dados, 2)), 1)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("EXPECTATIVA - INVESTIMENTO").Cells(Expectativas_Investimentos_Linha_An�lise, Expectativas_Investimentos_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Expectativas_Investimentos_Linha_Dados = Expectativas_Investimentos_Linha_Dados + 1
    Expectativas_Investimentos_Linha_An�lise = Expectativas_Investimentos_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior e define a vari�vel Coluna_DadosP que representa a primeira coluna de dados
Expectativas_Investimentos_Coluna_Dados1 = Sheets("EXPECTATIVA - INVESTIMENTO").Range("A9").End(xlToRight).Column
Expectativas_Investimentos_Coluna_DadosP = 2

Do Until Expectativas_Investimentos_Coluna_DadosP = Expectativas_Investimentos_Coluna_Dados1 + 1 ' Faz at� a variavel Coluna_DadosP ser igual a variavel Coluna_Dados1 mais uma unidade
    'Confere se o m�s da coluna em quest�o � igual ao m�s do dado mais recente
    If Month(Sheets("EXPECTATIVA - INVESTIMENTO").Cells(8, Expectativas_Investimentos_Coluna_DadosP)) = Month(Sheets("EXPECTATIVA - INVESTIMENTO").Cells(8, Expectativas_Investimentos_Coluna_Dados1)) Then
    'Caso seja igual, copia a coluna com os dados mais abaixo, a partir da linha 110
        Sheets("EXPECTATIVA - INVESTIMENTO").Range(Cells(9, Expectativas_Investimentos_Coluna_DadosP), (Cells(54, Expectativas_Investimentos_Coluna_DadosP))).Copy (Sheets("EXPECTATIVA - INVESTIMENTO").Cells(110, Expectativas_Investimentos_Coluna_DadosP))
    End If
    'Vai para a pr�xima coluna
    Expectativas_Investimentos_Coluna_DadosP = Expectativas_Investimentos_Coluna_DadosP + 1
'Repete a conferencia com a pr�xima coluna
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Expectativas_Investimentos_Linha_Dados = 110
Expectativas_Investimentos_Coluna_Dados1 = Sheets("EXPECTATIVA - INVESTIMENTO").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Expectativas_Investimentos_Linha_An�lise = 59
Expectativas_Investimentos_Coluna_An�lise = 7

'Ordena��o decrescente da s�rie hist�rica dos meses do dado mais recente
Do Until Expectativas_Investimentos_Linha_Dados = 156 'Faz o calculo at� a vari�vel Linha_Dados ser 156
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo m�s do mais recente
    posi��o = Application.WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVA - INVESTIMENTO").Cells(Expectativas_Investimentos_Linha_Dados, Expectativas_Investimentos_Coluna_Dados1), Sheets("EXPECTATIVA - INVESTIMENTO").Range(Cells(Expectativas_Investimentos_Linha_Dados, Expectativas_Investimentos_Coluna_Dados1), Cells(Expectativas_Investimentos_Linha_Dados, 2)))
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("EXPECTATIVA - INVESTIMENTO").Cells(Expectativas_Investimentos_Linha_An�lise, Expectativas_Investimentos_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Expectativas_Investimentos_Linha_Dados = Expectativas_Investimentos_Linha_Dados + 1
    Expectativas_Investimentos_Linha_An�lise = Expectativas_Investimentos_Linha_An�lise + 1
'Repete a ordena��o com a pr�xima linha
Loop

'Atribui os valores originais das variaveis ap�s o loop anterior
Expectativas_Investimentos_Linha_Dados = 110
Expectativas_Investimentos_Coluna_Dados1 = Sheets("EXPECTATIVA - INVESTIMENTO").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Expectativas_Investimentos_Linha_An�lise = 59
Expectativas_Investimentos_Coluna_An�lise = 8
'Ordena��o crescente da s�rie hist�rica dos meses do dado mais recente
Do Until Expectativas_Investimentos_Linha_Dados = 156 'Faz o calculo at� a vari�vel Linha_Dados ser 156
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo m�s do mais recente
    posi��o = Application.WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVA - INVESTIMENTO").Cells(Expectativas_Investimentos_Linha_Dados, Expectativas_Investimentos_Coluna_Dados1), Sheets("EXPECTATIVA - INVESTIMENTO").Range(Cells(Expectativas_Investimentos_Linha_Dados, Expectativas_Investimentos_Coluna_Dados1), Cells(Expectativas_Investimentos_Linha_Dados, 2)), 1)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("EXPECTATIVA - INVESTIMENTO").Cells(Expectativas_Investimentos_Linha_An�lise, Expectativas_Investimentos_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Expectativas_Investimentos_Linha_Dados = Expectativas_Investimentos_Linha_Dados + 1
    Expectativas_Investimentos_Linha_An�lise = Expectativas_Investimentos_Linha_An�lise + 1
'Repete a ordena��o com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Expectativas_Investimentos_Linha_Dados = 9
Expectativas_Investimentos_Coluna_Dados1 = Sheets("EXPECTATIVA - INVESTIMENTO").Range("A9").End(xlToRight).Column
Expectativas_Investimentos_Coluna_Dados2 = Expectativas_Investimentos_Coluna_Dados1 - 1
Expectativas_Investimentos_Linha_An�lise = 59
Expectativas_Investimentos_Coluna_An�lise = 9

'Avalia��o se cruzou ou n�o a linha de 50 e o sentido
Do Until Expectativas_Investimentos_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'se o dado mais recente for menor que 50 e o dado do m�s anterior for maior ou igual a 50 ent�o...
    If Cells(Expectativas_Investimentos_Linha_Dados, Expectativas_Investimentos_Coluna_Dados1) < 50 And Cells(Expectativas_Investimentos_Linha_Dados, Expectativas_Investimentos_Coluna_Dados2) >= 50 Then
    
    'a c�lula de an�lise recebe cruzou para baixo
    Cells(Expectativas_Investimentos_Linha_An�lise, Expectativas_Investimentos_Coluna_An�lise).Value = "Cruzou para baixo"
    'Caso n�o seja..
    Else
        'se o dado mais recente for maior ou igual a 50 e o dado do m~es anterior for menor ou igual a 50 ent�o...
        If Cells(Expectativas_Investimentos_Linha_Dados, Expectativas_Investimentos_Coluna_Dados1) >= 50 And Cells(Expectativas_Investimentos_Linha_Dados, Expectativas_Investimentos_Coluna_Dados2) <= 50 Then
        'a c�lula de an�lise recebe cruzou para cima
        Cells(Expectativas_Investimentos_Linha_An�lise, Expectativas_Investimentos_Coluna_An�lise).Value = "Cruzou para cima"
        'Caso n�o seja..
        Else
        'a c�lula de an�lise recebe n�o cruzou
        Cells(Expectativas_Investimentos_Linha_An�lise, Expectativas_Investimentos_Coluna_An�lise).Value = "N�o Cruzou"
        End If
    End If
    'Vai para a pr�xima linha de dados e de an�lise
    Expectativas_Investimentos_Linha_Dados = Expectativas_Investimentos_Linha_Dados + 1
    Expectativas_Investimentos_Linha_An�lise = Expectativas_Investimentos_Linha_An�lise + 1
'repete o processo com a nova linha
Loop

'Apaga as linhas com erros/dados faltantes/t�tulos e subt�tulos
Range(Cells(60, 2), Cells(60, 9)).ClearContents
Range(Cells(66, 2), Cells(66, 9)).ClearContents
Range(Cells(70, 2), Cells(70, 9)).ClearContents
Range(Cells(72, 2), Cells(73, 9)).Value = "-"
Range(Cells(75, 2), Cells(75, 9)).Value = "-"
Range(Cells(79, 2), Cells(79, 9)).Value = "-"
Range(Cells(87, 2), Cells(87, 9)).Value = "-"

'Inserindo valores nas celulas vazias para fugir de bugs
Range(Cells(10, Expectativas_Investimentos_Coluna_Dados3), Cells(10, Expectativas_Investimentos_Coluna_Dados1)).ClearContents
Range(Cells(16, Expectativas_Investimentos_Coluna_Dados3), Cells(16, Expectativas_Investimentos_Coluna_Dados1)).ClearContents
Range(Cells(20, Expectativas_Investimentos_Coluna_Dados3), Cells(20, Expectativas_Investimentos_Coluna_Dados1)).ClearContents
Range(Cells(22, Expectativas_Investimentos_Coluna_Dados3), Cells(23, Expectativas_Investimentos_Coluna_Dados1)).Value = "-"
Range(Cells(25, Expectativas_Investimentos_Coluna_Dados3), Cells(25, Expectativas_Investimentos_Coluna_Dados1)).Value = "-"
Range(Cells(29, Expectativas_Investimentos_Coluna_Dados3), Cells(29, Expectativas_Investimentos_Coluna_Dados1)).Value = "-"
Range(Cells(37, Expectativas_Investimentos_Coluna_Dados3), Cells(37, Expectativas_Investimentos_Coluna_Dados1)).Value = "-"

Range("E59:H104").NumberFormat = "0"

End Sub

Sub An�lise_Verde()


Dim Situa��oFinanceira_Lucro_Linha_Dados As Integer 'Define a linha que cont�m o dado a ser usado
Dim Situa��oFinanceira_Lucro_Coluna_Dados1 As Integer ' Define a coluna com o dado mais recente
Dim Situa��oFinanceira_Lucro_Coluna_Dados2 As Integer ' Define a coluna com o dado do m�s anterior
Dim Situa��oFinanceira_Lucro_Coluna_Dados3 As Integer ' Defie a coluna com o dado do mesmo m�s do ano anterior
Dim Situa��oFinanceira_Lucro_Linha_An�lise As Integer ' Define a linha que ser� feita a an�lise
Dim Situa��oFinanceira_Lucro_Coluna_An�lise As Integer 'Define a coluna que ser� feita a an�lise

Sheets("SITUACAO FINANCEIRA LUCRO").Select


'Copia os t�tulos das categorias e cola onde ser� formada a tabela de an�lise
Sheets("SITUACAO FINANCEIRA LUCRO").Range(Cells(9, 1), Cells(54, 9)).Copy (Sheets("SITUACAO FINANCEIRA LUCRO").Cells(59, 1))
'Limpa os n�meros que foram colados mas mant�m a formata��o
Sheets("SITUACAO FINANCEIRA LUCRO").Range(Cells(59, 2), Cells(105, 9)).ClearContents



'Nomeia as colunas de acordo com o dado que ser� calculado nelas
Sheets("SITUACAO FINANCEIRA LUCRO").Cells(58, 2).Value = "Diferen�a para o trimestre anterior"
Sheets("SITUACAO FINANCEIRA LUCRO").Cells(58, 3).Value = "Diferen�a para ao mesmo trimestre do ano anterior"
Sheets("SITUACAO FINANCEIRA LUCRO").Cells(58, 4).Value = "Diferen�a para a m�dia hist�rica"
Sheets("SITUACAO FINANCEIRA LUCRO").Cells(58, 5).Value = "Posi��o Decrescente (Maior valor 1�, menor valor �ltimo)"
Sheets("SITUACAO FINANCEIRA LUCRO").Cells(58, 6).Value = "Posi��o Crescente  (Menor valor 1�, maior valor �ltimo)"
Sheets("SITUACAO FINANCEIRA LUCRO").Cells(58, 7).Value = "Posi��o Crescente - Mesmo trimestre  (Menor valor 1�, maior valor �ltimo)"
Sheets("SITUACAO FINANCEIRA LUCRO").Cells(58, 8).Value = "Posi��o Decrescente -Mesmo trimestre  (Maior valor 1�, menor valor �ltimo)"
Sheets("SITUACAO FINANCEIRA LUCRO").Cells(58, 9).Value = "cruzou a linha de 50?"

'Atribui valores as variaveis definidas acima
Situa��oFinanceira_Lucro_Linha_Dados = 9 'Define o n�mero da primeira linha de dados
Situa��oFinanceira_Lucro_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA LUCRO").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Situa��oFinanceira_Lucro_Coluna_Dados2 = Situa��oFinanceira_Lucro_Coluna_Dados1 - 1 'Define o n�mero da coluna do m�s anterior
Situa��oFinanceira_Lucro_Coluna_Dados3 = Situa��oFinanceira_Lucro_Coluna_Dados1 - 4
Situa��oFinanceira_Lucro_Linha_An�lise = 59 'Define a primeira linhas de an�lises
Situa��oFinanceira_Lucro_Coluna_An�lise = 2 'Define a coluna de an�lises

'Inserindo valores nas celulas vazias para fugir de bugs
Sheets("SITUACAO FINANCEIRA LUCRO").Range(Cells(10, Situa��oFinanceira_Lucro_Coluna_Dados3), Cells(10, Situa��oFinanceira_Lucro_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA LUCRO").Range(Cells(16, Situa��oFinanceira_Lucro_Coluna_Dados3), Cells(16, Situa��oFinanceira_Lucro_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA LUCRO").Range(Cells(20, Situa��oFinanceira_Lucro_Coluna_Dados3), Cells(20, Situa��oFinanceira_Lucro_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA LUCRO").Range(Cells(22, Situa��oFinanceira_Lucro_Coluna_Dados3), Cells(23, Situa��oFinanceira_Lucro_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA LUCRO").Range(Cells(25, Situa��oFinanceira_Lucro_Coluna_Dados3), Cells(25, Situa��oFinanceira_Lucro_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA LUCRO").Range(Cells(29, Situa��oFinanceira_Lucro_Coluna_Dados3), Cells(29, Situa��oFinanceira_Lucro_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA LUCRO").Range(Cells(37, Situa��oFinanceira_Lucro_Coluna_Dados3), Cells(37, Situa��oFinanceira_Lucro_Coluna_Dados1)).Value = "0"


'Calculo da difern�a em pontos do valor mais recente em rela��o ao valor do m�s anterior
Do Until Situa��oFinanceira_Lucro_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Calculo da diferen�a em si: o valor da celula de analise � igual ao valor mais recente menos o valor do m�s anterior
   Sheets("SITUACAO FINANCEIRA LUCRO").Cells(Situa��oFinanceira_Lucro_Linha_An�lise, Situa��oFinanceira_Lucro_Coluna_An�lise).Value = Sheets("SITUACAO FINANCEIRA LUCRO").Cells(Situa��oFinanceira_Lucro_Linha_Dados, Situa��oFinanceira_Lucro_Coluna_Dados1).Value - Sheets("SITUACAO FINANCEIRA LUCRO").Cells(Situa��oFinanceira_Lucro_Linha_Dados, Situa��oFinanceira_Lucro_Coluna_Dados2).Value
    'Vai para a pr�xima linha de dados e de an�lise
   Situa��oFinanceira_Lucro_Linha_Dados = Situa��oFinanceira_Lucro_Linha_Dados + 1
   Situa��oFinanceira_Lucro_Linha_An�lise = Situa��oFinanceira_Lucro_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Situa��oFinanceira_Lucro_Linha_Dados = 9
Situa��oFinanceira_Lucro_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA LUCRO").Range("A9").End(xlToRight).Column
Situa��oFinanceira_Lucro_Coluna_Dados3 = Situa��oFinanceira_Lucro_Coluna_Dados1 - 4
Situa��oFinanceira_Lucro_Linha_An�lise = 59
Situa��oFinanceira_Lucro_Coluna_An�lise = 3

'C�lculo da diferen�a em pontos do valor mais recente em rela��o ao valor do mesmo m�s do ano anterior
Do Until Situa��oFinanceira_Lucro_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Calculo da diferen�a em si: o valor da celula de analise � igual ao valor mais recente menos o valor do mesmo m�s do ano anterior
    Sheets("SITUACAO FINANCEIRA LUCRO").Cells(Situa��oFinanceira_Lucro_Linha_An�lise, Situa��oFinanceira_Lucro_Coluna_An�lise).Value = Sheets("SITUACAO FINANCEIRA LUCRO").Cells(Situa��oFinanceira_Lucro_Linha_Dados, Situa��oFinanceira_Lucro_Coluna_Dados1).Value - Sheets("SITUACAO FINANCEIRA LUCRO").Cells(Situa��oFinanceira_Lucro_Linha_Dados, Situa��oFinanceira_Lucro_Coluna_Dados3).Value
    'Vai para a pr�xima linha de dados e de an�lise
    Situa��oFinanceira_Lucro_Linha_Dados = Situa��oFinanceira_Lucro_Linha_Dados + 1
    Situa��oFinanceira_Lucro_Linha_An�lise = Situa��oFinanceira_Lucro_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Situa��oFinanceira_Lucro_Linha_Dados = 9
Situa��oFinanceira_Lucro_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA LUCRO").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Situa��oFinanceira_Lucro_Linha_An�lise = 59
Situa��oFinanceira_Lucro_Coluna_An�lise = 4

'C�lculo da diferen�a em pontos do valor mais recente em rela��o ao valor da m�dia hist�rica
Do Until Situa��oFinanceira_Lucro_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a vari�vel media como a m�dia do intervalo entre a coluna com o dado mais recente e o primeiro
    media = Application.Average(Sheets("SITUACAO FINANCEIRA LUCRO").Range(Cells(Situa��oFinanceira_Lucro_Linha_Dados, Situa��oFinanceira_Lucro_Coluna_Dados1), Cells(Situa��oFinanceira_Lucro_Linha_Dados, 2)))
    'Calculo em si: o valor da celula de analise � igual ao valor mais recente menos o valor da m�dia
    Sheets("SITUACAO FINANCEIRA LUCRO").Cells(Situa��oFinanceira_Lucro_Linha_An�lise, Situa��oFinanceira_Lucro_Coluna_An�lise).Value = Sheets("SITUACAO FINANCEIRA LUCRO").Cells(Situa��oFinanceira_Lucro_Linha_Dados, Situa��oFinanceira_Lucro_Coluna_Dados1).Value - media
    'Vai para a pr�xima linha de dados e de an�lise
    Situa��oFinanceira_Lucro_Linha_Dados = Situa��oFinanceira_Lucro_Linha_Dados + 1
    Situa��oFinanceira_Lucro_Linha_An�lise = Situa��oFinanceira_Lucro_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Situa��oFinanceira_Lucro_Linha_Dados = 9
Situa��oFinanceira_Lucro_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA LUCRO").Range("A9").End(xlToRight).Column
Situa��oFinanceira_Lucro_Linha_An�lise = 59
Situa��oFinanceira_Lucro_Coluna_An�lise = 5

'Ordena��o decrescente da s�rie hist�rica completa
Do Until Situa��oFinanceira_Lucro_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posi��o = WorksheetFunction.Rank_Eq(Sheets("SITUACAO FINANCEIRA LUCRO").Cells(Situa��oFinanceira_Lucro_Linha_Dados, Situa��oFinanceira_Lucro_Coluna_Dados1), Sheets("SITUACAO FINANCEIRA LUCRO").Range(Cells(Situa��oFinanceira_Lucro_Linha_Dados, Situa��oFinanceira_Lucro_Coluna_Dados1), Cells(Situa��oFinanceira_Lucro_Linha_Dados, 2)), 0)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("SITUACAO FINANCEIRA LUCRO").Cells(Situa��oFinanceira_Lucro_Linha_An�lise, Situa��oFinanceira_Lucro_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Situa��oFinanceira_Lucro_Linha_Dados = Situa��oFinanceira_Lucro_Linha_Dados + 1
    Situa��oFinanceira_Lucro_Linha_An�lise = Situa��oFinanceira_Lucro_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Situa��oFinanceira_Lucro_Linha_Dados = 9
Situa��oFinanceira_Lucro_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA LUCRO").Range("A9").End(xlToRight).Column
Situa��oFinanceira_Lucro_Linha_An�lise = 59
Situa��oFinanceira_Lucro_Coluna_An�lise = 6

'Ordena��o Crescente da s�rie hist�rica completa
Do Until Situa��oFinanceira_Lucro_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posi��o = WorksheetFunction.Rank_Eq(Sheets("SITUACAO FINANCEIRA LUCRO").Cells(Situa��oFinanceira_Lucro_Linha_Dados, Situa��oFinanceira_Lucro_Coluna_Dados1), Sheets("SITUACAO FINANCEIRA LUCRO").Range(Cells(Situa��oFinanceira_Lucro_Linha_Dados, Situa��oFinanceira_Lucro_Coluna_Dados1), Cells(Situa��oFinanceira_Lucro_Linha_Dados, 2)), 1)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("SITUACAO FINANCEIRA LUCRO").Cells(Situa��oFinanceira_Lucro_Linha_An�lise, Situa��oFinanceira_Lucro_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Situa��oFinanceira_Lucro_Linha_Dados = Situa��oFinanceira_Lucro_Linha_Dados + 1
    Situa��oFinanceira_Lucro_Linha_An�lise = Situa��oFinanceira_Lucro_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior e define a vari�vel Coluna_DadosP que representa a primeira coluna de dados
Situa��oFinanceira_Lucro_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA LUCRO").Range("A9").End(xlToRight).Column
Situa��oFinanceira_Lucro_Coluna_DadosP = 2

Do Until Situa��oFinanceira_Lucro_Coluna_DadosP = Situa��oFinanceira_Lucro_Coluna_Dados1 + 1 ' Faz at� a variavel Coluna_DadosP ser igual a variavel Coluna_Dados1 mais uma unidade
    'Confere se o m�s da coluna em quest�o � igual ao m�s do dado mais recente
    If Left(Sheets("SITUACAO FINANCEIRA LUCRO").Cells(8, Situa��oFinanceira_Lucro_Coluna_DadosP), 1) = Left(Sheets("SITUACAO FINANCEIRA LUCRO").Cells(8, Situa��oFinanceira_Lucro_Coluna_Dados1), 1) Then
    'Caso seja igual, copia a coluna com os dados mais abaixo, a partir da linha 110
        Sheets("SITUACAO FINANCEIRA LUCRO").Range(Cells(9, Situa��oFinanceira_Lucro_Coluna_DadosP), (Cells(54, Situa��oFinanceira_Lucro_Coluna_DadosP))).Copy (Sheets("SITUACAO FINANCEIRA LUCRO").Cells(110, Situa��oFinanceira_Lucro_Coluna_DadosP))
    End If
    'Vai para a pr�xima coluna
    Situa��oFinanceira_Lucro_Coluna_DadosP = Situa��oFinanceira_Lucro_Coluna_DadosP + 1
'Repete a conferencia com a pr�xima coluna
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Situa��oFinanceira_Lucro_Linha_Dados = 110
Situa��oFinanceira_Lucro_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA LUCRO").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Situa��oFinanceira_Lucro_Linha_An�lise = 59
Situa��oFinanceira_Lucro_Coluna_An�lise = 7

'Ordena��o decrescente da s�rie hist�rica dos meses do dado mais recente
Do Until Situa��oFinanceira_Lucro_Linha_Dados = 156 'Faz o calculo at� a vari�vel Linha_Dados ser 156
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo m�s do mais recente
    posi��o = Application.WorksheetFunction.Rank_Eq(Sheets("SITUACAO FINANCEIRA LUCRO").Cells(Situa��oFinanceira_Lucro_Linha_Dados, Situa��oFinanceira_Lucro_Coluna_Dados1), Sheets("SITUACAO FINANCEIRA LUCRO").Range(Cells(Situa��oFinanceira_Lucro_Linha_Dados, Situa��oFinanceira_Lucro_Coluna_Dados1), Cells(Situa��oFinanceira_Lucro_Linha_Dados, 2)))
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("SITUACAO FINANCEIRA LUCRO").Cells(Situa��oFinanceira_Lucro_Linha_An�lise, Situa��oFinanceira_Lucro_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Situa��oFinanceira_Lucro_Linha_Dados = Situa��oFinanceira_Lucro_Linha_Dados + 1
    Situa��oFinanceira_Lucro_Linha_An�lise = Situa��oFinanceira_Lucro_Linha_An�lise + 1
'Repete a ordena��o com a pr�xima linha
Loop

'Atribui os valores originais das variaveis ap�s o loop anterior
Situa��oFinanceira_Lucro_Linha_Dados = 110
Situa��oFinanceira_Lucro_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA LUCRO").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Situa��oFinanceira_Lucro_Linha_An�lise = 59
Situa��oFinanceira_Lucro_Coluna_An�lise = 8
'Ordena��o crescente da s�rie hist�rica dos meses do dado mais recente
Do Until Situa��oFinanceira_Lucro_Linha_Dados = 156 'Faz o calculo at� a vari�vel Linha_Dados ser 156
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo m�s do mais recente
    posi��o = Application.WorksheetFunction.Rank_Eq(Sheets("SITUACAO FINANCEIRA LUCRO").Cells(Situa��oFinanceira_Lucro_Linha_Dados, Situa��oFinanceira_Lucro_Coluna_Dados1), Sheets("SITUACAO FINANCEIRA LUCRO").Range(Cells(Situa��oFinanceira_Lucro_Linha_Dados, Situa��oFinanceira_Lucro_Coluna_Dados1), Cells(Situa��oFinanceira_Lucro_Linha_Dados, 2)), 1)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("SITUACAO FINANCEIRA LUCRO").Cells(Situa��oFinanceira_Lucro_Linha_An�lise, Situa��oFinanceira_Lucro_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Situa��oFinanceira_Lucro_Linha_Dados = Situa��oFinanceira_Lucro_Linha_Dados + 1
    Situa��oFinanceira_Lucro_Linha_An�lise = Situa��oFinanceira_Lucro_Linha_An�lise + 1
'Repete a ordena��o com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Situa��oFinanceira_Lucro_Linha_Dados = 9
Situa��oFinanceira_Lucro_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA LUCRO").Range("A9").End(xlToRight).Column
Situa��oFinanceira_Lucro_Coluna_Dados2 = Situa��oFinanceira_Lucro_Coluna_Dados1 - 1
Situa��oFinanceira_Lucro_Linha_An�lise = 59
Situa��oFinanceira_Lucro_Coluna_An�lise = 9

'Avalia��o se cruzou ou n�o a linha de 50 e o sentido
Do Until Situa��oFinanceira_Lucro_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'se o dado mais recente for menor que 50 e o dado do m�s anterior for maior ou igual a 50 ent�o...
    If Cells(Situa��oFinanceira_Lucro_Linha_Dados, Situa��oFinanceira_Lucro_Coluna_Dados1) < 50 And Cells(Situa��oFinanceira_Lucro_Linha_Dados, Situa��oFinanceira_Lucro_Coluna_Dados2) >= 50 Then
    
    'a c�lula de an�lise recebe cruzou para baixo
    Cells(Situa��oFinanceira_Lucro_Linha_An�lise, Situa��oFinanceira_Lucro_Coluna_An�lise).Value = "Cruzou para baixo"
    'Caso n�o seja..
    Else
        'se o dado mais recente for maior ou igual a 50 e o dado do m~es anterior for menor ou igual a 50 ent�o...
        If Cells(Situa��oFinanceira_Lucro_Linha_Dados, Situa��oFinanceira_Lucro_Coluna_Dados1) >= 50 And Cells(Situa��oFinanceira_Lucro_Linha_Dados, Situa��oFinanceira_Lucro_Coluna_Dados2) <= 50 Then
        'a c�lula de an�lise recebe cruzou para cima
        Cells(Situa��oFinanceira_Lucro_Linha_An�lise, Situa��oFinanceira_Lucro_Coluna_An�lise).Value = "Cruzou para cima"
        'Caso n�o seja..
        Else
        'a c�lula de an�lise recebe n�o cruzou
        Cells(Situa��oFinanceira_Lucro_Linha_An�lise, Situa��oFinanceira_Lucro_Coluna_An�lise).Value = "N�o Cruzou"
        End If
    End If
    'Vai para a pr�xima linha de dados e de an�lise
    Situa��oFinanceira_Lucro_Linha_Dados = Situa��oFinanceira_Lucro_Linha_Dados + 1
    Situa��oFinanceira_Lucro_Linha_An�lise = Situa��oFinanceira_Lucro_Linha_An�lise + 1
'repete o processo com a nova linha
Loop

'Apaga as linhas com erros/dados faltantes/t�tulos e subt�tulos
Range(Cells(60, 2), Cells(60, 9)).ClearContents
Range(Cells(66, 2), Cells(66, 9)).ClearContents
Range(Cells(70, 2), Cells(70, 9)).ClearContents
Range(Cells(72, 2), Cells(73, 9)).Value = "-"
Range(Cells(75, 2), Cells(75, 9)).Value = "-"
Range(Cells(79, 2), Cells(79, 9)).Value = "-"
Range(Cells(87, 2), Cells(87, 9)).Value = "-"

'Inserindo valores nas celulas vazias para fugir de bugs
Range(Cells(10, Situa��oFinanceira_Lucro_Coluna_Dados3), Cells(10, Situa��oFinanceira_Lucro_Coluna_Dados1)).ClearContents
Range(Cells(16, Situa��oFinanceira_Lucro_Coluna_Dados3), Cells(16, Situa��oFinanceira_Lucro_Coluna_Dados1)).ClearContents
Range(Cells(20, Situa��oFinanceira_Lucro_Coluna_Dados3), Cells(20, Situa��oFinanceira_Lucro_Coluna_Dados1)).ClearContents
Range(Cells(22, Situa��oFinanceira_Lucro_Coluna_Dados3), Cells(23, Situa��oFinanceira_Lucro_Coluna_Dados1)).Value = "-"
Range(Cells(25, Situa��oFinanceira_Lucro_Coluna_Dados3), Cells(25, Situa��oFinanceira_Lucro_Coluna_Dados1)).Value = "-"
Range(Cells(29, Situa��oFinanceira_Lucro_Coluna_Dados3), Cells(29, Situa��oFinanceira_Lucro_Coluna_Dados1)).Value = "-"
Range(Cells(37, Situa��oFinanceira_Lucro_Coluna_Dados3), Cells(37, Situa��oFinanceira_Lucro_Coluna_Dados1)).Value = "-"

Range("E59:H104").NumberFormat = "0"




'**********************************         Situa��oFinanceira_Pre�oM�dio              **********************************************




Dim Situa��oFinanceira_Pre�oM�dio_Linha_Dados As Integer 'Define a linha que cont�m o dado a ser usado
Dim Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1 As Integer ' Define a coluna com o dado mais recente
Dim Situa��oFinanceira_Pre�oM�dio_Coluna_Dados2 As Integer ' Define a coluna com o dado do m�s anterior
Dim Situa��oFinanceira_Pre�oM�dio_Coluna_Dados3 As Integer ' Defie a coluna com o dado do mesmo m�s do ano anterior
Dim Situa��oFinanceira_Pre�oM�dio_Linha_An�lise As Integer ' Define a linha que ser� feita a an�lise
Dim Situa��oFinanceira_Pre�oM�dio_Coluna_An�lise As Integer 'Define a coluna que ser� feita a an�lise

Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Select


'Copia os t�tulos das categorias e cola onde ser� formada a tabela de an�lise
Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Range(Cells(9, 1), Cells(54, 9)).Copy (Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(59, 1))
'Limpa os n�meros que foram colados mas mant�m a formata��o
Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Range(Cells(59, 2), Cells(105, 9)).ClearContents



'Nomeia as colunas de acordo com o dado que ser� calculado nelas
Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(58, 2).Value = "Diferen�a para o trimestre anterior"
Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(58, 3).Value = "Diferen�a para ao mesmo trimestre do ano anterior"
Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(58, 4).Value = "Diferen�a para a m�dia hist�rica"
Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(58, 5).Value = "Posi��o Decrescente (Maior valor 1�, menor valor �ltimo)"
Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(58, 6).Value = "Posi��o Crescente  (Menor valor 1�, maior valor �ltimo)"
Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(58, 7).Value = "Posi��o Crescente - Mesmo trimestre  (Menor valor 1�, maior valor �ltimo)"
Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(58, 8).Value = "Posi��o Decrescente -Mesmo trimestre  (Maior valor 1�, menor valor �ltimo)"
Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(58, 9).Value = "cruzou a linha de 50?"

'Atribui valores as variaveis definidas acima
Situa��oFinanceira_Pre�oM�dio_Linha_Dados = 9 'Define o n�mero da primeira linha de dados
Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Situa��oFinanceira_Pre�oM�dio_Coluna_Dados2 = Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1 - 1 'Define o n�mero da coluna do m�s anterior
Situa��oFinanceira_Pre�oM�dio_Coluna_Dados3 = Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1 - 4
Situa��oFinanceira_Pre�oM�dio_Linha_An�lise = 59 'Define a primeira linhas de an�lises
Situa��oFinanceira_Pre�oM�dio_Coluna_An�lise = 2 'Define a coluna de an�lises

'Inserindo valores nas celulas vazias para fugir de bugs
Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Range(Cells(10, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados3), Cells(10, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Range(Cells(16, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados3), Cells(16, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Range(Cells(20, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados3), Cells(20, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Range(Cells(22, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados3), Cells(23, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Range(Cells(25, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados3), Cells(25, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Range(Cells(29, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados3), Cells(29, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Range(Cells(37, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados3), Cells(37, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1)).Value = "0"


'Calculo da difern�a em pontos do valor mais recente em rela��o ao valor do m�s anterior
Do Until Situa��oFinanceira_Pre�oM�dio_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Calculo da diferen�a em si: o valor da celula de analise � igual ao valor mais recente menos o valor do m�s anterior
   Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(Situa��oFinanceira_Pre�oM�dio_Linha_An�lise, Situa��oFinanceira_Pre�oM�dio_Coluna_An�lise).Value = Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(Situa��oFinanceira_Pre�oM�dio_Linha_Dados, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1).Value - Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(Situa��oFinanceira_Pre�oM�dio_Linha_Dados, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados2).Value
    'Vai para a pr�xima linha de dados e de an�lise
   Situa��oFinanceira_Pre�oM�dio_Linha_Dados = Situa��oFinanceira_Pre�oM�dio_Linha_Dados + 1
   Situa��oFinanceira_Pre�oM�dio_Linha_An�lise = Situa��oFinanceira_Pre�oM�dio_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Situa��oFinanceira_Pre�oM�dio_Linha_Dados = 9
Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Range("A9").End(xlToRight).Column
Situa��oFinanceira_Pre�oM�dio_Coluna_Dados3 = Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1 - 4
Situa��oFinanceira_Pre�oM�dio_Linha_An�lise = 59
Situa��oFinanceira_Pre�oM�dio_Coluna_An�lise = 3

'C�lculo da diferen�a em pontos do valor mais recente em rela��o ao valor do mesmo m�s do ano anterior
Do Until Situa��oFinanceira_Pre�oM�dio_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Calculo da diferen�a em si: o valor da celula de analise � igual ao valor mais recente menos o valor do mesmo m�s do ano anterior
    Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(Situa��oFinanceira_Pre�oM�dio_Linha_An�lise, Situa��oFinanceira_Pre�oM�dio_Coluna_An�lise).Value = Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(Situa��oFinanceira_Pre�oM�dio_Linha_Dados, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1).Value - Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(Situa��oFinanceira_Pre�oM�dio_Linha_Dados, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados3).Value
    'Vai para a pr�xima linha de dados e de an�lise
    Situa��oFinanceira_Pre�oM�dio_Linha_Dados = Situa��oFinanceira_Pre�oM�dio_Linha_Dados + 1
    Situa��oFinanceira_Pre�oM�dio_Linha_An�lise = Situa��oFinanceira_Pre�oM�dio_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Situa��oFinanceira_Pre�oM�dio_Linha_Dados = 9
Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Situa��oFinanceira_Pre�oM�dio_Linha_An�lise = 59
Situa��oFinanceira_Pre�oM�dio_Coluna_An�lise = 4

'C�lculo da diferen�a em pontos do valor mais recente em rela��o ao valor da m�dia hist�rica
Do Until Situa��oFinanceira_Pre�oM�dio_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a vari�vel media como a m�dia do intervalo entre a coluna com o dado mais recente e o primeiro
    media = Application.Average(Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Range(Cells(Situa��oFinanceira_Pre�oM�dio_Linha_Dados, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1), Cells(Situa��oFinanceira_Pre�oM�dio_Linha_Dados, 2)))
    'Calculo em si: o valor da celula de analise � igual ao valor mais recente menos o valor da m�dia
    Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(Situa��oFinanceira_Pre�oM�dio_Linha_An�lise, Situa��oFinanceira_Pre�oM�dio_Coluna_An�lise).Value = Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(Situa��oFinanceira_Pre�oM�dio_Linha_Dados, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1).Value - media
    'Vai para a pr�xima linha de dados e de an�lise
    Situa��oFinanceira_Pre�oM�dio_Linha_Dados = Situa��oFinanceira_Pre�oM�dio_Linha_Dados + 1
    Situa��oFinanceira_Pre�oM�dio_Linha_An�lise = Situa��oFinanceira_Pre�oM�dio_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Situa��oFinanceira_Pre�oM�dio_Linha_Dados = 9
Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Range("A9").End(xlToRight).Column
Situa��oFinanceira_Pre�oM�dio_Linha_An�lise = 59
Situa��oFinanceira_Pre�oM�dio_Coluna_An�lise = 5

'Ordena��o decrescente da s�rie hist�rica completa
Do Until Situa��oFinanceira_Pre�oM�dio_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posi��o = WorksheetFunction.Rank_Eq(Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(Situa��oFinanceira_Pre�oM�dio_Linha_Dados, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1), Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Range(Cells(Situa��oFinanceira_Pre�oM�dio_Linha_Dados, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1), Cells(Situa��oFinanceira_Pre�oM�dio_Linha_Dados, 2)), 0)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(Situa��oFinanceira_Pre�oM�dio_Linha_An�lise, Situa��oFinanceira_Pre�oM�dio_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Situa��oFinanceira_Pre�oM�dio_Linha_Dados = Situa��oFinanceira_Pre�oM�dio_Linha_Dados + 1
    Situa��oFinanceira_Pre�oM�dio_Linha_An�lise = Situa��oFinanceira_Pre�oM�dio_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Situa��oFinanceira_Pre�oM�dio_Linha_Dados = 9
Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Range("A9").End(xlToRight).Column
Situa��oFinanceira_Pre�oM�dio_Linha_An�lise = 59
Situa��oFinanceira_Pre�oM�dio_Coluna_An�lise = 6

'Ordena��o Crescente da s�rie hist�rica completa
Do Until Situa��oFinanceira_Pre�oM�dio_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posi��o = WorksheetFunction.Rank_Eq(Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(Situa��oFinanceira_Pre�oM�dio_Linha_Dados, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1), Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Range(Cells(Situa��oFinanceira_Pre�oM�dio_Linha_Dados, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1), Cells(Situa��oFinanceira_Pre�oM�dio_Linha_Dados, 2)), 1)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(Situa��oFinanceira_Pre�oM�dio_Linha_An�lise, Situa��oFinanceira_Pre�oM�dio_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Situa��oFinanceira_Pre�oM�dio_Linha_Dados = Situa��oFinanceira_Pre�oM�dio_Linha_Dados + 1
    Situa��oFinanceira_Pre�oM�dio_Linha_An�lise = Situa��oFinanceira_Pre�oM�dio_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior e define a vari�vel Coluna_DadosP que representa a primeira coluna de dados
Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Range("A9").End(xlToRight).Column
Situa��oFinanceira_Pre�oM�dio_Coluna_DadosP = 2

Do Until Situa��oFinanceira_Pre�oM�dio_Coluna_DadosP = Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1 + 1 ' Faz at� a variavel Coluna_DadosP ser igual a variavel Coluna_Dados1 mais uma unidade
    'Confere se o m�s da coluna em quest�o � igual ao m�s do dado mais recente
    If Left(Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(8, Situa��oFinanceira_Pre�oM�dio_Coluna_DadosP), 2) = Left(Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(8, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1), 2) Then
    'Caso seja igual, copia a coluna com os dados mais abaixo, a partir da linha 110
        Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Range(Cells(9, Situa��oFinanceira_Pre�oM�dio_Coluna_DadosP), (Cells(54, Situa��oFinanceira_Pre�oM�dio_Coluna_DadosP))).Copy (Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(110, Situa��oFinanceira_Pre�oM�dio_Coluna_DadosP))
    End If
    'Vai para a pr�xima coluna
    Situa��oFinanceira_Pre�oM�dio_Coluna_DadosP = Situa��oFinanceira_Pre�oM�dio_Coluna_DadosP + 1
'Repete a conferencia com a pr�xima coluna
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Situa��oFinanceira_Pre�oM�dio_Linha_Dados = 110
Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Situa��oFinanceira_Pre�oM�dio_Linha_An�lise = 59
Situa��oFinanceira_Pre�oM�dio_Coluna_An�lise = 7

'Ordena��o decrescente da s�rie hist�rica dos meses do dado mais recente
Do Until Situa��oFinanceira_Pre�oM�dio_Linha_Dados = 156 'Faz o calculo at� a vari�vel Linha_Dados ser 156
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo m�s do mais recente
    posi��o = Application.WorksheetFunction.Rank_Eq(Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(Situa��oFinanceira_Pre�oM�dio_Linha_Dados, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1), Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Range(Cells(Situa��oFinanceira_Pre�oM�dio_Linha_Dados, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1), Cells(Situa��oFinanceira_Pre�oM�dio_Linha_Dados, 2)))
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(Situa��oFinanceira_Pre�oM�dio_Linha_An�lise, Situa��oFinanceira_Pre�oM�dio_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Situa��oFinanceira_Pre�oM�dio_Linha_Dados = Situa��oFinanceira_Pre�oM�dio_Linha_Dados + 1
    Situa��oFinanceira_Pre�oM�dio_Linha_An�lise = Situa��oFinanceira_Pre�oM�dio_Linha_An�lise + 1
'Repete a ordena��o com a pr�xima linha
Loop

'Atribui os valores originais das variaveis ap�s o loop anterior
Situa��oFinanceira_Pre�oM�dio_Linha_Dados = 110
Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Situa��oFinanceira_Pre�oM�dio_Linha_An�lise = 59
Situa��oFinanceira_Pre�oM�dio_Coluna_An�lise = 8
'Ordena��o crescente da s�rie hist�rica dos meses do dado mais recente
Do Until Situa��oFinanceira_Pre�oM�dio_Linha_Dados = 156 'Faz o calculo at� a vari�vel Linha_Dados ser 156
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo m�s do mais recente
    posi��o = Application.WorksheetFunction.Rank_Eq(Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(Situa��oFinanceira_Pre�oM�dio_Linha_Dados, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1), Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Range(Cells(Situa��oFinanceira_Pre�oM�dio_Linha_Dados, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1), Cells(Situa��oFinanceira_Pre�oM�dio_Linha_Dados, 2)), 1)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Cells(Situa��oFinanceira_Pre�oM�dio_Linha_An�lise, Situa��oFinanceira_Pre�oM�dio_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Situa��oFinanceira_Pre�oM�dio_Linha_Dados = Situa��oFinanceira_Pre�oM�dio_Linha_Dados + 1
    Situa��oFinanceira_Pre�oM�dio_Linha_An�lise = Situa��oFinanceira_Pre�oM�dio_Linha_An�lise + 1
'Repete a ordena��o com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Situa��oFinanceira_Pre�oM�dio_Linha_Dados = 9
Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA PRE�O MEDIO").Range("A9").End(xlToRight).Column
Situa��oFinanceira_Pre�oM�dio_Coluna_Dados2 = Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1 - 1
Situa��oFinanceira_Pre�oM�dio_Linha_An�lise = 59
Situa��oFinanceira_Pre�oM�dio_Coluna_An�lise = 9

'Avalia��o se cruzou ou n�o a linha de 50 e o sentido
Do Until Situa��oFinanceira_Pre�oM�dio_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'se o dado mais recente for menor que 50 e o dado do m�s anterior for maior ou igual a 50 ent�o...
    If Cells(Situa��oFinanceira_Pre�oM�dio_Linha_Dados, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1) < 50 And Cells(Situa��oFinanceira_Pre�oM�dio_Linha_Dados, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados2) >= 50 Then
    
    'a c�lula de an�lise recebe cruzou para baixo
    Cells(Situa��oFinanceira_Pre�oM�dio_Linha_An�lise, Situa��oFinanceira_Pre�oM�dio_Coluna_An�lise).Value = "Cruzou para baixo"
    'Caso n�o seja..
    Else
        'se o dado mais recente for maior ou igual a 50 e o dado do m~es anterior for menor ou igual a 50 ent�o...
        If Cells(Situa��oFinanceira_Pre�oM�dio_Linha_Dados, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1) >= 50 And Cells(Situa��oFinanceira_Pre�oM�dio_Linha_Dados, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados2) <= 50 Then
        'a c�lula de an�lise recebe cruzou para cima
        Cells(Situa��oFinanceira_Pre�oM�dio_Linha_An�lise, Situa��oFinanceira_Pre�oM�dio_Coluna_An�lise).Value = "Cruzou para cima"
        'Caso n�o seja..
        Else
        'a c�lula de an�lise recebe n�o cruzou
        Cells(Situa��oFinanceira_Pre�oM�dio_Linha_An�lise, Situa��oFinanceira_Pre�oM�dio_Coluna_An�lise).Value = "N�o Cruzou"
        End If
    End If
    'Vai para a pr�xima linha de dados e de an�lise
    Situa��oFinanceira_Pre�oM�dio_Linha_Dados = Situa��oFinanceira_Pre�oM�dio_Linha_Dados + 1
    Situa��oFinanceira_Pre�oM�dio_Linha_An�lise = Situa��oFinanceira_Pre�oM�dio_Linha_An�lise + 1
'repete o processo com a nova linha
Loop

'Apaga as linhas com erros/dados faltantes/t�tulos e subt�tulos
Range(Cells(60, 2), Cells(60, 9)).ClearContents
Range(Cells(66, 2), Cells(66, 9)).ClearContents
Range(Cells(70, 2), Cells(70, 9)).ClearContents
Range(Cells(72, 2), Cells(73, 9)).Value = "-"
Range(Cells(75, 2), Cells(75, 9)).Value = "-"
Range(Cells(79, 2), Cells(79, 9)).Value = "-"
Range(Cells(87, 2), Cells(87, 9)).Value = "-"

'Inserindo valores nas celulas vazias para fugir de bugs
Range(Cells(10, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados3), Cells(10, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1)).ClearContents
Range(Cells(16, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados3), Cells(16, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1)).ClearContents
Range(Cells(20, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados3), Cells(20, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1)).ClearContents
Range(Cells(22, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados3), Cells(23, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1)).Value = "-"
Range(Cells(25, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados3), Cells(25, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1)).Value = "-"
Range(Cells(29, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados3), Cells(29, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1)).Value = "-"
Range(Cells(37, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados3), Cells(37, Situa��oFinanceira_Pre�oM�dio_Coluna_Dados1)).Value = "-"

Range("E59:H104").NumberFormat = "0"





'***************************     Situa��oFinanceira         ****************************************************************



Dim Situa��oFinanceira_Linha_Dados As Integer 'Define a linha que cont�m o dado a ser usado
Dim Situa��oFinanceira_Coluna_Dados1 As Integer ' Define a coluna com o dado mais recente
Dim Situa��oFinanceira_Coluna_Dados2 As Integer ' Define a coluna com o dado do m�s anterior
Dim Situa��oFinanceira_Coluna_Dados3 As Integer ' Defie a coluna com o dado do mesmo m�s do ano anterior
Dim Situa��oFinanceira_Linha_An�lise As Integer ' Define a linha que ser� feita a an�lise
Dim Situa��oFinanceira_Coluna_An�lise As Integer 'Define a coluna que ser� feita a an�lise

Sheets("SITUACAO FINANCEIRA").Select

'Copia os t�tulos das categorias e cola onde ser� formada a tabela de an�lise
Sheets("SITUACAO FINANCEIRA").Range(Cells(9, 1), Cells(54, 9)).Copy (Sheets("SITUACAO FINANCEIRA").Cells(59, 1))
'Limpa os n�meros que foram colados mas mant�m a formata��o
Sheets("SITUACAO FINANCEIRA").Range(Cells(59, 2), Cells(105, 9)).ClearContents



'Nomeia as colunas de acordo com o dado que ser� calculado nelas
Sheets("SITUACAO FINANCEIRA").Cells(58, 2).Value = "Diferen�a para o trimestre anterior"
Sheets("SITUACAO FINANCEIRA").Cells(58, 3).Value = "Diferen�a para ao mesmo trimestre do ano anterior"
Sheets("SITUACAO FINANCEIRA").Cells(58, 4).Value = "Diferen�a para a m�dia hist�rica"
Sheets("SITUACAO FINANCEIRA").Cells(58, 5).Value = "Posi��o Decrescente (Maior valor 1�, menor valor �ltimo)"
Sheets("SITUACAO FINANCEIRA").Cells(58, 6).Value = "Posi��o Crescente  (Menor valor 1�, maior valor �ltimo)"
Sheets("SITUACAO FINANCEIRA").Cells(58, 7).Value = "Posi��o Crescente - Mesmo trimestre  (Menor valor 1�, maior valor �ltimo)"
Sheets("SITUACAO FINANCEIRA").Cells(58, 8).Value = "Posi��o Decrescente -Mesmo trimestre  (Maior valor 1�, menor valor �ltimo)"
Sheets("SITUACAO FINANCEIRA").Cells(58, 9).Value = "cruzou a linha de 50?"

'Atribui valores as variaveis definidas acima
Situa��oFinanceira_Linha_Dados = 9 'Define o n�mero da primeira linha de dados
Situa��oFinanceira_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Situa��oFinanceira_Coluna_Dados2 = Situa��oFinanceira_Coluna_Dados1 - 1 'Define o n�mero da coluna do m�s anterior
Situa��oFinanceira_Coluna_Dados3 = Situa��oFinanceira_Coluna_Dados1 - 4
Situa��oFinanceira_Linha_An�lise = 59 'Define a primeira linhas de an�lises
Situa��oFinanceira_Coluna_An�lise = 2 'Define a coluna de an�lises

'Inserindo valores nas celulas vazias para fugir de bugs
Sheets("SITUACAO FINANCEIRA").Range(Cells(10, Situa��oFinanceira_Coluna_Dados3), Cells(10, Situa��oFinanceira_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA").Range(Cells(16, Situa��oFinanceira_Coluna_Dados3), Cells(16, Situa��oFinanceira_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA").Range(Cells(20, Situa��oFinanceira_Coluna_Dados3), Cells(20, Situa��oFinanceira_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA").Range(Cells(22, Situa��oFinanceira_Coluna_Dados3), Cells(23, Situa��oFinanceira_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA").Range(Cells(25, Situa��oFinanceira_Coluna_Dados3), Cells(25, Situa��oFinanceira_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA").Range(Cells(29, Situa��oFinanceira_Coluna_Dados3), Cells(29, Situa��oFinanceira_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA").Range(Cells(37, Situa��oFinanceira_Coluna_Dados3), Cells(37, Situa��oFinanceira_Coluna_Dados1)).Value = "0"


'Calculo da difern�a em pontos do valor mais recente em rela��o ao valor do m�s anterior
Do Until Situa��oFinanceira_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Calculo da diferen�a em si: o valor da celula de analise � igual ao valor mais recente menos o valor do m�s anterior
   Sheets("SITUACAO FINANCEIRA").Cells(Situa��oFinanceira_Linha_An�lise, Situa��oFinanceira_Coluna_An�lise).Value = Sheets("SITUACAO FINANCEIRA").Cells(Situa��oFinanceira_Linha_Dados, Situa��oFinanceira_Coluna_Dados1).Value - Sheets("SITUACAO FINANCEIRA").Cells(Situa��oFinanceira_Linha_Dados, Situa��oFinanceira_Coluna_Dados2).Value
    'Vai para a pr�xima linha de dados e de an�lise
   Situa��oFinanceira_Linha_Dados = Situa��oFinanceira_Linha_Dados + 1
   Situa��oFinanceira_Linha_An�lise = Situa��oFinanceira_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Situa��oFinanceira_Linha_Dados = 9
Situa��oFinanceira_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA").Range("A9").End(xlToRight).Column
Situa��oFinanceira_Coluna_Dados3 = Situa��oFinanceira_Coluna_Dados1 - 4
Situa��oFinanceira_Linha_An�lise = 59
Situa��oFinanceira_Coluna_An�lise = 3

'C�lculo da diferen�a em pontos do valor mais recente em rela��o ao valor do mesmo m�s do ano anterior
Do Until Situa��oFinanceira_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Calculo da diferen�a em si: o valor da celula de analise � igual ao valor mais recente menos o valor do mesmo m�s do ano anterior
    Sheets("SITUACAO FINANCEIRA").Cells(Situa��oFinanceira_Linha_An�lise, Situa��oFinanceira_Coluna_An�lise).Value = Sheets("SITUACAO FINANCEIRA").Cells(Situa��oFinanceira_Linha_Dados, Situa��oFinanceira_Coluna_Dados1).Value - Sheets("SITUACAO FINANCEIRA").Cells(Situa��oFinanceira_Linha_Dados, Situa��oFinanceira_Coluna_Dados3).Value
    'Vai para a pr�xima linha de dados e de an�lise
    Situa��oFinanceira_Linha_Dados = Situa��oFinanceira_Linha_Dados + 1
    Situa��oFinanceira_Linha_An�lise = Situa��oFinanceira_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Situa��oFinanceira_Linha_Dados = 9
Situa��oFinanceira_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Situa��oFinanceira_Linha_An�lise = 59
Situa��oFinanceira_Coluna_An�lise = 4

'C�lculo da diferen�a em pontos do valor mais recente em rela��o ao valor da m�dia hist�rica
Do Until Situa��oFinanceira_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a vari�vel media como a m�dia do intervalo entre a coluna com o dado mais recente e o primeiro
    media = Application.Average(Sheets("SITUACAO FINANCEIRA").Range(Cells(Situa��oFinanceira_Linha_Dados, Situa��oFinanceira_Coluna_Dados1), Cells(Situa��oFinanceira_Linha_Dados, 2)))
    'Calculo em si: o valor da celula de analise � igual ao valor mais recente menos o valor da m�dia
    Sheets("SITUACAO FINANCEIRA").Cells(Situa��oFinanceira_Linha_An�lise, Situa��oFinanceira_Coluna_An�lise).Value = Sheets("SITUACAO FINANCEIRA").Cells(Situa��oFinanceira_Linha_Dados, Situa��oFinanceira_Coluna_Dados1).Value - media
    'Vai para a pr�xima linha de dados e de an�lise
    Situa��oFinanceira_Linha_Dados = Situa��oFinanceira_Linha_Dados + 1
    Situa��oFinanceira_Linha_An�lise = Situa��oFinanceira_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Situa��oFinanceira_Linha_Dados = 9
Situa��oFinanceira_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA").Range("A9").End(xlToRight).Column
Situa��oFinanceira_Linha_An�lise = 59
Situa��oFinanceira_Coluna_An�lise = 5

'Ordena��o decrescente da s�rie hist�rica completa
Do Until Situa��oFinanceira_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posi��o = WorksheetFunction.Rank_Eq(Sheets("SITUACAO FINANCEIRA").Cells(Situa��oFinanceira_Linha_Dados, Situa��oFinanceira_Coluna_Dados1), Sheets("SITUACAO FINANCEIRA").Range(Cells(Situa��oFinanceira_Linha_Dados, Situa��oFinanceira_Coluna_Dados1), Cells(Situa��oFinanceira_Linha_Dados, 2)), 0)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("SITUACAO FINANCEIRA").Cells(Situa��oFinanceira_Linha_An�lise, Situa��oFinanceira_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Situa��oFinanceira_Linha_Dados = Situa��oFinanceira_Linha_Dados + 1
    Situa��oFinanceira_Linha_An�lise = Situa��oFinanceira_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Situa��oFinanceira_Linha_Dados = 9
Situa��oFinanceira_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA").Range("A9").End(xlToRight).Column
Situa��oFinanceira_Linha_An�lise = 59
Situa��oFinanceira_Coluna_An�lise = 6

'Ordena��o Crescente da s�rie hist�rica completa
Do Until Situa��oFinanceira_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posi��o = WorksheetFunction.Rank_Eq(Sheets("SITUACAO FINANCEIRA").Cells(Situa��oFinanceira_Linha_Dados, Situa��oFinanceira_Coluna_Dados1), Sheets("SITUACAO FINANCEIRA").Range(Cells(Situa��oFinanceira_Linha_Dados, Situa��oFinanceira_Coluna_Dados1), Cells(Situa��oFinanceira_Linha_Dados, 2)), 1)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("SITUACAO FINANCEIRA").Cells(Situa��oFinanceira_Linha_An�lise, Situa��oFinanceira_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Situa��oFinanceira_Linha_Dados = Situa��oFinanceira_Linha_Dados + 1
    Situa��oFinanceira_Linha_An�lise = Situa��oFinanceira_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior e define a vari�vel Coluna_DadosP que representa a primeira coluna de dados
Situa��oFinanceira_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA").Range("A9").End(xlToRight).Column
Situa��oFinanceira_Coluna_DadosP = 2

Do Until Situa��oFinanceira_Coluna_DadosP = Situa��oFinanceira_Coluna_Dados1 + 1 ' Faz at� a variavel Coluna_DadosP ser igual a variavel Coluna_Dados1 mais uma unidade
    'Confere se o m�s da coluna em quest�o � igual ao m�s do dado mais recente
    If Left(Sheets("SITUACAO FINANCEIRA").Cells(8, Situa��oFinanceira_Coluna_DadosP), 2) = Left(Sheets("SITUACAO FINANCEIRA").Cells(8, Situa��oFinanceira_Coluna_Dados1), 2) Then
    'Caso seja igual, copia a coluna com os dados mais abaixo, a partir da linha 110
        Sheets("SITUACAO FINANCEIRA").Range(Cells(9, Situa��oFinanceira_Coluna_DadosP), (Cells(54, Situa��oFinanceira_Coluna_DadosP))).Copy (Sheets("SITUACAO FINANCEIRA").Cells(110, Situa��oFinanceira_Coluna_DadosP))
    End If
    'Vai para a pr�xima coluna
    Situa��oFinanceira_Coluna_DadosP = Situa��oFinanceira_Coluna_DadosP + 1
'Repete a conferencia com a pr�xima coluna
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Situa��oFinanceira_Linha_Dados = 110
Situa��oFinanceira_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Situa��oFinanceira_Linha_An�lise = 59
Situa��oFinanceira_Coluna_An�lise = 7

'Ordena��o decrescente da s�rie hist�rica dos meses do dado mais recente
Do Until Situa��oFinanceira_Linha_Dados = 156 'Faz o calculo at� a vari�vel Linha_Dados ser 156
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo m�s do mais recente
    posi��o = Application.WorksheetFunction.Rank_Eq(Sheets("SITUACAO FINANCEIRA").Cells(Situa��oFinanceira_Linha_Dados, Situa��oFinanceira_Coluna_Dados1), Sheets("SITUACAO FINANCEIRA").Range(Cells(Situa��oFinanceira_Linha_Dados, Situa��oFinanceira_Coluna_Dados1), Cells(Situa��oFinanceira_Linha_Dados, 2)))
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("SITUACAO FINANCEIRA").Cells(Situa��oFinanceira_Linha_An�lise, Situa��oFinanceira_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Situa��oFinanceira_Linha_Dados = Situa��oFinanceira_Linha_Dados + 1
    Situa��oFinanceira_Linha_An�lise = Situa��oFinanceira_Linha_An�lise + 1
'Repete a ordena��o com a pr�xima linha
Loop

'Atribui os valores originais das variaveis ap�s o loop anterior
Situa��oFinanceira_Linha_Dados = 110
Situa��oFinanceira_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Situa��oFinanceira_Linha_An�lise = 59
Situa��oFinanceira_Coluna_An�lise = 8
'Ordena��o crescente da s�rie hist�rica dos meses do dado mais recente
Do Until Situa��oFinanceira_Linha_Dados = 156 'Faz o calculo at� a vari�vel Linha_Dados ser 156
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo m�s do mais recente
    posi��o = Application.WorksheetFunction.Rank_Eq(Sheets("SITUACAO FINANCEIRA").Cells(Situa��oFinanceira_Linha_Dados, Situa��oFinanceira_Coluna_Dados1), Sheets("SITUACAO FINANCEIRA").Range(Cells(Situa��oFinanceira_Linha_Dados, Situa��oFinanceira_Coluna_Dados1), Cells(Situa��oFinanceira_Linha_Dados, 2)), 1)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("SITUACAO FINANCEIRA").Cells(Situa��oFinanceira_Linha_An�lise, Situa��oFinanceira_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Situa��oFinanceira_Linha_Dados = Situa��oFinanceira_Linha_Dados + 1
    Situa��oFinanceira_Linha_An�lise = Situa��oFinanceira_Linha_An�lise + 1
'Repete a ordena��o com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Situa��oFinanceira_Linha_Dados = 9
Situa��oFinanceira_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA").Range("A9").End(xlToRight).Column
Situa��oFinanceira_Coluna_Dados2 = Situa��oFinanceira_Coluna_Dados1 - 1
Situa��oFinanceira_Linha_An�lise = 59
Situa��oFinanceira_Coluna_An�lise = 9

'Avalia��o se cruzou ou n�o a linha de 50 e o sentido
Do Until Situa��oFinanceira_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'se o dado mais recente for menor que 50 e o dado do m�s anterior for maior ou igual a 50 ent�o...
    If Cells(Situa��oFinanceira_Linha_Dados, Situa��oFinanceira_Coluna_Dados1) < 50 And Cells(Situa��oFinanceira_Linha_Dados, Situa��oFinanceira_Coluna_Dados2) >= 50 Then
    
    'a c�lula de an�lise recebe cruzou para baixo
    Cells(Situa��oFinanceira_Linha_An�lise, Situa��oFinanceira_Coluna_An�lise).Value = "Cruzou para baixo"
    'Caso n�o seja..
    Else
        'se o dado mais recente for maior ou igual a 50 e o dado do m~es anterior for menor ou igual a 50 ent�o...
        If Cells(Situa��oFinanceira_Linha_Dados, Situa��oFinanceira_Coluna_Dados1) >= 50 And Cells(Situa��oFinanceira_Linha_Dados, Situa��oFinanceira_Coluna_Dados2) <= 50 Then
        'a c�lula de an�lise recebe cruzou para cima
        Cells(Situa��oFinanceira_Linha_An�lise, Situa��oFinanceira_Coluna_An�lise).Value = "Cruzou para cima"
        'Caso n�o seja..
        Else
        'a c�lula de an�lise recebe n�o cruzou
        Cells(Situa��oFinanceira_Linha_An�lise, Situa��oFinanceira_Coluna_An�lise).Value = "N�o Cruzou"
        End If
    End If
    'Vai para a pr�xima linha de dados e de an�lise
    Situa��oFinanceira_Linha_Dados = Situa��oFinanceira_Linha_Dados + 1
    Situa��oFinanceira_Linha_An�lise = Situa��oFinanceira_Linha_An�lise + 1
'repete o processo com a nova linha
Loop

'Apaga as linhas com erros/dados faltantes/t�tulos e subt�tulos
Range(Cells(60, 2), Cells(60, 9)).ClearContents
Range(Cells(66, 2), Cells(66, 9)).ClearContents
Range(Cells(70, 2), Cells(70, 9)).ClearContents
Range(Cells(72, 2), Cells(73, 9)).Value = "-"
Range(Cells(75, 2), Cells(75, 9)).Value = "-"
Range(Cells(79, 2), Cells(79, 9)).Value = "-"
Range(Cells(87, 2), Cells(87, 9)).Value = "-"

'Inserindo valores nas celulas vazias para fugir de bugs
Range(Cells(10, Situa��oFinanceira_Coluna_Dados3), Cells(10, Situa��oFinanceira_Coluna_Dados1)).ClearContents
Range(Cells(16, Situa��oFinanceira_Coluna_Dados3), Cells(16, Situa��oFinanceira_Coluna_Dados1)).ClearContents
Range(Cells(20, Situa��oFinanceira_Coluna_Dados3), Cells(20, Situa��oFinanceira_Coluna_Dados1)).ClearContents
Range(Cells(22, Situa��oFinanceira_Coluna_Dados3), Cells(23, Situa��oFinanceira_Coluna_Dados1)).Value = "-"
Range(Cells(25, Situa��oFinanceira_Coluna_Dados3), Cells(25, Situa��oFinanceira_Coluna_Dados1)).Value = "-"
Range(Cells(29, Situa��oFinanceira_Coluna_Dados3), Cells(29, Situa��oFinanceira_Coluna_Dados1)).Value = "-"
Range(Cells(37, Situa��oFinanceira_Coluna_Dados3), Cells(37, Situa��oFinanceira_Coluna_Dados1)).Value = "-"

Range("E59:H104").NumberFormat = "0"


'****************************************                Situa��oFinanceira_Credito           *********************************************

Dim Situa��oFinanceira_Credito_Linha_Dados As Integer 'Define a linha que cont�m o dado a ser usado
Dim Situa��oFinanceira_Credito_Coluna_Dados1 As Integer ' Define a coluna com o dado mais recente
Dim Situa��oFinanceira_Credito_Coluna_Dados2 As Integer ' Define a coluna com o dado do m�s anterior
Dim Situa��oFinanceira_Credito_Coluna_Dados3 As Integer ' Defie a coluna com o dado do mesmo m�s do ano anterior
Dim Situa��oFinanceira_Credito_Linha_An�lise As Integer ' Define a linha que ser� feita a an�lise
Dim Situa��oFinanceira_Credito_Coluna_An�lise As Integer 'Define a coluna que ser� feita a an�lise

Sheets("SITUACAO FINANCEIRA CREDITO").Select

'Copia os t�tulos das categorias e cola onde ser� formada a tabela de an�lise
Sheets("SITUACAO FINANCEIRA CREDITO").Range(Cells(9, 1), Cells(54, 9)).Copy (Sheets("SITUACAO FINANCEIRA CREDITO").Cells(59, 1))
'Limpa os n�meros que foram colados mas mant�m a formata��o
Sheets("SITUACAO FINANCEIRA CREDITO").Range(Cells(59, 2), Cells(105, 9)).ClearContents



'Nomeia as colunas de acordo com o dado que ser� calculado nelas
Sheets("SITUACAO FINANCEIRA CREDITO").Cells(58, 2).Value = "Diferen�a para o trimestre anterior"
Sheets("SITUACAO FINANCEIRA CREDITO").Cells(58, 3).Value = "Diferen�a para ao mesmo trimestre do ano anterior"
Sheets("SITUACAO FINANCEIRA CREDITO").Cells(58, 4).Value = "Diferen�a para a m�dia hist�rica"
Sheets("SITUACAO FINANCEIRA CREDITO").Cells(58, 5).Value = "Posi��o Decrescente (Maior valor 1�, menor valor �ltimo)"
Sheets("SITUACAO FINANCEIRA CREDITO").Cells(58, 6).Value = "Posi��o Crescente  (Menor valor 1�, maior valor �ltimo)"
Sheets("SITUACAO FINANCEIRA CREDITO").Cells(58, 7).Value = "Posi��o Crescente - Mesmo trimestre  (Menor valor 1�, maior valor �ltimo)"
Sheets("SITUACAO FINANCEIRA CREDITO").Cells(58, 8).Value = "Posi��o Decrescente -Mesmo trimestre  (Maior valor 1�, menor valor �ltimo)"
Sheets("SITUACAO FINANCEIRA CREDITO").Cells(58, 9).Value = "cruzou a linha de 50?"

'Atribui valores as variaveis definidas acima
Situa��oFinanceira_Credito_Linha_Dados = 9 'Define o n�mero da primeira linha de dados
Situa��oFinanceira_Credito_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA CREDITO").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Situa��oFinanceira_Credito_Coluna_Dados2 = Situa��oFinanceira_Credito_Coluna_Dados1 - 1 'Define o n�mero da coluna do m�s anterior
Situa��oFinanceira_Credito_Coluna_Dados3 = Situa��oFinanceira_Credito_Coluna_Dados1 - 4
Situa��oFinanceira_Credito_Linha_An�lise = 59 'Define a primeira linhas de an�lises
Situa��oFinanceira_Credito_Coluna_An�lise = 2 'Define a coluna de an�lises

'Inserindo valores nas celulas vazias para fugir de bugs
Sheets("SITUACAO FINANCEIRA CREDITO").Range(Cells(10, Situa��oFinanceira_Credito_Coluna_Dados3), Cells(10, Situa��oFinanceira_Credito_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA CREDITO").Range(Cells(16, Situa��oFinanceira_Credito_Coluna_Dados3), Cells(16, Situa��oFinanceira_Credito_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA CREDITO").Range(Cells(20, Situa��oFinanceira_Credito_Coluna_Dados3), Cells(20, Situa��oFinanceira_Credito_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA CREDITO").Range(Cells(22, Situa��oFinanceira_Credito_Coluna_Dados3), Cells(23, Situa��oFinanceira_Credito_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA CREDITO").Range(Cells(25, Situa��oFinanceira_Credito_Coluna_Dados3), Cells(25, Situa��oFinanceira_Credito_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA CREDITO").Range(Cells(29, Situa��oFinanceira_Credito_Coluna_Dados3), Cells(29, Situa��oFinanceira_Credito_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA CREDITO").Range(Cells(37, Situa��oFinanceira_Credito_Coluna_Dados3), Cells(37, Situa��oFinanceira_Credito_Coluna_Dados1)).Value = "0"


'Calculo da difern�a em pontos do valor mais recente em rela��o ao valor do m�s anterior
Do Until Situa��oFinanceira_Credito_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Calculo da diferen�a em si: o valor da celula de analise � igual ao valor mais recente menos o valor do m�s anterior
   Sheets("SITUACAO FINANCEIRA CREDITO").Cells(Situa��oFinanceira_Credito_Linha_An�lise, Situa��oFinanceira_Credito_Coluna_An�lise).Value = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(Situa��oFinanceira_Credito_Linha_Dados, Situa��oFinanceira_Credito_Coluna_Dados1).Value - Sheets("SITUACAO FINANCEIRA CREDITO").Cells(Situa��oFinanceira_Credito_Linha_Dados, Situa��oFinanceira_Credito_Coluna_Dados2).Value
    'Vai para a pr�xima linha de dados e de an�lise
   Situa��oFinanceira_Credito_Linha_Dados = Situa��oFinanceira_Credito_Linha_Dados + 1
   Situa��oFinanceira_Credito_Linha_An�lise = Situa��oFinanceira_Credito_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Situa��oFinanceira_Credito_Linha_Dados = 9
Situa��oFinanceira_Credito_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA CREDITO").Range("A9").End(xlToRight).Column
Situa��oFinanceira_Credito_Coluna_Dados3 = Situa��oFinanceira_Credito_Coluna_Dados1 - 4
Situa��oFinanceira_Credito_Linha_An�lise = 59
Situa��oFinanceira_Credito_Coluna_An�lise = 3

'C�lculo da diferen�a em pontos do valor mais recente em rela��o ao valor do mesmo m�s do ano anterior
Do Until Situa��oFinanceira_Credito_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Calculo da diferen�a em si: o valor da celula de analise � igual ao valor mais recente menos o valor do mesmo m�s do ano anterior
    Sheets("SITUACAO FINANCEIRA CREDITO").Cells(Situa��oFinanceira_Credito_Linha_An�lise, Situa��oFinanceira_Credito_Coluna_An�lise).Value = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(Situa��oFinanceira_Credito_Linha_Dados, Situa��oFinanceira_Credito_Coluna_Dados1).Value - Sheets("SITUACAO FINANCEIRA CREDITO").Cells(Situa��oFinanceira_Credito_Linha_Dados, Situa��oFinanceira_Credito_Coluna_Dados3).Value
    'Vai para a pr�xima linha de dados e de an�lise
    Situa��oFinanceira_Credito_Linha_Dados = Situa��oFinanceira_Credito_Linha_Dados + 1
    Situa��oFinanceira_Credito_Linha_An�lise = Situa��oFinanceira_Credito_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Situa��oFinanceira_Credito_Linha_Dados = 9
Situa��oFinanceira_Credito_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA CREDITO").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Situa��oFinanceira_Credito_Linha_An�lise = 59
Situa��oFinanceira_Credito_Coluna_An�lise = 4

'C�lculo da diferen�a em pontos do valor mais recente em rela��o ao valor da m�dia hist�rica
Do Until Situa��oFinanceira_Credito_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a vari�vel media como a m�dia do intervalo entre a coluna com o dado mais recente e o primeiro
    media = Application.Average(Sheets("SITUACAO FINANCEIRA CREDITO").Range(Cells(Situa��oFinanceira_Credito_Linha_Dados, Situa��oFinanceira_Credito_Coluna_Dados1), Cells(Situa��oFinanceira_Credito_Linha_Dados, 2)))
    'Calculo em si: o valor da celula de analise � igual ao valor mais recente menos o valor da m�dia
    Sheets("SITUACAO FINANCEIRA CREDITO").Cells(Situa��oFinanceira_Credito_Linha_An�lise, Situa��oFinanceira_Credito_Coluna_An�lise).Value = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(Situa��oFinanceira_Credito_Linha_Dados, Situa��oFinanceira_Credito_Coluna_Dados1).Value - media
    'Vai para a pr�xima linha de dados e de an�lise
    Situa��oFinanceira_Credito_Linha_Dados = Situa��oFinanceira_Credito_Linha_Dados + 1
    Situa��oFinanceira_Credito_Linha_An�lise = Situa��oFinanceira_Credito_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Situa��oFinanceira_Credito_Linha_Dados = 9
Situa��oFinanceira_Credito_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA CREDITO").Range("A9").End(xlToRight).Column
Situa��oFinanceira_Credito_Linha_An�lise = 59
Situa��oFinanceira_Credito_Coluna_An�lise = 5

'Ordena��o decrescente da s�rie hist�rica completa
Do Until Situa��oFinanceira_Credito_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posi��o = WorksheetFunction.Rank_Eq(Sheets("SITUACAO FINANCEIRA CREDITO").Cells(Situa��oFinanceira_Credito_Linha_Dados, Situa��oFinanceira_Credito_Coluna_Dados1), Sheets("SITUACAO FINANCEIRA CREDITO").Range(Cells(Situa��oFinanceira_Credito_Linha_Dados, Situa��oFinanceira_Credito_Coluna_Dados1), Cells(Situa��oFinanceira_Credito_Linha_Dados, 2)), 0)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("SITUACAO FINANCEIRA CREDITO").Cells(Situa��oFinanceira_Credito_Linha_An�lise, Situa��oFinanceira_Credito_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Situa��oFinanceira_Credito_Linha_Dados = Situa��oFinanceira_Credito_Linha_Dados + 1
    Situa��oFinanceira_Credito_Linha_An�lise = Situa��oFinanceira_Credito_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Situa��oFinanceira_Credito_Linha_Dados = 9
Situa��oFinanceira_Credito_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA CREDITO").Range("A9").End(xlToRight).Column
Situa��oFinanceira_Credito_Linha_An�lise = 59
Situa��oFinanceira_Credito_Coluna_An�lise = 6

'Ordena��o Crescente da s�rie hist�rica completa
Do Until Situa��oFinanceira_Credito_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posi��o = WorksheetFunction.Rank_Eq(Sheets("SITUACAO FINANCEIRA CREDITO").Cells(Situa��oFinanceira_Credito_Linha_Dados, Situa��oFinanceira_Credito_Coluna_Dados1), Sheets("SITUACAO FINANCEIRA CREDITO").Range(Cells(Situa��oFinanceira_Credito_Linha_Dados, Situa��oFinanceira_Credito_Coluna_Dados1), Cells(Situa��oFinanceira_Credito_Linha_Dados, 2)), 1)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("SITUACAO FINANCEIRA CREDITO").Cells(Situa��oFinanceira_Credito_Linha_An�lise, Situa��oFinanceira_Credito_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Situa��oFinanceira_Credito_Linha_Dados = Situa��oFinanceira_Credito_Linha_Dados + 1
    Situa��oFinanceira_Credito_Linha_An�lise = Situa��oFinanceira_Credito_Linha_An�lise + 1
'Refaz o calculo com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior e define a vari�vel Coluna_DadosP que representa a primeira coluna de dados
Situa��oFinanceira_Credito_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA CREDITO").Range("A9").End(xlToRight).Column
Situa��oFinanceira_Credito_Coluna_DadosP = 2

Do Until Situa��oFinanceira_Credito_Coluna_DadosP = Situa��oFinanceira_Credito_Coluna_Dados1 + 1 ' Faz at� a variavel Coluna_DadosP ser igual a variavel Coluna_Dados1 mais uma unidade
    'Confere se o m�s da coluna em quest�o � igual ao m�s do dado mais recente
    If Left(Sheets("SITUACAO FINANCEIRA CREDITO").Cells(8, Situa��oFinanceira_Credito_Coluna_DadosP), 2) = Left(Sheets("SITUACAO FINANCEIRA CREDITO").Cells(8, Situa��oFinanceira_Credito_Coluna_Dados1), 2) Then
    'Caso seja igual, copia a coluna com os dados mais abaixo, a partir da linha 110
        Sheets("SITUACAO FINANCEIRA CREDITO").Range(Cells(9, Situa��oFinanceira_Credito_Coluna_DadosP), (Cells(54, Situa��oFinanceira_Credito_Coluna_DadosP))).Copy (Sheets("SITUACAO FINANCEIRA CREDITO").Cells(110, Situa��oFinanceira_Credito_Coluna_DadosP))
    End If
    'Vai para a pr�xima coluna
    Situa��oFinanceira_Credito_Coluna_DadosP = Situa��oFinanceira_Credito_Coluna_DadosP + 1
'Repete a conferencia com a pr�xima coluna
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Situa��oFinanceira_Credito_Linha_Dados = 110
Situa��oFinanceira_Credito_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA CREDITO").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Situa��oFinanceira_Credito_Linha_An�lise = 59
Situa��oFinanceira_Credito_Coluna_An�lise = 7

'Ordena��o decrescente da s�rie hist�rica dos meses do dado mais recente
Do Until Situa��oFinanceira_Credito_Linha_Dados = 156 'Faz o calculo at� a vari�vel Linha_Dados ser 156
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo m�s do mais recente
    posi��o = Application.WorksheetFunction.Rank_Eq(Sheets("SITUACAO FINANCEIRA CREDITO").Cells(Situa��oFinanceira_Credito_Linha_Dados, Situa��oFinanceira_Credito_Coluna_Dados1), Sheets("SITUACAO FINANCEIRA CREDITO").Range(Cells(Situa��oFinanceira_Credito_Linha_Dados, Situa��oFinanceira_Credito_Coluna_Dados1), Cells(Situa��oFinanceira_Credito_Linha_Dados, 2)))
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("SITUACAO FINANCEIRA CREDITO").Cells(Situa��oFinanceira_Credito_Linha_An�lise, Situa��oFinanceira_Credito_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Situa��oFinanceira_Credito_Linha_Dados = Situa��oFinanceira_Credito_Linha_Dados + 1
    Situa��oFinanceira_Credito_Linha_An�lise = Situa��oFinanceira_Credito_Linha_An�lise + 1
'Repete a ordena��o com a pr�xima linha
Loop

'Atribui os valores originais das variaveis ap�s o loop anterior
Situa��oFinanceira_Credito_Linha_Dados = 110
Situa��oFinanceira_Credito_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA CREDITO").Range("A9").End(xlToRight).Column 'Define o n�mero da �ltima coluna
Situa��oFinanceira_Credito_Linha_An�lise = 59
Situa��oFinanceira_Credito_Coluna_An�lise = 8
'Ordena��o crescente da s�rie hist�rica dos meses do dado mais recente
Do Until Situa��oFinanceira_Credito_Linha_Dados = 156 'Faz o calculo at� a vari�vel Linha_Dados ser 156
    'Define a var�vel posi��o como a aplica��o da f�mula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo m�s do mais recente
    posi��o = Application.WorksheetFunction.Rank_Eq(Sheets("SITUACAO FINANCEIRA CREDITO").Cells(Situa��oFinanceira_Credito_Linha_Dados, Situa��oFinanceira_Credito_Coluna_Dados1), Sheets("SITUACAO FINANCEIRA CREDITO").Range(Cells(Situa��oFinanceira_Credito_Linha_Dados, Situa��oFinanceira_Credito_Coluna_Dados1), Cells(Situa��oFinanceira_Credito_Linha_Dados, 2)), 1)
    'Define que a c�lula da an�lise seja igual a posi��o do dado mais recente
    Sheets("SITUACAO FINANCEIRA CREDITO").Cells(Situa��oFinanceira_Credito_Linha_An�lise, Situa��oFinanceira_Credito_Coluna_An�lise).Value = posi��o
    'Vai para a pr�xima linha de dados e de an�lise
    Situa��oFinanceira_Credito_Linha_Dados = Situa��oFinanceira_Credito_Linha_Dados + 1
    Situa��oFinanceira_Credito_Linha_An�lise = Situa��oFinanceira_Credito_Linha_An�lise + 1
'Repete a ordena��o com a pr�xima linha
Loop


'Atribui os valores originais das variaveis ap�s o loop anterior
Situa��oFinanceira_Credito_Linha_Dados = 9
Situa��oFinanceira_Credito_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA CREDITO").Range("A9").End(xlToRight).Column
Situa��oFinanceira_Credito_Coluna_Dados2 = Situa��oFinanceira_Credito_Coluna_Dados1 - 1
Situa��oFinanceira_Credito_Linha_An�lise = 59
Situa��oFinanceira_Credito_Coluna_An�lise = 9

'Avalia��o se cruzou ou n�o a linha de 50 e o sentido
Do Until Situa��oFinanceira_Credito_Linha_Dados = 55 'Faz o calculo at� a vari�vel Linha_Dados ser 55
    'se o dado mais recente for menor que 50 e o dado do m�s anterior for maior ou igual a 50 ent�o...
    If Cells(Situa��oFinanceira_Credito_Linha_Dados, Situa��oFinanceira_Credito_Coluna_Dados1) < 50 And Cells(Situa��oFinanceira_Credito_Linha_Dados, Situa��oFinanceira_Credito_Coluna_Dados2) >= 50 Then
    
    'a c�lula de an�lise recebe cruzou para baixo
    Cells(Situa��oFinanceira_Credito_Linha_An�lise, Situa��oFinanceira_Credito_Coluna_An�lise).Value = "Cruzou para baixo"
    'Caso n�o seja..
    Else
        'se o dado mais recente for maior ou igual a 50 e o dado do m~es anterior for menor ou igual a 50 ent�o...
        If Cells(Situa��oFinanceira_Credito_Linha_Dados, Situa��oFinanceira_Credito_Coluna_Dados1) >= 50 And Cells(Situa��oFinanceira_Credito_Linha_Dados, Situa��oFinanceira_Credito_Coluna_Dados2) <= 50 Then
        'a c�lula de an�lise recebe cruzou para cima
        Cells(Situa��oFinanceira_Credito_Linha_An�lise, Situa��oFinanceira_Credito_Coluna_An�lise).Value = "Cruzou para cima"
        'Caso n�o seja..
        Else
        'a c�lula de an�lise recebe n�o cruzou
        Cells(Situa��oFinanceira_Credito_Linha_An�lise, Situa��oFinanceira_Credito_Coluna_An�lise).Value = "N�o Cruzou"
        End If
    End If
    'Vai para a pr�xima linha de dados e de an�lise
    Situa��oFinanceira_Credito_Linha_Dados = Situa��oFinanceira_Credito_Linha_Dados + 1
    Situa��oFinanceira_Credito_Linha_An�lise = Situa��oFinanceira_Credito_Linha_An�lise + 1
'repete o processo com a nova linha
Loop

'Apaga as linhas com erros/dados faltantes/t�tulos e subt�tulos
Range(Cells(60, 2), Cells(60, 9)).ClearContents
Range(Cells(66, 2), Cells(66, 9)).ClearContents
Range(Cells(70, 2), Cells(70, 9)).ClearContents
Range(Cells(72, 2), Cells(73, 9)).Value = "-"
Range(Cells(75, 2), Cells(75, 9)).Value = "-"
Range(Cells(79, 2), Cells(79, 9)).Value = "-"
Range(Cells(87, 2), Cells(87, 9)).Value = "-"

'Inserindo valores nas celulas vazias para fugir de bugs
Range(Cells(10, Situa��oFinanceira_Credito_Coluna_Dados3), Cells(10, Situa��oFinanceira_Credito_Coluna_Dados1)).ClearContents
Range(Cells(16, Situa��oFinanceira_Credito_Coluna_Dados3), Cells(16, Situa��oFinanceira_Credito_Coluna_Dados1)).ClearContents
Range(Cells(20, Situa��oFinanceira_Credito_Coluna_Dados3), Cells(20, Situa��oFinanceira_Credito_Coluna_Dados1)).ClearContents
Range(Cells(22, Situa��oFinanceira_Credito_Coluna_Dados3), Cells(23, Situa��oFinanceira_Credito_Coluna_Dados1)).Value = "-"
Range(Cells(25, Situa��oFinanceira_Credito_Coluna_Dados3), Cells(25, Situa��oFinanceira_Credito_Coluna_Dados1)).Value = "-"
Range(Cells(29, Situa��oFinanceira_Credito_Coluna_Dados3), Cells(29, Situa��oFinanceira_Credito_Coluna_Dados1)).Value = "-"
Range(Cells(37, Situa��oFinanceira_Credito_Coluna_Dados3), Cells(37, Situa��oFinanceira_Credito_Coluna_Dados1)).Value = "-"

Range("E59:H104").NumberFormat = "0"

End Sub

Sub Formata��o()

Dim Sondagem As Workbook
Dim Modelo As Workbook
    
'   Capture current workbook
    Set Sondagem = ActiveWorkbook
    
'   Open new workbook
    Workbooks.Open ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Sondagem Industrial\Automa��o\Templates\Gr�ficos e Tabelas - Modelo Trimestral.xlsm")

'   Capture new workbook
    Set Modelo = ActiveWorkbook
    
Modelo.Activate
Sheets("TABELAS").Select
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

