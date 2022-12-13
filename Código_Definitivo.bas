Attribute VB_Name = "Código_Definitivo"
Sub PraDarPlay()
Call Aba_Gráfico
Call Tabelas
Call Análise_Vermelho
Call Análise_Azul
Call Análise_Verde
Call Formatação
End Sub


Sub Aba_Gráfico()

Sheets.Add(Before:=Sheets("PRODUÇÃO")).Name = "GRÁFICO" 'Adiciona a aba gráficos

'Adiciona o titúlo dos gráficos, que serão alocados de acordo com a posiçao desses títulos
ActiveSheet.Range("A1").Value = "Evolução da Produção"
ActiveSheet.Range("A2").Value = "Índice de difusão (0 a 100 pontos)*"
ActiveSheet.Range("J1").Value = "Evolução do número de empregados"
ActiveSheet.Range("J2").Value = "Índice de difusão (0 a 100 pontos)*"
ActiveSheet.Range("S1").Value = "Evolução do nível de estoques e do estoque efetivo em relação ao planejado"
ActiveSheet.Range("S2").Value = "Índice de difusão (0 a 100 pontos)*"
ActiveSheet.Range("AC1").Value = "Utilização média da capacidade instalada"
ActiveSheet.Range("AC2").Value = "Percentual (%)"
ActiveSheet.Range("AM1").Value = "Utilização da capacidade instalada efetiva em relação ao usual"
ActiveSheet.Range("AM2").Value = "Índice de difusão (0 a 100 pontos)*"
ActiveSheet.Range("A27").Value = "Índice de expectativa (Compra de Matérias-primas e Número de empregados)"
ActiveSheet.Range("A28").Value = "Índice de difusão (0 a 100 pontos)*"
ActiveSheet.Range("J27").Value = "Índice de expectativa (Demanda e Exporação)"
ActiveSheet.Range("J28").Value = "Índice de difusão (0 a 100 pontos)*"
ActiveSheet.Range("S27").Value = "Intenção de investimento"
ActiveSheet.Range("AC27").Value = "Principais problemas enfrentados pela indústria no trimestre"
ActiveSheet.Range("AC28").Value = "Percentual (%)"
ActiveSheet.Range("A53").Value = "Facilidade de acesso ao crédito"
ActiveSheet.Range("A54").Value = "Índice de difusão (0 a 100 pontos)*"
ActiveSheet.Range("J53").Value = "Preço médio das matérias-primas"
ActiveSheet.Range("J54").Value = "Índice de difusão (0 a 100 pontos)*"
ActiveSheet.Range("S53").Value = "Satisfação com o lucro operacional e com a situação financeira"
ActiveSheet.Range("S54").Value = "Índice de difusão (0 a 100 pontos)*"

'********************************************************  Gráfico Produção     ***************************************************************************

Dim U As Integer 'Número da última Coluna
Dim P As Integer 'Número da primeira coluna
Dim cht As Object 'Gráfico

U = Sheets("PRODUÇÃO").Range("A9").End(xlToRight).Column 'Define o número da última coluna
P = U - 12 'Define o número da primeira coluna

Sheets("PRODUÇÃO").Select 'Seleciona a aba Produção
Sheets("PRODUÇÃO").Range(Cells(55, P), Cells(55, U)).Value = "50" ' Insere a série da linha divisória
Sheets("PRODUÇÃO").Cells(55, 1).Value = "Linha divisória" 'Nomeia a celula que será usada como referencia para o título da série
Sheets("PRODUÇÃO").Cells(7, 2).Value = "Produção" 'Nomeia a celula que será usada como referencia para o título da série


Set cht = Sheets("GRÁFICO").Shapes.AddChart2 'Adiciona o gráfico

Sheets("GRÁFICO").Select 'Seleciona a aba gráfico

cht.Select ' Seleciona o Gráfico
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confederação Nacional da Indústria\ECON - 1 Indicadores Econômicos CNI\1 Indicadores de Atividade Industrial\Sondagem Industrial\Automação\Templates\Emprego.crtx") ' Aplica o template do gráfico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gráfico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gráfico
    ActiveChart.Parent.Top = Parent.Range("A3").Top 'reposiciona o grafico em relação ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("A3").Left ' reposiciona o gráfico em relação à borda esquerda da planilha
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova série ao gráfico
    ActiveChart.FullSeriesCollection(1).Name = "=PRODUÇÃO!" & Cells(7, 2).Address 'Determina o nome da série
    ActiveChart.FullSeriesCollection(1).Values = "=PRODUÇÃO!" & Range(Cells(9, P), Cells(9, U)).Address 'determina os valores da série
    ActiveChart.FullSeriesCollection(1).XValues = "=PRODUÇÃO!" & Range(Cells(8, P), Cells(8, U)).Address 'determina os valores referentes ao eixo x da série adicionada
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova série ao gráfico
    ActiveChart.FullSeriesCollection(2).Name = "=PRODUÇÃO!$A$55" 'Determina o nome da série
    ActiveChart.FullSeriesCollection(2).Values = "=PRODUÇÃO!" & Range(Cells(55, P), Cells(55, U)).Address 'determina os valores da série
    ActiveChart.FullSeriesCollection(2).XValues = "=PRODUÇÃO!" & Range(Cells(8, P), Cells(8, U)).Address 'determina os valores referentes ao eixo x da série adicionada
    ActiveChart.FullSeriesCollection(3).Delete ' Deleta os lixos importados do template

'********************************************************  Gráfico Emprego    ********************************************************************

Dim A As Integer 'Número da última Coluna
Dim B As Integer 'Número da primeira coluna
Dim GrafEmp As Object 'Gráfico

A = Sheets("EMPREGADOS").Range("A9").End(xlToRight).Column 'Define o número da última coluna
B = A - 12 'Define o número da primeira coluna

Sheets("EMPREGADOS").Select 'Seleciona a aba EMPREGADOS
Sheets("EMPREGADOS").Range(Cells(55, B), Cells(55, A)).Value = "50" ' Insere a série da linha divisória
Sheets("EMPREGADOS").Cells(55, 1).Value = "Linha divisória" 'Nomeia a celula que será usada como referencia para o título da série
Sheets("EMPREGADOS").Cells(7, 2).Value = "Emprego" 'Nomeia a celula que será usada como referencia para o título da série


Set GrafEmp = Sheets("GRÁFICO").Shapes.AddChart2 'Adiciona o gráfico

Sheets("GRÁFICO").Select 'Seleciona a aba gráfico

GrafEmp.Select ' Seleciona o Gráfico
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confederação Nacional da Indústria\ECON - 1 Indicadores Econômicos CNI\1 Indicadores de Atividade Industrial\Sondagem Industrial\Automação\Templates\Emprego.crtx") ' Aplica o template do gráfico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gráfico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gráfico
    ActiveChart.Parent.Top = Parent.Range("J3").Top 'reposiciona o grafico em relação ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("J3").Left ' reposiciona o gráfico em relação à borda esquerda da planilha
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova série ao gráfico
    ActiveChart.FullSeriesCollection(1).Name = "=EMPREGADOS!" & Cells(7, 2).Address 'Determina o nome da série
    ActiveChart.FullSeriesCollection(1).Values = "=EMPREGADOS!" & Range(Cells(9, B), Cells(9, A)).Address 'determina os valores da série
    ActiveChart.FullSeriesCollection(1).XValues = "=EMPREGADOS!" & Range(Cells(8, B), Cells(8, A)).Address 'determina os valores referentes ao eixo x da série adicionada
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova série ao gráfico
    ActiveChart.FullSeriesCollection(2).Name = "=EMPREGADOS!$A$55" 'Determina o nome da série
    ActiveChart.FullSeriesCollection(2).Values = "=EMPREGADOS!" & Range(Cells(55, B), Cells(55, A)).Address 'determina os valores da série
    ActiveChart.FullSeriesCollection(2).XValues = "=EMPREGADOS!" & Range(Cells(8, B), Cells(8, A)).Address 'determina os valores referentes ao eixo x da série adicionada
    ActiveChart.FullSeriesCollection(3).Delete ' Deleta os lixos importados do template

'********************************************************  Gráfico Estoques   ********************************************************************

Dim F As Integer 'Número da última Coluna
Dim G As Integer 'Número da primeira coluna
Dim GrafEst As Object 'Gráfico

F = Sheets("ESTOQUES (evolução)").Range("A9").End(xlToRight).Column 'Define o número da última coluna
G = F - 12 'Define o número da primeira coluna

Sheets("ESTOQUES (evolução)").Select 'Seleciona a aba ESTOQUES (evolução)
Sheets("ESTOQUES (evolução)").Range(Cells(55, G), Cells(55, F)).Value = "50" ' Insere a série da linha divisória
Sheets("ESTOQUES (evolução)").Cells(55, 1).Value = "Linha divisória" 'Nomeia a celula que será usada como referencia para o título da série
Sheets("ESTOQUES (evolução)").Cells(7, 2).Value = "Evolução" 'Nomeia a celula que será usada como referencia para o título da série
Sheets("ESTOQUES (efetivo-planejado)").Cells(7, 2).Value = "Efetivo-planejado"

Set GrafEst = Sheets("GRÁFICO").Shapes.AddChart2 'Adiciona o gráfico

Sheets("GRÁFICO").Select 'Seleciona a aba gráfico

GrafEst.Select ' Seleciona o Gráfico
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confederação Nacional da Indústria\ECON - 1 Indicadores Econômicos CNI\1 Indicadores de Atividade Industrial\Sondagem Industrial\Automação\Templates\Estoque.crtx") ' Aplica o template do gráfico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gráfico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gráfico
    ActiveChart.Parent.Top = Parent.Range("S3").Top 'reposiciona o grafico em relação ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("S3").Left ' reposiciona o gráfico em relação à borda esquerda da planilha
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova série ao gráfico
    ActiveChart.FullSeriesCollection(1).Name = "='ESTOQUES (evolução)'!" & Cells(7, 2).Address 'Determina o nome da série
    ActiveChart.FullSeriesCollection(1).Values = "='ESTOQUES (evolução)'!" & Range(Cells(9, G), Cells(9, F)).Address 'determina os valores da série
    ActiveChart.FullSeriesCollection(1).XValues = "='ESTOQUES (evolução)'!" & Range(Cells(8, G), Cells(8, F)).Address 'determina os valores referentes ao eixo x da série adicionada
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.NumberFormat = "#.##0,0"
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova série ao gráfico
    ActiveChart.FullSeriesCollection(2).Name = "='ESTOQUES (efetivo-planejado)'!" & Cells(7, 2).Address 'Determina o nome da série
    ActiveChart.FullSeriesCollection(2).Values = "='ESTOQUES (efetivo-planejado)'!" & Range(Cells(9, G + 12), Cells(9, F + 12)).Address 'determina os valores da série
    ActiveChart.FullSeriesCollection(2).XValues = "='ESTOQUES (efetivo-planejado'!" & Range(Cells(8, G + 12), Cells(8, F + 12)).Address 'determina os valores referentes ao eixo x da série adicionada
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova série ao gráfico
    ActiveChart.FullSeriesCollection(3).Name = "='ESTOQUES (evolução)'!$A$55" 'Determina o nome da série
    ActiveChart.FullSeriesCollection(3).Values = "='ESTOQUES (evolução)'!" & Range(Cells(55, G), Cells(55, F)).Address 'determina os valores da série
    ActiveChart.FullSeriesCollection(3).XValues = "='ESTOQUES (evolução)'!" & Range(Cells(8, G), Cells(8, F)).Address 'determina os valores referentes ao eixo x da série adicionada

   

'********************************************************  Gráfico UCI    ********************************************************************

Dim C As Integer 'Número da última Coluna
Dim GrafUCI As Object 'Gráfico

C = Sheets("UCI (%)").Range("A9").End(xlToRight).Column 'Define o número da última coluna

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

'Copia os dados nos de acordo com a tabela criada no código anterior
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

'Calcula e nomeia a média dos meses com os valores de 2011 a 2019
ActiveSheet.Range("K56").Value = "Média 2011 - 2019"
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

Set GrafUCI = Sheets("GRÁFICO").Shapes.AddChart2 'Adiciona o gráfico

Sheets("GRÁFICO").Select 'Seleciona a aba gráfico

GrafUCI.Select ' Seleciona o Gráfico
ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confederação Nacional da Indústria\ECON - 1 Indicadores Econômicos CNI\1 Indicadores de Atividade Industrial\Sondagem Industrial\Automação\Templates\UCI.crtx") ' Aplica o template do gráfico
ActiveChart.Parent.Height = 300 'ajusta a altura do gráfico
ActiveChart.Parent.Width = 425 ' ajusta a largura do gráfico
ActiveChart.Parent.Top = Parent.Range("AC3").Top 'reposiciona o grafico em relação ao topo da planilha
ActiveChart.Parent.Left = Parent.Range("AC3").Left ' reposiciona o gráfico em relação à borda esquerda da planilha
ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova série ao gráfico
ActiveChart.FullSeriesCollection(1).Name = "='UCI (%)'!$K$56" 'Determina o nome da série
ActiveChart.FullSeriesCollection(1).Values = "='UCI (%)'!$L$56:$W$56" 'determina os valores da série
ActiveChart.FullSeriesCollection(1).XValues = "='UCI (%)'!$L$57:$W$57" 'determina os valores referentes ao eixo x da série adicionada
ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova série ao gráfico
ActiveChart.FullSeriesCollection(2).Name = "='UCI (%)'!$K$67" 'Determina o nome da série
ActiveChart.FullSeriesCollection(2).Values = "='UCI (%)'!$L$67:$W$67" 'determina os valores da série
ActiveChart.FullSeriesCollection(2).XValues = "='UCI (%)'!$L$57:$W$57" 'determina os valores referentes ao eixo x da série adicionada
ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova série ao gráfico
ActiveChart.FullSeriesCollection(3).Name = "='UCI (%)'!$K$68" 'Determina o nome da série
ActiveChart.FullSeriesCollection(3).Values = "='UCI (%)'!$L$68:$W$68" 'determina os valores da série
ActiveChart.FullSeriesCollection(3).XValues = "='UCI (%)'!$L$57:$W$57" 'determina os valores referentes ao eixo x da série adicionada
ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova série ao gráfico
ActiveChart.FullSeriesCollection(4).Name = "='UCI (%)'!$K$69" 'Determina o nome da série
ActiveChart.FullSeriesCollection(4).Values = "='UCI (%)'!$L$69:$W$69" 'determina os valores da série
ActiveChart.FullSeriesCollection(4).XValues = "='UCI (%)'!$B$57:$M$57" 'determina os valores referentes ao eixo x da série adicionada
ActiveChart.SetElement (msoElementLegendBottom)



'                   * Os códigos abaixo contém um exemplo do que deve ser feito para o ano de 2022.
    



'Sempre que um ano novo se iniciar é necessário ajustar este código, a começar pela consolidação do ano passado e a adição
'do novo ano na parte do 'código descrita por "'Copia os dados nos de acordo com a tabela criada no código anterior" a partir da seguencia abaixo:

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

'É necessário também  a adição da linha com o ano novo com na tabela de anos e meses com o código que segue na sessão
'descrita por "'Cria a tabela com os anos nas linhas e os meses nas colunas"

'                      ActiveSheet.Range("A69").Value = "2022"


'Há duas maneiras de prossegui a partir deste momento 1 adicionando o ano passado à média para manter a estrutura de 3 linhas
'ou 2 adicionar uma nova série com o novo ano.

'1) Para adicionar o ano passado à média basta fazer os seguintes ajustes na seção "'Calcula e nomeia a média dos meses com os valores de 2011 a 2019" :


'   ActiveSheet.Range("A56").Value = "Média 2011 - 2020"
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

'1.1) Para manter a estrutura de 3 linhas basta ajustar a seção ""'Seleciona o Gráfico" da fprma que segue abaixo:

'   ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova série ao gráfico
'   ActiveChart.FullSeriesCollection(1).Name = "='UCI (%)'!$A$56" 'Determina o nome da série
'   ActiveChart.FullSeriesCollection(1).Values = "='UCI (%)'!$B$56:$M$56" 'determina os valores da série
'   ActiveChart.FullSeriesCollection(1).XValues = "='UCI (%)'!$B$57:$M$57" 'determina os valores referentes ao eixo x da série adicionada
'   ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova série ao gráfico
'   ActiveChart.FullSeriesCollection(2).Name = "='UCI (%)'!$A$67" 'Determina o nome da série
'   ActiveChart.FullSeriesCollection(2).Values = "='UCI (%)'!$B$68:$M$68" 'determina os valores da série
'   ActiveChart.FullSeriesCollection(2).XValues = "='UCI (%)'!$B$57:$M$57" 'determina os valores referentes ao eixo x da série adicionada
'   ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova série ao gráfico
'   ActiveChart.FullSeriesCollection(3).Name = "='UCI (%)'!$A$68" 'Determina o nome da série
'   ActiveChart.FullSeriesCollection(3).Values = "='UCI (%)'!$B$69:$M$69" 'determina os valores da série
'   ActiveChart.FullSeriesCollection(3).XValues = "='UCI (%)'!$B$57:$M$57" 'determina os valores referentes ao eixo x da série adicionada
'   ActiveChart.FullSeriesCollection(6).Delete ' Deleta os lixos importados do template
'   ActiveChart.FullSeriesCollection(4).Delete ' Deleta os lixos importados do template
'   ActiveChart.FullSeriesCollection(4).Delete ' Deleta os lixos importados do template
    
    
'2) para adicionarar uma nova série o novo ano a basta fazer os seguintes ajustes na seção "'Seleciona o Gráfico":

  
'   ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova série ao gráfico
'   ActiveChart.FullSeriesCollection(1).Name = "='UCI (%)'!$A$56" 'Determina o nome da série
'   ActiveChart.FullSeriesCollection(1).Values = "='UCI (%)'!$B$56:$M$56" 'determina os valores da série
'   ActiveChart.FullSeriesCollection(1).XValues = "='UCI (%)'!$B$57:$M$57" 'determina os valores referentes ao eixo x da série adicionada
'   ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova série ao gráfico
'   ActiveChart.FullSeriesCollection(2).Name = "='UCI (%)'!$A$67" 'Determina o nome da série
'   ActiveChart.FullSeriesCollection(2).Values = "='UCI (%)'!$B$67:$M$67" 'determina os valores da série
'   ActiveChart.FullSeriesCollection(2).XValues = "='UCI (%)'!$B$57:$M$57" 'determina os valores referentes ao eixo x da série adicionada
'   ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova série ao gráfico
'   ActiveChart.FullSeriesCollection(3).Name = "='UCI (%)'!$A$68" 'Determina o nome da série
'   ActiveChart.FullSeriesCollection(3).Values = "='UCI (%)'!$B$68:$M$68" 'determina os valores da série
'   ActiveChart.FullSeriesCollection(3).XValues = "='UCI (%)'!$B$57:$M$57" 'determina os valores referentes ao eixo x da série adicionada
'   ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova série ao gráfico
'   ActiveChart.FullSeriesCollection(4).Name = "='UCI (%)'!$A$69" 'Determina o nome da série
'   ActiveChart.FullSeriesCollection(4).Values = "='UCI (%)'!$B$69:$M$69" 'determina os valores da série
'   ActiveChart.FullSeriesCollection(4).XValues = "='UCI (%)'!$B$57:$M$57" 'determina os valores referentes ao eixo x da série adicionada
'   ActiveChart.FullSeriesCollection(5).Delete ' Deleta os lixos importados do template'
'   ActiveChart.FullSeriesCollection(5).Delete ' Deleta os lixos importados do template

'********************************************************  Gráfico UCI Efetivo Usual    ********************************************************************

Dim D As Integer 'Número da última Coluna
Dim E As Integer
Dim GrafUCIEU As Object 'Gráfico

D = Sheets("UCI (efetiva-usual)").Range("A9").End(xlToRight).Column 'Define o número da última coluna
E = D - 132

Sheets("UCI (efetiva-usual)").Select 'Seleciona a aba UCI (efetiva-usual)
Sheets("UCI (efetiva-usual)").Range(Cells(55, E), Cells(55, D)).Value = "50" ' Insere a série da linha divisória
Sheets("UCI (efetiva-usual)").Cells(55, 1).Value = "Linha divisória" 'Nomeia a celula que será usada como referencia para o título da série
Sheets("UCI (efetiva-usual)").Cells(7, 2).Value = "UCI (efetiva-usual)" 'Nomeia a celula que será usada como referencia para o título da série



Set GrafUCIEU = Sheets("GRÁFICO").Shapes.AddChart2 'Adiciona o gráfico

Sheets("GRÁFICO").Select 'Seleciona a aba gráfico

GrafUCIEU.Select ' Seleciona o Gráfico
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confederação Nacional da Indústria\ECON - 1 Indicadores Econômicos CNI\1 Indicadores de Atividade Industrial\Sondagem Industrial\Automação\Templates\UCI(Efetiva Usual).crtx") ' Aplica o template do gráfico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gráfico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gráfico
    ActiveChart.Parent.Top = Parent.Range("AM3").Top 'reposiciona o grafico em relação ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("AM3").Left ' reposiciona o gráfico em relação à borda esquerda da planilha
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova série ao gráfico
    ActiveChart.FullSeriesCollection(1).Name = "='UCI (efetiva-usual)'!" & Cells(7, 2).Address 'Determina o nome da série
    ActiveChart.FullSeriesCollection(1).Values = "='UCI (efetiva-usual)'!" & Range(Cells(9, E), Cells(9, D)).Address 'determina os valores da série
    ActiveChart.FullSeriesCollection(1).XValues = "='UCI (efetiva-usual)'!" & Range(Cells(8, E), Cells(8, D)).Address 'determina os valores referentes ao eixo x da série adicionada
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova série ao gráfico
    ActiveChart.FullSeriesCollection(2).Name = "='UCI (efetiva-usual)'!$A$55" 'Determina o nome da série
    ActiveChart.FullSeriesCollection(2).Values = "='UCI (efetiva-usual)'!" & Range(Cells(55, E), Cells(55, D)).Address 'determina os valores da série
    ActiveChart.FullSeriesCollection(2).XValues = "='UCI (efetiva-usual)'!" & Range(Cells(8, E), Cells(8, D)).Address 'determina os valores referentes ao eixo x da série adicionada
   

'********************************************************  Gráfico Expectativa Compras e Empregados    ********************************************************************

Dim J As Integer
Dim K As Integer
Dim GrafComEmp As Object

J = Sheets("EXPECTATIVA - COMPRAS").Range("A9").End(xlToRight).Column
K = J - 120

Sheets("EXPECTATIVA - COMPRAS").Select 'Seleciona a aba EXPECTATIVA - COMPRAS
Sheets("EXPECTATIVA - COMPRAS").Range(Cells(55, K), Cells(55, J)).Value = "50" ' Insere a série da linha divisória
Sheets("EXPECTATIVA - COMPRAS").Cells(55, 1).Value = "Linha divisória" 'Nomeia a celula que será usada como referencia para o título da série
Sheets("EXPECTATIVA - COMPRAS").Cells(7, 2).Value = "Expectativa de compras de matérias-primas" 'Nomeia a celula que será usada como referencia para o título da série
Sheets("EXPECTATIVA - EMPREGADOS").Cells(7, 2).Value = "Expectativa de número de empregados"


Set GrafGrafComEmp = Sheets("GRÁFICO").Shapes.AddChart2 'Adiciona o gráfico

Sheets("GRÁFICO").Select 'Seleciona a aba gráfico

GrafGrafComEmp.Select ' Seleciona o Gráfico
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confederação Nacional da Indústria\ECON - 1 Indicadores Econômicos CNI\1 Indicadores de Atividade Industrial\Sondagem Industrial\Automação\Templates\Expectativa - Demanda e Exportação.crtx") ' Aplica o template do gráfico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gráfico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gráfico
    ActiveChart.Parent.Top = Parent.Range("A29").Top 'reposiciona o grafico em relação ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("A29").Left ' reposiciona o gráfico em relação à borda esquerda da planilha
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova série ao gráfico
    ActiveChart.FullSeriesCollection(1).Name = "='EXPECTATIVA - COMPRAS'!" & Cells(7, 2).Address 'Determina o nome da série
    ActiveChart.FullSeriesCollection(1).Values = "='EXPECTATIVA - COMPRAS'!" & Range(Cells(9, K), Cells(9, J)).Address 'determina os valores da série
    ActiveChart.FullSeriesCollection(1).XValues = "='EXPECTATIVA - COMPRAS'!" & Range(Cells(8, K), Cells(8, J)).Address 'determina os valores referentes ao eixo x da série adicionada
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova série ao gráfico
    ActiveChart.FullSeriesCollection(2).Name = "='EXPECTATIVA - EMPREGADOS'!" & Cells(7, 2).Address 'Determina o nome da série
    ActiveChart.FullSeriesCollection(2).Values = "='EXPECTATIVA - EMPREGADOS'!" & Range(Cells(9, K - 8), Cells(9, J - 8)).Address 'determina os valores da série
    ActiveChart.FullSeriesCollection(2).XValues = "='EXPECTATIVA - EMPREGADOS'!" & Range(Cells(8, K - 8), Cells(8, J - 8)).Address 'determina os valores referentes ao eixo x da série adicionada
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova série ao gráfico
    ActiveChart.FullSeriesCollection(3).Name = "='EXPECTATIVA - COMPRAS'!$A$55" 'Determina o nome da série
    ActiveChart.FullSeriesCollection(3).Values = "='EXPECTATIVA - COMPRAS'!" & Range(Cells(55, K), Cells(55, J)).Address 'determina os valores da série
    ActiveChart.FullSeriesCollection(3).XValues = "='EXPECTATIVA - COMPRAS'!" & Range(Cells(8, K), Cells(8, J)).Address 'determina os valores referentes ao eixo x da série adicionada


'********************************************************  Gráfico Expectativa Demanda e Exportação    ********************************************************************

Dim H As Integer
Dim I As Integer
Dim GrafDemExt As Object

H = Sheets("EXPECTATIVAS - DEMANDA").Range("A9").End(xlToRight).Column
I = H - 132

Sheets("EXPECTATIVAS - DEMANDA").Select 'Seleciona a aba EXPECTATIVAS - DEMANDA
Sheets("EXPECTATIVAS - DEMANDA").Range(Cells(55, I), Cells(55, H)).Value = "50" ' Insere a série da linha divisória
Sheets("EXPECTATIVAS - DEMANDA").Cells(55, 1).Value = "Linha divisória" 'Nomeia a celula que será usada como referencia para o título da série
Sheets("EXPECTATIVAS - DEMANDA").Cells(7, 2).Value = "Expectativa de demanda" 'Nomeia a celula que será usada como referencia para o título da série
Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(7, 2).Value = "Expectativa de exportação"

Set GrafDemExt = Sheets("GRÁFICO").Shapes.AddChart2 'Adiciona o gráfico

Sheets("GRÁFICO").Select 'Seleciona a aba gráfico

GrafDemExt.Select ' Seleciona o Gráfico
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confederação Nacional da Indústria\ECON - 1 Indicadores Econômicos CNI\1 Indicadores de Atividade Industrial\Sondagem Industrial\Automação\Templates\Expectativa - Compra e empregados.crtx") ' Aplica o template do gráfico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gráfico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gráfico
    ActiveChart.Parent.Top = Parent.Range("J29").Top 'reposiciona o grafico em relação ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("J29").Left ' reposiciona o gráfico em relação à borda esquerda da planilha
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova série ao gráfico
    ActiveChart.FullSeriesCollection(1).Name = "='EXPECTATIVAS - DEMANDA'!" & Cells(7, 2).Address 'Determina o nome da série
    ActiveChart.FullSeriesCollection(1).Values = "='EXPECTATIVAS - DEMANDA'!" & Range(Cells(9, I), Cells(9, H)).Address 'determina os valores da série
    ActiveChart.FullSeriesCollection(1).XValues = "='EXPECTATIVAS - DEMANDA'!" & Range(Cells(8, I), Cells(8, H)).Address 'determina os valores referentes ao eixo x da série adicionada
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova série ao gráfico
    ActiveChart.FullSeriesCollection(2).Name = "='EXPECTATIVA - EXPORTAÇÃO'!" & Cells(7, 2).Address 'Determina o nome da série
    ActiveChart.FullSeriesCollection(2).Values = "='EXPECTATIVA - EXPORTAÇÃO'!" & Range(Cells(9, I - 12), Cells(9, H - 12)).Address 'determina os valores da série
    ActiveChart.FullSeriesCollection(2).XValues = "='EXPECTATIVA - EXPORTAÇÃO'!" & Range(Cells(8, I - 12), Cells(8, H - 12)).Address 'determina os valores referentes ao eixo x da série adicionada
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova série ao gráfico
    ActiveChart.FullSeriesCollection(3).Name = "='EXPECTATIVAS - DEMANDA'!$A$55" 'Determina o nome da série
    ActiveChart.FullSeriesCollection(3).Values = "='EXPECTATIVAS - DEMANDA'!" & Range(Cells(55, I), Cells(55, H)).Address 'determina os valores da série
    ActiveChart.FullSeriesCollection(3).XValues = "='EXPECTATIVAS - DEMANDA'!" & Range(Cells(8, I), Cells(8, H)).Address 'determina os valores referentes ao eixo x da série adicionada
   

'********************************************************  Gráfico Intenção de investimento    ********************************************************************

Dim L As Integer 'Número da última Coluna
Dim M As Integer 'Número da primeira Coluna
Dim GrafIntInv As Object 'Gráfico

L = Sheets("EXPECTATIVA - INVESTIMENTO").Range("A9").End(xlToRight).Column 'Define o número da última coluna
M = L - 84 'Define o número da primeira coluna

Sheets("EXPECTATIVA - INVESTIMENTO").Select 'Seleciona a aba EXPECTATIVA - INVESTIMENTO
Sheets("EXPECTATIVA - INVESTIMENTO").Cells(7, 2).Value = "Intenção de investimento" 'Nomeia a celula que será usada como referencia para o título da série

Set GrafIntInv = Sheets("GRÁFICO").Shapes.AddChart2 'Adiciona o gráfico

Sheets("GRÁFICO").Select 'Seleciona a aba gráfico

GrafIntInv.Select ' Seleciona o Gráfico
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confederação Nacional da Indústria\ECON - 1 Indicadores Econômicos CNI\1 Indicadores de Atividade Industrial\Sondagem Industrial\Automação\Templates\Investimento.crtx") ' Aplica o template do gráfico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gráfico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gráfico
    ActiveChart.Parent.Top = Parent.Range("S29").Top 'reposiciona o gráfico em relação ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("S29").Left ' reposiciona o gráfico em relação à borda esquerda da planilha
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova série ao gráfico
    ActiveChart.FullSeriesCollection(1).Name = "='EXPECTATIVA - INVESTIMENTO'!" & Cells(7, 2).Address 'Determina o nome da série
    ActiveChart.FullSeriesCollection(1).Values = "='EXPECTATIVA - INVESTIMENTO'!" & Range(Cells(9, M), Cells(9, L)).Address 'determina os valores da série
    ActiveChart.FullSeriesCollection(1).XValues = "='EXPECTATIVA - INVESTIMENTO'!" & Range(Cells(8, M), Cells(8, L)).Address 'determina os valores referentes ao eixo x da série adicionada
    

'********************************************************  Gráfico Credito    ********************************************************************

Dim Q As Integer 'Número da última Coluna
Dim R As Integer 'Número da primeira coluna
Dim GrafCredito As Object 'Gráfico

Q = Sheets("SITUACAO FINANCEIRA CREDITO").Range("A9").End(xlToRight).Column 'Define o número da última coluna
R = Q - 36 'Define o número da primeira coluna

Sheets("SITUACAO FINANCEIRA CREDITO").Select 'Seleciona a aba SITUACAO FINANCEIRA CREDITO
Sheets("SITUACAO FINANCEIRA CREDITO").Range(Cells(55, R), Cells(55, Q)).Value = "50" ' Insere a série da linha divisória
Sheets("SITUACAO FINANCEIRA CREDITO").Cells(55, 1).Value = "Linha divisória" 'Nomeia a celula que será usada como referencia para o título da série
Sheets("SITUACAO FINANCEIRA CREDITO").Cells(7, 2).Value = "Facilidade de acesso ao crédito" 'Nomeia a celula que será usada como referencia para o título da série

Set GrafCredito = Sheets("GRÁFICO").Shapes.AddChart2 'Adiciona o gráfico
Sheets("GRÁFICO").Select 'Seleciona a aba gráfico
GrafCredito.Select ' Seleciona o Gráfico
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confederação Nacional da Indústria\ECON - 1 Indicadores Econômicos CNI\1 Indicadores de Atividade Industrial\Sondagem Industrial\Automação\Templates\Crédito.crtx") ' Aplica o template do gráfico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gráfico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gráfico
    ActiveChart.Parent.Top = Parent.Range("A55").Top 'reposiciona o grafico em relação ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("A55").Left ' reposiciona o gráfico em relação à borda esquerda da planilha
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova série ao gráfico
    ActiveChart.FullSeriesCollection(1).Name = "='SITUACAO FINANCEIRA CREDITO'!" & Cells(7, 2).Address   'Determina o nome da série
    ActiveChart.FullSeriesCollection(1).Values = "='SITUACAO FINANCEIRA CREDITO'!" & Range(Cells(9, R), Cells(9, Q)).Address 'determina os valores da série
    ActiveChart.FullSeriesCollection(1).XValues = "='SITUACAO FINANCEIRA CREDITO'!" & Range(Cells(8, R), Cells(8, Q)).Address 'determina os valores referentes ao eixo x da série adicionada
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova série ao gráfico
    ActiveChart.FullSeriesCollection(2).Name = "='SITUACAO FINANCEIRA CREDITO'!$A$55" 'Determina o nome da série
    ActiveChart.FullSeriesCollection(2).Values = "='SITUACAO FINANCEIRA CREDITO'!" & Range(Cells(55, R), Cells(55, Q)).Address 'determina os valores da série
    ActiveChart.FullSeriesCollection(2).XValues = "='SITUACAO FINANCEIRA CREDITO'!" & Range(Cells(8, R), Cells(8, Q)).Address 'determina os valores referentes ao eixo x da série adicionada

'********************************************************  Gráfico Preço Médio    ********************************************************************

Dim S As Integer 'Número da última Coluna
Dim T As Integer 'Número da primeira coluna
Dim GrafPM As Object 'Gráfico

S = Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Range("A9").End(xlToRight).Column 'Define o número da última coluna
T = S - 36 'Define o número da primeira coluna

Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Select 'Seleciona a aba SITUACAO FINANCEIRA PREÇO MEDIO
Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Range(Cells(55, T), Cells(55, S)).Value = "50" ' Insere a série da linha divisória
Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(55, 1).Value = "Linha divisória" 'Nomeia a celula que será usada como referencia para o título da série
Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(7, 2).Value = "Preço médio das matérias-primas" 'Nomeia a celula que será usada como referencia para o título da série

Set GrafPM = Sheets("GRÁFICO").Shapes.AddChart2 'Adiciona o gráfico
Sheets("GRÁFICO").Select 'Seleciona a aba gráfico
GrafPM.Select ' Seleciona o Gráfico
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confederação Nacional da Indústria\ECON - 1 Indicadores Econômicos CNI\1 Indicadores de Atividade Industrial\Sondagem Industrial\Automação\Templates\Preço.crtx") ' Aplica o template do gráfico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gráfico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gráfico
    ActiveChart.Parent.Top = Parent.Range("J55").Top 'reposiciona o grafico em relação ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("J55").Left ' reposiciona o gráfico em relação à borda esquerda da planilha
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova série ao gráfico
    ActiveChart.FullSeriesCollection(1).Name = "='SITUACAO FINANCEIRA PREÇO MEDIO'!" & Cells(7, 2).Address  'Determina o nome da série
    ActiveChart.FullSeriesCollection(1).Values = "='SITUACAO FINANCEIRA PREÇO MEDIO'!" & Range(Cells(9, T), Cells(9, S)).Address 'determina os valores da série
    ActiveChart.FullSeriesCollection(1).XValues = "='SITUACAO FINANCEIRA PREÇO MEDIO'!" & Range(Cells(8, T), Cells(8, S)).Address 'determina os valores referentes ao eixo x da série adicionada
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova série ao gráfico
    ActiveChart.FullSeriesCollection(2).Name = "='SITUACAO FINANCEIRA PREÇO MEDIO'!$A$55" 'Determina o nome da série
    ActiveChart.FullSeriesCollection(2).Values = "='SITUACAO FINANCEIRA PREÇO MEDIO'!" & Range(Cells(55, T), Cells(55, S)).Address 'determina os valores da série
    ActiveChart.FullSeriesCollection(2).XValues = "='SITUACAO FINANCEIRA PREÇO MEDIO'!" & Range(Cells(8, T), Cells(8, S)).Address 'determina os valores referentes ao eixo x da série adicionada
    ActiveChart.Axes(xlValue).MinimumScale = 40
    ActiveChart.Axes(xlValue).MaximumScale = 85


'********************************************************  Gráfico Lucro    ********************************************************************

Dim N As Integer 'Número da última Coluna
Dim O As Integer 'Número da primeira coluna
Dim GrafSFL As Object 'Gráfico

N = Sheets("SITUACAO FINANCEIRA LUCRO").Range("A9").End(xlToRight).Column 'Define o número da última coluna
O = N - 36 'Define o número da primeira coluna

Sheets("SITUACAO FINANCEIRA LUCRO").Select 'Seleciona a aba Produção
Sheets("SITUACAO FINANCEIRA LUCRO").Range(Cells(55, O), Cells(55, N)).Value = "50" ' Insere a série da linha divisória
Sheets("SITUACAO FINANCEIRA LUCRO").Cells(55, 1).Value = "Linha divisória" 'Nomeia a celula que será usada como referencia para o título da série
Sheets("SITUACAO FINANCEIRA LUCRO").Cells(7, 2).Value = "Lucro Operacional" 'Nomeia a celula que será usada como referencia para o título da série
Sheets("SITUACAO FINANCEIRA").Cells(7, 2).Value = "Situação financeira" 'Nomeia a celula que será usada como referencia para o título da série

Set GrafSFL = Sheets("GRÁFICO").Shapes.AddChart2 'Adiciona o gráfico
Sheets("GRÁFICO").Select 'Seleciona a aba gráfico
GrafSFL.Select ' Seleciona o Gráfico
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confederação Nacional da Indústria\ECON - 1 Indicadores Econômicos CNI\1 Indicadores de Atividade Industrial\Sondagem Industrial\Automação\Templates\Lucro e situação financeira.crtx") ' Aplica o template do gráfico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gráfico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gráfico
    ActiveChart.Parent.Top = Parent.Range("S55").Top 'reposiciona o grafico em relação ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("S55").Left ' reposiciona o gráfico em relação à borda esquerda da planilha
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova série ao gráfico
    ActiveChart.FullSeriesCollection(1).Name = "='SITUACAO FINANCEIRA LUCRO'!" & Cells(7, 2).Address 'Determina o nome da série
    ActiveChart.FullSeriesCollection(1).Values = "='SITUACAO FINANCEIRA LUCRO'!" & Range(Cells(9, O), Cells(9, N)).Address 'determina os valores da série
    ActiveChart.FullSeriesCollection(1).XValues = "='SITUACAO FINANCEIRA LUCRO'!" & Range(Cells(8, O), Cells(8, N)).Address 'determina os valores referentes ao eixo x da série adicionada
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova série ao gráfico
    ActiveChart.FullSeriesCollection(2).Name = "='SITUACAO FINANCEIRA LUCRO'!$A$55" 'Determina o nome da série
    ActiveChart.FullSeriesCollection(2).Values = "='SITUACAO FINANCEIRA LUCRO'!" & Range(Cells(55, O), Cells(55, N)).Address 'determina os valores da série
    ActiveChart.FullSeriesCollection(2).XValues = "='SITUACAO FINANCEIRA LUCRO'!" & Range(Cells(8, O), Cells(8, N)).Address 'determina os valores referentes ao eixo x da série adicionada
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova série ao gráfico
    ActiveChart.FullSeriesCollection(3).Name = "='SITUACAO FINANCEIRA'!" & Cells(7, 2).Address 'Determina o nome da série
    ActiveChart.FullSeriesCollection(3).Values = "='SITUACAO FINANCEIRA'!" & Range(Cells(9, O), Cells(9, N)).Address 'determina os valores da série
    ActiveChart.FullSeriesCollection(3).XValues = "='SITUACAO FINANCEIRA'!" & Range(Cells(8, O), Cells(8, N)).Address 'determina os valores referentes ao eixo x da série adicionada


'********************************************************  Gráfico Principais problemas    ********************************************************************
 
Sheets("Principais_Problemas").Select ' Seleciona a aba Principais_Problemas
ActiveSheet.Range("B12:B28").Copy ActiveSheet.Range("B110") ' Copia e cola o nome das categorias menos outros e nehum.

Dim V As Integer 'Numero do trimestre mais recente
Dim X As Integer 'Número do trimestre anterior
Dim GrafProblemas As Object ' Gráfico
 
V = Sheets("Principais_Problemas").Range("B13").End(xlToRight).Column 'Define o número da última coluna
X = V - 1 'Define o número da primeira coluna

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

Set GrafProblemas = Sheets("GRÁFICO").Shapes.AddChart2 'Adiciona o gráfico
Sheets("GRÁFICO").Select 'Seleciona a aba gráfico
GrafProblemas.Select ' Seleciona o Gráfico
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confederação Nacional da Indústria\ECON - 1 Indicadores Econômicos CNI\1 Indicadores de Atividade Industrial\Sondagem Industrial\Automação\Templates\Principais Problemas.crtx") ' Aplica o template do gráfico
    ActiveChart.Parent.Height = 630 'ajusta a altura do gráfico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gráfico
    ActiveChart.Parent.Top = Parent.Range("AC29").Top 'reposiciona o grafico em relação ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("AC29").Left ' reposiciona o gráfico em relação à borda esquerda da planilha
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova série ao gráfico
    ActiveChart.FullSeriesCollection(1).Name = "='PRINCIPAIS_PROBLEMAS'!$D$110" 'Determina o nome da série
    ActiveChart.FullSeriesCollection(1).Values = "='PRINCIPAIS_PROBLEMAS'!$D$111:$D$128" 'determina os valores da série
    ActiveChart.FullSeriesCollection(1).XValues = "='PRINCIPAIS_PROBLEMAS'!$B$111:$B$128" 'determina os valores referentes ao eixo x da série adicionada
    ActiveChart.SeriesCollection.NewSeries 'adiciona uma nova série ao gráfico
    ActiveChart.FullSeriesCollection(2).Name = "='PRINCIPAIS_PROBLEMAS'!$C$110" 'Determina o nome da série
    ActiveChart.FullSeriesCollection(2).Values = "='PRINCIPAIS_PROBLEMAS'!$C$111:$C$128" 'determina os valores da série
    ActiveChart.FullSeriesCollection(2).XValues = "='PRINCIPAIS_PROBLEMAS'!$B$111:$B$128" 'determina os valores referentes ao eixo x da série adicionada

End Sub

Sub Tabelas()

'Adiciona a aba tabela
Sheets.Add(Before:=Sheets("PRODUÇÃO")).Name = "TABELAS"

'Nomeia os titulos das colunas e mescla as celulas
Sheets("TABELAS").Cells(1, 1).Value = "Desempenho da indústria"

Sheets("TABELAS").Range(Cells(2, 2), Cells(3, 4)).Merge
Sheets("TABELAS").Cells(2, 2).Value = "EVOLUÇÃO DA PRODUÇÃO"

Sheets("TABELAS").Range(Cells(2, 5), Cells(3, 7)).Merge
Sheets("TABELAS").Cells(2, 5).Value = "EVOLUÇÃO DO NO DE EMPREGADOS"

Sheets("TABELAS").Range(Cells(2, 8), Cells(3, 10)).Merge
Sheets("TABELAS").Cells(2, 8).Value = "UCI (%)"

Sheets("TABELAS").Range(Cells(2, 11), Cells(3, 13)).Merge
Sheets("TABELAS").Cells(2, 11).Value = " UCI EFETIVA-USUAL"

Sheets("TABELAS").Range(Cells(2, 14), Cells(3, 16)).Merge
Sheets("TABELAS").Cells(2, 14).Value = "EVOLUÇÃO DOS ESTOQUES"

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
Sheets("TABELAS").Range("A5").Value = "Indústria Geral"
Sheets("TABELAS").Range("A8").Value = "Indústria extrativa"
Sheets("TABELAS").Range("A9").Value = "Indústria de transformação"
Sheets("TABELAS").Range("A12").Value = "Pequena"
Sheets("TABELAS").Range("A13").Value = "Média"
Sheets("TABELAS").Range("A14").Value = "Grande"

'Define as variavies que serão usadas para preencher as celulas
Coluna_Produção_1 = Sheets("PRODUÇÃO").Range("B8").End(xlToRight).Column
Coluna_Produção_2 = Coluna_Produção_1 - 1
Coluna_Produção_3 = Coluna_Produção_1 - 12

'Define, atribui e copia e cola as datas
Datas_1 = Sheets("PRODUÇÃO").Cells(8, Coluna_Produção_1).Value
Datas_2 = Sheets("PRODUÇÃO").Cells(8, Coluna_Produção_2).Value
Datas_3 = Sheets("PRODUÇÃO").Cells(8, Coluna_Produção_3).Value

Sheets("TABELAS").Cells(4, 2).Value = Datas_3
Sheets("TABELAS").Cells(4, 3).Value = Datas_2
Sheets("TABELAS").Cells(4, 4).Value = Datas_1

Sheets("TABELAS").Range("B4:D4").Copy (Sheets("TABELAS").Range("E4:G4"))
Sheets("TABELAS").Range("B4:D4").Copy (Sheets("TABELAS").Range("H4:J4"))
Sheets("TABELAS").Range("B4:D4").Copy (Sheets("TABELAS").Range("K4:M4"))
Sheets("TABELAS").Range("B4:D4").Copy (Sheets("TABELAS").Range("N4:P4"))
Sheets("TABELAS").Range("B4:D4").Copy (Sheets("TABELAS").Range("Q4:S4"))

'Atribui os valores da coluna Evolução da produção
'Indústria Geral
ValoresIGP_1 = Sheets("PRODUÇÃO").Cells(9, Coluna_Produção_1).Value
ValoresIGP_2 = Sheets("PRODUÇÃO").Cells(9, Coluna_Produção_2).Value
ValoresIGP_3 = Sheets("PRODUÇÃO").Cells(9, Coluna_Produção_3).Value
Sheets("TABELAS").Cells(5, 2).Value = ValoresIGP_3
Sheets("TABELAS").Cells(5, 3).Value = ValoresIGP_2
Sheets("TABELAS").Cells(5, 4).Value = ValoresIGP_1
'Indústria Extrativa
ValoresIEP_1 = Sheets("PRODUÇÃO").Cells(21, Coluna_Produção_1).Value
ValoresIEP_2 = Sheets("PRODUÇÃO").Cells(21, Coluna_Produção_2).Value
ValoresIEP_3 = Sheets("PRODUÇÃO").Cells(21, Coluna_Produção_3).Value
Sheets("TABELAS").Cells(8, 2).Value = ValoresIEP_3
Sheets("TABELAS").Cells(8, 3).Value = ValoresIEP_2
Sheets("TABELAS").Cells(8, 4).Value = ValoresIEP_1
'Indústria da Transformação
ValoresITP_1 = Sheets("PRODUÇÃO").Cells(26, Coluna_Produção_1).Value
ValoresITP_2 = Sheets("PRODUÇÃO").Cells(26, Coluna_Produção_2).Value
ValoresITP_3 = Sheets("PRODUÇÃO").Cells(26, Coluna_Produção_3).Value
Sheets("TABELAS").Cells(9, 2).Value = ValoresITP_3
Sheets("TABELAS").Cells(9, 3).Value = ValoresITP_2
Sheets("TABELAS").Cells(9, 4).Value = ValoresITP_1
'Pequena
ValoresPP_1 = Sheets("PRODUÇÃO").Cells(17, Coluna_Produção_1).Value
ValoresPP_2 = Sheets("PRODUÇÃO").Cells(17, Coluna_Produção_2).Value
ValoresPP_3 = Sheets("PRODUÇÃO").Cells(17, Coluna_Produção_3).Value
Sheets("TABELAS").Cells(12, 2).Value = ValoresPP_3
Sheets("TABELAS").Cells(12, 3).Value = ValoresPP_2
Sheets("TABELAS").Cells(12, 4).Value = ValoresPP_1
'Média
ValoresMP_1 = Sheets("PRODUÇÃO").Cells(18, Coluna_Produção_1).Value
ValoresMP_2 = Sheets("PRODUÇÃO").Cells(18, Coluna_Produção_2).Value
ValoresMP_3 = Sheets("PRODUÇÃO").Cells(18, Coluna_Produção_3).Value
Sheets("TABELAS").Cells(13, 2).Value = ValoresMP_3
Sheets("TABELAS").Cells(13, 3).Value = ValoresMP_2
Sheets("TABELAS").Cells(13, 4).Value = ValoresMP_1
'Grande
ValoresGP_1 = Sheets("PRODUÇÃO").Cells(19, Coluna_Produção_1).Value
ValoresGP_2 = Sheets("PRODUÇÃO").Cells(19, Coluna_Produção_2).Value
ValoresGP_3 = Sheets("PRODUÇÃO").Cells(19, Coluna_Produção_3).Value
Sheets("TABELAS").Cells(14, 2).Value = ValoresGP_3
Sheets("TABELAS").Cells(14, 3).Value = ValoresGP_2
Sheets("TABELAS").Cells(14, 4).Value = ValoresGP_1

'Atribui os valores da coluna Evolução do Nº de Empregoados
Coluna_Emprego_1 = Sheets("EMPREGADOS").Range("B8").End(xlToRight).Column
Coluna_Emprego_2 = Coluna_Emprego_1 - 1
Coluna_Emprego_3 = Coluna_Emprego_1 - 12

'Indústria Geral
ValoresIGE_1 = Sheets("EMPREGADOS").Cells(9, Coluna_Emprego_1).Value
ValoresIGE_2 = Sheets("EMPREGADOS").Cells(9, Coluna_Emprego_2).Value
ValoresIGE_3 = Sheets("EMPREGADOS").Cells(9, Coluna_Emprego_3).Value
Sheets("TABELAS").Cells(5, 5).Value = ValoresIGE_3
Sheets("TABELAS").Cells(5, 6).Value = ValoresIGE_2
Sheets("TABELAS").Cells(5, 7).Value = ValoresIGE_1
'Indústria Extrativa
ValoresIEE_1 = Sheets("EMPREGADOS").Cells(21, Coluna_Emprego_1).Value
ValoresIEE_2 = Sheets("EMPREGADOS").Cells(21, Coluna_Emprego_2).Value
ValoresIEE_3 = Sheets("EMPREGADOS").Cells(21, Coluna_Emprego_3).Value
Sheets("TABELAS").Cells(8, 5).Value = ValoresIEE_3
Sheets("TABELAS").Cells(8, 6).Value = ValoresIEE_2
Sheets("TABELAS").Cells(8, 7).Value = ValoresIEE_1
'Indústria Transformação
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
'Média
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

'Indústria Geral
ValoresIG_UCI_1 = Sheets("UCI (%)").Cells(9, Coluna_UCI_1).Value
ValoresIG_UCI_2 = Sheets("UCI (%)").Cells(9, Coluna_UCI_2).Value
ValoresIG_UCI_3 = Sheets("UCI (%)").Cells(9, Coluna_UCI_3).Value
Sheets("TABELAS").Cells(5, 8).Value = ValoresIG_UCI_3
Sheets("TABELAS").Cells(5, 9).Value = ValoresIG_UCI_2
Sheets("TABELAS").Cells(5, 10).Value = ValoresIG_UCI_1
'Indústria extrativa
ValoresIE_UCI_1 = Sheets("UCI (%)").Cells(21, Coluna_UCI_1).Value
ValoresIE_UCI_2 = Sheets("UCI (%)").Cells(21, Coluna_UCI_2).Value
ValoresIE_UCI_3 = Sheets("UCI (%)").Cells(21, Coluna_UCI_3).Value
Sheets("TABELAS").Cells(8, 8).Value = ValoresIE_UCI_3
Sheets("TABELAS").Cells(8, 9).Value = ValoresIE_UCI_2
Sheets("TABELAS").Cells(8, 10).Value = ValoresIE_UCI_1
'Indústria Transformação
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
'Média
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

'Indústria Geral
ValoresIG_UCI_EU_1 = Sheets("UCI (efetiva-usual)").Cells(9, Coluna_UCI_EU_1).Value
ValoresIG_UCI_EU_2 = Sheets("UCI (efetiva-usual)").Cells(9, Coluna_UCI_EU_2).Value
ValoresIG_UCI_EU_3 = Sheets("UCI (efetiva-usual)").Cells(9, Coluna_UCI_EU_3).Value
Sheets("TABELAS").Cells(5, 11).Value = ValoresIG_UCI_EU_3
Sheets("TABELAS").Cells(5, 12).Value = ValoresIG_UCI_EU_2
Sheets("TABELAS").Cells(5, 13).Value = ValoresIG_UCI_EU_1
'Indústria Extrativa
ValoresIE_UCI_EU_1 = Sheets("UCI (efetiva-usual)").Cells(21, Coluna_UCI_EU_1).Value
ValoresIE_UCI_EU_2 = Sheets("UCI (efetiva-usual)").Cells(21, Coluna_UCI_EU_2).Value
ValoresIE_UCI_EU_3 = Sheets("UCI (efetiva-usual)").Cells(21, Coluna_UCI_EU_3).Value
Sheets("TABELAS").Cells(8, 11).Value = ValoresIE_UCI_EU_3
Sheets("TABELAS").Cells(8, 12).Value = ValoresIE_UCI_EU_2
Sheets("TABELAS").Cells(8, 13).Value = ValoresIE_UCI_EU_1
'Indústria transformação
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
'Média
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

'Atribui os valores da coluna Evolução dos estoques
Coluna_Estoques_1 = Sheets("ESTOQUES (evolução)").Range("B8").End(xlToRight).Column
Coluna_Estoques_2 = Coluna_Estoques_1 - 1
Coluna_Estoques_3 = Coluna_Estoques_1 - 12

'Indústria Geral
ValoresIG_Estoques_1 = Sheets("ESTOQUES (evolução)").Cells(9, Coluna_Estoques_1).Value
ValoresIG_Estoques_2 = Sheets("ESTOQUES (evolução)").Cells(9, Coluna_Estoques_2).Value
ValoresIG_Estoques_3 = Sheets("ESTOQUES (evolução)").Cells(9, Coluna_Estoques_3).Value
Sheets("TABELAS").Cells(5, 14).Value = ValoresIG_Estoques_3
Sheets("TABELAS").Cells(5, 15).Value = ValoresIG_Estoques_2
Sheets("TABELAS").Cells(5, 16).Value = ValoresIG_Estoques_1
'Indústria Extrativa
ValoresIE_Estoques_1 = Sheets("ESTOQUES (evolução)").Cells(21, Coluna_Estoques_1).Value
ValoresIE_Estoques_2 = Sheets("ESTOQUES (evolução)").Cells(21, Coluna_Estoques_2).Value
ValoresIE_Estoques_3 = Sheets("ESTOQUES (evolução)").Cells(21, Coluna_Estoques_3).Value
Sheets("TABELAS").Cells(8, 14).Value = ValoresIE_Estoques_3
Sheets("TABELAS").Cells(8, 15).Value = ValoresIE_Estoques_2
Sheets("TABELAS").Cells(8, 16).Value = ValoresIE_Estoques_1
'Indústria Transformação
ValoresIT_Estoques_1 = Sheets("ESTOQUES (evolução)").Cells(26, Coluna_Estoques_1).Value
ValoresIT_Estoques_2 = Sheets("ESTOQUES (evolução)").Cells(26, Coluna_Estoques_2).Value
ValoresIT_Estoques_3 = Sheets("ESTOQUES (evolução)").Cells(26, Coluna_Estoques_3).Value
Sheets("TABELAS").Cells(9, 14).Value = ValoresIT_Estoques_3
Sheets("TABELAS").Cells(9, 15).Value = ValoresIT_Estoques_2
Sheets("TABELAS").Cells(9, 16).Value = ValoresIT_Estoques_1
'Pequena
ValoresP_Estoques_1 = Sheets("ESTOQUES (evolução)").Cells(17, Coluna_Estoques_1).Value
ValoresP_Estoques_2 = Sheets("ESTOQUES (evolução)").Cells(17, Coluna_Estoques_2).Value
ValoresP_Estoques_3 = Sheets("ESTOQUES (evolução)").Cells(17, Coluna_Estoques_3).Value
Sheets("TABELAS").Cells(12, 14).Value = ValoresP_Estoques_3
Sheets("TABELAS").Cells(12, 15).Value = ValoresP_Estoques_2
Sheets("TABELAS").Cells(12, 16).Value = ValoresP_Estoques_1
'Média
ValoresM_Estoques_1 = Sheets("ESTOQUES (evolução)").Cells(18, Coluna_Estoques_1).Value
ValoresM_Estoques_2 = Sheets("ESTOQUES (evolução)").Cells(18, Coluna_Estoques_2).Value
ValoresM_Estoques_3 = Sheets("ESTOQUES (evolução)").Cells(18, Coluna_Estoques_3).Value
Sheets("TABELAS").Cells(13, 14).Value = ValoresM_Estoques_3
Sheets("TABELAS").Cells(13, 15).Value = ValoresM_Estoques_2
Sheets("TABELAS").Cells(13, 16).Value = ValoresM_Estoques_1
'Grande
ValoresG_Estoques_1 = Sheets("ESTOQUES (evolução)").Cells(19, Coluna_Estoques_1).Value
ValoresG_Estoques_2 = Sheets("ESTOQUES (evolução)").Cells(19, Coluna_Estoques_2).Value
ValoresG_Estoques_3 = Sheets("ESTOQUES (evolução)").Cells(19, Coluna_Estoques_3).Value
Sheets("TABELAS").Cells(14, 14).Value = ValoresG_Estoques_3
Sheets("TABELAS").Cells(14, 15).Value = ValoresG_Estoques_2
Sheets("TABELAS").Cells(14, 16).Value = ValoresG_Estoques_1

'Atribui os valores da coluna Estoque efetivo-planejado
Coluna_Estoques_EP_1 = Sheets("ESTOQUES (efetivo-planejado)").Range("B8").End(xlToRight).Column
Coluna_Estoques_EP_2 = Coluna_Estoques_EP_1 - 1
Coluna_Estoques_EP_3 = Coluna_Estoques_EP_1 - 12

'Indústria Geral
ValoresIG_Estoques_EP_1 = Sheets("ESTOQUES (efetivo-planejado)").Cells(9, Coluna_Estoques_EP_1).Value
ValoresIG_Estoques_EP_2 = Sheets("ESTOQUES (efetivo-planejado)").Cells(9, Coluna_Estoques_EP_2).Value
ValoresIG_Estoques_EP_3 = Sheets("ESTOQUES (efetivo-planejado)").Cells(9, Coluna_Estoques_EP_3).Value
Sheets("TABELAS").Cells(5, 17).Value = ValoresIG_Estoques_EP_3
Sheets("TABELAS").Cells(5, 18).Value = ValoresIG_Estoques_EP_2
Sheets("TABELAS").Cells(5, 19).Value = ValoresIG_Estoques_EP_1
'Indústria extrativa
ValoresIE_Estoques_EP_1 = Sheets("ESTOQUES (efetivo-planejado)").Cells(21, Coluna_Estoques_EP_1).Value
ValoresIE_Estoques_EP_2 = Sheets("ESTOQUES (efetivo-planejado)").Cells(21, Coluna_Estoques_EP_2).Value
ValoresIE_Estoques_EP_3 = Sheets("ESTOQUES (efetivo-planejado)").Cells(21, Coluna_Estoques_EP_3).Value
Sheets("TABELAS").Cells(8, 17).Value = ValoresIE_Estoques_EP_3
Sheets("TABELAS").Cells(8, 18).Value = ValoresIE_Estoques_EP_2
Sheets("TABELAS").Cells(8, 19).Value = ValoresIE_Estoques_EP_1
'Indústria Transformação
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
'Média
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

'*************************************************** Código da parte de Expectativas **********************************************************

'Nomeia os titulos das colunas e mescla as celulas
Sheets("TABELAS").Cells(16, 1).Value = "Expectativas da indústria"

Sheets("TABELAS").Range(Cells(17, 2), Cells(18, 4)).Merge
Sheets("TABELAS").Cells(17, 2).Value = "DEMANDA"

Sheets("TABELAS").Range(Cells(17, 5), Cells(18, 7)).Merge
Sheets("TABELAS").Cells(17, 5).Value = "QUANTIDADE EXPORTADA"

Sheets("TABELAS").Range(Cells(17, 8), Cells(18, 10)).Merge
Sheets("TABELAS").Cells(17, 8).Value = "COMPRAS DE MATÉRIA-PRIMA"

Sheets("TABELAS").Range(Cells(17, 11), Cells(18, 13)).Merge
Sheets("TABELAS").Cells(17, 11).Value = "Nº DE EMPREGADOS"

Sheets("TABELAS").Range(Cells(17, 14), Cells(18, 16)).Merge
Sheets("TABELAS").Cells(17, 14).Value = "INTENÇÃO DE INVESTIMENTO"

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
Sheets("TABELAS").Range("A20").Value = "Indústria Geral"
Sheets("TABELAS").Range("A23").Value = "Indústria extrativa"
Sheets("TABELAS").Range("A24").Value = "Indústria de transformação"
Sheets("TABELAS").Range("A27").Value = "Pequena"
Sheets("TABELAS").Range("A28").Value = "Média"
Sheets("TABELAS").Range("A29").Value = "Grande"

'Define as variavies que serão usadas para preencher as celulas
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
'Indústria Geral
ValoresIG_Demanda_1 = Sheets("EXPECTATIVAS - DEMANDA").Cells(9, Coluna_Demanda_1).Value
ValoresIG_Demanda_2 = Sheets("EXPECTATIVAS - DEMANDA").Cells(9, Coluna_Demanda_2).Value
ValoresIG_Demanda_3 = Sheets("EXPECTATIVAS - DEMANDA").Cells(9, Coluna_Demanda_3).Value
Sheets("TABELAS").Cells(20, 2).Value = ValoresIG_Demanda_3
Sheets("TABELAS").Cells(20, 3).Value = ValoresIG_Demanda_2
Sheets("TABELAS").Cells(20, 4).Value = ValoresIG_Demanda_1
'Indústria Extrativa
ValoresIE_Demanda_1 = Sheets("EXPECTATIVAS - DEMANDA").Cells(21, Coluna_Demanda_1).Value
ValoresIE_Demanda_2 = Sheets("EXPECTATIVAS - DEMANDA").Cells(21, Coluna_Demanda_2).Value
ValoresIE_Demanda_3 = Sheets("EXPECTATIVAS - DEMANDA").Cells(21, Coluna_Demanda_3).Value
Sheets("TABELAS").Cells(23, 2).Value = ValoresIE_Demanda_3
Sheets("TABELAS").Cells(23, 3).Value = ValoresIE_Demanda_2
Sheets("TABELAS").Cells(23, 4).Value = ValoresIE_Demanda_1
'Indústria Transformação
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
'Média
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
Coluna_Exportação_1 = Sheets("EXPECTATIVA - EXPORTAÇÃO").Range("B8").End(xlToRight).Column
Coluna_Exportação_2 = Coluna_Exportação_1 - 1
Coluna_Exportação_3 = Coluna_Exportação_1 - 12

'Indústria Geral
ValoresIG_Exportação_1 = Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(9, Coluna_Exportação_1).Value
ValoresIG_Exportação_2 = Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(9, Coluna_Exportação_2).Value
ValoresIG_Exportação_3 = Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(9, Coluna_Exportação_3).Value
Sheets("TABELAS").Cells(20, 5).Value = ValoresIG_Exportação_3
Sheets("TABELAS").Cells(20, 6).Value = ValoresIG_Exportação_2
Sheets("TABELAS").Cells(20, 7).Value = ValoresIG_Exportação_1
'Indústria Extrativa
ValoresIE_Exportação_1 = Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(21, Coluna_Exportação_1).Value
ValoresIE_Exportação_2 = Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(21, Coluna_Exportação_2).Value
ValoresIE_Exportação_3 = Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(21, Coluna_Exportação_3).Value
Sheets("TABELAS").Cells(23, 5).Value = ValoresIE_Exportação_3
Sheets("TABELAS").Cells(23, 6).Value = ValoresIE_Exportação_2
Sheets("TABELAS").Cells(23, 7).Value = ValoresIE_Exportação_1
'Indústria Tansformação
ValoresIT_Exportação_1 = Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(26, Coluna_Exportação_1).Value
ValoresIT_Exportação_2 = Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(26, Coluna_Exportação_2).Value
ValoresIT_Exportação_3 = Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(26, Coluna_Exportação_3).Value
Sheets("TABELAS").Cells(24, 5).Value = ValoresIT_Exportação_3
Sheets("TABELAS").Cells(24, 6).Value = ValoresIT_Exportação_2
Sheets("TABELAS").Cells(24, 7).Value = ValoresIT_Exportação_1
'Pequena
ValoresP_Exportação_1 = Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(17, Coluna_Exportação_1).Value
ValoresP_Exportação_2 = Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(17, Coluna_Exportação_2).Value
ValoresP_Exportação_3 = Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(17, Coluna_Exportação_3).Value
Sheets("TABELAS").Cells(27, 5).Value = ValoresP_Exportação_3
Sheets("TABELAS").Cells(27, 6).Value = ValoresP_Exportação_2
Sheets("TABELAS").Cells(27, 7).Value = ValoresP_Exportação_1
'Média
ValoresM_Exportação_1 = Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(18, Coluna_Exportação_1).Value
ValoresM_Exportação_2 = Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(18, Coluna_Exportação_2).Value
ValoresM_Exportação_3 = Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(18, Coluna_Exportação_3).Value
Sheets("TABELAS").Cells(28, 5).Value = ValoresM_Exportação_3
Sheets("TABELAS").Cells(28, 6).Value = ValoresM_Exportação_2
Sheets("TABELAS").Cells(28, 7).Value = ValoresM_Exportação_1
'Grande
ValoresG_Exportação_1 = Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(19, Coluna_Exportação_1).Value
ValoresG_Exportação_2 = Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(19, Coluna_Exportação_2).Value
ValoresG_Exportação_3 = Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(19, Coluna_Exportação_3).Value
Sheets("TABELAS").Cells(29, 5).Value = ValoresG_Exportação_3
Sheets("TABELAS").Cells(29, 6).Value = ValoresG_Exportação_2
Sheets("TABELAS").Cells(29, 7).Value = ValoresG_Exportação_1

'Atribui os valores da coluna Compras de matéria prima
Coluna_Compras_1 = Sheets("EXPECTATIVA - COMPRAS").Range("B8").End(xlToRight).Column
Coluna_Compras_2 = Coluna_Compras_1 - 1
Coluna_Compras_3 = Coluna_Compras_1 - 12

'Indústria Geral
ValoresIG_Compras_1 = Sheets("EXPECTATIVA - COMPRAS").Cells(9, Coluna_Compras_1).Value
ValoresIG_Compras_2 = Sheets("EXPECTATIVA - COMPRAS").Cells(9, Coluna_Compras_2).Value
ValoresIG_Compras_3 = Sheets("EXPECTATIVA - COMPRAS").Cells(9, Coluna_Compras_3).Value
Sheets("TABELAS").Cells(20, 8).Value = ValoresIG_Compras_3
Sheets("TABELAS").Cells(20, 9).Value = ValoresIG_Compras_2
Sheets("TABELAS").Cells(20, 10).Value = ValoresIG_Compras_1
'Indústria Extrativa
ValoresIE_Compras_1 = Sheets("EXPECTATIVA - COMPRAS").Cells(21, Coluna_Compras_1).Value
ValoresIE_Compras_2 = Sheets("EXPECTATIVA - COMPRAS").Cells(21, Coluna_Compras_2).Value
ValoresIE_Compras_3 = Sheets("EXPECTATIVA - COMPRAS").Cells(21, Coluna_Compras_3).Value
Sheets("TABELAS").Cells(23, 8).Value = ValoresIE_Compras_3
Sheets("TABELAS").Cells(23, 9).Value = ValoresIE_Compras_2
Sheets("TABELAS").Cells(23, 10).Value = ValoresIE_Compras_1
'Indústria Tranformação
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
'Média
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

'Atribui os valores da coluna Nº de empregados
Coluna_EXEmpregados_1 = Sheets("EXPECTATIVA - EMPREGADOS").Range("B8").End(xlToRight).Column
Coluna_EXEmpregados_2 = Coluna_EXEmpregados_1 - 1
Coluna_EXEmpregados_3 = Coluna_EXEmpregados_1 - 12

'Indústria Geral
ValoresIG_EXEmpregados_1 = Sheets("EXPECTATIVA - EMPREGADOS").Cells(9, Coluna_EXEmpregados_1).Value
ValoresIG_EXEmpregados_2 = Sheets("EXPECTATIVA - EMPREGADOS").Cells(9, Coluna_EXEmpregados_2).Value
ValoresIG_EXEmpregados_3 = Sheets("EXPECTATIVA - EMPREGADOS").Cells(9, Coluna_EXEmpregados_3).Value
Sheets("TABELAS").Cells(20, 11).Value = ValoresIG_EXEmpregados_3
Sheets("TABELAS").Cells(20, 12).Value = ValoresIG_EXEmpregados_2
Sheets("TABELAS").Cells(20, 13).Value = ValoresIG_EXEmpregados_1
'Indústria Extrativa
ValoresIE_EXEmpregados_1 = Sheets("EXPECTATIVA - EMPREGADOS").Cells(21, Coluna_EXEmpregados_1).Value
ValoresIE_EXEmpregados_2 = Sheets("EXPECTATIVA - EMPREGADOS").Cells(21, Coluna_EXEmpregados_2).Value
ValoresIE_EXEmpregados_3 = Sheets("EXPECTATIVA - EMPREGADOS").Cells(21, Coluna_EXEmpregados_3).Value
Sheets("TABELAS").Cells(23, 11).Value = ValoresIE_EXEmpregados_3
Sheets("TABELAS").Cells(23, 12).Value = ValoresIE_EXEmpregados_2
Sheets("TABELAS").Cells(23, 13).Value = ValoresIE_EXEmpregados_1
'Indústria Transformação
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
'Média
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

'Atribui os valores da coluna Intenção de investimento
Coluna_Investimento_1 = Sheets("EXPECTATIVA - INVESTIMENTO").Range("B8").End(xlToRight).Column
Coluna_Investimento_2 = Coluna_Investimento_1 - 1
Coluna_Investimento_3 = Coluna_Investimento_1 - 12

'Indústria Geral
ValoresIG_Investimento_1 = Sheets("EXPECTATIVA - INVESTIMENTO").Cells(9, Coluna_Investimento_1).Value
ValoresIG_Investimento_2 = Sheets("EXPECTATIVA - INVESTIMENTO").Cells(9, Coluna_Investimento_2).Value
ValoresIG_Investimento_3 = Sheets("EXPECTATIVA - INVESTIMENTO").Cells(9, Coluna_Investimento_3).Value
Sheets("TABELAS").Cells(20, 14).Value = ValoresIG_Investimento_3
Sheets("TABELAS").Cells(20, 15).Value = ValoresIG_Investimento_2
Sheets("TABELAS").Cells(20, 16).Value = ValoresIG_Investimento_1
'Indústria Extrativa
ValoresIE_Investimento_1 = Sheets("EXPECTATIVA - INVESTIMENTO").Cells(21, Coluna_Investimento_1).Value
ValoresIE_Investimento_2 = Sheets("EXPECTATIVA - INVESTIMENTO").Cells(21, Coluna_Investimento_2).Value
ValoresIE_Investimento_3 = Sheets("EXPECTATIVA - INVESTIMENTO").Cells(21, Coluna_Investimento_3).Value
Sheets("TABELAS").Cells(23, 14).Value = ValoresIE_Investimento_3
Sheets("TABELAS").Cells(23, 15).Value = ValoresIE_Investimento_2
Sheets("TABELAS").Cells(23, 16).Value = ValoresIE_Investimento_1
'Indústria Transformação
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
'Média
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


'*************************************************** Código da parte de Condições Financeiras **********************************************************

'Nomeia os titulos das colunas e mescla as celulas
Sheets("TABELAS").Cells(31, 1).Value = "Condições Financeiras no trimestre"

Sheets("TABELAS").Range(Cells(32, 2), Cells(33, 4)).Merge
Sheets("TABELAS").Cells(32, 2).Value = "MARGEM DE LUCRO OPERACIONAL"

Sheets("TABELAS").Range(Cells(32, 5), Cells(33, 7)).Merge
Sheets("TABELAS").Cells(32, 5).Value = "PREÇO MÉDIO DAS MATÉRIAS-PRIMAS"

Sheets("TABELAS").Range(Cells(32, 8), Cells(33, 10)).Merge
Sheets("TABELAS").Cells(32, 8).Value = "SITUAÇÃO FINANCEIRA"

Sheets("TABELAS").Range(Cells(32, 11), Cells(33, 13)).Merge
Sheets("TABELAS").Cells(32, 11).Value = "ACESSO AO CRÉDITO"

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
Sheets("TABELAS").Range("A35").Value = "Indústria Geral"
Sheets("TABELAS").Range("A38").Value = "Indústria extrativa"
Sheets("TABELAS").Range("A39").Value = "Indústria de transformação"
Sheets("TABELAS").Range("A42").Value = "Pequena"
Sheets("TABELAS").Range("A43").Value = "Média"
Sheets("TABELAS").Range("A44").Value = "Grande"

'Define as variavies que serão usadas para preencher as celulas
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
'Indústria Geral
ValoresIG_Lucro_1 = Sheets("SITUACAO FINANCEIRA LUCRO").Cells(9, Coluna_Lucro_1).Value
ValoresIG_Lucro_2 = Sheets("SITUACAO FINANCEIRA LUCRO").Cells(9, Coluna_Lucro_2).Value
ValoresIG_Lucro_3 = Sheets("SITUACAO FINANCEIRA LUCRO").Cells(9, Coluna_Lucro_3).Value
Sheets("TABELAS").Cells(35, 2).Value = ValoresIG_Lucro_3
Sheets("TABELAS").Cells(35, 3).Value = ValoresIG_Lucro_2
Sheets("TABELAS").Cells(35, 4).Value = ValoresIG_Lucro_1
'Indústria Extrativa
ValoresIE_Lucro_1 = Sheets("SITUACAO FINANCEIRA LUCRO").Cells(21, Coluna_Lucro_1).Value
ValoresIE_Lucro_2 = Sheets("SITUACAO FINANCEIRA LUCRO").Cells(21, Coluna_Lucro_2).Value
ValoresIE_Lucro_3 = Sheets("SITUACAO FINANCEIRA LUCRO").Cells(21, Coluna_Lucro_3).Value
Sheets("TABELAS").Cells(38, 2).Value = ValoresIE_Lucro_3
Sheets("TABELAS").Cells(38, 3).Value = ValoresIE_Lucro_2
Sheets("TABELAS").Cells(38, 4).Value = ValoresIE_Lucro_1
'Indústria Transformação
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
'Média
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

'Atribui os valores da coluna Preço médio de matérias primas
Coluna_Preço_1 = Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Range("B8").End(xlToRight).Column
Coluna_Preço_2 = Coluna_Preço_1 - 1
Coluna_Preço_3 = Coluna_Preço_1 - 12

'Indústria Geral
ValoresIG_Preço_1 = Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(9, Coluna_Preço_1).Value
ValoresIG_Preço_2 = Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(9, Coluna_Preço_2).Value
ValoresIG_Preço_3 = Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(9, Coluna_Preço_3).Value
Sheets("TABELAS").Cells(35, 5).Value = ValoresIG_Preço_3
Sheets("TABELAS").Cells(35, 6).Value = ValoresIG_Preço_2
Sheets("TABELAS").Cells(35, 7).Value = ValoresIG_Preço_1
'Indústria Extrativa
ValoresIE_Preço_1 = Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(21, Coluna_Preço_1).Value
ValoresIE_Preço_2 = Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(21, Coluna_Preço_2).Value
ValoresIE_Preço_3 = Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(21, Coluna_Preço_3).Value
Sheets("TABELAS").Cells(38, 5).Value = ValoresIE_Preço_3
Sheets("TABELAS").Cells(38, 6).Value = ValoresIE_Preço_2
Sheets("TABELAS").Cells(38, 7).Value = ValoresIE_Preço_1
'Indústria Tansformação
ValoresIT_Preço_1 = Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(26, Coluna_Preço_1).Value
ValoresIT_Preço_2 = Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(26, Coluna_Preço_2).Value
ValoresIT_Preço_3 = Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(26, Coluna_Preço_3).Value
Sheets("TABELAS").Cells(39, 5).Value = ValoresIT_Preço_3
Sheets("TABELAS").Cells(39, 6).Value = ValoresIT_Preço_2
Sheets("TABELAS").Cells(39, 7).Value = ValoresIT_Preço_1
'Pequena
ValoresP_Preço_1 = Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(17, Coluna_Preço_1).Value
ValoresP_Preço_2 = Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(17, Coluna_Preço_2).Value
ValoresP_Preço_3 = Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(17, Coluna_Preço_3).Value
Sheets("TABELAS").Cells(42, 5).Value = ValoresP_Preço_3
Sheets("TABELAS").Cells(42, 6).Value = ValoresP_Preço_2
Sheets("TABELAS").Cells(42, 7).Value = ValoresP_Preço_1
'Média
ValoresM_Preço_1 = Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(18, Coluna_Preço_1).Value
ValoresM_Preço_2 = Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(18, Coluna_Preço_2).Value
ValoresM_Preço_3 = Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(18, Coluna_Preço_3).Value
Sheets("TABELAS").Cells(43, 5).Value = ValoresM_Preço_3
Sheets("TABELAS").Cells(43, 6).Value = ValoresM_Preço_2
Sheets("TABELAS").Cells(43, 7).Value = ValoresM_Preço_1
'Grande
ValoresG_Preço_1 = Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(19, Coluna_Preço_1).Value
ValoresG_Preço_2 = Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(19, Coluna_Preço_2).Value
ValoresG_Preço_3 = Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(19, Coluna_Preço_3).Value
Sheets("TABELAS").Cells(44, 5).Value = ValoresG_Preço_3
Sheets("TABELAS").Cells(44, 6).Value = ValoresG_Preço_2
Sheets("TABELAS").Cells(44, 7).Value = ValoresG_Preço_1

'Atribui os valores da coluna Situação FInanceira
Coluna_Situação_1 = Sheets("SITUACAO FINANCEIRA").Range("B8").End(xlToRight).Column
Coluna_Situação_2 = Coluna_Situação_1 - 1
Coluna_Situação_3 = Coluna_Situação_1 - 12

'Indústria Geral
ValoresIG_Situação_1 = Sheets("SITUACAO FINANCEIRA").Cells(9, Coluna_Situação_1).Value
ValoresIG_Situação_2 = Sheets("SITUACAO FINANCEIRA").Cells(9, Coluna_Situação_2).Value
ValoresIG_Situação_3 = Sheets("SITUACAO FINANCEIRA").Cells(9, Coluna_Situação_3).Value
Sheets("TABELAS").Cells(35, 8).Value = ValoresIG_Situação_3
Sheets("TABELAS").Cells(35, 9).Value = ValoresIG_Situação_2
Sheets("TABELAS").Cells(35, 10).Value = ValoresIG_Situação_1
'Indústria Extrativa
ValoresIE_Situação_1 = Sheets("SITUACAO FINANCEIRA").Cells(21, Coluna_Situação_1).Value
ValoresIE_Situação_2 = Sheets("SITUACAO FINANCEIRA").Cells(21, Coluna_Situação_2).Value
ValoresIE_Situação_3 = Sheets("SITUACAO FINANCEIRA").Cells(21, Coluna_Situação_3).Value
Sheets("TABELAS").Cells(38, 8).Value = ValoresIE_Situação_3
Sheets("TABELAS").Cells(38, 9).Value = ValoresIE_Situação_2
Sheets("TABELAS").Cells(38, 10).Value = ValoresIE_Situação_1
'Indústria Tranformação
ValoresIT_Situação_1 = Sheets("SITUACAO FINANCEIRA").Cells(26, Coluna_Situação_1).Value
ValoresIT_Situação_2 = Sheets("SITUACAO FINANCEIRA").Cells(26, Coluna_Situação_2).Value
ValoresIT_Situação_3 = Sheets("SITUACAO FINANCEIRA").Cells(26, Coluna_Situação_3).Value
Sheets("TABELAS").Cells(39, 8).Value = ValoresIT_Situação_3
Sheets("TABELAS").Cells(39, 9).Value = ValoresIT_Situação_2
Sheets("TABELAS").Cells(39, 10).Value = ValoresIT_Situação_1
'Pequena
ValoresP_Situação_1 = Sheets("SITUACAO FINANCEIRA").Cells(17, Coluna_Situação_1).Value
ValoresP_Situação_2 = Sheets("SITUACAO FINANCEIRA").Cells(17, Coluna_Situação_2).Value
ValoresP_Situação_3 = Sheets("SITUACAO FINANCEIRA").Cells(17, Coluna_Situação_3).Value
Sheets("TABELAS").Cells(42, 8).Value = ValoresP_Situação_3
Sheets("TABELAS").Cells(42, 9).Value = ValoresP_Situação_2
Sheets("TABELAS").Cells(42, 10).Value = ValoresP_Situação_1
'Média
ValoresM_Situação_1 = Sheets("SITUACAO FINANCEIRA").Cells(18, Coluna_Situação_1).Value
ValoresM_Situação_2 = Sheets("SITUACAO FINANCEIRA").Cells(18, Coluna_Situação_2).Value
ValoresM_Situação_3 = Sheets("SITUACAO FINANCEIRA").Cells(18, Coluna_Situação_3).Value
Sheets("TABELAS").Cells(43, 8).Value = ValoresM_Situação_3
Sheets("TABELAS").Cells(43, 9).Value = ValoresM_Situação_2
Sheets("TABELAS").Cells(43, 10).Value = ValoresM_Situação_1
'Grande
ValoresG_Situação_1 = Sheets("SITUACAO FINANCEIRA").Cells(19, Coluna_Situação_1).Value
ValoresG_Situação_2 = Sheets("SITUACAO FINANCEIRA").Cells(19, Coluna_Situação_2).Value
ValoresG_Situação_3 = Sheets("SITUACAO FINANCEIRA").Cells(19, Coluna_Situação_3).Value
Sheets("TABELAS").Cells(44, 8).Value = ValoresG_Situação_3
Sheets("TABELAS").Cells(44, 9).Value = ValoresG_Situação_2
Sheets("TABELAS").Cells(44, 10).Value = ValoresG_Situação_1

'Atribui os valores da coluna Acesso ao crédito
Coluna_Crédito_1 = Sheets("SITUACAO FINANCEIRA CREDITO").Range("B8").End(xlToRight).Column
Coluna_Crédito_2 = Coluna_Crédito_1 - 1
Coluna_Crédito_3 = Coluna_Crédito_1 - 12

'Indústria Geral
ValoresIG_Crédito_1 = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(9, Coluna_Crédito_1).Value
ValoresIG_Crédito_2 = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(9, Coluna_Crédito_2).Value
ValoresIG_Crédito_3 = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(9, Coluna_Crédito_3).Value
Sheets("TABELAS").Cells(35, 11).Value = ValoresIG_Crédito_3
Sheets("TABELAS").Cells(35, 12).Value = ValoresIG_Crédito_2
Sheets("TABELAS").Cells(35, 13).Value = ValoresIG_Crédito_1
'Indústria Extrativa
ValoresIE_Crédito_1 = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(21, Coluna_Crédito_1).Value
ValoresIE_Crédito_2 = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(21, Coluna_Crédito_2).Value
ValoresIE_Crédito_3 = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(21, Coluna_Crédito_3).Value
Sheets("TABELAS").Cells(38, 11).Value = ValoresIE_Crédito_3
Sheets("TABELAS").Cells(38, 12).Value = ValoresIE_Crédito_2
Sheets("TABELAS").Cells(38, 13).Value = ValoresIE_Crédito_1
'Indústria Transformação
ValoresIT_Crédito_1 = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(26, Coluna_Crédito_1).Value
ValoresIT_Crédito_2 = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(26, Coluna_Crédito_2).Value
ValoresIT_Crédito_3 = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(26, Coluna_Crédito_3).Value
Sheets("TABELAS").Cells(39, 11).Value = ValoresIT_Crédito_3
Sheets("TABELAS").Cells(39, 12).Value = ValoresIT_Crédito_2
Sheets("TABELAS").Cells(39, 13).Value = ValoresIT_Crédito_1
'Pequena
ValoresP_Crédito_1 = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(17, Coluna_Crédito_1).Value
ValoresP_Crédito_2 = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(17, Coluna_Crédito_2).Value
ValoresP_Crédito_3 = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(17, Coluna_Crédito_3).Value
Sheets("TABELAS").Cells(42, 11).Value = ValoresP_Crédito_3
Sheets("TABELAS").Cells(42, 12).Value = ValoresP_Crédito_2
Sheets("TABELAS").Cells(42, 13).Value = ValoresP_Crédito_1
'Média
ValoresM_Crédito_1 = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(18, Coluna_Crédito_1).Value
ValoresM_Crédito_2 = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(18, Coluna_Crédito_2).Value
ValoresM_Crédito_3 = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(18, Coluna_Crédito_3).Value
Sheets("TABELAS").Cells(43, 11).Value = ValoresM_Crédito_3
Sheets("TABELAS").Cells(43, 12).Value = ValoresM_Crédito_2
Sheets("TABELAS").Cells(43, 13).Value = ValoresM_Crédito_1
'Grande
ValoresG_Crédito_1 = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(19, Coluna_Crédito_1).Value
ValoresG_Crédito_2 = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(19, Coluna_Crédito_2).Value
ValoresG_Crédito_3 = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(19, Coluna_Crédito_3).Value
Sheets("TABELAS").Cells(44, 11).Value = ValoresG_Crédito_3
Sheets("TABELAS").Cells(44, 12).Value = ValoresG_Crédito_2
Sheets("TABELAS").Cells(44, 13).Value = ValoresG_Crédito_1

'*******************************************************Princiapais Problemas******************************************************


Sheets("PRINCIPAIS_PROBLEMAS").Select
Range("C109").Value = "Geral"
Range("C109:E109").Merge
Range("F109").Value = "Pequenas"
Range("F109:H109").Merge
Range("F109").Value = "Pequenas"
Range("F109:H109").Merge
Range("I109").Value = "Média"
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
Range("E111").Value = "Posição"

Range("C111:E111").Copy
Range("F111:H111").PasteSpecial
Range("I111:K111").PasteSpecial
Range("L111:M111").PasteSpecial


Coluna_Ultimo_Tri = Sheets("PRINCIPAIS_PROBLEMAS").Range("C10").End(xlToRight).Column
Coluna_Tri_Anterior = Coluna_Ultimo_Tri - 1
linha = 112

'Rank geral
Do Until linha = 128
posiçãoG = Application.WorksheetFunction.Rank_Eq(Cells(linha, 4), Range("D112:D127").Cells, 0)
Cells(linha, 5).Value = posiçãoG
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
posiçãoP = Application.WorksheetFunction.Rank_Eq(Cells(linha, 7), Range("G112:G127").Cells, 0)
Cells(linha, 8).Value = posiçãoP
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
posiçãoM = Application.WorksheetFunction.Rank_Eq(Cells(linha, 10), Range("J112:J127").Cells, 0)
Cells(linha, 11).Value = posiçãoM
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
posiçãoGr = Application.WorksheetFunction.Rank_Eq(Cells(linha, 13), Range("M112:M127").Cells, 0)
Cells(linha, 14).Value = posiçãoGr
linha = linha + 1
Loop

Sheets("PRINCIPAIS_PROBLEMAS").Select
Range("B109:N129").Copy
Sheets("TABELAS").Select
Range("V2").PasteSpecial
Range("V1").Value = "Principais Problemas"

End Sub


Sub Análise_Vermelho()

Dim Linha_Dados As Integer 'Define a linha que contém o dado a ser usado
Dim Coluna_Dados1 As Integer ' Define a coluna com o dado mais recente
Dim Coluna_Dados2 As Integer ' Define a coluna com o dado do mês anterior
Dim Coluna_Dados3 As Integer ' Defie a coluna com o dado do mesmo mês do ano anterior
Dim Linha_Análise As Integer ' Define a linha que será feita a análise
Dim Coluna_Análise As Integer 'Define a coluna que será feita a análise

Sheets("PRODUÇÃO").Select


'Copia os títulos das categorias e cola onde será formada a tabela de análise
Range(Cells(9, 1), Cells(54, 9)).Copy (Cells(59, 1))
'Limpa os números que foram colados mas mantém a formatação
Range(Cells(59, 2), Cells(105, 9)).ClearContents



'Nomeia as colunas de acordo com o dado que será calculado nelas
Cells(58, 2).Value = "Diferença para o mês anterior"
Cells(58, 3).Value = "Diferença para ao mesmo mês do ano anterior"
Cells(58, 4).Value = "Diferença para a média histórica"
Cells(58, 5).Value = "Posição Decrescente (Maior valor 1º, menor valor último)"
Cells(58, 6).Value = "Posição Crescente  (Menor valor 1º, maior valor último)"
Cells(58, 7).Value = "Posição Crescente - Mesmo mês  (Menor valor 1º, maior valor último)"
Cells(58, 8).Value = "Posição Decrescente -Mesmo mês  (Maior valor 1º, menor valor último)"
Cells(58, 9).Value = "cruzou a linha de 50?"

'Atribui valores as variaveis definidas acima
Linha_Dados = 9 'Define o número da primeira linha de dados
Coluna_Dados1 = Sheets("PRODUÇÃO").Range("A9").End(xlToRight).Column 'Define o número da última coluna
Coluna_Dados2 = Coluna_Dados1 - 1 'Define o número da coluna do mês anterior
Coluna_Dados3 = Coluna_Dados1 - 12
Linha_Análise = 59 'Define a primeira linhas de análises
Coluna_Análise = 2 'Define a coluna de análises

'Inserindo valores nas celulas vazias para fugir de bugs
Range(Cells(10, Coluna_Dados3), Cells(10, Coluna_Dados1)).Value = "0"
Range(Cells(16, Coluna_Dados3), Cells(16, Coluna_Dados1)).Value = "0"
Range(Cells(20, Coluna_Dados3), Cells(20, Coluna_Dados1)).Value = "0"
Range(Cells(22, Coluna_Dados3), Cells(23, Coluna_Dados1)).Value = "0"
Range(Cells(25, Coluna_Dados3), Cells(25, Coluna_Dados1)).Value = "0"
Range(Cells(29, Coluna_Dados3), Cells(29, Coluna_Dados1)).Value = "0"
Range(Cells(37, Coluna_Dados3), Cells(37, Coluna_Dados1)).Value = "0"


'Calculo da difernça em pontos do valor mais recente em relação ao valor do mês anterior
Do Until Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Calculo da diferença em si: o valor da celula de analise é igual ao valor mais recente menos o valor do mês anterior
   Sheets("PRODUÇÃO").Cells(Linha_Análise, Coluna_Análise).Value = Sheets("PRODUÇÃO").Cells(Linha_Dados, Coluna_Dados1).Value - Sheets("PRODUÇÃO").Cells(Linha_Dados, Coluna_Dados2).Value
    'Vai para a próxima linha de dados e de análise
   Linha_Dados = Linha_Dados + 1
   Linha_Análise = Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Linha_Dados = 9
Coluna_Dados1 = Sheets("PRODUÇÃO").Range("A9").End(xlToRight).Column
Coluna_Dados3 = Coluna_Dados1 - 12
Linha_Análise = 59
Coluna_Análise = 3

'Cálculo da diferença em pontos do valor mais recente em relação ao valor do mesmo mês do ano anterior
Do Until Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Calculo da diferença em si: o valor da celula de analise é igual ao valor mais recente menos o valor do mesmo mês do ano anterior
    Sheets("PRODUÇÃO").Cells(Linha_Análise, Coluna_Análise).Value = Sheets("PRODUÇÃO").Cells(Linha_Dados, Coluna_Dados1).Value - Sheets("PRODUÇÃO").Cells(Linha_Dados, Coluna_Dados3).Value
    'Vai para a próxima linha de dados e de análise
    Linha_Dados = Linha_Dados + 1
    Linha_Análise = Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Linha_Dados = 9
Coluna_Dados1 = Sheets("PRODUÇÃO").Range("A9").End(xlToRight).Column 'Define o número da última coluna
Linha_Análise = 59
Coluna_Análise = 4

'Cálculo da diferença em pontos do valor mais recente em relação ao valor da média histórica
Do Until Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a variável media como a média do intervalo entre a coluna com o dado mais recente e o primeiro
    media = Application.Average(Range(Cells(Linha_Dados, Coluna_Dados1), Cells(Linha_Dados, 2)))
    'Calculo em si: o valor da celula de analise é igual ao valor mais recente menos o valor da média
    Sheets("PRODUÇÃO").Cells(Linha_Análise, Coluna_Análise).Value = Sheets("PRODUÇÃO").Cells(Linha_Dados, Coluna_Dados1).Value - media
    'Vai para a próxima linha de dados e de análise
    Linha_Dados = Linha_Dados + 1
    Linha_Análise = Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Linha_Dados = 9
Coluna_Dados1 = Sheets("PRODUÇÃO").Range("A9").End(xlToRight).Column
Linha_Análise = 59
Coluna_Análise = 5

'Ordenação decrescente da série histórica completa
Do Until Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posição = WorksheetFunction.Rank_Eq(Cells(Linha_Dados, Coluna_Dados1), Range(Cells(Linha_Dados, Coluna_Dados1), Cells(Linha_Dados, 2)), 0)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("PRODUÇÃO").Cells(Linha_Análise, Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    Linha_Dados = Linha_Dados + 1
    Linha_Análise = Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Linha_Dados = 9
Coluna_Dados1 = Sheets("PRODUÇÃO").Range("A9").End(xlToRight).Column
Linha_Análise = 59
Coluna_Análise = 6

'Ordenação Crescente da série histórica completa
Do Until Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posição = WorksheetFunction.Rank_Eq(Cells(Linha_Dados, Coluna_Dados1), Range(Cells(Linha_Dados, Coluna_Dados1), Cells(Linha_Dados, 2)), 1)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("PRODUÇÃO").Cells(Linha_Análise, Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    Linha_Dados = Linha_Dados + 1
    Linha_Análise = Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior e define a variável Coluna_DadosP que representa a primeira coluna de dados
Coluna_Dados1 = Sheets("PRODUÇÃO").Range("A9").End(xlToRight).Column
Coluna_DadosP = 2

Do Until Coluna_DadosP = Coluna_Dados1 + 1 ' Faz até a variavel Coluna_DadosP ser igual a variavel Coluna_Dados1 mais uma unidade
    'Confere se o mês da coluna em questão é igual ao mês do dado mais recente
    If Month(Cells(8, Coluna_DadosP)) = Month(Cells(8, Coluna_Dados1)) Then
    'Caso seja igual, copia a coluna com os dados mais abaixo, a partir da linha 110
        Range(Cells(9, Coluna_DadosP), (Cells(54, Coluna_DadosP))).Copy (Cells(110, Coluna_DadosP))
    End If
    'Vai para a próxima coluna
    Coluna_DadosP = Coluna_DadosP + 1
'Repete a conferencia com a próxima coluna
Loop


'Atribui os valores originais das variaveis após o loop anterior
Linha_Dados = 110
Coluna_Dados1 = Sheets("PRODUÇÃO").Range("A9").End(xlToRight).Column 'Define o número da última coluna
Linha_Análise = 59
Coluna_Análise = 7

'Ordenação decrescente da série histórica dos meses do dado mais recente
Do Until Linha_Dados = 156 'Faz o calculo até a variável Linha_Dados ser 156
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo mês do mais recente
    posição = Application.WorksheetFunction.Rank_Eq(Cells(Linha_Dados, Coluna_Dados1), Range(Cells(Linha_Dados, Coluna_Dados1), Cells(Linha_Dados, 2)))
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("PRODUÇÃO").Cells(Linha_Análise, Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    Linha_Dados = Linha_Dados + 1
    Linha_Análise = Linha_Análise + 1
'Repete a ordenação com a próxima linha
Loop

'Atribui os valores originais das variaveis após o loop anterior
Linha_Dados = 110
Coluna_Dados1 = Sheets("PRODUÇÃO").Range("A9").End(xlToRight).Column 'Define o número da última coluna
Linha_Análise = 59
Coluna_Análise = 8
'Ordenação crescente da série histórica dos meses do dado mais recente
Do Until Linha_Dados = 156 'Faz o calculo até a variável Linha_Dados ser 156
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo mês do mais recente
    posição = Application.WorksheetFunction.Rank_Eq(Cells(Linha_Dados, Coluna_Dados1), Range(Cells(Linha_Dados, Coluna_Dados1), Cells(Linha_Dados, 2)), 1)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("PRODUÇÃO").Cells(Linha_Análise, Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    Linha_Dados = Linha_Dados + 1
    Linha_Análise = Linha_Análise + 1
'Repete a ordenação com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Linha_Dados = 9
Coluna_Dados1 = Sheets("PRODUÇÃO").Range("A9").End(xlToRight).Column
Coluna_Dados2 = Coluna_Dados1 - 1
Linha_Análise = 59
Coluna_Análise = 9

'Avaliação se cruzou ou não a linha de 50 e o sentido
Do Until Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'se o dado mais recente for menor que 50 e o dado do mês anterior for maior ou igual a 50 então...
    If Cells(Linha_Dados, Coluna_Dados1) < 50 And Cells(Linha_Dados, Coluna_Dados2) >= 50 Then
    'a célula de análise recebe cruzou para baixo
    Cells(Linha_Análise, Coluna_Análise).Value = "Cruzou para baixo"
    'Caso não seja..
    Else
        'se o dado mais recente for maior ou igual a 50 e o dado do m~es anterior for menor ou igual a 50 então...
        If Cells(Linha_Dados, Coluna_Dados1) >= 50 And Cells(Linha_Dados, Coluna_Dados2) <= 50 Then
        'a célula de análise recebe cruzou para cima
        Cells(Linha_Análise, Coluna_Análise).Value = "Cruzou para cima"
        'Caso não seja..
        Else
        'a célula de análise recebe não cruzou
        Cells(Linha_Análise, Coluna_Análise).Value = "Não Cruzou"
        End If
    End If
    'Vai para a próxima linha de dados e de análise
    Linha_Dados = Linha_Dados + 1
    Linha_Análise = Linha_Análise + 1
'repete o processo com a nova linha
Loop

'Apaga as linhas com erros/dados faltantes/títulos e subtítulos
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



'**********************************************              Análise_Empregados                ***********************************************************************


Dim Empregados_Linha_Dados As Integer 'Define a linha que contém o dado a ser usado
Dim Empregados_Coluna_Dados1 As Integer ' Define a coluna com o dado mais recente
Dim Empregados_Coluna_Dados2 As Integer ' Define a coluna com o dado do mês anterior
Dim Empregados_Coluna_Dados3 As Integer ' Defie a coluna com o dado do mesmo mês do ano anterior
Dim Empregados_Linha_Análise As Integer ' Define a linha que será feita a análise
Dim Empregados_Coluna_Análise As Integer 'Define a coluna que será feita a análise


Sheets("EMPREGADOS").Select


'Copia os títulos das categorias e cola onde será formada a tabela de análise
Sheets("EMPREGADOS").Range(Cells(9, 1), Cells(54, 9)).Copy (Sheets("EMPREGADOS").Cells(59, 1))
'Limpa os números que foram colados mas mantém a formatação
Sheets("EMPREGADOS").Range(Cells(59, 2), Cells(105, 9)).ClearContents



'Nomeia as colunas de acordo com o dado que será calculado nelas
Sheets("EMPREGADOS").Cells(58, 2).Value = "Diferença para o mês anterior"
Sheets("EMPREGADOS").Cells(58, 3).Value = "Diferença para ao mesmo mês do ano anterior"
Sheets("EMPREGADOS").Cells(58, 4).Value = "Diferença para a média histórica"
Sheets("EMPREGADOS").Cells(58, 5).Value = "Posição Decrescente (Maior valor 1º, menor valor último)"
Sheets("EMPREGADOS").Cells(58, 6).Value = "Posição Crescente  (Menor valor 1º, maior valor último)"
Sheets("EMPREGADOS").Cells(58, 7).Value = "Posição Crescente - Mesmo mês  (Menor valor 1º, maior valor último)"
Sheets("EMPREGADOS").Cells(58, 8).Value = "Posição Decrescente -Mesmo mês  (Maior valor 1º, menor valor último)"
Sheets("EMPREGADOS").Cells(58, 9).Value = "cruzou a linha de 50?"

'Atribui valores as variaveis definidas acima
Empregados_Linha_Dados = 9 'Define o número da primeira linha de dados
Empregados_Coluna_Dados1 = Sheets("EMPREGADOS").Range("A9").End(xlToRight).Column 'Define o número da última coluna
Empregados_Coluna_Dados2 = Empregados_Coluna_Dados1 - 1 'Define o número da coluna do mês anterior
Empregados_Coluna_Dados3 = Empregados_Coluna_Dados1 - 12
Empregados_Linha_Análise = 59 'Define a primeira linhas de análises
Empregados_Coluna_Análise = 2 'Define a coluna de análises

'Inserindo valores nas celulas vazias para fugir de bugs
Sheets("EMPREGADOS").Range(Cells(10, Empregados_Coluna_Dados3), Cells(10, Empregados_Coluna_Dados1)).Value = "0"
Sheets("EMPREGADOS").Range(Cells(16, Empregados_Coluna_Dados3), Cells(16, Empregados_Coluna_Dados1)).Value = "0"
Sheets("EMPREGADOS").Range(Cells(20, Empregados_Coluna_Dados3), Cells(20, Empregados_Coluna_Dados1)).Value = "0"
Sheets("EMPREGADOS").Range(Cells(22, Empregados_Coluna_Dados3), Cells(23, Empregados_Coluna_Dados1)).Value = "0"
Sheets("EMPREGADOS").Range(Cells(25, Empregados_Coluna_Dados3), Cells(25, Empregados_Coluna_Dados1)).Value = "0"
Sheets("EMPREGADOS").Range(Cells(29, Empregados_Coluna_Dados3), Cells(29, Empregados_Coluna_Dados1)).Value = "0"
Sheets("EMPREGADOS").Range(Cells(37, Empregados_Coluna_Dados3), Cells(37, Empregados_Coluna_Dados1)).Value = "0"


'Calculo da difernça em pontos do valor mais recente em relação ao valor do mês anterior
Do Until Empregados_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Calculo da diferença em si: o valor da celula de analise é igual ao valor mais recente menos o valor do mês anterior
   Sheets("EMPREGADOS").Cells(Empregados_Linha_Análise, Empregados_Coluna_Análise).Value = Sheets("EMPREGADOS").Cells(Empregados_Linha_Dados, Empregados_Coluna_Dados1).Value - Sheets("EMPREGADOS").Cells(Empregados_Linha_Dados, Empregados_Coluna_Dados2).Value
    'Vai para a próxima linha de dados e de análise
   Empregados_Linha_Dados = Empregados_Linha_Dados + 1
   Empregados_Linha_Análise = Empregados_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Empregados_Linha_Dados = 9
Empregados_Coluna_Dados1 = Sheets("EMPREGADOS").Range("A9").End(xlToRight).Column
Empregados_Coluna_Dados3 = Empregados_Coluna_Dados1 - 12
Empregados_Linha_Análise = 59
Empregados_Coluna_Análise = 3

'Cálculo da diferença em pontos do valor mais recente em relação ao valor do mesmo mês do ano anterior
Do Until Empregados_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Calculo da diferença em si: o valor da celula de analise é igual ao valor mais recente menos o valor do mesmo mês do ano anterior
    Sheets("EMPREGADOS").Cells(Empregados_Linha_Análise, Empregados_Coluna_Análise).Value = Sheets("EMPREGADOS").Cells(Empregados_Linha_Dados, Empregados_Coluna_Dados1).Value - Sheets("EMPREGADOS").Cells(Empregados_Linha_Dados, Empregados_Coluna_Dados3).Value
    'Vai para a próxima linha de dados e de análise
    Empregados_Linha_Dados = Empregados_Linha_Dados + 1
    Empregados_Linha_Análise = Empregados_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Empregados_Linha_Dados = 9
Empregados_Coluna_Dados1 = Sheets("EMPREGADOS").Range("A9").End(xlToRight).Column 'Define o número da última coluna
Empregados_Linha_Análise = 59
Empregados_Coluna_Análise = 4

'Cálculo da diferença em pontos do valor mais recente em relação ao valor da média histórica
Do Until Empregados_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a variável media como a média do intervalo entre a coluna com o dado mais recente e o primeiro
    media = Application.Average(Sheets("EMPREGADOS").Range(Cells(Empregados_Linha_Dados, Empregados_Coluna_Dados1), Cells(Empregados_Linha_Dados, 2)))
    'Calculo em si: o valor da celula de analise é igual ao valor mais recente menos o valor da média
    Sheets("EMPREGADOS").Cells(Empregados_Linha_Análise, Empregados_Coluna_Análise).Value = Sheets("EMPREGADOS").Cells(Empregados_Linha_Dados, Empregados_Coluna_Dados1).Value - media
    'Vai para a próxima linha de dados e de análise
    Empregados_Linha_Dados = Empregados_Linha_Dados + 1
    Empregados_Linha_Análise = Empregados_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Empregados_Linha_Dados = 9
Empregados_Coluna_Dados1 = Sheets("EMPREGADOS").Range("A9").End(xlToRight).Column
Empregados_Linha_Análise = 59
Empregados_Coluna_Análise = 5

'Ordenação decrescente da série histórica completa
Do Until Empregados_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posição = WorksheetFunction.Rank_Eq(Sheets("EMPREGADOS").Cells(Empregados_Linha_Dados, Empregados_Coluna_Dados1), Sheets("EMPREGADOS").Range(Cells(Empregados_Linha_Dados, Empregados_Coluna_Dados1), Cells(Empregados_Linha_Dados, 2)), 0)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("EMPREGADOS").Cells(Empregados_Linha_Análise, Empregados_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    Empregados_Linha_Dados = Empregados_Linha_Dados + 1
    Empregados_Linha_Análise = Empregados_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Empregados_Linha_Dados = 9
Empregados_Coluna_Dados1 = Sheets("EMPREGADOS").Range("A9").End(xlToRight).Column
Empregados_Linha_Análise = 59
Empregados_Coluna_Análise = 6

'Ordenação Crescente da série histórica completa
Do Until Empregados_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posição = WorksheetFunction.Rank_Eq(Sheets("EMPREGADOS").Cells(Empregados_Linha_Dados, Empregados_Coluna_Dados1), Sheets("EMPREGADOS").Range(Cells(Empregados_Linha_Dados, Empregados_Coluna_Dados1), Cells(Empregados_Linha_Dados, 2)), 1)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("EMPREGADOS").Cells(Empregados_Linha_Análise, Empregados_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    Empregados_Linha_Dados = Empregados_Linha_Dados + 1
    Empregados_Linha_Análise = Empregados_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior e define a variável Coluna_DadosP que representa a primeira coluna de dados
Empregados_Coluna_Dados1 = Sheets("EMPREGADOS").Range("A9").End(xlToRight).Column
Empregados_Coluna_DadosP = 2

Do Until Empregados_Coluna_DadosP = Empregados_Coluna_Dados1 + 1 ' Faz até a variavel Coluna_DadosP ser igual a variavel Coluna_Dados1 mais uma unidade
    'Confere se o mês da coluna em questão é igual ao mês do dado mais recente
    If Month(Sheets("EMPREGADOS").Cells(8, Empregados_Coluna_DadosP)) = Month(Sheets("EMPREGADOS").Cells(8, Empregados_Coluna_Dados1)) Then
    'Caso seja igual, copia a coluna com os dados mais abaixo, a partir da linha 110
        Sheets("EMPREGADOS").Range(Cells(9, Empregados_Coluna_DadosP), (Cells(54, Empregados_Coluna_DadosP))).Copy (Sheets("EMPREGADOS").Cells(110, Empregados_Coluna_DadosP))
    End If
    'Vai para a próxima coluna
    Empregados_Coluna_DadosP = Empregados_Coluna_DadosP + 1
'Repete a conferencia com a próxima coluna
Loop


'Atribui os valores originais das variaveis após o loop anterior
Empregados_Linha_Dados = 110
Empregados_Coluna_Dados1 = Sheets("EMPREGADOS").Range("A9").End(xlToRight).Column 'Define o número da última coluna
Empregados_Linha_Análise = 59
Empregados_Coluna_Análise = 7

'Ordenação decrescente da série histórica dos meses do dado mais recente
Do Until Empregados_Linha_Dados = 156 'Faz o calculo até a variável Linha_Dados ser 156
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo mês do mais recente
    posição = Application.WorksheetFunction.Rank_Eq(Sheets("EMPREGADOS").Cells(Empregados_Linha_Dados, Empregados_Coluna_Dados1), Sheets("EMPREGADOS").Range(Cells(Empregados_Linha_Dados, Empregados_Coluna_Dados1), Cells(Empregados_Linha_Dados, 2)))
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("EMPREGADOS").Cells(Empregados_Linha_Análise, Empregados_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    Empregados_Linha_Dados = Empregados_Linha_Dados + 1
    Empregados_Linha_Análise = Empregados_Linha_Análise + 1
'Repete a ordenação com a próxima linha
Loop

'Atribui os valores originais das variaveis após o loop anterior
Empregados_Linha_Dados = 110
Empregados_Coluna_Dados1 = Sheets("EMPREGADOS").Range("A9").End(xlToRight).Column 'Define o número da última coluna
Empregados_Linha_Análise = 59
Empregados_Coluna_Análise = 8
'Ordenação crescente da série histórica dos meses do dado mais recente
Do Until Empregados_Linha_Dados = 156 'Faz o calculo até a variável Linha_Dados ser 156
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo mês do mais recente
    posição = Application.WorksheetFunction.Rank_Eq(Sheets("EMPREGADOS").Cells(Empregados_Linha_Dados, Empregados_Coluna_Dados1), Sheets("EMPREGADOS").Range(Cells(Empregados_Linha_Dados, Empregados_Coluna_Dados1), Cells(Empregados_Linha_Dados, 2)), 1)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("EMPREGADOS").Cells(Empregados_Linha_Análise, Empregados_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    Empregados_Linha_Dados = Empregados_Linha_Dados + 1
    Empregados_Linha_Análise = Empregados_Linha_Análise + 1
'Repete a ordenação com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Empregados_Linha_Dados = 9
Empregados_Coluna_Dados1 = Sheets("EMPREGADOS").Range("A9").End(xlToRight).Column
Empregados_Coluna_Dados2 = Empregados_Coluna_Dados1 - 1
Empregados_Linha_Análise = 59
Empregados_Coluna_Análise = 9

'Avaliação se cruzou ou não a linha de 50 e o sentido
Do Until Empregados_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'se o dado mais recente for menor que 50 e o dado do mês anterior for maior ou igual a 50 então...
    If Cells(Empregados_Linha_Dados, Empregados_Coluna_Dados1) < 50 And Cells(Empregados_Linha_Dados, Empregados_Coluna_Dados2) >= 50 Then
    
    'a célula de análise recebe cruzou para baixo
    Cells(Empregados_Linha_Análise, Empregados_Coluna_Análise).Value = "Cruzou para baixo"
    'Caso não seja..
    Else
        'se o dado mais recente for maior ou igual a 50 e o dado do m~es anterior for menor ou igual a 50 então...
        If Cells(Empregados_Linha_Dados, Empregados_Coluna_Dados1) >= 50 And Cells(Empregados_Linha_Dados, Empregados_Coluna_Dados2) <= 50 Then
        'a célula de análise recebe cruzou para cima
        Cells(Empregados_Linha_Análise, Empregados_Coluna_Análise).Value = "Cruzou para cima"
        'Caso não seja..
        Else
        'a célula de análise recebe não cruzou
        Cells(Empregados_Linha_Análise, Empregados_Coluna_Análise).Value = "Não Cruzou"
        End If
    End If
    'Vai para a próxima linha de dados e de análise
    Empregados_Linha_Dados = Empregados_Linha_Dados + 1
    Empregados_Linha_Análise = Empregados_Linha_Análise + 1
'repete o processo com a nova linha
Loop

'Apaga as linhas com erros/dados faltantes/títulos e subtítulos
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




'**************************************************            Análise_UCI                        ********************************************************



Dim UCI_Linha_Dados As Integer 'Define a linha que contém o dado a ser usado
Dim UCI_Coluna_Dados1 As Integer ' Define a coluna com o dado mais recente
Dim UCI_Coluna_Dados2 As Integer ' Define a coluna com o dado do mês anterior
Dim UCI_Coluna_Dados3 As Integer ' Defie a coluna com o dado do mesmo mês do ano anterior
Dim UCI_Linha_Análise As Integer ' Define a linha que será feita a análise
Dim UCI_Coluna_Análise As Integer 'Define a coluna que será feita a análise

Sheets("UCI (%)").Select


'Copia os títulos das categorias e cola onde será formada a tabela de análise
Sheets("UCI (%)").Range(Cells(9, 1), Cells(54, 9)).Copy (Sheets("UCI (%)").Cells(59, 1))
'Limpa os números que foram colados mas mantém a formatação
Sheets("UCI (%)").Range(Cells(59, 2), Cells(105, 9)).ClearContents



'Nomeia as colunas de acordo com o dado que será calculado nelas
Sheets("UCI (%)").Cells(58, 2).Value = "Diferença para o mês anterior"
Sheets("UCI (%)").Cells(58, 3).Value = "Diferença para ao mesmo mês do ano anterior"
Sheets("UCI (%)").Cells(58, 4).Value = "Diferença para a média histórica"
Sheets("UCI (%)").Cells(58, 5).Value = "Posição Decrescente (Maior valor 1º, menor valor último)"
Sheets("UCI (%)").Cells(58, 6).Value = "Posição Crescente  (Menor valor 1º, maior valor último)"
Sheets("UCI (%)").Cells(58, 7).Value = "Posição Crescente - Mesmo mês  (Menor valor 1º, maior valor último)"
Sheets("UCI (%)").Cells(58, 8).Value = "Posição Decrescente -Mesmo mês  (Maior valor 1º, menor valor último)"
Sheets("UCI (%)").Cells(58, 9).Value = "cruzou a linha de 50?"

'Atribui valores as variaveis definidas acima
UCI_Linha_Dados = 9 'Define o número da primeira linha de dados
UCI_Coluna_Dados1 = Sheets("UCI (%)").Range("A9").End(xlToRight).Column 'Define o número da última coluna
UCI_Coluna_Dados2 = UCI_Coluna_Dados1 - 1 'Define o número da coluna do mês anterior
UCI_Coluna_Dados3 = UCI_Coluna_Dados1 - 12
UCI_Linha_Análise = 59 'Define a primeira linhas de análises
UCI_Coluna_Análise = 2 'Define a coluna de análises

'Inserindo valores nas celulas vazias para fugir de bugs
Sheets("UCI (%)").Range(Cells(10, UCI_Coluna_Dados3), Cells(10, UCI_Coluna_Dados1)).Value = "0"
Sheets("UCI (%)").Range(Cells(16, UCI_Coluna_Dados3), Cells(16, UCI_Coluna_Dados1)).Value = "0"
Sheets("UCI (%)").Range(Cells(20, UCI_Coluna_Dados3), Cells(20, UCI_Coluna_Dados1)).Value = "0"
Sheets("UCI (%)").Range(Cells(22, UCI_Coluna_Dados3), Cells(23, UCI_Coluna_Dados1)).Value = "0"
Sheets("UCI (%)").Range(Cells(25, UCI_Coluna_Dados3), Cells(25, UCI_Coluna_Dados1)).Value = "0"
Sheets("UCI (%)").Range(Cells(29, UCI_Coluna_Dados3), Cells(29, UCI_Coluna_Dados1)).Value = "0"
Sheets("UCI (%)").Range(Cells(37, UCI_Coluna_Dados3), Cells(37, UCI_Coluna_Dados1)).Value = "0"


'Calculo da difernça em pontos do valor mais recente em relação ao valor do mês anterior
Do Until UCI_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Calculo da diferença em si: o valor da celula de analise é igual ao valor mais recente menos o valor do mês anterior
   Sheets("UCI (%)").Cells(UCI_Linha_Análise, UCI_Coluna_Análise).Value = Sheets("UCI (%)").Cells(UCI_Linha_Dados, UCI_Coluna_Dados1).Value - Sheets("UCI (%)").Cells(UCI_Linha_Dados, UCI_Coluna_Dados2).Value
    'Vai para a próxima linha de dados e de análise
   UCI_Linha_Dados = UCI_Linha_Dados + 1
   UCI_Linha_Análise = UCI_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
UCI_Linha_Dados = 9
UCI_Coluna_Dados1 = Sheets("UCI (%)").Range("A9").End(xlToRight).Column
UCI_Coluna_Dados3 = UCI_Coluna_Dados1 - 12
UCI_Linha_Análise = 59
UCI_Coluna_Análise = 3

'Cálculo da diferença em pontos do valor mais recente em relação ao valor do mesmo mês do ano anterior
Do Until UCI_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Calculo da diferença em si: o valor da celula de analise é igual ao valor mais recente menos o valor do mesmo mês do ano anterior
    Sheets("UCI (%)").Cells(UCI_Linha_Análise, UCI_Coluna_Análise).Value = Sheets("UCI (%)").Cells(UCI_Linha_Dados, UCI_Coluna_Dados1).Value - Sheets("UCI (%)").Cells(UCI_Linha_Dados, UCI_Coluna_Dados3).Value
    'Vai para a próxima linha de dados e de análise
    UCI_Linha_Dados = UCI_Linha_Dados + 1
    UCI_Linha_Análise = UCI_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
UCI_Linha_Dados = 9
UCI_Coluna_Dados1 = Sheets("UCI (%)").Range("A9").End(xlToRight).Column 'Define o número da última coluna
UCI_Linha_Análise = 59
UCI_Coluna_Análise = 4

'Cálculo da diferença em pontos do valor mais recente em relação ao valor da média histórica
Do Until UCI_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a variável media como a média do intervalo entre a coluna com o dado mais recente e o primeiro
    media = Application.Average(Sheets("UCI (%)").Range(Cells(UCI_Linha_Dados, UCI_Coluna_Dados1), Cells(UCI_Linha_Dados, 2)))
    'Calculo em si: o valor da celula de analise é igual ao valor mais recente menos o valor da média
    Sheets("UCI (%)").Cells(UCI_Linha_Análise, UCI_Coluna_Análise).Value = Sheets("UCI (%)").Cells(UCI_Linha_Dados, UCI_Coluna_Dados1).Value - media
    'Vai para a próxima linha de dados e de análise
    UCI_Linha_Dados = UCI_Linha_Dados + 1
    UCI_Linha_Análise = UCI_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
UCI_Linha_Dados = 9
UCI_Coluna_Dados1 = Sheets("UCI (%)").Range("A9").End(xlToRight).Column
UCI_Linha_Análise = 59
UCI_Coluna_Análise = 5

'Ordenação decrescente da série histórica completa
Do Until UCI_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posição = WorksheetFunction.Rank_Eq(Sheets("UCI (%)").Cells(UCI_Linha_Dados, UCI_Coluna_Dados1), Sheets("UCI (%)").Range(Cells(UCI_Linha_Dados, UCI_Coluna_Dados1), Cells(UCI_Linha_Dados, 2)), 0)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("UCI (%)").Cells(UCI_Linha_Análise, UCI_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    UCI_Linha_Dados = UCI_Linha_Dados + 1
    UCI_Linha_Análise = UCI_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
UCI_Linha_Dados = 9
UCI_Coluna_Dados1 = Sheets("UCI (%)").Range("A9").End(xlToRight).Column
UCI_Linha_Análise = 59
UCI_Coluna_Análise = 6

'Ordenação Crescente da série histórica completa
Do Until UCI_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posição = WorksheetFunction.Rank_Eq(Sheets("UCI (%)").Cells(UCI_Linha_Dados, UCI_Coluna_Dados1), Sheets("UCI (%)").Range(Cells(UCI_Linha_Dados, UCI_Coluna_Dados1), Cells(UCI_Linha_Dados, 2)), 1)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("UCI (%)").Cells(UCI_Linha_Análise, UCI_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    UCI_Linha_Dados = UCI_Linha_Dados + 1
    UCI_Linha_Análise = UCI_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior e define a variável Coluna_DadosP que representa a primeira coluna de dados
UCI_Coluna_Dados1 = Sheets("UCI (%)").Range("A9").End(xlToRight).Column
UCI_Coluna_DadosP = 2

Do Until UCI_Coluna_DadosP = UCI_Coluna_Dados1 + 1 ' Faz até a variavel Coluna_DadosP ser igual a variavel Coluna_Dados1 mais uma unidade
    'Confere se o mês da coluna em questão é igual ao mês do dado mais recente
    If Month(Sheets("UCI (%)").Cells(8, UCI_Coluna_DadosP)) = Month(Sheets("UCI (%)").Cells(8, UCI_Coluna_Dados1)) Then
    'Caso seja igual, copia a coluna com os dados mais abaixo, a partir da linha 110
        Sheets("UCI (%)").Range(Cells(9, UCI_Coluna_DadosP), (Cells(54, UCI_Coluna_DadosP))).Copy (Sheets("UCI (%)").Cells(110, UCI_Coluna_DadosP))
    End If
    'Vai para a próxima coluna
    UCI_Coluna_DadosP = UCI_Coluna_DadosP + 1
'Repete a conferencia com a próxima coluna
Loop


'Atribui os valores originais das variaveis após o loop anterior
UCI_Linha_Dados = 110
UCI_Coluna_Dados1 = Sheets("UCI (%)").Range("A9").End(xlToRight).Column 'Define o número da última coluna
UCI_Linha_Análise = 59
UCI_Coluna_Análise = 7

'Ordenação decrescente da série histórica dos meses do dado mais recente
Do Until UCI_Linha_Dados = 156 'Faz o calculo até a variável Linha_Dados ser 156
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo mês do mais recente
    posição = Application.WorksheetFunction.Rank_Eq(Sheets("UCI (%)").Cells(UCI_Linha_Dados, UCI_Coluna_Dados1), Sheets("UCI (%)").Range(Cells(UCI_Linha_Dados, UCI_Coluna_Dados1), Cells(UCI_Linha_Dados, 2)))
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("UCI (%)").Cells(UCI_Linha_Análise, UCI_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    UCI_Linha_Dados = UCI_Linha_Dados + 1
    UCI_Linha_Análise = UCI_Linha_Análise + 1
'Repete a ordenação com a próxima linha
Loop

'Atribui os valores originais das variaveis após o loop anterior
UCI_Linha_Dados = 110
UCI_Coluna_Dados1 = Sheets("UCI (%)").Range("A9").End(xlToRight).Column 'Define o número da última coluna
UCI_Linha_Análise = 59
UCI_Coluna_Análise = 8
'Ordenação crescente da série histórica dos meses do dado mais recente
Do Until UCI_Linha_Dados = 156 'Faz o calculo até a variável Linha_Dados ser 156
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo mês do mais recente
    posição = Application.WorksheetFunction.Rank_Eq(Sheets("UCI (%)").Cells(UCI_Linha_Dados, UCI_Coluna_Dados1), Sheets("UCI (%)").Range(Cells(UCI_Linha_Dados, UCI_Coluna_Dados1), Cells(UCI_Linha_Dados, 2)), 1)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("UCI (%)").Cells(UCI_Linha_Análise, UCI_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    UCI_Linha_Dados = UCI_Linha_Dados + 1
    UCI_Linha_Análise = UCI_Linha_Análise + 1
'Repete a ordenação com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
UCI_Linha_Dados = 9
UCI_Coluna_Dados1 = Sheets("UCI (%)").Range("A9").End(xlToRight).Column
UCI_Coluna_Dados2 = UCI_Coluna_Dados1 - 1
UCI_Linha_Análise = 59
UCI_Coluna_Análise = 9

'Avaliação se cruzou ou não a linha de 50 e o sentido
Do Until UCI_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'se o dado mais recente for menor que 50 e o dado do mês anterior for maior ou igual a 50 então...
    If Cells(UCI_Linha_Dados, UCI_Coluna_Dados1) < 50 And Cells(UCI_Linha_Dados, UCI_Coluna_Dados2) >= 50 Then
    
    'a célula de análise recebe cruzou para baixo
    Cells(UCI_Linha_Análise, UCI_Coluna_Análise).Value = "Cruzou para baixo"
    'Caso não seja..
    Else
        'se o dado mais recente for maior ou igual a 50 e o dado do m~es anterior for menor ou igual a 50 então...
        If Cells(UCI_Linha_Dados, UCI_Coluna_Dados1) >= 50 And Cells(UCI_Linha_Dados, UCI_Coluna_Dados2) <= 50 Then
        'a célula de análise recebe cruzou para cima
        Cells(UCI_Linha_Análise, UCI_Coluna_Análise).Value = "Cruzou para cima"
        'Caso não seja..
        Else
        'a célula de análise recebe não cruzou
        Cells(UCI_Linha_Análise, UCI_Coluna_Análise).Value = "Não Cruzou"
        End If
    End If
    'Vai para a próxima linha de dados e de análise
    UCI_Linha_Dados = UCI_Linha_Dados + 1
    UCI_Linha_Análise = UCI_Linha_Análise + 1
'repete o processo com a nova linha
Loop

'Apaga as linhas com erros/dados faltantes/títulos e subtítulos
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





'*************************************       Análise_UCI_Efetiva_Usual                  **************************************************************



Dim UCI_Efetiva_Usual_Linha_Dados As Integer 'Define a linha que contém o dado a ser usado
Dim UCI_Efetiva_Usual_Coluna_Dados1 As Integer ' Define a coluna com o dado mais recente
Dim UCI_Efetiva_Usual_Coluna_Dados2 As Integer ' Define a coluna com o dado do mês anterior
Dim UCI_Efetiva_Usual_Coluna_Dados3 As Integer ' Defie a coluna com o dado do mesmo mês do ano anterior
Dim UCI_Efetiva_Usual_Linha_Análise As Integer ' Define a linha que será feita a análise
Dim UCI_Efetiva_Usual_Coluna_Análise As Integer 'Define a coluna que será feita a análise

Sheets("UCI (efetiva-usual)").Select


'Copia os títulos das categorias e cola onde será formada a tabela de análise
Sheets("UCI (efetiva-usual)").Range(Cells(9, 1), Cells(54, 9)).Copy (Sheets("UCI (efetiva-usual)").Cells(59, 1))
'Limpa os números que foram colados mas mantém a formatação
Sheets("UCI (efetiva-usual)").Range(Cells(59, 2), Cells(105, 9)).ClearContents



'Nomeia as colunas de acordo com o dado que será calculado nelas
Sheets("UCI (efetiva-usual)").Cells(58, 2).Value = "Diferença para o mês anterior"
Sheets("UCI (efetiva-usual)").Cells(58, 3).Value = "Diferença para ao mesmo mês do ano anterior"
Sheets("UCI (efetiva-usual)").Cells(58, 4).Value = "Diferença para a média histórica"
Sheets("UCI (efetiva-usual)").Cells(58, 5).Value = "Posição Decrescente (Maior valor 1º, menor valor último)"
Sheets("UCI (efetiva-usual)").Cells(58, 6).Value = "Posição Crescente  (Menor valor 1º, maior valor último)"
Sheets("UCI (efetiva-usual)").Cells(58, 7).Value = "Posição Crescente - Mesmo mês  (Menor valor 1º, maior valor último)"
Sheets("UCI (efetiva-usual)").Cells(58, 8).Value = "Posição Decrescente -Mesmo mês  (Maior valor 1º, menor valor último)"
Sheets("UCI (efetiva-usual)").Cells(58, 9).Value = "cruzou a linha de 50?"

'Atribui valores as variaveis definidas acima
UCI_Efetiva_Usual_Linha_Dados = 9 'Define o número da primeira linha de dados
UCI_Efetiva_Usual_Coluna_Dados1 = Sheets("UCI (efetiva-usual)").Range("A9").End(xlToRight).Column 'Define o número da última coluna
UCI_Efetiva_Usual_Coluna_Dados2 = UCI_Efetiva_Usual_Coluna_Dados1 - 1 'Define o número da coluna do mês anterior
UCI_Efetiva_Usual_Coluna_Dados3 = UCI_Efetiva_Usual_Coluna_Dados1 - 12
UCI_Efetiva_Usual_Linha_Análise = 59 'Define a primeira linhas de análises
UCI_Efetiva_Usual_Coluna_Análise = 2 'Define a coluna de análises

'Inserindo valores nas celulas vazias para fugir de bugs
Sheets("UCI (efetiva-usual)").Range(Cells(10, UCI_Efetiva_Usual_Coluna_Dados3), Cells(10, UCI_Efetiva_Usual_Coluna_Dados1)).Value = "0"
Sheets("UCI (efetiva-usual)").Range(Cells(16, UCI_Efetiva_Usual_Coluna_Dados3), Cells(16, UCI_Efetiva_Usual_Coluna_Dados1)).Value = "0"
Sheets("UCI (efetiva-usual)").Range(Cells(20, UCI_Efetiva_Usual_Coluna_Dados3), Cells(20, UCI_Efetiva_Usual_Coluna_Dados1)).Value = "0"
Sheets("UCI (efetiva-usual)").Range(Cells(22, UCI_Efetiva_Usual_Coluna_Dados3), Cells(23, UCI_Efetiva_Usual_Coluna_Dados1)).Value = "0"
Sheets("UCI (efetiva-usual)").Range(Cells(25, UCI_Efetiva_Usual_Coluna_Dados3), Cells(25, UCI_Efetiva_Usual_Coluna_Dados1)).Value = "0"
Sheets("UCI (efetiva-usual)").Range(Cells(29, UCI_Efetiva_Usual_Coluna_Dados3), Cells(29, UCI_Efetiva_Usual_Coluna_Dados1)).Value = "0"
Sheets("UCI (efetiva-usual)").Range(Cells(37, UCI_Efetiva_Usual_Coluna_Dados3), Cells(37, UCI_Efetiva_Usual_Coluna_Dados1)).Value = "0"


'Calculo da difernça em pontos do valor mais recente em relação ao valor do mês anterior
Do Until UCI_Efetiva_Usual_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Calculo da diferença em si: o valor da celula de analise é igual ao valor mais recente menos o valor do mês anterior
   Sheets("UCI (efetiva-usual)").Cells(UCI_Efetiva_Usual_Linha_Análise, UCI_Efetiva_Usual_Coluna_Análise).Value = Sheets("UCI (efetiva-usual)").Cells(UCI_Efetiva_Usual_Linha_Dados, UCI_Efetiva_Usual_Coluna_Dados1).Value - Sheets("UCI (efetiva-usual)").Cells(UCI_Efetiva_Usual_Linha_Dados, UCI_Efetiva_Usual_Coluna_Dados2).Value
    'Vai para a próxima linha de dados e de análise
   UCI_Efetiva_Usual_Linha_Dados = UCI_Efetiva_Usual_Linha_Dados + 1
   UCI_Efetiva_Usual_Linha_Análise = UCI_Efetiva_Usual_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
UCI_Efetiva_Usual_Linha_Dados = 9
UCI_Efetiva_Usual_Coluna_Dados1 = Sheets("UCI (efetiva-usual)").Range("A9").End(xlToRight).Column
UCI_Efetiva_Usual_Coluna_Dados3 = UCI_Efetiva_Usual_Coluna_Dados1 - 12
UCI_Efetiva_Usual_Linha_Análise = 59
UCI_Efetiva_Usual_Coluna_Análise = 3

'Cálculo da diferença em pontos do valor mais recente em relação ao valor do mesmo mês do ano anterior
Do Until UCI_Efetiva_Usual_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Calculo da diferença em si: o valor da celula de analise é igual ao valor mais recente menos o valor do mesmo mês do ano anterior
    Sheets("UCI (efetiva-usual)").Cells(UCI_Efetiva_Usual_Linha_Análise, UCI_Efetiva_Usual_Coluna_Análise).Value = Sheets("UCI (efetiva-usual)").Cells(UCI_Efetiva_Usual_Linha_Dados, UCI_Efetiva_Usual_Coluna_Dados1).Value - Sheets("UCI (efetiva-usual)").Cells(UCI_Efetiva_Usual_Linha_Dados, UCI_Efetiva_Usual_Coluna_Dados3).Value
    'Vai para a próxima linha de dados e de análise
    UCI_Efetiva_Usual_Linha_Dados = UCI_Efetiva_Usual_Linha_Dados + 1
    UCI_Efetiva_Usual_Linha_Análise = UCI_Efetiva_Usual_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
UCI_Efetiva_Usual_Linha_Dados = 9
UCI_Efetiva_Usual_Coluna_Dados1 = Sheets("UCI (efetiva-usual)").Range("A9").End(xlToRight).Column 'Define o número da última coluna
UCI_Efetiva_Usual_Linha_Análise = 59
UCI_Efetiva_Usual_Coluna_Análise = 4

'Cálculo da diferença em pontos do valor mais recente em relação ao valor da média histórica
Do Until UCI_Efetiva_Usual_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a variável media como a média do intervalo entre a coluna com o dado mais recente e o primeiro
    media = Application.Average(Sheets("UCI (efetiva-usual)").Range(Cells(UCI_Efetiva_Usual_Linha_Dados, UCI_Efetiva_Usual_Coluna_Dados1), Cells(UCI_Efetiva_Usual_Linha_Dados, 2)))
    'Calculo em si: o valor da celula de analise é igual ao valor mais recente menos o valor da média
    Sheets("UCI (efetiva-usual)").Cells(UCI_Efetiva_Usual_Linha_Análise, UCI_Efetiva_Usual_Coluna_Análise).Value = Sheets("UCI (efetiva-usual)").Cells(UCI_Efetiva_Usual_Linha_Dados, UCI_Efetiva_Usual_Coluna_Dados1).Value - media
    'Vai para a próxima linha de dados e de análise
    UCI_Efetiva_Usual_Linha_Dados = UCI_Efetiva_Usual_Linha_Dados + 1
    UCI_Efetiva_Usual_Linha_Análise = UCI_Efetiva_Usual_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
UCI_Efetiva_Usual_Linha_Dados = 9
UCI_Efetiva_Usual_Coluna_Dados1 = Sheets("UCI (efetiva-usual)").Range("A9").End(xlToRight).Column
UCI_Efetiva_Usual_Linha_Análise = 59
UCI_Efetiva_Usual_Coluna_Análise = 5

'Ordenação decrescente da série histórica completa
Do Until UCI_Efetiva_Usual_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posição = WorksheetFunction.Rank_Eq(Sheets("UCI (efetiva-usual)").Cells(UCI_Efetiva_Usual_Linha_Dados, UCI_Efetiva_Usual_Coluna_Dados1), Sheets("UCI (efetiva-usual)").Range(Cells(UCI_Efetiva_Usual_Linha_Dados, UCI_Efetiva_Usual_Coluna_Dados1), Cells(UCI_Efetiva_Usual_Linha_Dados, 2)), 0)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("UCI (efetiva-usual)").Cells(UCI_Efetiva_Usual_Linha_Análise, UCI_Efetiva_Usual_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    UCI_Efetiva_Usual_Linha_Dados = UCI_Efetiva_Usual_Linha_Dados + 1
    UCI_Efetiva_Usual_Linha_Análise = UCI_Efetiva_Usual_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
UCI_Efetiva_Usual_Linha_Dados = 9
UCI_Efetiva_Usual_Coluna_Dados1 = Sheets("UCI (efetiva-usual)").Range("A9").End(xlToRight).Column
UCI_Efetiva_Usual_Linha_Análise = 59
UCI_Efetiva_Usual_Coluna_Análise = 6

'Ordenação Crescente da série histórica completa
Do Until UCI_Efetiva_Usual_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posição = WorksheetFunction.Rank_Eq(Sheets("UCI (efetiva-usual)").Cells(UCI_Efetiva_Usual_Linha_Dados, UCI_Efetiva_Usual_Coluna_Dados1), Sheets("UCI (efetiva-usual)").Range(Cells(UCI_Efetiva_Usual_Linha_Dados, UCI_Efetiva_Usual_Coluna_Dados1), Cells(UCI_Efetiva_Usual_Linha_Dados, 2)), 1)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("UCI (efetiva-usual)").Cells(UCI_Efetiva_Usual_Linha_Análise, UCI_Efetiva_Usual_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    UCI_Efetiva_Usual_Linha_Dados = UCI_Efetiva_Usual_Linha_Dados + 1
    UCI_Efetiva_Usual_Linha_Análise = UCI_Efetiva_Usual_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior e define a variável Coluna_DadosP que representa a primeira coluna de dados
UCI_Efetiva_Usual_Coluna_Dados1 = Sheets("UCI (efetiva-usual)").Range("A9").End(xlToRight).Column
UCI_Efetiva_Usual_Coluna_DadosP = 2

Do Until UCI_Efetiva_Usual_Coluna_DadosP = UCI_Efetiva_Usual_Coluna_Dados1 + 1 ' Faz até a variavel Coluna_DadosP ser igual a variavel Coluna_Dados1 mais uma unidade
    'Confere se o mês da coluna em questão é igual ao mês do dado mais recente
    If Month(Sheets("UCI (efetiva-usual)").Cells(8, UCI_Efetiva_Usual_Coluna_DadosP)) = Month(Sheets("UCI (efetiva-usual)").Cells(8, UCI_Efetiva_Usual_Coluna_Dados1)) Then
    'Caso seja igual, copia a coluna com os dados mais abaixo, a partir da linha 110
        Sheets("UCI (efetiva-usual)").Range(Cells(9, UCI_Efetiva_Usual_Coluna_DadosP), (Cells(54, UCI_Efetiva_Usual_Coluna_DadosP))).Copy (Sheets("UCI (efetiva-usual)").Cells(110, UCI_Efetiva_Usual_Coluna_DadosP))
    End If
    'Vai para a próxima coluna
    UCI_Efetiva_Usual_Coluna_DadosP = UCI_Efetiva_Usual_Coluna_DadosP + 1
'Repete a conferencia com a próxima coluna
Loop


'Atribui os valores originais das variaveis após o loop anterior
UCI_Efetiva_Usual_Linha_Dados = 110
UCI_Efetiva_Usual_Coluna_Dados1 = Sheets("UCI (efetiva-usual)").Range("A9").End(xlToRight).Column 'Define o número da última coluna
UCI_Efetiva_Usual_Linha_Análise = 59
UCI_Efetiva_Usual_Coluna_Análise = 7

'Ordenação decrescente da série histórica dos meses do dado mais recente
Do Until UCI_Efetiva_Usual_Linha_Dados = 156 'Faz o calculo até a variável Linha_Dados ser 156
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo mês do mais recente
    posição = Application.WorksheetFunction.Rank_Eq(Sheets("UCI (efetiva-usual)").Cells(UCI_Efetiva_Usual_Linha_Dados, UCI_Efetiva_Usual_Coluna_Dados1), Sheets("UCI (efetiva-usual)").Range(Cells(UCI_Efetiva_Usual_Linha_Dados, UCI_Efetiva_Usual_Coluna_Dados1), Cells(UCI_Efetiva_Usual_Linha_Dados, 2)))
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("UCI (efetiva-usual)").Cells(UCI_Efetiva_Usual_Linha_Análise, UCI_Efetiva_Usual_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    UCI_Efetiva_Usual_Linha_Dados = UCI_Efetiva_Usual_Linha_Dados + 1
    UCI_Efetiva_Usual_Linha_Análise = UCI_Efetiva_Usual_Linha_Análise + 1
'Repete a ordenação com a próxima linha
Loop

'Atribui os valores originais das variaveis após o loop anterior
UCI_Efetiva_Usual_Linha_Dados = 110
UCI_Efetiva_Usual_Coluna_Dados1 = Sheets("UCI (efetiva-usual)").Range("A9").End(xlToRight).Column 'Define o número da última coluna
UCI_Efetiva_Usual_Linha_Análise = 59
UCI_Efetiva_Usual_Coluna_Análise = 8
'Ordenação crescente da série histórica dos meses do dado mais recente
Do Until UCI_Efetiva_Usual_Linha_Dados = 156 'Faz o calculo até a variável Linha_Dados ser 156
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo mês do mais recente
    posição = Application.WorksheetFunction.Rank_Eq(Sheets("UCI (efetiva-usual)").Cells(UCI_Efetiva_Usual_Linha_Dados, UCI_Efetiva_Usual_Coluna_Dados1), Sheets("UCI (efetiva-usual)").Range(Cells(UCI_Efetiva_Usual_Linha_Dados, UCI_Efetiva_Usual_Coluna_Dados1), Cells(UCI_Efetiva_Usual_Linha_Dados, 2)), 1)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("UCI (efetiva-usual)").Cells(UCI_Efetiva_Usual_Linha_Análise, UCI_Efetiva_Usual_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    UCI_Efetiva_Usual_Linha_Dados = UCI_Efetiva_Usual_Linha_Dados + 1
    UCI_Efetiva_Usual_Linha_Análise = UCI_Efetiva_Usual_Linha_Análise + 1
'Repete a ordenação com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
UCI_Efetiva_Usual_Linha_Dados = 9
UCI_Efetiva_Usual_Coluna_Dados1 = Sheets("UCI (efetiva-usual)").Range("A9").End(xlToRight).Column
UCI_Efetiva_Usual_Coluna_Dados2 = UCI_Efetiva_Usual_Coluna_Dados1 - 1
UCI_Efetiva_Usual_Linha_Análise = 59
UCI_Efetiva_Usual_Coluna_Análise = 9

'Avaliação se cruzou ou não a linha de 50 e o sentido
Do Until UCI_Efetiva_Usual_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'se o dado mais recente for menor que 50 e o dado do mês anterior for maior ou igual a 50 então...
    If Cells(UCI_Efetiva_Usual_Linha_Dados, UCI_Efetiva_Usual_Coluna_Dados1) < 50 And Cells(UCI_Efetiva_Usual_Linha_Dados, UCI_Efetiva_Usual_Coluna_Dados2) >= 50 Then
    
    'a célula de análise recebe cruzou para baixo
    Cells(UCI_Efetiva_Usual_Linha_Análise, UCI_Efetiva_Usual_Coluna_Análise).Value = "Cruzou para baixo"
    'Caso não seja..
    Else
        'se o dado mais recente for maior ou igual a 50 e o dado do m~es anterior for menor ou igual a 50 então...
        If Cells(UCI_Efetiva_Usual_Linha_Dados, UCI_Efetiva_Usual_Coluna_Dados1) >= 50 And Cells(UCI_Efetiva_Usual_Linha_Dados, UCI_Efetiva_Usual_Coluna_Dados2) <= 50 Then
        'a célula de análise recebe cruzou para cima
        Cells(UCI_Efetiva_Usual_Linha_Análise, UCI_Efetiva_Usual_Coluna_Análise).Value = "Cruzou para cima"
        'Caso não seja..
        Else
        'a célula de análise recebe não cruzou
        Cells(UCI_Efetiva_Usual_Linha_Análise, UCI_Efetiva_Usual_Coluna_Análise).Value = "Não Cruzou"
        End If
    End If
    'Vai para a próxima linha de dados e de análise
    UCI_Efetiva_Usual_Linha_Dados = UCI_Efetiva_Usual_Linha_Dados + 1
    UCI_Efetiva_Usual_Linha_Análise = UCI_Efetiva_Usual_Linha_Análise + 1
'repete o processo com a nova linha
Loop

'Apaga as linhas com erros/dados faltantes/títulos e subtítulos
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




'*************************************          ESTOQUES_evolução            *********************************************************************




Dim Estoques_Evolução_Linha_Dados As Integer 'Define a linha que contém o dado a ser usado
Dim Estoques_Evolução_Coluna_Dados1 As Integer ' Define a coluna com o dado mais recente
Dim Estoques_Evolução_Coluna_Dados2 As Integer ' Define a coluna com o dado do mês anterior
Dim Estoques_Evolução_Coluna_Dados3 As Integer ' Defie a coluna com o dado do mesmo mês do ano anterior
Dim Estoques_Evolução_Linha_Análise As Integer ' Define a linha que será feita a análise
Dim Estoques_Evolução_Coluna_Análise As Integer 'Define a coluna que será feita a análise

Sheets("ESTOQUES (evolução)").Select

'Copia os títulos das categorias e cola onde será formada a tabela de análise
Sheets("ESTOQUES (evolução)").Range(Cells(9, 1), Cells(54, 9)).Copy (Sheets("ESTOQUES (evolução)").Cells(59, 1))
'Limpa os números que foram colados mas mantém a formatação
Sheets("ESTOQUES (evolução)").Range(Cells(59, 2), Cells(105, 9)).ClearContents



'Nomeia as colunas de acordo com o dado que será calculado nelas
Sheets("ESTOQUES (evolução)").Cells(58, 2).Value = "Diferença para o mês anterior"
Sheets("ESTOQUES (evolução)").Cells(58, 3).Value = "Diferença para ao mesmo mês do ano anterior"
Sheets("ESTOQUES (evolução)").Cells(58, 4).Value = "Diferença para a média histórica"
Sheets("ESTOQUES (evolução)").Cells(58, 5).Value = "Posição Decrescente (Maior valor 1º, menor valor último)"
Sheets("ESTOQUES (evolução)").Cells(58, 6).Value = "Posição Crescente  (Menor valor 1º, maior valor último)"
Sheets("ESTOQUES (evolução)").Cells(58, 7).Value = "Posição Crescente - Mesmo mês  (Menor valor 1º, maior valor último)"
Sheets("ESTOQUES (evolução)").Cells(58, 8).Value = "Posição Decrescente -Mesmo mês  (Maior valor 1º, menor valor último)"
Sheets("ESTOQUES (evolução)").Cells(58, 9).Value = "cruzou a linha de 50?"

'Atribui valores as variaveis definidas acima
Estoques_Evolução_Linha_Dados = 9 'Define o número da primeira linha de dados
Estoques_Evolução_Coluna_Dados1 = Sheets("ESTOQUES (evolução)").Range("A9").End(xlToRight).Column 'Define o número da última coluna
Estoques_Evolução_Coluna_Dados2 = Estoques_Evolução_Coluna_Dados1 - 1 'Define o número da coluna do mês anterior
Estoques_Evolução_Coluna_Dados3 = Estoques_Evolução_Coluna_Dados1 - 12
Estoques_Evolução_Linha_Análise = 59 'Define a primeira linhas de análises
Estoques_Evolução_Coluna_Análise = 2 'Define a coluna de análises

'Inserindo valores nas celulas vazias para fugir de bugs
Sheets("ESTOQUES (evolução)").Range(Cells(10, Estoques_Evolução_Coluna_Dados3), Cells(10, Estoques_Evolução_Coluna_Dados1)).Value = "0"
Sheets("ESTOQUES (evolução)").Range(Cells(16, Estoques_Evolução_Coluna_Dados3), Cells(16, Estoques_Evolução_Coluna_Dados1)).Value = "0"
Sheets("ESTOQUES (evolução)").Range(Cells(20, Estoques_Evolução_Coluna_Dados3), Cells(20, Estoques_Evolução_Coluna_Dados1)).Value = "0"
Sheets("ESTOQUES (evolução)").Range(Cells(22, Estoques_Evolução_Coluna_Dados3), Cells(23, Estoques_Evolução_Coluna_Dados1)).Value = "0"
Sheets("ESTOQUES (evolução)").Range(Cells(25, Estoques_Evolução_Coluna_Dados3), Cells(25, Estoques_Evolução_Coluna_Dados1)).Value = "0"
Sheets("ESTOQUES (evolução)").Range(Cells(29, Estoques_Evolução_Coluna_Dados3), Cells(29, Estoques_Evolução_Coluna_Dados1)).Value = "0"
Sheets("ESTOQUES (evolução)").Range(Cells(37, Estoques_Evolução_Coluna_Dados3), Cells(37, Estoques_Evolução_Coluna_Dados1)).Value = "0"


'Calculo da difernça em pontos do valor mais recente em relação ao valor do mês anterior
Do Until Estoques_Evolução_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Calculo da diferença em si: o valor da celula de analise é igual ao valor mais recente menos o valor do mês anterior
   Sheets("ESTOQUES (evolução)").Cells(Estoques_Evolução_Linha_Análise, Estoques_Evolução_Coluna_Análise).Value = Sheets("ESTOQUES (evolução)").Cells(Estoques_Evolução_Linha_Dados, Estoques_Evolução_Coluna_Dados1).Value - Sheets("ESTOQUES (evolução)").Cells(Estoques_Evolução_Linha_Dados, Estoques_Evolução_Coluna_Dados2).Value
    'Vai para a próxima linha de dados e de análise
   Estoques_Evolução_Linha_Dados = Estoques_Evolução_Linha_Dados + 1
   Estoques_Evolução_Linha_Análise = Estoques_Evolução_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Estoques_Evolução_Linha_Dados = 9
Estoques_Evolução_Coluna_Dados1 = Sheets("ESTOQUES (evolução)").Range("A9").End(xlToRight).Column
Estoques_Evolução_Coluna_Dados3 = Estoques_Evolução_Coluna_Dados1 - 12
Estoques_Evolução_Linha_Análise = 59
Estoques_Evolução_Coluna_Análise = 3

'Cálculo da diferença em pontos do valor mais recente em relação ao valor do mesmo mês do ano anterior
Do Until Estoques_Evolução_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Calculo da diferença em si: o valor da celula de analise é igual ao valor mais recente menos o valor do mesmo mês do ano anterior
    Sheets("ESTOQUES (evolução)").Cells(Estoques_Evolução_Linha_Análise, Estoques_Evolução_Coluna_Análise).Value = Sheets("ESTOQUES (evolução)").Cells(Estoques_Evolução_Linha_Dados, Estoques_Evolução_Coluna_Dados1).Value - Sheets("ESTOQUES (evolução)").Cells(Estoques_Evolução_Linha_Dados, Estoques_Evolução_Coluna_Dados3).Value
    'Vai para a próxima linha de dados e de análise
    Estoques_Evolução_Linha_Dados = Estoques_Evolução_Linha_Dados + 1
    Estoques_Evolução_Linha_Análise = Estoques_Evolução_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Estoques_Evolução_Linha_Dados = 9
Estoques_Evolução_Coluna_Dados1 = Sheets("ESTOQUES (evolução)").Range("A9").End(xlToRight).Column 'Define o número da última coluna
Estoques_Evolução_Linha_Análise = 59
Estoques_Evolução_Coluna_Análise = 4

'Cálculo da diferença em pontos do valor mais recente em relação ao valor da média histórica
Do Until Estoques_Evolução_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a variável media como a média do intervalo entre a coluna com o dado mais recente e o primeiro
    media = Application.Average(Sheets("ESTOQUES (evolução)").Range(Cells(Estoques_Evolução_Linha_Dados, Estoques_Evolução_Coluna_Dados1), Cells(Estoques_Evolução_Linha_Dados, 2)))
    'Calculo em si: o valor da celula de analise é igual ao valor mais recente menos o valor da média
    Sheets("ESTOQUES (evolução)").Cells(Estoques_Evolução_Linha_Análise, Estoques_Evolução_Coluna_Análise).Value = Sheets("ESTOQUES (evolução)").Cells(Estoques_Evolução_Linha_Dados, Estoques_Evolução_Coluna_Dados1).Value - media
    'Vai para a próxima linha de dados e de análise
    Estoques_Evolução_Linha_Dados = Estoques_Evolução_Linha_Dados + 1
    Estoques_Evolução_Linha_Análise = Estoques_Evolução_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Estoques_Evolução_Linha_Dados = 9
Estoques_Evolução_Coluna_Dados1 = Sheets("ESTOQUES (evolução)").Range("A9").End(xlToRight).Column
Estoques_Evolução_Linha_Análise = 59
Estoques_Evolução_Coluna_Análise = 5

'Ordenação decrescente da série histórica completa
Do Until Estoques_Evolução_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posição = WorksheetFunction.Rank_Eq(Sheets("ESTOQUES (evolução)").Cells(Estoques_Evolução_Linha_Dados, Estoques_Evolução_Coluna_Dados1), Sheets("ESTOQUES (evolução)").Range(Cells(Estoques_Evolução_Linha_Dados, Estoques_Evolução_Coluna_Dados1), Cells(Estoques_Evolução_Linha_Dados, 2)), 0)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("ESTOQUES (evolução)").Cells(Estoques_Evolução_Linha_Análise, Estoques_Evolução_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    Estoques_Evolução_Linha_Dados = Estoques_Evolução_Linha_Dados + 1
    Estoques_Evolução_Linha_Análise = Estoques_Evolução_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Estoques_Evolução_Linha_Dados = 9
Estoques_Evolução_Coluna_Dados1 = Sheets("ESTOQUES (evolução)").Range("A9").End(xlToRight).Column
Estoques_Evolução_Linha_Análise = 59
Estoques_Evolução_Coluna_Análise = 6

'Ordenação Crescente da série histórica completa
Do Until Estoques_Evolução_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posição = WorksheetFunction.Rank_Eq(Sheets("ESTOQUES (evolução)").Cells(Estoques_Evolução_Linha_Dados, Estoques_Evolução_Coluna_Dados1), Sheets("ESTOQUES (evolução)").Range(Cells(Estoques_Evolução_Linha_Dados, Estoques_Evolução_Coluna_Dados1), Cells(Estoques_Evolução_Linha_Dados, 2)), 1)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("ESTOQUES (evolução)").Cells(Estoques_Evolução_Linha_Análise, Estoques_Evolução_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    Estoques_Evolução_Linha_Dados = Estoques_Evolução_Linha_Dados + 1
    Estoques_Evolução_Linha_Análise = Estoques_Evolução_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior e define a variável Coluna_DadosP que representa a primeira coluna de dados
Estoques_Evolução_Coluna_Dados1 = Sheets("ESTOQUES (evolução)").Range("A9").End(xlToRight).Column
Estoques_Evolução_Coluna_DadosP = 2

Do Until Estoques_Evolução_Coluna_DadosP = Estoques_Evolução_Coluna_Dados1 + 1 ' Faz até a variavel Coluna_DadosP ser igual a variavel Coluna_Dados1 mais uma unidade
    'Confere se o mês da coluna em questão é igual ao mês do dado mais recente
    If Month(Sheets("ESTOQUES (evolução)").Cells(8, Estoques_Evolução_Coluna_DadosP)) = Month(Sheets("ESTOQUES (evolução)").Cells(8, Estoques_Evolução_Coluna_Dados1)) Then
    'Caso seja igual, copia a coluna com os dados mais abaixo, a partir da linha 110
        Sheets("ESTOQUES (evolução)").Range(Cells(9, Estoques_Evolução_Coluna_DadosP), (Cells(54, Estoques_Evolução_Coluna_DadosP))).Copy (Sheets("ESTOQUES (evolução)").Cells(110, Estoques_Evolução_Coluna_DadosP))
    End If
    'Vai para a próxima coluna
    Estoques_Evolução_Coluna_DadosP = Estoques_Evolução_Coluna_DadosP + 1
'Repete a conferencia com a próxima coluna
Loop


'Atribui os valores originais das variaveis após o loop anterior
Estoques_Evolução_Linha_Dados = 110
Estoques_Evolução_Coluna_Dados1 = Sheets("ESTOQUES (evolução)").Range("A9").End(xlToRight).Column 'Define o número da última coluna
Estoques_Evolução_Linha_Análise = 59
Estoques_Evolução_Coluna_Análise = 7

'Ordenação decrescente da série histórica dos meses do dado mais recente
Do Until Estoques_Evolução_Linha_Dados = 156 'Faz o calculo até a variável Linha_Dados ser 156
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo mês do mais recente
    posição = Application.WorksheetFunction.Rank_Eq(Sheets("ESTOQUES (evolução)").Cells(Estoques_Evolução_Linha_Dados, Estoques_Evolução_Coluna_Dados1), Sheets("ESTOQUES (evolução)").Range(Cells(Estoques_Evolução_Linha_Dados, Estoques_Evolução_Coluna_Dados1), Cells(Estoques_Evolução_Linha_Dados, 2)))
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("ESTOQUES (evolução)").Cells(Estoques_Evolução_Linha_Análise, Estoques_Evolução_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    Estoques_Evolução_Linha_Dados = Estoques_Evolução_Linha_Dados + 1
    Estoques_Evolução_Linha_Análise = Estoques_Evolução_Linha_Análise + 1
'Repete a ordenação com a próxima linha
Loop

'Atribui os valores originais das variaveis após o loop anterior
Estoques_Evolução_Linha_Dados = 110
Estoques_Evolução_Coluna_Dados1 = Sheets("ESTOQUES (evolução)").Range("A9").End(xlToRight).Column 'Define o número da última coluna
Estoques_Evolução_Linha_Análise = 59
Estoques_Evolução_Coluna_Análise = 8
'Ordenação crescente da série histórica dos meses do dado mais recente
Do Until Estoques_Evolução_Linha_Dados = 156 'Faz o calculo até a variável Linha_Dados ser 156
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo mês do mais recente
    posição = Application.WorksheetFunction.Rank_Eq(Sheets("ESTOQUES (evolução)").Cells(Estoques_Evolução_Linha_Dados, Estoques_Evolução_Coluna_Dados1), Sheets("ESTOQUES (evolução)").Range(Cells(Estoques_Evolução_Linha_Dados, Estoques_Evolução_Coluna_Dados1), Cells(Estoques_Evolução_Linha_Dados, 2)), 1)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("ESTOQUES (evolução)").Cells(Estoques_Evolução_Linha_Análise, Estoques_Evolução_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    Estoques_Evolução_Linha_Dados = Estoques_Evolução_Linha_Dados + 1
    Estoques_Evolução_Linha_Análise = Estoques_Evolução_Linha_Análise + 1
'Repete a ordenação com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Estoques_Evolução_Linha_Dados = 9
Estoques_Evolução_Coluna_Dados1 = Sheets("ESTOQUES (evolução)").Range("A9").End(xlToRight).Column
Estoques_Evolução_Coluna_Dados2 = Estoques_Evolução_Coluna_Dados1 - 1
Estoques_Evolução_Linha_Análise = 59
Estoques_Evolução_Coluna_Análise = 9

'Avaliação se cruzou ou não a linha de 50 e o sentido
Do Until Estoques_Evolução_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'se o dado mais recente for menor que 50 e o dado do mês anterior for maior ou igual a 50 então...
    If Cells(Estoques_Evolução_Linha_Dados, Estoques_Evolução_Coluna_Dados1) < 50 And Cells(Estoques_Evolução_Linha_Dados, Estoques_Evolução_Coluna_Dados2) >= 50 Then
    
    'a célula de análise recebe cruzou para baixo
    Cells(Estoques_Evolução_Linha_Análise, Estoques_Evolução_Coluna_Análise).Value = "Cruzou para baixo"
    'Caso não seja..
    Else
        'se o dado mais recente for maior ou igual a 50 e o dado do m~es anterior for menor ou igual a 50 então...
        If Cells(Estoques_Evolução_Linha_Dados, Estoques_Evolução_Coluna_Dados1) >= 50 And Cells(Estoques_Evolução_Linha_Dados, Estoques_Evolução_Coluna_Dados2) <= 50 Then
        'a célula de análise recebe cruzou para cima
        Cells(Estoques_Evolução_Linha_Análise, Estoques_Evolução_Coluna_Análise).Value = "Cruzou para cima"
        'Caso não seja..
        Else
        'a célula de análise recebe não cruzou
        Cells(Estoques_Evolução_Linha_Análise, Estoques_Evolução_Coluna_Análise).Value = "Não Cruzou"
        End If
    End If
    'Vai para a próxima linha de dados e de análise
    Estoques_Evolução_Linha_Dados = Estoques_Evolução_Linha_Dados + 1
    Estoques_Evolução_Linha_Análise = Estoques_Evolução_Linha_Análise + 1
'repete o processo com a nova linha
Loop

'Apaga as linhas com erros/dados faltantes/títulos e subtítulos
Range(Cells(60, 2), Cells(60, 9)).ClearContents
Range(Cells(66, 2), Cells(66, 9)).ClearContents
Range(Cells(70, 2), Cells(70, 9)).ClearContents
Range(Cells(72, 2), Cells(73, 9)).Value = "-"
Range(Cells(75, 2), Cells(75, 9)).Value = "-"
Range(Cells(79, 2), Cells(79, 9)).Value = "-"
Range(Cells(87, 2), Cells(87, 9)).Value = "-"

'Inserindo valores nas celulas vazias para fugir de bugs
Range(Cells(10, Estoques_Evolução_Coluna_Dados3), Cells(10, Estoques_Evolução_Coluna_Dados1)).ClearContents
Range(Cells(16, Estoques_Evolução_Coluna_Dados3), Cells(16, Estoques_Evolução_Coluna_Dados1)).ClearContents
Range(Cells(20, Estoques_Evolução_Coluna_Dados3), Cells(20, Estoques_Evolução_Coluna_Dados1)).ClearContents
Range(Cells(22, Estoques_Evolução_Coluna_Dados3), Cells(23, Estoques_Evolução_Coluna_Dados1)).Value = "-"
Range(Cells(25, Estoques_Evolução_Coluna_Dados3), Cells(25, Estoques_Evolução_Coluna_Dados1)).Value = "-"
Range(Cells(29, Estoques_Evolução_Coluna_Dados3), Cells(29, Estoques_Evolução_Coluna_Dados1)).Value = "-"
Range(Cells(37, Estoques_Evolução_Coluna_Dados3), Cells(37, Estoques_Evolução_Coluna_Dados1)).Value = "-"

Range("E59:H104").NumberFormat = "0"





'*******************************************           Estoques_Efetivo_Planejado           *******************************************************
  
 
 

Dim Estoques_Efetivo_Planejado_Linha_Dados As Integer 'Define a linha que contém o dado a ser usado
Dim Estoques_Efetivo_Planejado_Coluna_Dados1 As Integer ' Define a coluna com o dado mais recente
Dim Estoques_Efetivo_Planejado_Coluna_Dados2 As Integer ' Define a coluna com o dado do mês anterior
Dim Estoques_Efetivo_Planejado_Coluna_Dados3 As Integer ' Defie a coluna com o dado do mesmo mês do ano anterior
Dim Estoques_Efetivo_Planejado_Linha_Análise As Integer ' Define a linha que será feita a análise
Dim Estoques_Efetivo_Planejado_Coluna_Análise As Integer 'Define a coluna que será feita a análise
Sheets("Estoques (efetivo-planejado)").Select


'Copia os títulos das categorias e cola onde será formada a tabela de análise
Sheets("Estoques (efetivo-planejado)").Range(Cells(9, 1), Cells(54, 9)).Copy (Sheets("Estoques (efetivo-planejado)").Cells(59, 1))
'Limpa os números que foram colados mas mantém a formatação
Sheets("Estoques (efetivo-planejado)").Range(Cells(59, 2), Cells(105, 9)).ClearContents



'Nomeia as colunas de acordo com o dado que será calculado nelas
Sheets("Estoques (efetivo-planejado)").Cells(58, 2).Value = "Diferença para o mês anterior"
Sheets("Estoques (efetivo-planejado)").Cells(58, 3).Value = "Diferença para ao mesmo mês do ano anterior"
Sheets("Estoques (efetivo-planejado)").Cells(58, 4).Value = "Diferença para a média histórica"
Sheets("Estoques (efetivo-planejado)").Cells(58, 5).Value = "Posição Decrescente (Maior valor 1º, menor valor último)"
Sheets("Estoques (efetivo-planejado)").Cells(58, 6).Value = "Posição Crescente  (Menor valor 1º, maior valor último)"
Sheets("Estoques (efetivo-planejado)").Cells(58, 7).Value = "Posição Crescente - Mesmo mês  (Menor valor 1º, maior valor último)"
Sheets("Estoques (efetivo-planejado)").Cells(58, 8).Value = "Posição Decrescente -Mesmo mês  (Maior valor 1º, menor valor último)"
Sheets("Estoques (efetivo-planejado)").Cells(58, 9).Value = "cruzou a linha de 50?"

'Atribui valores as variaveis definidas acima
Estoques_Efetivo_Planejado_Linha_Dados = 9 'Define o número da primeira linha de dados
Estoques_Efetivo_Planejado_Coluna_Dados1 = Sheets("Estoques (efetivo-planejado)").Range("A9").End(xlToRight).Column 'Define o número da última coluna
Estoques_Efetivo_Planejado_Coluna_Dados2 = Estoques_Efetivo_Planejado_Coluna_Dados1 - 1 'Define o número da coluna do mês anterior
Estoques_Efetivo_Planejado_Coluna_Dados3 = Estoques_Efetivo_Planejado_Coluna_Dados1 - 12
Estoques_Efetivo_Planejado_Linha_Análise = 59 'Define a primeira linhas de análises
Estoques_Efetivo_Planejado_Coluna_Análise = 2 'Define a coluna de análises

'Inserindo valores nas celulas vazias para fugir de bugs
Sheets("Estoques (efetivo-planejado)").Range(Cells(10, Estoques_Efetivo_Planejado_Coluna_Dados3), Cells(10, Estoques_Efetivo_Planejado_Coluna_Dados1)).Value = "0"
Sheets("Estoques (efetivo-planejado)").Range(Cells(16, Estoques_Efetivo_Planejado_Coluna_Dados3), Cells(16, Estoques_Efetivo_Planejado_Coluna_Dados1)).Value = "0"
Sheets("Estoques (efetivo-planejado)").Range(Cells(20, Estoques_Efetivo_Planejado_Coluna_Dados3), Cells(20, Estoques_Efetivo_Planejado_Coluna_Dados1)).Value = "0"
Sheets("Estoques (efetivo-planejado)").Range(Cells(22, Estoques_Efetivo_Planejado_Coluna_Dados3), Cells(23, Estoques_Efetivo_Planejado_Coluna_Dados1)).Value = "0"
Sheets("Estoques (efetivo-planejado)").Range(Cells(25, Estoques_Efetivo_Planejado_Coluna_Dados3), Cells(25, Estoques_Efetivo_Planejado_Coluna_Dados1)).Value = "0"
Sheets("Estoques (efetivo-planejado)").Range(Cells(29, Estoques_Efetivo_Planejado_Coluna_Dados3), Cells(29, Estoques_Efetivo_Planejado_Coluna_Dados1)).Value = "0"
Sheets("Estoques (efetivo-planejado)").Range(Cells(37, Estoques_Efetivo_Planejado_Coluna_Dados3), Cells(37, Estoques_Efetivo_Planejado_Coluna_Dados1)).Value = "0"


'Calculo da difernça em pontos do valor mais recente em relação ao valor do mês anterior
Do Until Estoques_Efetivo_Planejado_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Calculo da diferença em si: o valor da celula de analise é igual ao valor mais recente menos o valor do mês anterior
   Sheets("Estoques (efetivo-planejado)").Cells(Estoques_Efetivo_Planejado_Linha_Análise, Estoques_Efetivo_Planejado_Coluna_Análise).Value = Sheets("Estoques (efetivo-planejado)").Cells(Estoques_Efetivo_Planejado_Linha_Dados, Estoques_Efetivo_Planejado_Coluna_Dados1).Value - Sheets("Estoques (efetivo-planejado)").Cells(Estoques_Efetivo_Planejado_Linha_Dados, Estoques_Efetivo_Planejado_Coluna_Dados2).Value
    'Vai para a próxima linha de dados e de análise
   Estoques_Efetivo_Planejado_Linha_Dados = Estoques_Efetivo_Planejado_Linha_Dados + 1
   Estoques_Efetivo_Planejado_Linha_Análise = Estoques_Efetivo_Planejado_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Estoques_Efetivo_Planejado_Linha_Dados = 9
Estoques_Efetivo_Planejado_Coluna_Dados1 = Sheets("Estoques (efetivo-planejado)").Range("A9").End(xlToRight).Column
Estoques_Efetivo_Planejado_Coluna_Dados3 = Estoques_Efetivo_Planejado_Coluna_Dados1 - 12
Estoques_Efetivo_Planejado_Linha_Análise = 59
Estoques_Efetivo_Planejado_Coluna_Análise = 3

'Cálculo da diferença em pontos do valor mais recente em relação ao valor do mesmo mês do ano anterior
Do Until Estoques_Efetivo_Planejado_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Calculo da diferença em si: o valor da celula de analise é igual ao valor mais recente menos o valor do mesmo mês do ano anterior
    Sheets("Estoques (efetivo-planejado)").Cells(Estoques_Efetivo_Planejado_Linha_Análise, Estoques_Efetivo_Planejado_Coluna_Análise).Value = Sheets("Estoques (efetivo-planejado)").Cells(Estoques_Efetivo_Planejado_Linha_Dados, Estoques_Efetivo_Planejado_Coluna_Dados1).Value - Sheets("Estoques (efetivo-planejado)").Cells(Estoques_Efetivo_Planejado_Linha_Dados, Estoques_Efetivo_Planejado_Coluna_Dados3).Value
    'Vai para a próxima linha de dados e de análise
    Estoques_Efetivo_Planejado_Linha_Dados = Estoques_Efetivo_Planejado_Linha_Dados + 1
    Estoques_Efetivo_Planejado_Linha_Análise = Estoques_Efetivo_Planejado_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Estoques_Efetivo_Planejado_Linha_Dados = 9
Estoques_Efetivo_Planejado_Coluna_Dados1 = Sheets("Estoques (efetivo-planejado)").Range("A9").End(xlToRight).Column 'Define o número da última coluna
Estoques_Efetivo_Planejado_Linha_Análise = 59
Estoques_Efetivo_Planejado_Coluna_Análise = 4

'Cálculo da diferença em pontos do valor mais recente em relação ao valor da média histórica
Do Until Estoques_Efetivo_Planejado_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a variável media como a média do intervalo entre a coluna com o dado mais recente e o primeiro
    media = Application.Average(Sheets("Estoques (efetivo-planejado)").Range(Cells(Estoques_Efetivo_Planejado_Linha_Dados, Estoques_Efetivo_Planejado_Coluna_Dados1), Cells(Estoques_Efetivo_Planejado_Linha_Dados, 2)))
    'Calculo em si: o valor da celula de analise é igual ao valor mais recente menos o valor da média
    Sheets("Estoques (efetivo-planejado)").Cells(Estoques_Efetivo_Planejado_Linha_Análise, Estoques_Efetivo_Planejado_Coluna_Análise).Value = Sheets("Estoques (efetivo-planejado)").Cells(Estoques_Efetivo_Planejado_Linha_Dados, Estoques_Efetivo_Planejado_Coluna_Dados1).Value - media
    'Vai para a próxima linha de dados e de análise
    Estoques_Efetivo_Planejado_Linha_Dados = Estoques_Efetivo_Planejado_Linha_Dados + 1
    Estoques_Efetivo_Planejado_Linha_Análise = Estoques_Efetivo_Planejado_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Estoques_Efetivo_Planejado_Linha_Dados = 9
Estoques_Efetivo_Planejado_Coluna_Dados1 = Sheets("Estoques (efetivo-planejado)").Range("A9").End(xlToRight).Column
Estoques_Efetivo_Planejado_Linha_Análise = 59
Estoques_Efetivo_Planejado_Coluna_Análise = 5

'Ordenação decrescente da série histórica completa
Do Until Estoques_Efetivo_Planejado_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posição = WorksheetFunction.Rank_Eq(Sheets("Estoques (efetivo-planejado)").Cells(Estoques_Efetivo_Planejado_Linha_Dados, Estoques_Efetivo_Planejado_Coluna_Dados1), Sheets("Estoques (efetivo-planejado)").Range(Cells(Estoques_Efetivo_Planejado_Linha_Dados, Estoques_Efetivo_Planejado_Coluna_Dados1), Cells(Estoques_Efetivo_Planejado_Linha_Dados, 2)), 0)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("Estoques (efetivo-planejado)").Cells(Estoques_Efetivo_Planejado_Linha_Análise, Estoques_Efetivo_Planejado_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    Estoques_Efetivo_Planejado_Linha_Dados = Estoques_Efetivo_Planejado_Linha_Dados + 1
    Estoques_Efetivo_Planejado_Linha_Análise = Estoques_Efetivo_Planejado_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Estoques_Efetivo_Planejado_Linha_Dados = 9
Estoques_Efetivo_Planejado_Coluna_Dados1 = Sheets("Estoques (efetivo-planejado)").Range("A9").End(xlToRight).Column
Estoques_Efetivo_Planejado_Linha_Análise = 59
Estoques_Efetivo_Planejado_Coluna_Análise = 6

'Ordenação Crescente da série histórica completa
Do Until Estoques_Efetivo_Planejado_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posição = WorksheetFunction.Rank_Eq(Sheets("Estoques (efetivo-planejado)").Cells(Estoques_Efetivo_Planejado_Linha_Dados, Estoques_Efetivo_Planejado_Coluna_Dados1), Sheets("Estoques (efetivo-planejado)").Range(Cells(Estoques_Efetivo_Planejado_Linha_Dados, Estoques_Efetivo_Planejado_Coluna_Dados1), Cells(Estoques_Efetivo_Planejado_Linha_Dados, 2)), 1)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("Estoques (efetivo-planejado)").Cells(Estoques_Efetivo_Planejado_Linha_Análise, Estoques_Efetivo_Planejado_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    Estoques_Efetivo_Planejado_Linha_Dados = Estoques_Efetivo_Planejado_Linha_Dados + 1
    Estoques_Efetivo_Planejado_Linha_Análise = Estoques_Efetivo_Planejado_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior e define a variável Coluna_DadosP que representa a primeira coluna de dados
Estoques_Efetivo_Planejado_Coluna_Dados1 = Sheets("Estoques (efetivo-planejado)").Range("A9").End(xlToRight).Column
Estoques_Efetivo_Planejado_Coluna_DadosP = 2

Do Until Estoques_Efetivo_Planejado_Coluna_DadosP = Estoques_Efetivo_Planejado_Coluna_Dados1 + 1 ' Faz até a variavel Coluna_DadosP ser igual a variavel Coluna_Dados1 mais uma unidade
    'Confere se o mês da coluna em questão é igual ao mês do dado mais recente
    If Month(Sheets("Estoques (efetivo-planejado)").Cells(8, Estoques_Efetivo_Planejado_Coluna_DadosP)) = Month(Sheets("Estoques (efetivo-planejado)").Cells(8, Estoques_Efetivo_Planejado_Coluna_Dados1)) Then
    'Caso seja igual, copia a coluna com os dados mais abaixo, a partir da linha 110
        Sheets("Estoques (efetivo-planejado)").Range(Cells(9, Estoques_Efetivo_Planejado_Coluna_DadosP), (Cells(54, Estoques_Efetivo_Planejado_Coluna_DadosP))).Copy (Sheets("Estoques (efetivo-planejado)").Cells(110, Estoques_Efetivo_Planejado_Coluna_DadosP))
    End If
    'Vai para a próxima coluna
    Estoques_Efetivo_Planejado_Coluna_DadosP = Estoques_Efetivo_Planejado_Coluna_DadosP + 1
'Repete a conferencia com a próxima coluna
Loop


'Atribui os valores originais das variaveis após o loop anterior
Estoques_Efetivo_Planejado_Linha_Dados = 110
Estoques_Efetivo_Planejado_Coluna_Dados1 = Sheets("Estoques (efetivo-planejado)").Range("A9").End(xlToRight).Column 'Define o número da última coluna
Estoques_Efetivo_Planejado_Linha_Análise = 59
Estoques_Efetivo_Planejado_Coluna_Análise = 7

'Ordenação decrescente da série histórica dos meses do dado mais recente
Do Until Estoques_Efetivo_Planejado_Linha_Dados = 156 'Faz o calculo até a variável Linha_Dados ser 156
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo mês do mais recente
    posição = Application.WorksheetFunction.Rank_Eq(Sheets("Estoques (efetivo-planejado)").Cells(Estoques_Efetivo_Planejado_Linha_Dados, Estoques_Efetivo_Planejado_Coluna_Dados1), Sheets("Estoques (efetivo-planejado)").Range(Cells(Estoques_Efetivo_Planejado_Linha_Dados, Estoques_Efetivo_Planejado_Coluna_Dados1), Cells(Estoques_Efetivo_Planejado_Linha_Dados, 2)))
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("Estoques (efetivo-planejado)").Cells(Estoques_Efetivo_Planejado_Linha_Análise, Estoques_Efetivo_Planejado_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    Estoques_Efetivo_Planejado_Linha_Dados = Estoques_Efetivo_Planejado_Linha_Dados + 1
    Estoques_Efetivo_Planejado_Linha_Análise = Estoques_Efetivo_Planejado_Linha_Análise + 1
'Repete a ordenação com a próxima linha
Loop

'Atribui os valores originais das variaveis após o loop anterior
Estoques_Efetivo_Planejado_Linha_Dados = 110
Estoques_Efetivo_Planejado_Coluna_Dados1 = Sheets("Estoques (efetivo-planejado)").Range("A9").End(xlToRight).Column 'Define o número da última coluna
Estoques_Efetivo_Planejado_Linha_Análise = 59
Estoques_Efetivo_Planejado_Coluna_Análise = 8
'Ordenação crescente da série histórica dos meses do dado mais recente
Do Until Estoques_Efetivo_Planejado_Linha_Dados = 156 'Faz o calculo até a variável Linha_Dados ser 156
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo mês do mais recente
    posição = Application.WorksheetFunction.Rank_Eq(Sheets("Estoques (efetivo-planejado)").Cells(Estoques_Efetivo_Planejado_Linha_Dados, Estoques_Efetivo_Planejado_Coluna_Dados1), Sheets("Estoques (efetivo-planejado)").Range(Cells(Estoques_Efetivo_Planejado_Linha_Dados, Estoques_Efetivo_Planejado_Coluna_Dados1), Cells(Estoques_Efetivo_Planejado_Linha_Dados, 2)), 1)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("Estoques (efetivo-planejado)").Cells(Estoques_Efetivo_Planejado_Linha_Análise, Estoques_Efetivo_Planejado_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    Estoques_Efetivo_Planejado_Linha_Dados = Estoques_Efetivo_Planejado_Linha_Dados + 1
    Estoques_Efetivo_Planejado_Linha_Análise = Estoques_Efetivo_Planejado_Linha_Análise + 1
'Repete a ordenação com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Estoques_Efetivo_Planejado_Linha_Dados = 9
Estoques_Efetivo_Planejado_Coluna_Dados1 = Sheets("Estoques (efetivo-planejado)").Range("A9").End(xlToRight).Column
Estoques_Efetivo_Planejado_Coluna_Dados2 = Estoques_Efetivo_Planejado_Coluna_Dados1 - 1
Estoques_Efetivo_Planejado_Linha_Análise = 59
Estoques_Efetivo_Planejado_Coluna_Análise = 9

'Avaliação se cruzou ou não a linha de 50 e o sentido
Do Until Estoques_Efetivo_Planejado_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'se o dado mais recente for menor que 50 e o dado do mês anterior for maior ou igual a 50 então...
    If Cells(Estoques_Efetivo_Planejado_Linha_Dados, Estoques_Efetivo_Planejado_Coluna_Dados1) < 50 And Cells(Estoques_Efetivo_Planejado_Linha_Dados, Estoques_Efetivo_Planejado_Coluna_Dados2) >= 50 Then
    
    'a célula de análise recebe cruzou para baixo
    Cells(Estoques_Efetivo_Planejado_Linha_Análise, Estoques_Efetivo_Planejado_Coluna_Análise).Value = "Cruzou para baixo"
    'Caso não seja..
    Else
        'se o dado mais recente for maior ou igual a 50 e o dado do m~es anterior for menor ou igual a 50 então...
        If Cells(Estoques_Efetivo_Planejado_Linha_Dados, Estoques_Efetivo_Planejado_Coluna_Dados1) >= 50 And Cells(Estoques_Efetivo_Planejado_Linha_Dados, Estoques_Efetivo_Planejado_Coluna_Dados2) <= 50 Then
        'a célula de análise recebe cruzou para cima
        Cells(Estoques_Efetivo_Planejado_Linha_Análise, Estoques_Efetivo_Planejado_Coluna_Análise).Value = "Cruzou para cima"
        'Caso não seja..
        Else
        'a célula de análise recebe não cruzou
        Cells(Estoques_Efetivo_Planejado_Linha_Análise, Estoques_Efetivo_Planejado_Coluna_Análise).Value = "Não Cruzou"
        End If
    End If
    'Vai para a próxima linha de dados e de análise
    Estoques_Efetivo_Planejado_Linha_Dados = Estoques_Efetivo_Planejado_Linha_Dados + 1
    Estoques_Efetivo_Planejado_Linha_Análise = Estoques_Efetivo_Planejado_Linha_Análise + 1
'repete o processo com a nova linha
Loop

'Apaga as linhas com erros/dados faltantes/títulos e subtítulos
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


Sub Análise_Azul()

 
Dim Expectativas_Demanda_Linha_Dados As Integer 'Define a linha que contém o dado a ser usado
Dim Expectativas_Demanda_Coluna_Dados1 As Integer ' Define a coluna com o dado mais recente
Dim Expectativas_Demanda_Coluna_Dados2 As Integer ' Define a coluna com o dado do mês anterior
Dim Expectativas_Demanda_Coluna_Dados3 As Integer ' Defie a coluna com o dado do mesmo mês do ano anterior
Dim Expectativas_Demanda_Linha_Análise As Integer ' Define a linha que será feita a análise
Dim Expectativas_Demanda_Coluna_Análise As Integer 'Define a coluna que será feita a análise

Sheets("EXPECTATIVAS - DEMANDA").Select

'Copia os títulos das categorias e cola onde será formada a tabela de análise
Sheets("EXPECTATIVAS - DEMANDA").Range(Cells(9, 1), Cells(54, 9)).Copy (Sheets("EXPECTATIVAS - DEMANDA").Cells(59, 1))
'Limpa os números que foram colados mas mantém a formatação
Sheets("EXPECTATIVAS - DEMANDA").Range(Cells(59, 2), Cells(105, 9)).ClearContents



'Nomeia as colunas de acordo com o dado que será calculado nelas
Sheets("EXPECTATIVAS - DEMANDA").Cells(58, 2).Value = "Diferença para o mês anterior"
Sheets("EXPECTATIVAS - DEMANDA").Cells(58, 3).Value = "Diferença para ao mesmo mês do ano anterior"
Sheets("EXPECTATIVAS - DEMANDA").Cells(58, 4).Value = "Diferença para a média histórica"
Sheets("EXPECTATIVAS - DEMANDA").Cells(58, 5).Value = "Posição Decrescente (Maior valor 1º, menor valor último)"
Sheets("EXPECTATIVAS - DEMANDA").Cells(58, 6).Value = "Posição Crescente  (Menor valor 1º, maior valor último)"
Sheets("EXPECTATIVAS - DEMANDA").Cells(58, 7).Value = "Posição Crescente - Mesmo mês  (Menor valor 1º, maior valor último)"
Sheets("EXPECTATIVAS - DEMANDA").Cells(58, 8).Value = "Posição Decrescente -Mesmo mês  (Maior valor 1º, menor valor último)"
Sheets("EXPECTATIVAS - DEMANDA").Cells(58, 9).Value = "cruzou a linha de 50?"

'Atribui valores as variaveis definidas acima
Expectativas_Demanda_Linha_Dados = 9 'Define o número da primeira linha de dados
Expectativas_Demanda_Coluna_Dados1 = Sheets("EXPECTATIVAS - DEMANDA").Range("A9").End(xlToRight).Column 'Define o número da última coluna
Expectativas_Demanda_Coluna_Dados2 = Expectativas_Demanda_Coluna_Dados1 - 1 'Define o número da coluna do mês anterior
Expectativas_Demanda_Coluna_Dados3 = Expectativas_Demanda_Coluna_Dados1 - 12
Expectativas_Demanda_Linha_Análise = 59 'Define a primeira linhas de análises
Expectativas_Demanda_Coluna_Análise = 2 'Define a coluna de análises

'Inserindo valores nas celulas vazias para fugir de bugs
Sheets("EXPECTATIVAS - DEMANDA").Range(Cells(10, Expectativas_Demanda_Coluna_Dados3), Cells(10, Expectativas_Demanda_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVAS - DEMANDA").Range(Cells(16, Expectativas_Demanda_Coluna_Dados3), Cells(16, Expectativas_Demanda_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVAS - DEMANDA").Range(Cells(20, Expectativas_Demanda_Coluna_Dados3), Cells(20, Expectativas_Demanda_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVAS - DEMANDA").Range(Cells(22, Expectativas_Demanda_Coluna_Dados3), Cells(23, Expectativas_Demanda_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVAS - DEMANDA").Range(Cells(25, Expectativas_Demanda_Coluna_Dados3), Cells(25, Expectativas_Demanda_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVAS - DEMANDA").Range(Cells(29, Expectativas_Demanda_Coluna_Dados3), Cells(29, Expectativas_Demanda_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVAS - DEMANDA").Range(Cells(37, Expectativas_Demanda_Coluna_Dados3), Cells(37, Expectativas_Demanda_Coluna_Dados1)).Value = "0"


'Calculo da difernça em pontos do valor mais recente em relação ao valor do mês anterior
Do Until Expectativas_Demanda_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Calculo da diferença em si: o valor da celula de analise é igual ao valor mais recente menos o valor do mês anterior
   Sheets("EXPECTATIVAS - DEMANDA").Cells(Expectativas_Demanda_Linha_Análise, Expectativas_Demanda_Coluna_Análise).Value = Sheets("EXPECTATIVAS - DEMANDA").Cells(Expectativas_Demanda_Linha_Dados, Expectativas_Demanda_Coluna_Dados1).Value - Sheets("EXPECTATIVAS - DEMANDA").Cells(Expectativas_Demanda_Linha_Dados, Expectativas_Demanda_Coluna_Dados2).Value
    'Vai para a próxima linha de dados e de análise
   Expectativas_Demanda_Linha_Dados = Expectativas_Demanda_Linha_Dados + 1
   Expectativas_Demanda_Linha_Análise = Expectativas_Demanda_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Expectativas_Demanda_Linha_Dados = 9
Expectativas_Demanda_Coluna_Dados1 = Sheets("EXPECTATIVAS - DEMANDA").Range("A9").End(xlToRight).Column
Expectativas_Demanda_Coluna_Dados3 = Expectativas_Demanda_Coluna_Dados1 - 12
Expectativas_Demanda_Linha_Análise = 59
Expectativas_Demanda_Coluna_Análise = 3

'Cálculo da diferença em pontos do valor mais recente em relação ao valor do mesmo mês do ano anterior
Do Until Expectativas_Demanda_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Calculo da diferença em si: o valor da celula de analise é igual ao valor mais recente menos o valor do mesmo mês do ano anterior
    Sheets("EXPECTATIVAS - DEMANDA").Cells(Expectativas_Demanda_Linha_Análise, Expectativas_Demanda_Coluna_Análise).Value = Sheets("EXPECTATIVAS - DEMANDA").Cells(Expectativas_Demanda_Linha_Dados, Expectativas_Demanda_Coluna_Dados1).Value - Sheets("EXPECTATIVAS - DEMANDA").Cells(Expectativas_Demanda_Linha_Dados, Expectativas_Demanda_Coluna_Dados3).Value
    'Vai para a próxima linha de dados e de análise
    Expectativas_Demanda_Linha_Dados = Expectativas_Demanda_Linha_Dados + 1
    Expectativas_Demanda_Linha_Análise = Expectativas_Demanda_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Expectativas_Demanda_Linha_Dados = 9
Expectativas_Demanda_Coluna_Dados1 = Sheets("EXPECTATIVAS - DEMANDA").Range("A9").End(xlToRight).Column 'Define o número da última coluna
Expectativas_Demanda_Linha_Análise = 59
Expectativas_Demanda_Coluna_Análise = 4

'Cálculo da diferença em pontos do valor mais recente em relação ao valor da média histórica
Do Until Expectativas_Demanda_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a variável media como a média do intervalo entre a coluna com o dado mais recente e o primeiro
    media = Application.Average(Sheets("EXPECTATIVAS - DEMANDA").Range(Cells(Expectativas_Demanda_Linha_Dados, Expectativas_Demanda_Coluna_Dados1), Cells(Expectativas_Demanda_Linha_Dados, 2)))
    'Calculo em si: o valor da celula de analise é igual ao valor mais recente menos o valor da média
    Sheets("EXPECTATIVAS - DEMANDA").Cells(Expectativas_Demanda_Linha_Análise, Expectativas_Demanda_Coluna_Análise).Value = Sheets("EXPECTATIVAS - DEMANDA").Cells(Expectativas_Demanda_Linha_Dados, Expectativas_Demanda_Coluna_Dados1).Value - media
    'Vai para a próxima linha de dados e de análise
    Expectativas_Demanda_Linha_Dados = Expectativas_Demanda_Linha_Dados + 1
    Expectativas_Demanda_Linha_Análise = Expectativas_Demanda_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Expectativas_Demanda_Linha_Dados = 9
Expectativas_Demanda_Coluna_Dados1 = Sheets("EXPECTATIVAS - DEMANDA").Range("A9").End(xlToRight).Column
Expectativas_Demanda_Linha_Análise = 59
Expectativas_Demanda_Coluna_Análise = 5

'Ordenação decrescente da série histórica completa
Do Until Expectativas_Demanda_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posição = WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVAS - DEMANDA").Cells(Expectativas_Demanda_Linha_Dados, Expectativas_Demanda_Coluna_Dados1), Sheets("EXPECTATIVAS - DEMANDA").Range(Cells(Expectativas_Demanda_Linha_Dados, Expectativas_Demanda_Coluna_Dados1), Cells(Expectativas_Demanda_Linha_Dados, 2)), 0)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("EXPECTATIVAS - DEMANDA").Cells(Expectativas_Demanda_Linha_Análise, Expectativas_Demanda_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    Expectativas_Demanda_Linha_Dados = Expectativas_Demanda_Linha_Dados + 1
    Expectativas_Demanda_Linha_Análise = Expectativas_Demanda_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Expectativas_Demanda_Linha_Dados = 9
Expectativas_Demanda_Coluna_Dados1 = Sheets("EXPECTATIVAS - DEMANDA").Range("A9").End(xlToRight).Column
Expectativas_Demanda_Linha_Análise = 59
Expectativas_Demanda_Coluna_Análise = 6

'Ordenação Crescente da série histórica completa
Do Until Expectativas_Demanda_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posição = WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVAS - DEMANDA").Cells(Expectativas_Demanda_Linha_Dados, Expectativas_Demanda_Coluna_Dados1), Sheets("EXPECTATIVAS - DEMANDA").Range(Cells(Expectativas_Demanda_Linha_Dados, Expectativas_Demanda_Coluna_Dados1), Cells(Expectativas_Demanda_Linha_Dados, 2)), 1)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("EXPECTATIVAS - DEMANDA").Cells(Expectativas_Demanda_Linha_Análise, Expectativas_Demanda_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    Expectativas_Demanda_Linha_Dados = Expectativas_Demanda_Linha_Dados + 1
    Expectativas_Demanda_Linha_Análise = Expectativas_Demanda_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior e define a variável Coluna_DadosP que representa a primeira coluna de dados
Expectativas_Demanda_Coluna_Dados1 = Sheets("EXPECTATIVAS - DEMANDA").Range("A9").End(xlToRight).Column
Expectativas_Demanda_Coluna_DadosP = 2

Do Until Expectativas_Demanda_Coluna_DadosP = Expectativas_Demanda_Coluna_Dados1 + 1 ' Faz até a variavel Coluna_DadosP ser igual a variavel Coluna_Dados1 mais uma unidade
    'Confere se o mês da coluna em questão é igual ao mês do dado mais recente
    If Month(Sheets("EXPECTATIVAS - DEMANDA").Cells(8, Expectativas_Demanda_Coluna_DadosP)) = Month(Sheets("EXPECTATIVAS - DEMANDA").Cells(8, Expectativas_Demanda_Coluna_Dados1)) Then
    'Caso seja igual, copia a coluna com os dados mais abaixo, a partir da linha 110
        Sheets("EXPECTATIVAS - DEMANDA").Range(Cells(9, Expectativas_Demanda_Coluna_DadosP), (Cells(54, Expectativas_Demanda_Coluna_DadosP))).Copy (Sheets("EXPECTATIVAS - DEMANDA").Cells(110, Expectativas_Demanda_Coluna_DadosP))
    End If
    'Vai para a próxima coluna
    Expectativas_Demanda_Coluna_DadosP = Expectativas_Demanda_Coluna_DadosP + 1
'Repete a conferencia com a próxima coluna
Loop


'Atribui os valores originais das variaveis após o loop anterior
Expectativas_Demanda_Linha_Dados = 110
Expectativas_Demanda_Coluna_Dados1 = Sheets("EXPECTATIVAS - DEMANDA").Range("A9").End(xlToRight).Column 'Define o número da última coluna
Expectativas_Demanda_Linha_Análise = 59
Expectativas_Demanda_Coluna_Análise = 7

'Ordenação decrescente da série histórica dos meses do dado mais recente
Do Until Expectativas_Demanda_Linha_Dados = 156 'Faz o calculo até a variável Linha_Dados ser 156
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo mês do mais recente
    posição = Application.WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVAS - DEMANDA").Cells(Expectativas_Demanda_Linha_Dados, Expectativas_Demanda_Coluna_Dados1), Sheets("EXPECTATIVAS - DEMANDA").Range(Cells(Expectativas_Demanda_Linha_Dados, Expectativas_Demanda_Coluna_Dados1), Cells(Expectativas_Demanda_Linha_Dados, 2)))
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("EXPECTATIVAS - DEMANDA").Cells(Expectativas_Demanda_Linha_Análise, Expectativas_Demanda_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    Expectativas_Demanda_Linha_Dados = Expectativas_Demanda_Linha_Dados + 1
    Expectativas_Demanda_Linha_Análise = Expectativas_Demanda_Linha_Análise + 1
'Repete a ordenação com a próxima linha
Loop

'Atribui os valores originais das variaveis após o loop anterior
Expectativas_Demanda_Linha_Dados = 110
Expectativas_Demanda_Coluna_Dados1 = Sheets("EXPECTATIVAS - DEMANDA").Range("A9").End(xlToRight).Column 'Define o número da última coluna
Expectativas_Demanda_Linha_Análise = 59
Expectativas_Demanda_Coluna_Análise = 8
'Ordenação crescente da série histórica dos meses do dado mais recente
Do Until Expectativas_Demanda_Linha_Dados = 156 'Faz o calculo até a variável Linha_Dados ser 156
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo mês do mais recente
    posição = Application.WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVAS - DEMANDA").Cells(Expectativas_Demanda_Linha_Dados, Expectativas_Demanda_Coluna_Dados1), Sheets("EXPECTATIVAS - DEMANDA").Range(Cells(Expectativas_Demanda_Linha_Dados, Expectativas_Demanda_Coluna_Dados1), Cells(Expectativas_Demanda_Linha_Dados, 2)), 1)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("EXPECTATIVAS - DEMANDA").Cells(Expectativas_Demanda_Linha_Análise, Expectativas_Demanda_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    Expectativas_Demanda_Linha_Dados = Expectativas_Demanda_Linha_Dados + 1
    Expectativas_Demanda_Linha_Análise = Expectativas_Demanda_Linha_Análise + 1
'Repete a ordenação com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Expectativas_Demanda_Linha_Dados = 9
Expectativas_Demanda_Coluna_Dados1 = Sheets("EXPECTATIVAS - DEMANDA").Range("A9").End(xlToRight).Column
Expectativas_Demanda_Coluna_Dados2 = Expectativas_Demanda_Coluna_Dados1 - 1
Expectativas_Demanda_Linha_Análise = 59
Expectativas_Demanda_Coluna_Análise = 9

'Avaliação se cruzou ou não a linha de 50 e o sentido
Do Until Expectativas_Demanda_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'se o dado mais recente for menor que 50 e o dado do mês anterior for maior ou igual a 50 então...
    If Cells(Expectativas_Demanda_Linha_Dados, Expectativas_Demanda_Coluna_Dados1) < 50 And Cells(Expectativas_Demanda_Linha_Dados, Expectativas_Demanda_Coluna_Dados2) >= 50 Then
    
    'a célula de análise recebe cruzou para baixo
    Cells(Expectativas_Demanda_Linha_Análise, Expectativas_Demanda_Coluna_Análise).Value = "Cruzou para baixo"
    'Caso não seja..
    Else
        'se o dado mais recente for maior ou igual a 50 e o dado do m~es anterior for menor ou igual a 50 então...
        If Cells(Expectativas_Demanda_Linha_Dados, Expectativas_Demanda_Coluna_Dados1) >= 50 And Cells(Expectativas_Demanda_Linha_Dados, Expectativas_Demanda_Coluna_Dados2) <= 50 Then
        'a célula de análise recebe cruzou para cima
        Cells(Expectativas_Demanda_Linha_Análise, Expectativas_Demanda_Coluna_Análise).Value = "Cruzou para cima"
        'Caso não seja..
        Else
        'a célula de análise recebe não cruzou
        Cells(Expectativas_Demanda_Linha_Análise, Expectativas_Demanda_Coluna_Análise).Value = "Não Cruzou"
        End If
    End If
    'Vai para a próxima linha de dados e de análise
    Expectativas_Demanda_Linha_Dados = Expectativas_Demanda_Linha_Dados + 1
    Expectativas_Demanda_Linha_Análise = Expectativas_Demanda_Linha_Análise + 1
'repete o processo com a nova linha
Loop

'Apaga as linhas com erros/dados faltantes/títulos e subtítulos
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


'******************************************             Expectativas_Exportação            *********************************************************

Dim Expectativas_Exportação_Linha_Dados As Integer 'Define a linha que contém o dado a ser usado
Dim Expectativas_Exportação_Coluna_Dados1 As Integer ' Define a coluna com o dado mais recente
Dim Expectativas_Exportação_Coluna_Dados2 As Integer ' Define a coluna com o dado do mês anterior
Dim Expectativas_Exportação_Coluna_Dados3 As Integer ' Defie a coluna com o dado do mesmo mês do ano anterior
Dim Expectativas_Exportação_Linha_Análise As Integer ' Define a linha que será feita a análise
Dim Expectativas_Exportação_Coluna_Análise As Integer 'Define a coluna que será feita a análise

Sheets("EXPECTATIVA - EXPORTAÇÃO").Select


'Copia os títulos das categorias e cola onde será formada a tabela de análise
Sheets("EXPECTATIVA - EXPORTAÇÃO").Range(Cells(9, 1), Cells(54, 9)).Copy (Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(59, 1))
'Limpa os números que foram colados mas mantém a formatação
Sheets("EXPECTATIVA - EXPORTAÇÃO").Range(Cells(59, 2), Cells(105, 9)).ClearContents



'Nomeia as colunas de acordo com o dado que será calculado nelas
Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(58, 2).Value = "Diferença para o mês anterior"
Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(58, 3).Value = "Diferença para ao mesmo mês do ano anterior"
Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(58, 4).Value = "Diferença para a média histórica"
Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(58, 5).Value = "Posição Decrescente (Maior valor 1º, menor valor último)"
Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(58, 6).Value = "Posição Crescente  (Menor valor 1º, maior valor último)"
Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(58, 7).Value = "Posição Crescente - Mesmo mês  (Menor valor 1º, maior valor último)"
Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(58, 8).Value = "Posição Decrescente -Mesmo mês  (Maior valor 1º, menor valor último)"
Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(58, 9).Value = "cruzou a linha de 50?"

'Atribui valores as variaveis definidas acima
Expectativas_Exportação_Linha_Dados = 9 'Define o número da primeira linha de dados
Expectativas_Exportação_Coluna_Dados1 = Sheets("EXPECTATIVA - EXPORTAÇÃO").Range("A9").End(xlToRight).Column 'Define o número da última coluna
Expectativas_Exportação_Coluna_Dados2 = Expectativas_Exportação_Coluna_Dados1 - 1 'Define o número da coluna do mês anterior
Expectativas_Exportação_Coluna_Dados3 = Expectativas_Exportação_Coluna_Dados1 - 12
Expectativas_Exportação_Linha_Análise = 59 'Define a primeira linhas de análises
Expectativas_Exportação_Coluna_Análise = 2 'Define a coluna de análises

'Inserindo valores nas celulas vazias para fugir de bugs
Sheets("EXPECTATIVA - EXPORTAÇÃO").Range(Cells(10, Expectativas_Exportação_Coluna_Dados3), Cells(10, Expectativas_Exportação_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - EXPORTAÇÃO").Range(Cells(16, Expectativas_Exportação_Coluna_Dados3), Cells(16, Expectativas_Exportação_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - EXPORTAÇÃO").Range(Cells(20, Expectativas_Exportação_Coluna_Dados3), Cells(20, Expectativas_Exportação_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - EXPORTAÇÃO").Range(Cells(22, Expectativas_Exportação_Coluna_Dados3), Cells(23, Expectativas_Exportação_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - EXPORTAÇÃO").Range(Cells(25, Expectativas_Exportação_Coluna_Dados3), Cells(25, Expectativas_Exportação_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - EXPORTAÇÃO").Range(Cells(29, Expectativas_Exportação_Coluna_Dados3), Cells(29, Expectativas_Exportação_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - EXPORTAÇÃO").Range(Cells(37, Expectativas_Exportação_Coluna_Dados3), Cells(37, Expectativas_Exportação_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - EXPORTAÇÃO").Range(Cells(54, Expectativas_Exportação_Coluna_Dados3), Cells(54, Expectativas_Exportação_Coluna_Dados1)).Value = "0"

'Calculo da difernça em pontos do valor mais recente em relação ao valor do mês anterior
Do Until Expectativas_Exportação_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Calculo da diferença em si: o valor da celula de analise é igual ao valor mais recente menos o valor do mês anterior
   Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(Expectativas_Exportação_Linha_Análise, Expectativas_Exportação_Coluna_Análise).Value = Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(Expectativas_Exportação_Linha_Dados, Expectativas_Exportação_Coluna_Dados1).Value - Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(Expectativas_Exportação_Linha_Dados, Expectativas_Exportação_Coluna_Dados2).Value
    'Vai para a próxima linha de dados e de análise
   Expectativas_Exportação_Linha_Dados = Expectativas_Exportação_Linha_Dados + 1
   Expectativas_Exportação_Linha_Análise = Expectativas_Exportação_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Expectativas_Exportação_Linha_Dados = 9
Expectativas_Exportação_Coluna_Dados1 = Sheets("EXPECTATIVA - EXPORTAÇÃO").Range("A9").End(xlToRight).Column
Expectativas_Exportação_Coluna_Dados3 = Expectativas_Exportação_Coluna_Dados1 - 12
Expectativas_Exportação_Linha_Análise = 59
Expectativas_Exportação_Coluna_Análise = 3

'Cálculo da diferença em pontos do valor mais recente em relação ao valor do mesmo mês do ano anterior
Do Until Expectativas_Exportação_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Calculo da diferença em si: o valor da celula de analise é igual ao valor mais recente menos o valor do mesmo mês do ano anterior
    Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(Expectativas_Exportação_Linha_Análise, Expectativas_Exportação_Coluna_Análise).Value = Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(Expectativas_Exportação_Linha_Dados, Expectativas_Exportação_Coluna_Dados1).Value - Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(Expectativas_Exportação_Linha_Dados, Expectativas_Exportação_Coluna_Dados3).Value
    'Vai para a próxima linha de dados e de análise
    Expectativas_Exportação_Linha_Dados = Expectativas_Exportação_Linha_Dados + 1
    Expectativas_Exportação_Linha_Análise = Expectativas_Exportação_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Expectativas_Exportação_Linha_Dados = 9
Expectativas_Exportação_Coluna_Dados1 = Sheets("EXPECTATIVA - EXPORTAÇÃO").Range("A9").End(xlToRight).Column 'Define o número da última coluna
Expectativas_Exportação_Linha_Análise = 59
Expectativas_Exportação_Coluna_Análise = 4

'Cálculo da diferença em pontos do valor mais recente em relação ao valor da média histórica
Do Until Expectativas_Exportação_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a variável media como a média do intervalo entre a coluna com o dado mais recente e o primeiro
    media = Application.Average(Sheets("EXPECTATIVA - EXPORTAÇÃO").Range(Cells(Expectativas_Exportação_Linha_Dados, Expectativas_Exportação_Coluna_Dados1), Cells(Expectativas_Exportação_Linha_Dados, 2)))
    'Calculo em si: o valor da celula de analise é igual ao valor mais recente menos o valor da média
    Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(Expectativas_Exportação_Linha_Análise, Expectativas_Exportação_Coluna_Análise).Value = Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(Expectativas_Exportação_Linha_Dados, Expectativas_Exportação_Coluna_Dados1).Value - media
    'Vai para a próxima linha de dados e de análise
    Expectativas_Exportação_Linha_Dados = Expectativas_Exportação_Linha_Dados + 1
    Expectativas_Exportação_Linha_Análise = Expectativas_Exportação_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Expectativas_Exportação_Linha_Dados = 9
Expectativas_Exportação_Coluna_Dados1 = Sheets("EXPECTATIVA - EXPORTAÇÃO").Range("A9").End(xlToRight).Column
Expectativas_Exportação_Linha_Análise = 59
Expectativas_Exportação_Coluna_Análise = 5

'Ordenação decrescente da série histórica completa
Do Until Expectativas_Exportação_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posição = WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(Expectativas_Exportação_Linha_Dados, Expectativas_Exportação_Coluna_Dados1), Sheets("EXPECTATIVA - EXPORTAÇÃO").Range(Cells(Expectativas_Exportação_Linha_Dados, Expectativas_Exportação_Coluna_Dados1), Cells(Expectativas_Exportação_Linha_Dados, 2)), 0)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(Expectativas_Exportação_Linha_Análise, Expectativas_Exportação_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    Expectativas_Exportação_Linha_Dados = Expectativas_Exportação_Linha_Dados + 1
    Expectativas_Exportação_Linha_Análise = Expectativas_Exportação_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Expectativas_Exportação_Linha_Dados = 9
Expectativas_Exportação_Coluna_Dados1 = Sheets("EXPECTATIVA - EXPORTAÇÃO").Range("A9").End(xlToRight).Column
Expectativas_Exportação_Linha_Análise = 59
Expectativas_Exportação_Coluna_Análise = 6

'Ordenação Crescente da série histórica completa
Do Until Expectativas_Exportação_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posição = WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(Expectativas_Exportação_Linha_Dados, Expectativas_Exportação_Coluna_Dados1), Sheets("EXPECTATIVA - EXPORTAÇÃO").Range(Cells(Expectativas_Exportação_Linha_Dados, Expectativas_Exportação_Coluna_Dados1), Cells(Expectativas_Exportação_Linha_Dados, 2)), 1)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(Expectativas_Exportação_Linha_Análise, Expectativas_Exportação_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    Expectativas_Exportação_Linha_Dados = Expectativas_Exportação_Linha_Dados + 1
    Expectativas_Exportação_Linha_Análise = Expectativas_Exportação_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior e define a variável Coluna_DadosP que representa a primeira coluna de dados
Expectativas_Exportação_Coluna_Dados1 = Sheets("EXPECTATIVA - EXPORTAÇÃO").Range("A9").End(xlToRight).Column
Expectativas_Exportação_Coluna_DadosP = 2

Do Until Expectativas_Exportação_Coluna_DadosP = Expectativas_Exportação_Coluna_Dados1 + 1 ' Faz até a variavel Coluna_DadosP ser igual a variavel Coluna_Dados1 mais uma unidade
    'Confere se o mês da coluna em questão é igual ao mês do dado mais recente
    If Month(Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(8, Expectativas_Exportação_Coluna_DadosP)) = Month(Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(8, Expectativas_Exportação_Coluna_Dados1)) Then
    'Caso seja igual, copia a coluna com os dados mais abaixo, a partir da linha 110
        Sheets("EXPECTATIVA - EXPORTAÇÃO").Range(Cells(9, Expectativas_Exportação_Coluna_DadosP), (Cells(54, Expectativas_Exportação_Coluna_DadosP))).Copy (Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(110, Expectativas_Exportação_Coluna_DadosP))
    End If
    'Vai para a próxima coluna
    Expectativas_Exportação_Coluna_DadosP = Expectativas_Exportação_Coluna_DadosP + 1
'Repete a conferencia com a próxima coluna
Loop


'Atribui os valores originais das variaveis após o loop anterior
Expectativas_Exportação_Linha_Dados = 110
Expectativas_Exportação_Coluna_Dados1 = Sheets("EXPECTATIVA - EXPORTAÇÃO").Range("A9").End(xlToRight).Column 'Define o número da última coluna
Expectativas_Exportação_Linha_Análise = 59
Expectativas_Exportação_Coluna_Análise = 7

'Ordenação decrescente da série histórica dos meses do dado mais recente
Do Until Expectativas_Exportação_Linha_Dados = 156 'Faz o calculo até a variável Linha_Dados ser 156
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo mês do mais recente
    posição = Application.WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(Expectativas_Exportação_Linha_Dados, Expectativas_Exportação_Coluna_Dados1), Sheets("EXPECTATIVA - EXPORTAÇÃO").Range(Cells(Expectativas_Exportação_Linha_Dados, Expectativas_Exportação_Coluna_Dados1), Cells(Expectativas_Exportação_Linha_Dados, 2)))
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(Expectativas_Exportação_Linha_Análise, Expectativas_Exportação_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    Expectativas_Exportação_Linha_Dados = Expectativas_Exportação_Linha_Dados + 1
    Expectativas_Exportação_Linha_Análise = Expectativas_Exportação_Linha_Análise + 1
'Repete a ordenação com a próxima linha
Loop

'Atribui os valores originais das variaveis após o loop anterior
Expectativas_Exportação_Linha_Dados = 110
Expectativas_Exportação_Coluna_Dados1 = Sheets("EXPECTATIVA - EXPORTAÇÃO").Range("A9").End(xlToRight).Column 'Define o número da última coluna
Expectativas_Exportação_Linha_Análise = 59
Expectativas_Exportação_Coluna_Análise = 8
'Ordenação crescente da série histórica dos meses do dado mais recente
Do Until Expectativas_Exportação_Linha_Dados = 156 'Faz o calculo até a variável Linha_Dados ser 156
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo mês do mais recente
    posição = Application.WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(Expectativas_Exportação_Linha_Dados, Expectativas_Exportação_Coluna_Dados1), Sheets("EXPECTATIVA - EXPORTAÇÃO").Range(Cells(Expectativas_Exportação_Linha_Dados, Expectativas_Exportação_Coluna_Dados1), Cells(Expectativas_Exportação_Linha_Dados, 2)), 1)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("EXPECTATIVA - EXPORTAÇÃO").Cells(Expectativas_Exportação_Linha_Análise, Expectativas_Exportação_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    Expectativas_Exportação_Linha_Dados = Expectativas_Exportação_Linha_Dados + 1
    Expectativas_Exportação_Linha_Análise = Expectativas_Exportação_Linha_Análise + 1
'Repete a ordenação com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Expectativas_Exportação_Linha_Dados = 9
Expectativas_Exportação_Coluna_Dados1 = Sheets("EXPECTATIVA - EXPORTAÇÃO").Range("A9").End(xlToRight).Column
Expectativas_Exportação_Coluna_Dados2 = Expectativas_Exportação_Coluna_Dados1 - 1
Expectativas_Exportação_Linha_Análise = 59
Expectativas_Exportação_Coluna_Análise = 9

'Avaliação se cruzou ou não a linha de 50 e o sentido
Do Until Expectativas_Exportação_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'se o dado mais recente for menor que 50 e o dado do mês anterior for maior ou igual a 50 então...
    If Cells(Expectativas_Exportação_Linha_Dados, Expectativas_Exportação_Coluna_Dados1) < 50 And Cells(Expectativas_Exportação_Linha_Dados, Expectativas_Exportação_Coluna_Dados2) >= 50 Then
    
    'a célula de análise recebe cruzou para baixo
    Cells(Expectativas_Exportação_Linha_Análise, Expectativas_Exportação_Coluna_Análise).Value = "Cruzou para baixo"
    'Caso não seja..
    Else
        'se o dado mais recente for maior ou igual a 50 e o dado do m~es anterior for menor ou igual a 50 então...
        If Cells(Expectativas_Exportação_Linha_Dados, Expectativas_Exportação_Coluna_Dados1) >= 50 And Cells(Expectativas_Exportação_Linha_Dados, Expectativas_Exportação_Coluna_Dados2) <= 50 Then
        'a célula de análise recebe cruzou para cima
        Cells(Expectativas_Exportação_Linha_Análise, Expectativas_Exportação_Coluna_Análise).Value = "Cruzou para cima"
        'Caso não seja..
        Else
        'a célula de análise recebe não cruzou
        Cells(Expectativas_Exportação_Linha_Análise, Expectativas_Exportação_Coluna_Análise).Value = "Não Cruzou"
        End If
    End If
    'Vai para a próxima linha de dados e de análise
    Expectativas_Exportação_Linha_Dados = Expectativas_Exportação_Linha_Dados + 1
    Expectativas_Exportação_Linha_Análise = Expectativas_Exportação_Linha_Análise + 1
'repete o processo com a nova linha
Loop

'Apaga as linhas com erros/dados faltantes/títulos e subtítulos
Range(Cells(60, 2), Cells(60, 9)).ClearContents
Range(Cells(66, 2), Cells(66, 9)).ClearContents
Range(Cells(70, 2), Cells(70, 9)).ClearContents
Range(Cells(72, 2), Cells(73, 9)).Value = "-"
Range(Cells(75, 2), Cells(75, 9)).Value = "-"
Range(Cells(79, 2), Cells(79, 9)).Value = "-"
Range(Cells(87, 2), Cells(87, 9)).Value = "-"
Range(Cells(104, 2), Cells(104, 9)).Value = "-"

'Inserindo valores nas celulas vazias para fugir de bugs
Range(Cells(10, Expectativas_Exportação_Coluna_Dados3), Cells(10, Expectativas_Exportação_Coluna_Dados1)).ClearContents
Range(Cells(16, Expectativas_Exportação_Coluna_Dados3), Cells(16, Expectativas_Exportação_Coluna_Dados1)).ClearContents
Range(Cells(20, Expectativas_Exportação_Coluna_Dados3), Cells(20, Expectativas_Exportação_Coluna_Dados1)).ClearContents
Range(Cells(22, Expectativas_Exportação_Coluna_Dados3), Cells(23, Expectativas_Exportação_Coluna_Dados1)).Value = "-"
Range(Cells(25, Expectativas_Exportação_Coluna_Dados3), Cells(25, Expectativas_Exportação_Coluna_Dados1)).Value = "-"
Range(Cells(29, Expectativas_Exportação_Coluna_Dados3), Cells(29, Expectativas_Exportação_Coluna_Dados1)).Value = "-"
Range(Cells(37, Expectativas_Exportação_Coluna_Dados3), Cells(37, Expectativas_Exportação_Coluna_Dados1)).Value = "-"
Range(Cells(54, Expectativas_Exportação_Coluna_Dados3), Cells(54, Expectativas_Exportação_Coluna_Dados1)).Value = "-"


Range("E59:H104").NumberFormat = "0"


'****************************************      Expectativa_Compras                ***********************************************************/


Dim Expectativas_Compras_Linha_Dados As Integer 'Define a linha que contém o dado a ser usado
Dim Expectativas_Compras_Coluna_Dados1 As Integer ' Define a coluna com o dado mais recente
Dim Expectativas_Compras_Coluna_Dados2 As Integer ' Define a coluna com o dado do mês anterior
Dim Expectativas_Compras_Coluna_Dados3 As Integer ' Defie a coluna com o dado do mesmo mês do ano anterior
Dim Expectativas_Compras_Linha_Análise As Integer ' Define a linha que será feita a análise
Dim Expectativas_Compras_Coluna_Análise As Integer 'Define a coluna que será feita a análise


Sheets("EXPECTATIVA - COMPRAS").Select



'Copia os títulos das categorias e cola onde será formada a tabela de análise
Sheets("EXPECTATIVA - COMPRAS").Range(Cells(9, 1), Cells(54, 9)).Copy (Sheets("EXPECTATIVA - COMPRAS").Cells(59, 1))
'Limpa os números que foram colados mas mantém a formatação
Sheets("EXPECTATIVA - COMPRAS").Range(Cells(59, 2), Cells(105, 9)).ClearContents



'Nomeia as colunas de acordo com o dado que será calculado nelas
Sheets("EXPECTATIVA - COMPRAS").Cells(58, 2).Value = "Diferença para o mês anterior"
Sheets("EXPECTATIVA - COMPRAS").Cells(58, 3).Value = "Diferença para ao mesmo mês do ano anterior"
Sheets("EXPECTATIVA - COMPRAS").Cells(58, 4).Value = "Diferença para a média histórica"
Sheets("EXPECTATIVA - COMPRAS").Cells(58, 5).Value = "Posição Decrescente (Maior valor 1º, menor valor último)"
Sheets("EXPECTATIVA - COMPRAS").Cells(58, 6).Value = "Posição Crescente  (Menor valor 1º, maior valor último)"
Sheets("EXPECTATIVA - COMPRAS").Cells(58, 7).Value = "Posição Crescente - Mesmo mês  (Menor valor 1º, maior valor último)"
Sheets("EXPECTATIVA - COMPRAS").Cells(58, 8).Value = "Posição Decrescente -Mesmo mês  (Maior valor 1º, menor valor último)"
Sheets("EXPECTATIVA - COMPRAS").Cells(58, 9).Value = "cruzou a linha de 50?"

'Atribui valores as variaveis definidas acima
Expectativas_Compras_Linha_Dados = 9 'Define o número da primeira linha de dados
Expectativas_Compras_Coluna_Dados1 = Sheets("EXPECTATIVA - COMPRAS").Range("A9").End(xlToRight).Column 'Define o número da última coluna
Expectativas_Compras_Coluna_Dados2 = Expectativas_Compras_Coluna_Dados1 - 1 'Define o número da coluna do mês anterior
Expectativas_Compras_Coluna_Dados3 = Expectativas_Compras_Coluna_Dados1 - 12
Expectativas_Compras_Linha_Análise = 59 'Define a primeira linhas de análises
Expectativas_Compras_Coluna_Análise = 2 'Define a coluna de análises

'Inserindo valores nas celulas vazias para fugir de bugs
Sheets("EXPECTATIVA - COMPRAS").Range(Cells(10, Expectativas_Compras_Coluna_Dados3), Cells(10, Expectativas_Compras_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - COMPRAS").Range(Cells(16, Expectativas_Compras_Coluna_Dados3), Cells(16, Expectativas_Compras_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - COMPRAS").Range(Cells(20, Expectativas_Compras_Coluna_Dados3), Cells(20, Expectativas_Compras_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - COMPRAS").Range(Cells(22, Expectativas_Compras_Coluna_Dados3), Cells(23, Expectativas_Compras_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - COMPRAS").Range(Cells(25, Expectativas_Compras_Coluna_Dados3), Cells(25, Expectativas_Compras_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - COMPRAS").Range(Cells(29, Expectativas_Compras_Coluna_Dados3), Cells(29, Expectativas_Compras_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - COMPRAS").Range(Cells(37, Expectativas_Compras_Coluna_Dados3), Cells(37, Expectativas_Compras_Coluna_Dados1)).Value = "0"


'Calculo da difernça em pontos do valor mais recente em relação ao valor do mês anterior
Do Until Expectativas_Compras_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Calculo da diferença em si: o valor da celula de analise é igual ao valor mais recente menos o valor do mês anterior
   Sheets("EXPECTATIVA - COMPRAS").Cells(Expectativas_Compras_Linha_Análise, Expectativas_Compras_Coluna_Análise).Value = Sheets("EXPECTATIVA - COMPRAS").Cells(Expectativas_Compras_Linha_Dados, Expectativas_Compras_Coluna_Dados1).Value - Sheets("EXPECTATIVA - COMPRAS").Cells(Expectativas_Compras_Linha_Dados, Expectativas_Compras_Coluna_Dados2).Value
    'Vai para a próxima linha de dados e de análise
   Expectativas_Compras_Linha_Dados = Expectativas_Compras_Linha_Dados + 1
   Expectativas_Compras_Linha_Análise = Expectativas_Compras_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Expectativas_Compras_Linha_Dados = 9
Expectativas_Compras_Coluna_Dados1 = Sheets("EXPECTATIVA - COMPRAS").Range("A9").End(xlToRight).Column
Expectativas_Compras_Coluna_Dados3 = Expectativas_Compras_Coluna_Dados1 - 12
Expectativas_Compras_Linha_Análise = 59
Expectativas_Compras_Coluna_Análise = 3

'Cálculo da diferença em pontos do valor mais recente em relação ao valor do mesmo mês do ano anterior
Do Until Expectativas_Compras_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Calculo da diferença em si: o valor da celula de analise é igual ao valor mais recente menos o valor do mesmo mês do ano anterior
    Sheets("EXPECTATIVA - COMPRAS").Cells(Expectativas_Compras_Linha_Análise, Expectativas_Compras_Coluna_Análise).Value = Sheets("EXPECTATIVA - COMPRAS").Cells(Expectativas_Compras_Linha_Dados, Expectativas_Compras_Coluna_Dados1).Value - Sheets("EXPECTATIVA - COMPRAS").Cells(Expectativas_Compras_Linha_Dados, Expectativas_Compras_Coluna_Dados3).Value
    'Vai para a próxima linha de dados e de análise
    Expectativas_Compras_Linha_Dados = Expectativas_Compras_Linha_Dados + 1
    Expectativas_Compras_Linha_Análise = Expectativas_Compras_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Expectativas_Compras_Linha_Dados = 9
Expectativas_Compras_Coluna_Dados1 = Sheets("EXPECTATIVA - COMPRAS").Range("A9").End(xlToRight).Column 'Define o número da última coluna
Expectativas_Compras_Linha_Análise = 59
Expectativas_Compras_Coluna_Análise = 4

'Cálculo da diferença em pontos do valor mais recente em relação ao valor da média histórica
Do Until Expectativas_Compras_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a variável media como a média do intervalo entre a coluna com o dado mais recente e o primeiro
    media = Application.Average(Sheets("EXPECTATIVA - COMPRAS").Range(Cells(Expectativas_Compras_Linha_Dados, Expectativas_Compras_Coluna_Dados1), Cells(Expectativas_Compras_Linha_Dados, 2)))
    'Calculo em si: o valor da celula de analise é igual ao valor mais recente menos o valor da média
    Sheets("EXPECTATIVA - COMPRAS").Cells(Expectativas_Compras_Linha_Análise, Expectativas_Compras_Coluna_Análise).Value = Sheets("EXPECTATIVA - COMPRAS").Cells(Expectativas_Compras_Linha_Dados, Expectativas_Compras_Coluna_Dados1).Value - media
    'Vai para a próxima linha de dados e de análise
    Expectativas_Compras_Linha_Dados = Expectativas_Compras_Linha_Dados + 1
    Expectativas_Compras_Linha_Análise = Expectativas_Compras_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Expectativas_Compras_Linha_Dados = 9
Expectativas_Compras_Coluna_Dados1 = Sheets("EXPECTATIVA - COMPRAS").Range("A9").End(xlToRight).Column
Expectativas_Compras_Linha_Análise = 59
Expectativas_Compras_Coluna_Análise = 5

'Ordenação decrescente da série histórica completa
Do Until Expectativas_Compras_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posição = WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVA - COMPRAS").Cells(Expectativas_Compras_Linha_Dados, Expectativas_Compras_Coluna_Dados1), Sheets("EXPECTATIVA - COMPRAS").Range(Cells(Expectativas_Compras_Linha_Dados, Expectativas_Compras_Coluna_Dados1), Cells(Expectativas_Compras_Linha_Dados, 2)), 0)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("EXPECTATIVA - COMPRAS").Cells(Expectativas_Compras_Linha_Análise, Expectativas_Compras_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    Expectativas_Compras_Linha_Dados = Expectativas_Compras_Linha_Dados + 1
    Expectativas_Compras_Linha_Análise = Expectativas_Compras_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Expectativas_Compras_Linha_Dados = 9
Expectativas_Compras_Coluna_Dados1 = Sheets("EXPECTATIVA - COMPRAS").Range("A9").End(xlToRight).Column
Expectativas_Compras_Linha_Análise = 59
Expectativas_Compras_Coluna_Análise = 6

'Ordenação Crescente da série histórica completa
Do Until Expectativas_Compras_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posição = WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVA - COMPRAS").Cells(Expectativas_Compras_Linha_Dados, Expectativas_Compras_Coluna_Dados1), Sheets("EXPECTATIVA - COMPRAS").Range(Cells(Expectativas_Compras_Linha_Dados, Expectativas_Compras_Coluna_Dados1), Cells(Expectativas_Compras_Linha_Dados, 2)), 1)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("EXPECTATIVA - COMPRAS").Cells(Expectativas_Compras_Linha_Análise, Expectativas_Compras_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    Expectativas_Compras_Linha_Dados = Expectativas_Compras_Linha_Dados + 1
    Expectativas_Compras_Linha_Análise = Expectativas_Compras_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior e define a variável Coluna_DadosP que representa a primeira coluna de dados
Expectativas_Compras_Coluna_Dados1 = Sheets("EXPECTATIVA - COMPRAS").Range("A9").End(xlToRight).Column
Expectativas_Compras_Coluna_DadosP = 2

Do Until Expectativas_Compras_Coluna_DadosP = Expectativas_Compras_Coluna_Dados1 + 1 ' Faz até a variavel Coluna_DadosP ser igual a variavel Coluna_Dados1 mais uma unidade
    'Confere se o mês da coluna em questão é igual ao mês do dado mais recente
    If Month(Sheets("EXPECTATIVA - COMPRAS").Cells(8, Expectativas_Compras_Coluna_DadosP)) = Month(Sheets("EXPECTATIVA - COMPRAS").Cells(8, Expectativas_Compras_Coluna_Dados1)) Then
    'Caso seja igual, copia a coluna com os dados mais abaixo, a partir da linha 110
        Sheets("EXPECTATIVA - COMPRAS").Range(Cells(9, Expectativas_Compras_Coluna_DadosP), (Cells(54, Expectativas_Compras_Coluna_DadosP))).Copy (Sheets("EXPECTATIVA - COMPRAS").Cells(110, Expectativas_Compras_Coluna_DadosP))
    End If
    'Vai para a próxima coluna
    Expectativas_Compras_Coluna_DadosP = Expectativas_Compras_Coluna_DadosP + 1
'Repete a conferencia com a próxima coluna
Loop


'Atribui os valores originais das variaveis após o loop anterior
Expectativas_Compras_Linha_Dados = 110
Expectativas_Compras_Coluna_Dados1 = Sheets("EXPECTATIVA - COMPRAS").Range("A9").End(xlToRight).Column 'Define o número da última coluna
Expectativas_Compras_Linha_Análise = 59
Expectativas_Compras_Coluna_Análise = 7

'Ordenação decrescente da série histórica dos meses do dado mais recente
Do Until Expectativas_Compras_Linha_Dados = 156 'Faz o calculo até a variável Linha_Dados ser 156
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo mês do mais recente
    posição = Application.WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVA - COMPRAS").Cells(Expectativas_Compras_Linha_Dados, Expectativas_Compras_Coluna_Dados1), Sheets("EXPECTATIVA - COMPRAS").Range(Cells(Expectativas_Compras_Linha_Dados, Expectativas_Compras_Coluna_Dados1), Cells(Expectativas_Compras_Linha_Dados, 2)))
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("EXPECTATIVA - COMPRAS").Cells(Expectativas_Compras_Linha_Análise, Expectativas_Compras_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    Expectativas_Compras_Linha_Dados = Expectativas_Compras_Linha_Dados + 1
    Expectativas_Compras_Linha_Análise = Expectativas_Compras_Linha_Análise + 1
'Repete a ordenação com a próxima linha
Loop

'Atribui os valores originais das variaveis após o loop anterior
Expectativas_Compras_Linha_Dados = 110
Expectativas_Compras_Coluna_Dados1 = Sheets("EXPECTATIVA - COMPRAS").Range("A9").End(xlToRight).Column 'Define o número da última coluna
Expectativas_Compras_Linha_Análise = 59
Expectativas_Compras_Coluna_Análise = 8
'Ordenação crescente da série histórica dos meses do dado mais recente
Do Until Expectativas_Compras_Linha_Dados = 156 'Faz o calculo até a variável Linha_Dados ser 156
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo mês do mais recente
    posição = Application.WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVA - COMPRAS").Cells(Expectativas_Compras_Linha_Dados, Expectativas_Compras_Coluna_Dados1), Sheets("EXPECTATIVA - COMPRAS").Range(Cells(Expectativas_Compras_Linha_Dados, Expectativas_Compras_Coluna_Dados1), Cells(Expectativas_Compras_Linha_Dados, 2)), 1)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("EXPECTATIVA - COMPRAS").Cells(Expectativas_Compras_Linha_Análise, Expectativas_Compras_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    Expectativas_Compras_Linha_Dados = Expectativas_Compras_Linha_Dados + 1
    Expectativas_Compras_Linha_Análise = Expectativas_Compras_Linha_Análise + 1
'Repete a ordenação com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Expectativas_Compras_Linha_Dados = 9
Expectativas_Compras_Coluna_Dados1 = Sheets("EXPECTATIVA - COMPRAS").Range("A9").End(xlToRight).Column
Expectativas_Compras_Coluna_Dados2 = Expectativas_Compras_Coluna_Dados1 - 1
Expectativas_Compras_Linha_Análise = 59
Expectativas_Compras_Coluna_Análise = 9

'Avaliação se cruzou ou não a linha de 50 e o sentido
Do Until Expectativas_Compras_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'se o dado mais recente for menor que 50 e o dado do mês anterior for maior ou igual a 50 então...
    If Cells(Expectativas_Compras_Linha_Dados, Expectativas_Compras_Coluna_Dados1) < 50 And Cells(Expectativas_Compras_Linha_Dados, Expectativas_Compras_Coluna_Dados2) >= 50 Then
    
    'a célula de análise recebe cruzou para baixo
    Cells(Expectativas_Compras_Linha_Análise, Expectativas_Compras_Coluna_Análise).Value = "Cruzou para baixo"
    'Caso não seja..
    Else
        'se o dado mais recente for maior ou igual a 50 e o dado do m~es anterior for menor ou igual a 50 então...
        If Cells(Expectativas_Compras_Linha_Dados, Expectativas_Compras_Coluna_Dados1) >= 50 And Cells(Expectativas_Compras_Linha_Dados, Expectativas_Compras_Coluna_Dados2) <= 50 Then
        'a célula de análise recebe cruzou para cima
        Cells(Expectativas_Compras_Linha_Análise, Expectativas_Compras_Coluna_Análise).Value = "Cruzou para cima"
        'Caso não seja..
        Else
        'a célula de análise recebe não cruzou
        Cells(Expectativas_Compras_Linha_Análise, Expectativas_Compras_Coluna_Análise).Value = "Não Cruzou"
        End If
    End If
    'Vai para a próxima linha de dados e de análise
    Expectativas_Compras_Linha_Dados = Expectativas_Compras_Linha_Dados + 1
    Expectativas_Compras_Linha_Análise = Expectativas_Compras_Linha_Análise + 1
'repete o processo com a nova linha
Loop

'Apaga as linhas com erros/dados faltantes/títulos e subtítulos
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

Dim Expectativas_Empregados_Linha_Dados As Integer 'Define a linha que contém o dado a ser usado
Dim Expectativas_Empregados_Coluna_Dados1 As Integer ' Define a coluna com o dado mais recente
Dim Expectativas_Empregados_Coluna_Dados2 As Integer ' Define a coluna com o dado do mês anterior
Dim Expectativas_Empregados_Coluna_Dados3 As Integer ' Defie a coluna com o dado do mesmo mês do ano anterior
Dim Expectativas_Empregados_Linha_Análise As Integer ' Define a linha que será feita a análise
Dim Expectativas_Empregados_Coluna_Análise As Integer 'Define a coluna que será feita a análise

Sheets("EXPECTATIVA - EMPREGADOS").Select


'Copia os títulos das categorias e cola onde será formada a tabela de análise
Sheets("EXPECTATIVA - EMPREGADOS").Range(Cells(9, 1), Cells(54, 9)).Copy (Sheets("EXPECTATIVA - EMPREGADOS").Cells(59, 1))
'Limpa os números que foram colados mas mantém a formatação
Sheets("EXPECTATIVA - EMPREGADOS").Range(Cells(59, 2), Cells(105, 9)).ClearContents



'Nomeia as colunas de acordo com o dado que será calculado nelas
Sheets("EXPECTATIVA - EMPREGADOS").Cells(58, 2).Value = "Diferença para o mês anterior"
Sheets("EXPECTATIVA - EMPREGADOS").Cells(58, 3).Value = "Diferença para ao mesmo mês do ano anterior"
Sheets("EXPECTATIVA - EMPREGADOS").Cells(58, 4).Value = "Diferença para a média histórica"
Sheets("EXPECTATIVA - EMPREGADOS").Cells(58, 5).Value = "Posição Decrescente (Maior valor 1º, menor valor último)"
Sheets("EXPECTATIVA - EMPREGADOS").Cells(58, 6).Value = "Posição Crescente  (Menor valor 1º, maior valor último)"
Sheets("EXPECTATIVA - EMPREGADOS").Cells(58, 7).Value = "Posição Crescente - Mesmo mês  (Menor valor 1º, maior valor último)"
Sheets("EXPECTATIVA - EMPREGADOS").Cells(58, 8).Value = "Posição Decrescente -Mesmo mês  (Maior valor 1º, menor valor último)"
Sheets("EXPECTATIVA - EMPREGADOS").Cells(58, 9).Value = "cruzou a linha de 50?"

'Atribui valores as variaveis definidas acima
Expectativas_Empregados_Linha_Dados = 9 'Define o número da primeira linha de dados
Expectativas_Empregados_Coluna_Dados1 = Sheets("EXPECTATIVA - EMPREGADOS").Range("A9").End(xlToRight).Column 'Define o número da última coluna
Expectativas_Empregados_Coluna_Dados2 = Expectativas_Empregados_Coluna_Dados1 - 1 'Define o número da coluna do mês anterior
Expectativas_Empregados_Coluna_Dados3 = Expectativas_Empregados_Coluna_Dados1 - 12
Expectativas_Empregados_Linha_Análise = 59 'Define a primeira linhas de análises
Expectativas_Empregados_Coluna_Análise = 2 'Define a coluna de análises

'Inserindo valores nas celulas vazias para fugir de bugs
Sheets("EXPECTATIVA - EMPREGADOS").Range(Cells(10, Expectativas_Empregados_Coluna_Dados3), Cells(10, Expectativas_Empregados_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - EMPREGADOS").Range(Cells(16, Expectativas_Empregados_Coluna_Dados3), Cells(16, Expectativas_Empregados_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - EMPREGADOS").Range(Cells(20, Expectativas_Empregados_Coluna_Dados3), Cells(20, Expectativas_Empregados_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - EMPREGADOS").Range(Cells(22, Expectativas_Empregados_Coluna_Dados3), Cells(23, Expectativas_Empregados_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - EMPREGADOS").Range(Cells(25, Expectativas_Empregados_Coluna_Dados3), Cells(25, Expectativas_Empregados_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - EMPREGADOS").Range(Cells(29, Expectativas_Empregados_Coluna_Dados3), Cells(29, Expectativas_Empregados_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - EMPREGADOS").Range(Cells(37, Expectativas_Empregados_Coluna_Dados3), Cells(37, Expectativas_Empregados_Coluna_Dados1)).Value = "0"


'Calculo da difernça em pontos do valor mais recente em relação ao valor do mês anterior
Do Until Expectativas_Empregados_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Calculo da diferença em si: o valor da celula de analise é igual ao valor mais recente menos o valor do mês anterior
   Sheets("EXPECTATIVA - EMPREGADOS").Cells(Expectativas_Empregados_Linha_Análise, Expectativas_Empregados_Coluna_Análise).Value = Sheets("EXPECTATIVA - EMPREGADOS").Cells(Expectativas_Empregados_Linha_Dados, Expectativas_Empregados_Coluna_Dados1).Value - Sheets("EXPECTATIVA - EMPREGADOS").Cells(Expectativas_Empregados_Linha_Dados, Expectativas_Empregados_Coluna_Dados2).Value
    'Vai para a próxima linha de dados e de análise
   Expectativas_Empregados_Linha_Dados = Expectativas_Empregados_Linha_Dados + 1
   Expectativas_Empregados_Linha_Análise = Expectativas_Empregados_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Expectativas_Empregados_Linha_Dados = 9
Expectativas_Empregados_Coluna_Dados1 = Sheets("EXPECTATIVA - EMPREGADOS").Range("A9").End(xlToRight).Column
Expectativas_Empregados_Coluna_Dados3 = Expectativas_Empregados_Coluna_Dados1 - 12
Expectativas_Empregados_Linha_Análise = 59
Expectativas_Empregados_Coluna_Análise = 3

'Cálculo da diferença em pontos do valor mais recente em relação ao valor do mesmo mês do ano anterior
Do Until Expectativas_Empregados_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Calculo da diferença em si: o valor da celula de analise é igual ao valor mais recente menos o valor do mesmo mês do ano anterior
    Sheets("EXPECTATIVA - EMPREGADOS").Cells(Expectativas_Empregados_Linha_Análise, Expectativas_Empregados_Coluna_Análise).Value = Sheets("EXPECTATIVA - EMPREGADOS").Cells(Expectativas_Empregados_Linha_Dados, Expectativas_Empregados_Coluna_Dados1).Value - Sheets("EXPECTATIVA - EMPREGADOS").Cells(Expectativas_Empregados_Linha_Dados, Expectativas_Empregados_Coluna_Dados3).Value
    'Vai para a próxima linha de dados e de análise
    Expectativas_Empregados_Linha_Dados = Expectativas_Empregados_Linha_Dados + 1
    Expectativas_Empregados_Linha_Análise = Expectativas_Empregados_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Expectativas_Empregados_Linha_Dados = 9
Expectativas_Empregados_Coluna_Dados1 = Sheets("EXPECTATIVA - EMPREGADOS").Range("A9").End(xlToRight).Column 'Define o número da última coluna
Expectativas_Empregados_Linha_Análise = 59
Expectativas_Empregados_Coluna_Análise = 4

'Cálculo da diferença em pontos do valor mais recente em relação ao valor da média histórica
Do Until Expectativas_Empregados_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a variável media como a média do intervalo entre a coluna com o dado mais recente e o primeiro
    media = Application.Average(Sheets("EXPECTATIVA - EMPREGADOS").Range(Cells(Expectativas_Empregados_Linha_Dados, Expectativas_Empregados_Coluna_Dados1), Cells(Expectativas_Empregados_Linha_Dados, 2)))
    'Calculo em si: o valor da celula de analise é igual ao valor mais recente menos o valor da média
    Sheets("EXPECTATIVA - EMPREGADOS").Cells(Expectativas_Empregados_Linha_Análise, Expectativas_Empregados_Coluna_Análise).Value = Sheets("EXPECTATIVA - EMPREGADOS").Cells(Expectativas_Empregados_Linha_Dados, Expectativas_Empregados_Coluna_Dados1).Value - media
    'Vai para a próxima linha de dados e de análise
    Expectativas_Empregados_Linha_Dados = Expectativas_Empregados_Linha_Dados + 1
    Expectativas_Empregados_Linha_Análise = Expectativas_Empregados_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Expectativas_Empregados_Linha_Dados = 9
Expectativas_Empregados_Coluna_Dados1 = Sheets("EXPECTATIVA - EMPREGADOS").Range("A9").End(xlToRight).Column
Expectativas_Empregados_Linha_Análise = 59
Expectativas_Empregados_Coluna_Análise = 5

'Ordenação decrescente da série histórica completa
Do Until Expectativas_Empregados_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posição = WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVA - EMPREGADOS").Cells(Expectativas_Empregados_Linha_Dados, Expectativas_Empregados_Coluna_Dados1), Sheets("EXPECTATIVA - EMPREGADOS").Range(Cells(Expectativas_Empregados_Linha_Dados, Expectativas_Empregados_Coluna_Dados1), Cells(Expectativas_Empregados_Linha_Dados, 2)), 0)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("EXPECTATIVA - EMPREGADOS").Cells(Expectativas_Empregados_Linha_Análise, Expectativas_Empregados_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    Expectativas_Empregados_Linha_Dados = Expectativas_Empregados_Linha_Dados + 1
    Expectativas_Empregados_Linha_Análise = Expectativas_Empregados_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Expectativas_Empregados_Linha_Dados = 9
Expectativas_Empregados_Coluna_Dados1 = Sheets("EXPECTATIVA - EMPREGADOS").Range("A9").End(xlToRight).Column
Expectativas_Empregados_Linha_Análise = 59
Expectativas_Empregados_Coluna_Análise = 6

'Ordenação Crescente da série histórica completa
Do Until Expectativas_Empregados_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posição = WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVA - EMPREGADOS").Cells(Expectativas_Empregados_Linha_Dados, Expectativas_Empregados_Coluna_Dados1), Sheets("EXPECTATIVA - EMPREGADOS").Range(Cells(Expectativas_Empregados_Linha_Dados, Expectativas_Empregados_Coluna_Dados1), Cells(Expectativas_Empregados_Linha_Dados, 2)), 1)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("EXPECTATIVA - EMPREGADOS").Cells(Expectativas_Empregados_Linha_Análise, Expectativas_Empregados_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    Expectativas_Empregados_Linha_Dados = Expectativas_Empregados_Linha_Dados + 1
    Expectativas_Empregados_Linha_Análise = Expectativas_Empregados_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior e define a variável Coluna_DadosP que representa a primeira coluna de dados
Expectativas_Empregados_Coluna_Dados1 = Sheets("EXPECTATIVA - EMPREGADOS").Range("A9").End(xlToRight).Column
Expectativas_Empregados_Coluna_DadosP = 2

Do Until Expectativas_Empregados_Coluna_DadosP = Expectativas_Empregados_Coluna_Dados1 + 1 ' Faz até a variavel Coluna_DadosP ser igual a variavel Coluna_Dados1 mais uma unidade
    'Confere se o mês da coluna em questão é igual ao mês do dado mais recente
    If Month(Sheets("EXPECTATIVA - EMPREGADOS").Cells(8, Expectativas_Empregados_Coluna_DadosP)) = Month(Sheets("EXPECTATIVA - EMPREGADOS").Cells(8, Expectativas_Empregados_Coluna_Dados1)) Then
    'Caso seja igual, copia a coluna com os dados mais abaixo, a partir da linha 110
        Sheets("EXPECTATIVA - EMPREGADOS").Range(Cells(9, Expectativas_Empregados_Coluna_DadosP), (Cells(54, Expectativas_Empregados_Coluna_DadosP))).Copy (Sheets("EXPECTATIVA - EMPREGADOS").Cells(110, Expectativas_Empregados_Coluna_DadosP))
    End If
    'Vai para a próxima coluna
    Expectativas_Empregados_Coluna_DadosP = Expectativas_Empregados_Coluna_DadosP + 1
'Repete a conferencia com a próxima coluna
Loop


'Atribui os valores originais das variaveis após o loop anterior
Expectativas_Empregados_Linha_Dados = 110
Expectativas_Empregados_Coluna_Dados1 = Sheets("EXPECTATIVA - EMPREGADOS").Range("A9").End(xlToRight).Column 'Define o número da última coluna
Expectativas_Empregados_Linha_Análise = 59
Expectativas_Empregados_Coluna_Análise = 7

'Ordenação decrescente da série histórica dos meses do dado mais recente
Do Until Expectativas_Empregados_Linha_Dados = 156 'Faz o calculo até a variável Linha_Dados ser 156
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo mês do mais recente
    posição = Application.WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVA - EMPREGADOS").Cells(Expectativas_Empregados_Linha_Dados, Expectativas_Empregados_Coluna_Dados1), Sheets("EXPECTATIVA - EMPREGADOS").Range(Cells(Expectativas_Empregados_Linha_Dados, Expectativas_Empregados_Coluna_Dados1), Cells(Expectativas_Empregados_Linha_Dados, 2)))
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("EXPECTATIVA - EMPREGADOS").Cells(Expectativas_Empregados_Linha_Análise, Expectativas_Empregados_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    Expectativas_Empregados_Linha_Dados = Expectativas_Empregados_Linha_Dados + 1
    Expectativas_Empregados_Linha_Análise = Expectativas_Empregados_Linha_Análise + 1
'Repete a ordenação com a próxima linha
Loop

'Atribui os valores originais das variaveis após o loop anterior
Expectativas_Empregados_Linha_Dados = 110
Expectativas_Empregados_Coluna_Dados1 = Sheets("EXPECTATIVA - EMPREGADOS").Range("A9").End(xlToRight).Column 'Define o número da última coluna
Expectativas_Empregados_Linha_Análise = 59
Expectativas_Empregados_Coluna_Análise = 8
'Ordenação crescente da série histórica dos meses do dado mais recente
Do Until Expectativas_Empregados_Linha_Dados = 156 'Faz o calculo até a variável Linha_Dados ser 156
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo mês do mais recente
    posição = Application.WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVA - EMPREGADOS").Cells(Expectativas_Empregados_Linha_Dados, Expectativas_Empregados_Coluna_Dados1), Sheets("EXPECTATIVA - EMPREGADOS").Range(Cells(Expectativas_Empregados_Linha_Dados, Expectativas_Empregados_Coluna_Dados1), Cells(Expectativas_Empregados_Linha_Dados, 2)), 1)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("EXPECTATIVA - EMPREGADOS").Cells(Expectativas_Empregados_Linha_Análise, Expectativas_Empregados_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    Expectativas_Empregados_Linha_Dados = Expectativas_Empregados_Linha_Dados + 1
    Expectativas_Empregados_Linha_Análise = Expectativas_Empregados_Linha_Análise + 1
'Repete a ordenação com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Expectativas_Empregados_Linha_Dados = 9
Expectativas_Empregados_Coluna_Dados1 = Sheets("EXPECTATIVA - EMPREGADOS").Range("A9").End(xlToRight).Column
Expectativas_Empregados_Coluna_Dados2 = Expectativas_Empregados_Coluna_Dados1 - 1
Expectativas_Empregados_Linha_Análise = 59
Expectativas_Empregados_Coluna_Análise = 9

'Avaliação se cruzou ou não a linha de 50 e o sentido
Do Until Expectativas_Empregados_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'se o dado mais recente for menor que 50 e o dado do mês anterior for maior ou igual a 50 então...
    If Cells(Expectativas_Empregados_Linha_Dados, Expectativas_Empregados_Coluna_Dados1) < 50 And Cells(Expectativas_Empregados_Linha_Dados, Expectativas_Empregados_Coluna_Dados2) >= 50 Then
    
    'a célula de análise recebe cruzou para baixo
    Cells(Expectativas_Empregados_Linha_Análise, Expectativas_Empregados_Coluna_Análise).Value = "Cruzou para baixo"
    'Caso não seja..
    Else
        'se o dado mais recente for maior ou igual a 50 e o dado do m~es anterior for menor ou igual a 50 então...
        If Cells(Expectativas_Empregados_Linha_Dados, Expectativas_Empregados_Coluna_Dados1) >= 50 And Cells(Expectativas_Empregados_Linha_Dados, Expectativas_Empregados_Coluna_Dados2) <= 50 Then
        'a célula de análise recebe cruzou para cima
        Cells(Expectativas_Empregados_Linha_Análise, Expectativas_Empregados_Coluna_Análise).Value = "Cruzou para cima"
        'Caso não seja..
        Else
        'a célula de análise recebe não cruzou
        Cells(Expectativas_Empregados_Linha_Análise, Expectativas_Empregados_Coluna_Análise).Value = "Não Cruzou"
        End If
    End If
    'Vai para a próxima linha de dados e de análise
    Expectativas_Empregados_Linha_Dados = Expectativas_Empregados_Linha_Dados + 1
    Expectativas_Empregados_Linha_Análise = Expectativas_Empregados_Linha_Análise + 1
'repete o processo com a nova linha
Loop

'Apaga as linhas com erros/dados faltantes/títulos e subtítulos
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



Dim Expectativas_Investimentos_Linha_Dados As Integer 'Define a linha que contém o dado a ser usado
Dim Expectativas_Investimentos_Coluna_Dados1 As Integer ' Define a coluna com o dado mais recente
Dim Expectativas_Investimentos_Coluna_Dados2 As Integer ' Define a coluna com o dado do mês anterior
Dim Expectativas_Investimentos_Coluna_Dados3 As Integer ' Defie a coluna com o dado do mesmo mês do ano anterior
Dim Expectativas_Investimentos_Linha_Análise As Integer ' Define a linha que será feita a análise
Dim Expectativas_Investimentos_Coluna_Análise As Integer 'Define a coluna que será feita a análise

Sheets("EXPECTATIVA - INVESTIMENTO").Select


'Copia os títulos das categorias e cola onde será formada a tabela de análise
Sheets("EXPECTATIVA - INVESTIMENTO").Range(Cells(9, 1), Cells(54, 9)).Copy (Sheets("EXPECTATIVA - INVESTIMENTO").Cells(59, 1))
'Limpa os números que foram colados mas mantém a formatação
Sheets("EXPECTATIVA - INVESTIMENTO").Range(Cells(59, 2), Cells(105, 9)).ClearContents



'Nomeia as colunas de acordo com o dado que será calculado nelas
Sheets("EXPECTATIVA - INVESTIMENTO").Cells(58, 2).Value = "Diferença para o mês anterior"
Sheets("EXPECTATIVA - INVESTIMENTO").Cells(58, 3).Value = "Diferença para ao mesmo mês do ano anterior"
Sheets("EXPECTATIVA - INVESTIMENTO").Cells(58, 4).Value = "Diferença para a média histórica"
Sheets("EXPECTATIVA - INVESTIMENTO").Cells(58, 5).Value = "Posição Decrescente (Maior valor 1º, menor valor último)"
Sheets("EXPECTATIVA - INVESTIMENTO").Cells(58, 6).Value = "Posição Crescente  (Menor valor 1º, maior valor último)"
Sheets("EXPECTATIVA - INVESTIMENTO").Cells(58, 7).Value = "Posição Crescente - Mesmo mês  (Menor valor 1º, maior valor último)"
Sheets("EXPECTATIVA - INVESTIMENTO").Cells(58, 8).Value = "Posição Decrescente -Mesmo mês  (Maior valor 1º, menor valor último)"
Sheets("EXPECTATIVA - INVESTIMENTO").Cells(58, 9).Value = "cruzou a linha de 50?"

'Atribui valores as variaveis definidas acima
Expectativas_Investimentos_Linha_Dados = 9 'Define o número da primeira linha de dados
Expectativas_Investimentos_Coluna_Dados1 = Sheets("EXPECTATIVA - INVESTIMENTO").Range("A9").End(xlToRight).Column 'Define o número da última coluna
Expectativas_Investimentos_Coluna_Dados2 = Expectativas_Investimentos_Coluna_Dados1 - 1 'Define o número da coluna do mês anterior
Expectativas_Investimentos_Coluna_Dados3 = Expectativas_Investimentos_Coluna_Dados1 - 12
Expectativas_Investimentos_Linha_Análise = 59 'Define a primeira linhas de análises
Expectativas_Investimentos_Coluna_Análise = 2 'Define a coluna de análises

'Inserindo valores nas celulas vazias para fugir de bugs
Sheets("EXPECTATIVA - INVESTIMENTO").Range(Cells(10, Expectativas_Investimentos_Coluna_Dados3), Cells(10, Expectativas_Investimentos_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - INVESTIMENTO").Range(Cells(16, Expectativas_Investimentos_Coluna_Dados3), Cells(16, Expectativas_Investimentos_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - INVESTIMENTO").Range(Cells(20, Expectativas_Investimentos_Coluna_Dados3), Cells(20, Expectativas_Investimentos_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - INVESTIMENTO").Range(Cells(22, Expectativas_Investimentos_Coluna_Dados3), Cells(23, Expectativas_Investimentos_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - INVESTIMENTO").Range(Cells(25, Expectativas_Investimentos_Coluna_Dados3), Cells(25, Expectativas_Investimentos_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - INVESTIMENTO").Range(Cells(29, Expectativas_Investimentos_Coluna_Dados3), Cells(29, Expectativas_Investimentos_Coluna_Dados1)).Value = "0"
Sheets("EXPECTATIVA - INVESTIMENTO").Range(Cells(37, Expectativas_Investimentos_Coluna_Dados3), Cells(37, Expectativas_Investimentos_Coluna_Dados1)).Value = "0"


'Calculo da difernça em pontos do valor mais recente em relação ao valor do mês anterior
Do Until Expectativas_Investimentos_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Calculo da diferença em si: o valor da celula de analise é igual ao valor mais recente menos o valor do mês anterior
   Sheets("EXPECTATIVA - INVESTIMENTO").Cells(Expectativas_Investimentos_Linha_Análise, Expectativas_Investimentos_Coluna_Análise).Value = Sheets("EXPECTATIVA - INVESTIMENTO").Cells(Expectativas_Investimentos_Linha_Dados, Expectativas_Investimentos_Coluna_Dados1).Value - Sheets("EXPECTATIVA - INVESTIMENTO").Cells(Expectativas_Investimentos_Linha_Dados, Expectativas_Investimentos_Coluna_Dados2).Value
    'Vai para a próxima linha de dados e de análise
   Expectativas_Investimentos_Linha_Dados = Expectativas_Investimentos_Linha_Dados + 1
   Expectativas_Investimentos_Linha_Análise = Expectativas_Investimentos_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Expectativas_Investimentos_Linha_Dados = 9
Expectativas_Investimentos_Coluna_Dados1 = Sheets("EXPECTATIVA - INVESTIMENTO").Range("A9").End(xlToRight).Column
Expectativas_Investimentos_Coluna_Dados3 = Expectativas_Investimentos_Coluna_Dados1 - 12
Expectativas_Investimentos_Linha_Análise = 59
Expectativas_Investimentos_Coluna_Análise = 3

'Cálculo da diferença em pontos do valor mais recente em relação ao valor do mesmo mês do ano anterior
Do Until Expectativas_Investimentos_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Calculo da diferença em si: o valor da celula de analise é igual ao valor mais recente menos o valor do mesmo mês do ano anterior
    Sheets("EXPECTATIVA - INVESTIMENTO").Cells(Expectativas_Investimentos_Linha_Análise, Expectativas_Investimentos_Coluna_Análise).Value = Sheets("EXPECTATIVA - INVESTIMENTO").Cells(Expectativas_Investimentos_Linha_Dados, Expectativas_Investimentos_Coluna_Dados1).Value - Sheets("EXPECTATIVA - INVESTIMENTO").Cells(Expectativas_Investimentos_Linha_Dados, Expectativas_Investimentos_Coluna_Dados3).Value
    'Vai para a próxima linha de dados e de análise
    Expectativas_Investimentos_Linha_Dados = Expectativas_Investimentos_Linha_Dados + 1
    Expectativas_Investimentos_Linha_Análise = Expectativas_Investimentos_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Expectativas_Investimentos_Linha_Dados = 9
Expectativas_Investimentos_Coluna_Dados1 = Sheets("EXPECTATIVA - INVESTIMENTO").Range("A9").End(xlToRight).Column 'Define o número da última coluna
Expectativas_Investimentos_Linha_Análise = 59
Expectativas_Investimentos_Coluna_Análise = 4

'Cálculo da diferença em pontos do valor mais recente em relação ao valor da média histórica
Do Until Expectativas_Investimentos_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a variável media como a média do intervalo entre a coluna com o dado mais recente e o primeiro
    media = Application.Average(Sheets("EXPECTATIVA - INVESTIMENTO").Range(Cells(Expectativas_Investimentos_Linha_Dados, Expectativas_Investimentos_Coluna_Dados1), Cells(Expectativas_Investimentos_Linha_Dados, 2)))
    'Calculo em si: o valor da celula de analise é igual ao valor mais recente menos o valor da média
    Sheets("EXPECTATIVA - INVESTIMENTO").Cells(Expectativas_Investimentos_Linha_Análise, Expectativas_Investimentos_Coluna_Análise).Value = Sheets("EXPECTATIVA - INVESTIMENTO").Cells(Expectativas_Investimentos_Linha_Dados, Expectativas_Investimentos_Coluna_Dados1).Value - media
    'Vai para a próxima linha de dados e de análise
    Expectativas_Investimentos_Linha_Dados = Expectativas_Investimentos_Linha_Dados + 1
    Expectativas_Investimentos_Linha_Análise = Expectativas_Investimentos_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Expectativas_Investimentos_Linha_Dados = 9
Expectativas_Investimentos_Coluna_Dados1 = Sheets("EXPECTATIVA - INVESTIMENTO").Range("A9").End(xlToRight).Column
Expectativas_Investimentos_Linha_Análise = 59
Expectativas_Investimentos_Coluna_Análise = 5

'Ordenação decrescente da série histórica completa
Do Until Expectativas_Investimentos_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posição = WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVA - INVESTIMENTO").Cells(Expectativas_Investimentos_Linha_Dados, Expectativas_Investimentos_Coluna_Dados1), Sheets("EXPECTATIVA - INVESTIMENTO").Range(Cells(Expectativas_Investimentos_Linha_Dados, Expectativas_Investimentos_Coluna_Dados1), Cells(Expectativas_Investimentos_Linha_Dados, 2)), 0)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("EXPECTATIVA - INVESTIMENTO").Cells(Expectativas_Investimentos_Linha_Análise, Expectativas_Investimentos_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    Expectativas_Investimentos_Linha_Dados = Expectativas_Investimentos_Linha_Dados + 1
    Expectativas_Investimentos_Linha_Análise = Expectativas_Investimentos_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Expectativas_Investimentos_Linha_Dados = 9
Expectativas_Investimentos_Coluna_Dados1 = Sheets("EXPECTATIVA - INVESTIMENTO").Range("A9").End(xlToRight).Column
Expectativas_Investimentos_Linha_Análise = 59
Expectativas_Investimentos_Coluna_Análise = 6

'Ordenação Crescente da série histórica completa
Do Until Expectativas_Investimentos_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posição = WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVA - INVESTIMENTO").Cells(Expectativas_Investimentos_Linha_Dados, Expectativas_Investimentos_Coluna_Dados1), Sheets("EXPECTATIVA - INVESTIMENTO").Range(Cells(Expectativas_Investimentos_Linha_Dados, Expectativas_Investimentos_Coluna_Dados1), Cells(Expectativas_Investimentos_Linha_Dados, 2)), 1)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("EXPECTATIVA - INVESTIMENTO").Cells(Expectativas_Investimentos_Linha_Análise, Expectativas_Investimentos_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    Expectativas_Investimentos_Linha_Dados = Expectativas_Investimentos_Linha_Dados + 1
    Expectativas_Investimentos_Linha_Análise = Expectativas_Investimentos_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior e define a variável Coluna_DadosP que representa a primeira coluna de dados
Expectativas_Investimentos_Coluna_Dados1 = Sheets("EXPECTATIVA - INVESTIMENTO").Range("A9").End(xlToRight).Column
Expectativas_Investimentos_Coluna_DadosP = 2

Do Until Expectativas_Investimentos_Coluna_DadosP = Expectativas_Investimentos_Coluna_Dados1 + 1 ' Faz até a variavel Coluna_DadosP ser igual a variavel Coluna_Dados1 mais uma unidade
    'Confere se o mês da coluna em questão é igual ao mês do dado mais recente
    If Month(Sheets("EXPECTATIVA - INVESTIMENTO").Cells(8, Expectativas_Investimentos_Coluna_DadosP)) = Month(Sheets("EXPECTATIVA - INVESTIMENTO").Cells(8, Expectativas_Investimentos_Coluna_Dados1)) Then
    'Caso seja igual, copia a coluna com os dados mais abaixo, a partir da linha 110
        Sheets("EXPECTATIVA - INVESTIMENTO").Range(Cells(9, Expectativas_Investimentos_Coluna_DadosP), (Cells(54, Expectativas_Investimentos_Coluna_DadosP))).Copy (Sheets("EXPECTATIVA - INVESTIMENTO").Cells(110, Expectativas_Investimentos_Coluna_DadosP))
    End If
    'Vai para a próxima coluna
    Expectativas_Investimentos_Coluna_DadosP = Expectativas_Investimentos_Coluna_DadosP + 1
'Repete a conferencia com a próxima coluna
Loop


'Atribui os valores originais das variaveis após o loop anterior
Expectativas_Investimentos_Linha_Dados = 110
Expectativas_Investimentos_Coluna_Dados1 = Sheets("EXPECTATIVA - INVESTIMENTO").Range("A9").End(xlToRight).Column 'Define o número da última coluna
Expectativas_Investimentos_Linha_Análise = 59
Expectativas_Investimentos_Coluna_Análise = 7

'Ordenação decrescente da série histórica dos meses do dado mais recente
Do Until Expectativas_Investimentos_Linha_Dados = 156 'Faz o calculo até a variável Linha_Dados ser 156
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo mês do mais recente
    posição = Application.WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVA - INVESTIMENTO").Cells(Expectativas_Investimentos_Linha_Dados, Expectativas_Investimentos_Coluna_Dados1), Sheets("EXPECTATIVA - INVESTIMENTO").Range(Cells(Expectativas_Investimentos_Linha_Dados, Expectativas_Investimentos_Coluna_Dados1), Cells(Expectativas_Investimentos_Linha_Dados, 2)))
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("EXPECTATIVA - INVESTIMENTO").Cells(Expectativas_Investimentos_Linha_Análise, Expectativas_Investimentos_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    Expectativas_Investimentos_Linha_Dados = Expectativas_Investimentos_Linha_Dados + 1
    Expectativas_Investimentos_Linha_Análise = Expectativas_Investimentos_Linha_Análise + 1
'Repete a ordenação com a próxima linha
Loop

'Atribui os valores originais das variaveis após o loop anterior
Expectativas_Investimentos_Linha_Dados = 110
Expectativas_Investimentos_Coluna_Dados1 = Sheets("EXPECTATIVA - INVESTIMENTO").Range("A9").End(xlToRight).Column 'Define o número da última coluna
Expectativas_Investimentos_Linha_Análise = 59
Expectativas_Investimentos_Coluna_Análise = 8
'Ordenação crescente da série histórica dos meses do dado mais recente
Do Until Expectativas_Investimentos_Linha_Dados = 156 'Faz o calculo até a variável Linha_Dados ser 156
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo mês do mais recente
    posição = Application.WorksheetFunction.Rank_Eq(Sheets("EXPECTATIVA - INVESTIMENTO").Cells(Expectativas_Investimentos_Linha_Dados, Expectativas_Investimentos_Coluna_Dados1), Sheets("EXPECTATIVA - INVESTIMENTO").Range(Cells(Expectativas_Investimentos_Linha_Dados, Expectativas_Investimentos_Coluna_Dados1), Cells(Expectativas_Investimentos_Linha_Dados, 2)), 1)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("EXPECTATIVA - INVESTIMENTO").Cells(Expectativas_Investimentos_Linha_Análise, Expectativas_Investimentos_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    Expectativas_Investimentos_Linha_Dados = Expectativas_Investimentos_Linha_Dados + 1
    Expectativas_Investimentos_Linha_Análise = Expectativas_Investimentos_Linha_Análise + 1
'Repete a ordenação com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
Expectativas_Investimentos_Linha_Dados = 9
Expectativas_Investimentos_Coluna_Dados1 = Sheets("EXPECTATIVA - INVESTIMENTO").Range("A9").End(xlToRight).Column
Expectativas_Investimentos_Coluna_Dados2 = Expectativas_Investimentos_Coluna_Dados1 - 1
Expectativas_Investimentos_Linha_Análise = 59
Expectativas_Investimentos_Coluna_Análise = 9

'Avaliação se cruzou ou não a linha de 50 e o sentido
Do Until Expectativas_Investimentos_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'se o dado mais recente for menor que 50 e o dado do mês anterior for maior ou igual a 50 então...
    If Cells(Expectativas_Investimentos_Linha_Dados, Expectativas_Investimentos_Coluna_Dados1) < 50 And Cells(Expectativas_Investimentos_Linha_Dados, Expectativas_Investimentos_Coluna_Dados2) >= 50 Then
    
    'a célula de análise recebe cruzou para baixo
    Cells(Expectativas_Investimentos_Linha_Análise, Expectativas_Investimentos_Coluna_Análise).Value = "Cruzou para baixo"
    'Caso não seja..
    Else
        'se o dado mais recente for maior ou igual a 50 e o dado do m~es anterior for menor ou igual a 50 então...
        If Cells(Expectativas_Investimentos_Linha_Dados, Expectativas_Investimentos_Coluna_Dados1) >= 50 And Cells(Expectativas_Investimentos_Linha_Dados, Expectativas_Investimentos_Coluna_Dados2) <= 50 Then
        'a célula de análise recebe cruzou para cima
        Cells(Expectativas_Investimentos_Linha_Análise, Expectativas_Investimentos_Coluna_Análise).Value = "Cruzou para cima"
        'Caso não seja..
        Else
        'a célula de análise recebe não cruzou
        Cells(Expectativas_Investimentos_Linha_Análise, Expectativas_Investimentos_Coluna_Análise).Value = "Não Cruzou"
        End If
    End If
    'Vai para a próxima linha de dados e de análise
    Expectativas_Investimentos_Linha_Dados = Expectativas_Investimentos_Linha_Dados + 1
    Expectativas_Investimentos_Linha_Análise = Expectativas_Investimentos_Linha_Análise + 1
'repete o processo com a nova linha
Loop

'Apaga as linhas com erros/dados faltantes/títulos e subtítulos
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

Sub Análise_Verde()


Dim SituaçãoFinanceira_Lucro_Linha_Dados As Integer 'Define a linha que contém o dado a ser usado
Dim SituaçãoFinanceira_Lucro_Coluna_Dados1 As Integer ' Define a coluna com o dado mais recente
Dim SituaçãoFinanceira_Lucro_Coluna_Dados2 As Integer ' Define a coluna com o dado do mês anterior
Dim SituaçãoFinanceira_Lucro_Coluna_Dados3 As Integer ' Defie a coluna com o dado do mesmo mês do ano anterior
Dim SituaçãoFinanceira_Lucro_Linha_Análise As Integer ' Define a linha que será feita a análise
Dim SituaçãoFinanceira_Lucro_Coluna_Análise As Integer 'Define a coluna que será feita a análise

Sheets("SITUACAO FINANCEIRA LUCRO").Select


'Copia os títulos das categorias e cola onde será formada a tabela de análise
Sheets("SITUACAO FINANCEIRA LUCRO").Range(Cells(9, 1), Cells(54, 9)).Copy (Sheets("SITUACAO FINANCEIRA LUCRO").Cells(59, 1))
'Limpa os números que foram colados mas mantém a formatação
Sheets("SITUACAO FINANCEIRA LUCRO").Range(Cells(59, 2), Cells(105, 9)).ClearContents



'Nomeia as colunas de acordo com o dado que será calculado nelas
Sheets("SITUACAO FINANCEIRA LUCRO").Cells(58, 2).Value = "Diferença para o trimestre anterior"
Sheets("SITUACAO FINANCEIRA LUCRO").Cells(58, 3).Value = "Diferença para ao mesmo trimestre do ano anterior"
Sheets("SITUACAO FINANCEIRA LUCRO").Cells(58, 4).Value = "Diferença para a média histórica"
Sheets("SITUACAO FINANCEIRA LUCRO").Cells(58, 5).Value = "Posição Decrescente (Maior valor 1º, menor valor último)"
Sheets("SITUACAO FINANCEIRA LUCRO").Cells(58, 6).Value = "Posição Crescente  (Menor valor 1º, maior valor último)"
Sheets("SITUACAO FINANCEIRA LUCRO").Cells(58, 7).Value = "Posição Crescente - Mesmo trimestre  (Menor valor 1º, maior valor último)"
Sheets("SITUACAO FINANCEIRA LUCRO").Cells(58, 8).Value = "Posição Decrescente -Mesmo trimestre  (Maior valor 1º, menor valor último)"
Sheets("SITUACAO FINANCEIRA LUCRO").Cells(58, 9).Value = "cruzou a linha de 50?"

'Atribui valores as variaveis definidas acima
SituaçãoFinanceira_Lucro_Linha_Dados = 9 'Define o número da primeira linha de dados
SituaçãoFinanceira_Lucro_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA LUCRO").Range("A9").End(xlToRight).Column 'Define o número da última coluna
SituaçãoFinanceira_Lucro_Coluna_Dados2 = SituaçãoFinanceira_Lucro_Coluna_Dados1 - 1 'Define o número da coluna do mês anterior
SituaçãoFinanceira_Lucro_Coluna_Dados3 = SituaçãoFinanceira_Lucro_Coluna_Dados1 - 4
SituaçãoFinanceira_Lucro_Linha_Análise = 59 'Define a primeira linhas de análises
SituaçãoFinanceira_Lucro_Coluna_Análise = 2 'Define a coluna de análises

'Inserindo valores nas celulas vazias para fugir de bugs
Sheets("SITUACAO FINANCEIRA LUCRO").Range(Cells(10, SituaçãoFinanceira_Lucro_Coluna_Dados3), Cells(10, SituaçãoFinanceira_Lucro_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA LUCRO").Range(Cells(16, SituaçãoFinanceira_Lucro_Coluna_Dados3), Cells(16, SituaçãoFinanceira_Lucro_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA LUCRO").Range(Cells(20, SituaçãoFinanceira_Lucro_Coluna_Dados3), Cells(20, SituaçãoFinanceira_Lucro_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA LUCRO").Range(Cells(22, SituaçãoFinanceira_Lucro_Coluna_Dados3), Cells(23, SituaçãoFinanceira_Lucro_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA LUCRO").Range(Cells(25, SituaçãoFinanceira_Lucro_Coluna_Dados3), Cells(25, SituaçãoFinanceira_Lucro_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA LUCRO").Range(Cells(29, SituaçãoFinanceira_Lucro_Coluna_Dados3), Cells(29, SituaçãoFinanceira_Lucro_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA LUCRO").Range(Cells(37, SituaçãoFinanceira_Lucro_Coluna_Dados3), Cells(37, SituaçãoFinanceira_Lucro_Coluna_Dados1)).Value = "0"


'Calculo da difernça em pontos do valor mais recente em relação ao valor do mês anterior
Do Until SituaçãoFinanceira_Lucro_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Calculo da diferença em si: o valor da celula de analise é igual ao valor mais recente menos o valor do mês anterior
   Sheets("SITUACAO FINANCEIRA LUCRO").Cells(SituaçãoFinanceira_Lucro_Linha_Análise, SituaçãoFinanceira_Lucro_Coluna_Análise).Value = Sheets("SITUACAO FINANCEIRA LUCRO").Cells(SituaçãoFinanceira_Lucro_Linha_Dados, SituaçãoFinanceira_Lucro_Coluna_Dados1).Value - Sheets("SITUACAO FINANCEIRA LUCRO").Cells(SituaçãoFinanceira_Lucro_Linha_Dados, SituaçãoFinanceira_Lucro_Coluna_Dados2).Value
    'Vai para a próxima linha de dados e de análise
   SituaçãoFinanceira_Lucro_Linha_Dados = SituaçãoFinanceira_Lucro_Linha_Dados + 1
   SituaçãoFinanceira_Lucro_Linha_Análise = SituaçãoFinanceira_Lucro_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
SituaçãoFinanceira_Lucro_Linha_Dados = 9
SituaçãoFinanceira_Lucro_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA LUCRO").Range("A9").End(xlToRight).Column
SituaçãoFinanceira_Lucro_Coluna_Dados3 = SituaçãoFinanceira_Lucro_Coluna_Dados1 - 4
SituaçãoFinanceira_Lucro_Linha_Análise = 59
SituaçãoFinanceira_Lucro_Coluna_Análise = 3

'Cálculo da diferença em pontos do valor mais recente em relação ao valor do mesmo mês do ano anterior
Do Until SituaçãoFinanceira_Lucro_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Calculo da diferença em si: o valor da celula de analise é igual ao valor mais recente menos o valor do mesmo mês do ano anterior
    Sheets("SITUACAO FINANCEIRA LUCRO").Cells(SituaçãoFinanceira_Lucro_Linha_Análise, SituaçãoFinanceira_Lucro_Coluna_Análise).Value = Sheets("SITUACAO FINANCEIRA LUCRO").Cells(SituaçãoFinanceira_Lucro_Linha_Dados, SituaçãoFinanceira_Lucro_Coluna_Dados1).Value - Sheets("SITUACAO FINANCEIRA LUCRO").Cells(SituaçãoFinanceira_Lucro_Linha_Dados, SituaçãoFinanceira_Lucro_Coluna_Dados3).Value
    'Vai para a próxima linha de dados e de análise
    SituaçãoFinanceira_Lucro_Linha_Dados = SituaçãoFinanceira_Lucro_Linha_Dados + 1
    SituaçãoFinanceira_Lucro_Linha_Análise = SituaçãoFinanceira_Lucro_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
SituaçãoFinanceira_Lucro_Linha_Dados = 9
SituaçãoFinanceira_Lucro_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA LUCRO").Range("A9").End(xlToRight).Column 'Define o número da última coluna
SituaçãoFinanceira_Lucro_Linha_Análise = 59
SituaçãoFinanceira_Lucro_Coluna_Análise = 4

'Cálculo da diferença em pontos do valor mais recente em relação ao valor da média histórica
Do Until SituaçãoFinanceira_Lucro_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a variável media como a média do intervalo entre a coluna com o dado mais recente e o primeiro
    media = Application.Average(Sheets("SITUACAO FINANCEIRA LUCRO").Range(Cells(SituaçãoFinanceira_Lucro_Linha_Dados, SituaçãoFinanceira_Lucro_Coluna_Dados1), Cells(SituaçãoFinanceira_Lucro_Linha_Dados, 2)))
    'Calculo em si: o valor da celula de analise é igual ao valor mais recente menos o valor da média
    Sheets("SITUACAO FINANCEIRA LUCRO").Cells(SituaçãoFinanceira_Lucro_Linha_Análise, SituaçãoFinanceira_Lucro_Coluna_Análise).Value = Sheets("SITUACAO FINANCEIRA LUCRO").Cells(SituaçãoFinanceira_Lucro_Linha_Dados, SituaçãoFinanceira_Lucro_Coluna_Dados1).Value - media
    'Vai para a próxima linha de dados e de análise
    SituaçãoFinanceira_Lucro_Linha_Dados = SituaçãoFinanceira_Lucro_Linha_Dados + 1
    SituaçãoFinanceira_Lucro_Linha_Análise = SituaçãoFinanceira_Lucro_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
SituaçãoFinanceira_Lucro_Linha_Dados = 9
SituaçãoFinanceira_Lucro_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA LUCRO").Range("A9").End(xlToRight).Column
SituaçãoFinanceira_Lucro_Linha_Análise = 59
SituaçãoFinanceira_Lucro_Coluna_Análise = 5

'Ordenação decrescente da série histórica completa
Do Until SituaçãoFinanceira_Lucro_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posição = WorksheetFunction.Rank_Eq(Sheets("SITUACAO FINANCEIRA LUCRO").Cells(SituaçãoFinanceira_Lucro_Linha_Dados, SituaçãoFinanceira_Lucro_Coluna_Dados1), Sheets("SITUACAO FINANCEIRA LUCRO").Range(Cells(SituaçãoFinanceira_Lucro_Linha_Dados, SituaçãoFinanceira_Lucro_Coluna_Dados1), Cells(SituaçãoFinanceira_Lucro_Linha_Dados, 2)), 0)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("SITUACAO FINANCEIRA LUCRO").Cells(SituaçãoFinanceira_Lucro_Linha_Análise, SituaçãoFinanceira_Lucro_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    SituaçãoFinanceira_Lucro_Linha_Dados = SituaçãoFinanceira_Lucro_Linha_Dados + 1
    SituaçãoFinanceira_Lucro_Linha_Análise = SituaçãoFinanceira_Lucro_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
SituaçãoFinanceira_Lucro_Linha_Dados = 9
SituaçãoFinanceira_Lucro_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA LUCRO").Range("A9").End(xlToRight).Column
SituaçãoFinanceira_Lucro_Linha_Análise = 59
SituaçãoFinanceira_Lucro_Coluna_Análise = 6

'Ordenação Crescente da série histórica completa
Do Until SituaçãoFinanceira_Lucro_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posição = WorksheetFunction.Rank_Eq(Sheets("SITUACAO FINANCEIRA LUCRO").Cells(SituaçãoFinanceira_Lucro_Linha_Dados, SituaçãoFinanceira_Lucro_Coluna_Dados1), Sheets("SITUACAO FINANCEIRA LUCRO").Range(Cells(SituaçãoFinanceira_Lucro_Linha_Dados, SituaçãoFinanceira_Lucro_Coluna_Dados1), Cells(SituaçãoFinanceira_Lucro_Linha_Dados, 2)), 1)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("SITUACAO FINANCEIRA LUCRO").Cells(SituaçãoFinanceira_Lucro_Linha_Análise, SituaçãoFinanceira_Lucro_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    SituaçãoFinanceira_Lucro_Linha_Dados = SituaçãoFinanceira_Lucro_Linha_Dados + 1
    SituaçãoFinanceira_Lucro_Linha_Análise = SituaçãoFinanceira_Lucro_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior e define a variável Coluna_DadosP que representa a primeira coluna de dados
SituaçãoFinanceira_Lucro_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA LUCRO").Range("A9").End(xlToRight).Column
SituaçãoFinanceira_Lucro_Coluna_DadosP = 2

Do Until SituaçãoFinanceira_Lucro_Coluna_DadosP = SituaçãoFinanceira_Lucro_Coluna_Dados1 + 1 ' Faz até a variavel Coluna_DadosP ser igual a variavel Coluna_Dados1 mais uma unidade
    'Confere se o mês da coluna em questão é igual ao mês do dado mais recente
    If Left(Sheets("SITUACAO FINANCEIRA LUCRO").Cells(8, SituaçãoFinanceira_Lucro_Coluna_DadosP), 1) = Left(Sheets("SITUACAO FINANCEIRA LUCRO").Cells(8, SituaçãoFinanceira_Lucro_Coluna_Dados1), 1) Then
    'Caso seja igual, copia a coluna com os dados mais abaixo, a partir da linha 110
        Sheets("SITUACAO FINANCEIRA LUCRO").Range(Cells(9, SituaçãoFinanceira_Lucro_Coluna_DadosP), (Cells(54, SituaçãoFinanceira_Lucro_Coluna_DadosP))).Copy (Sheets("SITUACAO FINANCEIRA LUCRO").Cells(110, SituaçãoFinanceira_Lucro_Coluna_DadosP))
    End If
    'Vai para a próxima coluna
    SituaçãoFinanceira_Lucro_Coluna_DadosP = SituaçãoFinanceira_Lucro_Coluna_DadosP + 1
'Repete a conferencia com a próxima coluna
Loop


'Atribui os valores originais das variaveis após o loop anterior
SituaçãoFinanceira_Lucro_Linha_Dados = 110
SituaçãoFinanceira_Lucro_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA LUCRO").Range("A9").End(xlToRight).Column 'Define o número da última coluna
SituaçãoFinanceira_Lucro_Linha_Análise = 59
SituaçãoFinanceira_Lucro_Coluna_Análise = 7

'Ordenação decrescente da série histórica dos meses do dado mais recente
Do Until SituaçãoFinanceira_Lucro_Linha_Dados = 156 'Faz o calculo até a variável Linha_Dados ser 156
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo mês do mais recente
    posição = Application.WorksheetFunction.Rank_Eq(Sheets("SITUACAO FINANCEIRA LUCRO").Cells(SituaçãoFinanceira_Lucro_Linha_Dados, SituaçãoFinanceira_Lucro_Coluna_Dados1), Sheets("SITUACAO FINANCEIRA LUCRO").Range(Cells(SituaçãoFinanceira_Lucro_Linha_Dados, SituaçãoFinanceira_Lucro_Coluna_Dados1), Cells(SituaçãoFinanceira_Lucro_Linha_Dados, 2)))
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("SITUACAO FINANCEIRA LUCRO").Cells(SituaçãoFinanceira_Lucro_Linha_Análise, SituaçãoFinanceira_Lucro_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    SituaçãoFinanceira_Lucro_Linha_Dados = SituaçãoFinanceira_Lucro_Linha_Dados + 1
    SituaçãoFinanceira_Lucro_Linha_Análise = SituaçãoFinanceira_Lucro_Linha_Análise + 1
'Repete a ordenação com a próxima linha
Loop

'Atribui os valores originais das variaveis após o loop anterior
SituaçãoFinanceira_Lucro_Linha_Dados = 110
SituaçãoFinanceira_Lucro_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA LUCRO").Range("A9").End(xlToRight).Column 'Define o número da última coluna
SituaçãoFinanceira_Lucro_Linha_Análise = 59
SituaçãoFinanceira_Lucro_Coluna_Análise = 8
'Ordenação crescente da série histórica dos meses do dado mais recente
Do Until SituaçãoFinanceira_Lucro_Linha_Dados = 156 'Faz o calculo até a variável Linha_Dados ser 156
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo mês do mais recente
    posição = Application.WorksheetFunction.Rank_Eq(Sheets("SITUACAO FINANCEIRA LUCRO").Cells(SituaçãoFinanceira_Lucro_Linha_Dados, SituaçãoFinanceira_Lucro_Coluna_Dados1), Sheets("SITUACAO FINANCEIRA LUCRO").Range(Cells(SituaçãoFinanceira_Lucro_Linha_Dados, SituaçãoFinanceira_Lucro_Coluna_Dados1), Cells(SituaçãoFinanceira_Lucro_Linha_Dados, 2)), 1)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("SITUACAO FINANCEIRA LUCRO").Cells(SituaçãoFinanceira_Lucro_Linha_Análise, SituaçãoFinanceira_Lucro_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    SituaçãoFinanceira_Lucro_Linha_Dados = SituaçãoFinanceira_Lucro_Linha_Dados + 1
    SituaçãoFinanceira_Lucro_Linha_Análise = SituaçãoFinanceira_Lucro_Linha_Análise + 1
'Repete a ordenação com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
SituaçãoFinanceira_Lucro_Linha_Dados = 9
SituaçãoFinanceira_Lucro_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA LUCRO").Range("A9").End(xlToRight).Column
SituaçãoFinanceira_Lucro_Coluna_Dados2 = SituaçãoFinanceira_Lucro_Coluna_Dados1 - 1
SituaçãoFinanceira_Lucro_Linha_Análise = 59
SituaçãoFinanceira_Lucro_Coluna_Análise = 9

'Avaliação se cruzou ou não a linha de 50 e o sentido
Do Until SituaçãoFinanceira_Lucro_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'se o dado mais recente for menor que 50 e o dado do mês anterior for maior ou igual a 50 então...
    If Cells(SituaçãoFinanceira_Lucro_Linha_Dados, SituaçãoFinanceira_Lucro_Coluna_Dados1) < 50 And Cells(SituaçãoFinanceira_Lucro_Linha_Dados, SituaçãoFinanceira_Lucro_Coluna_Dados2) >= 50 Then
    
    'a célula de análise recebe cruzou para baixo
    Cells(SituaçãoFinanceira_Lucro_Linha_Análise, SituaçãoFinanceira_Lucro_Coluna_Análise).Value = "Cruzou para baixo"
    'Caso não seja..
    Else
        'se o dado mais recente for maior ou igual a 50 e o dado do m~es anterior for menor ou igual a 50 então...
        If Cells(SituaçãoFinanceira_Lucro_Linha_Dados, SituaçãoFinanceira_Lucro_Coluna_Dados1) >= 50 And Cells(SituaçãoFinanceira_Lucro_Linha_Dados, SituaçãoFinanceira_Lucro_Coluna_Dados2) <= 50 Then
        'a célula de análise recebe cruzou para cima
        Cells(SituaçãoFinanceira_Lucro_Linha_Análise, SituaçãoFinanceira_Lucro_Coluna_Análise).Value = "Cruzou para cima"
        'Caso não seja..
        Else
        'a célula de análise recebe não cruzou
        Cells(SituaçãoFinanceira_Lucro_Linha_Análise, SituaçãoFinanceira_Lucro_Coluna_Análise).Value = "Não Cruzou"
        End If
    End If
    'Vai para a próxima linha de dados e de análise
    SituaçãoFinanceira_Lucro_Linha_Dados = SituaçãoFinanceira_Lucro_Linha_Dados + 1
    SituaçãoFinanceira_Lucro_Linha_Análise = SituaçãoFinanceira_Lucro_Linha_Análise + 1
'repete o processo com a nova linha
Loop

'Apaga as linhas com erros/dados faltantes/títulos e subtítulos
Range(Cells(60, 2), Cells(60, 9)).ClearContents
Range(Cells(66, 2), Cells(66, 9)).ClearContents
Range(Cells(70, 2), Cells(70, 9)).ClearContents
Range(Cells(72, 2), Cells(73, 9)).Value = "-"
Range(Cells(75, 2), Cells(75, 9)).Value = "-"
Range(Cells(79, 2), Cells(79, 9)).Value = "-"
Range(Cells(87, 2), Cells(87, 9)).Value = "-"

'Inserindo valores nas celulas vazias para fugir de bugs
Range(Cells(10, SituaçãoFinanceira_Lucro_Coluna_Dados3), Cells(10, SituaçãoFinanceira_Lucro_Coluna_Dados1)).ClearContents
Range(Cells(16, SituaçãoFinanceira_Lucro_Coluna_Dados3), Cells(16, SituaçãoFinanceira_Lucro_Coluna_Dados1)).ClearContents
Range(Cells(20, SituaçãoFinanceira_Lucro_Coluna_Dados3), Cells(20, SituaçãoFinanceira_Lucro_Coluna_Dados1)).ClearContents
Range(Cells(22, SituaçãoFinanceira_Lucro_Coluna_Dados3), Cells(23, SituaçãoFinanceira_Lucro_Coluna_Dados1)).Value = "-"
Range(Cells(25, SituaçãoFinanceira_Lucro_Coluna_Dados3), Cells(25, SituaçãoFinanceira_Lucro_Coluna_Dados1)).Value = "-"
Range(Cells(29, SituaçãoFinanceira_Lucro_Coluna_Dados3), Cells(29, SituaçãoFinanceira_Lucro_Coluna_Dados1)).Value = "-"
Range(Cells(37, SituaçãoFinanceira_Lucro_Coluna_Dados3), Cells(37, SituaçãoFinanceira_Lucro_Coluna_Dados1)).Value = "-"

Range("E59:H104").NumberFormat = "0"




'**********************************         SituaçãoFinanceira_PreçoMédio              **********************************************




Dim SituaçãoFinanceira_PreçoMédio_Linha_Dados As Integer 'Define a linha que contém o dado a ser usado
Dim SituaçãoFinanceira_PreçoMédio_Coluna_Dados1 As Integer ' Define a coluna com o dado mais recente
Dim SituaçãoFinanceira_PreçoMédio_Coluna_Dados2 As Integer ' Define a coluna com o dado do mês anterior
Dim SituaçãoFinanceira_PreçoMédio_Coluna_Dados3 As Integer ' Defie a coluna com o dado do mesmo mês do ano anterior
Dim SituaçãoFinanceira_PreçoMédio_Linha_Análise As Integer ' Define a linha que será feita a análise
Dim SituaçãoFinanceira_PreçoMédio_Coluna_Análise As Integer 'Define a coluna que será feita a análise

Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Select


'Copia os títulos das categorias e cola onde será formada a tabela de análise
Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Range(Cells(9, 1), Cells(54, 9)).Copy (Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(59, 1))
'Limpa os números que foram colados mas mantém a formatação
Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Range(Cells(59, 2), Cells(105, 9)).ClearContents



'Nomeia as colunas de acordo com o dado que será calculado nelas
Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(58, 2).Value = "Diferença para o trimestre anterior"
Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(58, 3).Value = "Diferença para ao mesmo trimestre do ano anterior"
Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(58, 4).Value = "Diferença para a média histórica"
Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(58, 5).Value = "Posição Decrescente (Maior valor 1º, menor valor último)"
Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(58, 6).Value = "Posição Crescente  (Menor valor 1º, maior valor último)"
Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(58, 7).Value = "Posição Crescente - Mesmo trimestre  (Menor valor 1º, maior valor último)"
Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(58, 8).Value = "Posição Decrescente -Mesmo trimestre  (Maior valor 1º, menor valor último)"
Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(58, 9).Value = "cruzou a linha de 50?"

'Atribui valores as variaveis definidas acima
SituaçãoFinanceira_PreçoMédio_Linha_Dados = 9 'Define o número da primeira linha de dados
SituaçãoFinanceira_PreçoMédio_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Range("A9").End(xlToRight).Column 'Define o número da última coluna
SituaçãoFinanceira_PreçoMédio_Coluna_Dados2 = SituaçãoFinanceira_PreçoMédio_Coluna_Dados1 - 1 'Define o número da coluna do mês anterior
SituaçãoFinanceira_PreçoMédio_Coluna_Dados3 = SituaçãoFinanceira_PreçoMédio_Coluna_Dados1 - 4
SituaçãoFinanceira_PreçoMédio_Linha_Análise = 59 'Define a primeira linhas de análises
SituaçãoFinanceira_PreçoMédio_Coluna_Análise = 2 'Define a coluna de análises

'Inserindo valores nas celulas vazias para fugir de bugs
Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Range(Cells(10, SituaçãoFinanceira_PreçoMédio_Coluna_Dados3), Cells(10, SituaçãoFinanceira_PreçoMédio_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Range(Cells(16, SituaçãoFinanceira_PreçoMédio_Coluna_Dados3), Cells(16, SituaçãoFinanceira_PreçoMédio_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Range(Cells(20, SituaçãoFinanceira_PreçoMédio_Coluna_Dados3), Cells(20, SituaçãoFinanceira_PreçoMédio_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Range(Cells(22, SituaçãoFinanceira_PreçoMédio_Coluna_Dados3), Cells(23, SituaçãoFinanceira_PreçoMédio_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Range(Cells(25, SituaçãoFinanceira_PreçoMédio_Coluna_Dados3), Cells(25, SituaçãoFinanceira_PreçoMédio_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Range(Cells(29, SituaçãoFinanceira_PreçoMédio_Coluna_Dados3), Cells(29, SituaçãoFinanceira_PreçoMédio_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Range(Cells(37, SituaçãoFinanceira_PreçoMédio_Coluna_Dados3), Cells(37, SituaçãoFinanceira_PreçoMédio_Coluna_Dados1)).Value = "0"


'Calculo da difernça em pontos do valor mais recente em relação ao valor do mês anterior
Do Until SituaçãoFinanceira_PreçoMédio_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Calculo da diferença em si: o valor da celula de analise é igual ao valor mais recente menos o valor do mês anterior
   Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(SituaçãoFinanceira_PreçoMédio_Linha_Análise, SituaçãoFinanceira_PreçoMédio_Coluna_Análise).Value = Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(SituaçãoFinanceira_PreçoMédio_Linha_Dados, SituaçãoFinanceira_PreçoMédio_Coluna_Dados1).Value - Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(SituaçãoFinanceira_PreçoMédio_Linha_Dados, SituaçãoFinanceira_PreçoMédio_Coluna_Dados2).Value
    'Vai para a próxima linha de dados e de análise
   SituaçãoFinanceira_PreçoMédio_Linha_Dados = SituaçãoFinanceira_PreçoMédio_Linha_Dados + 1
   SituaçãoFinanceira_PreçoMédio_Linha_Análise = SituaçãoFinanceira_PreçoMédio_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
SituaçãoFinanceira_PreçoMédio_Linha_Dados = 9
SituaçãoFinanceira_PreçoMédio_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Range("A9").End(xlToRight).Column
SituaçãoFinanceira_PreçoMédio_Coluna_Dados3 = SituaçãoFinanceira_PreçoMédio_Coluna_Dados1 - 4
SituaçãoFinanceira_PreçoMédio_Linha_Análise = 59
SituaçãoFinanceira_PreçoMédio_Coluna_Análise = 3

'Cálculo da diferença em pontos do valor mais recente em relação ao valor do mesmo mês do ano anterior
Do Until SituaçãoFinanceira_PreçoMédio_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Calculo da diferença em si: o valor da celula de analise é igual ao valor mais recente menos o valor do mesmo mês do ano anterior
    Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(SituaçãoFinanceira_PreçoMédio_Linha_Análise, SituaçãoFinanceira_PreçoMédio_Coluna_Análise).Value = Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(SituaçãoFinanceira_PreçoMédio_Linha_Dados, SituaçãoFinanceira_PreçoMédio_Coluna_Dados1).Value - Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(SituaçãoFinanceira_PreçoMédio_Linha_Dados, SituaçãoFinanceira_PreçoMédio_Coluna_Dados3).Value
    'Vai para a próxima linha de dados e de análise
    SituaçãoFinanceira_PreçoMédio_Linha_Dados = SituaçãoFinanceira_PreçoMédio_Linha_Dados + 1
    SituaçãoFinanceira_PreçoMédio_Linha_Análise = SituaçãoFinanceira_PreçoMédio_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
SituaçãoFinanceira_PreçoMédio_Linha_Dados = 9
SituaçãoFinanceira_PreçoMédio_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Range("A9").End(xlToRight).Column 'Define o número da última coluna
SituaçãoFinanceira_PreçoMédio_Linha_Análise = 59
SituaçãoFinanceira_PreçoMédio_Coluna_Análise = 4

'Cálculo da diferença em pontos do valor mais recente em relação ao valor da média histórica
Do Until SituaçãoFinanceira_PreçoMédio_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a variável media como a média do intervalo entre a coluna com o dado mais recente e o primeiro
    media = Application.Average(Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Range(Cells(SituaçãoFinanceira_PreçoMédio_Linha_Dados, SituaçãoFinanceira_PreçoMédio_Coluna_Dados1), Cells(SituaçãoFinanceira_PreçoMédio_Linha_Dados, 2)))
    'Calculo em si: o valor da celula de analise é igual ao valor mais recente menos o valor da média
    Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(SituaçãoFinanceira_PreçoMédio_Linha_Análise, SituaçãoFinanceira_PreçoMédio_Coluna_Análise).Value = Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(SituaçãoFinanceira_PreçoMédio_Linha_Dados, SituaçãoFinanceira_PreçoMédio_Coluna_Dados1).Value - media
    'Vai para a próxima linha de dados e de análise
    SituaçãoFinanceira_PreçoMédio_Linha_Dados = SituaçãoFinanceira_PreçoMédio_Linha_Dados + 1
    SituaçãoFinanceira_PreçoMédio_Linha_Análise = SituaçãoFinanceira_PreçoMédio_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
SituaçãoFinanceira_PreçoMédio_Linha_Dados = 9
SituaçãoFinanceira_PreçoMédio_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Range("A9").End(xlToRight).Column
SituaçãoFinanceira_PreçoMédio_Linha_Análise = 59
SituaçãoFinanceira_PreçoMédio_Coluna_Análise = 5

'Ordenação decrescente da série histórica completa
Do Until SituaçãoFinanceira_PreçoMédio_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posição = WorksheetFunction.Rank_Eq(Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(SituaçãoFinanceira_PreçoMédio_Linha_Dados, SituaçãoFinanceira_PreçoMédio_Coluna_Dados1), Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Range(Cells(SituaçãoFinanceira_PreçoMédio_Linha_Dados, SituaçãoFinanceira_PreçoMédio_Coluna_Dados1), Cells(SituaçãoFinanceira_PreçoMédio_Linha_Dados, 2)), 0)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(SituaçãoFinanceira_PreçoMédio_Linha_Análise, SituaçãoFinanceira_PreçoMédio_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    SituaçãoFinanceira_PreçoMédio_Linha_Dados = SituaçãoFinanceira_PreçoMédio_Linha_Dados + 1
    SituaçãoFinanceira_PreçoMédio_Linha_Análise = SituaçãoFinanceira_PreçoMédio_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
SituaçãoFinanceira_PreçoMédio_Linha_Dados = 9
SituaçãoFinanceira_PreçoMédio_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Range("A9").End(xlToRight).Column
SituaçãoFinanceira_PreçoMédio_Linha_Análise = 59
SituaçãoFinanceira_PreçoMédio_Coluna_Análise = 6

'Ordenação Crescente da série histórica completa
Do Until SituaçãoFinanceira_PreçoMédio_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posição = WorksheetFunction.Rank_Eq(Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(SituaçãoFinanceira_PreçoMédio_Linha_Dados, SituaçãoFinanceira_PreçoMédio_Coluna_Dados1), Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Range(Cells(SituaçãoFinanceira_PreçoMédio_Linha_Dados, SituaçãoFinanceira_PreçoMédio_Coluna_Dados1), Cells(SituaçãoFinanceira_PreçoMédio_Linha_Dados, 2)), 1)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(SituaçãoFinanceira_PreçoMédio_Linha_Análise, SituaçãoFinanceira_PreçoMédio_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    SituaçãoFinanceira_PreçoMédio_Linha_Dados = SituaçãoFinanceira_PreçoMédio_Linha_Dados + 1
    SituaçãoFinanceira_PreçoMédio_Linha_Análise = SituaçãoFinanceira_PreçoMédio_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior e define a variável Coluna_DadosP que representa a primeira coluna de dados
SituaçãoFinanceira_PreçoMédio_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Range("A9").End(xlToRight).Column
SituaçãoFinanceira_PreçoMédio_Coluna_DadosP = 2

Do Until SituaçãoFinanceira_PreçoMédio_Coluna_DadosP = SituaçãoFinanceira_PreçoMédio_Coluna_Dados1 + 1 ' Faz até a variavel Coluna_DadosP ser igual a variavel Coluna_Dados1 mais uma unidade
    'Confere se o mês da coluna em questão é igual ao mês do dado mais recente
    If Left(Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(8, SituaçãoFinanceira_PreçoMédio_Coluna_DadosP), 2) = Left(Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(8, SituaçãoFinanceira_PreçoMédio_Coluna_Dados1), 2) Then
    'Caso seja igual, copia a coluna com os dados mais abaixo, a partir da linha 110
        Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Range(Cells(9, SituaçãoFinanceira_PreçoMédio_Coluna_DadosP), (Cells(54, SituaçãoFinanceira_PreçoMédio_Coluna_DadosP))).Copy (Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(110, SituaçãoFinanceira_PreçoMédio_Coluna_DadosP))
    End If
    'Vai para a próxima coluna
    SituaçãoFinanceira_PreçoMédio_Coluna_DadosP = SituaçãoFinanceira_PreçoMédio_Coluna_DadosP + 1
'Repete a conferencia com a próxima coluna
Loop


'Atribui os valores originais das variaveis após o loop anterior
SituaçãoFinanceira_PreçoMédio_Linha_Dados = 110
SituaçãoFinanceira_PreçoMédio_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Range("A9").End(xlToRight).Column 'Define o número da última coluna
SituaçãoFinanceira_PreçoMédio_Linha_Análise = 59
SituaçãoFinanceira_PreçoMédio_Coluna_Análise = 7

'Ordenação decrescente da série histórica dos meses do dado mais recente
Do Until SituaçãoFinanceira_PreçoMédio_Linha_Dados = 156 'Faz o calculo até a variável Linha_Dados ser 156
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo mês do mais recente
    posição = Application.WorksheetFunction.Rank_Eq(Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(SituaçãoFinanceira_PreçoMédio_Linha_Dados, SituaçãoFinanceira_PreçoMédio_Coluna_Dados1), Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Range(Cells(SituaçãoFinanceira_PreçoMédio_Linha_Dados, SituaçãoFinanceira_PreçoMédio_Coluna_Dados1), Cells(SituaçãoFinanceira_PreçoMédio_Linha_Dados, 2)))
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(SituaçãoFinanceira_PreçoMédio_Linha_Análise, SituaçãoFinanceira_PreçoMédio_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    SituaçãoFinanceira_PreçoMédio_Linha_Dados = SituaçãoFinanceira_PreçoMédio_Linha_Dados + 1
    SituaçãoFinanceira_PreçoMédio_Linha_Análise = SituaçãoFinanceira_PreçoMédio_Linha_Análise + 1
'Repete a ordenação com a próxima linha
Loop

'Atribui os valores originais das variaveis após o loop anterior
SituaçãoFinanceira_PreçoMédio_Linha_Dados = 110
SituaçãoFinanceira_PreçoMédio_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Range("A9").End(xlToRight).Column 'Define o número da última coluna
SituaçãoFinanceira_PreçoMédio_Linha_Análise = 59
SituaçãoFinanceira_PreçoMédio_Coluna_Análise = 8
'Ordenação crescente da série histórica dos meses do dado mais recente
Do Until SituaçãoFinanceira_PreçoMédio_Linha_Dados = 156 'Faz o calculo até a variável Linha_Dados ser 156
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo mês do mais recente
    posição = Application.WorksheetFunction.Rank_Eq(Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(SituaçãoFinanceira_PreçoMédio_Linha_Dados, SituaçãoFinanceira_PreçoMédio_Coluna_Dados1), Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Range(Cells(SituaçãoFinanceira_PreçoMédio_Linha_Dados, SituaçãoFinanceira_PreçoMédio_Coluna_Dados1), Cells(SituaçãoFinanceira_PreçoMédio_Linha_Dados, 2)), 1)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Cells(SituaçãoFinanceira_PreçoMédio_Linha_Análise, SituaçãoFinanceira_PreçoMédio_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    SituaçãoFinanceira_PreçoMédio_Linha_Dados = SituaçãoFinanceira_PreçoMédio_Linha_Dados + 1
    SituaçãoFinanceira_PreçoMédio_Linha_Análise = SituaçãoFinanceira_PreçoMédio_Linha_Análise + 1
'Repete a ordenação com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
SituaçãoFinanceira_PreçoMédio_Linha_Dados = 9
SituaçãoFinanceira_PreçoMédio_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA PREÇO MEDIO").Range("A9").End(xlToRight).Column
SituaçãoFinanceira_PreçoMédio_Coluna_Dados2 = SituaçãoFinanceira_PreçoMédio_Coluna_Dados1 - 1
SituaçãoFinanceira_PreçoMédio_Linha_Análise = 59
SituaçãoFinanceira_PreçoMédio_Coluna_Análise = 9

'Avaliação se cruzou ou não a linha de 50 e o sentido
Do Until SituaçãoFinanceira_PreçoMédio_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'se o dado mais recente for menor que 50 e o dado do mês anterior for maior ou igual a 50 então...
    If Cells(SituaçãoFinanceira_PreçoMédio_Linha_Dados, SituaçãoFinanceira_PreçoMédio_Coluna_Dados1) < 50 And Cells(SituaçãoFinanceira_PreçoMédio_Linha_Dados, SituaçãoFinanceira_PreçoMédio_Coluna_Dados2) >= 50 Then
    
    'a célula de análise recebe cruzou para baixo
    Cells(SituaçãoFinanceira_PreçoMédio_Linha_Análise, SituaçãoFinanceira_PreçoMédio_Coluna_Análise).Value = "Cruzou para baixo"
    'Caso não seja..
    Else
        'se o dado mais recente for maior ou igual a 50 e o dado do m~es anterior for menor ou igual a 50 então...
        If Cells(SituaçãoFinanceira_PreçoMédio_Linha_Dados, SituaçãoFinanceira_PreçoMédio_Coluna_Dados1) >= 50 And Cells(SituaçãoFinanceira_PreçoMédio_Linha_Dados, SituaçãoFinanceira_PreçoMédio_Coluna_Dados2) <= 50 Then
        'a célula de análise recebe cruzou para cima
        Cells(SituaçãoFinanceira_PreçoMédio_Linha_Análise, SituaçãoFinanceira_PreçoMédio_Coluna_Análise).Value = "Cruzou para cima"
        'Caso não seja..
        Else
        'a célula de análise recebe não cruzou
        Cells(SituaçãoFinanceira_PreçoMédio_Linha_Análise, SituaçãoFinanceira_PreçoMédio_Coluna_Análise).Value = "Não Cruzou"
        End If
    End If
    'Vai para a próxima linha de dados e de análise
    SituaçãoFinanceira_PreçoMédio_Linha_Dados = SituaçãoFinanceira_PreçoMédio_Linha_Dados + 1
    SituaçãoFinanceira_PreçoMédio_Linha_Análise = SituaçãoFinanceira_PreçoMédio_Linha_Análise + 1
'repete o processo com a nova linha
Loop

'Apaga as linhas com erros/dados faltantes/títulos e subtítulos
Range(Cells(60, 2), Cells(60, 9)).ClearContents
Range(Cells(66, 2), Cells(66, 9)).ClearContents
Range(Cells(70, 2), Cells(70, 9)).ClearContents
Range(Cells(72, 2), Cells(73, 9)).Value = "-"
Range(Cells(75, 2), Cells(75, 9)).Value = "-"
Range(Cells(79, 2), Cells(79, 9)).Value = "-"
Range(Cells(87, 2), Cells(87, 9)).Value = "-"

'Inserindo valores nas celulas vazias para fugir de bugs
Range(Cells(10, SituaçãoFinanceira_PreçoMédio_Coluna_Dados3), Cells(10, SituaçãoFinanceira_PreçoMédio_Coluna_Dados1)).ClearContents
Range(Cells(16, SituaçãoFinanceira_PreçoMédio_Coluna_Dados3), Cells(16, SituaçãoFinanceira_PreçoMédio_Coluna_Dados1)).ClearContents
Range(Cells(20, SituaçãoFinanceira_PreçoMédio_Coluna_Dados3), Cells(20, SituaçãoFinanceira_PreçoMédio_Coluna_Dados1)).ClearContents
Range(Cells(22, SituaçãoFinanceira_PreçoMédio_Coluna_Dados3), Cells(23, SituaçãoFinanceira_PreçoMédio_Coluna_Dados1)).Value = "-"
Range(Cells(25, SituaçãoFinanceira_PreçoMédio_Coluna_Dados3), Cells(25, SituaçãoFinanceira_PreçoMédio_Coluna_Dados1)).Value = "-"
Range(Cells(29, SituaçãoFinanceira_PreçoMédio_Coluna_Dados3), Cells(29, SituaçãoFinanceira_PreçoMédio_Coluna_Dados1)).Value = "-"
Range(Cells(37, SituaçãoFinanceira_PreçoMédio_Coluna_Dados3), Cells(37, SituaçãoFinanceira_PreçoMédio_Coluna_Dados1)).Value = "-"

Range("E59:H104").NumberFormat = "0"





'***************************     SituaçãoFinanceira         ****************************************************************



Dim SituaçãoFinanceira_Linha_Dados As Integer 'Define a linha que contém o dado a ser usado
Dim SituaçãoFinanceira_Coluna_Dados1 As Integer ' Define a coluna com o dado mais recente
Dim SituaçãoFinanceira_Coluna_Dados2 As Integer ' Define a coluna com o dado do mês anterior
Dim SituaçãoFinanceira_Coluna_Dados3 As Integer ' Defie a coluna com o dado do mesmo mês do ano anterior
Dim SituaçãoFinanceira_Linha_Análise As Integer ' Define a linha que será feita a análise
Dim SituaçãoFinanceira_Coluna_Análise As Integer 'Define a coluna que será feita a análise

Sheets("SITUACAO FINANCEIRA").Select

'Copia os títulos das categorias e cola onde será formada a tabela de análise
Sheets("SITUACAO FINANCEIRA").Range(Cells(9, 1), Cells(54, 9)).Copy (Sheets("SITUACAO FINANCEIRA").Cells(59, 1))
'Limpa os números que foram colados mas mantém a formatação
Sheets("SITUACAO FINANCEIRA").Range(Cells(59, 2), Cells(105, 9)).ClearContents



'Nomeia as colunas de acordo com o dado que será calculado nelas
Sheets("SITUACAO FINANCEIRA").Cells(58, 2).Value = "Diferença para o trimestre anterior"
Sheets("SITUACAO FINANCEIRA").Cells(58, 3).Value = "Diferença para ao mesmo trimestre do ano anterior"
Sheets("SITUACAO FINANCEIRA").Cells(58, 4).Value = "Diferença para a média histórica"
Sheets("SITUACAO FINANCEIRA").Cells(58, 5).Value = "Posição Decrescente (Maior valor 1º, menor valor último)"
Sheets("SITUACAO FINANCEIRA").Cells(58, 6).Value = "Posição Crescente  (Menor valor 1º, maior valor último)"
Sheets("SITUACAO FINANCEIRA").Cells(58, 7).Value = "Posição Crescente - Mesmo trimestre  (Menor valor 1º, maior valor último)"
Sheets("SITUACAO FINANCEIRA").Cells(58, 8).Value = "Posição Decrescente -Mesmo trimestre  (Maior valor 1º, menor valor último)"
Sheets("SITUACAO FINANCEIRA").Cells(58, 9).Value = "cruzou a linha de 50?"

'Atribui valores as variaveis definidas acima
SituaçãoFinanceira_Linha_Dados = 9 'Define o número da primeira linha de dados
SituaçãoFinanceira_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA").Range("A9").End(xlToRight).Column 'Define o número da última coluna
SituaçãoFinanceira_Coluna_Dados2 = SituaçãoFinanceira_Coluna_Dados1 - 1 'Define o número da coluna do mês anterior
SituaçãoFinanceira_Coluna_Dados3 = SituaçãoFinanceira_Coluna_Dados1 - 4
SituaçãoFinanceira_Linha_Análise = 59 'Define a primeira linhas de análises
SituaçãoFinanceira_Coluna_Análise = 2 'Define a coluna de análises

'Inserindo valores nas celulas vazias para fugir de bugs
Sheets("SITUACAO FINANCEIRA").Range(Cells(10, SituaçãoFinanceira_Coluna_Dados3), Cells(10, SituaçãoFinanceira_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA").Range(Cells(16, SituaçãoFinanceira_Coluna_Dados3), Cells(16, SituaçãoFinanceira_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA").Range(Cells(20, SituaçãoFinanceira_Coluna_Dados3), Cells(20, SituaçãoFinanceira_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA").Range(Cells(22, SituaçãoFinanceira_Coluna_Dados3), Cells(23, SituaçãoFinanceira_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA").Range(Cells(25, SituaçãoFinanceira_Coluna_Dados3), Cells(25, SituaçãoFinanceira_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA").Range(Cells(29, SituaçãoFinanceira_Coluna_Dados3), Cells(29, SituaçãoFinanceira_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA").Range(Cells(37, SituaçãoFinanceira_Coluna_Dados3), Cells(37, SituaçãoFinanceira_Coluna_Dados1)).Value = "0"


'Calculo da difernça em pontos do valor mais recente em relação ao valor do mês anterior
Do Until SituaçãoFinanceira_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Calculo da diferença em si: o valor da celula de analise é igual ao valor mais recente menos o valor do mês anterior
   Sheets("SITUACAO FINANCEIRA").Cells(SituaçãoFinanceira_Linha_Análise, SituaçãoFinanceira_Coluna_Análise).Value = Sheets("SITUACAO FINANCEIRA").Cells(SituaçãoFinanceira_Linha_Dados, SituaçãoFinanceira_Coluna_Dados1).Value - Sheets("SITUACAO FINANCEIRA").Cells(SituaçãoFinanceira_Linha_Dados, SituaçãoFinanceira_Coluna_Dados2).Value
    'Vai para a próxima linha de dados e de análise
   SituaçãoFinanceira_Linha_Dados = SituaçãoFinanceira_Linha_Dados + 1
   SituaçãoFinanceira_Linha_Análise = SituaçãoFinanceira_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
SituaçãoFinanceira_Linha_Dados = 9
SituaçãoFinanceira_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA").Range("A9").End(xlToRight).Column
SituaçãoFinanceira_Coluna_Dados3 = SituaçãoFinanceira_Coluna_Dados1 - 4
SituaçãoFinanceira_Linha_Análise = 59
SituaçãoFinanceira_Coluna_Análise = 3

'Cálculo da diferença em pontos do valor mais recente em relação ao valor do mesmo mês do ano anterior
Do Until SituaçãoFinanceira_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Calculo da diferença em si: o valor da celula de analise é igual ao valor mais recente menos o valor do mesmo mês do ano anterior
    Sheets("SITUACAO FINANCEIRA").Cells(SituaçãoFinanceira_Linha_Análise, SituaçãoFinanceira_Coluna_Análise).Value = Sheets("SITUACAO FINANCEIRA").Cells(SituaçãoFinanceira_Linha_Dados, SituaçãoFinanceira_Coluna_Dados1).Value - Sheets("SITUACAO FINANCEIRA").Cells(SituaçãoFinanceira_Linha_Dados, SituaçãoFinanceira_Coluna_Dados3).Value
    'Vai para a próxima linha de dados e de análise
    SituaçãoFinanceira_Linha_Dados = SituaçãoFinanceira_Linha_Dados + 1
    SituaçãoFinanceira_Linha_Análise = SituaçãoFinanceira_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
SituaçãoFinanceira_Linha_Dados = 9
SituaçãoFinanceira_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA").Range("A9").End(xlToRight).Column 'Define o número da última coluna
SituaçãoFinanceira_Linha_Análise = 59
SituaçãoFinanceira_Coluna_Análise = 4

'Cálculo da diferença em pontos do valor mais recente em relação ao valor da média histórica
Do Until SituaçãoFinanceira_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a variável media como a média do intervalo entre a coluna com o dado mais recente e o primeiro
    media = Application.Average(Sheets("SITUACAO FINANCEIRA").Range(Cells(SituaçãoFinanceira_Linha_Dados, SituaçãoFinanceira_Coluna_Dados1), Cells(SituaçãoFinanceira_Linha_Dados, 2)))
    'Calculo em si: o valor da celula de analise é igual ao valor mais recente menos o valor da média
    Sheets("SITUACAO FINANCEIRA").Cells(SituaçãoFinanceira_Linha_Análise, SituaçãoFinanceira_Coluna_Análise).Value = Sheets("SITUACAO FINANCEIRA").Cells(SituaçãoFinanceira_Linha_Dados, SituaçãoFinanceira_Coluna_Dados1).Value - media
    'Vai para a próxima linha de dados e de análise
    SituaçãoFinanceira_Linha_Dados = SituaçãoFinanceira_Linha_Dados + 1
    SituaçãoFinanceira_Linha_Análise = SituaçãoFinanceira_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
SituaçãoFinanceira_Linha_Dados = 9
SituaçãoFinanceira_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA").Range("A9").End(xlToRight).Column
SituaçãoFinanceira_Linha_Análise = 59
SituaçãoFinanceira_Coluna_Análise = 5

'Ordenação decrescente da série histórica completa
Do Until SituaçãoFinanceira_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posição = WorksheetFunction.Rank_Eq(Sheets("SITUACAO FINANCEIRA").Cells(SituaçãoFinanceira_Linha_Dados, SituaçãoFinanceira_Coluna_Dados1), Sheets("SITUACAO FINANCEIRA").Range(Cells(SituaçãoFinanceira_Linha_Dados, SituaçãoFinanceira_Coluna_Dados1), Cells(SituaçãoFinanceira_Linha_Dados, 2)), 0)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("SITUACAO FINANCEIRA").Cells(SituaçãoFinanceira_Linha_Análise, SituaçãoFinanceira_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    SituaçãoFinanceira_Linha_Dados = SituaçãoFinanceira_Linha_Dados + 1
    SituaçãoFinanceira_Linha_Análise = SituaçãoFinanceira_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
SituaçãoFinanceira_Linha_Dados = 9
SituaçãoFinanceira_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA").Range("A9").End(xlToRight).Column
SituaçãoFinanceira_Linha_Análise = 59
SituaçãoFinanceira_Coluna_Análise = 6

'Ordenação Crescente da série histórica completa
Do Until SituaçãoFinanceira_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posição = WorksheetFunction.Rank_Eq(Sheets("SITUACAO FINANCEIRA").Cells(SituaçãoFinanceira_Linha_Dados, SituaçãoFinanceira_Coluna_Dados1), Sheets("SITUACAO FINANCEIRA").Range(Cells(SituaçãoFinanceira_Linha_Dados, SituaçãoFinanceira_Coluna_Dados1), Cells(SituaçãoFinanceira_Linha_Dados, 2)), 1)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("SITUACAO FINANCEIRA").Cells(SituaçãoFinanceira_Linha_Análise, SituaçãoFinanceira_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    SituaçãoFinanceira_Linha_Dados = SituaçãoFinanceira_Linha_Dados + 1
    SituaçãoFinanceira_Linha_Análise = SituaçãoFinanceira_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior e define a variável Coluna_DadosP que representa a primeira coluna de dados
SituaçãoFinanceira_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA").Range("A9").End(xlToRight).Column
SituaçãoFinanceira_Coluna_DadosP = 2

Do Until SituaçãoFinanceira_Coluna_DadosP = SituaçãoFinanceira_Coluna_Dados1 + 1 ' Faz até a variavel Coluna_DadosP ser igual a variavel Coluna_Dados1 mais uma unidade
    'Confere se o mês da coluna em questão é igual ao mês do dado mais recente
    If Left(Sheets("SITUACAO FINANCEIRA").Cells(8, SituaçãoFinanceira_Coluna_DadosP), 2) = Left(Sheets("SITUACAO FINANCEIRA").Cells(8, SituaçãoFinanceira_Coluna_Dados1), 2) Then
    'Caso seja igual, copia a coluna com os dados mais abaixo, a partir da linha 110
        Sheets("SITUACAO FINANCEIRA").Range(Cells(9, SituaçãoFinanceira_Coluna_DadosP), (Cells(54, SituaçãoFinanceira_Coluna_DadosP))).Copy (Sheets("SITUACAO FINANCEIRA").Cells(110, SituaçãoFinanceira_Coluna_DadosP))
    End If
    'Vai para a próxima coluna
    SituaçãoFinanceira_Coluna_DadosP = SituaçãoFinanceira_Coluna_DadosP + 1
'Repete a conferencia com a próxima coluna
Loop


'Atribui os valores originais das variaveis após o loop anterior
SituaçãoFinanceira_Linha_Dados = 110
SituaçãoFinanceira_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA").Range("A9").End(xlToRight).Column 'Define o número da última coluna
SituaçãoFinanceira_Linha_Análise = 59
SituaçãoFinanceira_Coluna_Análise = 7

'Ordenação decrescente da série histórica dos meses do dado mais recente
Do Until SituaçãoFinanceira_Linha_Dados = 156 'Faz o calculo até a variável Linha_Dados ser 156
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo mês do mais recente
    posição = Application.WorksheetFunction.Rank_Eq(Sheets("SITUACAO FINANCEIRA").Cells(SituaçãoFinanceira_Linha_Dados, SituaçãoFinanceira_Coluna_Dados1), Sheets("SITUACAO FINANCEIRA").Range(Cells(SituaçãoFinanceira_Linha_Dados, SituaçãoFinanceira_Coluna_Dados1), Cells(SituaçãoFinanceira_Linha_Dados, 2)))
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("SITUACAO FINANCEIRA").Cells(SituaçãoFinanceira_Linha_Análise, SituaçãoFinanceira_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    SituaçãoFinanceira_Linha_Dados = SituaçãoFinanceira_Linha_Dados + 1
    SituaçãoFinanceira_Linha_Análise = SituaçãoFinanceira_Linha_Análise + 1
'Repete a ordenação com a próxima linha
Loop

'Atribui os valores originais das variaveis após o loop anterior
SituaçãoFinanceira_Linha_Dados = 110
SituaçãoFinanceira_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA").Range("A9").End(xlToRight).Column 'Define o número da última coluna
SituaçãoFinanceira_Linha_Análise = 59
SituaçãoFinanceira_Coluna_Análise = 8
'Ordenação crescente da série histórica dos meses do dado mais recente
Do Until SituaçãoFinanceira_Linha_Dados = 156 'Faz o calculo até a variável Linha_Dados ser 156
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo mês do mais recente
    posição = Application.WorksheetFunction.Rank_Eq(Sheets("SITUACAO FINANCEIRA").Cells(SituaçãoFinanceira_Linha_Dados, SituaçãoFinanceira_Coluna_Dados1), Sheets("SITUACAO FINANCEIRA").Range(Cells(SituaçãoFinanceira_Linha_Dados, SituaçãoFinanceira_Coluna_Dados1), Cells(SituaçãoFinanceira_Linha_Dados, 2)), 1)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("SITUACAO FINANCEIRA").Cells(SituaçãoFinanceira_Linha_Análise, SituaçãoFinanceira_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    SituaçãoFinanceira_Linha_Dados = SituaçãoFinanceira_Linha_Dados + 1
    SituaçãoFinanceira_Linha_Análise = SituaçãoFinanceira_Linha_Análise + 1
'Repete a ordenação com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
SituaçãoFinanceira_Linha_Dados = 9
SituaçãoFinanceira_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA").Range("A9").End(xlToRight).Column
SituaçãoFinanceira_Coluna_Dados2 = SituaçãoFinanceira_Coluna_Dados1 - 1
SituaçãoFinanceira_Linha_Análise = 59
SituaçãoFinanceira_Coluna_Análise = 9

'Avaliação se cruzou ou não a linha de 50 e o sentido
Do Until SituaçãoFinanceira_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'se o dado mais recente for menor que 50 e o dado do mês anterior for maior ou igual a 50 então...
    If Cells(SituaçãoFinanceira_Linha_Dados, SituaçãoFinanceira_Coluna_Dados1) < 50 And Cells(SituaçãoFinanceira_Linha_Dados, SituaçãoFinanceira_Coluna_Dados2) >= 50 Then
    
    'a célula de análise recebe cruzou para baixo
    Cells(SituaçãoFinanceira_Linha_Análise, SituaçãoFinanceira_Coluna_Análise).Value = "Cruzou para baixo"
    'Caso não seja..
    Else
        'se o dado mais recente for maior ou igual a 50 e o dado do m~es anterior for menor ou igual a 50 então...
        If Cells(SituaçãoFinanceira_Linha_Dados, SituaçãoFinanceira_Coluna_Dados1) >= 50 And Cells(SituaçãoFinanceira_Linha_Dados, SituaçãoFinanceira_Coluna_Dados2) <= 50 Then
        'a célula de análise recebe cruzou para cima
        Cells(SituaçãoFinanceira_Linha_Análise, SituaçãoFinanceira_Coluna_Análise).Value = "Cruzou para cima"
        'Caso não seja..
        Else
        'a célula de análise recebe não cruzou
        Cells(SituaçãoFinanceira_Linha_Análise, SituaçãoFinanceira_Coluna_Análise).Value = "Não Cruzou"
        End If
    End If
    'Vai para a próxima linha de dados e de análise
    SituaçãoFinanceira_Linha_Dados = SituaçãoFinanceira_Linha_Dados + 1
    SituaçãoFinanceira_Linha_Análise = SituaçãoFinanceira_Linha_Análise + 1
'repete o processo com a nova linha
Loop

'Apaga as linhas com erros/dados faltantes/títulos e subtítulos
Range(Cells(60, 2), Cells(60, 9)).ClearContents
Range(Cells(66, 2), Cells(66, 9)).ClearContents
Range(Cells(70, 2), Cells(70, 9)).ClearContents
Range(Cells(72, 2), Cells(73, 9)).Value = "-"
Range(Cells(75, 2), Cells(75, 9)).Value = "-"
Range(Cells(79, 2), Cells(79, 9)).Value = "-"
Range(Cells(87, 2), Cells(87, 9)).Value = "-"

'Inserindo valores nas celulas vazias para fugir de bugs
Range(Cells(10, SituaçãoFinanceira_Coluna_Dados3), Cells(10, SituaçãoFinanceira_Coluna_Dados1)).ClearContents
Range(Cells(16, SituaçãoFinanceira_Coluna_Dados3), Cells(16, SituaçãoFinanceira_Coluna_Dados1)).ClearContents
Range(Cells(20, SituaçãoFinanceira_Coluna_Dados3), Cells(20, SituaçãoFinanceira_Coluna_Dados1)).ClearContents
Range(Cells(22, SituaçãoFinanceira_Coluna_Dados3), Cells(23, SituaçãoFinanceira_Coluna_Dados1)).Value = "-"
Range(Cells(25, SituaçãoFinanceira_Coluna_Dados3), Cells(25, SituaçãoFinanceira_Coluna_Dados1)).Value = "-"
Range(Cells(29, SituaçãoFinanceira_Coluna_Dados3), Cells(29, SituaçãoFinanceira_Coluna_Dados1)).Value = "-"
Range(Cells(37, SituaçãoFinanceira_Coluna_Dados3), Cells(37, SituaçãoFinanceira_Coluna_Dados1)).Value = "-"

Range("E59:H104").NumberFormat = "0"


'****************************************                SituaçãoFinanceira_Credito           *********************************************

Dim SituaçãoFinanceira_Credito_Linha_Dados As Integer 'Define a linha que contém o dado a ser usado
Dim SituaçãoFinanceira_Credito_Coluna_Dados1 As Integer ' Define a coluna com o dado mais recente
Dim SituaçãoFinanceira_Credito_Coluna_Dados2 As Integer ' Define a coluna com o dado do mês anterior
Dim SituaçãoFinanceira_Credito_Coluna_Dados3 As Integer ' Defie a coluna com o dado do mesmo mês do ano anterior
Dim SituaçãoFinanceira_Credito_Linha_Análise As Integer ' Define a linha que será feita a análise
Dim SituaçãoFinanceira_Credito_Coluna_Análise As Integer 'Define a coluna que será feita a análise

Sheets("SITUACAO FINANCEIRA CREDITO").Select

'Copia os títulos das categorias e cola onde será formada a tabela de análise
Sheets("SITUACAO FINANCEIRA CREDITO").Range(Cells(9, 1), Cells(54, 9)).Copy (Sheets("SITUACAO FINANCEIRA CREDITO").Cells(59, 1))
'Limpa os números que foram colados mas mantém a formatação
Sheets("SITUACAO FINANCEIRA CREDITO").Range(Cells(59, 2), Cells(105, 9)).ClearContents



'Nomeia as colunas de acordo com o dado que será calculado nelas
Sheets("SITUACAO FINANCEIRA CREDITO").Cells(58, 2).Value = "Diferença para o trimestre anterior"
Sheets("SITUACAO FINANCEIRA CREDITO").Cells(58, 3).Value = "Diferença para ao mesmo trimestre do ano anterior"
Sheets("SITUACAO FINANCEIRA CREDITO").Cells(58, 4).Value = "Diferença para a média histórica"
Sheets("SITUACAO FINANCEIRA CREDITO").Cells(58, 5).Value = "Posição Decrescente (Maior valor 1º, menor valor último)"
Sheets("SITUACAO FINANCEIRA CREDITO").Cells(58, 6).Value = "Posição Crescente  (Menor valor 1º, maior valor último)"
Sheets("SITUACAO FINANCEIRA CREDITO").Cells(58, 7).Value = "Posição Crescente - Mesmo trimestre  (Menor valor 1º, maior valor último)"
Sheets("SITUACAO FINANCEIRA CREDITO").Cells(58, 8).Value = "Posição Decrescente -Mesmo trimestre  (Maior valor 1º, menor valor último)"
Sheets("SITUACAO FINANCEIRA CREDITO").Cells(58, 9).Value = "cruzou a linha de 50?"

'Atribui valores as variaveis definidas acima
SituaçãoFinanceira_Credito_Linha_Dados = 9 'Define o número da primeira linha de dados
SituaçãoFinanceira_Credito_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA CREDITO").Range("A9").End(xlToRight).Column 'Define o número da última coluna
SituaçãoFinanceira_Credito_Coluna_Dados2 = SituaçãoFinanceira_Credito_Coluna_Dados1 - 1 'Define o número da coluna do mês anterior
SituaçãoFinanceira_Credito_Coluna_Dados3 = SituaçãoFinanceira_Credito_Coluna_Dados1 - 4
SituaçãoFinanceira_Credito_Linha_Análise = 59 'Define a primeira linhas de análises
SituaçãoFinanceira_Credito_Coluna_Análise = 2 'Define a coluna de análises

'Inserindo valores nas celulas vazias para fugir de bugs
Sheets("SITUACAO FINANCEIRA CREDITO").Range(Cells(10, SituaçãoFinanceira_Credito_Coluna_Dados3), Cells(10, SituaçãoFinanceira_Credito_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA CREDITO").Range(Cells(16, SituaçãoFinanceira_Credito_Coluna_Dados3), Cells(16, SituaçãoFinanceira_Credito_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA CREDITO").Range(Cells(20, SituaçãoFinanceira_Credito_Coluna_Dados3), Cells(20, SituaçãoFinanceira_Credito_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA CREDITO").Range(Cells(22, SituaçãoFinanceira_Credito_Coluna_Dados3), Cells(23, SituaçãoFinanceira_Credito_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA CREDITO").Range(Cells(25, SituaçãoFinanceira_Credito_Coluna_Dados3), Cells(25, SituaçãoFinanceira_Credito_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA CREDITO").Range(Cells(29, SituaçãoFinanceira_Credito_Coluna_Dados3), Cells(29, SituaçãoFinanceira_Credito_Coluna_Dados1)).Value = "0"
Sheets("SITUACAO FINANCEIRA CREDITO").Range(Cells(37, SituaçãoFinanceira_Credito_Coluna_Dados3), Cells(37, SituaçãoFinanceira_Credito_Coluna_Dados1)).Value = "0"


'Calculo da difernça em pontos do valor mais recente em relação ao valor do mês anterior
Do Until SituaçãoFinanceira_Credito_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Calculo da diferença em si: o valor da celula de analise é igual ao valor mais recente menos o valor do mês anterior
   Sheets("SITUACAO FINANCEIRA CREDITO").Cells(SituaçãoFinanceira_Credito_Linha_Análise, SituaçãoFinanceira_Credito_Coluna_Análise).Value = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(SituaçãoFinanceira_Credito_Linha_Dados, SituaçãoFinanceira_Credito_Coluna_Dados1).Value - Sheets("SITUACAO FINANCEIRA CREDITO").Cells(SituaçãoFinanceira_Credito_Linha_Dados, SituaçãoFinanceira_Credito_Coluna_Dados2).Value
    'Vai para a próxima linha de dados e de análise
   SituaçãoFinanceira_Credito_Linha_Dados = SituaçãoFinanceira_Credito_Linha_Dados + 1
   SituaçãoFinanceira_Credito_Linha_Análise = SituaçãoFinanceira_Credito_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
SituaçãoFinanceira_Credito_Linha_Dados = 9
SituaçãoFinanceira_Credito_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA CREDITO").Range("A9").End(xlToRight).Column
SituaçãoFinanceira_Credito_Coluna_Dados3 = SituaçãoFinanceira_Credito_Coluna_Dados1 - 4
SituaçãoFinanceira_Credito_Linha_Análise = 59
SituaçãoFinanceira_Credito_Coluna_Análise = 3

'Cálculo da diferença em pontos do valor mais recente em relação ao valor do mesmo mês do ano anterior
Do Until SituaçãoFinanceira_Credito_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Calculo da diferença em si: o valor da celula de analise é igual ao valor mais recente menos o valor do mesmo mês do ano anterior
    Sheets("SITUACAO FINANCEIRA CREDITO").Cells(SituaçãoFinanceira_Credito_Linha_Análise, SituaçãoFinanceira_Credito_Coluna_Análise).Value = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(SituaçãoFinanceira_Credito_Linha_Dados, SituaçãoFinanceira_Credito_Coluna_Dados1).Value - Sheets("SITUACAO FINANCEIRA CREDITO").Cells(SituaçãoFinanceira_Credito_Linha_Dados, SituaçãoFinanceira_Credito_Coluna_Dados3).Value
    'Vai para a próxima linha de dados e de análise
    SituaçãoFinanceira_Credito_Linha_Dados = SituaçãoFinanceira_Credito_Linha_Dados + 1
    SituaçãoFinanceira_Credito_Linha_Análise = SituaçãoFinanceira_Credito_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
SituaçãoFinanceira_Credito_Linha_Dados = 9
SituaçãoFinanceira_Credito_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA CREDITO").Range("A9").End(xlToRight).Column 'Define o número da última coluna
SituaçãoFinanceira_Credito_Linha_Análise = 59
SituaçãoFinanceira_Credito_Coluna_Análise = 4

'Cálculo da diferença em pontos do valor mais recente em relação ao valor da média histórica
Do Until SituaçãoFinanceira_Credito_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a variável media como a média do intervalo entre a coluna com o dado mais recente e o primeiro
    media = Application.Average(Sheets("SITUACAO FINANCEIRA CREDITO").Range(Cells(SituaçãoFinanceira_Credito_Linha_Dados, SituaçãoFinanceira_Credito_Coluna_Dados1), Cells(SituaçãoFinanceira_Credito_Linha_Dados, 2)))
    'Calculo em si: o valor da celula de analise é igual ao valor mais recente menos o valor da média
    Sheets("SITUACAO FINANCEIRA CREDITO").Cells(SituaçãoFinanceira_Credito_Linha_Análise, SituaçãoFinanceira_Credito_Coluna_Análise).Value = Sheets("SITUACAO FINANCEIRA CREDITO").Cells(SituaçãoFinanceira_Credito_Linha_Dados, SituaçãoFinanceira_Credito_Coluna_Dados1).Value - media
    'Vai para a próxima linha de dados e de análise
    SituaçãoFinanceira_Credito_Linha_Dados = SituaçãoFinanceira_Credito_Linha_Dados + 1
    SituaçãoFinanceira_Credito_Linha_Análise = SituaçãoFinanceira_Credito_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
SituaçãoFinanceira_Credito_Linha_Dados = 9
SituaçãoFinanceira_Credito_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA CREDITO").Range("A9").End(xlToRight).Column
SituaçãoFinanceira_Credito_Linha_Análise = 59
SituaçãoFinanceira_Credito_Coluna_Análise = 5

'Ordenação decrescente da série histórica completa
Do Until SituaçãoFinanceira_Credito_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posição = WorksheetFunction.Rank_Eq(Sheets("SITUACAO FINANCEIRA CREDITO").Cells(SituaçãoFinanceira_Credito_Linha_Dados, SituaçãoFinanceira_Credito_Coluna_Dados1), Sheets("SITUACAO FINANCEIRA CREDITO").Range(Cells(SituaçãoFinanceira_Credito_Linha_Dados, SituaçãoFinanceira_Credito_Coluna_Dados1), Cells(SituaçãoFinanceira_Credito_Linha_Dados, 2)), 0)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("SITUACAO FINANCEIRA CREDITO").Cells(SituaçãoFinanceira_Credito_Linha_Análise, SituaçãoFinanceira_Credito_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    SituaçãoFinanceira_Credito_Linha_Dados = SituaçãoFinanceira_Credito_Linha_Dados + 1
    SituaçãoFinanceira_Credito_Linha_Análise = SituaçãoFinanceira_Credito_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
SituaçãoFinanceira_Credito_Linha_Dados = 9
SituaçãoFinanceira_Credito_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA CREDITO").Range("A9").End(xlToRight).Column
SituaçãoFinanceira_Credito_Linha_Análise = 59
SituaçãoFinanceira_Credito_Coluna_Análise = 6

'Ordenação Crescente da série histórica completa
Do Until SituaçãoFinanceira_Credito_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado
    posição = WorksheetFunction.Rank_Eq(Sheets("SITUACAO FINANCEIRA CREDITO").Cells(SituaçãoFinanceira_Credito_Linha_Dados, SituaçãoFinanceira_Credito_Coluna_Dados1), Sheets("SITUACAO FINANCEIRA CREDITO").Range(Cells(SituaçãoFinanceira_Credito_Linha_Dados, SituaçãoFinanceira_Credito_Coluna_Dados1), Cells(SituaçãoFinanceira_Credito_Linha_Dados, 2)), 1)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("SITUACAO FINANCEIRA CREDITO").Cells(SituaçãoFinanceira_Credito_Linha_Análise, SituaçãoFinanceira_Credito_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    SituaçãoFinanceira_Credito_Linha_Dados = SituaçãoFinanceira_Credito_Linha_Dados + 1
    SituaçãoFinanceira_Credito_Linha_Análise = SituaçãoFinanceira_Credito_Linha_Análise + 1
'Refaz o calculo com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior e define a variável Coluna_DadosP que representa a primeira coluna de dados
SituaçãoFinanceira_Credito_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA CREDITO").Range("A9").End(xlToRight).Column
SituaçãoFinanceira_Credito_Coluna_DadosP = 2

Do Until SituaçãoFinanceira_Credito_Coluna_DadosP = SituaçãoFinanceira_Credito_Coluna_Dados1 + 1 ' Faz até a variavel Coluna_DadosP ser igual a variavel Coluna_Dados1 mais uma unidade
    'Confere se o mês da coluna em questão é igual ao mês do dado mais recente
    If Left(Sheets("SITUACAO FINANCEIRA CREDITO").Cells(8, SituaçãoFinanceira_Credito_Coluna_DadosP), 2) = Left(Sheets("SITUACAO FINANCEIRA CREDITO").Cells(8, SituaçãoFinanceira_Credito_Coluna_Dados1), 2) Then
    'Caso seja igual, copia a coluna com os dados mais abaixo, a partir da linha 110
        Sheets("SITUACAO FINANCEIRA CREDITO").Range(Cells(9, SituaçãoFinanceira_Credito_Coluna_DadosP), (Cells(54, SituaçãoFinanceira_Credito_Coluna_DadosP))).Copy (Sheets("SITUACAO FINANCEIRA CREDITO").Cells(110, SituaçãoFinanceira_Credito_Coluna_DadosP))
    End If
    'Vai para a próxima coluna
    SituaçãoFinanceira_Credito_Coluna_DadosP = SituaçãoFinanceira_Credito_Coluna_DadosP + 1
'Repete a conferencia com a próxima coluna
Loop


'Atribui os valores originais das variaveis após o loop anterior
SituaçãoFinanceira_Credito_Linha_Dados = 110
SituaçãoFinanceira_Credito_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA CREDITO").Range("A9").End(xlToRight).Column 'Define o número da última coluna
SituaçãoFinanceira_Credito_Linha_Análise = 59
SituaçãoFinanceira_Credito_Coluna_Análise = 7

'Ordenação decrescente da série histórica dos meses do dado mais recente
Do Until SituaçãoFinanceira_Credito_Linha_Dados = 156 'Faz o calculo até a variável Linha_Dados ser 156
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo mês do mais recente
    posição = Application.WorksheetFunction.Rank_Eq(Sheets("SITUACAO FINANCEIRA CREDITO").Cells(SituaçãoFinanceira_Credito_Linha_Dados, SituaçãoFinanceira_Credito_Coluna_Dados1), Sheets("SITUACAO FINANCEIRA CREDITO").Range(Cells(SituaçãoFinanceira_Credito_Linha_Dados, SituaçãoFinanceira_Credito_Coluna_Dados1), Cells(SituaçãoFinanceira_Credito_Linha_Dados, 2)))
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("SITUACAO FINANCEIRA CREDITO").Cells(SituaçãoFinanceira_Credito_Linha_Análise, SituaçãoFinanceira_Credito_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    SituaçãoFinanceira_Credito_Linha_Dados = SituaçãoFinanceira_Credito_Linha_Dados + 1
    SituaçãoFinanceira_Credito_Linha_Análise = SituaçãoFinanceira_Credito_Linha_Análise + 1
'Repete a ordenação com a próxima linha
Loop

'Atribui os valores originais das variaveis após o loop anterior
SituaçãoFinanceira_Credito_Linha_Dados = 110
SituaçãoFinanceira_Credito_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA CREDITO").Range("A9").End(xlToRight).Column 'Define o número da última coluna
SituaçãoFinanceira_Credito_Linha_Análise = 59
SituaçãoFinanceira_Credito_Coluna_Análise = 8
'Ordenação crescente da série histórica dos meses do dado mais recente
Do Until SituaçãoFinanceira_Credito_Linha_Dados = 156 'Faz o calculo até a variável Linha_Dados ser 156
    'Define a varável posição como a aplicação da fómula Rank.EQ no intervalo dado pelo dado mais recente e o primeiro dado com o mesmo mês do mais recente
    posição = Application.WorksheetFunction.Rank_Eq(Sheets("SITUACAO FINANCEIRA CREDITO").Cells(SituaçãoFinanceira_Credito_Linha_Dados, SituaçãoFinanceira_Credito_Coluna_Dados1), Sheets("SITUACAO FINANCEIRA CREDITO").Range(Cells(SituaçãoFinanceira_Credito_Linha_Dados, SituaçãoFinanceira_Credito_Coluna_Dados1), Cells(SituaçãoFinanceira_Credito_Linha_Dados, 2)), 1)
    'Define que a célula da análise seja igual a posição do dado mais recente
    Sheets("SITUACAO FINANCEIRA CREDITO").Cells(SituaçãoFinanceira_Credito_Linha_Análise, SituaçãoFinanceira_Credito_Coluna_Análise).Value = posição
    'Vai para a próxima linha de dados e de análise
    SituaçãoFinanceira_Credito_Linha_Dados = SituaçãoFinanceira_Credito_Linha_Dados + 1
    SituaçãoFinanceira_Credito_Linha_Análise = SituaçãoFinanceira_Credito_Linha_Análise + 1
'Repete a ordenação com a próxima linha
Loop


'Atribui os valores originais das variaveis após o loop anterior
SituaçãoFinanceira_Credito_Linha_Dados = 9
SituaçãoFinanceira_Credito_Coluna_Dados1 = Sheets("SITUACAO FINANCEIRA CREDITO").Range("A9").End(xlToRight).Column
SituaçãoFinanceira_Credito_Coluna_Dados2 = SituaçãoFinanceira_Credito_Coluna_Dados1 - 1
SituaçãoFinanceira_Credito_Linha_Análise = 59
SituaçãoFinanceira_Credito_Coluna_Análise = 9

'Avaliação se cruzou ou não a linha de 50 e o sentido
Do Until SituaçãoFinanceira_Credito_Linha_Dados = 55 'Faz o calculo até a variável Linha_Dados ser 55
    'se o dado mais recente for menor que 50 e o dado do mês anterior for maior ou igual a 50 então...
    If Cells(SituaçãoFinanceira_Credito_Linha_Dados, SituaçãoFinanceira_Credito_Coluna_Dados1) < 50 And Cells(SituaçãoFinanceira_Credito_Linha_Dados, SituaçãoFinanceira_Credito_Coluna_Dados2) >= 50 Then
    
    'a célula de análise recebe cruzou para baixo
    Cells(SituaçãoFinanceira_Credito_Linha_Análise, SituaçãoFinanceira_Credito_Coluna_Análise).Value = "Cruzou para baixo"
    'Caso não seja..
    Else
        'se o dado mais recente for maior ou igual a 50 e o dado do m~es anterior for menor ou igual a 50 então...
        If Cells(SituaçãoFinanceira_Credito_Linha_Dados, SituaçãoFinanceira_Credito_Coluna_Dados1) >= 50 And Cells(SituaçãoFinanceira_Credito_Linha_Dados, SituaçãoFinanceira_Credito_Coluna_Dados2) <= 50 Then
        'a célula de análise recebe cruzou para cima
        Cells(SituaçãoFinanceira_Credito_Linha_Análise, SituaçãoFinanceira_Credito_Coluna_Análise).Value = "Cruzou para cima"
        'Caso não seja..
        Else
        'a célula de análise recebe não cruzou
        Cells(SituaçãoFinanceira_Credito_Linha_Análise, SituaçãoFinanceira_Credito_Coluna_Análise).Value = "Não Cruzou"
        End If
    End If
    'Vai para a próxima linha de dados e de análise
    SituaçãoFinanceira_Credito_Linha_Dados = SituaçãoFinanceira_Credito_Linha_Dados + 1
    SituaçãoFinanceira_Credito_Linha_Análise = SituaçãoFinanceira_Credito_Linha_Análise + 1
'repete o processo com a nova linha
Loop

'Apaga as linhas com erros/dados faltantes/títulos e subtítulos
Range(Cells(60, 2), Cells(60, 9)).ClearContents
Range(Cells(66, 2), Cells(66, 9)).ClearContents
Range(Cells(70, 2), Cells(70, 9)).ClearContents
Range(Cells(72, 2), Cells(73, 9)).Value = "-"
Range(Cells(75, 2), Cells(75, 9)).Value = "-"
Range(Cells(79, 2), Cells(79, 9)).Value = "-"
Range(Cells(87, 2), Cells(87, 9)).Value = "-"

'Inserindo valores nas celulas vazias para fugir de bugs
Range(Cells(10, SituaçãoFinanceira_Credito_Coluna_Dados3), Cells(10, SituaçãoFinanceira_Credito_Coluna_Dados1)).ClearContents
Range(Cells(16, SituaçãoFinanceira_Credito_Coluna_Dados3), Cells(16, SituaçãoFinanceira_Credito_Coluna_Dados1)).ClearContents
Range(Cells(20, SituaçãoFinanceira_Credito_Coluna_Dados3), Cells(20, SituaçãoFinanceira_Credito_Coluna_Dados1)).ClearContents
Range(Cells(22, SituaçãoFinanceira_Credito_Coluna_Dados3), Cells(23, SituaçãoFinanceira_Credito_Coluna_Dados1)).Value = "-"
Range(Cells(25, SituaçãoFinanceira_Credito_Coluna_Dados3), Cells(25, SituaçãoFinanceira_Credito_Coluna_Dados1)).Value = "-"
Range(Cells(29, SituaçãoFinanceira_Credito_Coluna_Dados3), Cells(29, SituaçãoFinanceira_Credito_Coluna_Dados1)).Value = "-"
Range(Cells(37, SituaçãoFinanceira_Credito_Coluna_Dados3), Cells(37, SituaçãoFinanceira_Credito_Coluna_Dados1)).Value = "-"

Range("E59:H104").NumberFormat = "0"

End Sub

Sub Formatação()

Dim Sondagem As Workbook
Dim Modelo As Workbook
    
'   Capture current workbook
    Set Sondagem = ActiveWorkbook
    
'   Open new workbook
    Workbooks.Open ("C:\Users\e-gustavo.oliveira\CNI - Confederação Nacional da Indústria\ECON - 1 Indicadores Econômicos CNI\1 Indicadores de Atividade Industrial\Sondagem Industrial\Automação\Templates\Gráficos e Tabelas - Modelo Trimestral.xlsm")

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

