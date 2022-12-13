Attribute VB_Name = "Automa��o"
Sub Aperte_o_Play()
'Executa a rotina
Call Desempenho
Call Situa��o_Financeira
Call ICEI
Call Perspectivas
Call Principais_Problemas
'Oculta as abas
Sheets("Indicadores").Visible = False
Sheets("Principais Problemas").Visible = False

End Sub
Sub Desempenho()
'Define as vari�veis
Dim coluna As Integer
Dim Gr�fico_Desempenho As Object
Dim m�dia As Single

'Seleciona a aba Indicadores
Sheets("Indicadores").Select
'Pega o n�mero da pultima coluna
coluna = Range("D9").End(xlToRight).Column
'Calcula a m�dia hist�rica
m�dia = Application.WorksheetFunction.Average(Range(Cells(10, 4), Cells(10, coluna)))
'Adiciona a aba Desempenho
Sheets.Add(Before:=Sheets("Indicadores")).Name = "Desempenho"
'Seleciona a aba desempenho
Sheets("Desempenho").Select
'Escreve os dois prmeiros meses da s�rie
Range("B1").Value = "01/01/12"
Range("C1").Value = "02/01/12"
'Completa a linha dos meses at� o �ltimo m�s
Range("B1:C1").Select
    Selection.NumberFormat = "mmm-yy"
    Selection.AutoFill Destination:=Range(Cells(1, 2), Cells(1, coluna - 2)), Type:=xlFillDefault
    
'Nomeia as s�ries
Range("A2").Value = "�ndice de Desempenho da Pequena Empresa"
Range("A3").Value = "M�dia hist�rica"

'Copia e cola as s�ries
Sheets("Indicadores").Select
Range(Cells(10, 4), Cells(10, coluna)).Copy
Sheets("Desempenho").Select
Range("B2").PasteSpecial xlPasteValues
'Atribui o valor da m�dia
Range(Cells(3, 2), Cells(3, coluna - 2)).Value = m�dia
Range("A1").Select

'Cria o gr�fico
Set Gr�fico_Desempenho = Sheets("Desempenho").Shapes.AddChart2

Gr�fico_Desempenho.Select ' Seleciona o Gr�fico
    '1/7 ActiveChart.ApplyChartTemplate ("C:\Users\paula.verlangeira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Panorama da Pequena Ind�stria\Automa��o\Desempenho_PPI.crtx") ' Aplica o template do gr�fico
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Panorama da Pequena Ind�stria\Automa��o\Desempenho_PPI.crtx") ' Aplica o template do gr�fico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gr�fico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gr�fico
    ActiveChart.Parent.Top = Parent.Range("B5").Top 'reposiciona o grafico em rela��o ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("B5").Left ' reposiciona o gr�fico em rela��o � borda esquerda da planilha
    ActiveChart.FullSeriesCollection(1).Name = "=Desempenho!" & Cells(2, 1).Address 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(1).Values = "=Desempenho!" & Range(Cells(2, coluna - 121), Cells(2, coluna - 2)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(1).XValues = "=Desempenho!" & Range(Cells(1, coluna - 121), Cells(1, coluna - 2)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.FullSeriesCollection(2).Name = "=Desempenho!$A$3" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(2).Values = "=Desempenho!" & Range(Cells(3, coluna - 121), Cells(3, coluna - 2)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(2).XValues = "=Desempenho!" & Range(Cells(1, coluna - 121), Cells(1, coluna - 2)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
End Sub

Sub Situa��o_Financeira()
'Define as vari�veis
Dim coluna As Integer
Dim Gr�fico_Situa��o As Object
Dim m�dia As Single
'Seleciona a aba indicadores
Sheets("Indicadores").Select
'pega o n�mero da �ltima coluna
coluna = Range("D20").End(xlToRight).Column
'Calcula a m��edia hist�rica
m�dia = Application.WorksheetFunction.Average(Range(Cells(20, 4), Cells(20, coluna)))
'Adiciona a aba Situa��o Financiera
Sheets.Add(Before:=Sheets("Indicadores")).Name = "Situa��o Financeira"
'Seleciona a aba situa��o fincanceira
Sheets("Situa��o Financeira").Select
'Escreve os primeiros trimestres da serie
Range("B1").Value = "I-12"
Range("C1").Value = "II-12"
Range("D1").Value = "III-12"
Range("E1").Value = "IV-12"
'Completa a s�rie dos trimestres
Range("B1:E1").Select
    Selection.NumberFormat = "mmm-yy"
    Selection.AutoFill Destination:=Range(Cells(1, 2), Cells(1, coluna - 2)), Type:=xlFillDefault
'Nomeia as s�ries
Range("A2").Value = "�ndice de Situa��o Financeira da Pequena Empresa"
Range("A3").Value = "M�dia hist�rica"
'Copia e cola os valores da s�rie
Sheets("Indicadores").Select
Range(Cells(20, 4), Cells(20, coluna)).Copy
Sheets("Situa��o Financeira").Select
Range("B2").PasteSpecial xlPasteValues
'Atribui os valores da m�dia
Range(Cells(3, 2), Cells(3, coluna - 2)).Value = m�dia
Range("A1").Select

'Adiciona o gr�fico
Set Gr�fico_Situa��o = Sheets("Situa��o Financeira").Shapes.AddChart2

Gr�fico_Situa��o.Select ' Seleciona o Gr�fico
    '2/7 ActiveChart.ApplyChartTemplate ("C:\Users\paula.verlangeira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Panorama da Pequena Ind�stria\Automa��o\Situa��o Financeira PPI.crtx") ' Aplica o template do gr�fico
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Panorama da Pequena Ind�stria\Automa��o\Situa��o Financeira PPI.crtx") ' Aplica o template do gr�fico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gr�fico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gr�fico
    ActiveChart.Parent.Top = Parent.Range("B5").Top 'reposiciona o grafico em rela��o ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("B5").Left ' reposiciona o gr�fico em rela��o � borda esquerda da planilha
    ActiveChart.FullSeriesCollection(1).Name = "='Situa��o Financeira'!$A$2" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(1).Values = "='Situa��o Financeira'!" & Range(Cells(2, coluna - 42), Cells(2, coluna - 2)).Address
    ActiveChart.FullSeriesCollection(1).XValues = "='Situa��o Financeira'!" & Range(Cells(1, coluna - 42), Cells(1, coluna - 2)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.FullSeriesCollection(2).Name = "='Situa��o Financeira'!$A$3" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(2).Values = "='Situa��o Financeira'!" & Range(Cells(3, coluna - 40), Cells(3, coluna - 2)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(2).XValues = "='Situa��o Financeira'!" & Range(Cells(1, coluna - 40), Cells(1, coluna - 2)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
 
End Sub

Sub ICEI()
'Define as vari�veis
Dim coluna As Integer
Dim Gr�fico_ICEI As Object
Dim m�dia As Single
'Seleciona a aba indicadores
Sheets("Indicadores").Select
'pega o n�mero da �ltima coluna
coluna = Range("D40").End(xlToRight).Column
'Calcula a m�dia hist�rica
m�dia = Application.WorksheetFunction.Average(Range(Cells(40, 4), Cells(40, coluna)))
'Adiciona a aba Situa��o Financiera
Sheets.Add(Before:=Sheets("Indicadores")).Name = "ICEI"
'Seleciona a aba situa��o fincanceira
Sheets("ICEI").Select
'Escreve os primeiros trimestres da serie
Range("B1").Value = "01/01/10"
Range("C1").Value = "02/01/10"
'Completa a s�rie dos trimestres
Range("B1:C1").Select
    Selection.NumberFormat = "mmm-yy"
    Selection.AutoFill Destination:=Range(Cells(1, 2), Cells(1, coluna - 2)), Type:=xlFillDefault
'Atribui os nomes das s�ries
Range("A2").Value = "ICEI"
Range("A3").Value = "M�dia hist�rica"
'Copia e cola os valores da s�rie
Sheets("Indicadores").Select
Range(Cells(40, 4), Cells(40, coluna)).Copy
Sheets("ICEI").Select
Range("B2").PasteSpecial xlPasteValues
'Atribui os valores da m�dia
Range(Cells(3, 2), Cells(3, coluna - 2)).Value = m�dia
Range("A1").Select

'Adiciona o gr�fico
Set Gr�fico_ICEI = Sheets("ICEI").Shapes.AddChart2

Gr�fico_ICEI.Select ' Seleciona o Gr�fico
    '3/7 ActiveChart.ApplyChartTemplate ("C:\Users\paula.verlangeiro\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Panorama da Pequena Ind�stria\Automa��o\ICEI_PPI.crtx") ' Aplica o template do gr�fico
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Panorama da Pequena Ind�stria\Automa��o\ICEI_PPI.crtx") ' Aplica o template do gr�fico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gr�fico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gr�fico
    ActiveChart.Parent.Top = Parent.Range("B5").Top 'reposiciona o grafico em rela��o ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("B5").Left ' reposiciona o gr�fico em rela��o � borda esquerda da planilha
    ActiveChart.FullSeriesCollection(1).Name = "=ICEI!" & Cells(2, 1).Address 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(1).Values = "=ICEI!" & Range(Cells(2, coluna - 122), Cells(2, coluna - 2)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(1).XValues = "=ICEI!" & Range(Cells(1, coluna - 122), Cells(1, coluna - 2)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.FullSeriesCollection(2).Name = "=ICEI!$A$3" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(2).Values = "=ICEI!" & Range(Cells(3, coluna - 122), Cells(3, coluna - 2)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(2).XValues = "=ICEI!" & Range(Cells(1, coluna - 122), Cells(1, coluna - 2)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
End Sub
Sub Perspectivas()
'Define as vari�veis
Dim coluna As Integer
Dim Gr�fico_Perspectivas As Object
Dim m�dia As Single
'Seleciona a aba indicadores
Sheets("Indicadores").Select
'pega o n�mero da �ltima coluna
coluna = Range("D30").End(xlToRight).Column
'Calcula a m�dia hist�rica
m�dia = Application.WorksheetFunction.Average(Range(Cells(30, 4), Cells(30, coluna)))
'Adiciona a aba Situa��o Financiera
Sheets.Add(Before:=Sheets("Indicadores")).Name = "Perspectivas"
'Seleciona a aba situa��o fincanceira
Sheets("Perspectivas").Select
'Escreve os primeiros trimestres da serie
Range("B1").Value = "11/01/13"
Range("C1").Value = "12/01/13"
'Completa a s�rie dos trimestres
Range("B1:C1").Select
    Selection.NumberFormat = "mmm-yy"
    Selection.AutoFill Destination:=Range(Cells(1, 2), Cells(1, coluna - 2)), Type:=xlFillDefault
'Atribui os nomes das s�ries
Range("A2").Value = "�ndice de Perspectivas da Pequena Empresa"
Range("A3").Value = "M�dia hist�rica"
'Copia e cola os valores da s�rie
Sheets("Indicadores").Select
Range(Cells(30, 4), Cells(30, coluna)).Copy
Sheets("Perspectivas").Select
Range("B2").PasteSpecial xlPasteValues
'Atribui os valores da m�dia
Range(Cells(3, 2), Cells(3, coluna - 2)).Value = m�dia
Range("A1").Select

'Adiciona o gr�fico
Set Gr�fico_Perspectivas = Sheets("Perspectivas").Shapes.AddChart2

Gr�fico_Perspectivas.Select ' Seleciona o Gr�fico
    '4/7 ActiveChart.ApplyChartTemplate ("C:\Users\paula.verlangeiro\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Panorama da Pequena Ind�stria\Automa��o\Perspectivas_PPI.crtx") ' Aplica o template do gr�fico
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Panorama da Pequena Ind�stria\Automa��o\Perspectivas_PPI.crtx") ' Aplica o template do gr�fico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gr�fico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gr�fico
    ActiveChart.Parent.Top = Parent.Range("B5").Top 'reposiciona o grafico em rela��o ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("B5").Left ' reposiciona o gr�fico em rela��o � borda esquerda da planilha
    ActiveChart.FullSeriesCollection(1).Name = "=Perspectivas!$A$3" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(1).Values = "=Perspectivas!" & Range(Cells(3, coluna - 98), Cells(3, coluna - 2)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(1).XValues = "=Perspectivas!" & Range(Cells(1, coluna - 98), Cells(1, coluna - 2)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
    ActiveChart.FullSeriesCollection(2).Name = "=Perspectivas!" & Cells(2, 1).Address 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(2).Values = "=Perspectivas!" & Range(Cells(2, coluna - 98), Cells(2, coluna - 2)).Address 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(2).XValues = "=Perspectivas!" & Range(Cells(1, coluna - 98), Cells(1, coluna - 2)).Address 'determina os valores referentes ao eixo x da s�rie adicionada
End Sub

Sub Principais_Problemas()
'Define as vari�veis
Dim coluna As Integer
Dim Gr�fico_Extrativa As Object
Dim Gr�fico_Transforma��o As Object
Dim Gr�fico_Constru��o As Object

'Seleciona a aba indicadores
Sheets("Principais Problemas").Select
'pega o n�mero da �ltima coluna
coluna = Range("C9").End(xlToRight).Column
'Adiciona a aba Situa��o Financiera
Sheets.Add(Before:=Sheets("Indicadores")).Name = "Principais Problemas Gr�fico"
'Seleciona a aba situa��o fincanceira
Sheets("Principais Problemas").Select
'Escreve os primeiros trimestres da serie
Range("C9").Value = "I-15"
Range("D9").Value = "II-15"
Range("E9").Value = "III-15"
Range("F9").Value = "IV-15"
'Completa a s�rie dos trimestres
Range("C9:F9").Select
    Selection.NumberFormat = "mmm-yy"
    Selection.AutoFill Destination:=Range(Cells(9, 3), Cells(9, coluna)), Type:=xlFillDefault
    
'Seleciona a aba situa��o fincanceira
Sheets("Principais Problemas Gr�fico").Select

'Atribui os nomes das s�ries
Range("A1").Value = "Extrativa"
Range("E1").Value = "Transforma��o"
Range("I1").Value = "Constru��o"


'Extrativa................................................................................................

'copia e cola Categorias sem outros e nenhum
Sheets("Principais Problemas").Select
Range(Cells(11, 2), Cells(26, 2)).Copy
Sheets("Principais Problemas Gr�fico").Select
Range("A3").PasteSpecial xlPasteValues

'copia e colaValores sem outros e nenhum
Sheets("Principais Problemas").Select
Range(Cells(11, coluna - 1), Cells(26, coluna)).Copy
Sheets("Principais Problemas Gr�fico").Select
Range("B3").PasteSpecial xlPasteValues

'copia e cola Data
Sheets("Principais Problemas").Select
Range(Cells(9, coluna - 1), Cells(9, coluna)).Copy
Sheets("Principais Problemas Gr�fico").Select
Range("B2").PasteSpecial xlPasteValues

'Ordenando os valores
ActiveSheet.Range("A2:C2").Select
Selection.AutoFilter
ActiveSheet.AutoFilter.Sort.SortFields. _
        Add2 Key:=Range("C2"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
ActiveSheet.Range("A2:C2").Select
Selection.AutoFilter
    
'copia e cola Categoria outros e nenhum
Sheets("Principais Problemas").Select
Range(Cells(27, 2), Cells(28, 2)).Copy
Sheets("Principais Problemas Gr�fico").Select
Range("A19").PasteSpecial xlPasteValues

'copia e cola Valores outros e nenhum
Sheets("Principais Problemas").Select
Range(Cells(27, coluna - 1), Cells(28, coluna)).Copy
Sheets("Principais Problemas Gr�fico").Select
Range("B19").PasteSpecial xlPasteValues
Range("B1").Select

'Adiciona o gr�fico
Set Gr�fico_Extrativa = Sheets("Principais Problemas Gr�fico").Shapes.AddChart2

Gr�fico_Extrativa.Select ' Seleciona o Gr�fico
    '5/7 ActiveChart.ApplyChartTemplate ("C:\Users\paula.verlangeiro\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Panorama da Pequena Ind�stria\Automa��o\Problemas_Extrativa_PPI.crtx") ' Aplica o template do gr�fico
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Panorama da Pequena Ind�stria\Automa��o\Problemas_Extrativa_PPI.crtx") ' Aplica o template do gr�fico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gr�fico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gr�fico
    ActiveChart.Parent.Top = Parent.Range("A21").Top 'reposiciona o grafico em rela��o ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("A21").Left ' reposiciona o gr�fico em rela��o � borda esquerda da planilha
    
    ActiveChart.FullSeriesCollection(1).Name = "='Principais Problemas Gr�fico'!$B$2" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(1).Values = "='Principais Problemas Gr�fico'!$B$3:$B$7" 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(1).XValues = "='Principais Problemas Gr�fico'!$A$3:$A$7"  'determina os valores referentes ao eixo x da s�rie adicionada
    
    ActiveChart.FullSeriesCollection(2).Name = "='Principais Problemas Gr�fico'!$C$2" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(2).Values = "='Principais Problemas Gr�fico'!$C$3:$C$7" 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(2).XValues = "='Principais Problemas Gr�fico'!$A$3:$A$7"  'determina os valores referentes ao eixo x da s�rie adicionada
  

'Transforma��o................................................................................................................................

'copia e cola Categorias sem outros e nenhum
Sheets("Principais Problemas").Select
Range("B33:B48").Copy
Sheets("Principais Problemas Gr�fico").Select
Range("E3").PasteSpecial xlPasteValues

'copia e cola Valores sem outros e nenhum
Sheets("Principais Problemas").Select
Range(Cells(33, coluna - 1), Cells(48, coluna)).Copy
Sheets("Principais Problemas Gr�fico").Select
Range("F3").PasteSpecial xlPasteValues

'copia e cola Data
Range("B2:C2").Copy
Range("F2").PasteSpecial xlPasteValues

'Ordenando os valores
ActiveSheet.Range("E2:G2").Select
Selection.AutoFilter
ActiveSheet.AutoFilter.Sort.SortFields. _
        Add2 Key:=Range("G2"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
ActiveSheet.Range("E2:G2").Select
Selection.AutoFilter

'copia e cola Categoria outros e nenhum
Range("A19:A20").Copy
Range("E19").PasteSpecial xlPasteValues

'copia e cola Valores outros e nenhum
Sheets("Principais Problemas").Select
Range(Cells(49, coluna - 1), Cells(50, coluna)).Copy
Sheets("Principais Problemas Gr�fico").Select
Range("F19").PasteSpecial xlPasteValues
Range("A1").Select

'Adiciona o gr�fico
Set Gr�fico_Transforma��o = Sheets("Principais Problemas Gr�fico").Shapes.AddChart2

Gr�fico_Transforma��o.Select ' Seleciona o Gr�fico
    '6/7 ActiveChart.ApplyChartTemplate ("C:\Users\paula.verlangeiro\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Panorama da Pequena Ind�stria\Automa��o\Problemas_Transforma��o_PPI.crtx") ' Aplica o template do gr�fico
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Panorama da Pequena Ind�stria\Automa��o\Problemas_Transforma��o_PPI.crtx") ' Aplica o template do gr�fico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gr�fico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gr�fico
    ActiveChart.Parent.Top = Parent.Range("E21").Top 'reposiciona o grafico em rela��o ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("E21").Left ' reposiciona o gr�fico em rela��o � borda esquerda da planilha
    
    ActiveChart.FullSeriesCollection(1).Name = "='Principais Problemas Gr�fico'!$F$2" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(1).Values = "='Principais Problemas Gr�fico'!$F$3:$F$7" 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(1).XValues = "='Principais Problemas Gr�fico'!$E$3:$E$7"  'determina os valores referentes ao eixo x da s�rie adicionada
    
    ActiveChart.FullSeriesCollection(2).Name = "='Principais Problemas Gr�fico'!$G$2" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(2).Values = "='Principais Problemas Gr�fico'!$G$3:$G$7" 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(2).XValues = "='Principais Problemas Gr�fico'!$E$3:$E$7"  'determina os valores referentes ao eixo x da s�rie adicionada


'Constru��o..............................................................................................................................

'copia e cola Categorias sem outros e nenhum
Sheets("Principais Problemas").Select
Range("B55:B72").Copy
Sheets("Principais Problemas Gr�fico").Select
Range("I3").PasteSpecial xlPasteValues

'copia e cola Valores sem outros e nenhum
Sheets("Principais Problemas").Select
Range(Cells(55, coluna - 1), Cells(72, coluna)).Copy
Sheets("Principais Problemas Gr�fico").Select
Range("J3").PasteSpecial xlPasteValues

'copia e cola Data
Range("B2:C2").Copy
Range("J2").PasteSpecial xlPasteValues

'Ordenando os valores
ActiveSheet.Range("I2:K2").Select
Selection.AutoFilter
ActiveSheet.AutoFilter.Sort.SortFields. _
        Add2 Key:=Range("K2"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
ActiveSheet.Range("I2:K2").Select
Selection.AutoFilter
    
'copia e cola Categoria outros e nenhum
Range("A19:A20").Copy (Sheets("Principais Problemas Gr�fico").Range("I19"))


'copia e cola Valores outros e nenhum
Sheets("Principais Problemas").Select
Range(Cells(73, coluna - 1), Cells(74, coluna)).Copy
Sheets("Principais Problemas Gr�fico").Select
Range("J19").PasteSpecial xlPasteValues
Range("A1").Select


'Adiociona o gr�fico
Set Gr�fico_Constru��o = Sheets("Principais Problemas Gr�fico").Shapes.AddChart2

Gr�fico_Constru��o.Select ' Seleciona o Gr�fico
    '7/7 ActiveChart.ApplyChartTemplate ("C:\Users\paula.verlangeiro\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Panorama da Pequena Ind�stria\Automa��o\Problemas_Constru��o_PPI.crtx") ' Aplica o template do gr�fico
    ActiveChart.ApplyChartTemplate ("C:\Users\e-gustavo.oliveira\CNI - Confedera��o Nacional da Ind�stria\ECON - 1 Indicadores Econ�micos CNI\1 Indicadores de Atividade Industrial\Panorama da Pequena Ind�stria\Automa��o\Problemas_Constru��o_PPI.crtx") ' Aplica o template do gr�fico
    ActiveChart.Parent.Height = 300 'ajusta a altura do gr�fico
    ActiveChart.Parent.Width = 425 ' ajusta a largura do gr�fico
    ActiveChart.Parent.Top = Parent.Range("I21").Top 'reposiciona o grafico em rela��o ao topo da planilha
    ActiveChart.Parent.Left = Parent.Range("I21").Left ' reposiciona o gr�fico em rela��o � borda esquerda da planilha
    
    ActiveChart.FullSeriesCollection(1).Name = "='Principais Problemas Gr�fico'!$J$2" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(1).Values = "='Principais Problemas Gr�fico'!$J$3:$J$7" 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(1).XValues = "='Principais Problemas Gr�fico'!$I$3:$I$7"  'determina os valores referentes ao eixo x da s�rie adicionada
    
    ActiveChart.FullSeriesCollection(2).Name = "='Principais Problemas Gr�fico'!$K$2" 'Determina o nome da s�rie
    ActiveChart.FullSeriesCollection(2).Values = "='Principais Problemas Gr�fico'!$K$3:$K$7" 'determina os valores da s�rie
    ActiveChart.FullSeriesCollection(2).XValues = "='Principais Problemas Gr�fico'!$I$3:$I$7"  'determina os valores referentes ao eixo x da s�rie adicionada

End Sub
