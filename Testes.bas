Attribute VB_Name = "Testes"
Option Explicit
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" ( _
    ByVal pCaller As LongPtr, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As LongPtr, _
    ByVal lpfnCB As LongPtr) As LongPtr

Sub Web_Scraping()

  Dim Internet_Explorer As InternetExplorer
  Set Internet_Explorer = New InternetExplorer
  Internet_Explorer.Visible = True
  
Internet_Explorer.Navigate "https://www3.bcb.gov.br/sgspub/localizarseries/localizarSeries.do?method=prepararTelaLocalizarSeries"

Application.Wait

SendKeys "Enter"

End Sub

Sub exemplo()

Dim IE As Object
Dim Data As Object

Set IE = CreateObject("InternetExplorer.Application")

IE.Visible = True

With IE
    .Navigate ("https://www3.bcb.gov.br/sgspub/localizarseries/localizarSeries.do?method=prepararTelaLocalizarSeries")
    Application.Wait (Now + TimeValue("00:00:05"))
    Application.SendKeys "~"
    Application.Wait (Now + TimeValue("00:00:05"))
    Application.SendKeys "10813"
    Application.SendKeys "~"
        
        
End With
     
End Sub




Sub cambio2()

Dim xmlhttp As New MSXML2.XMLHTTP60

xmlhttp.Open "GET", myURL, False
xmlhttp.send

Dim H As Long
Dim PictureToSave() As Byte
Dim FileName As String

H = FreeFile
FileName = "filepath"
PictureToSave() = xmlhttp.responseBody

Open FileName For Binary As #H
Put #H, 1, PictureToSave()
Close #H

End Sub

Sub teste_vinicius()
Dim Setor_Externo As Workbook 'Define a planilha do setor externo
Dim Dado As Workbook ' Define a planilha com o dado atualizado
Dim FileURL As String 'Link para baixar o dado
Dim DestinationFile As String 'Caminho na nuvem que download será realizado
Dim linha_dado As Integer ' ultima linha com dado da planilha com o dado atualizado
Dim linha_SE As Integer 'ultima linha com dado da planilha do setor externo
Dim nome As String 'nome do usuário (parte do e-mail da cni antes do @)

'captura o nome do usuário que será usado para definir o caminho na nuvem para o download
nome = VBA.Interaction.Environ$("UserName")

'Define o link de onde o arquivo será baixado
FileURL = "https://balanca.economia.gov.br/balanca/semanal/Tabela_Resumo.xlsx"

'Define o caminho que o download será realizado
DestinationFile = "C:\Users\" & nome & "\CNI - Confederação Nacional da Indústria\ECON - 2 Informe Conjuntural\Setor Externo\Base de Dados - Setor Externo\BC semanal\Tabela_Resumo" & Format(Date, "ddmmmyy") & ".xlsx"

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


End Sub
Sub Macro1()
'
' Macro1 Macro
'

'
    Range("B9").Select
    ActiveCell.FormulaR1C1 = "='BP mensal'!RC/1000"
    Range("B10").Select
End Sub



Sub IP_IQ()

Dim FileURL As String
Dim DestinationFile As String
Dim Setor_Externo As Workbook
Dim Dado As Workbook
Dim Intervalo_Dados As Double
Dim Nome_aba As String
Dim ultima_linha As Integer
Dim linha_tabela As Integer
 
 
FileURL = "https://balanca.economia.gov.br/balanca/IPQ/arquivos/Dados_totais_mensal.csv"

DestinationFile = "C:\Users\e-gustavo.oliveira\CNI - Confederação Nacional da Indústria\ECON - 2 Análise Conjuntural\2 Informe Conjuntural\Setor Externo\Base de Dados - Setor Externo\IP e IQ\TOTAL\totais" & Format(Date, "ddmmmyy") & ".xls"

If URLDownloadToFile(0, FileURL, DestinationFile, 0, 0) = 0 Then
    Debug.Print "File download started"
Else
    Debug.Print "File download not started"
End If

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
        TableDestination:="Sheet1!R3C1", TableName:="PivotTable1", DefaultVersion _
        :=7
        
    Sheets("Sheet1").Select
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
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CO_ANO")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CO_MES")
        .Orientation = xlRowField
        .Position = 4
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("NO_CLASSIFICACAO")
        .Orientation = xlRowField
        .Position = 5
    End With
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("INDICE"), "Sum of INDICE", xlSum
        
        
'Formatando para forma tabular e repetir em cada linha******************************************************************************************************************************
          
          
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO")
        .LayoutForm = xlTabular
        .RepeatLabels = True
    End With
    
   
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE")
        .LayoutForm = xlTabular
        .RepeatLabels = True
    End With
  
  
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CO_ANO")
        .LayoutForm = xlTabular
        .RepeatLabels = True
    End With
   
   
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CO_MES")
        .LayoutForm = xlTabular
        .RepeatLabels = True
    End With
    
    ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    
     ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE").Subtotals = Array( _
    False, False, False, False, False, False, False, False, False, False, False, False)
    
    ActiveSheet.PivotTables("PivotTable1").PivotFields("CO_MES").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
   
    ActiveSheet.PivotTables("PivotTable1").PivotFields("CO_ANO").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    

    ActiveSheet.PivotTables("PivotTable1").ColumnGrand = False
    
'Filtro ********************************************************************************************************************************

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO")
        .PivotItems("EXP_DESSAZONALIZADA").Visible = False
        .PivotItems("IMP").Visible = False
        .PivotItems("IMP_DESSAZONALIZADA").Visible = False
        .PivotItems("TERMOS_DE_TROCA").Visible = False
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE")
        .PivotItems("QUANTUM").Visible = False
    End With
    
'Copia e cola ***************************************************************************************************************************

linha_tabela = Range("A3").End(xlDown).Row
   
Range(Cells(4, 1), Cells(6, linha_tabela)).Copy

Setor_Externo.Activate
Sheets("IP e IQ - ME").Select
Range("B9").PasteSpecial xlPasteAll

'Filtro ********************************************************************************************************************************

Dado.Activate

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE")
        .PivotItems("QUANTUM").Visible = True
        .PivotItems("PRECO").Visible = False
    End With
    
'Copia e cola ***************************************************************************************************************************

Range(Cells(4, 1), Cells(6, linha_tabela)).Copy

Setor_Externo.Activate
Sheets("IP e IQ - ME").Select
Range("I9").PasteSpecial xlPasteAll

'Filtro ********************************************************************************************************************************
Dado.Activate

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO")
        .PivotItems("EXP_DESSAZONALIZADA").Visible = False
        .PivotItems("IMP").Visible = True
        .PivotItems("EXP").Visible = False
        .PivotItems("IMP_DESSAZONALIZADA").Visible = False
        .PivotItems("TERMOS_DE_TROCA").Visible = False
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE")
        .PivotItems("QUANTUM").Visible = False
        .PivotItems("PRECO").Visible = True
    End With
    
'Copia e cola ***************************************************************************************************************************

Range(Cells(4, 1), Cells(6, linha_tabela)).Copy

Setor_Externo.Activate
Sheets("IP e IQ - ME").Select
Range("P9").PasteSpecial xlPasteAll


'Filtro ********************************************************************************************************************************
Dado.Activate

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO")
        .PivotItems("EXP_DESSAZONALIZADA").Visible = False
        .PivotItems("IMP").Visible = True
        .PivotItems("EXP").Visible = False
        .PivotItems("IMP_DESSAZONALIZADA").Visible = False
        .PivotItems("TERMOS_DE_TROCA").Visible = False
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE")
        .PivotItems("QUANTUM").Visible = True
        .PivotItems("PRECO").Visible = False
    End With
    
'Copia e cola ***************************************************************************************************************************

Range(Cells(4, 1), Cells(6, linha_tabela)).Copy

Setor_Externo.Activate
Sheets("IP e IQ - ME").Select
Range("W9").PasteSpecial xlPasteAll


'Filtro ********************************************************************************************************************************

Dado.Activate

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO")
        .PivotItems("EXP_DESSAZONALIZADA").Visible = True
        .PivotItems("IMP").Visible = False
        .PivotItems("EXP").Visible = False
        .PivotItems("IMP_DESSAZONALIZADA").Visible = False
        .PivotItems("TERMOS_DE_TROCA").Visible = False
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE")
        .PivotItems("QUANTUM").Visible = True
        .PivotItems("PRECO").Visible = False
    End With
    
'Copia e cola ***************************************************************************************************************************

Range(Cells(4, 1), Cells(6, linha_tabela)).Copy

Setor_Externo.Activate
Sheets("IP e IQ - ME").Select
Range("AD9").PasteSpecial xlPasteAll

'Filtro ********************************************************************************************************************************

Dado.Activate

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO")
        .PivotItems("EXP_DESSAZONALIZADA").Visible = False
        .PivotItems("IMP").Visible = False
        .PivotItems("EXP").Visible = False
        .PivotItems("IMP_DESSAZONALIZADA").Visible = True
        .PivotItems("TERMOS_DE_TROCA").Visible = False
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("TIPO_INDICE")
        .PivotItems("QUANTUM").Visible = True
        .PivotItems("PRECO").Visible = False
    End With
    
'Copia e cola ***************************************************************************************************************************

Range(Cells(4, 1), Cells(6, linha_tabela)).Copy

Setor_Externo.Activate
Sheets("IP e IQ - ME").Select
Range("AK9").PasteSpecial xlPasteAll

End Sub


Sub user()

Sheets("IP IQ - GCE").Range("F1").Value = Application.UserName

End Sub


Sub GetUserName_Environ()
    Dim idx As Integer
    'To Directly the value of a Environment Variable with its Name
    MsgBox VBA.Interaction.Environ$("UserName")
End Sub


