Attribute VB_Name = "Módulo1"
Sub formatação()

UL = Range("B3").End(xlDown).Row
Range(Cells(1, 11), Cells(UL, 14)).Clear
Range("1:1").ClearContents


Range(Cells(1, 4), Cells(UL, 4)).Cut
Range("C120").Select
ActiveSheet.Paste

Range(Cells(1, 5), Cells(UL, 5)).Cut
Range("E120").Select
ActiveSheet.Paste

Range(Cells(1, 6), Cells(UL, 6)).Cut
Range("G120").Select
ActiveSheet.Paste

Range(Cells(1, 7), Cells(UL, 7)).Cut
Range("D1").Select
ActiveSheet.Paste

Range(Cells(1, 8), Cells(UL, 8)).Cut
Range("D120").Select
ActiveSheet.Paste

Range(Cells(1, 9), Cells(UL, 9)).Cut
Range("F120").Select
ActiveSheet.Paste

Range(Cells(120, 3), Cells(120 + UL - 1, 7)).Cut
Range("E1").Select
ActiveSheet.Paste

Columns("A:A").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

Rows("1:1").Select
Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
Rows("1:1").Select
Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
Rows("1:1").Select
Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
Rows("1:1").Select
Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
Rows("1:1").Select
Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
Rows("1:1").Select
Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove


Range("B2").Value = ""
Range("B3").Value = "Distrito Federal, 2022."
Range("B2").Font.Bold = True
Range("B3").Font.Bold = True

Range("B5:C7").Merge

Range("D5:K5").Merge
Range("D5:K5").Value = "Região"
Range("D5:K5").Font.Bold = True

Range("D6:E6").Merge
Range("D6:M6").Font.Bold = True

Range("D6:E6").Copy
Range("F6:G6").Select
ActiveSheet.Paste

Range("H6:I6").Select
ActiveSheet.Paste

Range("J6:K6").Select
ActiveSheet.Paste

Application.CutCopyMode = False

Range("D6:E6").Value = "Alta"
Range("F6:G6").Value = "Média-Alta"
Range("H6:I6").Value = "Média-Baixa"
Range("J6:K6").Value = "Baixa"

Range("D7:E7").Font.Bold = True
Range("D7").Value = "Nº Crianças"
Range("E7").Value = "Percentual (%)"

Range("D7:E7").Copy
Range("F7:G7").Select
ActiveSheet.Paste

Range("H7:I7").Select
ActiveSheet.Paste

Range("J7:K7").Select
ActiveSheet.Paste

Range("L7:M7").Select
ActiveSheet.Paste

Application.CutCopyMode = False

Range("L5:M6").Clear
Range("L5:M6").Select
Selection.Merge
Selection.Value = "DF"
Selection.Font.Bold = True

UL = Range("C8").End(xlDown).Row

Range("B" & UL + 1 & ":C" & UL + 1 & "").Merge

Cells(UL + 1, 2).Value = "Total"

Cells(UL + 1, 2).Font.Bold = True

Cells(UL + 1, 4).Formula = "=SUM(D8:D" & UL & ")"

Cells(UL + 1, 4).AutoFill Destination:=Range(Cells(UL + 1, 4), Cells(UL + 1, 13)), Type:=xlFillDefault

Range(Cells(UL + 1, 4), Cells(UL + 1, 13)).Font.Bold = True

Nlinhadados = 8 - UL
linhaagr = 8

Do Until Nlinhadados = -1

Cells(linhaagr, 12).Formula = "=SUM(D" & linhaagr & "+ F" & linhaagr & "+ H" & linhaagr & "+ J" & linhaagr & ")"

linhaagr = linhaagr + 1

Nlinhadados = UL - linhaagr

Loop


Nlinhadados = 8 - UL
linhaagr = 8

Do Until Nlinhadados = -1

Cells(linhaagr, 13).Value = "=L" & linhaagr & "/$L$" & UL + 1
linhaagr = linhaagr + 1

Nlinhadados = UL - linhaagr

Loop

UL = UL + 1

Range("D5:K6").HorizontalAlignment = xlCenter
Range("L5:M6").HorizontalAlignment = xlCenter
Range("L5:M6").VerticalAlignment = xlBottom
Range("D7:M7").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With


Range(Cells(8, 3 + 1), Cells(UL, 3 + 1)).NumberFormat = "#,##0"
Range(Cells(8, 5 + 1), Cells(UL, 5 + 1)).NumberFormat = "#,##0"
Range(Cells(8, 7 + 1), Cells(UL, 7 + 1)).NumberFormat = "#,##0"
Range(Cells(8, 9 + 1), Cells(UL, 9 + 1)).NumberFormat = "#,##0"
Range(Cells(8, 11 + 1), Cells(UL, 11 + 1)).NumberFormat = "#,##0"

Range(Cells(8, 4 + 1), Cells(UL, 4 + 1)).Style = "Percent"
Range(Cells(8, 4 + 1), Cells(UL, 4 + 1)).NumberFormat = "0.0%"

Range(Cells(8, 6 + 1), Cells(UL, 6 + 1)).Style = "Percent"
Range(Cells(8, 6 + 1), Cells(UL, 6 + 1)).NumberFormat = "0.0%"

Range(Cells(8, 8 + 1), Cells(UL, 8 + 1)).Style = "Percent"
Range(Cells(8, 8 + 1), Cells(UL, 8 + 1)).NumberFormat = "0.0%"

Range(Cells(8, 10 + 1), Cells(UL, 10 + 1)).Style = "Percent"
Range(Cells(8, 10 + 1), Cells(UL, 10 + 1)).NumberFormat = "0.0%"

Range(Cells(8, 12 + 1), Cells(UL, 12 + 1)).Style = "Percent"
Range(Cells(8, 12 + 1), Cells(UL, 12 + 1)).NumberFormat = "0.0%"


Range(Cells(6, 3 + 1), Cells(UL, 4 + 1)).Select
     With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
    
Range(Cells(6, 7 + 1), Cells(UL, 8 + 1)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
    
Range(Cells(6, 11 + 1), Cells(UL, 12 + 1)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With

Range(Cells(5, 2), Cells(UL, 13)).Select
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
Range(Cells(UL, 2), Cells(UL, 13)).Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
  
Range("B5:M6").Select
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    
Cells(UL + 1, 2).Value = "Fonte: IPE DF Codeplan. Pesquisa sobre desenvolvimento infantil e parentalidades (DiP)."

nome = ThisWorkbook.Name
If InStr(nome, ".") > 0 Then
   nome = Left(nome, InStr(nome, ".xlsx") - 1)
End If

diretorio = ThisWorkbook.Path & "\"

ThisWorkbook.SaveAs diretorio & nome & "_formatada.xlsm", FileFormat:=52

End Sub
