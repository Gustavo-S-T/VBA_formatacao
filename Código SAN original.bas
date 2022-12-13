Attribute VB_Name = "Módulo1"
Sub formatação()

UL = Range("A1").End(xlDown).Row
Range(Cells(1, 10), Cells(UL, 13)).Clear
Range("1:1").ClearContents


Range(Cells(1, 3), Cells(UL, 3)).Cut
Range("C20").Select
ActiveSheet.Paste

Range(Cells(1, 4), Cells(UL, 4)).Cut
Range("E20").Select
ActiveSheet.Paste

Range(Cells(1, 5), Cells(UL, 5)).Cut
Range("G20").Select
ActiveSheet.Paste

Range(Cells(1, 6), Cells(UL, 6)).Cut
Range("C1").Select
ActiveSheet.Paste

Range(Cells(1, 7), Cells(UL, 7)).Cut
Range("D20").Select
ActiveSheet.Paste

Range(Cells(1, 8), Cells(UL, 8)).Cut
Range("F20").Select
ActiveSheet.Paste

Range(Cells(20, 3), Cells(20 + UL - 1, 7)).Cut
Range("D1").Select
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
    
Rows("6:6").Select
Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
Rows("6:6").Select
Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove


Range("B2").Value = ""
Range("B3").Value = "Distrito Federal, 2022."
Range("B2").Font.Bold = True
Range("B3").Font.Bold = True

Range("B5:B7").Merge

Range("C5:J5").Merge
Range("C5:J5").Value = "Região"
Range("C5:J5").Font.Bold = True

Range("C6:D6").Merge
Range("C6:J6").Font.Bold = True

Range("C6:D6").Copy
Range("E6:F6").Select
ActiveSheet.Paste

Range("G6:H6").Select
ActiveSheet.Paste

Range("I6:J6").Select
ActiveSheet.Paste

Application.CutCopyMode = False

Range("C6:D6").Value = "Alta"
Range("E6:F6").Value = "Média-Alta"
Range("G6:H6").Value = "Média-Baixa"
Range("I6:J6").Value = "Baixa"

Range("C7:D7").Font.Bold = True
Range("C7").Value = "Nº Crianças"
Range("D7").Value = "Percentual (%)"

Range("C7:D7").Copy
Range("E7:F7").Select
ActiveSheet.Paste

Range("G7:H7").Select
ActiveSheet.Paste

Range("I7:J7").Select
ActiveSheet.Paste

Range("K7:L7").Select
ActiveSheet.Paste

Application.CutCopyMode = False

Range("K5:L6").Clear
Range("K5:L6").Select
Selection.Merge
Selection.Value = "DF"
Selection.Font.Bold = True

UL = Range("B8").End(xlDown).Row

Cells(UL + 1, 2).Value = "Total"

Cells(UL + 1, 2).Font.Bold = True

Cells(UL + 1, 3).Formula = "=SUM(C8:C" & UL & ")"

Cells(UL + 1, 3).AutoFill Destination:=Range(Cells(UL + 1, 3), Cells(UL + 1, 12)), Type:=xlFillDefault

Range(Cells(UL + 1, 3), Cells(UL + 1, 12)).Font.Bold = True

Nlinhadados = 8 - UL
linhaagr = 8

Do Until Nlinhadados = -1

Cells(linhaagr, 11).Formula = "=SUM(C" & linhaagr & "+ E" & linhaagr & "+ G" & linhaagr & "+ I" & linhaagr & ")"

linhaagr = linhaagr + 1

Nlinhadados = UL - linhaagr

Loop


Nlinhadados = 8 - UL
linhaagr = 8

Do Until Nlinhadados = -1

Cells(linhaagr, 12).Value = "=K" & linhaagr & "/$K$" & UL + 1
linhaagr = linhaagr + 1

Nlinhadados = UL - linhaagr

Loop

UL = UL + 1

Range("C5:J5").HorizontalAlignment = xlCenter
Range("C6:J6").HorizontalAlignment = xlCenter
Range("K5:L6").HorizontalAlignment = xlCenter
Range("K5:L6").VerticalAlignment = xlBottom
Range("C7:L7").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With


Range(Cells(8, 3), Cells(UL, 3)).NumberFormat = "#,##0"
Range(Cells(8, 5), Cells(UL, 5)).NumberFormat = "#,##0"
Range(Cells(8, 7), Cells(UL, 7)).NumberFormat = "#,##0"
Range(Cells(8, 9), Cells(UL, 9)).NumberFormat = "#,##0"
Range(Cells(8, 11), Cells(UL, 11)).NumberFormat = "#,##0"

Range(Cells(8, 4), Cells(UL, 4)).Style = "Percent"
Range(Cells(8, 4), Cells(UL, 4)).NumberFormat = "0.0%"

Range(Cells(8, 6), Cells(UL, 6)).Style = "Percent"
Range(Cells(8, 6), Cells(UL, 6)).NumberFormat = "0.0%"

Range(Cells(8, 8), Cells(UL, 8)).Style = "Percent"
Range(Cells(8, 8), Cells(UL, 8)).NumberFormat = "0.0%"

Range(Cells(8, 10), Cells(UL, 10)).Style = "Percent"
Range(Cells(8, 10), Cells(UL, 10)).NumberFormat = "0.0%"

Range(Cells(8, 12), Cells(UL, 12)).Style = "Percent"
Range(Cells(8, 12), Cells(UL, 12)).NumberFormat = "0.0%"


Range(Cells(6, 3), Cells(UL, 4)).Select
     With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
    
Range(Cells(6, 7), Cells(UL, 8)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
    
Range(Cells(6, 11), Cells(UL, 12)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With

Range(Cells(5, 2), Cells(UL, 12)).Select
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
Range(Cells(UL, 2), Cells(UL, 12)).Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
  
Range("K5:L6").Select
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
Range("C5:J5").Select
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
Range("B5:B7").Select
    Selection.Merge
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
Cells(UL + 1, 2).Value = "Fonte: IPE DF Codeplan. Pesquisa sobre desenvolvimento infantil e parentalidades (DiP)."

nome = ThisWorkbook.Name
If InStr(nome, ".") > 0 Then
   nome = Left(nome, InStr(nome, ".csv") - 1)
End If

diretorio = ThisWorkbook.Path & "\"

ThisWorkbook.SaveAs diretorio & nome & "_formatada.xlsm", FileFormat:=52

End Sub

