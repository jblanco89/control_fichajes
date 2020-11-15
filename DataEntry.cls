VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Private Sub DelButton_Click()

'Aplicación botón para deshacer hoja "output"

Application.DisplayAlerts = False
For Each Sheet In Application.Worksheets
   If Sheet.name <> "DataEntry" Then
   Sheet.Delete
   End If
Next Sheet
Application.DisplayAlerts = True


End Sub

'Generador automático de fichajes

Private Sub GenButton_Click()
Dim hour, endDays, workDays, diff As Integer
Dim sch, jour, name, lastRow, chrc As String
Dim startDate, endDate, dateCnt As Variant
Dim total As Double

Application.ScreenUpdating = False

name = Sheets("DataEntry").Range("B1").Value 'nombre empleado
hour = Sheets("DataEntry").Range("B3").Value 'horas contratadas
sch = Sheets("DataEntry").Range("B5").Value 'tipo de horario (partido o completo)
jour = Sheets("DataEntry").Range("B7").Value 'jornada (lunes a viernes)
startDate = Sheets("DataEntry").Range("B9").Value 'Fecha de inicio
endDate = Sheets("DataEntry").Range("B11").Value 'Fecha final
chrc = Mid(name, 1, 9)

If sheetExt(chrc) = False Then

        
        Sheets.Add After:=Sheets(1)
        ActiveSheet.name = chrc
        
        Sheets(chrc).Range("A2:B2").Merge
        Sheets(chrc).Range("A2").Value = "Periodo"
        Sheets(chrc).Range("A2").HorizontalAlignment = xlCenter
        
        Sheets(chrc).Range("A3").Value = Sheets("DataEntry").Range("A9").Value
        Sheets(chrc).Range("B3").Value = startDate
        
        Sheets(chrc).Range("A4").Value = Sheets("DataEntry").Range("A11").Value
        Sheets(chrc).Range("B4").Value = endDate
        
        Sheets(chrc).Range("A6").Value = "Nombre Empleado"
        Sheets(chrc).Range("B6").Value = name
        
                If jour = "Lun-Vie" Then
                   startDate = DateValue(startDate)
                   endDate = DateValue(endDate)
                   wholeWeeks = DateDiff("w", startDate, endDate)
                   dateCnt = DateAdd("ww", wholeWeeks, startDate)
                   endDays = 0
            
                   Do While dateCnt <= endDate
                    If Format(dateCnt, "ddd") <> "Sun" And _
                    Format(dateCnt, "ddd") <> "Sat" Then
                    endDays = endDays + 1
                    End If
                   dateCnt = DateAdd("d", 1, dateCnt)
                   Loop
            
                   workDays = ((wholeWeeks * 5) + endDays)
                   
                   Sheets(chrc).Range("A7").Value = "Dias Trabajados"
                   
                   Sheets(chrc).Range("B7").Value = workDays
                
                End If
                
                If sch = "Completo" And hour = 8 Then
                
                Sheets(chrc).Range("A10").Value = "Entradas"
                Sheets(chrc).Range("B10").Value = "Salidas"
                Sheets(chrc).Range("C10").Value = "Horas"
                
                    For i = 1 To workDays
                        Sheets(chrc).Cells(i + 10, 1) = "9:00"
                        Sheets(chrc).Cells(i + 10, 2) = "17:00"
                        Sheets(chrc).Cells(i + 10, 3) = 8
                    Next i
                    
                ElseIf sch = "Partido" And hour = 8 Then
                Sheets(chrc).Range("A10").Value = "Entradas"
                Sheets(chrc).Range("B10").Value = "Salidas"
                Sheets(chrc).Range("C10").Value = "Horas"
                
                    For i = 1 To workDays * 2
                        Sheets(chrc).Cells(i + 10, 1) = "16:00"
                        Sheets(chrc).Cells(i + 10, 2) = "20:00"
                        Sheets(chrc).Cells(i + 10, 3) = 4
        
                    Next i
                    
                    For i = 1 To workDays * 2 Step 2
                        Sheets(chrc).Cells(i + 10, 1) = "10:00"
                        Sheets(chrc).Cells(i + 10, 2) = "14:00"
                        Sheets(chrc).Cells(i + 10, 3) = 4
        
                    Next i
                    
                ElseIf sch = "Completo" And hour = 4 Then
                Sheets(chrc).Range("A10").Value = "Entradas"
                Sheets(chrc).Range("B10").Value = "Salidas"
                Sheets(chrc).Range("C10").Value = "Horas"
                
                    For i = 1 To workDays
                        Sheets(chrc).Cells(i + 10, 1) = "16:00"
                        Sheets(chrc).Cells(i + 10, 2) = "20:00"
                        Sheets(chrc).Cells(i + 10, 3) = 4
                    Next i
                    
                    For i = 1 To workDays Step 3
                        Sheets(chrc).Cells(i + 10, 1) = "10:00"
                        Sheets(chrc).Cells(i + 10, 2) = "14:00"
                        Sheets(chrc).Cells(i + 10, 3) = 4
                    Next i
                    
                ElseIf sch = "Partido" And hour = 6 Then
                    Sheets(chrc).Range("A10").Value = "Entradas"
                    Sheets(chrc).Range("B10").Value = "Salidas"
                    Sheets(chrc).Range("C10").Value = "Horas"
                
                    For i = 1 To workDays * 2
                        Sheets(chrc).Cells(i + 10, 1) = "16:00"
                        Sheets(chrc).Cells(i + 10, 2) = "18:00"
                        Sheets(chrc).Cells(i + 10, 3) = 2
        
                    Next i
                    
                    For i = 1 To workDays * 2 Step 2
                        Sheets(chrc).Cells(i + 10, 1) = "10:00"
                        Sheets(chrc).Cells(i + 10, 2) = "14:00"
                        Sheets(chrc).Cells(i + 10, 3) = 4
        
                    Next i
                        
                
                End If
                
                lastRow = Sheets(chrc).Cells(Rows.Count, "C").End(xlUp).Row
                Set addition = Sheets(chrc).Range("C11:C" & lastRow)
                total = Application.WorksheetFunction.Sum(addition)
        
                Sheets(chrc).Cells(lastRow, 4) = "Total de Horas"
                Sheets(chrc).Cells(lastRow, 5) = total
                Sheets(chrc).Cells(lastRow, 5).Font.Bold = True
                Sheets(chrc).Cells(lastRow + 3, 1) = "Firma Empleado"
                Sheets(chrc).Cells(lastRow + 3, 1).Font.Bold = True
                Sheets(chrc).Cells(lastRow + 3, 2) = "______________________"
                Sheets(chrc).Cells(lastRow + 3, 3) = "Empresa"
                Sheets(chrc).Cells(lastRow + 3, 3).Font.Bold = True
                Sheets(chrc).Cells(lastRow + 3, 4) = "______________________"
                
                Sheets(chrc).Cells(lastRow + 3, 1).RowHeight = 25
                
                For i = 1 To 4
                Sheets(chrc).Columns(i).ColumnWidth = 22
                Next i
        
        
        
        
        
        'formato
        
        Sheets(chrc).Range("A2:B4").Select
            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            With Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
        Sheets(chrc).Range("A2:B2").Select
            Selection.Font.Bold = True
        Sheets(chrc).Range("A3:A4").Select
            Selection.Style = "Énfasis1"
            Selection.Font.Bold = True
        Sheets(chrc).Range("B3:B4").Select
            Selection.Style = "Salida"
        Sheets(chrc).Range("A3:B4").Select
            With Selection.Font
                .name = "Arial"
                .Size = 11
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = xlUnderlineStyleNone
                .TintAndShade = 0
                .ThemeFont = xlThemeFontNone
            End With
            Selection.Font.Size = 12
        Sheets(chrc).Range("A2:B2").Select
            Selection.Style = "Énfasis3"
            Selection.Style = "Énfasis5"
            Selection.Font.Bold = True
            Selection.Font.Size = 12
            With Selection.Font
                .name = "Arial"
                .Size = 12
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = xlUnderlineStyleNone
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .ThemeFont = xlThemeFontNone
            End With
        Sheets(chrc).Range("A2:B2").Select
            Selection.Style = "Título 2"
            Selection.Style = "Normal"
            Selection.Style = "Título 2"
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = True
            End With
            With Selection.Font
                .name = "Arial"
                .Size = 13
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = xlUnderlineStyleNone
                .ThemeColor = xlThemeColorLight2
                .TintAndShade = 0
                .ThemeFont = xlThemeFontNone
            End With
        Sheets(chrc).Range("A3:B4").Select
            Selection.Copy
        Sheets(chrc).Range("A6:A7").Select
            Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False
            Application.CutCopyMode = False
            
        Sheets(chrc).Range("B6:B7").Select
            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            With Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            Selection.Font.Bold = True
        Sheets(chrc).Range("B6:B7").Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = 0.799981688894314
                .PatternTintAndShade = 0
            End With
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent3
                .TintAndShade = 0.799981688894314
                .PatternTintAndShade = 0
            End With
            
            
        Sheets(chrc).Range("A10:C10").Select
            Selection.Style = "Título 3"
            Selection.Font.Size = 12
            With Selection.Font
                .name = "Arial"
                .Size = 12
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = xlUnderlineStyleNone
                .ThemeColor = xlThemeColorLight2
                .TintAndShade = 0
                .ThemeFont = xlThemeFontNone
            End With
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
        Sheets(chrc).Range("F7").Select
            ActiveWindow.SmallScroll Down:=6
        Sheets(chrc).Range("A11:C" & lastRow).Select
            Application.CutCopyMode = False
            Application.CutCopyMode = False
        Sheets(chrc).Range("A10:C" & lastRow).Select
            Application.CutCopyMode = False
            
            
            
        Sheets(chrc).Range("A10:C" & lastRow).Select
            ActiveWindow.SmallScroll Down:=6
            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            With Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            ActiveWindow.SmallScroll Down:=-15
            
            
        Sheets(chrc).Cells(lastRow, 4).Select
            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            With Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            
            
        Sheets(chrc).Cells(lastRow, 4).Select
            ActiveWindow.SmallScroll Down:=-3
            Selection.Style = "Énfasis1"
            With Selection.Font
                .name = "Arial"
                .Size = 11
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = xlUnderlineStyleNone
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .ThemeFont = xlThemeFontNone
            End With
            Selection.Font.Size = 12
            Selection.Font.Bold = True
        Sheets(chrc).Cells(lastRow, 5).Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent3
                .TintAndShade = 0.799981688894314
                .PatternTintAndShade = 0
            End With
            
        ActiveWindow.DisplayGridlines = False
        'Application.ScreenUpdating = True
    Else
    MsgBox "El fichaje de " & name & " existe, reinicie el sistema antes de continuar", , "ERROR"
    
    End If
    

End Sub

Private Sub printPDF_Click()
'Función que genera PDF
Dim fName As String
Dim directFolder As String
Dim chrc As String
Dim name As String
Dim currentMonth, startDate, currentYear, newDate

name = Sheets("DataEntry").Range("B1").Value
chrc = Mid(name, 1, 9)

If sheetExt(chrc) = True Then
    startDate = Sheets("DataEntry").Range("B9").Value
    newDate = Format(startDate, "dd/mm/yyyy")
    currentYear = Year(newDate)
    currentMonth = Month(newDate)
    Application.ScreenUpdating = False
    
    
    For Each Sheet In Worksheets
        If Sheet.name <> "DataEntry" Then
        Sheet.PageSetup.LeftHeader = "&B CONTROL DE ENTRADAS Y SALIDAS &B"
        Sheet.PageSetup.RightHeader = "&I Página &P de &N &I"
        Sheet.PageSetup.Orientation = xlLandscape
        fName = Sheet.name
        directFolder = ActiveWorkbook.Path & "\" & fName & "_" & currentMonth & "_" & currentYear & ".pdf"
        Sheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=directFolder, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, OpenAfterPublish:=False
        End If
    Next Sheet
Else
MsgBox "Genere un fichaje antes de imprimir", , "ERROR"

End If

End Sub
