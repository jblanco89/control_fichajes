Attribute VB_Name = "Funciones"


Function workDays(begDate As Variant, endDate As Variant) As Integer
'Función que determina los días laborales el mes
 
 Dim wholeWeeks As Variant
 Dim dateCnt As Variant
 Dim endDays As Integer
 
 On Error GoTo Err_workDays
 
 begDate = DateValue(begDate)
 endDate = DateValue(endDate)
 wholeWeeks = DateDiff("w", begDate, endDate)
 dateCnt = DateAdd("ww", wholeWeeks, begDate)
 endDays = 0
 
 Do While dateCnt <= endDate
 If Format(dateCnt, "ddd") <> "Sun" And _
 Format(dateCnt, "ddd") <> "Sat" Then
 endDays = endDays + 1
 End If
 dateCnt = DateAdd("d", 1, dateCnt)
 Loop
 
 workDays = wholeWeeks * 5 + endDays
 
Exit Function
 
Err_workDays:
 
 
 If Err.Number = 94 Then
 workDays = 0
 Exit Function
 Else
 MsgBox "Error " & Err.Number & ": " & Err.Description
 End If
 
End Function

Function sheetExt(chrc As String) As Boolean
'Función que determina si una hoja existe o no
On Error Resume Next
sheetExt = (Sheets(chrc).name <> "")
End Function


