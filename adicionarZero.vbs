Function EscreverFormulas(vParametro)
	Set wExcel = CreateObject("Excel.Application")
	Set wBook = wExcel.Workbooks.Open(vParametro)
	wExcel.DisplayAlerts = False
	wExcel.Visible = True
	Set wAba = wExcel.Worksheets("M500")
	lastRowM500 = wAba.Range("A" & wAba.Rows.Count).End(-4162).Row
	For i = 3 To lastRowM500
		value = wExcel.ActiveSheet.Range("D" & i).Value
		If Len(value) <= 2 Then
			value = Right("000" & value, 3)
			wExcel.ActiveSheet.Range("D" & i).Value = value
		End If
	Next
	wBook.Save
	wBook.Close
	wExcel.Quit
End Function

Dim caminhoArquivo
caminhoArquivo = "C:\Users\julio\Desktop\M500_Conferencia.xlsx"
Call EscreverFormulas(caminhoArquivo)