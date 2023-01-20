Function CriarNovaPlanilha(parametro)
Dim objExcel, objWorkbook
Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Add
objWorkbook.SaveAs parametro
objWorkbook.Close
objExcel.Quit
End Function

Call CriarNovaPlanilha("C:\Users\julio\Desktop\python\vscodeandgit\vscodeandgit\M500 final.xlsx")