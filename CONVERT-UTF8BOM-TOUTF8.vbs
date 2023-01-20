Function ConverterUtf8BomToUtf8(parametros)
    SepararParametros  = Split(parametros, "|")    
    inputFile = SepararParametros(0)
    outputFile = SepararParametros(1)

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set inputStream = fso.OpenTextFile(inputFile, 1)
    contents = inputStream.ReadAll
    inputStream.Close
    contents = Mid(contents, 4)
    Set outputStream = fso.CreateTextFile(outputFile, True)
    outputStream.Write contents
    outputStream.Close    
end Function

file1 = "012021.txt"
file2 = "0012021.txt"
parametro = file1 & "|" & file2
msgbox parametro
call ConverterUtf8BomToUtf8(parametro)
