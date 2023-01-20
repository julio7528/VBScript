Function ApagarPasta(parametro)
    Dim fso, folder
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = parametro
    
    If fso.FolderExists(folder) Then
        fso.DeleteFolder folder, True
    End If
    End Function

call ApagarPasta("C:\Users\julio\Desktop\apagar")