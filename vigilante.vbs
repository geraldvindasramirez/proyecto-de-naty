Set fso = CreateObject("Scripting.FileSystemObject")
folderPath = fso.GetParentFolderName(WScript.ScriptFullName)
filePath = folderPath & "\datos.xlsx"
excelPath = folderPath & "\programa.xlsm"

' Obtener la fecha inicial del archivo si ya existe
lastModified = ""
If fso.FileExists(filePath) Then
    lastModified = fso.GetFile(filePath).DateLastModified
End If

' Bucle silencioso que revisa la carpeta cada segundo
Do
    WScript.Sleep 1000 ' Pausa de 1 segundo (consumo de CPU 0%)
    
    If fso.FileExists(filePath) Then
        currentModified = fso.GetFile(filePath).DateLastModified
        
        ' Si detecta que datos.xlsx fue modificado o creado recientemente
        If currentModified <> lastModified Then
            lastModified = currentModified
            
            ' Espera un segundo a que termine de escribirse el archivo por completo
            WScript.Sleep 1000
            
            ' Abre automáticamente el programa de macros de Excel
            Dim shell
            Set shell = CreateObject("WScript.Shell")
            shell.Run "excel.exe """ & excelPath & """", 1, False
            Set shell = Nothing
        End If
    End If
Loop