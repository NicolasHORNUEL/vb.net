Imports System.IO
Imports System.Drawing.Printing
Imports System.Threading
Imports System.Printing

Public Class Form1

    Dim filePathDesktop As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
    Dim folderPath As String = My.Computer.FileSystem.SpecialDirectories.Desktop & "\HOTFOLDER"
    Dim sProcess As String = "AcroRd32.exe"

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Not PrinterSettings.InstalledPrinters.Cast(Of String)().Any(Function(name) "SHARP_A3_jaune" = name) Then
            UsingShellExecute(szPrinter:="/c Cscript %WINDIR%\System32\Printing_Admin_Scripts\fr-FR\prnport.vbs -a -r 192.168.1.200_jaune -h 192.168.1.200", szDocumentPath:="")
            UsingShellExecute(szPrinter:="/c Cscript %WINDIR%\System32\Printing_Admin_Scripts\fr-FR\Prnport.vbs -t -r 192.168.1.200_jaune -md", szDocumentPath:="")
            UsingShellExecute(szPrinter:="/c rundll32 %WINDIR%\SysWOW64\printui.dll,PrintUIEntry /if /Z /b ""SHARP_A3_jaune"" /f Z:\UTILISATEURS\LOGICIELS\SHARP_Drivers\MX_D25_PCL6_PS_1505a_French_64bit\French\PCL6\64bit\ss0emfra.inf /r ""192.168.1.200_jaune"" /m ""SHARP MX-2614N PCL6"" ", szDocumentPath:="")
            Do Until PrinterSettings.InstalledPrinters.Cast(Of String)().Any(Function(name) "SHARP_A3_jaune" = name)
                Thread.Sleep(1000)
            Loop
            UsingShellExecute(szPrinter:="/c rundll32 %WINDIR%\SysWOW64\printui.dll,PrintUIEntry /Sr /n ""SHARP_A3_jaune"" /a ""Z:\UTILISATEURS\LOGICIELS\SHARP_Drivers\config_jaune.dat"" d g u r", szDocumentPath:="")
            Invoke(New MethodInvoker(Sub() TextBox1.AppendText("Imprimante SHARP_A3_jaune créée avec succès" & vbNewLine)))
        End If
        If Not PrinterSettings.InstalledPrinters.Cast(Of String)().Any(Function(name) "SHARP_A3_blanc" = name) Then
            UsingShellExecute(szPrinter:="/c Cscript %WINDIR%\System32\Printing_Admin_Scripts\fr-FR\prnport.vbs -a -r 192.168.1.200_blanc -h 192.168.1.200", szDocumentPath:="")
            UsingShellExecute(szPrinter:="/c Cscript %WINDIR%\System32\Printing_Admin_Scripts\fr-FR\Prnport.vbs -t -r 192.168.1.200_blanc -md", szDocumentPath:="")
            UsingShellExecute(szPrinter:="/c rundll32 %WINDIR%\SysWOW64\printui.dll,PrintUIEntry /if /Z /b ""SHARP_A3_blanc"" /f Z:\UTILISATEURS\LOGICIELS\SHARP_Drivers\MX_D25_PCL6_PS_1505a_French_64bit\French\PCL6\64bit\ss0emfra.inf /r ""192.168.1.200_blanc"" /m ""SHARP MX-2614N PCL6"" ", szDocumentPath:="")
            Do Until PrinterSettings.InstalledPrinters.Cast(Of String)().Any(Function(name) "SHARP_A3_blanc" = name)
                Thread.Sleep(1000)
            Loop
            UsingShellExecute(szPrinter:="/c rundll32 %WINDIR%\SysWOW64\printui.dll,PrintUIEntry /Sr /n ""SHARP_A3_blanc"" /a ""Z:\UTILISATEURS\LOGICIELS\SHARP_Drivers\config_blanc.dat"" d g u r", szDocumentPath:="")
            Invoke(New MethodInvoker(Sub() TextBox1.AppendText("Imprimante SHARP_A3_blanc créée avec succès" & vbNewLine)))
        End If
        If Not Directory.Exists(folderPath) Then
            Directory.CreateDirectory(folderPath)
            Do Until Directory.Exists(folderPath)
                Thread.Sleep(1000)
            Loop
            Invoke(New MethodInvoker(Sub() TextBox1.AppendText("Dossier \Desktop\HOTFOLDER créé avec succès" & vbNewLine)))
        End If
        FileSystemWatcher1.Path = folderPath
    End Sub



    Private Sub UsingShellExecute(ByVal szPrinter As String, ByVal szDocumentPath As String)
        Dim printProcess As Process = New Process
        printProcess.StartInfo.FileName = "cmd.exe"
        printProcess.StartInfo.Arguments = (szPrinter & """" & szDocumentPath & """")
        printProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
        printProcess.Start()
        printProcess.WaitForExit()
        'boucle d'attente de fin de process
        Do
            'avec une pause de 5 secondes pour soulager l'uc
            Threading.Thread.Sleep(100)
            'bouclage tant que le process n'a pas terminé l'impression
        Loop Until printProcess.HasExited
    End Sub



    Private Sub FileSystemWatcher1_Created(sender As Object, e As FileSystemEventArgs) Handles FileSystemWatcher1.Created
        Dim fileEntries As String() = Directory.GetFiles(folderPath)
        For Each filePath As String In fileEntries
            Dim fileExtension As String = Path.GetExtension(filePath)
            Dim fileName As String = Path.GetFileName(filePath)
            If fileExtension = ".pdf" Then
                Invoke(New MethodInvoker(Sub() TextBox1.AppendText(fileName & vbNewLine & vbNewLine)))
                ' ATTENDRE LA COPIE INTEGRALE DU FICHIER
                Dim FileTotalSize As Long = 5
                Dim FileNewSize As Long = 0
                Dim FileInfo As New FileInfo(filePath)
                Do
                    FileInfo.Refresh()
                    FileTotalSize = FileInfo.Length()
                    Thread.Sleep(5000)
                    FileInfo.Refresh()
                    FileNewSize = FileInfo.Length
                Loop Until FileNewSize = FileTotalSize And FileNewSize <> 0 And FileTotalSize <> 0
                ' LANCEMENT DES PROCEDURES D'IMPRESSION ADOBE READER
                ProcessAcroRd32(filePath, fileName)
            End If
        Next
    End Sub



    Private Sub ProcessAcroRd32(filePath, fileName)
        ' NOMBRE DE PROCESS ADOBE READER EN COURS
        Dim myProcessesA() As Process
        myProcessesA = Process.GetProcessesByName("AcroRd32")
        Dim NbProcessA As Integer = myProcessesA.Count
        ' LANCEMENT DES PROCEDURE D'IMPRESSION
        GoPrint(filePath)
        ' L'APPLICATION ADOBE READER EST-ELLE LIBRE ?
        Dim myProcessesB() As Process
        Dim myCurrentProcess As Boolean = True
        myProcessesB = Process.GetProcessesByName("AcroRd32")
        For Each myProcess In myProcessesB
            Dim Nom As String = myProcess.MainWindowTitle
            If Nom.Contains(fileName) = True Then
                Do
                    Thread.Sleep(1000)
                    myProcess.Refresh()
                    Nom = myProcess.MainWindowTitle
                    myCurrentProcess = Nom.Contains(fileName)
                Loop Until myCurrentProcess <> True
            End If
        Next myProcess
        ' SUPPRESSION DU FICHIER
        Thread.Sleep(1000)
        My.Computer.FileSystem.DeleteFile(filePath)
        Invoke(New MethodInvoker(Sub() TextBox1.AppendText("■  " & DateTime.Now.ToString("ddd d MMM à HH\hmm") & " Fichier supprimé." & vbNewLine & vbNewLine)))
        If NbProcessA = 0 Then Shell("TASKKILL /IM AcroRd32.exe /F")
    End Sub



    Private Sub GoPrint(FilePath)
        ' LISTE DES IMPRIMANTES
        Dim PcName As String = "\\" & My.Computer.Name
        Dim myPrintServer As New PrintServer(PcName, PrintSystemDesiredAccess.AdministrateServer)
        Dim myPrintQueues As PrintQueueCollection = myPrintServer.GetPrintQueues()
        For Each printer As PrintQueue In myPrintQueues
            If printer.Name.Contains("SHARP_A3") Then
                printer.Refresh()
                ' MESSAGE STATUS IMPRIMANTE
                Dim JobsCountA As Integer = printer.GetPrintJobInfoCollection.Count()
                Dim statusReport As String = ""
                Dim PingPrinter As String = SpotTroubleUsingProperties(statusReport, printer)
                Invoke(New MethodInvoker(Sub() TextBox1.AppendText(printer.Name & " : " & vbNewLine & PingPrinter & vbNewLine)))
                ' IMPRESSION
                Thread.Sleep(5000)
                Dim starter As New ProcessStartInfo(sProcess, "/t """ + FilePath + """  " + printer.Name + "")
                Dim Process As New Process()
                Process.StartInfo = starter
                Process.Start()
                Process.CloseMainWindow()
                ' LE TRAVAIL EST-IL EN IMPRESSION ?
                printer.Refresh()
                Dim JobsCountB As Integer = printer.GetPrintJobInfoCollection.Count()
                Do
                    Thread.Sleep(1000)
                    printer.Refresh()
                    JobsCountB = printer.GetPrintJobInfoCollection.Count()
                Loop Until JobsCountB <> JobsCountA
                If JobsCountB > 0 Then
                    Dim jobs As PrintJobInfoCollection = printer.GetPrintJobInfoCollection()
                    Dim theJob As PrintSystemJobInfo = jobs.Last
                    Do
                        printer.Refresh()
                        Dim jobsA As PrintJobInfoCollection = printer.GetPrintJobInfoCollection()
                        theJob = jobsA.Last
                    Loop Until FilePath.Contains(theJob.Name)
                    ' MESSAGE STATUS JOB
                    Dim statusJobReport As String = ""
                    Dim PingJobPrinter As String = SpotTroubleUsingJobAttributes(statusJobReport, theJob)
                    Invoke(New MethodInvoker(Sub() TextBox1.AppendText(PingJobPrinter & vbNewLine & vbNewLine)))
                End If
            End If
        Next printer
    End Sub


    ' Check for possible trouble states of a printer using its properties
    Function SpotTroubleUsingProperties(ByRef statusReport As String, ByVal pq As PrintQueue)
        If (pq.QueueStatus And PrintQueueStatus.Busy) = PrintQueueStatus.Busy Then statusReport = statusReport & "L'imprimante est occupée. "
        If (pq.QueueStatus And PrintQueueStatus.DoorOpen) = PrintQueueStatus.DoorOpen Then statusReport = statusReport & "Une porte sur l’imprimante est ouverte. "
        If (pq.QueueStatus And PrintQueueStatus.Error) = PrintQueueStatus.Error Then statusReport = statusReport & "L’imprimante ne peut pas imprimer en raison d’une condition d’erreur. "
        If (pq.QueueStatus And PrintQueueStatus.Initializing) = PrintQueueStatus.Initializing Then statusReport = statusReport & "L’initialisation de l’imprimante. "
        If (pq.QueueStatus And PrintQueueStatus.IOActive) = PrintQueueStatus.IOActive Then statusReport = statusReport & "L’imprimante échange des données avec le serveur d’impression. "
        If (pq.QueueStatus And PrintQueueStatus.ManualFeed) = PrintQueueStatus.ManualFeed Then statusReport = statusReport & "L’imprimante est en attente pour un utilisateur placer les supports d’impression dans le bac d’alimentation manuelle. "
        If (pq.QueueStatus And PrintQueueStatus.None) = PrintQueueStatus.None Then statusReport = statusReport & ""
        If (pq.QueueStatus And PrintQueueStatus.NotAvailable) = PrintQueueStatus.NotAvailable Then statusReport = statusReport & "Informations d’état ne sont pas disponibles. "
        If (pq.QueueStatus And PrintQueueStatus.NoToner) = PrintQueueStatus.NoToner Then statusReport = statusReport & "L’imprimante manque de toner. "
        If (pq.QueueStatus And PrintQueueStatus.Offline) = PrintQueueStatus.Offline Then statusReport = statusReport & "L’imprimante est hors connexion. "
        If (pq.QueueStatus And PrintQueueStatus.OutOfMemory) = PrintQueueStatus.OutOfMemory Then statusReport = statusReport & "L’imprimante n’a aucun mémoire disponible. "
        If (pq.QueueStatus And PrintQueueStatus.OutputBinFull) = PrintQueueStatus.OutputBinFull Then statusReport = statusReport & "La sortie imprimante bin est plein. "
        If (pq.QueueStatus And PrintQueueStatus.PagePunt) = PrintQueueStatus.PagePunt Then statusReport = statusReport & "L’imprimante ne peut pas imprimer la page en cours. "
        If (pq.QueueStatus And PrintQueueStatus.PaperJam) = PrintQueueStatus.PaperJam Then statusReport = statusReport & "Le document de l’imprimante est bloqué. "
        If (pq.QueueStatus And PrintQueueStatus.PaperOut) = PrintQueueStatus.PaperOut Then statusReport = statusReport & "L’imprimante n’a pas ou est du type de papier nécessaire pour le travail d’impression en cours. "
        If (pq.QueueStatus And PrintQueueStatus.PaperProblem) = PrintQueueStatus.PaperProblem Then statusReport = statusReport & "Le document de l’imprimante provoque une condition d’erreur non spécifiée. "
        If (pq.QueueStatus And PrintQueueStatus.Paused) = PrintQueueStatus.Paused Then statusReport = statusReport & "File d’attente suspendu. "
        If (pq.QueueStatus And PrintQueueStatus.PendingDeletion) = PrintQueueStatus.PendingDeletion Then statusReport = statusReport & "La file d’attente d’impression supprime un travail d’impression. "
        If (pq.QueueStatus And PrintQueueStatus.PowerSave) = PrintQueueStatus.PowerSave Then statusReport = statusReport & "L’imprimante est en mode veille. "
        If (pq.QueueStatus And PrintQueueStatus.Printing) = PrintQueueStatus.Printing Then statusReport = statusReport & "Le périphérique d’impression est. "
        If (pq.QueueStatus And PrintQueueStatus.Processing) = PrintQueueStatus.Processing Then statusReport = statusReport & "L’appareil est effectue un travail qui ne doive pas imprimer si le périphérique est une combinaison imprimante, télécopieur et moteur d’analyse. "
        If (pq.QueueStatus And PrintQueueStatus.ServerUnknown) = PrintQueueStatus.ServerUnknown Then statusReport = statusReport & "L’imprimante est en état d’erreur. "
        If (pq.QueueStatus And PrintQueueStatus.TonerLow) = PrintQueueStatus.TonerLow Then statusReport = statusReport & "Toner bas. "
        If (pq.QueueStatus And PrintQueueStatus.UserIntervention) = PrintQueueStatus.UserIntervention Then statusReport = statusReport & "L’imprimante requiert une action de l’utilisateur pour résoudre une condition d’erreur. "
        If (pq.QueueStatus And PrintQueueStatus.Waiting) = PrintQueueStatus.Waiting Then statusReport = statusReport & "L’imprimante est en attente pour un travail d’impression. "
        If (pq.QueueStatus And PrintQueueStatus.WarmingUp) = PrintQueueStatus.WarmingUp Then statusReport = statusReport & "L’imprimante est en cours de route. "
        Return statusReport
    End Function 'end SpotTroubleUsingProperties


    ' Check for possible trouble states of a print job using the flags of the JobStatus property
    Function SpotTroubleUsingJobAttributes(ByRef statusJobReport As String, ByVal theJob As PrintSystemJobInfo)
        If (theJob.JobStatus And PrintJobStatus.Blocked) = PrintJobStatus.Blocked Then statusJobReport = statusJobReport & "Une condition d’erreur, éventuellement sur un travail d’impression qui précède dans la file d’attente, a bloqué le travail d’impression. "
        If (theJob.JobStatus And PrintJobStatus.Completed) = PrintJobStatus.Completed Then statusJobReport = statusJobReport & "Le travail d’impression est terminé, y compris tout traitement de post impression. "
        If (theJob.JobStatus And PrintJobStatus.Deleted) = PrintJobStatus.Deleted Then statusJobReport = statusJobReport & "Le travail d’impression a été supprimé de la file d’attente, généralement après l’impression. "
        If (theJob.JobStatus And PrintJobStatus.Deleting) = PrintJobStatus.Deleting Then statusJobReport = statusJobReport & "Le travail d’impression est en cours de suppression. "
        If (theJob.JobStatus And PrintJobStatus.Error) = PrintJobStatus.Error Then statusJobReport = statusJobReport & "Le travail d’impression est dans un état d’erreur. "
        If (theJob.JobStatus And PrintJobStatus.None) = PrintJobStatus.None Then statusJobReport = statusJobReport & "Le travail d’impression n’a aucun état spécifié. "
        If (theJob.JobStatus And PrintJobStatus.Offline) = PrintJobStatus.Offline Then statusJobReport = statusJobReport & "L’imprimante est hors connexion. "
        If (theJob.JobStatus And PrintJobStatus.PaperOut) = PrintJobStatus.PaperOut Then statusJobReport = statusJobReport & "L’imprimante est hors de la taille du papier requis. "
        If (theJob.JobStatus And PrintJobStatus.Paused) = PrintJobStatus.Paused Then statusJobReport = statusJobReport & "Le travail d’impression est suspendu. "
        If (theJob.JobStatus And PrintJobStatus.Printed) = PrintJobStatus.Printed Then statusJobReport = statusJobReport & "Le travail d’impression est imprimé. "
        If (theJob.JobStatus And PrintJobStatus.Printing) = PrintJobStatus.Printing Then statusJobReport = statusJobReport & "Le travail d’impression est en cours d’impression. "
        If (theJob.JobStatus And PrintJobStatus.Restarted) = PrintJobStatus.Restarted Then statusJobReport = statusJobReport & "Le travail d’impression a été bloqué mais a redémarré. "
        If (theJob.JobStatus And PrintJobStatus.Retained) = PrintJobStatus.Retained Then statusJobReport = statusJobReport & "Le travail d’impression est conservé dans la file d’attente d’impression après l’impression. "
        If (theJob.JobStatus And PrintJobStatus.Spooling) = PrintJobStatus.Spooling Then statusJobReport = statusJobReport & "Le travail d’impression est mis en attente. "
        If (theJob.JobStatus And PrintJobStatus.UserIntervention) = PrintJobStatus.UserIntervention Then statusJobReport = statusJobReport & "L’imprimante requiert une action de l’utilisateur de résoudre une condition d’erreur. "
        Return statusJobReport
    End Function 'end SpotTroubleUsingJobAttributes

End Class