Imports System.IO
Imports System.Xml
Imports System.Threading

Module ZipFolder

    Dim strInputFolder As String
    Dim strParentFolderPath As String
    Dim Debug As Boolean = True
    Dim AppPath As String
    Dim ScriptRootDir As String
    Dim LogFile As String = "C:\DocumentProcessor.log"
    Dim ZipPath As String
    Dim ZipTimeout As Integer

    Sub Main()
        Initialize()
        ZipFolders()
    End Sub

    Sub Initialize()

        AppPath = Environment.GetCommandLineArgs(0).ToString
        ScriptRootDir = Left(AppPath, InStrRev(AppPath, "\") - 1)

        ReadSettingsFile()

        Dim strInstructions As String

        strInstructions = "This script requires a 2 Arguments.  [&InputFolder] and the " & vbCrLf & _
        "Parent Folder path to look for folders" & vbCrLf
        strInstructions += vbCrLf & "Usage: ZipFolder.exe [&InputFolder] C:\SearchFolder"
        strInstructions += vbCrLf & "This application will zip every folder found in the Parent Folder path "
        strInstructions += vbCrLf & "and place the ZIP contents in the InputFolder."
        strInstructions += vbCrLf & "(c)2009, Adlib Software, Licensed by Enterprise Rent-a-Car Company."

        'Do we have the correct # of Arguments?  (3 because the first is the name of the application)
        If Environment.GetCommandLineArgs.Length <> 3 Then
            Dim cal As Integer = Environment.GetCommandLineArgs.Length

            Console.WriteLine(strInstructions)
            Console.WriteLine("Arguments Provided")
            For i As Integer = 1 To Environment.GetCommandLineArgs.Length - 1
                Console.WriteLine(Environment.GetCommandLineArgs(i).ToString)
            Next

            End

        ElseIf Environment.GetCommandLineArgs(1) = "/?" Then
            Console.WriteLine(strInstructions)
            End
        End If

        If Debug = True Then
            WriteLog("Arguments", _
            Environment.GetCommandLineArgs(1).ToString & " " & _
            Environment.GetCommandLineArgs(2).ToString, Environment.GetCommandLineArgs(0).ToString)

        End If


        strInputFolder = Environment.GetCommandLineArgs(1)
        If strInputFolder.EndsWith("\") Then strInputFolder = Left(strInputFolder, Len(strInputFolder) - 1)

        strParentFolderPath = Environment.GetCommandLineArgs(2)
        If strParentFolderPath.EndsWith("\") Then strParentFolderPath = Left(strParentFolderPath, Len(strParentFolderPath) - 1)

    End Sub

    Sub ZipFolders()
        Try

            For Each fileFound As String In Directory.GetFiles(strParentFolderPath)

                If fileFound.EndsWith("zip", True, System.Globalization.CultureInfo.CurrentCulture) Then

                    'Only process a subfolder that's more than 1 Minute since last write.

                    Dim FileAge As Integer
                    FileAge = DateDiff(DateInterval.Second, File.GetLastWriteTime(fileFound), Now)

                    If FileAge > 10 Then

                        If Debug = True Then WriteLog("ZipFolders", "Found ZIP File - Copying to Input Folder", fileFound)
                        Dim ShortFileName As String = Right(fileFound, Len(fileFound) - InStrRev(fileFound, "\"))

                        Try
                            File.Copy(fileFound, strInputFolder & "\" & ShortFileName)
                        Catch ex As Exception
                            WriteLog("ZipFolders", "Error Copying ZIP File", fileFound)
                        End Try
                        Try
                            If Debug = True Then WriteLog("ZipFolder", "Deleting ZipFile", fileFound)

                            File.Delete(fileFound)

                            If Debug = True Then WriteLog("ZipFolder", "ZipFile Deleted", fileFound)
                        Catch ex As Exception
                            WriteLog("ZipFolder", "Error Deleting Folder", ex.Message)
                        End Try
                    End If


                End If
            Next


            For Each subFolder As String In Directory.GetDirectories(strParentFolderPath)

                'Only process a subfolder that's more than 1 Minute since last write.
                Dim FileAge As Integer
                FileAge = DateDiff(DateInterval.Second, Directory.GetLastWriteTime(subFolder), Now)
                'If Debug = True Then FileAge = 2
                If FileAge > 10 Then
                    Dim FolderName As String = Right(subFolder, Len(subFolder) - InStrRev(subFolder, "\"))
                    If Debug = True Then WriteLog(FolderName, "Found Folder", "Path = " & subFolder)

                    '20090401 - Thumbs.db causes Express to Error...So let's look for those and delete them.
                    For Each foundFile As String In Directory.GetFiles(subFolder)
                        Dim ShortName As String = Right(foundFile, Len(foundFile) - InStrRev(foundFile, "\"))
                        If ShortName = "Thumbs.db" Then
                            Try
                                If Debug = True Then WriteLog("FolderName", "Found Thumbs.db So Deleting!", foundFile)
                                File.Delete(foundFile)
                            Catch ex As Exception
                                WriteLog(FolderName, "Error deleting " & foundFile, ex.Message)
                            End Try

                        End If
                    Next

                    ZipFolder(subFolder, FolderName)
                    Cleanup(subFolder, FolderName)
                End If
            Next
        Catch ex As Exception
            WriteLog("Unknown", "Error processing folders", ex.Message)
        End Try


    End Sub

    Sub ZipFolder(ByVal strFolderPath As String, ByVal strFolderName As String)

        Dim strZipFileName As String = strInputFolder & "\" & strFolderName & ".zip"

        Dim CommandLine As String = """" & ZipPath & " a """
        CommandLine += " " & """" & strZipFileName & """"
        CommandLine += " """ & strFolderPath & "\*"""

        Try
            If Debug = True Then WriteLog(strFolderName, "Calling 7za.exe", CommandLine)

            Dim objProcess As New Process()
            Dim WaitTime As Integer = 0

            Dim StartTime As DateTime = System.DateTime.Now()

            ' Start the Command and redirect the output

            objProcess.StartInfo.UseShellExecute = False
            objProcess.StartInfo.RedirectStandardOutput = True
            objProcess.StartInfo.CreateNoWindow = True
            objProcess.StartInfo.RedirectStandardError = True
            objProcess.StartInfo.FileName() = ZipPath
            objProcess.StartInfo.Arguments() = " a """ & strZipFileName & """" & " """ & strFolderPath & "\*"""
            objProcess.Start()

            While objProcess.HasExited = False
                System.Threading.Thread.Sleep(1000)
                WaitTime += 1

                If WaitTime > ZipTimeout Then
                    WriteLog(strFolderName, "Zip Timeout (" & ZipTimeout.ToString & " Seconds) Exceeded", "Duration = " & WaitTime & " Seconds.")
                    objProcess.Close()
                End If
            End While

            objProcess.Close()

            If Debug = True Then WriteLog(strFolderName, "Zip Creation Successful.  Duration = " & WaitTime & " Seconds.", CommandLine)

        Catch ex As Exception
            WriteLog(strFolderName, "Error Creating ZIP: " + CommandLine, ex.Message)
        End Try

    End Sub

    Sub Cleanup(ByVal strFolderPath As String, ByVal strFolderName As String)
        Try
            If Debug = True Then WriteLog(strFolderName, "Deleting Folder", strFolderPath)

            Directory.Delete(strFolderPath, True)

            If Debug = True Then WriteLog(strFolderName, "Folder Deleted", strFolderPath)
        Catch ex As Exception
            WriteLog(strFolderName, "Error Deleting Folder", ex.Message)
        End Try

    End Sub

    Private Sub WriteLog(ByVal FolderName As String, ByVal LogEvent As String, ByVal LogText As String)
        Try
            Dim strLogRecord As String
            strLogRecord = Now.ToString("G") & vbTab & FolderName & vbTab & LogEvent & vbTab & LogText & vbCrLf
            File.AppendAllText(LogFile, strLogRecord)
        Catch ex As Exception
            Console.WriteLine("Error Writing to Log: " & LogFile)
            Console.WriteLine("Error was: " & ex.Message)
        End Try
    End Sub

    Private Sub ReadSettingsFile()
        'Load Settings from our XML Settings File
        Dim XMLSettings As XmlDocument = New XmlDocument
        Dim XMLSettingsNodes As XmlNodeList

        Try

            XMLSettings.Load(ScriptRootDir & "\PublishDocumentSettings.xml")

            XMLSettingsNodes = XMLSettings.GetElementsByTagName("Setting")

            For Each Setting As XmlNode In XMLSettingsNodes
                If Setting.Attributes.GetNamedItem("NAME").Value = "LogFile" Then
                    LogFile = Setting.Attributes.GetNamedItem("VALUE").Value
                End If
                If Setting.Attributes.GetNamedItem("NAME").Value = "ZipPath" Then
                    ZipPath = Setting.Attributes.GetNamedItem("VALUE").Value
                End If
                If Setting.Attributes.GetNamedItem("NAME").Value = "ZipTimeout" Then
                    ZipTimeout = Setting.Attributes.GetNamedItem("VALUE").Value
                End If
                If Setting.Attributes.GetNamedItem("NAME").Value = "Debug" Then
                    Debug = Setting.Attributes.GetNamedItem("VALUE").Value
                End If
            Next

        Catch ex As Exception
            WriteLog("Initialialization", "Error Reading Settings", ex.Message.ToString)
        End Try
    End Sub

End Module
