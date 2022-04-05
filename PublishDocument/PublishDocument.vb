Imports System.IO
Imports System.Xml


Module PublishDocument

    Dim OutputFolder1Page As String
    Dim OutputFolder2to5Pages As String
    Dim OutputFolder6to10Pages As String
    Dim OutputFolder10PlusPages As String
    Dim BarCodeImagesFolder As String
    Dim Pages() As String
    Dim PrintRotation As String

    Dim PortraitAlignment As String
    Dim PortraitHorizontal As String
    Dim PortraitVertical As String
    Dim LandscapeAlignment As String
    Dim LandscapeHorizontal As String
    Dim LandscapeVertical As String

    Dim JobTicketTemplatePath As String
    Dim JobTicketInputFolder As String

    Dim TTFCounterFileName As String
    Dim STTCounterFileName As String

    Dim InputFolder As String

    Dim PDFInfoFileName As String
    Dim PDFInfoFile As String

    Dim TTFCounter As Integer = 0
    Dim STTCounter As Integer = 0

    Dim P1BarCodeFileName(15) As String
    Dim P2BarCodeFileName(15) As String

    Dim InputPath As String
    Dim AdlibInputPath As String

    Dim LogFile As String = "C:\DocumentProcessor.log"

    Dim PDFDocFileName As String
    Dim PDFDocShortFileName As String

    Dim Debug As Boolean

    Dim AppPath As String
    Dim ScriptRootDir As String

    Sub Main()
        Initialize()
        ReadPDFInfo()
        Cleanup()
    End Sub

    Private Sub Cleanup()
        Try
            Dim ZipFileName As String = Left(PDFDocFileName, Len(PDFDocFileName) - 3) & "zip"

            If Debug = True Then WriteLog("Cleanup", "Deleting File", PDFInfoFileName)
            File.Delete(PDFInfoFileName)
            If Debug = True Then WriteLog("Cleanup", "Deleting File", ZipFileName)
            File.Delete(ZipFileName)

            If Debug = True Then WriteLog("Cleanup", "Cleaning up files older than 1 day.", InputPath)

            'Check any files more than 24 Hours Old, and delete them.
            Dim theFileDirectory As New DirectoryInfo(InputPath)
            Dim theFiles() As FileInfo = theFileDirectory.GetFiles()

            For Each theFile As FileInfo In theFiles

                Dim FileCreationDate = theFile.CreationTime.Date
                Dim FileAge As Integer = DateDiff(DateInterval.Day, FileCreationDate, Now())

                If FileAge > 1 Then
                    Try
                        If Debug = True Then WriteLog("Cleanup", "Deleting File", theFile.FullName.ToString)
                        theFile.Delete()
                    Catch ex As Exception
                        WriteLog("Cleanup", "Error Deleting " & theFile.FullName.ToString, ex.Message)
                    End Try
                End If
            Next

        Catch ex As Exception
            WriteLog("Cleanup", "Error deleting files", ex.Message)
        End Try
    End Sub

    Private Sub FailGracefully()
        'Copy ZIP Back to Input Folder

        WriteLog("ERROR", "Attempting to Fail Gracefully", "")

        Dim ZipFileName As String = Left(PDFDocFileName, Len(PDFDocFileName) - 3) & "zip"
        Dim ZipFileShortName = Right(ZipFileName, Len(ZipFileName) - InStrRev(ZipFileName, "\"))

        If File.Exists(ZipFileName) Then
            File.Copy(ZipFileName, AdlibInputPath & "\" & ZipFileShortName, True)
            WriteLog("ERROR", "Successfully copied ZIP back to Adlib Input Folder", AdlibInputPath & "\" & ZipFileName)
        End If

        End

    End Sub

    Private Sub Initialize()

        'Console.writeline("Attach Debugger to Process than press Enter")
        'Console.ReadLine()

        If Debug = True Then
            WriteLog("Arguments", _
            Environment.GetCommandLineArgs(1).ToString & " " & _
            Environment.GetCommandLineArgs(2).ToString, Environment.GetCommandLineArgs(0).ToString)

        End If

        AppPath = Environment.GetCommandLineArgs(0).ToString
        ScriptRootDir = Left(AppPath, InStrRev(AppPath, "\") - 1)

        Dim strInstructions As String

        strInstructions = "This script requires an Argument.  [&OutputFolder] and &[InputFolder]" & vbCrLf
        strInstructions += vbCrLf & "Usage: PublishDocument.exe [&OutputFolder] [&InputFolder]"

        'What we really want is the name of the output file, and if it is .XML we'll parse it.
        'We'll also make sure that the accompanying PDF file also exists.

        'Configure me as a Post Processing Script, AFTER documents are moved.
        'Make sure that Move Source Docuemnt to Output Folder is enabled.
        'This is because we'll move the source document BACK to the input folder in the event
        'of an unrecoverable failure.

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

        If Environment.GetCommandLineArgs(1).EndsWith(".tif") Then

            PDFInfoFileName = Environment.GetCommandLineArgs(1).ToString
            PDFInfoFileName = Left(PDFInfoFileName, Len(PDFInfoFileName) - 3) + "xml"
            'PDFDocFileName = Left(PDFInfoFileName, Len(PDFInfoFileName) - 3) + "pdf"
            PDFDocFileName = Left(PDFInfoFileName, Len(PDFInfoFileName) - 3) + "tif"
            PDFDocShortFileName = Right(PDFDocFileName, Len(PDFDocFileName) - InStrRev(PDFDocFileName, "\"))

            InputPath = Left(PDFInfoFileName, InStrRev(PDFInfoFileName, "\") - 1)

            AdlibInputPath = Environment.GetCommandLineArgs(2).ToString

            If AdlibInputPath.EndsWith("\") Then AdlibInputPath = Left(AdlibInputPath, Len(AdlibInputPath) - 1)

            ReadSettingsFile()

            If Directory.Exists(OutputFolder1Page) = False Then
                Try
                    Directory.CreateDirectory(OutputFolder1Page)
                Catch ex As Exception
                    WriteLog("Initialization", "Error-Directory Not Found and could not be Created.", OutputFolder1Page)
                    End
                End Try
            End If
            If Directory.Exists(OutputFolder2to5Pages) = False Then
                Try
                    Directory.CreateDirectory(OutputFolder2to5Pages)
                Catch ex As Exception
                    WriteLog("Initialization", "Error-Directory Not Found and could not be Created.", OutputFolder2to5Pages)
                    End
                End Try
            End If
            If Directory.Exists(OutputFolder6to10Pages) = False Then
                Try
                    Directory.CreateDirectory(OutputFolder6to10Pages)
                Catch ex As Exception
                    WriteLog("Initialization", "Error-Directory Not Found and could not be Created.", OutputFolder6to10Pages)
                    End
                End Try
            End If
            If Directory.Exists(OutputFolder10PlusPages) = False Then
                Try
                    Directory.CreateDirectory(OutputFolder10PlusPages)
                Catch ex As Exception
                    WriteLog("Initialization", "Error-Directory Not Found and could not be Created.", OutputFolder10PlusPages)
                    End
                End Try
            End If
            If File.Exists(PDFInfoFileName) = False Then
                WriteLog("Initialization", "Could not find PDFInfo File", PDFInfoFileName)
                End
            End If

            If File.Exists(PDFDocFileName) = False Then
                WriteLog("Initialization", "Could not find PDF File", PDFDocFileName)
                End
            End If



            Try '   ===========Initialize Counters

                TTFCounterFileName = ScriptRootDir & "\TTFCounter.dat"
                STTCounterFileName = ScriptRootDir & "\STTCounter.dat"

                If File.Exists(TTFCounterFileName) = False Then
                    File.WriteAllText(TTFCounterFileName, "1")
                End If

                If File.Exists(STTCounterFileName) = False Then
                    File.WriteAllText(STTCounterFileName, "1")
                End If

                TTFCounter = CInt(File.ReadAllText(TTFCounterFileName))
                STTCounter = CInt(File.ReadAllText(STTCounterFileName))

                'Check & Set Barcode Image Path
                P1BarCodeFileName(0) = BarCodeImagesFolder & "\O0-p"
                P2BarCodeFileName(0) = BarCodeImagesFolder & "\80-p"
                P1BarCodeFileName(1) = BarCodeImagesFolder & "\K0-p"
                P2BarCodeFileName(1) = BarCodeImagesFolder & "\40-p"
                P1BarCodeFileName(2) = BarCodeImagesFolder & "\S0-p"
                P2BarCodeFileName(2) = BarCodeImagesFolder & "\C0-p"
                P1BarCodeFileName(3) = BarCodeImagesFolder & "\I0-p"
                P2BarCodeFileName(3) = BarCodeImagesFolder & "\20-p"
                P1BarCodeFileName(4) = BarCodeImagesFolder & "\Q0-p"
                P2BarCodeFileName(4) = BarCodeImagesFolder & "\A0-p"
                P1BarCodeFileName(5) = BarCodeImagesFolder & "\M0-p"
                P2BarCodeFileName(5) = BarCodeImagesFolder & "\60-p"
                P1BarCodeFileName(6) = BarCodeImagesFolder & "\U0-p"
                P2BarCodeFileName(6) = BarCodeImagesFolder & "\E0-p"
                P1BarCodeFileName(7) = BarCodeImagesFolder & "\H0-p"
                P2BarCodeFileName(7) = BarCodeImagesFolder & "\10-p"
                P1BarCodeFileName(8) = BarCodeImagesFolder & "\P0-p"
                P2BarCodeFileName(8) = BarCodeImagesFolder & "\90-p"
                P1BarCodeFileName(9) = BarCodeImagesFolder & "\L0-p"
                P2BarCodeFileName(9) = BarCodeImagesFolder & "\50-p"
                P1BarCodeFileName(10) = BarCodeImagesFolder & "\T0-p"
                P2BarCodeFileName(10) = BarCodeImagesFolder & "\D0-p"
                P1BarCodeFileName(11) = BarCodeImagesFolder & "\J0-p"
                P2BarCodeFileName(11) = BarCodeImagesFolder & "\30-p"
                P1BarCodeFileName(12) = BarCodeImagesFolder & "\R0-p"
                P2BarCodeFileName(12) = BarCodeImagesFolder & "\B0-p"
                P1BarCodeFileName(13) = BarCodeImagesFolder & "\N0-p"
                P2BarCodeFileName(13) = BarCodeImagesFolder & "\70-p"
                P1BarCodeFileName(14) = BarCodeImagesFolder & "\V0-p"
                P2BarCodeFileName(14) = BarCodeImagesFolder & "\F0-p"

                Dim I As Integer

                For I = 0 To 14
                    CheckBarCodeImageExists(I)
                Next

            Catch ex As Exception

                If TTFCounter = 0 Then
                    WriteLog("Initialization", "Error Initializing Counters.", ex.Message)
                    TTFCounter = 1
                End If

                If STTCounter = 0 Then
                    WriteLog("Initialization", "Error Initializing Counters.", ex.Message)
                    STTCounter = 1
                End If

                WriteLog("Initialization", "Error.", ex.ToString)

                FailGracefully()

            End Try

        Else
            'The Output Doc was not .xml, so no sense in doing anything then.
            End
        End If

    End Sub

    Sub CheckBarCodeImageExists(ByVal Item As Integer)
        If File.Exists(P1BarCodeFileName(Item) + ".bmp") = False Then
            WriteLog("Initialization", "Barcode P1 Item-" & Item & " Not Found.", P1BarCodeFileName(Item))
            FailGracefully()
        End If
        If File.Exists(P2BarCodeFileName(Item) + ".bmp") = False Then
            WriteLog("Initialization", "Barcode 2 Item-" & Item & " Not Found.", P2BarCodeFileName(Item))
            FailGracefully()
        End If
    End Sub

    Sub ReadPDFInfo()
        Dim xmlPDFInfo As New XmlDocument
        Dim PageNodes As XmlNodeList
        Dim PageErrorMessage As String = ""
        Dim Page As Integer
        Dim OverlaySettings As String = ""
        Dim Width As Integer
        Dim Height As Integer
        Dim Orientation As String = "Portrait"


        If File.Exists(PDFInfoFileName) Then
            Try 'To Load the XML

                xmlPDFInfo.Load(PDFInfoFileName)

                PageNodes = xmlPDFInfo.GetElementsByTagName("PAGE")

                If PageNodes.Count = 1 Then
                    'MoveOutput
                    If Debug = True Then WriteLog("ReadPDFInfo", "1 Page Read", PDFInfoFileName)

                    'Commented out, no longer simple copy, because we had to create TIFF output
                    'We now need a JT to convert compound doc to PDF
                    'But no stamping is required.
                    '2009-04-15 Jeff Brand
                    '
                    'If File.Exists(PDFDocFileName) Then
                    '    Try
                    '        If File.Exists(OutputFolder1Page & "\" & PDFDocShortFileName) Then
                    '            File.Delete(OutputFolder1Page & "\" & PDFDocShortFileName)
                    '        Else
                    '            File.Copy(PDFDocFileName, OutputFolder1Page & "\" & PDFDocShortFileName)
                    '        End If
                    '    Catch ex As Exception
                    '        WriteLog("COPY FILE", "Error Copying File to " & OutputFolder1Page & "\" & PDFDocShortFileName, ex.Message)
                    '        FailGracefully()
                    '    End Try
                    'End If

                    Dim strJobTicketTemplate As String = ""
                    Try
                        strJobTicketTemplate = File.ReadAllText(JobTicketTemplatePath)
                    Catch ex As Exception
                        WriteLog("ReadPDFInfo", "Error Reading Job Ticket Template", JobTicketTemplatePath)
                        FailGracefully()
                    End Try

                    strJobTicketTemplate = strJobTicketTemplate.Replace("[DocInputFolder]", InputPath)
                    strJobTicketTemplate = strJobTicketTemplate.Replace("[DocInputFilename]", PDFDocShortFileName)
                    strJobTicketTemplate = strJobTicketTemplate.Replace("[DocOutputFolder]", OutputFolder1Page)
                    strJobTicketTemplate = strJobTicketTemplate.Replace("[DocOutputFilename]", Left(PDFDocShortFileName, Len(PDFDocShortFileName) - 3) & "pdf")
                    strJobTicketTemplate = strJobTicketTemplate.Replace("[OverlaySettings]", "")

                    Dim JTFileName As String = JobTicketInputFolder & "\" & "O_JT_" & PDFDocShortFileName & ".xml"

                    If Debug = True Then WriteLog("ReadPDFInfo", "Writing XML Job Ticket", JTFileName)

                    Try
                        File.WriteAllText(JTFileName, strJobTicketTemplate, System.Text.Encoding.Unicode)
                        If Debug = True Then WriteLog("ReadPDFInfo", "Writing XML Job Ticket - Complete", JTFileName)

                    Catch ex As Exception
                        WriteLog("ReadPDFInfo", "Error writing new Job Ticket " & JTFileName, ex.Message)
                    End Try

                ElseIf PageNodes.Count > 1 And PageNodes.Count < 6 Then

                    If Debug = True Then WriteLog("ReadPDFInfo", "2-5 Pages Read", PDFInfoFileName)
                    If Debug = True Then WriteLog("ReadPDFInfo", "Reading XML Job Ticket Template", JobTicketTemplatePath)


                    'Determine orientation for each page and add an Overlay item for each page
                    'Based on orientation and position.

                    OverlaySettings = ""
                    For Page = 0 To PageNodes.Count - 1

                        Dim BarCodeRotation As String = ""
                        Dim BarCodePath As String = ""
                        Dim Alignment As String = "BottomRight"
                        Dim Horizontal As String = ".1"
                        Dim Vertical As String = "4"



                        Width = CInt(PageNodes(Page).Attributes("WIDTH").Value.ToString)
                        Height = CInt(PageNodes(Page).Attributes("HEIGHT").Value.ToString)

                        If Height > Width Then Orientation = "Portrait"
                        If Width > Height Then Orientation = "Landscape"

                        If PrintRotation = "Left" Then
                            BarCodeRotation = "-r"
                        ElseIf PrintRotation = "Right" Then
                            BarCodeRotation = "-l"
                        End If

                        If Page = 0 Then BarCodePath = P1BarCodeFileName(TTFCounter)
                        If Page > 0 Then BarCodePath = P2BarCodeFileName(TTFCounter)

                        If Orientation = "Landscape" Then

                            BarCodePath += BarCodeRotation
                            Alignment = LandscapeAlignment
                            Horizontal = LandscapeHorizontal
                            Vertical = LandscapeVertical

                        ElseIf Orientation = "Portrait" Then

                            Alignment = PortraitAlignment
                            Horizontal = PortraitHorizontal
                            Vertical = PortraitVertical

                        End If

                        BarCodePath += ".bmp"

                        OverlaySettings += "			<JOB:OVERLAY ENABLED='Yes' " & _
                        "PATH='" & BarCodePath & "' LAYER='Foreground' PAGES='" & (Page + 1).ToString & "' " & _
                        "ALIGNMENT='" & Alignment & "' VERTICAL='" & Vertical & "' HORIZONTAL='" & Horizontal & "'/>" & vbCrLf

                    Next

                    Dim strJobTicketTemplate As String = ""
                    Try
                        strJobTicketTemplate = File.ReadAllText(JobTicketTemplatePath)
                    Catch ex As Exception
                        WriteLog("ReadPDFInfo", "Error Reading Job Ticket Template", JobTicketTemplatePath)
                        FailGracefully()
                    End Try

                    strJobTicketTemplate = strJobTicketTemplate.Replace("[DocInputFolder]", InputPath)
                    strJobTicketTemplate = strJobTicketTemplate.Replace("[DocInputFilename]", PDFDocShortFileName)
                    strJobTicketTemplate = strJobTicketTemplate.Replace("[DocOutputFolder]", OutputFolder2to5Pages)
                    strJobTicketTemplate = strJobTicketTemplate.Replace("[DocOutputFilename]", TTFCounter.ToString & "_" & Left(PDFDocShortFileName, Len(PDFDocShortFileName) - 3) & "pdf")
                    strJobTicketTemplate = strJobTicketTemplate.Replace("[OverlaySettings]", OverlaySettings)

                    Dim JTFileName As String = JobTicketInputFolder & "\" & "TTF_JT_" & _
                                                TTFCounter.ToString & "_" & PDFDocShortFileName & ".xml"

                    If Debug = True Then WriteLog("ReadPDFInfo", "Writing XML Job Ticket", JTFileName)

                    Try
                        File.WriteAllText(JTFileName, strJobTicketTemplate, System.Text.Encoding.Unicode)
                        If Debug = True Then WriteLog("ReadPDFInfo", "Writing XML Job Ticket - Complete", JTFileName)

                    Catch ex As Exception
                        WriteLog("ReadPDFInfo", "Error writing new Job Ticket " & JTFileName, ex.Message)
                    End Try

                    'Increment our counter
                    If Debug = True Then WriteLog("ReadPDFInfo", "Incrementing TTFCounter", JTFileName)

                    If TTFCounter = 14 Then
                        File.WriteAllText(TTFCounterFileName, "0")
                    ElseIf TTFCounter < 14 Then
                        File.WriteAllText(TTFCounterFileName, (TTFCounter + 1).ToString)
                    ElseIf TTFCounter > 14 Then
                        WriteLog("ReadPDFInfo", "Error - TTFCounter was above 14!", TTFCounter.ToString)
                        File.WriteAllText(TTFCounterFileName, "0")
                        TTFCounter = 1
                    End If

                ElseIf PageNodes.Count > 5 And PageNodes.Count < 11 Then

                    'Barcode Job Ticket
                    If Debug = True Then WriteLog("ReadPDFInfo", "6-10 Pages Read", PDFInfoFileName)
                    If Debug = True Then WriteLog("ReadPDFInfo", "Reading XML Job Ticket Template", JobTicketTemplatePath)

                    OverlaySettings = ""
                    For Page = 0 To PageNodes.Count - 1

                        Dim BarCodeRotation As String = ""
                        Dim BarCodePath As String = ""
                        Dim Alignment As String = "BottomRight"
                        Dim Horizontal As String = ".1"
                        Dim Vertical As String = "4"


                        Width = CInt(PageNodes(Page).Attributes("WIDTH").Value.ToString)
                        Height = CInt(PageNodes(Page).Attributes("HEIGHT").Value.ToString)

                        If Height > Width Then Orientation = "Portrait"
                        If Width > Height Then Orientation = "Landscape"

                        If PrintRotation = "Left" Then
                            BarCodeRotation = "-r"
                        ElseIf PrintRotation = "Right" Then
                            BarCodeRotation = "-l"
                        End If

                        If Page = 0 Then BarCodePath = P1BarCodeFileName(STTCounter)
                        If Page > 0 Then BarCodePath = P2BarCodeFileName(STTCounter)

                        If Orientation = "Landscape" Then

                            BarCodePath += BarCodeRotation
                            Alignment = LandscapeAlignment
                            Horizontal = LandscapeHorizontal
                            Vertical = LandscapeVertical

                        ElseIf Orientation = "Portrait" Then

                            Alignment = PortraitAlignment
                            Horizontal = PortraitHorizontal
                            Vertical = PortraitVertical

                        End If

                        BarCodePath += ".bmp"

                        OverlaySettings += "			<JOB:OVERLAY ENABLED='Yes' " & _
                        "PATH='" & BarCodePath & "' LAYER='Foreground' PAGES='" & (Page + 1).ToString & "' " & _
                        "ALIGNMENT='" & Alignment & "' VERTICAL='" & Vertical & "' HORIZONTAL='" & Horizontal & "'/>" & vbCrLf

                    Next

                    Dim strJobTicketTemplate As String = ""
                    Try
                        strJobTicketTemplate = File.ReadAllText(JobTicketTemplatePath)
                    Catch ex As Exception
                        WriteLog("ReadPDFInfo", "Error Reading Job Ticket Template", JobTicketTemplatePath)
                        FailGracefully()
                    End Try

                    strJobTicketTemplate = strJobTicketTemplate.Replace("[DocInputFolder]", InputPath)
                    strJobTicketTemplate = strJobTicketTemplate.Replace("[DocInputFilename]", PDFDocShortFileName)
                    strJobTicketTemplate = strJobTicketTemplate.Replace("[DocOutputFolder]", OutputFolder6to10Pages)
                    strJobTicketTemplate = strJobTicketTemplate.Replace("[DocOutputFilename]", STTCounter.ToString & "_" & Left(PDFDocShortFileName, Len(PDFDocShortFileName) - 3) & "pdf")
                    strJobTicketTemplate = strJobTicketTemplate.Replace("[OverlaySettings]", OverlaySettings)

                    Dim JTFileName As String = JobTicketInputFolder & "\" & "STT_JT_" & _
                                                STTCounter.ToString & "_" & PDFDocShortFileName & ".xml"

                    If Debug = True Then WriteLog("ReadPDFInfo", "Writing XML Job Ticket", JTFileName)

                    Try
                        File.WriteAllText(JTFileName, strJobTicketTemplate, System.Text.Encoding.Unicode)
                        If Debug = True Then WriteLog("ReadPDFInfo", "Writing XML Job Ticket - Complete", JTFileName)

                    Catch ex As Exception
                        WriteLog("ReadPDFInfo", "Error writing new Job Ticket " & JTFileName, ex.Message)
                    End Try

                    'Increment our counter
                    If Debug = True Then WriteLog("ReadPDFInfo", "Incrementing STTCounter", JTFileName)

                    'Increment our Counter
                    If STTCounter = 14 Then
                        File.WriteAllText(STTCounterFileName, "0")
                    ElseIf STTCounter < 14 Then
                        File.WriteAllText(STTCounterFileName, (STTCounter + 1).ToString)
                    ElseIf STTCounter > 14 Then
                        WriteLog("Initialization", "Error - STTCounter was above 14!", STTCounter.ToString)
                        File.WriteAllText(STTCounterFileName, "0")
                        STTCounter = 1
                    End If

                ElseIf PageNodes.Count > 10 Then

                    ''MoveOutput

                    'Commented out, no longer simple copy, because we had to create TIFF output
                    'We now need a JT to convert compound doc to PDF
                    'But no stamping is required.
                    '2009-04-15 Jeff Brand
                    '

                    'Try
                    '    If File.Exists(OutputFolder10PlusPages & "\" & PDFDocShortFileName) Then
                    '        File.Delete(OutputFolder10PlusPages & "\" & PDFDocShortFileName)
                    '    Else
                    '        File.Copy(PDFDocFileName, OutputFolder10PlusPages & "\" & PDFDocShortFileName)
                    '    End If

                    'Catch ex As Exception
                    '    WriteLog("COPY FILE", "Error Copying File to " & OutputFolder10PlusPages & "\" & PDFDocShortFileName, ex.Message)
                    '    FailGracefully()
                    'End Try

                    If Debug = True Then WriteLog("ReadPDFInfo", ">10 Pages Read", PDFInfoFileName)

                    Dim strJobTicketTemplate As String = ""
                    Try
                        strJobTicketTemplate = File.ReadAllText(JobTicketTemplatePath)
                    Catch ex As Exception
                        WriteLog("ReadPDFInfo", "Error Reading Job Ticket Template", JobTicketTemplatePath)
                        FailGracefully()
                    End Try

                    strJobTicketTemplate = strJobTicketTemplate.Replace("[DocInputFolder]", InputPath)
                    strJobTicketTemplate = strJobTicketTemplate.Replace("[DocInputFilename]", PDFDocShortFileName)
                    strJobTicketTemplate = strJobTicketTemplate.Replace("[DocOutputFolder]", OutputFolder10PlusPages)
                    strJobTicketTemplate = strJobTicketTemplate.Replace("[DocOutputFilename]", Left(PDFDocShortFileName, Len(PDFDocShortFileName) - 3) & "pdf")
                    strJobTicketTemplate = strJobTicketTemplate.Replace("[OverlaySettings]", "")

                    Dim JTFileName As String = JobTicketInputFolder & "\" & "GTT_JT_" & PDFDocShortFileName & ".xml"

                    If Debug = True Then WriteLog("ReadPDFInfo", "Writing XML Job Ticket", JTFileName)

                    Try
                        File.WriteAllText(JTFileName, strJobTicketTemplate, System.Text.Encoding.Unicode)
                        If Debug = True Then WriteLog("ReadPDFInfo", "Writing XML Job Ticket - Complete", JTFileName)
                    Catch ex As Exception
                        WriteLog("ReadPDFInfo", "Error writing new Job Ticket " & JTFileName, ex.Message)
                    End Try
                End If

                xmlPDFInfo = Nothing

            Catch ex As Exception
                WriteLog("ReadPDFInfo", "ERROR", ex.Message)
                FailGracefully()
            End Try
        End If
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

            Dim strName As String = ""
            Dim strValue As String = ""

            For Each Setting As XmlNode In XMLSettingsNodes

                strName = Setting.Attributes.GetNamedItem("NAME").Value
                strValue = Setting.Attributes.GetNamedItem("VALUE").Value

                If strName = "OutputFolder1Page" Then OutputFolder1Page = strValue
                If strName = "OutputFolder2to5Pages" Then OutputFolder2to5Pages = strValue
                If strName = "OutputFolder6to10Pages" Then OutputFolder6to10Pages = strValue
                If strName = "OutputFolder10PlusPages" Then OutputFolder10PlusPages = strValue
                If strName = "BarCodeImagesFolder" Then BarCodeImagesFolder = strValue
                If strName = "Debug" Then Debug = strValue
                If strName = "LogFile" Then LogFile = strValue
                If strName = "JobTicketTemplatePath" Then JobTicketTemplatePath = strValue
                If strName = "JobTicketInputFolder" Then JobTicketInputFolder = strValue
                If strName = "PrintRotation" Then PrintRotation = strValue
                If strName = "PortraitAlignment" Then PortraitAlignment = strValue
                If strName = "PortraitHorizontal" Then PortraitHorizontal = strValue
                If strName = "PortraitVertical" Then PortraitVertical = strValue
                If strName = "LandscapeAlignment" Then LandscapeAlignment = strValue
                If strName = "LandscapeHorizontal" Then LandscapeHorizontal = strValue
                If strName = "LandscapeVertical" Then LandscapeVertical = strValue

            Next

        Catch ex As Exception
            WriteLog("Initialialization", "Error Reading Settings", ex.Message.ToString)
        End Try
    End Sub
End Module
