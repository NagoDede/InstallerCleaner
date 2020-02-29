Imports System.IO
Imports Appature.Common
Imports System.Reflection
Imports Microsoft.Tools.WindowsInstallerXml
Imports System
Imports System.Linq

Module Module1

    Private Structure parserOptions
        Public toDelete As Boolean 'Request to delete files from InputDir
        Public toMove As Boolean 'Request to move files from InputDir to MoveDir
        Public toSimulate As Boolean 'Simulate the operation and generate a report
        Public toReport As Boolean 'Will create a report
        Public toShowHelp As Boolean 'To show the help file

        Public InputDir As String 'Directory where are stored the MSI/MSP files (should be C:\WINDOWS\Installer)
        Public MoveDir As String 'Taget directory when option Move is selected
        Public ReportPath As String 'Report file path
    End Structure

    Const WIN_DIR_INSTALLER = "C:\WINDOWS\Installer\"
    Const DFLT_DIR_MOVE = "C:\BackupInstaller"
    Const DFLT_PATH_REPORT = "C:\Temp\InstallerClean.txt"

    Private parseOptions As parserOptions
    Private FailedList As List(Of String) 'Messages generated when error occurs

    Private MaxPathLength As Integer 'max filename length

    Sub Main(args() As String)

        MaxPathLength = GetMaxPathLength() 'retrieve the max filename length
        FailedList = New List(Of String)

        Dim parser As New CommandParser
        InitParse(parser)
        parser.Parse()

        If parseOptions.toShowHelp Then
            Console.WriteLine(parser.GetHelp())
            PressKeyToQuit()
        End If

        'test the configuration and init
        TestConfigurationAndInit()

        'create the report file if necessary
        Dim sw As StreamWriter
        If parseOptions.toReport Then
            sw = CreateReportFile(parseOptions.ReportPath)
        End If

        'Retrieve the relevant files
        Dim KeptList As List(Of String) = GetMandatoryFiles()
        Console.WriteLine("** Total Size of remaining files: " & GetSizeOfFileList(KeptList) & " | Number of files: " & KeptList.Count)

        If parseOptions.toReport Then
            If parseOptions.toSimulate Then
#Disable Warning BC42030 ' La variable est transmise par référence avant de se voir attribuer une valeur
                WriteRowInSw(sw, "--- Process Simulation on MSI/MSP removal in " & parseOptions.InputDir)
#Enable Warning BC42030 ' La variable est transmise par référence avant de se voir attribuer une valeur
            End If
            WriteRowInSw(sw, "The following files will be kept in " & parseOptions.InputDir)
            WriteFileListInSw(sw, KeptList)
        End If

        'retrieve the files to be deleted or moved
        Dim toProcessList As List(Of String) = ListFiles(parseOptions.InputDir, KeptList)
        Console.WriteLine("** Total Size of files to be processed: " & GetSizeOfFileList(toProcessList) & " | Number of files: " & toProcessList.Count)

        If parseOptions.toMove And Not parseOptions.toDelete Then
            If GetConfirmation("Confirm move " & toProcessList.Count & " files from " &
                               parseOptions.InputDir & " to " & parseOptions.MoveDir) Then

                If parseOptions.toReport Then
                    WriteRowInSw(sw, "The following files will be moved in " & parseOptions.MoveDir)
                    WriteFileListInSw(sw, toProcessList)
                End If
                MoveListFiles(toProcessList, parseOptions.MoveDir, parseOptions.toSimulate)
            Else
                PressKeyToQuit()
            End If

        ElseIf Not parseOptions.toMove And parseOptions.toDelete Then
            'delete the files
            If GetConfirmation("Confirm Delete " & toProcessList.Count & " files from " &
                              parseOptions.InputDir) Then

                If parseOptions.toReport Then
                    WriteRowInSw(sw, "The following files will be deleted")
                    WriteFileListInSw(sw, toProcessList)
                End If

                DeleteListFiles(toProcessList, parseOptions.toSimulate)
            Else
                PressKeyToQuit()
            End If
        End If

        If parseOptions.toReport Then
            WriteRowInSw(sw, "Detected errors: " & FailedList.Count)
            If parseOptions.toSimulate Then
                WriteRowInSw(sw, "Simulation cannot guarantee there is no error during real process")
            End If

            If FailedList.Count > 0 Then
                WriteListInSw(sw, FailedList)
            End If
            Console.WriteLine("Report written in: " & parseOptions.ReportPath)
        End If

        If sw IsNot Nothing Then
            sw.Flush()
            sw.Close()
        End If

        Console.WriteLine("End of Process")
        PressKeyToQuit()
    End Sub

    ''' <summary>
    ''' Retrieves the MSI and MSP files required by the Windows Installer
    ''' </summary>
    ''' <returns></returns>
    Private Function GetMandatoryFiles() As List(Of String)
        Dim toBeKeptList As New List(Of String)

        Dim msi As Object
        msi = CreateObject("WindowsInstaller.Installer")

        ' Enumerate all products
        Dim products As IEnumerable = CType(msi.Products, IEnumerable)
        For Each productCode In products

            'retrieve the location of the product 
            Dim productName As String = CStr(msi.ProductInfo(productCode, "ProductName"))

            If productName = "" Then
                FailedList.Add("productCode: " & CStr(productCode) & " cannot be retrieved")
            Else
                Dim location As String = CStr(msi.ProductInfo(productCode, "LocalPackage"))
                If File.Exists(location) Then
                    ToBeKeptList.Add(location)
                    Console.WriteLine(location)
                    Console.WriteLine(PropertiesToString(location, msi, " "))

                    Dim patchList As List(Of String)
                    patchList = GetAssociatedPatches(CStr(productCode), msi)

                    'retrieve the pacthes associated to the product
                    If patchList.Count > 0 Then
                        Console.WriteLine("** Patches **")

                        For Each pth In patchList
                            Dim properties As Dictionary(Of String, String)
                            properties = GetMSIProperty(pth, msi)
                            'keep only the patches which have the appropriate allowremoval option
                            If properties.ContainsKey("AllowRemoval") Then
                                Dim allowRemoval As Integer = CInt(properties("AllowRemoval"))
                                If allowRemoval = 0 Then 'if set to 0, the patch cannot be removed, it means the msp can be removed of installer directory
                                    Console.ForegroundColor = ConsoleColor.Gray
                                    Console.WriteLine(PropertiesToString(properties, "   "))
                                    Console.ResetColor()
                                Else
                                    ToBeKeptList.Add(pth)
                                    Console.ForegroundColor = ConsoleColor.Cyan
                                    Console.WriteLine(PropertiesToString(properties, "   "))
                                    Console.ResetColor()
                                End If
                            End If
                        Next
                    End If
                Else
                    Console.ForegroundColor = ConsoleColor.Red
                    Console.WriteLine("File: " & productCode & ", " & location & " does not exist")
                    FailedList.Add("File: " & CStr(productCode) & ", " & location & " does not exist")
                    Console.ResetColor()
                End If
            End If
        Next
        Return ToBeKeptList
    End Function
    ''' <summary>
    ''' Return a list of msp files associated to the product.
    ''' </summary>
    ''' <param name="productCode">ProductCode of the product</param>
    ''' <param name="msi">msi object, if nothing a windowsinstaller object is created</param>
    ''' <returns>List of MSP file pathes</returns>
    Private Function GetAssociatedPatches(productCode As String, ByRef msi As Object) As List(Of String)

        If msi Is Nothing Then
            msi = CreateObject("WindowsInstaller.Installer")
        End If


        Dim toBeKeptPatchList As New List(Of String)

        ' For each product, enumerate its applied patches
        Dim patches As IEnumerable = CType(msi.Patches(productCode), IEnumerable)
        Dim patchCode

        For Each patchCode In patches
            ' Get the local patch location
            Dim patchLocation As String = CStr(msi.PatchInfo(patchCode, "LocalPackage"))

            If File.Exists(patchLocation) Then
                toBeKeptPatchList.Add(patchLocation)
            Else
                Console.ForegroundColor = ConsoleColor.Red
                Console.WriteLine("Patch: " & patchCode & ", " & patchLocation & " does not exist")
                FailedList.Add("Patch: " & CStr(patchCode) & ", " & patchLocation & " does not exist")
                Console.ResetColor()
            End If

        Next
        Return toBeKeptPatchList
    End Function

    ''' <summary>
    ''' Print MSI or MSp properties in a readable format
    ''' </summary>
    ''' <param name="location">Path of the MSI or MSP file</param>
    ''' <param name="msi">msi object, if nothing, a msi object is created</param>
    ''' <param name="offset"> string offset to enhance the print</param>
    ''' <returns>Properties in readable format</returns>
    Private Function PropertiesToString(location As String, ByRef msi As Object, Optional offset As String = "") As String
        If msi Is Nothing Then
            msi = CreateObject("WindowsInstaller.Installer")
        End If

        Dim properties As Dictionary(Of String, String)
        properties = GetMSIProperty(location, msi)
        Return PropertiesToString(properties, offset)
    End Function

    ''' <summary>
    ''' Print properties dictionary in a string format
    ''' </summary>
    ''' <param name="properties"></param>
    ''' <param name="offset">string offset to enhance the print</param>
    ''' <returns></returns>
    Private Function PropertiesToString(properties As Dictionary(Of String, String), Optional offset As String = "") As String

        Dim s As String
        s = offset & String.Join(vbCrLf & offset, properties.Select(Function(kvp) String.Format("{0}={1}", kvp.Key, kvp.Value)).ToArray()) & vbCrLf

        Return s
    End Function

    ''' <summary>
    ''' List the files MSI/MSP files that are not required by the Microsoft Installer
    ''' </summary>
    ''' <param name="folderPath">Directory to look for MSI/MSP files</param>
    ''' <param name="keptFile">Files required by Microsoft Installer</param>
    ''' <returns>List of Files not reauired by Microsoft Installer</returns>
    Private Function ListFiles(folderPath As String, ByRef keptFile As List(Of String)) As List(Of String)
        Dim fullList As New List(Of String)
        Dim pattern(1) As String
        pattern(0) = "*.msi"
        pattern(1) = "*.msp"
        For Each foundFile As String In My.Computer.FileSystem.GetFiles(folderPath,
                                                                        FileIO.SearchOption.SearchAllSubDirectories,
                                                                       pattern)
            If Not keptFile.Contains(foundFile) Then
                fullList.Add(foundFile)
            End If
        Next
        Return fullList
    End Function

    ''' <summary>
    ''' Move a list of file
    ''' </summary>
    ''' <param name="filesList"></param>
    ''' <param name="dirTo"></param>
    ''' <param name="simulate"></param>
    Private Sub MoveListFiles(ByRef filesList As List(Of String), dirTo As String, simulate As Boolean)
        Dim count As UInteger = 1
        For Each f In filesList
            Dim dest As String
            dest = Path.Combine(dirTo, Path.GetFileName(f))


            If Not simulate Then
                Console.ForegroundColor = ConsoleColor.Yellow
                Try
                    Console.WriteLine("(" & count & ") Move " & f & " to " & dest)

                    '**********************************************
                    ' Critical Code
                    'set as comment to avoid any error during debug

                    File.Move(f, dest)

                    ' End of Critical
                    '**********************************************
                    count = count + CUInt(1)
                Catch ex As Exception

                    Console.ForegroundColor = ConsoleColor.Red
                    Console.WriteLine("Unable to move: " & f)
                    Console.ResetColor()

                    If ex.GetType() Is GetType(IO.IOException) Then
                        Console.WriteLine(dest & " file exists")
                        FailedList.Add("Unable to move " & f & " as " & dest & " exists")

                    ElseIf ex.GetType() Is GetType(ArgumentNullException) Then
                        If String.IsNullOrEmpty(f) Then
                            Console.WriteLine("Empty file name")
                            FailedList.Add("Unable to move empty file name")
                        End If
                        If String.IsNullOrEmpty(dest) Then
                            Console.WriteLine("Empty destination file name")
                            FailedList.Add("Unable to move " & f & " to an empty file name")
                        End If

                    ElseIf ex.GetType() Is GetType(ArgumentException) Then
                        If String.IsNullOrEmpty(f) Then
                            Console.WriteLine("Empty file name")
                            FailedList.Add("Unable to move empty file name")
                        Else
                            If String.IsNullOrEmpty(dest) Then
                                Console.WriteLine("Empty destination file name")
                                FailedList.Add("Unable to move " & f & " to an empty file name")
                            Else
                                Console.WriteLine("File name or destination file contain invalid characters")
                                FailedList.Add("Unable to move " & f & " to " & dest & " there is invalid characters")
                            End If
                        End If
                    ElseIf ex.GetType() Is GetType(UnauthorizedAccessException) Then
                        Console.WriteLine("Unauthorized access")
                        FailedList.Add("Unable to move " & f & " to " & dest & " Unauthorized access")
                    ElseIf ex.GetType() Is GetType(PathTooLongException) Then
                        Console.WriteLine("Path tool long")
                        FailedList.Add("Unable to move " & f & " to " & dest & " path is too long")
                    ElseIf ex.GetType() Is GetType(DirectoryNotFoundException) Then
                        Console.WriteLine("Directory not found")
                        FailedList.Add("Unable to move " & f & " to " & dest & " directory does not exist")
                    ElseIf ex.GetType() Is GetType(NotSupportedException) Then
                        Console.WriteLine("Unsupported format")
                        FailedList.Add("Unable to move " & f & " to " & dest & " invalid format")
                    End If
                End Try
            Else

                Console.ForegroundColor = ConsoleColor.Green
                Console.WriteLine("Simulation - (" & count & ") Move " & f & " to " & dest)
                count = count + CUInt(1)
                If String.IsNullOrEmpty(f) Then
                    Console.WriteLine("Empty file name")
                    FailedList.Add("Unable to move empty file name")
                End If
                If String.IsNullOrEmpty(dest) Then
                    Console.WriteLine("Empty destination file name")
                    FailedList.Add("Unable to move " & f & " to an empty file name")
                End If

                If f.Length > MaxPathLength Then
                    Console.WriteLine("Path tool long")
                    FailedList.Add("Unable to move " & f & " to " & dest & " path is too long")
                End If

                If Not Directory.Exists(parseOptions.InputDir) Or Not Directory.Exists(parseOptions.MoveDir) Then
                    Console.WriteLine("Directory not found")
                    FailedList.Add("Unable to move " & f & " to " & dest & " directory does not exist")
                End If
            End If
        Next
        Console.ResetColor()

        Console.WriteLine("Move " & count - CUInt(1) & " Files")
    End Sub

    ''' <summary>
    ''' Delete a list of files
    ''' </summary>
    ''' <param name="filesList"></param>
    ''' <param name="simulate"></param>
    Private Sub DeleteListFiles(ByRef filesList As List(Of String), simulate As Boolean)

        Dim count As UInteger = 1

        For Each f In filesList
            If Not simulate Then
                Try
                    Console.ForegroundColor = ConsoleColor.Yellow
                    Console.WriteLine("(" & count & ") Delete " & f)

                    '**********************************************
                    ' Critical Code
                    'set as comment to avoid any error during debug

                    File.Delete(f)
                    ' End of Critical
                    '**********************************************

                    count = count + CUInt(1)
                Catch ex As Exception
                    FailedList.Add(f)
                    Console.ForegroundColor = ConsoleColor.Red
                    Console.WriteLine("Unable to delete: " & f)
                    Console.ResetColor()
                End Try
            Else

                Console.ForegroundColor = ConsoleColor.Green
                Console.WriteLine("Simulation - (" & count & ") Delete " & f)
                count = count + CUInt(1)

                If String.IsNullOrEmpty(f) Then
                    Console.WriteLine("Empty file name")
                    FailedList.Add("Unable to delete empty file name")
                End If


                If f.Length > MaxPathLength Then
                    Console.WriteLine("Path tool long")
                    FailedList.Add("Unable to delete " & f & " path is too long")
                End If

                If Not Directory.Exists(parseOptions.InputDir) Or Not Directory.Exists(parseOptions.MoveDir) Then
                    Console.WriteLine("Directory not found")
                    FailedList.Add("Unable to delete " & f & " directory does not exist")
                End If

            End If
        Next

        Console.WriteLine("Delete " & count - CUInt(1) & " Files")
    End Sub

    ''' <summary>
    ''' Retrieve the relevant properties of a MSi or MSP files.
    ''' To retrieve the information, the function use the OpenDataBase function of msi object.
    ''' It could have been possible to retrieve some information by using the ProductInfo property.
    ''' Significant information for MSP files come from: https://docs.microsoft.com/fr-fr/windows/win32/msi/msipatchmetadata-table
    ''' General knowledge about the Database: 
    ''' https://docs.microsoft.com/fr-fr/windows/win32/msi/database-tables
    ''' https://docs.microsoft.com/fr-fr/windows/win32/msi/getting-application-information
    ''' https://docs.microsoft.com/fr-fr/windows/win32/msi/installer-object
    ''' </summary>
    ''' <param name="msiFile">Path to the MSI or MSP file</param>
    ''' <param name="msi">msi object, if nothing a msi object is created</param>
    ''' <returns></returns>
    Public Function GetMSIProperty(msiFile As String, Optional ByRef msi As Object = Nothing) As Dictionary(Of String, String)

        Dim results As New Dictionary(Of String, String)
        Dim retVal As String = String.Empty

        Dim sql As String

        If msi Is Nothing Then
            msi = CreateObject("WindowsInstaller.Installer")
        End If

        If Not File.Exists(msiFile) Then
            Return results
        End If

        Dim openMode As Integer
        If Path.GetExtension(msiFile).Contains("msp") Then
            openMode = 32 ' = msiOpenDatabaseModePatchFile
            sql = String.Format("SELECT * FROM MsiPatchMetadata WHERE Property = 'MoreInfoURL'" &
                                " OR Property = 'ManufacturerName'" &
                                " OR Property = 'AllowRemoval'" &
                                " OR Property = 'DisplayName'" &
                                " OR Property = 'Description'" &
                                " OR Property = 'PatchVersion'" &
                                " OR Property = 'EULAID'" &
                                " OR Property = 'TargetProductName' ") 'MsiPatchMetadata, MsiPatchSequence _Tables
        ElseIf Path.GetExtension(msiFile).Contains("msi") Then
            openMode = 0 'msiOpenDatabaseModeReadOnly
            sql = String.Format("SELECT * FROM Property WHERE Property = 'ProductCode' OR Property = 'ProductName' OR Property = 'Manufacturer' OR Property = 'ProductCode'")
        Else
            FailedList.Add("File: " & msiFile & " is not a MSI or MSP file")
            Return results
        End If

        Try
            results("File") = msiFile
            results("Size") = GetSize(msiFile)

            Dim database = msi.OpenDatabase(msiFile, openMode)
            Dim view = database.OpenView(sql)
            view.Execute(Nothing)

            Dim record = view.Fetch()
            While record IsNot Nothing
                If openMode = 0 Then
                    results(CStr(record.StringData(1))) = CStr(record.StringData(2))
                ElseIf openMode = 32 Then
                    results(CStr(record.StringData(2))) = CStr(record.StringData(3))
                End If
                    record = view.Fetch()
            End While

            view.Close()

            Return results
        Catch
            Dim errorrecord = msi.LastErrorRecord()
            FailedList.Add("File: " & msiFile & " Error Code:" & CStr(errorrecord.StringData(1)))
            Return results
        End Try

    End Function

    Private Sub InitParse(ByRef parser As CommandParser)
        parseOptions.InputDir = WIN_DIR_INSTALLER
        parseOptions.MoveDir = DFLT_DIR_MOVE
        parseOptions.ReportPath = DFLT_PATH_REPORT
        parseOptions.toDelete = False
        parseOptions.toMove = False
        parseOptions.toShowHelp = False
        parseOptions.toReport = False
        parseOptions.toSimulate = False


        parser.Argument("d", "delete", "delete the relevant files", "delete",
                CommandArgumentFlags.HideInUsage,
                Sub(p, v)
                    parseOptions.toDelete = True
                    parseOptions.toMove = False
                End Sub
                )

        parser.Argument("m", "move", "Move the relevant files to the specificed directory - default: " &
                        parseOptions.MoveDir, "inputFile",
                CommandArgumentFlags.TakesParameter,
               Sub(p, v)
                   parseOptions.toDelete = False
                   parseOptions.toMove = True
                   If Not String.IsNullOrEmpty(v) Then
                       parseOptions.MoveDir = v
                   End If
               End Sub)

        parser.Argument("s", "simulate", "Simulate the action and write the report in the indicated file [" &
                        parseOptions.ReportPath & "]",
                        "simulate",
                CommandArgumentFlags.TakesParameter,
                 Sub(p, v)
                     parseOptions.toSimulate = True
                     parseOptions.toReport = True
                     If Not String.IsNullOrEmpty(v) Then
                         parseOptions.ReportPath = v
                     End If

                 End Sub)

        parser.Argument("h", "help", "Display this help message", "help",
                CommandArgumentFlags.HideInUsage,
             Sub(p, v)
                 parseOptions.toShowHelp = True
             End Sub)

        parser.Argument("r", "report", "Write the report at the indicated path [" & parseOptions.ReportPath & "]", "report",
                CommandArgumentFlags.TakesParameter,
                Sub(p, v)
                    parseOptions.toReport = True
                    If Not String.IsNullOrEmpty(v) Then
                        parseOptions.ReportPath = v
                    End If

                End Sub)
    End Sub

    Private Sub TestConfigurationAndInit()
        If parseOptions.toMove Then
            If Not Directory.Exists(parseOptions.MoveDir) Then
                Console.ForegroundColor = ConsoleColor.DarkYellow
                Console.WriteLine(parseOptions.MoveDir & "does not exists")
                Console.ResetColor()
                Console.WriteLine("Do you want to create [Y/N] ?")
                Dim rslt = Console.ReadKey()
                While (rslt.KeyChar <> "Y"c And rslt.KeyChar <> "N"c)
                    rslt = Console.ReadKey(True)
                End While
                If rslt.KeyChar <> "Y"c Then
                    Console.WriteLine("Process will stop")
                    PressKeyToQuit()
                Else
                    Console.WriteLine("Create the directory " & parseOptions.MoveDir)
                    Try
                        Directory.CreateDirectory(parseOptions.MoveDir)
                    Catch ex As Exception
                        Console.WriteLine("Unable to create " & parseOptions.MoveDir)
                        PressKeyToQuit(-1)
                    End Try
                End If
            End If
        End If

        If Not parseOptions.toDelete And Not parseOptions.toMove Then
            Console.ForegroundColor = ConsoleColor.Yellow
            Console.WriteLine("No action defined")
            Console.ResetColor()
            Console.WriteLine("Shall select Move or delete.")
            PressKeyToQuit(-2)
        End If

        If parseOptions.toDelete And parseOptions.toMove Then
            Console.ForegroundColor = ConsoleColor.Yellow
            Console.WriteLine("Two actions defined")
            Console.ResetColor()
            Console.WriteLine("Only one action can be perform, select delete or move.")
            PressKeyToQuit(-2)
        End If

        If parseOptions.toDelete Or parseOptions.toMove Then
            Console.ForegroundColor = ConsoleColor.Cyan
            If parseOptions.toSimulate Then
                Console.WriteLine("--- Simulation Mode ---")
            End If
            If parseOptions.toReport Then
                Console.WriteLine("Report will be written in: " & parseOptions.ReportPath)
            End If
            If parseOptions.toDelete Then
                Console.WriteLine("Will Delete the unecessary files set in: " & parseOptions.InputDir)
            End If
            If parseOptions.toMove Then
                Console.WriteLine("Will Move the unecessary files set in: " & parseOptions.InputDir & " to " &
                                  parseOptions.MoveDir)
            End If
            Console.ResetColor()
        End If
    End Sub


    Private Function CreateReportFile(pth As String) As StreamWriter
        If File.Exists(pth) Then
            'first delete the file
            Try
                File.Delete(pth)

            Catch ex As Exception
                Console.WriteLine("Unable to delete the report file: " & pth)
                PressKeyToQuit(-1)
            End Try
        End If

        Try
            Return File.CreateText(pth)
        Catch ex As Exception
            Console.WriteLine("Unable to create the report file: " & pth)
            PressKeyToQuit(-1)
        End Try
        Return Nothing
    End Function

    Private Sub WriteFileListInSw(ByRef sw As StreamWriter, ByRef lst As List(Of String))

        For Each r In lst
            If File.Exists(r) Then
                WriteRowInSw(sw, r)
                WriteRowInSw(sw, PropertiesToString(r, Nothing, "   "))
            Else
                WriteRowInSw(sw, "Unable to find file: " & r)
            End If

        Next r

    End Sub

    Private Sub WriteListInSw(ByRef sw As StreamWriter, ByRef lst As List(Of String))

        For Each r In lst
            WriteRowInSw(sw, r)
        Next r

    End Sub

    Private Sub WriteRowInSw(ByRef sw As StreamWriter, txt As String)
        If sw IsNot Nothing Then
            sw.WriteLine(txt)
        Else
            Console.WriteLine("Unable to Write in the report file.")
            PressKeyToQuit(-1)
        End If
    End Sub

    Private Function GetSize(f As String) As String
        Dim fi As FileInfo
        fi = New FileInfo(f)

        Dim size As Long = fi.Length

        If size = 0 Then Return ""

        Return GetSizeInStr(CType(size, ULong))

    End Function

    Private Function GetSizeInStr(size As ULong) As String
        Dim DoubleBytes As Double
        Try
            Select Case size
                Case Is >= 1099511627776
                    DoubleBytes = CDbl(size / 1099511627776) 'TB
                    Return FormatNumber(DoubleBytes, 2) & " TB"
                Case 1073741824 To 1099511627775
                    DoubleBytes = CDbl(size / 1073741824) 'GB
                    Return FormatNumber(DoubleBytes, 2) & " GB"
                Case 1048576 To 1073741823
                    DoubleBytes = CDbl(size / 1048576) 'MB
                    Return FormatNumber(DoubleBytes, 2) & " MB"
                Case 1024 To 1048575
                    DoubleBytes = CDbl(size / 1024) 'KB
                    Return FormatNumber(DoubleBytes, 2) & " KB"
                Case 0 To 1023
                    DoubleBytes = size ' bytes
                    Return FormatNumber(DoubleBytes, 2) & " bytes"
                Case Else
                    Return ""
            End Select
        Catch
            Return ""
        End Try
    End Function

    Private Function GetSizeOfFileList(fl As List(Of String)) As String
        Dim ul As ULong

        Dim fi As FileInfo
        For Each f In fl
            fi = New FileInfo(f)
            ul = ul + CType(fi.Length, ULong)
        Next

        Return GetSizeInStr(ul)
    End Function

    Private Function GetConfirmation(txt As String) As Boolean
        Console.WriteLine(txt)
        Console.WriteLine("Press [Y] for Yes, [N] for No")
        Dim rslt = Console.ReadKey(True)
        Do While (rslt.KeyChar <> "Y"c And rslt.KeyChar <> "N"c)
            rslt = Console.ReadKey(True)
        Loop

        If rslt.KeyChar <> "Y"c Then
            Return False
        Else
            Return True
        End If
    End Function

    Private Sub PressKeyToQuit(Optional errcode As Integer = 0)
        Console.WriteLine("Will quit - Press a key")
        Console.ReadKey()
        Environment.Exit(errcode)

    End Sub

    ''' <summary>
    ''' Retrieve the Max length of file name.
    ''' Usually it's 260 characters, but can change in regard of the configuration
    ''' Use the reflection to retrieve the correct value.
    ''' Refer to the discussion on https://stackoverflow.com/questions/3406494/what-is-the-maximum-amount-of-characters-or-length-for-a-directory
    ''' </summary>
    ''' <returns></returns>
    Private Function GetMaxPathLength() As Integer

        Dim myFieldInfo As FieldInfo
        Dim myType As Type = GetType(Path)
        ' Get the type and fields of FieldInfoClass.
        myFieldInfo = myType.GetField("MaxPath",
        BindingFlags.Static Or
                BindingFlags.GetField Or
                BindingFlags.NonPublic)

        Return CType(myFieldInfo.GetValue(Nothing), Integer)
    End Function
End Module
