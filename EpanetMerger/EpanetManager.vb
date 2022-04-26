Imports System.IO

Public Enum LogTypes
    NoLogging
    CompactLogging
    FullLogging
End Enum

Public Class EpanetManager
    Private network As Dictionary(Of String, List(Of String))
    Private processedItems As Dictionary(Of String, HashSet(Of String))

    Private section As String
    Private sections As List(Of String)
    Private filesProcessed As Integer
    Private logbook As List(Of String)

    Public Property LogType As LogTypes = LogTypes.NoLogging
    Public Property LogFilename As String
    Public ReadOnly Property EventsLogged As Boolean

    Public Sub New()
        network = New Dictionary(Of String, List(Of String))
        processedItems = New Dictionary(Of String, HashSet(Of String))
        logbook = New List(Of String)

        sections = New List(Of String)
        InitSectionsList()
    End Sub

    ''' <summary>
    ''' Adds all information the source Epanet file to a list of sections.
    ''' If a node or link already exists, then this information is skipped. Also
    ''' these elements are notes in the logbook - if the log option is enabled.
    ''' </summary>
    ''' <param name="filename">Name of the source Epanet file</param>
    ''' <returns>An exit code: 0 means success, > 0 means an error occurred.</returns>
    Public Function Load(ByVal filename As String) As Integer
        Try
            If Not IsEpanetFile(filename) Then Return ERROR_FAILURE

            'Write the source and target files in the logbook
            SetLogbookHeader(filename)

            'Process the file and all sections in the Epanet INP file
            ProcessFile(filename)

            filesProcessed += 1
            Return ERROR_SUCCESS

        Catch ex As DirectoryNotFoundException
            ProcessFatalError($"Path of the specified file '{filename}' could not be found!", ex)
            Return ERROR_PATH_NOT_FOUND

        Catch ex As FileNotFoundException
            ProcessFatalError($"Couldn't find the specified file '{filename}'!", ex)
            Return ERROR_FILE_NOT_FOUND

        Catch ex As AccessViolationException
            ProcessFatalError($"Couldn't access the path or file specified: '{filename}'!", ex)
            Return ERROR_ACCESS_DENIED

        Catch ex As Exception
            ProcessFatalError($"Loading file '{filename}' failed!", ex)
            Return ERROR_FAILURE
        End Try
    End Function

    ''' <summary>
    ''' Save all information from the source files to the target file
    ''' </summary>
    ''' <param name="filename">Name of the target Epanet file.</param>
    ''' <returns>An exit code: 0 means success, > 0 means an error occurred.</returns>
    Public Function Save(ByVal filename As String) As Integer
        Try
            'Delete old file, if it exists
            File.Delete(filename)

            'Write new combined file
            For Each s In sections
                File.AppendAllLines(filename, network(s))
            Next

            'Save logbook
            If LogType <> LogTypes.NoLogging Then SaveLogbook()

            Return ERROR_SUCCESS

        Catch ex As DirectoryNotFoundException
            ProcessFatalError($"Path of the specified file '{filename}' could not be found!", ex)
            Return ERROR_PATH_NOT_FOUND

        Catch ex As PathTooLongException
            ProcessFatalError($"Path of the specified file '{filename}' is to long!", ex)
            Return ERROR_BAD_PATHNAME

        Catch ex As AccessViolationException
            ProcessFatalError($"Couldn't access the path or file specified: '{filename}'!", ex)
            Return ERROR_ACCESS_DENIED

        Catch ex As Exception
            ProcessFatalError($"Saving file '{filename}' failed!", ex)
            Return ERROR_FAILURE

        End Try
    End Function

    Private Sub SetLogbookHeader(filename)
        Dim msg As String = ""

        If filesProcessed = 0 Then
            logbook.Add($"Logbook created: {Now.ToString("f")}")
            logbook.Add("")

            msg = "Combining the following two Epanet files:"
            network("Title").Add(msg)
            logbook.Add("----------------------------------------------")
            logbook.Add(msg)

            msg = $"File 1: {Path.GetFileName(filename)}"
            network("Title").Add(msg)
            logbook.Add(msg)
        Else
            msg = $"File 2: {Path.GetFileName(filename)}"
            network("Title").Add(msg)
            logbook.Add(msg)
            logbook.Add("----------------------------------------------")
        End If
    End Sub

    Private Sub ProcessFatalError(title As String, ex As Exception)
        AddToLogbook(title, ex.Message)

        'Save logbook
        If LogType <> LogTypes.NoLogging Then SaveLogbook()
    End Sub

    Private Sub AddToLogbook(title As String, Optional exceptionMessage As String = Nothing)
        logbook.Add(title)
        If exceptionMessage IsNot Nothing Then
            logbook.Add($"System error message: {exceptionMessage}")
        End If

        _EventsLogged = True
    End Sub

    Private Sub SaveLogbook()
        If LogType <> LogTypes.NoLogging And LogFilename <> "" Then
            'Delete old file, if it exists
            File.Delete(LogFilename)

            'Write logresults to file
            File.AppendAllLines(LogFilename, logbook)
        End If
    End Sub

    Private Sub ProcessFile(filename As String)
        Dim str As String() = File.ReadAllLines(filename)
        Dim lines As List(Of String) = str.ToList

        'Remove all empty lines
        lines.RemoveAll(Function(s) s.Trim.Length = 0)

        'process the sections in the Epanet INP file
        ProcessHeaders(lines)
    End Sub

    Private Sub ProcessHeaders(ByVal lines As List(Of String))
        'Look for headers and otherwise process the lines
        Dim totalcount = lines.Count

        For Each line As String In lines
            If line.StartsWith("[") Then
                GetSection(line)
            Else
                If section <> "" Then
                    ProcessLine(line)
                End If
            End If
        Next
    End Sub

    Private Sub GetSection(ByVal line As String)
        section = (line.ToLower.Replace("[", "").Replace("]", "").Trim).ToCamelCase
    End Sub

    Private Sub ProcessLine(ByVal line As String)
        line = line.Trim

        'Don't add comment lines twice
        If filesProcessed = 1 And line.StartsWith(";") Then Exit Sub

        Select Case section
            Case "Title", "Backdrop", "End"
                'Don't process these sections
                Exit Sub

            Case "Times", "Report", "Options", "Energy", "Reactions", "Backdrop"
                'Only add these sections for the first file
                If filesProcessed = 1 Then Exit Sub

            Case "Labels", "Rules", "Controls"
                'Add the labels of both files

            Case Else
                'Add the items from both files as long as they are unique 
                If ContainsDuplicateItem(line) Then Exit Sub
        End Select
        network(section).Add(line)
    End Sub

    Private Function ContainsDuplicateItem(ByVal line As String) As Boolean
        Dim item = line.Trim.Split(" ")(0)

        If Not processedItems.ContainsKey(section) Then
            processedItems.Add(section, New HashSet(Of String))
        End If

        If filesProcessed = 0 Then
            processedItems(section).Add(item)
        Else
            If processedItems(section).Contains(item) Then
                If LogType = LogTypes.FullLogging Then
                    Dim msg = $"Duplicate item In section '{section}': {item}"
                    If Not logbook.Contains(msg) Then AddToLogbook(msg)
                    Return True
                End If
            End If
        End If
        Return False
    End Function

    Private Function SectionContainsItem(line As String) As (String, Boolean)
        Dim item1 As String

        item1 = line.Trim.Split(" ")(0)

        For Each l In network(section)
            If l.Length = 0 Or l = $"[{section}]" Then
                'skip this line
            Else
                If item1 = l.Trim.Split(" ")(0) Then Return (item1, True)
            End If
        Next
        Return (item1, False)
    End Function

    Private Sub InitSectionsList()
        sections.Add("Title")
        sections.Add("Junctions")
        sections.Add("Reservoirs")
        sections.Add("Tanks")
        sections.Add("Pipes")
        sections.Add("Pumps")
        sections.Add("Valves")
        sections.Add("Tags")
        sections.Add("Demands")
        sections.Add("Status")
        sections.Add("Patterns")
        sections.Add("Curves")
        sections.Add("Controls")
        sections.Add("Rules")
        sections.Add("Energy")
        sections.Add("Emitters")
        sections.Add("Quality")
        sections.Add("Sources")
        sections.Add("Reactions")
        sections.Add("Mixing")
        sections.Add("Times")
        sections.Add("Report")
        sections.Add("Options")
        sections.Add("Coordinates")
        sections.Add("Vertices")
        sections.Add("Labels")
        sections.Add("Backdrop")
        sections.Add("End")

        For Each s In sections
            network.Add(s, New List(Of String))
            If s <> "Title" Then network(s).Add("")
            network(s).Add($"[{s}]")
        Next
    End Sub

    Private Function IsEpanetFile(filename As String) As Boolean
        If Path.GetExtension(filename).ToLower <> ".inp" Then
            Console.WriteLine($"Invalid filetype: {Path.GetFileName(filename)}. EPAMERGE can only process .INP files.")
            Return False
        Else
            Return True
        End If
    End Function
End Class
