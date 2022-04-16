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
    Public ReadOnly Property EventsLogged As Boolean

    Public Sub New()
        network = New Dictionary(Of String, List(Of String))
        processedItems = New Dictionary(Of String, HashSet(Of String))
        logbook = New List(Of String)

        sections = New List(Of String)
        InitSectionsList()
    End Sub

    Public Function Load(ByVal filename As String) As Integer
        Try
            If Not IsEpanetFile(filename) Then End

            SetLogbookHeader(filename)

            'Process the file and all sections in the Epanet INP file
            ProcessFile(filename)

            filesProcessed += 1
            Return 0

        Catch ex As Exception
            AddFatalError($"Processing file '{filename}' failed!", ex.Message)
            SaveLogbook(filename)

            Return 1
        End Try
    End Function

    Public Function Save(ByVal filename As String) As Integer
        Dim savingLogbook As Boolean = False

        Try
            'Delete old file, if it exists
            File.Delete(filename)

            'Write new combined file
            For Each s In sections
                File.AppendAllLines(filename, network(s))
            Next

            savingLogbook = True
            SaveLogbook(filename)

            Return 0

        Catch ex As Exception
            Console.WriteLine("*** FATAL ERROR ***")
            If Path.GetExtension(filename).ToLower = ".inp" Then
                Console.WriteLine($"*** Saving the combined Epanet file '{filename}' failed!  {ex.Message} ***")
            Else
                Console.WriteLine($"*** Saving the logbook '{filename}' failed!  {ex.Message} ***")
            End If

            If Not savingLogbook Then
                AddFatalError($"Saving the combined Epanet file '{filename}' failed!", ex.Message)
                SaveLogbook(filename)
            End If

            Return 1
        End Try
    End Function

    Private Function IsEpanetFile(filename As String) As Boolean
        If Path.GetExtension(filename).ToLower <> ".inp" Then
            Console.WriteLine($"Invalid filetype: {Path.GetFileName(filename)}. EPAMERGE can only process .INP files.")
            Return False
        Else
            Return True
        End If
    End Function

    Private Sub SetLogbookHeader(filename)
        Dim msg As String = ""

        If LogType <> LogTypes.NoLogging Then
            logbook.Add($"Logbook created: {Now.ToString("f")}")
            logbook.Add("")

            If filesProcessed = 0 Then
                msg = "Combined network from two Epanet files:"
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
        End If
    End Sub

    Private Sub SaveLogbook(filename As String)
        If LogType <> LogTypes.NoLogging Then
            Dim logfilename = IO.Path.GetDirectoryName(filename) + "\logresults.txt"

            'Delete old file, if it exists
            File.Delete(logfilename)

            'Write logresults to file
            File.AppendAllLines(logfilename, logbook)
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
            Case "Title", "End"
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
                    If Not logbook.Contains(msg) Then AddEvent(msg)
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

    Private Sub AddEvent(title As String)
        logbook.Add(title)
        _EventsLogged = True
    End Sub

    Private Sub AddFatalError(title As String, exceptionMessage As String)
        logbook.Add("----------------------------------------------")
        logbook.Add("*** FATAL ERROR ***")
        logbook.Add(title)
        logbook.Add(exceptionMessage)
        _EventsLogged = True
    End Sub
End Class
