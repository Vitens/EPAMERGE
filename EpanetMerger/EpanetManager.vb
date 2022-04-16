Imports System.IO

Public Enum LogTypes
    NoLogging
    CompactLogging
    FullLogging
End Enum

Public Class EpanetManager
    Private epanetFilename As String
    Private section As String
    Private sections As List(Of String)
    Private filesProcessed As Integer

    Public Logbook As List(Of String)
    Public LogType As LogTypes = LogTypes.NoLogging
    Public Network As Dictionary(Of String, List(Of String))

    Public Sub New()
        sections = New List(Of String)
        InitSectionsList()

        Logbook = New List(Of String)
        Network = New Dictionary(Of String, List(Of String))
    End Sub

    Public Function Open(ByVal filename As String) As Dictionary(Of String, List(Of String))
        Try
            epanetFilename = filename

            'Process the file and all sections in the Epanet INP file
            ProcessFile()
            filesProcessed += 1

            Return Network

        Catch ex As Exception
            Logbook.Add("*** FATAL ERROR ***")
            Logbook.Add($"*** Processing file '{filename}' failed!  {ex.Message} ***")
            Return Nothing
        End Try
    End Function

    Private Sub ProcessFile()
        Dim str As String() = File.ReadAllLines(epanetFilename)
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
                ProcessLine(line)
            End If
        Next
    End Sub

    Private Sub GetSection(ByVal line As String)
        section = (line.ToLower.Replace("[", "").Replace("]", "").Trim).ToCamelCase
    End Sub

    Private Sub ProcessLine(ByVal line As String)
        line = line.Trim

        Select Case section.ToUpper
            Case "TIMES", "REPORT", "OPTIONS", "BACKDROP", "END"
                'skip for second file
                If filesProcessed = 1 Then Exit Sub

            Case Else
                'Check for duplicate items
                If filesProcessed = 1 Then CheckForDuplicates(line)
        End Select
        Network(section).Add(line)
    End Sub

    Private Sub CheckForDuplicates(ByVal line As String)
        Dim result = SectionContainsItem(line)
        Dim item = result.Item1
        Dim duplicateItem = result.Item2

        If duplicateItem And LogType = LogTypes.FullLogging Then
            Logbook.Add($"Duplicate item In section '{section}': {item}")
        End If
    End Sub

    Private Function SectionContainsItem(line As String) As (String, Boolean)
        Dim item1 As String = line.Trim.Split(" ")(0).ToUpper

        For Each l In Network(section)
            Dim item2 As String = l.Trim.Split(" ")(0).ToUpper
            If item1 = item2 Then Return (item1, True)
            Exit For
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
            Network.Add(s, New List(Of String))
        Next
    End Sub
End Class
