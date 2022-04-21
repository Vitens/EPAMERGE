Imports System
Imports System.IO

Module Program
    Private Const HelpString = "Merges two Epanet files into one. Returns whether or not the merge was successful.

EPAMERGE [drive:][path][source1 filename] [drive:][path][source2 filename] [drive:][path][target filename] [/L:logtype]

  [drive:][path][filename]
               Specifies drive, directory, and source/target files.
 
  /L           Log results to [targetfile].log
  logtype       N  No logging; default option, when no parameter option is given.
                C  Compact logging: minimal information on exceptions.
                F  Full logging: detailed information on exceptions, e.g. on duplicate nodes and links.
"

    Public Const ERROR_SUCCESS = 0
    Public Const ERROR_FAILURE = 1
    Public Const ERROR_FILE_NOT_FOUND = 2
    Public Const ERROR_PATH_NOT_FOUND = 3
    Public Const ERROR_ACCESS_DENIED = 5
    Public Const ERROR_BAD_ARGUMENTS = 160
    Public Const ERROR_BAD_PATHNAME = 161

    Private filenames(2) As String
    Private epaManager As New EpanetManager
    Private logFilename As String

    Sub Main(args As String())
        'Check for valid number of arguments. Exit if not so.
        CheckForValidNumberOfArguments(args)

        'Process arguments
        ProcessArguments(args)

        'All seems fine, so merge the two Epanet files
        MergeEpanetFiles()

    End Sub

    Private Sub CheckForValidNumberOfArguments(args As String())
        If args.Count = 0 Or (args.Count > 0 AndAlso args(0) = "/?") Then
            Console.WriteLine(HelpString)
            Environment.Exit(ERROR_BAD_ARGUMENTS)

        ElseIf args.Count < 3 Then
            Console.WriteLine("Too few arguments specified. Type 'EPAMERGE /?' for how to use this application.")
            Environment.Exit(ERROR_BAD_ARGUMENTS)
        End If
    End Sub

    Private Sub SetFilenames(args As String())
        filenames(0) = args(0)
        filenames(1) = args(1)
        filenames(2) = args(2)
    End Sub

    Private Sub ProcessArguments(args As String())
        'Process filenames
        SetFilenames(args)

        'Process other arguments
        If args.Count > 3 Then
            For idx = 3 To args.Count - 1
                ProcessArgument(args(idx).ToUpper)
            Next
        End If
    End Sub

    Private Sub ProcessArgument(arg As String)
        If arg.StartsWith("/L") Then
            'Process logtype argument

            arg = arg.Replace(":", "").Replace("/L", "")

            Select Case arg.First
                Case "N"
                    epaManager.LogType = LogTypes.NoLogging
                Case "C"
                    epaManager.LogType = LogTypes.CompactLogging
                Case "F"
                    epaManager.LogType = LogTypes.FullLogging
            End Select

            logFilename = filenames(2).Replace(".inp", ".log")
            epaManager.LogFilename = logfilename

        Else
            'Unknown argument specified
            Console.WriteLine("Unknown argument. Type 'EPAMERGE /?' for how to use this command.")
            Environment.Exit(ERROR_BAD_ARGUMENTS)
        End If
    End Sub

    Private Sub MergeEpanetFiles()
        Dim result As Integer

        'Load source files 1 and 2 in memory
        For idx = 0 To 1
            result = epamanager.Load(filenames(idx))
            If result > 0 Then
                ShowErrorMessage("load", filenames(idx))
                Environment.Exit(result)
            End If
        Next

        'Write target file to disk
        result = epamanager.Save(filenames(2))
        If result > 0 Then
            ShowErrorMessage("save", filenames(2))
            Environment.Exit(result)
        Else
            ShowFinishedMessage()
        End If
    End Sub

    Private Sub ShowErrorMessage(process As String, filename As String)
        Console.WriteLine($"Error: couldn't {process} file '{filename}'.")
        If epamanager.LogType <> LogTypes.NoLogging Then
            Console.WriteLine("See the logfile 'logresults.txt' for details.")
        End If
    End Sub

    Private Sub ShowFinishedMessage()
        Console.WriteLine($"Successfully created a combined Epanet file.")
        If epamanager.LogType <> LogTypes.NoLogging And epamanager.EventsLogged Then
            Console.WriteLine($"See the logfile '{logFilename}' for details.")
        End If
    End Sub
End Module
