Imports System
Imports System.IO

Module Program
    Private Const HelpString = "Merges two EPAnet file into one. Returns whether or not the merge was successful.

EPAMERGE [drive:][path][filename first file] [drive:][path][filename second file] [drive:][path][filename combined file] [/L:logtype]

  [drive:][path][filename]
               Specifies drive, directory, and/or files to list.
 
  /L           Log results to 'logresults.txt' file.
  logtype       N  No logging; default option, when no parameter option is given.
                C  Compact logging: minimal information on exceptions.
                F  Full logging: detailed information on exceptions, e.g. on duplicate nodes and links.

"

    Private Const ERROR_SUCCESS = 0
    Private Const ERROR_FAILURE = 1

    Private filenames(2) As String
    Private epamanager As New EpanetManager

    Sub Main(args As String())
        If args.Count = 0 Or (args.Count > 0 AndAlso args(0) = "/?") Then
            Console.WriteLine(HelpString)
            Exit Sub
        End If

        If args.Count < 3 Then
            Console.WriteLine("Too few arguments. Type 'EPAMERGE /?' for how to use this command.")
        Else
            filenames(0) = args(0)
            filenames(1) = args(1)
            filenames(2) = args(2)

            If args.Count > 3 Then
                Dim arg = args(3).ToUpper
                If arg.StartsWith("/L") Then
                    arg = arg.Replace(":", "").Replace("/L", "")
                Else
                    Console.WriteLine("Unknown argument. Type 'EPAMERGE /?' for how to use this command.")
                End If

                Select Case arg.First
                    Case "N"
                        epamanager.LogType = LogTypes.NoLogging
                    Case "C"
                        epamanager.LogType = LogTypes.CompactLogging
                    Case "F"
                        epamanager.LogType = LogTypes.FullLogging
                End Select
            End If
        End If

        mergeEpanetFiles()

    End Sub

    Private Sub MergeEpanetFiles()
        Dim result As Integer

        'Load Epanet files 1 and 2 in memory
        For idx = 0 To 1
            result = epamanager.Load(filenames(idx))
            If result = 1 Then
                ShowErrorMessage("load", filenames(idx))
                Environment.Exit(ERROR_FAILURE)
            End If
        Next

        'Write combined file to disk
        result = epamanager.Save(filenames(2))
        If result = 1 Then
            ShowErrorMessage("save", filenames(2))
            Environment.Exit(ERROR_FAILURE)
        Else
            ShowFinishedMessage()
            Environment.Exit(ERROR_SUCCESS)
        End If
    End Sub

    Private Sub ShowErrorMessage(process As String, filename As String)
        Console.WriteLine($"Error: couldn't {process} file '{filename}'.")
        If epamanager.LogType <> LogTypes.NoLogging Then
            Console.WriteLine("See the logfile 'logresults.txt' for more details.")
        End If
    End Sub

    Private Sub ShowFinishedMessage()
        Console.WriteLine($"Successfully created a combined Epanet file.")
        If epamanager.LogType <> LogTypes.NoLogging And epamanager.EventsLogged Then
            Console.WriteLine("See the logfile 'logresults.txt' for events and more details.")
        End If
    End Sub
End Module
