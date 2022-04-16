Imports System
Imports System.IO

Module Program


    Private Const HelpString = "Merges two EPAnet file into one. Returns whether or not the merge was successful.

EPAMERGE [drive:][path][filename first file] [drive:][path][filename second file] [drive:][path][filename combined file] [/L:logtype]

  [drive:][path][filename]
               Specifies drive, directory, and/or files to list.
 
  /L           Log results to 'logresults.txt' file.
  logtype       N  No logging; default option when no parameter option is given.
                C  Compact logging: minimal information on exceptions.
                F  Full logging: detailed information on exceptions, e.g. on duplicate nodes and links.

"

    Private filenames(2) As String
    Private combinedFile As New Dictionary(Of String, List(Of String))
    Private logResults As String = ""
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

            If args.Count > 2 Then
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

    Private Sub mergeEpanetFiles()
        'Load Epanet files 1 and 2 in memory
        For idx = 0 To 1
            loadEpanetFile(filenames(idx))
        Next

        'Write combined file to disk
        saveEpanetFile()

        'Write logfile
        If epamanager.LogType <> LogTypes.NoLogging Then writeLogfile()
    End Sub

    Private Sub loadEpanetFile(filename As String)
        'Get contents of filename and store in combined file
        combinedFile = epamanager.Open(filename)
    End Sub

    Private Sub saveEpanetFile()
        'Save combined file

    End Sub

    Private Sub writeLogfile()
        Dim logfilename = IO.Path.GetFullPath(filenames(2)) + "logresults.txt"

        'Write logresults to file
        File.AppendAllLines(logfilename, epamanager.Logbook)
    End Sub

End Module
