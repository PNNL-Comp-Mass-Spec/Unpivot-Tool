Option Strict On

Imports System.IO
Imports System.Reflection
Imports System.Threading
Imports System.Windows.Forms
Imports PRISM

' This program uses clsFileUnpivoter to read in a tab delimited file that is in a crosstab / pivottable format
' and writes out a new file where the data has been unpivotted
'
' Example command Line
' /I:DataFile.txt /F:FixedColumnCount /O:OutputDirectoryPath

' -------------------------------------------------------------------------------
' Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
' Program started July 20, 2009
'
' E-mail: matthew.monroe@pnl.gov or matt@alchemistmatt.com
' Website: http://ncrr.pnl.gov/ or http://www.sysbio.org/resources/staff/
' -------------------------------------------------------------------------------
'
' Licensed under the Apache License, Version 2.0; you may not use this file except
' in compliance with the License.  You may obtain a copy of the License at
' http://www.apache.org/licenses/LICENSE-2.0
'
' Notice: This computer software was prepared by Battelle Memorial Institute,
' hereinafter the Contractor, under Contract No. DE-AC05-76RL0 1830 with the
' Department of Energy (DOE).  All rights in the computer software are reserved
' by DOE on behalf of the United States Government and the Contractor as
' provided in the Contract.  NEITHER THE GOVERNMENT NOR THE CONTRACTOR MAKES ANY
' WARRANTY, EXPRESS OR IMPLIED, OR ASSUMES ANY LIABILITY FOR THE USE OF THIS
' SOFTWARE.  This notice including this sentence must appear on any copies of
' this computer software.

Public Module modMain
    ' Ignore Spelling: unpivoter, unpivotted, Sep, unpivotting, pivottable

    Public Const PROGRAM_DATE As String = "February 11, 2021"

    Private mInputFilePath As String
    Private mOutputDirectoryName As String          ' Optional
    Private mParameterFilePath As String            ' Optional; not used by clsFileUnpivoter

    Private mFixedColumnCount As Integer
    Private mSkipBlankItems As Boolean
    Private mSkipNullItems As Boolean
    Private mTabDelimitedFile As Boolean
    Private mColumnSepChar As Char

    Private mOutputDirectoryAlternatePath As String                ' Optional
    Private mRecreateDirectoryHierarchyInAlternatePath As Boolean  ' Optional

    Private mRecurseDirectories As Boolean
    Private mRecurseDirectoriesMaxLevels As Integer

    Private mLogMessagesToFile As Boolean

    Private mLastProgressReportTime As DateTime
    Private mLastProgressReportValue As Integer

    Private Sub DisplayProgressPercent(percentComplete As Integer, addCarriageReturn As Boolean)
        If addCarriageReturn Then
            Console.WriteLine()
        End If
        If percentComplete > 100 Then percentComplete = 100
        Console.Write("Processing: " & percentComplete.ToString & "% ")
        If addCarriageReturn Then
            Console.WriteLine()
        End If
    End Sub

    Public Function Main() As Integer
        ' Returns 0 if no error, error code if an error
        Dim returnCode As Integer
        Dim commandLineParser As New clsParseCommandLine
        Dim proceed As Boolean

        returnCode = 0
        mInputFilePath = String.Empty
        mOutputDirectoryName = String.Empty
        mParameterFilePath = String.Empty

        mFixedColumnCount = 1
        mSkipBlankItems = False
        mSkipNullItems = False

        mTabDelimitedFile = True
        mColumnSepChar = ControlChars.Tab

        mOutputDirectoryAlternatePath = String.Empty
        mRecreateDirectoryHierarchyInAlternatePath = False

        mRecurseDirectories = False
        mRecurseDirectoriesMaxLevels = 0

        mLogMessagesToFile = False

        Try
            proceed = False
            If commandLineParser.ParseCommandLine Then
                If SetOptionsUsingCommandLineParameters(commandLineParser) Then proceed = True
            End If

            If Not proceed OrElse commandLineParser.NeedToShowHelp OrElse mInputFilePath.Length = 0 Then
                ShowProgramHelp()
                returnCode = -1
            Else
                Try
                    Dim unpivoter = New FileUnpivoter
                    AddHandler unpivoter.ProgressUpdate, AddressOf Unpivoter_ProgressChanged
                    AddHandler unpivoter.ProgressReset, AddressOf Unpivoter_ProgressReset

                    With unpivoter
                        .LogMessagesToFile = mLogMessagesToFile

                        .FixedColumnCount = mFixedColumnCount
                        .SkipBlankItems = mSkipBlankItems
                        .SkipNullItems = mSkipNullItems

                        If Not mTabDelimitedFile Then
                            .ColumnSepChar = mColumnSepChar
                        End If

                    End With

                    If mRecurseDirectories Then
                        If unpivoter.ProcessFilesAndRecurseDirectories(mInputFilePath, mOutputDirectoryName, mOutputDirectoryAlternatePath, mRecreateDirectoryHierarchyInAlternatePath, mParameterFilePath, mRecurseDirectoriesMaxLevels) Then
                            returnCode = 0
                        Else
                            returnCode = unpivoter.ErrorCode
                        End If
                    Else
                        If unpivoter.ProcessFilesWildcard(mInputFilePath, mOutputDirectoryName, mParameterFilePath) Then
                            returnCode = 0
                        Else
                            returnCode = unpivoter.ErrorCode
                            If returnCode <> 0 Then
                                ConsoleMsgUtils.ShowWarning("Error while processing: " & unpivoter.GetErrorMessage())
                            End If
                        End If
                    End If

                Catch ex As Exception
                    Console.WriteLine("Error initializing File Unpivoter " & ex.Message)
                End Try

                DisplayProgressPercent(mLastProgressReportValue, True)
            End If

        Catch ex As Exception
            ConsoleMsgUtils.ShowError("Error occurred in modMain->Main", ex)
            returnCode = -1
        End Try

        Return returnCode

    End Function

    Private Function SetOptionsUsingCommandLineParameters(commandLineParser As clsParseCommandLine) As Boolean
        ' Returns True if no problems; otherwise, returns false
        ' /I:PeptideInputFilePath /R: ProteinInputFilePath /O:OutputFolderPath /P:ParameterFilePath

        Dim value As String = String.Empty
        Dim validParameters = New String() {"I", "O", "F", "C", "B", "N", "S", "A", "R", "L", "Q"}

        Dim result As Integer

        Try
            ' Make sure no invalid parameters are present
            If commandLineParser.InvalidParametersPresent(validParameters) Then
                Return False
            Else
                With commandLineParser

                    If .NonSwitchParameterCount > 0 Then
                        ' Treat the first non-switch parameter as the input file
                        mInputFilePath = .RetrieveNonSwitchParameter(0)
                    End If

                    ' Query commandLineParser to see if various parameters are present
                    If .RetrieveValueForParameter("I", value) Then mInputFilePath = value
                    If .RetrieveValueForParameter("O", value) Then mOutputDirectoryName = value

                    If .RetrieveValueForParameter("F", value) Then
                        If Integer.TryParse(value, result) Then
                            mFixedColumnCount = result
                        Else
                            Console.WriteLine("Error parsing /F parameter; should be an integer")
                        End If
                    End If

                    If .RetrieveValueForParameter("C", value) Then
                        ' The user has defined a column separator character
                        mTabDelimitedFile = False
                        If value.ToLower = "space" Then
                            mColumnSepChar = " "c
                        ElseIf value.ToLower = "tab" Then
                            mColumnSepChar = ControlChars.Tab
                        Else
                            mColumnSepChar = value.Chars(0)
                        End If

                    End If

                    If .RetrieveValueForParameter("B", value) Then mSkipBlankItems = True
                    If .RetrieveValueForParameter("N", value) Then mSkipNullItems = True

                    If .RetrieveValueForParameter("S", value) Then
                        mRecurseDirectories = True
                        If IsNumeric(value) Then
                            mRecurseDirectoriesMaxLevels = CInt(value)
                        End If
                    End If
                    If .RetrieveValueForParameter("A", value) Then mOutputDirectoryAlternatePath = value
                    If .RetrieveValueForParameter("R", value) Then mRecreateDirectoryHierarchyInAlternatePath = True

                    If .RetrieveValueForParameter("L", value) Then mLogMessagesToFile = True

                End With

                Return True
            End If

        Catch ex As Exception
            ConsoleMsgUtils.ShowError("Error parsing the command line parameters", ex)
            Return False
        End Try

    End Function

    Private Sub ShowProgramHelp()

        Try
            Console.WriteLine("This program reads in a delimited text file that is in crosstab (aka pivot table) format and writes out a new file where the data has been unpivotted.")
            Console.WriteLine()
            Console.WriteLine("Program syntax:" & ControlChars.NewLine & Path.GetFileName(Assembly.GetExecutingAssembly().Location) &
                              " /I:InputFilePath [/O:OutputDirectoryName]")
            Console.WriteLine(" [/F:FixedColumnCount] [/C:ColumnSepChar] [/B] [/N]")
            Console.WriteLine(" [/S:[MaxLevel]] [/A:AlternateOutputDirectoryPath] [/R] [/L] [/Q]")
            Console.WriteLine()

            Console.WriteLine("The input file path can contain the wildcard character *.  If a wildcard is present, then all matching files will be processed")
            Console.WriteLine("The output directory name is optional.  If omitted, the output files will be created in the same directory as the input file.  If included, then a subdirectory is created with the name OutputDirectoryName.")
            Console.WriteLine()

            Console.WriteLine("Use /F to define the number of fixed columns (default is /F:1).  When unpivotting, data in these columns will be written to every row in the output file.")
            Console.WriteLine("The default column separation character is the tab character.  Use /C to define an alternate character.  For example, use /C:, for a comma.  For a space, use /C:space")
            Console.WriteLine("Use /B to skip writing blank column values to the output file")
            Console.WriteLine("Use /N to skip writing Null values to the output file (as indicated by the word 'null').")
            Console.WriteLine()

            Console.WriteLine("Use /S to process all valid files in the input directory and subdirectories. Include a number after /S (like /S:2) to limit the level of subdirectories to examine.")
            Console.WriteLine("When using /S, you can redirect the output of the results using /A.")
            Console.WriteLine("When using /S, you can use /R to re-create the input directory hierarchy in the alternate output directory (if defined).")
            Console.WriteLine()
            Console.WriteLine("Use /L to log messages to a file.  Use the optional /Q switch will suppress all error messages.")
            Console.WriteLine()

            Console.WriteLine("Program written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA) in 2009")
            Console.WriteLine()

            Console.WriteLine("This is version " & Application.ProductVersion & " (" & PROGRAM_DATE & ")")
            Console.WriteLine()

            Console.WriteLine("E-mail: matthew.monroe@pnl.gov or matt@alchemistmatt.com")
            Console.WriteLine("Website: http://ncrr.pnl.gov/ or http://www.sysbio.org/resources/staff/")
            Console.WriteLine()

            Console.WriteLine("Licensed under the Apache License, Version 2.0; you may not use this file except in compliance with the License.  " &
                              "You may obtain a copy of the License at http://www.apache.org/licenses/LICENSE-2.0")
            Console.WriteLine()

            Console.WriteLine("Notice: This computer software was prepared by Battelle Memorial Institute, " &
                              "hereinafter the Contractor, under Contract No. DE-AC05-76RL0 1830 with the " &
                              "Department of Energy (DOE).  All rights in the computer software are reserved " &
                              "by DOE on behalf of the United States Government and the Contractor as " &
                              "provided in the Contract.  NEITHER THE GOVERNMENT NOR THE CONTRACTOR MAKES ANY " &
                              "WARRANTY, EXPRESS OR IMPLIED, OR ASSUMES ANY LIABILITY FOR THE USE OF THIS " &
                              "SOFTWARE.  This notice including this sentence must appear on any copies of " &
                              "this computer software.")
            Console.WriteLine()

            ' Delay for 750 msec in case the user double clicked this file from within Windows Explorer (or started the program via a shortcut)
            Thread.Sleep(750)

        Catch ex As Exception
            Console.WriteLine("Error displaying the program syntax: " & ex.Message)
        End Try

    End Sub

    Private Sub Unpivoter_ProgressChanged(taskDescription As String, percentComplete As Single)
        Const PERCENT_REPORT_INTERVAL = 25
        Const PROGRESS_DOT_INTERVAL_MSEC = 250

        If percentComplete >= mLastProgressReportValue Then
            If mLastProgressReportValue > 0 Then
                Console.WriteLine()
            End If
            DisplayProgressPercent(mLastProgressReportValue, False)
            mLastProgressReportValue += PERCENT_REPORT_INTERVAL
            mLastProgressReportTime = DateTime.UtcNow
        Else
            If DateTime.UtcNow.Subtract(mLastProgressReportTime).TotalMilliseconds > PROGRESS_DOT_INTERVAL_MSEC Then
                mLastProgressReportTime = DateTime.UtcNow
                Console.Write(".")
            End If
        End If
    End Sub

    Private Sub Unpivoter_ProgressReset()
        mLastProgressReportTime = DateTime.UtcNow
        mLastProgressReportValue = 0
    End Sub
End Module
