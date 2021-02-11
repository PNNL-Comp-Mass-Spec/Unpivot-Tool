Option Strict On

Imports PRISM
' This program uses clsFileUnpivoter to read in a tab delimited file that is in a crosstab / pivottable format
' and writes out a new file where the data has been unpivotted
'
' Example command Line
' /I:DataFile.txt /F:FixedColumnCount /O:OutputFolderPath

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

    Public Const PROGRAM_DATE As String = "February 11, 2021"

    Private mInputFilePath As String
    Private mOutputFolderName As String             ' Optional
    Private mParameterFilePath As String            ' Optional; not used by clsFileUnpivoter

    Private mFixedColumnCount As Integer
    Private mSkipBlankItems As Boolean
    Private mSkipNullItems As Boolean
    Private mTabDelimitedFile As Boolean
    Private mColumnSepChar As Char

    Private mOutputFolderAlternatePath As String                ' Optional
    Private mRecreateFolderHierarchyInAlternatePath As Boolean  ' Optional

    Private mRecurseFolders As Boolean
    Private mRecurseFoldersMaxLevels As Integer

    Private mLogMessagesToFile As Boolean
    Private mQuietMode As Boolean = False

    Private WithEvents mUnpivoter As clsFileUnpivoter
    Private mLastProgressReportTime As System.DateTime
    Private mLastProgressReportValue As Integer

    Private Sub DisplayProgressPercent(ByVal intPercentComplete As Integer, ByVal blnAddCarriageReturn As Boolean)
        If blnAddCarriageReturn Then
            Console.WriteLine()
        End If
        If intPercentComplete > 100 Then intPercentComplete = 100
        Console.Write("Processing: " & intPercentComplete.ToString & "% ")
        If blnAddCarriageReturn Then
            Console.WriteLine()
        End If
    End Sub

    Public Function Main() As Integer
        ' Returns 0 if no error, error code if an error
        Dim intReturnCode As Integer
        Dim objParseCommandLine As New clsParseCommandLine
        Dim blnProceed As Boolean
        Dim blnSuccess As Boolean

        intReturnCode = 0
        mInputFilePath = String.Empty
        mOutputFolderName = String.Empty
        mParameterFilePath = String.Empty

        mFixedColumnCount = 1
        mSkipBlankItems = False
        mSkipNullItems = False

        mTabDelimitedFile = True
        mColumnSepChar = ControlChars.Tab

        mOutputFolderAlternatePath = String.Empty
        mRecreateFolderHierarchyInAlternatePath = False

        mRecurseFolders = False
        mRecurseFoldersMaxLevels = 0

        mQuietMode = False
        mLogMessagesToFile = False

        Try
            blnProceed = False
            If objParseCommandLine.ParseCommandLine Then
                If SetOptionsUsingCommandLineParameters(objParseCommandLine) Then blnProceed = True
            End If

            If Not blnProceed OrElse objParseCommandLine.NeedToShowHelp OrElse mInputFilePath.Length = 0 Then
                ShowProgramHelp()
                intReturnCode = -1
            Else
                Try
                    mUnpivoter = New clsFileUnpivoter

                    With mUnpivoter
                        .ShowMessages = Not mQuietMode
                        .LogMessagesToFile = mLogMessagesToFile

                        .FixedColumnCount = mFixedColumnCount
                        .SkipBlankItems = mSkipBlankItems
                        .SkipNullItems = mSkipNullItems

                        If Not mTabDelimitedFile Then
                            .ColumnSepChar = mColumnSepChar
                        End If

                    End With

                    If mRecurseFolders Then
                        If mUnpivoter.ProcessFilesAndRecurseDirectories(mInputFilePath, mOutputFolderName, mOutputFolderAlternatePath, mRecreateFolderHierarchyInAlternatePath, mParameterFilePath, mRecurseFoldersMaxLevels) Then
                            intReturnCode = 0
                        Else
                            intReturnCode = mUnpivoter.ErrorCode
                        End If
                    Else
                        If mUnpivoter.ProcessFilesWildcard(mInputFilePath, mOutputFolderName, mParameterFilePath) Then
                            intReturnCode = 0
                        Else
                            intReturnCode = mUnpivoter.ErrorCode
                            If intReturnCode <> 0 AndAlso Not mQuietMode Then
                                Console.WriteLine("Error while processing: " & mUnpivoter.GetErrorMessage())
                            End If
                        End If
                    End If

                Catch ex As Exception
                    blnSuccess = False
                    Console.WriteLine("Error initializing File Unpivoter " & ex.Message)
                End Try

                DisplayProgressPercent(mLastProgressReportValue, True)
            End If

        Catch ex As Exception
            If mQuietMode Then
                Throw ex
            Else
                Console.WriteLine("Error occurred in modMain->Main: " & ControlChars.NewLine & ex.Message)
            End If
            intReturnCode = -1
        End Try

        Return intReturnCode

    End Function


    Private Function SetOptionsUsingCommandLineParameters(ByVal objParseCommandLine As clsParseCommandLine) As Boolean
        ' Returns True if no problems; otherwise, returns false
        ' /I:PeptideInputFilePath /R: ProteinInputFilePath /O:OutputFolderPath /P:ParameterFilePath

        Dim strValue As String = String.Empty
        Dim strValidParameters() As String = New String() {"I", "O", "F", "C", "B", "N", "S", "A", "R", "L", "Q"}

        Dim intResult As Integer

        Try
            ' Make sure no invalid parameters are present
            If objParseCommandLine.InvalidParametersPresent(strValidParameters) Then
                Return False
            Else
                With objParseCommandLine

                    If .NonSwitchParameterCount > 0 Then
                        ' Treat the first non-switch parameter as the input file
                        mInputFilePath = .RetrieveNonSwitchParameter(0)
                    End If

                    ' Query objParseCommandLine to see if various parameters are present
                    If .RetrieveValueForParameter("I", strValue) Then mInputFilePath = strValue
                    If .RetrieveValueForParameter("O", strValue) Then mOutputFolderName = strValue

                    If .RetrieveValueForParameter("F", strValue) Then
                        If Integer.TryParse(strValue, intResult) Then
                            mFixedColumnCount = intResult
                        Else
                            Console.WriteLine("Error parsing /F parameter; should be an integer")
                        End If
                    End If

                    If .RetrieveValueForParameter("C", strValue) Then
                        ' The user has defined a column separator character
                        mTabDelimitedFile = False
                        If strValue.ToLower = "space" Then
                            mColumnSepChar = " "c
                        ElseIf strValue.ToLower = "tab" Then
                            mColumnSepChar = ControlChars.Tab
                        Else
                            mColumnSepChar = strValue.Chars(0)
                        End If

                    End If

                    If .RetrieveValueForParameter("B", strValue) Then mSkipBlankItems = True
                    If .RetrieveValueForParameter("N", strValue) Then mSkipNullItems = True

                    If .RetrieveValueForParameter("S", strValue) Then
                        mRecurseFolders = True
                        If IsNumeric(strValue) Then
                            mRecurseFoldersMaxLevels = CInt(strValue)
                        End If
                    End If
                    If .RetrieveValueForParameter("A", strValue) Then mOutputFolderAlternatePath = strValue
                    If .RetrieveValueForParameter("R", strValue) Then mRecreateFolderHierarchyInAlternatePath = True

                    If .RetrieveValueForParameter("L", strValue) Then mLogMessagesToFile = True
                    If .RetrieveValueForParameter("Q", strValue) Then mQuietMode = True

                End With

                Return True
            End If

        Catch ex As Exception
            If mQuietMode Then
                Throw New System.Exception("Error parsing the command line parameters", ex)
            Else
                Console.WriteLine("Error parsing the command line parameters: " & ControlChars.NewLine & ex.Message)
            End If
        End Try

    End Function

    Private Sub ShowProgramHelp()

        Try
            Console.WriteLine("This program reads in a delimited text file that is in crosstab (aka pivot table) format and writes out a new file where the data has been unpivotted.")
            Console.WriteLine()
            Console.WriteLine("Program syntax:" & ControlChars.NewLine & System.IO.Path.GetFileName(System.Reflection.Assembly.GetExecutingAssembly().Location) & _
                              " /I:InputFilePath [/O:OutputFolderName]")
            Console.WriteLine(" [/F:FixedColumnCount] [/C:ColumnSepChar] [/B] [/N]")
            Console.WriteLine(" [/S:[MaxLevel]] [/A:AlternateOutputFolderPath] [/R] [/L] [/Q]")
            Console.WriteLine()

            Console.WriteLine("The input file path can contain the wildcard character *.  If a wildcard is present, then all matching files will be processed")
            Console.WriteLine("The output folder name is optional.  If omitted, the output files will be created in the same folder as the input file.  If included, then a subfolder is created with the name OutputFolderName.")
            Console.WriteLine()

            Console.WriteLine("Use /F to define the number of fixed columns (default is /F:1).  When unpivotting, data in these columns will be written to every row in the output file.")
            Console.WriteLine("The default column separation character is the tab character.  Use /C to define an alternate character.  For example, use /C:, for a comma.  For a space, use /C:space")
            Console.WriteLine("Use /B to skip writing blank column values to the output file")
            Console.WriteLine("Use /N to skip writing Null values to the output file (as indicated by the word 'null').")
            Console.WriteLine()

            Console.WriteLine("Use /S to process all valid files in the input folder and subfolders. Include a number after /S (like /S:2) to limit the level of subfolders to examine.")
            Console.WriteLine("When using /S, you can redirect the output of the results using /A.")
            Console.WriteLine("When using /S, you can use /R to re-create the input folder hierarchy in the alternate output folder (if defined).")
            Console.WriteLine()
            Console.WriteLine("Use /L to log messages to a file.  Use the optional /Q switch will suppress all error messages.")
            Console.WriteLine()

            Console.WriteLine("Program written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA) in 2009")
            Console.WriteLine()

            Console.WriteLine("This is version " & System.Windows.Forms.Application.ProductVersion & " (" & PROGRAM_DATE & ")")
            Console.WriteLine()

            Console.WriteLine("E-mail: matthew.monroe@pnl.gov or matt@alchemistmatt.com")
            Console.WriteLine("Website: http://ncrr.pnl.gov/ or http://www.sysbio.org/resources/staff/")
            Console.WriteLine()

            Console.WriteLine("Licensed under the Apache License, Version 2.0; you may not use this file except in compliance with the License.  " & _
                              "You may obtain a copy of the License at http://www.apache.org/licenses/LICENSE-2.0")
            Console.WriteLine()

            Console.WriteLine("Notice: This computer software was prepared by Battelle Memorial Institute, " & _
                              "hereinafter the Contractor, under Contract No. DE-AC05-76RL0 1830 with the " & _
                              "Department of Energy (DOE).  All rights in the computer software are reserved " & _
                              "by DOE on behalf of the United States Government and the Contractor as " & _
                              "provided in the Contract.  NEITHER THE GOVERNMENT NOR THE CONTRACTOR MAKES ANY " & _
                              "WARRANTY, EXPRESS OR IMPLIED, OR ASSUMES ANY LIABILITY FOR THE USE OF THIS " & _
                              "SOFTWARE.  This notice including this sentence must appear on any copies of " & _
                              "this computer software.")
            Console.WriteLine()

            ' Delay for 750 msec in case the user double clicked this file from within Windows Explorer (or started the program via a shortcut)
            System.Threading.Thread.Sleep(750)

        Catch ex As Exception
            Console.WriteLine("Error displaying the program syntax: " & ex.Message)
        End Try

    End Sub

    Private Sub mUnpivoter_ProgressChanged(ByVal taskDescription As String, ByVal percentComplete As Single) Handles mUnpivoter.ProgressUpdate
        Const PERCENT_REPORT_INTERVAL As Integer = 25
        Const PROGRESS_DOT_INTERVAL_MSEC As Integer = 250

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

    Private Sub mUnpivoter_ProgressReset() Handles mUnpivoter.ProgressReset
        mLastProgressReportTime = DateTime.UtcNow
        mLastProgressReportValue = 0
    End Sub
End Module
