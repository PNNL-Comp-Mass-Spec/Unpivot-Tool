Option Strict On

Imports System.IO
Imports System.Reflection
Imports PRISM
Imports PRISM.FileProcessor

' This class reads in a tab delimited file that is in a crosstab / pivottable format
' and writes out a new file where the data has been unpivotted
'
' -------------------------------------------------------------------------------
' Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
' Program started July 20,2009
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

' Last updated July 20, 2009

Public Class FileUnpivoter
    Inherits ProcessFilesBase

    Public Sub New()
        MyBase.mFileDate = PROGRAM_DATE
        InitializeLocalVariables()
    End Sub

#Region "Constants and Enums"
    Public Enum FileUnpivoterErrorCodes
        NoError = 0
        UnspecifiedError = -1
    End Enum
#End Region

#Region "Structures"

#End Region

#Region "Classwide variables"
    Private mFixedColumnCount As Integer
    Private mSkipBlankItems As Boolean
    Private mSkipNullItems As Boolean

    Private mTabDelimitedFile As Boolean        ' When true (the default) then assumes the separation char is a tab; if false, then mColumnSepChar is used
    Private mColumnSepChar As Char

    Private mLocalErrorCode As FileUnpivoterErrorCodes
#End Region

#Region "Properties"

    Public Property ColumnSepChar As Char
        Get
            Return mColumnSepChar
        End Get
        Set
            mColumnSepChar = Value
            If mColumnSepChar = ControlChars.Tab Then
                mTabDelimitedFile = True
            Else
                mTabDelimitedFile = False
            End If
        End Set
    End Property

    Public Property FixedColumnCount As Integer
        Get
            Return mFixedColumnCount
        End Get
        Set
            If Value < 0 Then Value = 0
            mFixedColumnCount = Value
        End Set
    End Property

    Public Property SkipBlankItems As Boolean
        Get
            Return mSkipBlankItems
        End Get
        Set
            mSkipBlankItems = Value
        End Set
    End Property

    Public Property SkipNullItems As Boolean
        Get
            Return mSkipNullItems
        End Get
        Set
            mSkipNullItems = Value
        End Set
    End Property

    ' ReSharper disable once UnusedMember.Global
    Public Property TabDelimitedFile As Boolean
        Get
            Return mTabDelimitedFile
        End Get
        Set
            mTabDelimitedFile = Value
            If mTabDelimitedFile Then
                mColumnSepChar = ControlChars.Tab
            End If
        End Set
    End Property

#End Region

    Public Overrides Sub AbortProcessingNow()
        MyBase.AbortProcessingNow()
    End Sub

    Private Function DetermineLineTerminatorSize(inputFilePath As String) As Integer

        Dim terminatorSize = 2

        Try
            ' Open the input file and look for the first carriage return (byte code 13) or line feed (byte code 10)
            ' Examining, at most, the first 100000 bytes

            Using reader = New FileStream(inputFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)

                Do While reader.Position < reader.Length AndAlso reader.Position < 100000

                    Dim oneByte = reader.ReadByte()

                    If oneByte = 10 Or oneByte = 13 Then
                        ' Found linefeed or carriage return
                        If reader.Position < reader.Length Then
                            oneByte = reader.ReadByte()
                            If oneByte = 10 Or oneByte = 13 Then
                                ' CrLf or LfCr
                                terminatorSize = 2
                            Else
                                ' Lf only or Cr only
                                terminatorSize = 1
                            End If
                        Else
                            terminatorSize = 1
                        End If
                        Exit Do
                    End If

                Loop

            End Using

        Catch ex As Exception
            HandleException("Error in DetermineLineTerminatorSize", ex)
        End Try

        Return terminatorSize

    End Function

    Public Overrides Function GetErrorMessage() As String
        ' Returns "" if no error

        Dim message As String

        If MyBase.ErrorCode = ProcessFilesErrorCodes.LocalizedError Or
           MyBase.ErrorCode = ProcessFilesErrorCodes.NoError Then
            Select Case mLocalErrorCode
                Case FileUnpivoterErrorCodes.NoError
                    message = ""
                Case FileUnpivoterErrorCodes.UnspecifiedError
                    message = "Unspecified localized error"
                Case Else
                    ' This shouldn't happen
                    message = "Unknown error state"
            End Select
        Else
            message = MyBase.GetBaseClassErrorMessage()
        End If

        Return message
    End Function

    Private Sub InitializeLocalVariables()
        mFixedColumnCount = 1
        mSkipBlankItems = True
        mSkipNullItems = False

        mTabDelimitedFile = True
        mColumnSepChar = ControlChars.Tab


        mLocalErrorCode = FileUnpivoterErrorCodes.NoError

    End Sub

    Private Function LoadParameterFileSettings(parameterFilePath As String) As Boolean

        Const OPTIONS_SECTION = "FileUnpivoter"

        Dim settingsFile As New XmlSettingsFileAccessor()

        Try

            If parameterFilePath Is Nothing OrElse parameterFilePath.Length = 0 Then
                ' No parameter file specified; nothing to load
                Return True
            End If

            If Not File.Exists(parameterFilePath) Then
                ' See if parameterFilePath points to a file in the same directory as the application
                parameterFilePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), Path.GetFileName(parameterFilePath))
                If Not File.Exists(parameterFilePath) Then
                    MyBase.SetBaseClassErrorCode(ProcessFilesErrorCodes.ParameterFileNotFound)
                    Return False
                End If
            End If

            If settingsFile.LoadSettings(parameterFilePath) Then
                If Not settingsFile.SectionPresent(OPTIONS_SECTION) Then
                    ShowErrorMessage("The node '<section name=""" & OPTIONS_SECTION & """> was not found in the parameter file: " & parameterFilePath)
                    MyBase.SetBaseClassErrorCode(ProcessFilesErrorCodes.InvalidParameterFile)
                    Return False
                Else
                    ' mMySetting = settingsFile.GetParam(OPTIONS_SECTION, "MySetting", "Default value")
                End If
            End If

        Catch ex As Exception
            HandleException("Error in LoadParameterFileSettings", ex)
            Return False
        End Try

        Return True

    End Function

    ' Main processing function
    ' Returns True if success, False if failure
    Public Overloads Overrides Function ProcessFile(inputFilePath As String, outputDirectoryPath As String, parameterFilePath As String, resetErrorCode As Boolean) As Boolean

        Dim statusMessage As String
        Dim success As Boolean

        If resetErrorCode Then
            SetLocalErrorCode(FileUnpivoterErrorCodes.NoError)
        End If

        If Not LoadParameterFileSettings(parameterFilePath) Then
            statusMessage = "Parameter file load error: " & parameterFilePath
            ShowErrorMessage(statusMessage)

            If MyBase.ErrorCode = ProcessFilesErrorCodes.NoError Then
                MyBase.SetBaseClassErrorCode(ProcessFilesErrorCodes.InvalidParameterFile)
            End If
            Return False
        End If

        Try
            If inputFilePath Is Nothing OrElse inputFilePath.Length = 0 Then
                ShowErrorMessage("Input file name is empty")
                MyBase.SetBaseClassErrorCode(ProcessFilesErrorCodes.InvalidInputFilePath)
            Else
                ' Note that CleanupFilePaths() will update mOutputDirectoryPath, which is used by LogMessage()
                If Not CleanupFilePaths(inputFilePath, outputDirectoryPath) Then
                    MyBase.SetBaseClassErrorCode(ProcessFilesErrorCodes.FilePathError)
                Else
                    MyBase.UpdateProgress("Parsing " & Path.GetFileName(inputFilePath))
                    LogMessage(MyBase.ProgressStepDescription)
                    MyBase.ResetProgress()

                    ' Call UnpivotFile to perform the work
                    success = UnpivotFile(inputFilePath, outputDirectoryPath)

                    If success Then
                        LogMessage("Processing Complete")
                        OperationComplete()
                    End If
                End If
            End If

        Catch ex As Exception
            HandleException("Error in ProcessFile", ex)
            success = False
        End Try

        Return success

    End Function

    Private Sub SetLocalErrorCode(eNewErrorCode As FileUnpivoterErrorCodes)
        SetLocalErrorCode(eNewErrorCode, False)
    End Sub

    Private Sub SetLocalErrorCode(eNewErrorCode As FileUnpivoterErrorCodes, leaveExistingErrorCodeUnchanged As Boolean)

        If leaveExistingErrorCodeUnchanged AndAlso mLocalErrorCode <> FileUnpivoterErrorCodes.NoError Then
            ' An error code is already defined; do not change it
        Else
            mLocalErrorCode = eNewErrorCode

            If eNewErrorCode = FileUnpivoterErrorCodes.NoError Then
                If MyBase.ErrorCode = ProcessFilesErrorCodes.LocalizedError Then
                    MyBase.SetBaseClassErrorCode(ProcessFilesErrorCodes.NoError)
                End If
            Else
                MyBase.SetBaseClassErrorCode(ProcessFilesErrorCodes.LocalizedError)
            End If
        End If

    End Sub

    Protected Function UnpivotFile(inputFilePath As String, outputDirectoryPath As String) As Boolean

        Dim reader As StreamReader
        Dim writer As StreamWriter

        Dim outputFilePath As String
        Dim sepCharDescription As String

        Dim lineTerminatorBytes As Integer
        Dim linesRead As Integer
        Dim bytesRead As Long
        Dim percentComplete As Single

        Dim lineIn As String
        Dim lineOut As String

        Dim splitLine() As String

        Dim headerColumnCount As Integer
        Dim headerColumnNames() As String

        Dim index As Integer

        Dim headerProcessed As Boolean
        Dim skipValue As Boolean
        Dim success As Boolean

        Try
            outputFilePath = Path.GetFileNameWithoutExtension(inputFilePath) & "_Unpivot" & Path.GetExtension(inputFilePath)

            If outputDirectoryPath Is Nothing OrElse outputDirectoryPath.Length = 0 Then
                outputFilePath = Path.Combine(Path.GetDirectoryName(inputFilePath), outputFilePath)
            Else
                outputFilePath = Path.Combine(outputDirectoryPath, outputFilePath)
            End If

            Try
                ' Open the input file
                reader = New StreamReader(New FileStream(inputFilePath, FileMode.Open, FileAccess.Read, FileShare.Read))

            Catch ex As Exception
                HandleException("Error opening input file: " & inputFilePath, ex)
                Return False
            End Try

            Try
                ' Create the output file; it will be overwritten if it exists
                writer = New StreamWriter(New FileStream(outputFilePath, FileMode.Create, FileAccess.Write, FileShare.Read))

            Catch ex As Exception
                HandleException("Error creating the output file: " & outputFilePath, ex)
                Return False
            End Try

            ' Make sure mColumnSepChar is a tab if mTabDelimitedFile is true
            If mTabDelimitedFile Then mColumnSepChar = ControlChars.Tab

            ' Define the column sep char description
            If mColumnSepChar = ControlChars.Tab Then
                sepCharDescription = "<tab>"
            Else
                sepCharDescription = "'" & mColumnSepChar & "'"
            End If

            ' Initialize the tracking variables
            ReDim headerColumnNames(0)
            linesRead = 0
            headerProcessed = False

            lineTerminatorBytes = DetermineLineTerminatorSize(inputFilePath)

            ' Parse the input file and create the output file
            Do While Not reader.EndOfStream
                lineIn = reader.ReadLine()
                linesRead += 1

                If lineIn Is Nothing OrElse lineIn = String.Empty Then
                    ' Blank line; skip it
                Else
                    ' Bump up bytesRead
                    bytesRead += lineIn.Length + lineTerminatorBytes

                    ' Split the line on the separation char
                    splitLine = lineIn.Split(mColumnSepChar)

                    If splitLine.Length = 0 OrElse Not lineIn.Contains(mColumnSepChar) Then
                        ShowMessage("Warning: line " & linesRead.ToString & " did not contain the separation character " & sepCharDescription & "; this line will be skipped")
                    End If

                    ' Add the fixed column names to the output line (this is required on both the header line and subsequent lines)
                    lineOut = String.Empty
                    For index = 0 To mFixedColumnCount - 1
                        If index >= splitLine.Length Then
                            Exit For
                        End If
                        lineOut &= splitLine(index) & mColumnSepChar
                    Next

                    If Not headerProcessed Then
                        ' Write out the header line
                        writer.WriteLine(lineOut & "Header" & mColumnSepChar & "Value")

                        ' Cache the Header names in headerColumnNames
                        headerColumnCount = splitLine.Length - mFixedColumnCount - 1
                        If headerColumnCount < -1 Then headerColumnCount = -1
                        ReDim headerColumnNames(headerColumnCount)

                        For index = mFixedColumnCount To splitLine.Length - 1
                            headerColumnNames(index - mFixedColumnCount) = String.Copy(splitLine(index))
                        Next

                        headerProcessed = True
                    Else
                        ' Unpivot this line, writing out to the output file
                        For index = mFixedColumnCount To splitLine.Length - 1
                            If index - mFixedColumnCount >= headerColumnNames.Length Then
                                ' This data line has too many columns of data; ignore the remaining columns
                                ShowMessage("Warning: line " & linesRead.ToString & " has extra data columns (the header line had " & headerColumnNames.Length.ToString & " data columns; remaining data columns will be skipped")
                                Exit For
                            End If

                            skipValue = False
                            If mSkipBlankItems Then
                                If splitLine(index) Is Nothing OrElse splitLine(index).Trim.Length = 0 Then
                                    skipValue = True
                                End If
                            End If

                            If Not skipValue AndAlso mSkipNullItems Then
                                If splitLine(index) Is Nothing OrElse splitLine(index).ToLower = "null" Then
                                    skipValue = True
                                End If
                            End If

                            If Not skipValue Then
                                writer.WriteLine(lineOut & headerColumnNames(index - mFixedColumnCount) & mColumnSepChar & splitLine(index))
                            End If
                        Next

                    End If

                End If

                If linesRead Mod 100 = 0 Then
                    percentComplete = CSng(bytesRead / CSng(reader.BaseStream.Length) * 100)
                    UpdateProgress(percentComplete)
                End If
            Loop

            If reader IsNot Nothing Then reader.Close()
            If writer IsNot Nothing Then writer.Close()

            success = True

        Catch ex As Exception
            HandleException("Error in UnpivotFile", ex)
            success = False
        End Try

        Return success
    End Function
End Class
