Option Strict On

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

Public Class clsFileUnpivoter
    Inherits clsProcessFilesBaseClass

    Public Sub New()
        MyBase.mFileDate = PROGRAM_DATE
        InitializeLocalVariables()
    End Sub

#Region "Constants and Enums"
    Public Enum eFileUnpivoterErrorCodes
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

    Private mLocalErrorCode As eFileUnpivoterErrorCodes
#End Region

#Region "Properties"

    Public Property ColumnSepChar() As Char
        Get
            Return mColumnSepChar
        End Get
        Set(ByVal value As Char)
            mColumnSepChar = value
            If mColumnSepChar = ControlChars.Tab Then
                mTabDelimitedFile = True
            Else
                mTabDelimitedFile = False
            End If
        End Set
    End Property

    Public Property FixedColumnCount() As Integer
        Get
            Return mFixedColumnCount
        End Get
        Set(ByVal Value As Integer)
            If Value < 0 Then Value = 0
            mFixedColumnCount = Value
        End Set
    End Property

    Public Property SkipBlankItems() As Boolean
        Get
            Return mSkipBlankItems
        End Get
        Set(ByVal value As Boolean)
            mSkipBlankItems = value
        End Set
    End Property

    Public Property SkipNullItems() As Boolean
        Get
            Return mSkipNullItems
        End Get
        Set(ByVal value As Boolean)
            mSkipNullItems = value
        End Set
    End Property

    Public Property TabDelimitedFile() As Boolean
        Get
            Return mTabDelimitedFile
        End Get
        Set(ByVal value As Boolean)
            mTabDelimitedFile = value
            If mTabDelimitedFile Then
                mColumnSepChar = ControlChars.Tab
            End If
        End Set
    End Property

#End Region

    Public Overrides Sub AbortProcessingNow()
        MyBase.AbortProcessingNow()
    End Sub

    Private Function DetermineLineTerminatorSize(ByVal strInputFilePath As String) As Integer
        Dim fsInFile As System.IO.FileStream
        Dim intByte As Integer

        Dim intTerminatorSize As Integer = 2

        Try
            ' Open the input file and look for the first carriage return (byte code 13) or line feed (byte code 10)
            ' Examining, at most, the first 100000 bytes

            fsInFile = New System.IO.FileStream(strInputFilePath, System.IO.FileMode.Open, System.IO.FileAccess.Read, IO.FileShare.ReadWrite)

            Do While fsInFile.Position < fsInFile.Length AndAlso fsInFile.Position < 100000

                intByte = fsInFile.ReadByte()

                If intByte = 10 Or intByte = 13 Then
                    ' Found linefeed or carriage return
                    If fsInFile.Position < fsInFile.Length Then
                        intByte = fsInFile.ReadByte()
                        If intByte = 10 Or intByte = 13 Then
                            ' CrLf or LfCr
                            intTerminatorSize = 2
                        Else
                            ' Lf only or Cr only
                            intTerminatorSize = 1
                        End If
                    Else
                        intTerminatorSize = 1
                    End If
                    Exit Do
                End If

            Loop

        Catch ex As Exception
            HandleException("Error in DetermineLineTerminatorSize", ex)
        Finally
            If Not fsInFile Is Nothing Then
                fsInFile.Close()
            End If
        End Try

        Return intTerminatorSize

    End Function

    Public Overrides Function GetErrorMessage() As String
        ' Returns "" if no error

        Dim strErrorMessage As String

        If MyBase.ErrorCode = clsProcessFilesBaseClass.eProcessFilesErrorCodes.LocalizedError Or _
           MyBase.ErrorCode = clsProcessFilesBaseClass.eProcessFilesErrorCodes.NoError Then
            Select Case mLocalErrorCode
                Case eFileUnpivoterErrorCodes.NoError
                    strErrorMessage = ""
                Case eFileUnpivoterErrorCodes.UnspecifiedError
                    strErrorMessage = "Unspecified localized error"
                Case Else
                    ' This shouldn't happen
                    strErrorMessage = "Unknown error state"
            End Select
        Else
            strErrorMessage = MyBase.GetBaseClassErrorMessage()
        End If

        Return strErrorMessage
    End Function

    Private Sub InitializeLocalVariables()
        MyBase.ShowMessages = False

        mFixedColumnCount = 1
        mSkipBlankItems = True
        mSkipNullItems = False

        mTabDelimitedFile = True
        mColumnSepChar = ControlChars.Tab


        mLocalErrorCode = eFileUnpivoterErrorCodes.NoError

    End Sub

    Private Function LoadParameterFileSettings(ByVal strParameterFilePath As String) As Boolean

        Const OPTIONS_SECTION As String = "FileUnpivoter"

        Dim objSettingsFile As New XmlSettingsFileAccessor

        Try

            If strParameterFilePath Is Nothing OrElse strParameterFilePath.Length = 0 Then
                ' No parameter file specified; nothing to load
                Return True
            End If

            If Not System.IO.File.Exists(strParameterFilePath) Then
                ' See if strParameterFilePath points to a file in the same directory as the application
                strParameterFilePath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), System.IO.Path.GetFileName(strParameterFilePath))
                If Not System.IO.File.Exists(strParameterFilePath) Then
                    MyBase.SetBaseClassErrorCode(clsProcessFilesBaseClass.eProcessFilesErrorCodes.ParameterFileNotFound)
                    Return False
                End If
            End If

            If objSettingsFile.LoadSettings(strParameterFilePath) Then
                If Not objSettingsFile.SectionPresent(OPTIONS_SECTION) Then
                    ShowErrorMessage("The node '<section name=""" & OPTIONS_SECTION & """> was not found in the parameter file: " & strParameterFilePath)
                    MyBase.SetBaseClassErrorCode(clsProcessFilesBaseClass.eProcessFilesErrorCodes.InvalidParameterFile)
                    Return False
                Else
                    ' mMySetting = objSettingsFile.GetParam(OPTIONS_SECTION, "MySetting", "Default value")
                End If
            End If

        Catch ex As Exception
            HandleException("Error in LoadParameterFileSettings", ex)
            Return False
        End Try

        Return True

    End Function

    ' Main processing function
    Public Overloads Overrides Function ProcessFile(ByVal strInputFilePath As String, ByVal strOutputFolderPath As String, ByVal strParameterFilePath As String, ByVal blnResetErrorCode As Boolean) As Boolean
        ' Returns True if success, False if failure

        Dim strStatusMessage As String
        Dim blnSuccess As Boolean

        If blnResetErrorCode Then
            SetLocalErrorCode(eFileUnpivoterErrorCodes.NoError)
        End If

        If Not LoadParameterFileSettings(strParameterFilePath) Then
            strStatusMessage = "Parameter file load error: " & strParameterFilePath
            ShowErrorMessage(strStatusMessage)

            If MyBase.ErrorCode = clsProcessFilesBaseClass.eProcessFilesErrorCodes.NoError Then
                MyBase.SetBaseClassErrorCode(clsProcessFilesBaseClass.eProcessFilesErrorCodes.InvalidParameterFile)
            End If
            Return False
        End If

        Try
            If strInputFilePath Is Nothing OrElse strInputFilePath.Length = 0 Then
                ShowErrorMessage("Input file name is empty")
                MyBase.SetBaseClassErrorCode(clsProcessFilesBaseClass.eProcessFilesErrorCodes.InvalidInputFilePath)
            Else
                ' Note that CleanupFilePaths() will update mOutputFolderPath, which is used by LogMessage()
                If Not CleanupFilePaths(strInputFilePath, strOutputFolderPath) Then
                    MyBase.SetBaseClassErrorCode(clsProcessFilesBaseClass.eProcessFilesErrorCodes.FilePathError)
                Else
                    MyBase.mProgressStepDescription = "Parsing " & System.IO.Path.GetFileName(strInputFilePath)
                    LogMessage(MyBase.mProgressStepDescription)
                    MyBase.ResetProgress()

                    ' Call UnpivotFile to perform the work
                    blnSuccess = UnpivotFile(strInputFilePath, strOutputFolderPath)

                    If blnSuccess Then
                        LogMessage("Processing Complete")
                        OperationComplete()
                    End If
                End If
            End If

        Catch ex As Exception
            HandleException("Error in ProcessFile", ex)
            blnSuccess = False
        End Try

        Return blnSuccess

    End Function

    Private Sub SetLocalErrorCode(ByVal eNewErrorCode As eFileUnpivoterErrorCodes)
        SetLocalErrorCode(eNewErrorCode, False)
    End Sub

    Private Sub SetLocalErrorCode(ByVal eNewErrorCode As eFileUnpivoterErrorCodes, ByVal blnLeaveExistingErrorCodeUnchanged As Boolean)

        If blnLeaveExistingErrorCodeUnchanged AndAlso mLocalErrorCode <> eFileUnpivoterErrorCodes.NoError Then
            ' An error code is already defined; do not change it
        Else
            mLocalErrorCode = eNewErrorCode

            If eNewErrorCode = eFileUnpivoterErrorCodes.NoError Then
                If MyBase.ErrorCode = clsProcessFilesBaseClass.eProcessFilesErrorCodes.LocalizedError Then
                    MyBase.SetBaseClassErrorCode(clsProcessFilesBaseClass.eProcessFilesErrorCodes.NoError)
                End If
            Else
                MyBase.SetBaseClassErrorCode(clsProcessFilesBaseClass.eProcessFilesErrorCodes.LocalizedError)
            End If
        End If

    End Sub

    Protected Function UnpivotFile(ByVal strInputFilePath As String, ByVal strOutputFolderPath As String) As Boolean

        Dim srInFile As System.IO.StreamReader
        Dim swOutFile As System.IO.StreamWriter

        Dim strOutputFilePath As String
        Dim strSepCharDescription As String

        Dim intLineTerminatorBytes As Integer
        Dim intLinesRead As Integer
        Dim lngBytesRead As Long
        Dim sngPercentComplete As Single

        Dim strLineIn As String
        Dim strLineOut As String

        Dim strSplitLine() As String

        Dim intHeaderColumnCount As Integer
        Dim strHeaderColumnNames() As String

        Dim intIndex As Integer

        Dim blnHeaderProcessed As Boolean
        Dim blnSkipValue As Boolean
        Dim blnSuccess As Boolean

        Try
            strOutputFilePath = System.IO.Path.GetFileNameWithoutExtension(strInputFilePath) & "_Unpivot" & System.IO.Path.GetExtension(strInputFilePath)

            If strOutputFolderPath Is Nothing OrElse strOutputFolderPath.Length = 0 Then
                strOutputFilePath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(strInputFilePath), strOutputFilePath)
            Else
                strOutputFilePath = System.IO.Path.Combine(strOutputFolderPath, strOutputFilePath)
            End If

            Try
                ' Open the input file
                srInFile = New System.IO.StreamReader(New System.IO.FileStream(strInputFilePath, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read))

            Catch ex As Exception
                HandleException("Error opening input file: " & strInputFilePath, ex)
                Return False
            End Try

            Try
                ' Create the output file; it will be overwritten if it exists
                swOutFile = New System.IO.StreamWriter(New System.IO.FileStream(strOutputFilePath, IO.FileMode.Create, IO.FileAccess.Write, IO.FileShare.Read))

            Catch ex As Exception
                HandleException("Error creating the output file: " & strOutputFilePath, ex)
                Return False
            End Try

            ' Make sure mColumnSepChar is a tab if mTabDelimitedFile is true
            If mTabDelimitedFile Then mColumnSepChar = ControlChars.Tab

            ' Define the column sep char description
            If mColumnSepChar = ControlChars.Tab Then
                strSepCharDescription = "<tab>"
            Else
                strSepCharDescription = "'" & mColumnSepChar & "'"
            End If

            ' Initialize the tracking variables
            ReDim strHeaderColumnNames(0)
            intLinesRead = 0
            blnHeaderProcessed = False

            intLineTerminatorBytes = DetermineLineTerminatorSize(strInputFilePath)

            ' Parse the input file and create the output file
            Do While srInFile.Peek >= 0
                strLineIn = srInFile.ReadLine()
                intLinesRead += 1

                If strLineIn Is Nothing OrElse strLineIn = String.Empty Then
                    ' Blank line; skip it
                Else
                    ' Bump up lngBytesRead
                    lngBytesRead += strLineIn.Length + intLineTerminatorBytes

                    ' Split the line on the separation char
                    strSplitLine = strLineIn.Split(mColumnSepChar)

                    If strSplitLine.Length = 0 OrElse Not strLineIn.Contains(mColumnSepChar) Then
                        ShowMessage("Warning: line " & intLinesRead.ToString & " did not contain the separation character " & strSepCharDescription & "; this line will be skipped")
                    End If

                    ' Add the fixed column names to the output line (this is required on both the header line and subsequent lines)
                    strLineOut = String.Empty
                    For intIndex = 0 To mFixedColumnCount - 1
                        If intIndex >= strSplitLine.Length Then
                            Exit For
                        End If
                        strLineOut &= strSplitLine(intIndex) & mColumnSepChar
                    Next

                    If Not blnHeaderProcessed Then
                        ' Write out the header line
                        swOutFile.WriteLine(strLineOut & "Header" & mColumnSepChar & "Value")

                        ' Cache the Header names in strHeaderColumnNames
                        intHeaderColumnCount = strSplitLine.Length - mFixedColumnCount - 1
                        If intHeaderColumnCount < -1 Then intHeaderColumnCount = -1
                        ReDim strHeaderColumnNames(intHeaderColumnCount)

                        For intIndex = mFixedColumnCount To strSplitLine.Length - 1
                            strHeaderColumnNames(intIndex - mFixedColumnCount) = String.Copy(strSplitLine(intIndex))
                        Next

                        blnHeaderProcessed = True
                    Else
                        ' Unpivot this line, writing out to the output file
                        For intIndex = mFixedColumnCount To strSplitLine.Length - 1
                            If intIndex - mFixedColumnCount >= strHeaderColumnNames.Length Then
                                ' This data line has too many columns of data; ignore the remaining columns
                                ShowMessage("Warning: line " & intLinesRead.ToString & " has extra data columns (the header line had " & strHeaderColumnNames.Length.ToString & " data columns; remaining data columns will be skipped")
                                Exit For
                            End If

                            blnSkipValue = False
                            If mSkipBlankItems Then
                                If strSplitLine(intIndex) Is Nothing OrElse strSplitLine(intIndex).Trim.Length = 0 Then
                                    blnSkipValue = True
                                End If
                            End If

                            If Not blnSkipValue AndAlso mSkipNullItems Then
                                If strSplitLine(intIndex) Is Nothing OrElse strSplitLine(intIndex).ToLower = "null" Then
                                    blnSkipValue = True
                                End If
                            End If

                            If Not blnSkipValue Then
                                swOutFile.WriteLine(strLineOut & strHeaderColumnNames(intIndex - mFixedColumnCount) & mColumnSepChar & strSplitLine(intIndex))
                            End If
                        Next

                    End If

                End If

                If intLinesRead Mod 100 = 0 Then
                    sngPercentComplete = CSng(lngBytesRead / CSng(srInFile.BaseStream.Length) * 100)
                    UpdateProgress(sngPercentComplete)
                End If
            Loop

            If Not srInFile Is Nothing Then srInFile.Close()
            If Not swOutFile Is Nothing Then swOutFile.Close()

            blnSuccess = True

        Catch ex As Exception
            HandleException("Error in UnpivotFile", ex)
            blnSuccess = False
        End Try

        Return blnSuccess
    End Function
End Class
