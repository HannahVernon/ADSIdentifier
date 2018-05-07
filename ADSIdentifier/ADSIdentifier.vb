Option Compare Text ' NTFS is case-insensitive by default
Option Explicit On
Option Strict Off
Option Infer Off

Imports System
Imports System.Runtime.InteropServices

Module NativeMethods

    Public Enum StreamInfoLevels As Int32
        FindStreamInfoStandard = 0
        FindStreamInfoMaxInfoLevel = 1
    End Enum

    Public Const INVALID_HANDLE_VALUE As Int32 = -1
    Public Const ERROR_HANDLE_EOF As Int32 = 38
    Public Const FILE_ATTRIBUTE_DIRECTORY As UInt32 = 16
    Public Const FILE_ATTRIBUTE_NORMAL As UInt32 = 128

    <StructLayout(LayoutKind.Explicit)>
    Public Structure LARGE_INTEGER
        <FieldOffset(0)> Dim Low As Int32
        <FieldOffset(4)> Dim High As Int32
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Public Structure WIN32_FIND_STREAM_DATA
        Dim StreamSize As LARGE_INTEGER ' <MarshalAs(UnmanagedType.I8)> 
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=296)> Dim cStreamName As String ' 260 (max_path) + 36
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Structure WIN32_FIND_DATA
        Public dwFileAttributes As UInteger
        Public ftCreationTime As System.Runtime.InteropServices.ComTypes.FILETIME
        Public ftLastAccessTime As System.Runtime.InteropServices.ComTypes.FILETIME
        Public ftLastWriteTime As System.Runtime.InteropServices.ComTypes.FILETIME
        Public nFileSizeHigh As UInteger
        Public nFileSizeLow As UInteger
        Public dwReserved0 As UInteger
        Public dwReserved1 As UInteger
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)> Public cFileName As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=14)> Public cAlternateFileName As String
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Structure WIN32_SECURITY_ATTRIBUTES
        Public nLength As Int32
        Public lpSecurityDescriptor As IntPtr
        Public bInheritHandle As Boolean
    End Structure

    <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, EntryPoint:="FindFirstStreamW")>
    Function FindFirstStream(<MarshalAs(UnmanagedType.LPWStr)> ByVal lpFileName As String, ByVal InfoLevel As StreamInfoLevels, ByRef lpFindStreamData As WIN32_FIND_STREAM_DATA, ByVal dwFlags As Int32) As IntPtr
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, EntryPoint:="FindNextStreamW")>
    Function FindNextStream(ByVal hFindStream As IntPtr, ByRef lpFindStreamData As WIN32_FIND_STREAM_DATA) As Boolean
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, EntryPoint:="DeleteFileW")>
    Function DeleteFile(ByVal lpFileName As String) As Boolean
    End Function

    <DllImport("kernel32.dll")>
    Public Function FindClose(ByVal hFindFile As IntPtr) As Boolean
    End Function

    <DllImport("kernel32.dll")>
    Public Function GetLastError() As Int32
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, EntryPoint:="FindFirstFileW")>
    Public Function FindFirstFile(ByVal lpFileName As String, ByRef lpFindFileData As WIN32_FIND_DATA) As IntPtr
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, EntryPoint:="FindNextFileW")>
    Public Function FindNextFile(ByVal hFindFile As IntPtr, ByRef lpFindFileData As WIN32_FIND_DATA) As Boolean
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, EntryPoint:="CreateFileW")>
    Public Function CreateFile(ByVal lpFileName As String, ByVal dwDesiredAccess As Int32, ByVal dwShareMode As Int32, ByRef lpSecurityAttributes As WIN32_SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Int32, ByVal dwFlagsAndAttributes As Int32, ByVal hTemplateFile As IntPtr) As IntPtr
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, EntryPoint:="GetFileTime")>
    Public Function GetFileTime(ByVal hFile As IntPtr, ByRef lpCreationTime As FILETIME, ByRef lpLastAccessTime As FILETIME, ByRef lpLastWriteTime As FILETIME) As Boolean
    End Function

    <StructLayout(LayoutKind.Sequential)>
    Public Structure FILETIME
        Public dwLowDateTime As UInteger
        Public dwHighDateTime As UInteger

        Public ReadOnly Property Value() As ULong
            Get
                Dim h As ULong = dwHighDateTime
                Dim l As ULong = dwLowDateTime
                Return (h << 32) + l
            End Get
        End Property
        Public ReadOnly Property DateTime As DateTime
            Get
                Return DateTime.FromFileTime(Me.Value)
            End Get
        End Property
    End Structure

    <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, EntryPoint:="FileTimeToLocalFileTime")>
    Public Function FileTimeToLocalFileTime(ByRef lpFileTime As FILETIME, ByRef lpLocalFileTime As FILETIME) As Boolean
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, EntryPoint:="FileTimeToSystemTime")>
    Public Function FileTimeToSystemTime(ByRef lpFileTime As FILETIME, ByRef lpLocalFileTime As FILETIME) As Boolean
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, EntryPoint:="CloseHandle")>
    Public Function CloseHandle(ByVal handle As IntPtr) As Boolean
    End Function
End Module

Module ADSIdentifier

    Sub Main()
        Dim sFolder As String = ""
        Dim bPause As Boolean = False
        Dim bIgnoreZoneIdentifier As Boolean = False
        Dim sPattern As String = ""
        Dim bDeleteStreams As Boolean = False
        Dim bDebug As Boolean = False
        Dim bBare As Boolean = False
        For Each sItem As String In My.Application.CommandLineArgs
            If sItem.ToUpper.StartsWith("/FOLDER:") Then
                sFolder = Right(sItem, Len(sItem) - 8)
            End If
            If sItem.ToUpper = "/P" Or sItem.ToUpper = "/PAUSE" Then
                bPause = True
            End If
            If sItem.ToUpper = "/IZI" Or sItem.ToUpper = "/IGNOREZONEIDENTIFIER" Then
                bIgnoreZoneIdentifier = True
            End If
            If sItem.ToUpper.StartsWith("/PATTERN:") Then
                sPattern = Right(sItem, Len(sItem) - 9)
            End If
            If sItem.ToUpper = "/R" Or sItem.ToUpper = "/REMOVE" Then
                bDeleteStreams = True
            End If
            If sItem.ToUpper = "/D" Or sItem.ToUpper = "/DEBUG" Then
                bDebug = True
            End If
            If sItem.ToUpper = "/B" Or sItem.ToUpper = "/BARE" Then
                bBare = True
            End If
        Next
        If sPattern = "" Then sPattern = "*"
        If Left(sPattern, 1) <> "*" Then sPattern = "*" & sPattern
        If Right(sPattern, 1) <> "*" Then sPattern = sPattern & "*"
        Console.WriteLine("ADSIdentifier.exe v" & My.Application.Info.Version.ToString)
        If sFolder <> "" Then
            Console.WriteLine("Searching for Alternate Data Streams in " & sFolder)
            Console.WriteLine("Pattern Match: " & sPattern)
            Console.WriteLine("Alternate Data Streams will " & IIf(bDeleteStreams, "", "not") & " be deleted.")
            If bIgnoreZoneIdentifier Then
                Console.WriteLine("Zone.Identifier streams will be ignored.")
            End If
            If bBare = False Then
                Console.WriteLine("")
                Console.WriteLine("Create Date            Size in Bytes   Stream Name")
                Console.WriteLine("--------------------------------------------------")
            End If
            GetStreams(sFolder, bIgnoreZoneIdentifier, bPause, sPattern, bDeleteStreams, bDebug, bBare)
        Else
            Console.WriteLine("Useage is:   ADSIdentifier.exe /Folder:<starting_folder_name>")
            Console.WriteLine("                  [/P] or [/Pause] - pause before exiting")
            Console.WriteLine("                  [/IZI] or [/IgnoreZoneIdentifier] - ignore :Zone.Identifier streams")
            Console.WriteLine("                  [/Pattern:<xyz>] - only find Alternate Data Streams matching <xyz>")
            Console.WriteLine("                  [/r] or [/Remove] - remove Alternate Data Streams that have been found matching the other parameters")
            Console.WriteLine("                  [/ns] or [/nosize] - don't display alternate data stream sizes in the output")
            Console.WriteLine("                  [/d] or [/debug] - show debugging details")
        End If
        If bPause Then
            Do Until Console.KeyAvailable = False
                Console.ReadKey(True)
            Loop
            Console.WriteLine("")
            Console.WriteLine("Press any key to exit...")
            Console.ReadKey(True)
        End If
    End Sub

    Sub GetStreams(ByVal StartingFolder As String, ByVal IgnoreZoneIdentifier As Boolean, ByVal Pause As Boolean, ByVal Pattern As String, ByVal DeleteStreams As Boolean, ByVal Debug As Boolean, ByVal Bare As Boolean)
        Dim fsd As New WIN32_FIND_STREAM_DATA
        Dim sItem As String = StartingFolder
        Dim iErr As Int32 = 0
        If Debug Then Console.WriteLine(StartingFolder)
        Dim iResult As IntPtr = FindFirstStream(StartingFolder, StreamInfoLevels.FindStreamInfoStandard, fsd, 0)
        iErr = GetLastError()
        Dim iResDel As Int32
        Dim sSize As String = ""
        If iResult <> INVALID_HANDLE_VALUE Then
            If (Not fsd.cStreamName Like ":Zone.Identifier*" Or IgnoreZoneIdentifier = False) And fsd.cStreamName Like Pattern Then
                If Bare = False Then
                    sSize = Strings.RSet((fsd.StreamSize.Low + (fsd.StreamSize.High << 32)).ToString, 14) & "   "
                    sSize = Strings.LSet(GetFileCreateTime(StartingFolder & fsd.cStreamName.Replace(":$DATA", "")), 22) & " " & sSize
                Else
                    sSize = ""
                End If
                Console.WriteLine(sSize & StartingFolder & fsd.cStreamName.Replace(":$DATA", ""))
                If DeleteStreams Then
                    Dim bDelete As Boolean = True
                    If Pause Then
                        Console.WriteLine("Delete " & StartingFolder & fsd.cStreamName & " ?  [Y,N]")
                        Dim cKey As ConsoleKeyInfo = Console.ReadKey
                        If cKey.KeyChar.ToString.ToUpper <> "Y" Then
                            bDelete = False
                        End If
                        Console.WriteLine()
                    End If
                    If bDelete Then
                        iResDel = DeleteFile(StartingFolder & fsd.cStreamName.Replace(":$DATA", ""))
                        If iResDel = 0 Then
                            iErr = GetLastError()
                            Console.Error.WriteLine("DeleteFile failed with error: " & iErr.ToString & " while attempting to delete: '" & StartingFolder & fsd.cStreamName.Replace(":$DATA", "") & "'")
                        End If
                    End If
                End If
            End If
            Dim iRes As Int32 = -1
            While iRes = -1
                iRes = FindNextStream(iResult, fsd)
                If iRes = 0 Then ' failed
                    iErr = GetLastError()
                    If iErr <> ERROR_HANDLE_EOF Then
                        Console.Error.WriteLine("Error: " & iErr)
                    End If
                    Exit While
                Else ' we've found another stream, report it
                    If (Not fsd.cStreamName Like ":Zone.Identifier*" Or IgnoreZoneIdentifier = False) And fsd.cStreamName Like Pattern Then
                        If Bare = False Then
                            sSize = Strings.RSet((fsd.StreamSize.Low + (fsd.StreamSize.High << 32)).ToString, 14) & "   "
                            sSize = Strings.LSet(GetFileCreateTime(StartingFolder & fsd.cStreamName.Replace(":$DATA", "")), 22) & " " & sSize
                        Else
                            sSize = ""
                        End If
                        Console.WriteLine(sSize & StartingFolder & fsd.cStreamName.Replace(":$DATA", ""))
                        If DeleteStreams Then
                            Dim bDelete As Boolean = True
                            If Pause Then
                                Console.WriteLine("Delete " & StartingFolder & fsd.cStreamName & " ?  [Y,N]")
                                Dim cKey As ConsoleKeyInfo = Console.ReadKey
                                If cKey.KeyChar.ToString.ToUpper <> "Y" Then
                                    bDelete = False
                                End If
                                Console.WriteLine()
                            End If
                            If bDelete Then
                                DeleteFile(StartingFolder & fsd.cStreamName.Replace(":$DATA", ""))
                                If iResDel = 0 Then
                                    iErr = GetLastError()
                                    Console.Error.WriteLine("DeleteFile failed with error: " & iErr.ToString & " while attempting to delete: '" & StartingFolder & fsd.cStreamName.Replace(":$DATA", "") & "'")
                                End If
                            End If
                        End If
                    End If
                End If
            End While
        End If
        If iErr <> 38 And iResult <> -1 Then
            FindClose(iResult)
        End If
        Try
            For Each sFile As FileInfo In GetFiles(StartingFolder)
                If Debug Then Console.WriteLine(sFile.Name)
                If (sFile.Attributes And IO.FileAttributes.ReparsePoint) <> IO.FileAttributes.ReparsePoint Then
                    sItem = sFile.Name
                    iResult = FindFirstStream(sFile.Name, StreamInfoLevels.FindStreamInfoStandard, fsd, 0)
                    iErr = GetLastError()
                    If iResult <> INVALID_HANDLE_VALUE Then
                        If fsd.cStreamName <> "::$DATA" AndAlso (Not fsd.cStreamName Like ":Zone.Identifier*" Or IgnoreZoneIdentifier = False) AndAlso fsd.cStreamName Like Pattern Then
                            If Bare = False Then
                                sSize = Strings.RSet((fsd.StreamSize.Low + (fsd.StreamSize.High << 32)).ToString, 14) & "   "
                                sSize = Strings.LSet(GetFileCreateTime(StartingFolder & fsd.cStreamName.Replace(":$DATA", "")), 22) & " " & sSize
                            Else
                                sSize = ""
                            End If
                            Console.WriteLine(sSize & sFile.Name & fsd.cStreamName.Replace(":$DATA", ""))
                            If DeleteStreams Then
                                Dim bDelete As Boolean = True
                                If Pause Then
                                    Console.WriteLine("Delete " & StartingFolder & fsd.cStreamName & " ?  [Y,N]")
                                    Dim cKey As ConsoleKeyInfo = Console.ReadKey
                                    If cKey.KeyChar.ToString.ToUpper <> "Y" Then
                                        bDelete = False
                                    End If
                                    Console.WriteLine()
                                End If
                                If bDelete Then
                                    DeleteFile(sFile.Name & fsd.cStreamName.Replace(":$DATA", ""))
                                    If iResDel = 0 Then
                                        iErr = GetLastError()
                                        Console.Error.WriteLine("DeleteFile failed with error: " & iErr.ToString & " while attempting to delete: '" & sFile.Name & fsd.cStreamName.Replace(":$DATA", "") & "'")
                                    End If
                                End If
                            End If
                        End If
                        Dim iRes As Int32 = -1
                        While iRes = -1
                            iRes = FindNextStream(iResult, fsd)
                            If iRes = 0 Then ' failed
                                iErr = GetLastError()
                                If iErr <> ERROR_HANDLE_EOF Then
                                    Console.Error.WriteLine("Error: " & iErr)
                                Else
                                    If iRes <> -1 Then
                                        FindClose(iRes)
                                    End If
                                End If
                                Exit While
                            Else ' we've found another stream - report the details
                                If fsd.cStreamName <> "::$DATA" AndAlso (Not fsd.cStreamName Like ":Zone.Identifier*" Or IgnoreZoneIdentifier = False) AndAlso fsd.cStreamName Like Pattern Then
                                    If Bare = False Then
                                        sSize = Strings.RSet((fsd.StreamSize.Low + (fsd.StreamSize.High << 32)).ToString, 14) & "   "
                                        sSize = Strings.LSet(GetFileCreateTime(StartingFolder & fsd.cStreamName.Replace(":$DATA", "")), 22) & " " & sSize
                                    Else
                                        sSize = ""
                                    End If
                                    Console.WriteLine(sSize & sFile.Name & fsd.cStreamName.Replace(":$DATA", ""))
                                    If DeleteStreams Then
                                        Dim bDelete As Boolean = True
                                        If Pause Then
                                            Console.WriteLine("Delete " & StartingFolder & fsd.cStreamName & " ?  [Y,N]")
                                            Dim cKey As ConsoleKeyInfo = Console.ReadKey
                                            If cKey.KeyChar.ToString.ToUpper <> "Y" Then
                                                bDelete = False
                                            End If
                                            Console.WriteLine()
                                        End If
                                        If bDelete Then
                                            DeleteFile(sFile.Name & fsd.cStreamName.Replace(":$DATA", ""))
                                            If iResDel = 0 Then
                                                iErr = GetLastError()
                                                Console.Error.WriteLine("DeleteFile failed with error: " & iErr.ToString & " while attempting to delete: '" & sFile.Name & fsd.cStreamName.Replace(":$DATA", "") & "'")
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End While
                    End If
                End If
            Next
            For Each folder As FileInfo In GetFolders(StartingFolder)
                If (folder.Attributes And IO.FileAttributes.ReparsePoint) <> IO.FileAttributes.ReparsePoint Then
                    sItem = folder.Name
                    GetStreams(StartingFolder:=folder.Name, IgnoreZoneIdentifier:=IgnoreZoneIdentifier, Pause:=Pause, Pattern:=Pattern, DeleteStreams:=DeleteStreams, Debug:=Debug, Bare:=Bare)
                End If
            Next
        Catch ex As System.UnauthorizedAccessException
            Console.Error.WriteLine(ex.Message)
        Catch ex As System.IO.IOException
            Console.WriteLine("")
            Console.Error.WriteLine(sItem & " -> " & ex.Message)
            If Pause Then
                Console.Error.WriteLine("Press any key to continue...")
                Console.ReadKey(True)
            End If
        Catch ex As Exception
            Console.Error.WriteLine()
            Console.Error.WriteLine("A problem occurred while inspecting " & sItem)
            Console.Error.WriteLine(ex.ToString)
            If Pause Then
                Console.Error.WriteLine("Press any key to continue...")
                Console.ReadKey(True)
            End If
        End Try

    End Sub

    Public Class FileInfo
        Property Name As String
        Property Attributes As UInt32
        Public Sub New()
            Name = ""
            Attributes = 0
        End Sub
        Public Sub New(ByVal Name As String, ByVal Attributes As UInt32)
            Me.Name = Name
            Me.Attributes = Attributes
        End Sub
    End Class

    Public Function GetFiles(ByVal FolderName As String) As List(Of FileInfo)
        Dim iHandle As IntPtr
        Dim ffd As New WIN32_FIND_DATA
        Dim Files As New List(Of FileInfo)
        Dim sPath As String
        If Not FolderName.EndsWith("\") Then
            FolderName = FolderName + "\"
        End If
        sPath = FolderName
        If Not FolderName.EndsWith("*") Then
            FolderName = FolderName + "*"
        End If
        iHandle = FindFirstFile(FolderName, ffd)
        If iHandle <> INVALID_HANDLE_VALUE Then
            If (ffd.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = 0 Then ' only capture non-directories
                If ffd.cFileName <> "." And ffd.cFileName <> ".." Then
                    Files.Add(New FileInfo(sPath & ffd.cFileName, ffd.dwFileAttributes))
                End If
            End If
            Dim iRes As Int32 = 1
            While iRes <> 0
                iRes = FindNextFile(iHandle, ffd)
                If iRes = 0 Then
                    ' no more files exist
                    Exit While
                Else
                    If (ffd.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = 0 Then ' only capture non-directories
                        If ffd.cFileName <> "." And ffd.cFileName <> ".." Then
                            Files.Add(New FileInfo(sPath & ffd.cFileName, ffd.dwFileAttributes))
                        End If
                    End If
                End If
            End While
        End If
        If iHandle <> INVALID_HANDLE_VALUE Then
            FindClose(iHandle)
        End If
        Return Files
    End Function

    Public Function GetFolders(ByVal FolderName As String) As List(Of FileInfo)
        Dim iHandle As IntPtr
        Dim ffd As New WIN32_FIND_DATA
        Dim Files As New List(Of FileInfo)
        Dim sPath As String
        If Not FolderName.EndsWith("\") Then
            FolderName = FolderName + "\"
        End If
        sPath = FolderName
        If Not FolderName.EndsWith("*") Then
            FolderName = FolderName + "*"
        End If
        iHandle = FindFirstFile(FolderName, ffd)
        If iHandle <> INVALID_HANDLE_VALUE Then
            If (ffd.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then ' only capture directories
                If ffd.cFileName <> "." And ffd.cFileName <> ".." Then
                    Files.Add(New FileInfo(sPath & ffd.cFileName, ffd.dwFileAttributes))
                End If
            End If
            Dim iRes As Int32 = 1
            While iRes <> 0
                iRes = FindNextFile(iHandle, ffd)
                If iRes = 0 Then
                    ' no more files exist
                    Exit While
                Else
                    If (ffd.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then ' only capture directories
                        If ffd.cFileName <> "." And ffd.cFileName <> ".." Then
                            Files.Add(New FileInfo(sPath & ffd.cFileName, ffd.dwFileAttributes))
                        End If
                    End If
                End If
            End While
        End If
        If iHandle <> INVALID_HANDLE_VALUE Then
            FindClose(iHandle)
        End If
        Return Files
    End Function

    Public Function GetFileCreateTime(ByVal FileName As String) As DateTime
        Dim bRes As IntPtr = NativeMethods.CreateFile(FileName, &H80000000, 7, Nothing, 3, 0, Nothing)
        If bRes <> NativeMethods.INVALID_HANDLE_VALUE Then
            Dim iFileCreate As NativeMethods.FILETIME
            Dim iFileAccess As NativeMethods.FILETIME
            Dim iFileWrite As NativeMethods.FILETIME
            bRes = NativeMethods.GetFileTime(bRes, iFileCreate, iFileAccess, iFileWrite)
            If bRes Then
                NativeMethods.CloseHandle(bRes)
                Return iFileCreate.DateTime
            Else
                Return Nothing
            End If
        Else
            Return Nothing
        End If
    End Function

End Module
