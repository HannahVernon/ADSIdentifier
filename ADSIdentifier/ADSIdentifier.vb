Option Compare Text
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

    <StructLayout(LayoutKind.Explicit)>
    Public Structure LARGE_INTEGER
        <FieldOffset(0)> Dim Low As Int32
        <FieldOffset(4)> Dim High As Int32
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Public Structure _WIN32_FIND_STREAM_DATA
        Dim StreamSize As LARGE_INTEGER ' <MarshalAs(UnmanagedType.I8)> 
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=296)> Dim cStreamName As String ' 260 (max_path) + 36
    End Structure

    <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, EntryPoint:="FindFirstStreamW")>
    Function FindFirstStream(<MarshalAs(UnmanagedType.LPWStr)> ByVal lpFileName As String, ByVal InfoLevel As StreamInfoLevels, ByRef lpFindStreamData As _WIN32_FIND_STREAM_DATA, ByVal dwFlags As Int32) As IntPtr
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, EntryPoint:="FindNextStreamW")>
    Function FindNextStream(ByVal hFindStream As IntPtr, ByRef lpFindStreamData As _WIN32_FIND_STREAM_DATA) As Int32
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

End Module

Module ADSIdentifier

    Sub Main()
        Dim sFolder As String = ""
        Dim bPause As Boolean = False
        Dim bIgnoreZoneIdentifier As Boolean = False
        Dim sPattern As String = ""
        Dim bDeleteStreams As Boolean = False
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
        Next
        If sPattern = "" Then sPattern = "*"
        If Left(sPattern, 1) <> "*" Then sPattern = "*" & sPattern
        If Right(sPattern, 1) <> "*" Then sPattern = sPattern & "*"
        If sFolder <> "" Then
            GetStreams(sFolder, bIgnoreZoneIdentifier, bPause, sPattern, bDeleteStreams)
        Else
            Console.WriteLine("Useage is:   ADSIdentifier.exe /Folder:<starting_folder_name>")
            Console.WriteLine("                  [/P] or [/Pause] - pause before exiting")
            Console.WriteLine("                  [/IZI] or [/IgnoreZoneIdentifier] - ignore :Zone.Identifier streams")
            Console.WriteLine("                  [/Pattern:<xyz>] - only find Alternate Data Streams matching <xyz>")
            Console.WriteLine("                  [/Remove] - remove Alternate Data Streams that have been found matching the other parameters")
        End If
        If bPause Then
            Console.WriteLine("")
            Console.WriteLine("Press any key to exit...")
            Console.ReadKey()
        End If
    End Sub

    Sub GetStreams(ByVal StartingFolder As String, ByVal IgnoreZoneIdentifier As Boolean, ByVal Pause As Boolean, ByVal Pattern As String, ByVal DeleteStreams As Boolean)
        Dim fsd As New _WIN32_FIND_STREAM_DATA
        Dim sItem As String = StartingFolder
        Dim iErr As Int32 = 0
        Dim iResult As IntPtr = FindFirstStream(StartingFolder, StreamInfoLevels.FindStreamInfoStandard, fsd, 0)
        iErr = GetLastError()
        Dim iResDel As Int32
        If iResult <> INVALID_HANDLE_VALUE Then
            If (Not fsd.cStreamName Like ":Zone.Identifier*" Or IgnoreZoneIdentifier = False) And fsd.cStreamName Like Pattern Then
                Console.WriteLine(StartingFolder & fsd.cStreamName.Replace(":$DATA", ""))
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
                        iResDel = DeleteFile(StartingFolder & fsd.cStreamName)
                        If iResDel = 0 Then
                            iErr = GetLastError()
                            Console.Error.WriteLine("DeleteFile failed with error: " & iErr.ToString)
                        End If
                    End If
                End If
            End If
            Dim iRes As Int32 = 0
            While iRes = 0
                iRes = FindNextStream(iResult, fsd)
                If iRes = 0 Then ' failed
                    iErr = GetLastError()
                    If iErr <> ERROR_HANDLE_EOF Then
                        Console.Error.WriteLine("Error: " & iErr)
                    End If
                    Exit While
                Else
                    If (Not fsd.cStreamName Like ":Zone.Identifier*" Or IgnoreZoneIdentifier = False) And fsd.cStreamName Like Pattern Then
                        Console.WriteLine(StartingFolder & fsd.cStreamName.Replace(":$DATA", ""))
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
                                DeleteFile(StartingFolder & fsd.cStreamName)
                                If iResDel = 0 Then
                                    iErr = GetLastError()
                                    Console.Error.WriteLine("DeleteFile failed with error: " & iErr.ToString)
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
            For Each sFile As String In My.Computer.FileSystem.GetFiles(StartingFolder)
                Dim fa As IO.FileAttributes = My.Computer.FileSystem.GetFileInfo(sFile).Attributes
                If (fa And IO.FileAttributes.ReparsePoint) <> IO.FileAttributes.ReparsePoint Then
                    sItem = sFile
                    iResult = FindFirstStream(sFile, StreamInfoLevels.FindStreamInfoStandard, fsd, 0)
                    iErr = GetLastError()
                    If iResult <> INVALID_HANDLE_VALUE Then
                        If fsd.cStreamName <> "::$DATA" And (Not fsd.cStreamName Like ":Zone.Identifier*" Or IgnoreZoneIdentifier = False) And fsd.cStreamName Like Pattern Then
                            Console.WriteLine(sFile & fsd.cStreamName.Replace(":$DATA", ""))
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
                                    DeleteFile(sFile & fsd.cStreamName)
                                    If iResDel = 0 Then
                                        iErr = GetLastError()
                                        Console.Error.WriteLine("DeleteFile failed with error: " & iErr.ToString)
                                    End If
                                End If
                            End If
                        End If
                        Dim iRes As Int32 = 0
                        While iRes = 0
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
                            Else
                                If fsd.cStreamName <> "::$DATA" And (Not fsd.cStreamName Like ":Zone.Identifier*" Or IgnoreZoneIdentifier = False) And fsd.cStreamName Like Pattern Then
                                    Console.WriteLine(sFile & fsd.cStreamName.Replace(":$DATA", ""))
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
                                            DeleteFile(sFile & fsd.cStreamName)
                                            If iResDel = 0 Then
                                                iErr = GetLastError()
                                                Console.Error.WriteLine("DeleteFile failed with error: " & iErr.ToString)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End While
                    End If
                End If
            Next
            For Each folder As String In My.Computer.FileSystem.GetDirectories(StartingFolder, FileIO.SearchOption.SearchTopLevelOnly)
                Dim fa As IO.FileAttributes = My.Computer.FileSystem.GetFileInfo(folder).Attributes
                If (fa And IO.FileAttributes.ReparsePoint) <> IO.FileAttributes.ReparsePoint Then
                    sItem = folder
                    GetStreams(StartingFolder:=folder, IgnoreZoneIdentifier:=IgnoreZoneIdentifier, Pause:=Pause, Pattern:=Pattern, DeleteStreams:=DeleteStreams)
                End If
            Next
        Catch ex As System.UnauthorizedAccessException
            Console.Error.WriteLine(ex.Message)
        Catch ex As System.IO.IOException
            Console.WriteLine("")
            Console.Error.WriteLine(sItem & " -> " & ex.Message)
            If Pause Then
                Console.Error.WriteLine("Press any key to continue...")
                Console.ReadKey()
            End If
        Catch ex As Exception
            Console.Error.WriteLine()
            Console.Error.WriteLine("A problem occurred while inspecting " & sItem)
            Console.Error.WriteLine(ex.ToString)
            If Pause Then
                Console.Error.WriteLine("Press any key to continue...")
                Console.ReadKey()
            End If
        End Try

    End Sub

End Module
