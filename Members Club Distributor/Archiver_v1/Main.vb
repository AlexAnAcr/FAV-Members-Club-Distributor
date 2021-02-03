Module Main
    Sub Main()
        If My.Application.CommandLineArgs.Count = 0 Then
            MsgBox("This program is not intended for self-study!", , "Error")
            End
        End If
        Dim com_line_arg(My.Application.CommandLineArgs.Count - 1) As String
        For i As UShort = 0 To My.Application.CommandLineArgs.Count - 1
            com_line_arg(i) = My.Application.CommandLineArgs(i)
        Next
        If com_line_arg.Length = 4 Then
            If com_line_arg(0) = "MCDmode:archiver,Code:2801" Then
                If com_line_arg(1) = "compress" Then
                    If My.Computer.FileSystem.DirectoryExists(com_line_arg(3)) Then
                        Try
                            Dim zip As New Ionic.Zip.ZipFile
                            zip.CompressionLevel = Ionic.Zlib.CompressionLevel.BestSpeed
                            zip.AlternateEncoding = System.Text.Encoding.UTF8
                            zip.AlternateEncodingUsage = Ionic.Zip.ZipOption.Always
                            zip.AddDirectory(com_line_arg(3), "")
                            zip.Save(com_line_arg(2))
                            My.Computer.FileSystem.WriteAllText(My.Application.Info.DirectoryPath & "\work_result.tmp", "sucess", False)
                        Catch
                            My.Computer.FileSystem.WriteAllText(My.Application.Info.DirectoryPath & "\work_result.tmp", "failed", False)
                        End Try
                    End If
                ElseIf com_line_arg(1) = "decompress" Then
                    If My.Computer.FileSystem.DirectoryExists(com_line_arg(3)) And My.Computer.FileSystem.FileExists(com_line_arg(2)) Then
                        Try
                            Dim zip As New Ionic.Zip.ZipFile(com_line_arg(2))
                            zip.CompressionLevel = Ionic.Zlib.CompressionLevel.BestSpeed
                            zip.AlternateEncoding = System.Text.Encoding.UTF8
                            zip.AlternateEncodingUsage = Ionic.Zip.ZipOption.Always
                            zip.ExtractAll(com_line_arg(3), Ionic.Zip.ExtractExistingFileAction.OverwriteSilently)
                            My.Computer.FileSystem.WriteAllText(My.Application.Info.DirectoryPath & "\work_result.tmp", "sucess", False)
                        Catch
                            My.Computer.FileSystem.WriteAllText(My.Application.Info.DirectoryPath & "\work_result.tmp", "failed", False)
                        End Try
                    End If
                Else
                    MsgBox("This program is not intended for self-study!", , "Error")
                End If
            Else
                MsgBox("This program is not intended for self-study!", , "Error")
            End If
        ElseIf com_line_arg.Length = 2 Then
            If com_line_arg(0) = "MCDmode:dir-delete,Code:2802" Then
                If My.Computer.FileSystem.DirectoryExists(com_line_arg(1)) Then
                    Try
                        My.Computer.FileSystem.DeleteDirectory(com_line_arg(1), FileIO.DeleteDirectoryOption.DeleteAllContents)
                    Catch
                    End Try
                End If
            ElseIf com_line_arg(0) = "MCDmode:testing,Code:2803" Then
                My.Computer.FileSystem.WriteAllText(My.Application.Info.DirectoryPath & "\check_info.tmp", com_line_arg(1), False)
            Else
                MsgBox("This program is not intended for self-study!", , "Error")
            End If
        Else
            MsgBox("This program is not intended for self-study!", , "Error")
        End If
    End Sub
End Module
