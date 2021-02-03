Public Class Member_to_distribute
    Private Sub Member_to_distribute_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        CheckBox1.Checked = True
        CheckBox2.Checked = False
        CheckBox3.Checked = False
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim indexes(-1) As String, added As Short = 0
        For Each i As DataGridViewRow In Main.DataGridView1.SelectedRows
            Array.Resize(indexes, indexes.Length + 1)
            indexes(indexes.Length - 1) = i.Cells.Item(0).Value
        Next
        For i As Short = 0 To indexes.Length - 1
            Dim pos As Short = -1
            For i1 As Short = 0 To Main.db_data.Count - 1
                If Main.db_data(i1).key = indexes(i) Then
                    pos = i1
                    Exit For
                End If
            Next
            Dim exits As Boolean = False
            For Each i1 As Member In Main.participators
                If i1.name = Main.db_data(pos).name Then
                    exits = True
                End If
            Next
            If exits = False Then
                If CheckBox1.Checked = True Then
                    Dim pos1 As Short = -1
                    For i1 As Short = 0 To Main.db_data(pos).skills.Length - 1
                        If Main.db_data(pos).skills(i1).ToLower.IndexOf("/dn ") = 0 Then
                            If CheckBox3.Checked = True Then
                                pos1 = i1
                                Exit For
                            Else
                                If pos1 > -1 Then
                                    pos1 = -2
                                    Exit For
                                Else
                                    pos1 = i1
                                End If
                            End If
                        End If
                    Next
                    If pos1 = -1 Then
                        Main.participators.Add(New Member)
                        Main.participators.Item(Main.participators.Count - 1).name = Main.db_data(pos).name
                        Main.participators.Item(Main.participators.Count - 1).dir = 0
                        Main.UpDate_Lists(0)
                        Main.Members_commands_redacted(True, True)
                        added += 1
                    ElseIf pos1 > -1 Then
                        Dim parts() As String = Main.db_data(pos).skills(pos1).Split(" ")
                        If parts.Length = 2 Then
                            If parts(1).Length <= 4 Then
                                Try
                                    Dim temp As UShort = parts(1)
                                    Main.participators.Add(New Member)
                                    Main.participators.Item(Main.participators.Count - 1).name = Main.db_data(pos).name
                                    Main.participators.Item(Main.participators.Count - 1).dir = temp
                                    Main.UpDate_Lists(0)
                                    Main.Members_commands_redacted(True, True)
                                    added += 1
                                Catch
                                    If CheckBox3.Checked = True Then
                                        Main.participators.Add(New Member)
                                        Main.participators.Item(Main.participators.Count - 1).name = Main.db_data(pos).name
                                        Main.participators.Item(Main.participators.Count - 1).dir = 0
                                        Main.UpDate_Lists(0)
                                        Main.Members_commands_redacted(True, True)
                                        added += 1
                                    End If
                                End Try
                            Else
                                If CheckBox3.Checked = True Then
                                    Main.participators.Add(New Member)
                                    Main.participators.Item(Main.participators.Count - 1).name = Main.db_data(pos).name
                                    Main.participators.Item(Main.participators.Count - 1).dir = 0
                                    Main.UpDate_Lists(0)
                                    Main.Members_commands_redacted(True, True)
                                    added += 1
                                End If
                            End If
                        Else
                            If CheckBox3.Checked = True Then
                                Main.participators.Add(New Member)
                                Main.participators.Item(Main.participators.Count - 1).name = Main.db_data(pos).name
                                Main.participators.Item(Main.participators.Count - 1).dir = 0
                                Main.UpDate_Lists(0)
                                Main.Members_commands_redacted(True, True)
                                added += 1
                            End If
                        End If
                    End If
                Else
                    Main.participators.Add(New Member)
                    Main.participators.Item(Main.participators.Count - 1).name = Main.db_data(pos).name
                    Main.participators.Item(Main.participators.Count - 1).dir = 0
                    Main.UpDate_Lists(0)
                    Main.Members_commands_redacted(True, True)
                    added += 1
                End If
            Else
                If CheckBox2.Checked = True Then
                    If CheckBox1.Checked = True Then
                        Dim pos1 As Short = -1
                        For i1 As Short = 0 To Main.db_data(pos).skills.Length - 1
                            If Main.db_data(pos).skills(i1).ToLower.IndexOf("/dn ") = 0 Then
                                If CheckBox3.Checked = True Then
                                    pos1 = i1
                                    Exit For
                                Else
                                    If pos1 > -1 Then
                                        pos1 = -2
                                        Exit For
                                    Else
                                        pos1 = i1
                                    End If
                                End If
                            End If
                        Next
                        If pos1 = -1 Then
                            Dim i1 As UShort = 0
                            While exits = True
                                exits = False
                                i1 += 1
                                For Each i2 As Member In Main.participators
                                    If i2.name = Main.db_data(pos).name & "-" & i1 Then
                                        exits = True
                                    End If
                                Next
                            End While
                            Main.participators.Add(New Member)
                            Main.participators.Item(Main.participators.Count - 1).name = Main.db_data(pos).name & "-" & i1
                            Main.participators.Item(Main.participators.Count - 1).dir = 0
                            Main.UpDate_Lists(0)
                            Main.Members_commands_redacted(True, True)
                            added += 1
                        ElseIf pos1 > -1 Then
                            Dim parts() As String = Main.db_data(pos).skills(pos1).Split(" ")
                            If parts.Length = 2 Then
                                If parts(1).Length <= 4 Then
                                    Try
                                        Dim temp As UShort = parts(1)
                                        Dim i1 As UShort = 0
                                        While exits = True
                                            exits = False
                                            i1 += 1
                                            For Each i2 As Member In Main.participators
                                                If i2.name = Main.db_data(pos).name & "-" & i1 Then
                                                    exits = True
                                                End If
                                            Next
                                        End While
                                        Main.participators.Add(New Member)
                                        Main.participators.Item(Main.participators.Count - 1).name = Main.db_data(pos).name & "-" & i1
                                        Main.participators.Item(Main.participators.Count - 1).dir = temp
                                        Main.UpDate_Lists(0)
                                        Main.Members_commands_redacted(True, True)
                                        added += 1
                                    Catch
                                        If CheckBox3.Checked = True Then
                                            Dim i1 As UShort = 0
                                            While exits = True
                                                exits = False
                                                i1 += 1
                                                For Each i2 As Member In Main.participators
                                                    If i2.name = Main.db_data(pos).name & "-" & i1 Then
                                                        exits = True
                                                    End If
                                                Next
                                            End While
                                            Main.participators.Add(New Member)
                                            Main.participators.Item(Main.participators.Count - 1).name = Main.db_data(pos).name & "-" & i1
                                            Main.participators.Item(Main.participators.Count - 1).dir = 0
                                            Main.UpDate_Lists(0)
                                            Main.Members_commands_redacted(True, True)
                                            added += 1
                                        End If
                                    End Try
                                Else
                                    If CheckBox3.Checked = True Then
                                        Dim i1 As UShort = 0
                                        While exits = True
                                            exits = False
                                            i1 += 1
                                            For Each i2 As Member In Main.participators
                                                If i2.name = Main.db_data(pos).name & "-" & i1 Then
                                                    exits = True
                                                End If
                                            Next
                                        End While
                                        Main.participators.Add(New Member)
                                        Main.participators.Item(Main.participators.Count - 1).name = Main.db_data(pos).name & "-" & i1
                                        Main.participators.Item(Main.participators.Count - 1).dir = 0
                                        Main.UpDate_Lists(0)
                                        Main.Members_commands_redacted(True, True)
                                        added += 1
                                    End If
                                End If
                            Else
                                If CheckBox3.Checked = True Then
                                    Dim i1 As UShort = 0
                                    While exits = True
                                        exits = False
                                        i1 += 1
                                        For Each i2 As Member In Main.participators
                                            If i2.name = Main.db_data(pos).name & "-" & i1 Then
                                                exits = True
                                            End If
                                        Next
                                    End While
                                    Main.participators.Add(New Member)
                                    Main.participators.Item(Main.participators.Count - 1).name = Main.db_data(pos).name & "-" & i1
                                    Main.participators.Item(Main.participators.Count - 1).dir = 0
                                    Main.UpDate_Lists(0)
                                    Main.Members_commands_redacted(True, True)
                                    added += 1
                                End If
                            End If
                        End If
                    Else
                        Dim i1 As UShort = 0
                        While exits = True
                            exits = False
                            i1 += 1
                            For Each i2 As Member In Main.participators
                                If i2.name = Main.db_data(pos).name & "-" & i1 Then
                                    exits = True
                                End If
                            Next
                        End While
                        Main.participators.Add(New Member)
                        Main.participators.Item(Main.participators.Count - 1).name = Main.db_data(pos).name & "-" & i1
                        Main.participators.Item(Main.participators.Count - 1).dir = 0
                        Main.UpDate_Lists(0)
                        Main.Members_commands_redacted(True, True)
                        added += 1
                    End If
                End If
            End If
        Next
        MsgBox("Добавлено участников: " & added & " из " & indexes.Length, MsgBoxStyle.OkOnly, "Информация")
        Me.Close()
    End Sub
End Class
