Public Class Redact
    Public work_mode As UShort
    Private Sub Redact_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        If work_mode = 0 Then
            Label1.Text = "Измените имя:"
        Else
            Label1.Text = "Измените название:"
        End If
        If work_mode = 0 Then
            TextBox1.Text = Main.participators.Item(Main.ListBox1.SelectedIndex).name
            TextBox3.Text = Main.participators.Item(Main.ListBox1.SelectedIndex).dir
        ElseIf work_mode = 2 Then
            TextBox1.Text = Main.tourniers.Item(Main.ListBox6.SelectedIndex).name
            TextBox3.Text = Main.tourniers.Item(Main.ListBox6.SelectedIndex).dir
        Else
            TextBox1.Text = Main.commands.Item(Main.ListBox3.SelectedIndex).name
            TextBox3.Text = Main.commands.Item(Main.ListBox3.SelectedIndex).dir
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim temp As UInteger = TextBox3.Text
        Catch
            MsgBox("Вы неверно ввели данные!",, "Ошибка")
            Exit Sub
        End Try
        If work_mode = 0 Then
            Dim exits As Boolean = False
            For i As Integer = 0 To Main.participators.Count - 1
                If Main.participators(i).name = TextBox1.Text And i <> Main.ListBox1.SelectedIndex Then
                    exits = True
                    Exit For
                End If
            Next
            If exits = True Then
                MsgBox("Участник с таким именем уже имеется в списке!",, "Ошибка")
            Else
                If Main.participators.Item(Main.ListBox1.SelectedIndex).name <> TextBox1.Text Or Main.participators.Item(Main.ListBox1.SelectedIndex).dir <> TextBox3.Text Then
                    Main.Members_commands_redacted(True, True)
                End If
                Main.participators.Item(Main.ListBox1.SelectedIndex).name = TextBox1.Text
                Main.participators.Item(Main.ListBox1.SelectedIndex).dir = TextBox3.Text
                Main.UpDate_Lists(0)
                Me.Close()
            End If
        ElseIf work_mode = 2 Then
            Dim exits As Boolean = False
            For i As Integer = 0 To Main.commands.Count - 1
                If Main.commands(i).name = TextBox1.Text And i <> Main.ListBox3.SelectedIndex Then
                    exits = True
                    Exit For
                End If
            Next
            If exits = True Then
                MsgBox("Турнир с таким названием уже имеется в списке!",, "Ошибка")
            Else
                If Main.commands.Item(Main.ListBox6.SelectedIndex).name <> TextBox1.Text Or Main.commands.Item(Main.ListBox6.SelectedIndex).dir <> TextBox3.Text Then
                    Main.Members_commands_redacted(False, True)
                End If
                Main.tourniers.Item(Main.ListBox6.SelectedIndex).name = TextBox1.Text
                Main.tourniers.Item(Main.ListBox6.SelectedIndex).dir = TextBox3.Text
                Main.UpDate_Lists(0)
                Me.Close()
            End If
        Else
            Dim exits As Boolean = False
            For i As Integer = 0 To Main.commands.Count - 1
                If Main.commands(i).name = TextBox1.Text And i <> Main.ListBox3.SelectedIndex Then
                    exits = True
                    Exit For
                End If
            Next
            If exits = True Then
                MsgBox("Команда с таким названием уже имеется в списке!",, "Ошибка")
            Else
                If Main.commands.Item(Main.ListBox3.SelectedIndex).name <> TextBox1.Text Or Main.commands.Item(Main.ListBox3.SelectedIndex).dir <> TextBox3.Text Then
                    Main.Members_commands_redacted(True, True)
                    Main.Members_commands_redacted(False, True)
                End If
                Main.commands.Item(Main.ListBox3.SelectedIndex).name = TextBox1.Text
                Main.commands.Item(Main.ListBox3.SelectedIndex).dir = TextBox3.Text
                Main.UpDate_Lists(0)
                Me.Close()
            End If
        End If
    End Sub
End Class
