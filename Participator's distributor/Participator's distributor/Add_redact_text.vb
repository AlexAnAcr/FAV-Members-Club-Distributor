Public Class Add_redact_text
    Public work_mode As UShort
    Private Sub Add_redact_text_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        ListBox1.Items.Clear()
        If work_mode = 1 Then
            Label1.Text = "Выберите участника:"
            For Each i As Member In Main.participators
                ListBox1.Items.Add(i.name)
            Next
        ElseIf work_mode = 2 Then
            Label1.Text = "Выберите команду:"
            For Each i As Command In Main.commands
                ListBox1.Items.Add(i.name)
            Next
        ElseIf work_mode = 3 Then
            Label1.Text = "Выберите турнир:"
            For Each i As Tournier In Main.tourniers
                ListBox1.Items.Add(i.name)
            Next
        ElseIf work_mode = 4 Then
            Label1.Text = "Выберите таймер:"
            For Each i As Timer_temp In Main.timers
                ListBox1.Items.Add(i.name)
            Next
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ListBox1.SelectedIndex <> -1 Then
            Me.DialogResult = DialogResult.OK
            Me.Close()
        Else
            MsgBox("Вы ничего не выбрали!",, "Ошибка")
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
End Class
