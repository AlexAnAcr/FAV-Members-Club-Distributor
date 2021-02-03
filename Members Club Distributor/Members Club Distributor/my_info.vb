Public Class my_info
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
    Private Sub my_info_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        TextBox1.Text = Main.info_set(0)
        TextBox2.Text = Main.info_set(1)
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or Main.Banned_simvols_test(TextBox1.Text) = False Then
            MsgBox("Вы не задали имя пользователя или оно содержит запрещённые символы!",, "Сообщение")
        Else
            If TextBox2.Text = "" Or Main.Banned_simvols_test(TextBox2.Text) = False Then
                MsgBox("Вы не задали название организации или оно содержит запрещённые символы!",, "Сообщение")
            Else
                Main.info_set(0) = TextBox1.Text
                Main.info_set(1) = TextBox2.Text
                Main.Info_set_write()
                Me.Close()
            End If
        End If
    End Sub
End Class
