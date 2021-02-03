Public Class Rand
    Public work_mode As Short, smooth(1), nearby(1) As Boolean, smooth_strench(1) As Short
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
    Private Sub Rand_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        If work_mode = 0 Then
            CheckBox3.Checked = smooth(0)
            CheckBox1.Checked = nearby(0)
            ComboBox1.SelectedIndex = smooth_strench(0)
            CheckBox3.Text = "Уравнять кол-во участников в командах"
            CheckBox1.Text = "Допускать соседние элементы рядом"
            Label1.Text = "Допустимое различие команд:"
        Else
            CheckBox3.Checked = smooth(1)
            CheckBox1.Checked = nearby(1)
            ComboBox1.SelectedIndex = smooth_strench(1)
            CheckBox3.Text = "Уравнять кол-во команд в турнирах"
            CheckBox1.Text = "Допускать соседние элементы рядом"
            Label1.Text = "Допустимое различие турниров:"
        End If
        If CheckBox3.Checked = True Then
            ComboBox1.Enabled = True
        Else
            ComboBox1.Enabled = False
        End If
    End Sub
    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            ComboBox1.Enabled = True
        Else
            ComboBox1.Enabled = False
        End If
    End Sub
    Private Sub Rand_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If work_mode = 0 Then
            If nearby(0) <> CheckBox1.Checked Or smooth(0) <> CheckBox3.Checked Or smooth_strench(0) <> ComboBox1.SelectedIndex Then
                Main.Members_commands_redacted(True, True)
            End If
            smooth(0) = CheckBox3.Checked
            nearby(0) = CheckBox1.Checked
            smooth_strench(0) = ComboBox1.SelectedIndex
        Else
            If nearby(1) <> CheckBox1.Checked Or smooth(1) <> CheckBox3.Checked Or smooth_strench(1) <> ComboBox1.SelectedIndex Then
                Main.Members_commands_redacted(False, True)
            End If
            smooth(1) = CheckBox3.Checked
            nearby(1) = CheckBox1.Checked
            smooth_strench(1) = ComboBox1.SelectedIndex
        End If
    End Sub
End Class
