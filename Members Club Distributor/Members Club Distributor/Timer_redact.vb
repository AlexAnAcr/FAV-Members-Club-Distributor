Public Class Timer_redact
    Private Sub Timer_redact_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Dim time_data As hms_temp = Main.Get_hms(Main.timer_func.interval)
        TextBox1.Text = time_data.hours
        TextBox2.Text = time_data.minutes
        TextBox3.Text = time_data.seconds
        CheckBox1.Checked = Main.timer_func.work_messange(2)
        CheckBox2.Checked = Main.timer_func.work_messange(3)
        CheckBox3.Checked = Main.timer_func.work_messange(4)
        CheckBox4.Checked = Main.timer_func.work_messange(1)
        CheckBox5.Checked = Main.timer_func.work_messange(0)
        CheckBox6.Checked = False
        If Main.timer_func.select_time(0) = 5 Then
            ComboBox1.SelectedIndex = 0
        ElseIf Main.timer_func.select_time(0) = 10 Then
            ComboBox1.SelectedIndex = 1
        ElseIf Main.timer_func.select_time(0) = 30 Then
            ComboBox1.SelectedIndex = 2
        ElseIf Main.timer_func.select_time(0) = 60 Then
            ComboBox1.SelectedIndex = 3
        ElseIf Main.timer_func.select_time(0) = 300 Then
            ComboBox1.SelectedIndex = 4
        End If
        If Main.timer_func.select_time(1) = 5 Then
            ComboBox2.SelectedIndex = 0
        ElseIf Main.timer_func.select_time(1) = 10 Then
            ComboBox2.SelectedIndex = 1
        ElseIf Main.timer_func.select_time(1) = 30 Then
            ComboBox2.SelectedIndex = 2
        ElseIf Main.timer_func.select_time(1) = 60 Then
            ComboBox2.SelectedIndex = 3
        ElseIf Main.timer_func.select_time(1) = 300 Then
            ComboBox2.SelectedIndex = 4
        End If
        If Main.timer_func.select_time(2) = 5 Then
            ComboBox3.SelectedIndex = 0
        ElseIf Main.timer_func.select_time(2) = 10 Then
            ComboBox3.SelectedIndex = 1
        ElseIf Main.timer_func.select_time(2) = 30 Then
            ComboBox3.SelectedIndex = 2
        ElseIf Main.timer_func.select_time(2) = 60 Then
            ComboBox3.SelectedIndex = 3
        ElseIf Main.timer_func.select_time(2) = 300 Then
            ComboBox3.SelectedIndex = 4
        End If
        If Main.settings(3) = "" Then
            RadioButton1.Checked = True
        Else
            If My.Computer.FileSystem.FileExists(Main.settings(3)) Then
                If My.Computer.FileSystem.GetFileInfo(Main.settings(3)).Extension.ToLower = ".mp3" Then
                    RadioButton2.Checked = True
                    TextBox4.Text = Main.settings(3)
                Else
                    RadioButton1.Checked = True
                End If
            Else
                RadioButton1.Checked = True
            End If
        End If
    End Sub
    Private Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text.Length > 1 Then
            If Mid(TextBox1.Text, 1, 1) = "0" Then TextBox1.Text = Mid(TextBox1.Text, 2)
            TextBox1.SelectionStart = TextBox1.Text.Length
        End If
    End Sub
    Private Sub TextBox1_Leave(sender As Object, e As EventArgs) Handles TextBox1.Leave
        If TextBox1.Text = "" Then TextBox1.Text = 0
    End Sub
    Private Sub TextBox2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged
        If TextBox2.Text.Length > 1 Then
            If Mid(TextBox2.Text, 1, 1) = "0" Then TextBox2.Text = Mid(TextBox2.Text, 2)
            TextBox2.SelectionStart = TextBox2.Text.Length
        End If
    End Sub
    Private Sub TextBox2_Leave(sender As Object, e As EventArgs) Handles TextBox2.Leave
        If TextBox2.Text = "" Then TextBox2.Text = 0
    End Sub
    Private Sub TextBox3_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged
        If TextBox3.Text.Length > 1 Then
            If Mid(TextBox3.Text, 1, 1) = "0" Then TextBox3.Text = Mid(TextBox3.Text, 2)
            TextBox3.SelectionStart = TextBox3.Text.Length
        End If
    End Sub
    Private Sub TextBox3_Leave(sender As Object, e As EventArgs) Handles TextBox3.Leave
        If TextBox3.Text = "" Then TextBox3.Text = 0
    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            ComboBox1.Enabled = True
        Else
            ComboBox1.Enabled = False
        End If
    End Sub
    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked Then
            ComboBox2.Enabled = True
        Else
            ComboBox2.Enabled = False
        End If
    End Sub
    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked Then
            ComboBox3.Enabled = True
        Else
            ComboBox3.Enabled = False
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked Then
            TextBox4.Enabled = True
            Button3.Enabled = True
        End If
    End Sub
    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked Then
            TextBox4.Enabled = False
            Button3.Enabled = False
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim result As DialogResult = OpenFile.ShowDialog
        If result = DialogResult.OK Then
            If My.Computer.FileSystem.GetFileInfo(OpenFile.FileName).Extension.ToLower = ".mp3" Then
                TextBox4.Text = OpenFile.FileName
            Else
                MsgBox("Выбранный файл имеет формат не mp3!",, "Ошибка")
            End If
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim temp_num1 As UShort = TextBox1.Text
            Dim temp_num2 As UShort = TextBox2.Text
            Dim temp_num3 As UShort = TextBox3.Text
            If temp_num3 >= 60 Then Error 1
            If temp_num2 >= 60 Then Error 1
            If temp_num1 > 60 Then Error 1
            Dim time As UInteger = (temp_num1 * 60 + temp_num2) * 60 + temp_num3
            If time = 0 Then Error 1
            Main.timer_func.interval = time
            Main.timer_func.new_interval = time
            Main.timer_func.work_messange(2) = CheckBox1.Checked
            Main.timer_func.work_messange(3) = CheckBox2.Checked
            Main.timer_func.work_messange(4) = CheckBox3.Checked
            Main.timer_func.work_messange(1) = CheckBox4.Checked
            Main.timer_func.work_messange(0) = CheckBox5.Checked
            If ComboBox1.SelectedIndex = 0 Then
                Main.timer_func.select_time(0) = 5
            ElseIf ComboBox1.SelectedIndex = 1 Then
                Main.timer_func.select_time(0) = 10
            ElseIf ComboBox1.SelectedIndex = 2 Then
                Main.timer_func.select_time(0) = 30
            ElseIf ComboBox1.SelectedIndex = 3 Then
                Main.timer_func.select_time(0) = 60
            ElseIf ComboBox1.SelectedIndex = 4 Then
                Main.timer_func.select_time(0) = 300
            End If
            If ComboBox2.SelectedIndex = 0 Then
                Main.timer_func.select_time(1) = 5
            ElseIf ComboBox2.SelectedIndex = 1 Then
                Main.timer_func.select_time(1) = 10
            ElseIf ComboBox2.SelectedIndex = 2 Then
                Main.timer_func.select_time(1) = 30
            ElseIf ComboBox2.SelectedIndex = 3 Then
                Main.timer_func.select_time(1) = 60
            ElseIf ComboBox2.SelectedIndex = 4 Then
                Main.timer_func.select_time(1) = 300
            End If
            If ComboBox3.SelectedIndex = 0 Then
                Main.timer_func.select_time(2) = 5
            ElseIf ComboBox3.SelectedIndex = 1 Then
                Main.timer_func.select_time(2) = 10
            ElseIf ComboBox3.SelectedIndex = 2 Then
                Main.timer_func.select_time(2) = 30
            ElseIf ComboBox3.SelectedIndex = 3 Then
                Main.timer_func.select_time(2) = 60
            ElseIf ComboBox3.SelectedIndex = 4 Then
                Main.timer_func.select_time(2) = 300
            End If
            If RadioButton1.Checked = True Then
                Main.settings(3) = ""
                Main.timer_resourse = Main.player1.newMedia(Main.TEMP_WAY & "\MCD-temp-47591795970905\tds.mp3")
            ElseIf RadioButton2.Checked = True Then
                Main.settings(3) = TextBox4.Text
                Main.timer_resourse = Main.player1.newMedia(Main.settings(3))
            End If
            Dim reg As Microsoft.Win32.RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey("Software", True), exist_reg As Boolean = False
            For Each i As String In reg.GetSubKeyNames
                If i = "Members Club Distributor" Then
                    exist_reg = True
                    Exit For
                End If
            Next
            If exist_reg = True Then
                reg.Close()
                reg = My.Computer.Registry.CurrentUser.OpenSubKey("Software\Members Club Distributor", True)
                reg.OpenSubKey("Members Club Distributor", True)
                Dim names() As String = {"TimerSoundWay"}, reg_values() As String = reg.GetValueNames
                For i As Short = 0 To names.Length - 1
                    exist_reg = False
                    For Each i1 As String In reg_values
                        If names(i) = i1 Then
                            exist_reg = True
                            Exit For
                        End If
                    Next
                    If exist_reg = False Then
                        reg.Close()
                        MsgBox("Нет доступа к настройкам программы!",, "Ошибка")
                    End If
                Next
                reg.SetValue("TimerSoundWay", Main.settings(3), Microsoft.Win32.RegistryValueKind.String)
                reg.Close()
            Else
                reg.Close()
                MsgBox("Нет доступа к настройкам программы!",, "Ошибка")
            End If
            If CheckBox6.Checked Then
                Main.Timer_timer.Start()
                Main.timer_button_mode = 2
                Main.Button24.Text = "Стоп"
                Main.Button27.Enabled = False
                full_screen.Button2.Enabled = False
                If Main.timer_func.work_messange(1) Then
                    Messanger.Add_messange("Таймер запущен.", "timer start/stop", True)
                End If
                If Main.timer_func.work_messange(0) Then
                    Dim temp_time As hms_temp = Main.Get_hms(Main.timer_func.new_interval)
                    Messanger.Add_messange("Время таймера: " & Format(temp_time.hours, "00") & ":" & Format(temp_time.minutes, "00") & ":" & Format(temp_time.seconds, "00"), "timer progress", False)
                End If
            End If
            Main.Button24.Enabled = True
            Main.Button25.Enabled = True
            Main.Button26.Enabled = True
            Main.Renovate_timer(True)
            Me.Close()
        Catch
            MsgBox("Неверно задано время таймера!",, "Ошибка")
        End Try
    End Sub
End Class
