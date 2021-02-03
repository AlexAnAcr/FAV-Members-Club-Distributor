Public Class Add_redact_timer
    Public work_mode As UShort
    Dim temp_mem As New Timer_temp
    Private Sub Add_redact_timer_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        temp_mem = New Timer_temp
        If work_mode = 1 Then
            TextBox1.Text = ""
            TextBox2.Text = "0"
            TextBox3.Text = "0"
            TextBox4.Text = "0"
            TextBox5.Text = ""
            CheckBox1.Checked = False
            CheckBox2.Checked = True
            CheckBox3.Checked = True
            CheckBox4.Checked = True
            CheckBox5.Checked = True
            ComboBox1.SelectedIndex = 0
            ListBox1.Items.Clear()
            ComboBox2.SelectedIndex = 0
            ListBox1.Enabled = False
            Button1.Enabled = False
            Button2.Enabled = False
            Label10.Visible = False
            temp_mem.work_messange(0) = True
            temp_mem.work_messange(1) = True
            temp_mem.work_messange(2) = True
            temp_mem.work_messange(3) = True
        ElseIf work_mode = 2 Then
        End If
        If temp_mem.work_messange(0) = True Then
            Label10.Visible = False
            CheckBox2.Checked = True
        Else
            Label10.Visible = True
            CheckBox2.Checked = False
        End If
        If temp_mem.active = True Then
            CheckBox1.Checked = True
        Else
            CheckBox1.Checked = False
        End If
        If temp_mem.work_messange(1) = True Then
            CheckBox5.Checked = True
        Else
            CheckBox5.Checked = False
        End If
        If temp_mem.work_messange(2) = True Then
            CheckBox3.Checked = True
        Else
            CheckBox3.Checked = False
        End If
        If temp_mem.work_messange(3) = True Then
            CheckBox4.Checked = True
        Else
            CheckBox4.Checked = False
        End If
        TextBox1.Text = temp_mem.name
        ComboBox2.SelectedIndex = temp_mem.action
        ComboBox1.SelectedIndex = temp_mem.connect_type
        Dim hms As g_hms = Get_hms(temp_mem.interval)
        TextBox2.Text = hms.hour
        TextBox3.Text = hms.minute
        TextBox4.Text = hms.second
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.Close()
    End Sub
    Public Function Get_hms(seconds As String) As g_hms
        Dim left As Integer = seconds
        Dim hours_left As Integer = Int((left / 60) / 60)
        left -= hours_left * 3600
        Dim minutes_left As Integer = Int(left / 60)
        left -= minutes_left * 60
        Dim hms As New g_hms(hours_left, minutes_left, left)
        Return hms
    End Function
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        TextBox1.Text = ""
        TextBox2.Text = "0"
        TextBox3.Text = "0"
        TextBox4.Text = "0"
        TextBox5.Text = ""
        CheckBox1.Checked = False
        CheckBox2.Checked = True
        CheckBox3.Checked = True
        CheckBox4.Checked = True
        CheckBox5.Checked = True
        ComboBox1.SelectedIndex = 0
        ListBox1.Items.Clear()
        ComboBox2.SelectedIndex = 0
        ListBox1.Enabled = False
        Button1.Enabled = False
        Button2.Enabled = False
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Add_redact_text.work_mode = ComboBox1.SelectedIndex
        Dim result As DialogResult = Add_redact_text.ShowDialog
        If result = DialogResult.OK Then
            If temp_mem.connect_type = 1 Then
                temp_mem.members.Add(Main.participators(Add_redact_text.ListBox1.SelectedIndex))
            ElseIf temp_mem.connect_type = 2 Then
                temp_mem.commands.Add(Main.commands(Add_redact_text.ListBox1.SelectedIndex))
            ElseIf temp_mem.connect_type = 3 Then
                temp_mem.tourniers.Add(Main.tourniers(Add_redact_text.ListBox1.SelectedIndex))
            End If
            Renovate_list()
        End If
    End Sub
    Private Sub Renovate_list()
        ListBox1.Items.Clear()
        If temp_mem.connect_type = 1 Then
            For Each i As Member In temp_mem.members
                ListBox1.Items.Add(i.name)
            Next
        ElseIf temp_mem.connect_type = 2 Then
            For Each i As Command In temp_mem.commands
                ListBox1.Items.Add(i.name)
            Next
        ElseIf temp_mem.connect_type = 3 Then
            For Each i As Tournier In temp_mem.tourniers
                ListBox1.Items.Add(i.name)
            Next
        End If
        If ComboBox2.SelectedIndex = 2 Then
            TextBox5.Text = temp_mem.action_data
        Else
            TextBox5.Text = ""
        End If
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        temp_mem.connect_type = ComboBox1.SelectedIndex
        If ComboBox1.SelectedIndex = 0 Then
            ListBox1.Enabled = False
            Button1.Enabled = False
            Button2.Enabled = False
        Else
            ListBox1.Enabled = True
            Button1.Enabled = True
            Button2.Enabled = True
        End If
        Renovate_list()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If ListBox1.SelectedIndex <> -1 Then
            If temp_mem.connect_type = 1 Then
                temp_mem.members.RemoveAt(ListBox1.SelectedIndex)
            ElseIf temp_mem.connect_type = 2 Then
                temp_mem.commands.RemoveAt(ListBox1.SelectedIndex)
            ElseIf temp_mem.connect_type = 3 Then
                temp_mem.tourniers.RemoveAt(ListBox1.SelectedIndex)
            End If
            Renovate_list()
        Else
            MsgBox("Вы не выбрали элемент!",, "Ошибка")
        End If
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        temp_mem.action = ComboBox2.SelectedIndex
        If ComboBox2.SelectedIndex = 2 Then
            Button3.Enabled = True
        Else
            Button3.Enabled = False
        End If
        Renovate_list()
    End Sub
    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked Then
            Label10.Visible = False
            temp_mem.work_messange(0) = True
        Else
            Label10.Visible = True
            temp_mem.work_messange(0) = False
        End If
    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            temp_mem.active = True
        Else
            temp_mem.active = False
        End If
    End Sub
    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked Then
            temp_mem.work_messange(1) = True
        Else
            temp_mem.work_messange(1) = False
        End If
    End Sub
    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked Then
            temp_mem.work_messange(2) = True
        Else
            temp_mem.work_messange(2) = False
        End If
    End Sub
    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked Then
            temp_mem.work_messange(3) = True
        Else
            temp_mem.work_messange(3) = False
        End If
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If TextBox1.Text = "" Then
            MsgBox("Вы не указали имя таймера!",, "Ошибка")
            Exit Sub
        End If
        Dim exits As Boolean = False
        For Each i As Timer_temp In Main.timers
            If i.name = TextBox1.Text Then
                exits = True
            End If
        Next
        If exits = False Then
            Dim interval As UInteger
            Try
                Dim temp_dat1 As UShort = TextBox2.Text
                If temp_dat1 > 48 Then Error 1
                Dim temp_dat2 As UShort = TextBox3.Text
                If temp_dat2 > 59 Then Error 1
                Dim temp_dat3 As UShort = TextBox4.Text
                If temp_dat3 > 59 Then Error 1
                interval = (temp_dat1 * 60 + temp_dat2) * 60 + temp_dat3
                If interval = 0 Then Error 1
            Catch
                MsgBox("Вы неверно указали время таймера!",, "Ошибка")
                Exit Sub
            End Try
            temp_mem.interval = interval
            temp_mem.name = TextBox1.Text
            Main.timers.Add(temp_mem)
            Main.Renovate_timers_list()
            Me.Close()
        Else
            MsgBox("Таймер с таким именем уже существует!",, "Ошибка")
        End If
    End Sub
    Public Structure g_hms
        Public hour, minute, second As UShort
        Public Sub New(hour As UShort, minute As UShort, second As UShort)
            Me.hour = hour
            Me.minute = minute
            Me.second = second
        End Sub
    End Structure
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Add_redact_text.work_mode = 4
        Dim result As DialogResult = Add_redact_text.ShowDialog
        If result = DialogResult.OK Then
            temp_mem.action_data = Add_redact_text.ListBox1.SelectedItem
            Renovate_list()
        End If
    End Sub
End Class
