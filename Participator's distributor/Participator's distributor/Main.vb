Public Class Main
    Public participators As New List(Of Member)
    Public commands As New List(Of Command)
    Public tourniers As New List(Of Tournier)
    Public messanges As New List(Of Messange_temp), settings(6) As String
    Public timers As New List(Of Timer_temp)
    Dim pic_stage As Short = 0, sorted(1) As Boolean
    Private Sub Main_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove
        If pic_stage = 1 Then
            PictureBox1.BackColor = SystemColors.ControlLight
            pic_stage = 0
        End If
    End Sub
    Private Sub PictureBox1_MouseMove(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseMove
        If pic_stage = 0 Then
            PictureBox1.BackColor = Color.LightGray
            pic_stage = 1
        End If
    End Sub
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        PictureBox1.BackColor = SystemColors.ControlLight
        About.ShowDialog()
    End Sub
    Private Sub Main_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Rand.smooth_strench(0) = 0
        Rand.smooth_strench(1) = 0
        Rand.nearby(0) = True
        Rand.nearby(1) = True
        Rand.smooth(0) = True
        Rand.smooth(1) = True
        sorted(0) = False
        sorted(1) = False
        Button11.Enabled = False
        Button12.Enabled = False
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
            Dim names() As String = {"StartMessange", "InactiveMessangeTime", "MessangeSound", "TimerSoundWay", "DistributingStartMessange", "DistributingEndMessange", "SaveMessange"}, types() As Microsoft.Win32.RegistryValueKind = {Microsoft.Win32.RegistryValueKind.DWord, Microsoft.Win32.RegistryValueKind.DWord, Microsoft.Win32.RegistryValueKind.DWord, Microsoft.Win32.RegistryValueKind.String, Microsoft.Win32.RegistryValueKind.DWord, Microsoft.Win32.RegistryValueKind.DWord, Microsoft.Win32.RegistryValueKind.DWord}, values() As String = {"1", "1", "1", "", "1", "1", "1"}, reg_values() As String = reg.GetValueNames
            For i As Short = 0 To names.Length - 1
                exist_reg = False
                For Each i1 As String In reg_values
                    If names(i) = i1 Then
                        exist_reg = True
                        Exit For
                    End If
                Next
                If exist_reg = False Then
                    reg.SetValue(names(i), values(i), types(i))
                End If
            Next
            settings(0) = reg.GetValue("StartMessange")
            settings(1) = reg.GetValue("InactiveMessangeTime")
            settings(2) = reg.GetValue("MessangeSound")
            settings(3) = reg.GetValue("TimerSoundWay")
            settings(4) = reg.GetValue("DistributingStartMessange")
            settings(5) = reg.GetValue("DistributingEndMessange")
            settings(6) = reg.GetValue("SaveMessange")
            reg.Close()
            If settings(0) = "1" Then
                Mes_set.CheckBox2.Checked = True
            Else
                Mes_set.CheckBox2.Checked = False
            End If
            If settings(1) = "1" Then
                Mes_set.CheckBox1.Checked = True
            Else
                Mes_set.CheckBox1.Checked = False
            End If
            If settings(2) = "1" Then
                Mes_set.CheckBox3.Checked = True
            Else
                Mes_set.CheckBox3.Checked = False
            End If
            If settings(4) = "1" Then
                Mes_set.CheckBox4.Checked = True
            Else
                Mes_set.CheckBox4.Checked = False
            End If
            If settings(5) = "1" Then
                Mes_set.CheckBox5.Checked = True
            Else
                Mes_set.CheckBox5.Checked = False
            End If
            If settings(6) = "1" Then
                Mes_set.CheckBox6.Checked = True
            Else
                Mes_set.CheckBox6.Checked = False
            End If
        Else
            reg.CreateSubKey("Members Club Distributor")
            reg.Close()
            reg = My.Computer.Registry.CurrentUser.OpenSubKey("Software\Members Club Distributor", True)
            reg.SetValue("StartMessange", "1", Microsoft.Win32.RegistryValueKind.DWord)
            reg.SetValue("InactiveMessangeTime", "1", Microsoft.Win32.RegistryValueKind.DWord)
            reg.SetValue("MessangeSound", "1", Microsoft.Win32.RegistryValueKind.DWord)
            reg.SetValue("DistributingStartMessange", "1", Microsoft.Win32.RegistryValueKind.DWord)
            reg.SetValue("DistributingEndMessange", "1", Microsoft.Win32.RegistryValueKind.DWord)
            reg.SetValue("TimerSoundWay", "", Microsoft.Win32.RegistryValueKind.String)
            reg.SetValue("SaveMessange", "1", Microsoft.Win32.RegistryValueKind.DWord)
            settings(0) = "1"
            settings(1) = "1"
            settings(2) = "1"
            settings(3) = ""
            settings(4) = "1"
            settings(5) = "1"
            settings(6) = "1"
            reg.Close()
        End If
        Messanger.Add_link_lib("mem_dist inactual", 0, 145, 325, 210, 50)
        Messanger.Add_link_lib("com_dist inactual", 1, 145, 325, 210, 50)
        Messanger.Add_link_lib("distr_start command", 0, 520, 360, 157, 97)
        Messanger.Add_link_lib("distr_end command", 0, 20, 360, 515, 180)
        Messanger.Add_link_lib("distr_start tournier", 1, 520, 360, 157, 97)
        Messanger.Add_link_lib("distr_end tournier", 1, 20, 360, 515, 180)
        Messanger.Add_link_lib("reset com", 0, 520, 490, 157, 50)
        Messanger.Add_link_lib("reset tour", 1, 520, 490, 157, 50)
        Messanger.Add_link_lib("save com", 0, 520, 440, 157, 60)
        Messanger.Add_link_lib("save tour", 1, 520, 440, 157, 60)
        If Mes_set.CheckBox2.Checked Then
            Messanger.Add_messange("Добро пожаловать в MCD!", "", True)
        End If
        'If My.Computer.Network.IsAvailable = False Then
        '    Messanger.Add_messange("Вы не подключены к сети (Internet)!", "", True)
        'End If
    End Sub
    Public Sub Members_commands_redacted(mem_distr As Boolean, set_pos As Boolean)
        If mem_distr = True Then
            If set_pos = True Then
                If sorted(0) = True Then
                    Label2.Visible = True
                    Messanger.Add_messange("Распределение по командам не актуально!", "mem_dist inactual", False)
                End If
            Else
                sorted(0) = False
                Label2.Visible = False
                Messanger.Delete_messange(Messanger.Get_position_by_link("mem_dist inactual"))
            End If
        Else
            If set_pos = True Then
                If sorted(1) = True Then
                    Label10.Visible = True
                    Messanger.Add_messange("Распределение по турнирам не актуально!", "com_dist inactual", False)
                End If
            Else
                sorted(1) = False
                Label10.Visible = False
                Messanger.Delete_messange(Messanger.Get_position_by_link("com_dist inactual"))
            End If
        End If
    End Sub
    Public Sub UpDate_Lists(mode As Short)
        If mode = 0 Then
            ListBox1.Items.Clear()
            For Each i As Member In participators
                ListBox1.Items.Add(i.name & " - Н:" & i.dir)
            Next
        ElseIf mode = 1 Then
            ListBox6.Items.Clear()
            For Each i As Tournier In tourniers
                ListBox6.Items.Add(i.name & " - К:" & i.commands.Count & ", Н:" & i.dir)
            Next
        End If
        ListBox3.Items.Clear()
        For Each i As Command In commands
            ListBox3.Items.Add(i.name & " - К:" & i.members.Count & ", Н:" & i.dir)
        Next
        ListBox5.Items.Clear()
        For Each i As Command In commands
            ListBox5.Items.Add(i.name & " - К:" & i.members.Count & ", Н:" & i.dir)
        Next
    End Sub
    Public Sub UpDate_Sorter_List(mode As Short)
        If mode = 0 Then
            ListBox2.Items.Clear()
            For Each i As Command In commands
                If i.members.Count = 0 Then
                    ListBox2.Items.Add("Команда: " & i.name)
                Else
                    Dim result As String = "Команда: " & i.name & "; Участники: "
                    For i1 As Integer = 0 To i.members.Count - 1
                        If i1 = 0 Then
                            result &= i.members.Item(i1).name
                        Else
                            result &= ", " & i.members.Item(i1).name
                        End If
                    Next
                    ListBox2.Items.Add(result)
                End If
            Next
        ElseIf mode = 1 Then
            ListBox7.Items.Clear()
            For Each i As Tournier In tourniers
                If i.commands.Count = 0 Then
                    ListBox7.Items.Add("Турнир: " & i.name)
                Else
                    Dim result As String = "Турнир: " & i.name & "; Команда: "
                    For i1 As Integer = 0 To i.commands.Count - 1
                        If i1 = 0 Then
                            result &= i.commands.Item(i1).name
                        Else
                            result &= ", " & i.commands.Item(i1).name
                        End If
                    Next
                    ListBox7.Items.Add(result)
                End If
            Next
        End If
    End Sub
    Private Sub Panel1_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles Panel1.MouseEnter
        Panel1.Focus()
    End Sub
    Private Sub ListBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox4.SelectedIndexChanged
        If ListBox4.SelectedIndex > -1 Then
            Button11.Enabled = True
            If messanges(ListBox4.SelectedIndex).link <> "" Then
                Button12.Enabled = True
            Else
                Button12.Enabled = False
            End If
        End If
    End Sub
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Dim mess As Messanger.Link_lib_part = Messanger.Get_link_lib(messanges(ListBox4.SelectedIndex).link)
        If mess.tab_page <> -1 Then
            TabControl1.SelectedIndex = mess.tab_page
            Intro_View.Set_position(mess.x, mess.y, mess.width, mess.height)
            Intro_View.ShowDialog()
        Else
            MsgBox("Невозможно выполнить переход к объекту-отправителю этого уведомления!",, "Ошибка")
        End If
    End Sub
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        If My.Computer.Keyboard.ShiftKeyDown And button_p_press Then
            Messanger.Clear_inactive_messanger()
        ElseIf My.Computer.Keyboard.ShiftKeyDown And button_p_press = False Then
            Messanger.Clear_messanger()
        Else
            Messanger.Delete_messange(ListBox4.SelectedIndex)
        End If
    End Sub
    Private Sub cont_KeyDown(sender As Object, e As KeyEventArgs) Handles active_control.KeyDown
        If e.KeyCode = 80 And button_p_press = False Then
            button_p_press = True
        End If
    End Sub
    Private Sub cont_KeyUp(sender As Object, e As KeyEventArgs) Handles active_control.KeyUp
        If e.KeyCode = 80 And button_p_press = True Then
            button_p_press = False
        End If
    End Sub
    WithEvents active_control As Control
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles ActiveControl_Get.Tick
        active_control = Me.ActiveControl
    End Sub
    Private Sub Button11_EnabledChanged(sender As Object, e As EventArgs) Handles Button11.EnabledChanged
        If Button11.Enabled = True Then
            ActiveControl_Get.Start()
        Else
            ActiveControl_Get.Stop()
        End If
    End Sub
    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        Mes_set.ShowDialog()
    End Sub
    Dim sorting_position As Integer, stage, work_mode As Short, rand_commands As New List(Of MC_rand_memory), last_group As Integer, button_p_press As Boolean
    Private Sub Sorting_timer_Tick(sender As Object, e As EventArgs) Handles Sorting_timer.Tick
        Sorting_timer.Stop()
        Dim finded_commands(-1) As Integer, radius As Integer = 0
        If work_mode = 0 Then
            If stage = 0 Then
                While finded_commands.Length = 0
                    For i1 As Integer = 0 To commands.Count - 1
                        If Math.Max(participators.Item(sorting_position).dir, commands.Item(i1).dir) - Math.Min(participators.Item(sorting_position).dir, commands.Item(i1).dir) <= radius Then
                            Array.Resize(finded_commands, finded_commands.Length + 1)
                            finded_commands(finded_commands.Length - 1) = i1
                        End If
                    Next
                    radius += 1
                End While
                If finded_commands.Length = 1 Then
                    commands.Item(finded_commands(0)).members.Add(participators.Item(sorting_position))
                    last_group = finded_commands(0)
                ElseIf finded_commands.Length > 1 Then
                    Dim r As New Random
                    If Rand.smooth(0) = True Then
                        Dim min_value As Integer = commands.Item(0).members.Count, select_commands(-1) As Integer
                        For i As Integer = 1 To finded_commands.Length - 1
                            If commands.Item(finded_commands(i)).members.Count < min_value Then
                                min_value = commands.Item(finded_commands(i)).members.Count
                            End If
                        Next
                        For i As Integer = 0 To finded_commands.Length - 1
                            If commands.Item(finded_commands(i)).members.Count <= min_value + Rand.smooth_strench(0) Then
                                Array.Resize(select_commands, select_commands.Length + 1)
                                select_commands(select_commands.Length - 1) = finded_commands(i)
                            End If
                        Next
                        If Rand.nearby(0) = False And sorting_position > 0 Then
                            Dim group As Integer = select_commands(r.Next(select_commands.Length))
                            While group = last_group
                                group = select_commands(r.Next(select_commands.Length))
                            End While
                            commands.Item(group).members.Add(participators.Item(sorting_position))
                            rand_commands.Add(New MC_rand_memory(sorting_position, group))
                            last_group = group
                        Else
                            Dim group As Integer = select_commands(r.Next(select_commands.Length))
                            commands.Item(group).members.Add(participators.Item(sorting_position))
                            rand_commands.Add(New MC_rand_memory(sorting_position, group))
                            last_group = group
                        End If
                    Else
                        If Rand.nearby(0) = False And sorting_position > 0 Then
                            Dim group As Integer = finded_commands(r.Next(finded_commands.Length))
                            While group = last_group
                                group = finded_commands(r.Next(finded_commands.Length))
                            End While
                            commands.Item(group).members.Add(participators.Item(sorting_position))
                            rand_commands.Add(New MC_rand_memory(sorting_position, group))
                            last_group = group
                        Else
                            Dim group As Integer = finded_commands(r.Next(finded_commands.Length))
                            commands.Item(group).members.Add(participators.Item(sorting_position))
                            last_group = group
                        End If
                    End If
                End If
                sorting_position += 1
                If sorting_position = participators.Count Then
                    If Rand.smooth(0) = True Then
                        sorting_position = 0
                        stage = 1
                        Sorting_timer.Start()
                    Else
                        UpDate_Sorter_List(0)
                        UpDate_Lists(0)
                        For Each i As Control In Me.Controls
                            If i.Tag = "ToDisable" Then
                                i.Enabled = True
                            End If
                        Next
                        Members_commands_redacted(True, False)
                        sorted(0) = True
                        Progress.Close()
                        Me.UseWaitCursor = False
                        If Mes_set.CheckBox5.Checked Then
                            Messanger.Add_messange("Распределение по командам завершено!", "distr_end command", True)
                        End If
                    End If
                Else
                    Sorting_timer.Start()
                End If
            ElseIf stage = 1 Then
                Dim this_rand_command As MC_rand_memory, com_finded As Boolean = False
                For i1 As Integer = 0 To rand_commands.Count - 1
                    If rand_commands(i1).member = sorting_position Then
                        this_rand_command = rand_commands(i1)
                        com_finded = True
                        Exit For
                    End If
                Next
                If com_finded = True Then
                    While finded_commands.Length = 0
                        For i1 As Integer = 0 To commands.Count - 1
                            If Math.Max(participators.Item(sorting_position).dir, commands.Item(i1).dir) - Math.Min(participators.Item(sorting_position).dir, commands.Item(i1).dir) <= radius Then
                                Array.Resize(finded_commands, finded_commands.Length + 1)
                                finded_commands(finded_commands.Length - 1) = i1
                            End If
                        Next
                        radius += 1
                    End While
                    Dim min_value As Integer = commands.Item(finded_commands(0)).members.Count, select_commands(-1) As Integer, command_exits As Boolean = False
                    For i1 As Integer = 0 To finded_commands.Length - 1
                        If finded_commands(i1) = this_rand_command.com Then
                            If commands.Item(finded_commands(i1)).members.Count - 1 < min_value Then
                                min_value = commands.Item(finded_commands(i1)).members.Count - 1
                            End If
                        Else
                            If commands.Item(finded_commands(i1)).members.Count < min_value Then
                                min_value = commands.Item(finded_commands(i1)).members.Count
                            End If
                        End If
                    Next
                    For i1 As Integer = 0 To finded_commands.Length - 1
                        If finded_commands(i1) = this_rand_command.com Then
                            If commands.Item(finded_commands(i1)).members.Count - 1 <= min_value + Rand.smooth_strench(0) Then
                                command_exits = True
                            End If
                        Else
                            If commands.Item(finded_commands(i1)).members.Count <= min_value + Rand.smooth_strench(0) Then
                                Array.Resize(select_commands, select_commands.Length + 1)
                                select_commands(select_commands.Length - 1) = finded_commands(i1)
                            End If
                        End If
                    Next
                    If command_exits = False Then
                        If finded_commands.Length = 1 Then
                            commands.Item(this_rand_command.com).members.Remove(participators(this_rand_command.member))
                            commands.Item(0).members.Add(participators.Item(sorting_position))
                        ElseIf finded_commands.Length > 1 Then
                            commands.Item(this_rand_command.com).members.Remove(participators(this_rand_command.member))
                            Dim r As New Random
                            If Rand.nearby(0) = False And sorting_position > 0 Then
                                Dim group As Integer = select_commands(r.Next(select_commands.Length))
                                While group = last_group
                                    group = select_commands(r.Next(select_commands.Length))
                                End While
                                commands.Item(group).members.Add(participators.Item(sorting_position))
                                last_group = group
                            Else
                                Dim group As Integer = select_commands(r.Next(select_commands.Length))
                                commands.Item(group).members.Add(participators.Item(sorting_position))
                                last_group = group
                            End If
                        End If
                    End If
                End If
                sorting_position += 1
                If sorting_position = participators.Count Then
                    UpDate_Sorter_List(0)
                    UpDate_Lists(0)
                    For Each i As Control In Me.Controls
                        If i.Tag = "ToDisable" Then
                            i.Enabled = True
                        End If
                    Next
                    Members_commands_redacted(True, False)
                    sorted(0) = True
                    Progress.Close()
                    Me.UseWaitCursor = False
                    If Mes_set.CheckBox5.Checked Then
                        Messanger.Add_messange("Распределение по командам завершено!", "distr_end command", True)
                    End If
                Else
                    Sorting_timer.Start()
                End If
            End If
        ElseIf work_mode = 1 Then
            If stage = 0 Then
                While finded_commands.Length = 0
                    For i1 As Integer = 0 To tourniers.Count - 1
                        If Math.Max(commands.Item(sorting_position).dir, tourniers.Item(i1).dir) - Math.Min(commands.Item(sorting_position).dir, tourniers.Item(i1).dir) <= radius Then
                            Array.Resize(finded_commands, finded_commands.Length + 1)
                            finded_commands(finded_commands.Length - 1) = i1
                        End If
                    Next
                    radius += 1
                End While
                If finded_commands.Length = 1 Then
                    tourniers.Item(finded_commands(0)).commands.Add(commands.Item(sorting_position))
                    last_group = finded_commands(0)
                ElseIf finded_commands.Length > 1 Then
                    Dim r As New Random
                    If Rand.smooth(1) = True Then
                        Dim min_value As Integer = tourniers.Item(0).commands.Count, select_tourniers(-1) As Integer
                        For i As Integer = 1 To finded_commands.Length - 1
                            If tourniers.Item(finded_commands(i)).commands.Count < min_value Then
                                min_value = tourniers.Item(finded_commands(i)).commands.Count
                            End If
                        Next
                        For i As Integer = 0 To finded_commands.Length - 1
                            If tourniers.Item(finded_commands(i)).commands.Count <= min_value + Rand.smooth_strench(1) Then
                                Array.Resize(select_tourniers, select_tourniers.Length + 1)
                                select_tourniers(select_tourniers.Length - 1) = finded_commands(i)
                            End If
                        Next
                        If Rand.nearby(1) = False And sorting_position > 0 Then
                            Dim group As Integer = select_tourniers(r.Next(select_tourniers.Length))
                            While group = last_group
                                group = select_tourniers(r.Next(select_tourniers.Length))
                            End While
                            tourniers.Item(group).commands.Add(commands.Item(sorting_position))
                            rand_commands.Add(New MC_rand_memory(sorting_position, group))
                            last_group = group
                        Else
                            Dim group As Integer = select_tourniers(r.Next(select_tourniers.Length))
                            tourniers.Item(group).commands.Add(commands.Item(sorting_position))
                            rand_commands.Add(New MC_rand_memory(sorting_position, group))
                            last_group = group
                        End If
                    Else
                        If Rand.nearby(1) = False And sorting_position > 0 Then
                            Dim group As Integer = finded_commands(r.Next(finded_commands.Length))
                            While group = last_group
                                group = finded_commands(r.Next(finded_commands.Length))
                            End While
                            tourniers.Item(group).commands.Add(commands.Item(sorting_position))
                            rand_commands.Add(New MC_rand_memory(sorting_position, group))
                            last_group = group
                        Else
                            Dim group As Integer = finded_commands(r.Next(finded_commands.Length))
                            tourniers.Item(group).commands.Add(commands.Item(sorting_position))
                            last_group = group
                        End If
                    End If
                End If
                sorting_position += 1
                If sorting_position = commands.Count Then
                    If Rand.smooth(1) = True Then
                        sorting_position = 0
                        stage = 1
                        Sorting_timer.Start()
                    Else
                        UpDate_Sorter_List(1)
                        UpDate_Lists(1)
                        For Each i As Control In Me.Controls
                            If i.Tag = "ToDisable" Then
                                i.Enabled = True
                            End If
                        Next
                        Members_commands_redacted(False, False)
                        sorted(1) = True
                        Progress.Close()
                        Me.UseWaitCursor = False
                        If Mes_set.CheckBox5.Checked Then
                            Messanger.Add_messange("Распределение по турнирам завершено!", "distr_end tournier", True)
                        End If
                    End If
                Else
                    Sorting_timer.Start()
                End If
            ElseIf stage = 1 Then
                Dim this_rand_command As MC_rand_memory, com_finded As Boolean = False
                For i1 As Integer = 0 To rand_commands.Count - 1
                    If rand_commands(i1).member = sorting_position Then
                        this_rand_command = rand_commands(i1)
                        com_finded = True
                        Exit For
                    End If
                Next
                If com_finded = True Then
                    While finded_commands.Length = 0
                        For i1 As Integer = 0 To tourniers.Count - 1
                            If Math.Max(commands.Item(sorting_position).dir, tourniers.Item(i1).dir) - Math.Min(commands.Item(sorting_position).dir, tourniers.Item(i1).dir) <= radius Then
                                Array.Resize(finded_commands, finded_commands.Length + 1)
                                finded_commands(finded_commands.Length - 1) = i1
                            End If
                        Next
                        radius += 1
                    End While
                    Dim min_value As Integer = tourniers.Item(finded_commands(0)).commands.Count, select_tourniers(-1) As Integer, command_exits As Boolean = False
                    For i1 As Integer = 0 To finded_commands.Length - 1
                        If finded_commands(i1) = this_rand_command.com Then
                            If tourniers.Item(finded_commands(i1)).commands.Count - 1 < min_value Then
                                min_value = tourniers.Item(finded_commands(i1)).commands.Count - 1
                            End If
                        Else
                            If tourniers.Item(finded_commands(i1)).commands.Count < min_value Then
                                min_value = tourniers.Item(finded_commands(i1)).commands.Count
                            End If
                        End If
                    Next
                    For i1 As Integer = 0 To finded_commands.Length - 1
                        If finded_commands(i1) = this_rand_command.com Then
                            If tourniers.Item(finded_commands(i1)).commands.Count - 1 <= min_value + Rand.smooth_strench(1) Then
                                command_exits = True
                            End If
                        Else
                            If tourniers.Item(finded_commands(i1)).commands.Count <= min_value + Rand.smooth_strench(1) Then
                                Array.Resize(select_tourniers, select_tourniers.Length + 1)
                                select_tourniers(select_tourniers.Length - 1) = finded_commands(i1)
                            End If
                        End If
                    Next
                    If command_exits = False Then
                        If finded_commands.Length = 1 Then
                            tourniers.Item(this_rand_command.com).commands.Remove(commands(this_rand_command.member))
                            tourniers.Item(0).commands.Add(commands.Item(sorting_position))
                        ElseIf finded_commands.Length > 1 Then
                            tourniers.Item(this_rand_command.com).commands.Remove(commands(this_rand_command.member))
                            Dim r As New Random
                            If Rand.nearby(1) = False And sorting_position > 0 Then
                                Dim group As Integer = select_tourniers(r.Next(select_tourniers.Length))
                                While group = last_group
                                    group = select_tourniers(r.Next(select_tourniers.Length))
                                End While
                                tourniers.Item(group).commands.Add(commands.Item(sorting_position))
                                last_group = group
                            Else
                                Dim group As Integer = select_tourniers(r.Next(select_tourniers.Length))
                                tourniers.Item(group).commands.Add(commands.Item(sorting_position))
                                last_group = group
                            End If
                        End If
                    End If
                End If
                sorting_position += 1
                If sorting_position = commands.Count Then
                    UpDate_Sorter_List(1)
                    UpDate_Lists(1)
                    For Each i As Control In Me.Controls
                        If i.Tag = "ToDisable" Then
                            i.Enabled = True
                        End If
                    Next
                    Members_commands_redacted(False, False)
                    sorted(1) = True
                    Progress.Close()
                    Me.UseWaitCursor = False
                    If Mes_set.CheckBox5.Checked Then
                        Messanger.Add_messange("Распределение по турнирам завершено!", "distr_end tournier", True)
                    End If
                Else
                    Sorting_timer.Start()
                End If
            End If
        End If
        If Progress.progress_mode(0) = 1 Then
            If Math.IEEERemainder(sorting_position, Progress.progress_mode(1)) = 0 Then
                Progress.Progress_tick()
            End If
        ElseIf Progress.progress_mode(0) = 0 Then
            Progress.Progress_tick()
        End If
        Try
            If ActiveForm.Name = "Main" Then
                Progress.Activate()
            End If
        Catch
        End Try
    End Sub
#Region "Timer"
    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        Add_redact_timer.work_mode = 1
        Add_redact_timer.ShowDialog()
    End Sub
    Public Sub Renovate_timers_list()
        ListBox8.Items.Clear()
        ListBox9.Items.Clear()
        For Each i As Timer_temp In timers
            Dim result As String = "Имя: " & i.name & "; "
            If i.active Then
                result &= "активен; "
            Else
                result &= "не активен; "
            End If
            Dim hms As Add_redact_timer.g_hms = Add_redact_timer.Get_hms(i.interval)
            result &= "интервал: " & hms.hour & " ч. " & hms.minute & " м. " & hms.second & " с. ; "
            If i.connect_type = 0 Then
                result &= "связь: нет; "
            Else
                If i.connect_type = 1 Then
                    result &= "связь: да (участники); "
                ElseIf i.connect_type = 2 Then
                    result &= "связь: да (команды); "
                ElseIf i.connect_type = 3 Then
                    result &= "связь: да (турниры); "
                End If
            End If
            If i.action = 0 Then
                result &= "триггер: нет; "
            Else
                If i.action = Timer_temp.Trigger_action.Delete_this_select Then
                    result &= "триггер: удалить связанное;"
                ElseIf i.action = Timer_temp.Trigger_action.Repeat_timer Then
                    result &= "триггер: перезапустить таймер;"
                ElseIf i.action = Timer_temp.Trigger_action.Sort_commands Then
                    result &= "триггер: перераспределить команды;"
                ElseIf i.action = Timer_temp.Trigger_action.Sort_members Then
                    result &= "триггер: перераспределить участников;"
                ElseIf i.action = Timer_temp.Trigger_action.Start_over_timer Then
                    result &= "триггер: запустить другой таймер (" & i.action_data & ");"
                End If
            End If
            ListBox8.Items.Add(result)
        Next
    End Sub
    Private Sub Button29_Click(sender As Object, e As EventArgs) Handles Button29.Click
        If Button29.Text = "\/" Then
            TableLayoutPanel1.Width = 795
            Button29.Text = "/\"
        Else
            TableLayoutPanel1.Width = 632
            Button29.Text = "\/"
        End If
    End Sub
#End Region

#Region "Comands distributor"
    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        If My.Computer.Keyboard.ShiftKeyDown = False Then
            If TextBox2.Text <> "" Then
                Dim exits As Boolean = False
                For Each i As Command In commands
                    If i.name = TextBox2.Text Then
                        exits = True
                    End If
                Next
                If exits = False Then
                    commands.Add(New Command)
                    commands.Item(commands.Count - 1).name = TextBox2.Text
                    commands.Item(commands.Count - 1).dir = 0
                    TextBox2.Text = ""
                    UpDate_Lists(1)
                    Members_commands_redacted(True, True)
                Else
                    MsgBox("Команда с таким названием уже добавлена в список!",, "Ошибка")
                End If
            Else
                MsgBox("Вы не ввели название команды!",, "Ошибка")
            End If
        ElseIf My.Computer.Keyboard.ShiftKeyDown = True Then
            If TextBox2.Text <> "" Then
                If TextBox2.Text.IndexOf("/") > -1 Then
                    Dim par_name As String = "0"
                    Try
                        If Mid(TextBox2.Text, TextBox2.Text.LastIndexOf("/") + 2).Length > 4 Then Error 1
                        Dim r_number As UShort = Mid(TextBox2.Text, TextBox2.Text.LastIndexOf("/") + 2)
                        par_name = Mid(TextBox2.Text, 1, TextBox2.Text.LastIndexOf("/"))
                        If par_name = "" Then
                            MsgBox("Вы не ввели название команды!",, "Ошибка")
                            Error 1
                        End If
                        Dim exits As Boolean = False
                        For Each i As Command In commands
                            If i.name = par_name Then
                                exits = True
                            End If
                        Next
                        If exits = False Then
                            commands.Add(New Command)
                            commands.Item(commands.Count - 1).name = par_name
                            commands.Item(commands.Count - 1).dir = r_number
                            TextBox2.Text = ""
                            UpDate_Lists(1)
                            Members_commands_redacted(True, True)
                        Else
                            MsgBox("Команда с таким названием уже добавлена в список!",, "Ошибка")
                        End If
                    Catch
                        If par_name <> "" Then
                            MsgBox("Неправильно задан распределительный номер! Распределительный номер - 4-х значное число.",, "Ошибка")
                        End If
                    End Try
                Else
                    MsgBox("Вы не задали распределительный номер! Формат: имя/распределительный номер",, "Ошибка")
                End If
            Else
                MsgBox("Поле ввода пустое!",, "Ошибка")
            End If
        End If
    End Sub
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = 13 And e.Shift = False Then
            If TextBox2.Text <> "" Then
                Dim exits As Boolean = False
                For Each i As Command In commands
                    If i.name = TextBox2.Text Then
                        exits = True
                    End If
                Next
                If exits = False Then
                    commands.Add(New Command)
                    commands.Item(commands.Count - 1).name = TextBox2.Text
                    commands.Item(commands.Count - 1).dir = 0
                    TextBox2.Text = ""
                    UpDate_Lists(1)
                    Members_commands_redacted(True, True)
                Else
                    MsgBox("Команда с таким названием уже добавлена в список!",, "Ошибка")
                End If
            Else
                MsgBox("Вы не ввели название команды!",, "Ошибка")
            End If
        ElseIf e.KeyCode = 13 And e.Shift = True Then
            If TextBox2.Text <> "" Then
                If TextBox2.Text.IndexOf("/") > -1 Then
                    Dim par_name As String = "0"
                    Try
                        If Mid(TextBox2.Text, TextBox2.Text.LastIndexOf("/") + 2).Length > 4 Then Error 1
                        Dim r_number As UShort = Mid(TextBox2.Text, TextBox2.Text.LastIndexOf("/") + 2)
                        par_name = Mid(TextBox2.Text, 1, TextBox2.Text.LastIndexOf("/"))
                        If par_name = "" Then
                            MsgBox("Вы не ввели название команды!",, "Ошибка")
                            Error 1
                        End If
                        Dim exits As Boolean = False
                        For Each i As Command In commands
                            If i.name = par_name Then
                                exits = True
                            End If
                        Next
                        If exits = False Then
                            commands.Add(New Command)
                            commands.Item(commands.Count - 1).name = par_name
                            commands.Item(commands.Count - 1).dir = r_number
                            TextBox2.Text = ""
                            UpDate_Lists(1)
                            Members_commands_redacted(True, True)
                        Else
                            MsgBox("Команда с таким названием уже добавлена в список!",, "Ошибка")
                        End If
                    Catch
                        If par_name <> "" Then
                            MsgBox("Неправильно задан распределительный номер! Распределительный номер - 4-х значное число.",, "Ошибка")
                        End If
                    End Try
                Else
                    MsgBox("Вы не задали распределительный номер! Формат: имя/распределительный номер",, "Ошибка")
                End If
            Else
                MsgBox("Поле ввода пустое!",, "Ошибка")
            End If
        End If
    End Sub
    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        If My.Computer.Keyboard.ShiftKeyDown = False Then
            If TextBox4.Text <> "" Then
                Dim exits As Boolean = False
                For Each i As Tournier In tourniers
                    If i.name = TextBox4.Text Then
                        exits = True
                    End If
                Next
                If exits = False Then
                    tourniers.Add(New Tournier)
                    tourniers.Item(tourniers.Count - 1).name = TextBox4.Text
                    tourniers.Item(tourniers.Count - 1).dir = 0
                    TextBox4.Text = ""
                    UpDate_Lists(1)
                    Members_commands_redacted(False, True)
                Else
                    MsgBox("Турнир с таким названием уже добавлен в список!",, "Ошибка")
                End If
            Else
                MsgBox("Вы не ввели название турнира!",, "Ошибка")
            End If
        ElseIf My.Computer.Keyboard.ShiftKeyDown = True Then
            If TextBox4.Text <> "" Then
                If TextBox4.Text.IndexOf("/") > -1 Then
                    Dim par_name As String = "0"
                    Try
                        If Mid(TextBox4.Text, TextBox4.Text.LastIndexOf("/") + 2).Length > 4 Then Error 1
                        Dim r_number As UShort = Mid(TextBox4.Text, TextBox4.Text.LastIndexOf("/") + 2)
                        par_name = Mid(TextBox4.Text, 1, TextBox4.Text.LastIndexOf("/"))
                        If par_name = "" Then
                            MsgBox("Вы не ввели название турнира!",, "Ошибка")
                            Error 1
                        End If
                        Dim exits As Boolean = False
                        For Each i As Tournier In tourniers
                            If i.name = par_name Then
                                exits = True
                            End If
                        Next
                        If exits = False Then
                            tourniers.Add(New Tournier)
                            tourniers.Item(tourniers.Count - 1).name = par_name
                            tourniers.Item(tourniers.Count - 1).dir = r_number
                            TextBox4.Text = ""
                            UpDate_Lists(1)
                            Members_commands_redacted(False, True)
                        Else
                            MsgBox("Турнир с таким названием уже добавлен в список!",, "Ошибка")
                        End If
                    Catch
                        If par_name <> "" Then
                            MsgBox("Неправильно задан распределительный номер! Распределительный номер - 4-х значное число.",, "Ошибка")
                        End If
                    End Try
                Else
                    MsgBox("Вы не задали распределительный номер! Формат: имя/распределительный номер",, "Ошибка")
                End If
            Else
                MsgBox("Поле ввода пустое!",, "Ошибка")
            End If
        End If
    End Sub
    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = 13 And e.Shift = False Then
            If TextBox4.Text <> "" Then
                Dim exits As Boolean = False
                For Each i As Tournier In tourniers
                    If i.name = TextBox4.Text Then
                        exits = True
                    End If
                Next
                If exits = False Then
                    tourniers.Add(New Tournier)
                    tourniers.Item(tourniers.Count - 1).name = TextBox4.Text
                    tourniers.Item(tourniers.Count - 1).dir = 0
                    TextBox4.Text = ""
                    UpDate_Lists(1)
                    Members_commands_redacted(False, True)
                Else
                    MsgBox("Турнир с таким названием уже добавлен в список!",, "Ошибка")
                End If
            Else
                MsgBox("Вы не ввели название турнира!",, "Ошибка")
            End If
        ElseIf e.KeyCode = 13 And e.Shift = True Then
            If TextBox4.Text <> "" Then
                If TextBox4.Text.IndexOf("/") > -1 Then
                    Dim par_name As String = "0"
                    Try
                        If Mid(TextBox4.Text, TextBox4.Text.LastIndexOf("/") + 2).Length > 4 Then Error 1
                        Dim r_number As UShort = Mid(TextBox4.Text, TextBox4.Text.LastIndexOf("/") + 2)
                        par_name = Mid(TextBox4.Text, 1, TextBox4.Text.LastIndexOf("/"))
                        If par_name = "" Then
                            MsgBox("Вы не ввели название турнира!",, "Ошибка")
                            Error 1
                        End If
                        Dim exits As Boolean = False
                        For Each i As Tournier In tourniers
                            If i.name = par_name Then
                                exits = True
                            End If
                        Next
                        If exits = False Then
                            tourniers.Add(New Tournier)
                            tourniers.Item(tourniers.Count - 1).name = par_name
                            tourniers.Item(tourniers.Count - 1).dir = r_number
                            TextBox4.Text = ""
                            UpDate_Lists(1)
                            Members_commands_redacted(False, True)
                        Else
                            MsgBox("Турнир с таким названием уже добавлен в список!",, "Ошибка")
                        End If
                    Catch
                        If par_name <> "" Then
                            MsgBox("Неправильно задан распределительный номер! Распределительный номер - 4-х значное число.",, "Ошибка")
                        End If
                    End Try
                Else
                    MsgBox("Вы не задали распределительный номер! Формат: имя/распределительный номер",, "Ошибка")
                End If
            Else
                MsgBox("Поле ввода пустое!",, "Ошибка")
            End If
        End If
    End Sub
    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        If ListBox5.SelectedIndex = -1 Then
            MsgBox("Вы не выбрали элемент списка!",, "Ошибка")
        Else
            commands.RemoveAt(ListBox5.SelectedIndex)
            ListBox5.Items.RemoveAt(ListBox5.SelectedIndex)
            Members_commands_redacted(True, True)
            Members_commands_redacted(False, True)
        End If
    End Sub
    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        If ListBox5.SelectedIndex = -1 Then
            MsgBox("Вы не выбрали элемент списка!",, "Ошибка")
        Else
            Redact.work_mode = 1
            Redact.ShowDialog()
        End If
    End Sub
    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        If ListBox6.SelectedIndex = -1 Then
            MsgBox("Вы не выбрали элемент списка!",, "Ошибка")
        Else
            tourniers.RemoveAt(ListBox6.SelectedIndex)
            ListBox6.Items.RemoveAt(ListBox6.SelectedIndex)
            Members_commands_redacted(False, True)
        End If
    End Sub
    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        If ListBox6.SelectedIndex = -1 Then
            MsgBox("Вы не выбрали элемент списка!",, "Ошибка")
        Else
            Redact.work_mode = 2
            Redact.ShowDialog()
        End If
    End Sub
    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Rand.work_mode = 1
        Rand.ShowDialog()
    End Sub
    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        If commands.Count = 0 Then
            MsgBox("Отсутствуют команды!",, "Ошибка")
            Exit Sub
        End If
        If tourniers.Count = 0 Then
            MsgBox("Отсутствуют турниры!",, "Ошибка")
            Exit Sub
        End If
        If tourniers.Count > commands.Count Then
            Dim result As MsgBoxResult = MsgBox("Турниров больше, чем команд!" & Chr(10) & "Вы хотите продолжить?", MsgBoxStyle.YesNo, "Предупреждение")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
        End If
        If Math.IEEERemainder(commands.Count, tourniers.Count) <> 0 And Rand.smooth(1) = True Then
            Dim result As MsgBoxResult = MsgBox("Команды не делятся поровну на турниры!" & Chr(10) & "Вы хотите продолжить?", MsgBoxStyle.YesNo, "Предупреждение")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
        End If
        If tourniers.Count = 1 And Rand.nearby(1) = False Then
            Dim result As MsgBoxResult = MsgBox("Настройка 'Допускать соседние элементы рядом' отключена! В списке присутствует только 1 турнир. При продолжении, настройка 'Допускать соседние элементы рядом' будет включена." & Chr(10) & "Вы хотите продолжить?", MsgBoxStyle.YesNo, "Предупреждение")
            If result = MsgBoxResult.No Then
                Exit Sub
            Else
                Rand.nearby(1) = True
            End If
        End If
        Me.UseWaitCursor = True
        For Each i As Tournier In tourniers
            i.commands.Clear()
        Next
        rand_commands.Clear()
        sorting_position = 0
        stage = 0
        work_mode = 1
        For Each i As Control In Me.Controls
            If i.Tag = "ToDisable" Then
                i.Enabled = False
            End If
        Next
        Progress.Reset()
        If Rand.smooth(1) = True Then
            If commands.Count * 2 >= 100 Then
                Progress.progress_mode(0) = 1
                Progress.progress_mode(1) = Math.Round(commands.Count * 2 / 100)
            Else
                Progress.progress_mode(0) = 0
                Progress.progress_mode(1) = Math.Round(100 / commands.Count)
            End If
        Else
            If commands.Count >= 100 Then
                Progress.progress_mode(0) = 1
                Progress.progress_mode(1) = Math.Round(commands.Count / 100)
            Else
                Progress.progress_mode(0) = 0
                Progress.progress_mode(1) = Math.Round(100 / commands.Count)
            End If
        End If
        Progress.Show()
        Sorting_timer.Start()
        If Mes_set.CheckBox4.Checked Then
            Messanger.Add_messange("Распределение по турнирам начато.", "distr_start tournier", True)
        End If
    End Sub
    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        If sorted(1) = True Then
            SaveFormat.work_mode = 1
            Dim result As DialogResult = SaveFormat.ShowDialog
            If result = DialogResult.OK Then
                Dim arr(tourniers.Count - 1) As String
                For i As Integer = 0 To tourniers.Count - 1
                    Dim res As String = ""
                    If tourniers(i).commands.Count = 0 Then
                        If SaveFormat.CheckBox1.Checked And SaveFormat.CheckBox3.Checked Then
                            res = "Турнир: " & tourniers(i).name & " - Н:" & tourniers(i).dir & ", К:0"
                        ElseIf SaveFormat.CheckBox1.Checked Then
                            res = "Турнир: " & tourniers(i).name & " - Н:" & tourniers(i).dir
                        ElseIf SaveFormat.CheckBox3.Checked Then
                            res = "Турнир: " & tourniers(i).name & " - К:0"
                        Else
                            res = "Турнир: " & tourniers(i).name
                        End If
                        arr(i) = res
                    Else
                        If SaveFormat.CheckBox1.Checked And SaveFormat.CheckBox3.Checked Then
                            res = "Турнир: " & tourniers(i).name & " - Н:" & tourniers(i).dir & ", К:" & tourniers(i).commands.Count & "; Команды: "
                        ElseIf SaveFormat.CheckBox1.Checked Then
                            res = "Турнир: " & tourniers(i).name & " - Н:" & tourniers(i).dir & "; Команды: "
                        ElseIf SaveFormat.CheckBox3.Checked Then
                            res = "Турнир: " & tourniers(i).name & " - К:" & tourniers(i).commands.Count & "; Команды: "
                        Else
                            res = "Турнир: " & tourniers(i).name & "; Команды: "
                        End If
                        If SaveFormat.CheckBox2.Checked Then
                            For i1 As Integer = 0 To tourniers(i).commands.Count - 1
                                If i1 = 0 Then
                                    res &= tourniers(i).commands.Item(i1).name & " - Н:" & tourniers(i).commands.Item(i1).dir
                                Else
                                    res &= ", " & tourniers(i).commands.Item(i1).name & " - Н:" & tourniers(i).commands.Item(i1).dir
                                End If
                            Next
                        Else
                            For i1 As Integer = 0 To tourniers(i).commands.Count - 1
                                If i1 = 0 Then
                                    res &= tourniers(i).commands.Item(i1).name
                                Else
                                    res &= ", " & tourniers(i).commands.Item(i1).name
                                End If
                            Next
                        End If
                        arr(i) = res
                    End If
                Next
                Try
                    IO.File.WriteAllLines(SaveFormat.SaveFile.FileName, arr)
                    If Mes_set.CheckBox6.Checked Then
                        Messanger.Add_messange("Распределение по турнирам сохранено успешно!", "save tour", True)
                    End If
                Catch
                    If Mes_set.CheckBox6.Checked Then
                        Messanger.Add_messange("Распределение по турнирам не сохранено!", "save tour", True)
                    End If
                    MsgBox("Нет доступа к файлу или(и) папке!",, "Ошибка")
                End Try
            End If
        Else
            MsgBox("Команды не распределены по турнирам или текущее распределение не активно!",, "Ошибка")
        End If
    End Sub
    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        commands.Clear()
        tourniers.Clear()
        UpDate_Lists(1)
        TextBox2.Text = ""
        TextBox4.Text = ""
        ListBox7.Items.Clear()
        Rand.smooth(0) = True
        Rand.nearby(0) = True
        Rand.smooth_strench(0) = 0
        sorted(1) = False
        Members_commands_redacted(False, False)
        Members_commands_redacted(True, True)
        Messanger.Add_messange("Выполнен сброс раздела ""Турниры""!", "reset tour", True)
    End Sub
    Private Sub Label10_MouseDown(sender As Object, e As MouseEventArgs) Handles Label10.MouseDown
        If My.Computer.Keyboard.ShiftKeyDown = True Then
            Members_commands_redacted(False, False)
        End If
    End Sub
#End Region

#Region "Member distributor"
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If My.Computer.Keyboard.ShiftKeyDown = False Then
            If TextBox1.Text <> "" Then
                Dim exits As Boolean = False
                For Each i As Member In participators
                    If i.name = TextBox1.Text Then
                        exits = True
                    End If
                Next
                If exits = False Then
                    participators.Add(New Member)
                    participators.Item(participators.Count - 1).name = TextBox1.Text
                    participators.Item(participators.Count - 1).dir = 0
                    TextBox1.Text = ""
                    UpDate_Lists(0)
                    Members_commands_redacted(True, True)
                Else
                    MsgBox("Участник с таким именем уже добавлен в список!",, "Ошибка")
                End If
            Else
                MsgBox("Вы не ввели имя участника!",, "Ошибка")
            End If
        ElseIf My.Computer.Keyboard.ShiftKeyDown = True Then
            If TextBox1.Text <> "" Then
                If TextBox1.Text.IndexOf("/") > -1 Then
                    Dim par_name As String = "0"
                    Try
                        If Mid(TextBox1.Text, TextBox1.Text.LastIndexOf("/") + 2).Length > 4 Then Error 1
                        Dim r_number As UShort = Mid(TextBox1.Text, TextBox1.Text.LastIndexOf("/") + 2)
                        par_name = Mid(TextBox1.Text, 1, TextBox1.Text.LastIndexOf("/"))
                        If par_name = "" Then
                            MsgBox("Вы не ввели имя участника!",, "Ошибка")
                            Error 1
                        End If
                        Dim exits As Boolean = False
                        For Each i As Member In participators
                            If i.name = par_name Then
                                exits = True
                            End If
                        Next
                        If exits = False Then
                            participators.Add(New Member)
                            participators.Item(participators.Count - 1).name = par_name
                            participators.Item(participators.Count - 1).dir = r_number
                            TextBox1.Text = ""
                            UpDate_Lists(0)
                            Members_commands_redacted(True, True)
                        Else
                            MsgBox("Участник с таким именем уже добавлен в список!",, "Ошибка")
                        End If
                    Catch
                        If par_name <> "" Then
                            MsgBox("Неправильно задан распределительный номер! Распределительный номер - 4-х значное число.",, "Ошибка")
                        End If
                    End Try
                Else
                    MsgBox("Вы не задали распределительный номер! Формат: имя/распределительный номер",, "Ошибка")
                End If
            Else
                MsgBox("Поле ввода пустое!",, "Ошибка")
            End If
        End If
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If My.Computer.Keyboard.ShiftKeyDown = False Then
            If TextBox3.Text <> "" Then
                Dim exits As Boolean = False
                For Each i As Command In commands
                    If i.name = TextBox3.Text Then
                        exits = True
                    End If
                Next
                If exits = False Then
                    commands.Add(New Command)
                    commands.Item(commands.Count - 1).name = TextBox3.Text
                    commands.Item(commands.Count - 1).dir = 0
                    TextBox3.Text = ""
                    UpDate_Lists(0)
                    Members_commands_redacted(True, True)
                    Members_commands_redacted(False, True)
                Else
                    MsgBox("Команда с таким названием уже добавлена в список!",, "Ошибка")
                End If
            Else
                MsgBox("Вы не ввели название команды!",, "Ошибка")
            End If
        ElseIf My.Computer.Keyboard.ShiftKeyDown = True Then
            If TextBox3.Text <> "" Then
                If TextBox3.Text.IndexOf("/") > -1 Then
                    Dim par_name As String = "0"
                    Try
                        If Mid(TextBox3.Text, TextBox3.Text.LastIndexOf("/") + 2).Length > 4 Then Error 1
                        Dim r_number As UShort = Mid(TextBox3.Text, TextBox3.Text.LastIndexOf("/") + 2)
                        par_name = Mid(TextBox3.Text, 1, TextBox3.Text.LastIndexOf("/"))
                        If par_name = "" Then
                            MsgBox("Вы не ввели название команды!",, "Ошибка")
                            Error 1
                        End If
                        Dim exits As Boolean = False
                        For Each i As Command In commands
                            If i.name = par_name Then
                                exits = True
                            End If
                        Next
                        If exits = False Then
                            commands.Add(New Command)
                            commands.Item(commands.Count - 1).name = par_name
                            commands.Item(commands.Count - 1).dir = r_number
                            TextBox3.Text = ""
                            UpDate_Lists(0)
                            Members_commands_redacted(True, True)
                            Members_commands_redacted(False, True)
                        Else
                            MsgBox("Команда с таким названием уже добавлена в список!",, "Ошибка")
                        End If
                    Catch
                        If par_name <> "" Then
                            MsgBox("Неправильно задан распределительный номер! Распределительный номер - 4-х значное число.",, "Ошибка")
                        End If
                    End Try
                Else
                    MsgBox("Вы не задали распределительный номер! Формат: имя/распределительный номер",, "Ошибка")
                End If
            Else
                MsgBox("Поле ввода пустое!",, "Ошибка")
            End If
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If ListBox1.SelectedIndex = -1 Then
            MsgBox("Вы не выбрали элемент списка!",, "Ошибка")
        Else
            participators.RemoveAt(ListBox1.SelectedIndex)
            ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)
            Members_commands_redacted(True, True)
        End If
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If ListBox3.SelectedIndex = -1 Then
            MsgBox("Вы не выбрали элемент списка!",, "Ошибка")
        Else
            commands.RemoveAt(ListBox3.SelectedIndex)
            ListBox3.Items.RemoveAt(ListBox3.SelectedIndex)
            Members_commands_redacted(True, True)
            Members_commands_redacted(False, True)
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If ListBox1.SelectedIndex = -1 Then
            MsgBox("Вы не выбрали элемент списка!",, "Ошибка")
        Else
            Redact.work_mode = 0
            Redact.ShowDialog()
        End If
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If ListBox3.SelectedIndex = -1 Then
            MsgBox("Вы не выбрали элемент списка!",, "Ошибка")
        Else
            Redact.work_mode = 1
            Redact.ShowDialog()
        End If
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        participators.Clear()
        commands.Clear()
        UpDate_Lists(0)
        TextBox1.Text = ""
        TextBox3.Text = ""
        ListBox2.Items.Clear()
        Rand.smooth(0) = True
        Rand.nearby(0) = True
        Rand.smooth_strench(0) = 0
        sorted(0) = False
        Members_commands_redacted(True, False)
        Members_commands_redacted(False, True)
        Messanger.Add_messange("Выполнен сброс раздела ""Команды""!", "reset com", True)
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        If participators.Count = 0 Then
            MsgBox("Отсутствуют участники!",, "Ошибка")
            Exit Sub
        End If
        If commands.Count = 0 Then
            MsgBox("Отсутствуют команды!",, "Ошибка")
            Exit Sub
        End If
        If commands.Count > participators.Count Then
            Dim result As MsgBoxResult = MsgBox("Команд больше, чем участников!" & Chr(10) & "Вы хотите продолжить?", MsgBoxStyle.YesNo, "Предупреждение")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
        End If
        If Math.IEEERemainder(participators.Count, commands.Count) <> 0 And Rand.smooth(0) = True Then
            Dim result As MsgBoxResult = MsgBox("Участники не делятся поровну на команды!" & Chr(10) & "Вы хотите продолжить?", MsgBoxStyle.YesNo, "Предупреждение")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
        End If
        If commands.Count = 1 And Rand.nearby(0) = False Then
            Dim result As MsgBoxResult = MsgBox("Настройка 'Допускать соседние элементы рядом' отключена! В списке присутствует только 1 команда. При продолжении, настройка 'Допускать соседние элементы рядом' будет включена." & Chr(10) & "Вы хотите продолжить?", MsgBoxStyle.YesNo, "Предупреждение")
            If result = MsgBoxResult.No Then
                Exit Sub
            Else
                Rand.nearby(0) = True
            End If
        End If
        Me.UseWaitCursor = True
        For Each i As Command In commands
            i.members.Clear()
        Next
        rand_commands.Clear()
        sorting_position = 0
        stage = 0
        work_mode = 0
        For Each i As Control In Me.Controls
            If i.Tag = "ToDisable" Then
                i.Enabled = False
            End If
        Next
        Progress.Reset()
        If Rand.smooth(0) = True Then
            If participators.Count * 2 >= 100 Then
                Progress.progress_mode(0) = 1
                Progress.progress_mode(1) = Math.Round(participators.Count * 2 / 100)
            Else
                Progress.progress_mode(0) = 0
                Progress.progress_mode(1) = Math.Round(100 / participators.Count)
            End If
        Else
            If participators.Count >= 100 Then
                Progress.progress_mode(0) = 1
                Progress.progress_mode(1) = Math.Round(participators.Count / 100)
            Else
                Progress.progress_mode(0) = 0
                Progress.progress_mode(1) = Math.Round(100 / participators.Count)
            End If
        End If
        Progress.Show()
        Sorting_timer.Start()
        If Mes_set.CheckBox4.Checked Then
            Messanger.Add_messange("Распределение по командам начато.", "distr_start command", True)
        End If
    End Sub
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If sorted(0) = True Then
            SaveFormat.work_mode = 0
            Dim result As DialogResult = SaveFormat.ShowDialog
            If result = DialogResult.OK Then
                Dim arr(commands.Count - 1) As String
                For i As Integer = 0 To commands.Count - 1
                    Dim res As String = ""
                    If commands(i).members.Count = 0 Then
                        If SaveFormat.CheckBox1.Checked And SaveFormat.CheckBox3.Checked Then
                            res = "Команда: " & commands(i).name & " - Н:" & commands(i).dir & ", К:0"
                        ElseIf SaveFormat.CheckBox1.Checked Then
                            res = "Команда: " & commands(i).name & " - Н:" & commands(i).dir
                        ElseIf SaveFormat.CheckBox3.Checked Then
                            res = "Команда: " & commands(i).name & " - К:0"
                        Else
                            res = "Команда: " & commands(i).name
                        End If
                        arr(i) = res
                    Else
                        If SaveFormat.CheckBox1.Checked And SaveFormat.CheckBox3.Checked Then
                            res = "Команда: " & commands(i).name & " - Н:" & commands(i).dir & ", К:" & commands(i).members.Count & "; Участники: "
                        ElseIf SaveFormat.CheckBox1.Checked Then
                            res = "Команда: " & commands(i).name & " - Н:" & commands(i).dir & "; Участники: "
                        ElseIf SaveFormat.CheckBox3.Checked Then
                            res = "Команда: " & commands(i).name & " - К:" & commands(i).members.Count & "; Участники: "
                        Else
                            res = "Команда: " & commands(i).name & "; Участники: "
                        End If
                        If SaveFormat.CheckBox2.Checked Then
                            For i1 As Integer = 0 To commands(i).members.Count - 1
                                If i1 = 0 Then
                                    res &= commands(i).members.Item(i1).name & " - Н:" & commands(i).members.Item(i1).dir
                                Else
                                    res &= ", " & commands(i).members.Item(i1).name & " - Н:" & commands(i).members.Item(i1).dir
                                End If
                            Next
                        Else
                            For i1 As Integer = 0 To commands(i).members.Count - 1
                                If i1 = 0 Then
                                    res &= commands(i).members.Item(i1).name
                                Else
                                    res &= ", " & commands(i).members.Item(i1).name
                                End If
                            Next
                        End If
                        arr(i) = res
                    End If
                Next
                Try
                    If Mes_set.CheckBox6.Checked Then
                        Messanger.Add_messange("Распределение по командам сохранено успешно!", "save com", True)
                    End If
                    IO.File.WriteAllLines(SaveFormat.SaveFile.FileName, arr)
                Catch
                    If Mes_set.CheckBox6.Checked Then
                        Messanger.Add_messange("Распределение по командам сохранено успешно!", "save com", True)
                    End If
                    MsgBox("Нет доступа к файлу или(и) папке!",, "Ошибка")
                End Try
            End If
        Else
            MsgBox("Участники не распределены по командам или текущее распределение не активно!",, "Ошибка")
        End If
    End Sub
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = 13 And e.Shift = False Then
            If TextBox1.Text <> "" Then
                Dim exits As Boolean = False
                For Each i As Member In participators
                    If i.name = TextBox1.Text Then
                        exits = True
                    End If
                Next
                If exits = False Then
                    participators.Add(New Member)
                    participators.Item(participators.Count - 1).name = TextBox1.Text
                    participators.Item(participators.Count - 1).dir = 0
                    TextBox1.Text = ""
                    UpDate_Lists(0)
                    Members_commands_redacted(True, True)
                Else
                    MsgBox("Участник с таким именем уже добавлен в список!",, "Ошибка")
                End If
            Else
                MsgBox("Вы не ввели имя участника!",, "Ошибка")
            End If
        ElseIf e.KeyCode = 13 And e.Shift = True Then
            If TextBox1.Text <> "" Then
                If TextBox1.Text.IndexOf("/") > -1 Then
                    Dim par_name As String = "0"
                    Try
                        If Mid(TextBox1.Text, TextBox1.Text.LastIndexOf("/") + 2).Length > 4 Then Error 1
                        Dim r_number As UShort = Mid(TextBox1.Text, TextBox1.Text.LastIndexOf("/") + 2)
                        par_name = Mid(TextBox1.Text, 1, TextBox1.Text.LastIndexOf("/"))
                        If par_name = "" Then
                            MsgBox("Вы не ввели имя участника!",, "Ошибка")
                            Error 1
                        End If
                        Dim exits As Boolean = False
                        For Each i As Member In participators
                            If i.name = par_name Then
                                exits = True
                            End If
                        Next
                        If exits = False Then
                            participators.Add(New Member)
                            participators.Item(participators.Count - 1).name = par_name
                            participators.Item(participators.Count - 1).dir = r_number
                            TextBox1.Text = ""
                            UpDate_Lists(0)
                            Members_commands_redacted(True, True)
                        Else
                            MsgBox("Участник с таким именем уже добавлен в список!",, "Ошибка")
                        End If
                    Catch
                        If par_name <> "" Then
                            MsgBox("Неправильно задан распределительный номер! Распределительный номер - 4-х значное число.",, "Ошибка")
                        End If
                    End Try
                Else
                    MsgBox("Вы не задали распределительный номер! Формат: имя/распределительный номер",, "Ошибка")
                End If
            Else
                MsgBox("Поле ввода пустое!",, "Ошибка")
            End If
        End If
    End Sub
    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = 13 And e.Shift = False Then
            If TextBox3.Text <> "" Then
                Dim exits As Boolean = False
                For Each i As Command In commands
                    If i.name = TextBox3.Text Then
                        exits = True
                    End If
                Next
                If exits = False Then
                    commands.Add(New Command)
                    commands.Item(commands.Count - 1).name = TextBox3.Text
                    commands.Item(commands.Count - 1).dir = 0
                    TextBox3.Text = ""
                    UpDate_Lists(0)
                    Members_commands_redacted(True, True)
                Else
                    MsgBox("Команда с таким названием уже добавлена в список!",, "Ошибка")
                End If
            Else
                MsgBox("Вы не ввели название команды!",, "Ошибка")
            End If
        ElseIf e.KeyCode = 13 And e.Shift = True Then
            If TextBox3.Text <> "" Then
                If TextBox3.Text.IndexOf("/") > -1 Then
                    Dim par_name As String = "0"
                    Try
                        If Mid(TextBox3.Text, TextBox3.Text.LastIndexOf("/") + 2).Length > 4 Then Error 1
                        Dim r_number As UShort = Mid(TextBox3.Text, TextBox3.Text.LastIndexOf("/") + 2)
                        par_name = Mid(TextBox3.Text, 1, TextBox3.Text.LastIndexOf("/"))
                        If par_name = "" Then
                            MsgBox("Вы не ввели название команды!",, "Ошибка")
                            Error 1
                        End If
                        Dim exits As Boolean = False
                        For Each i As Command In commands
                            If i.name = par_name Then
                                exits = True
                            End If
                        Next
                        If exits = False Then
                            commands.Add(New Command)
                            commands.Item(commands.Count - 1).name = par_name
                            commands.Item(commands.Count - 1).dir = r_number
                            TextBox3.Text = ""
                            UpDate_Lists(0)
                            Members_commands_redacted(True, True)
                            Members_commands_redacted(False, True)
                        Else
                            MsgBox("Команда с таким названием уже добавлена в список!",, "Ошибка")
                        End If
                    Catch
                        If par_name <> "" Then
                            MsgBox("Неправильно задан распределительный номер! Распределительный номер - 4-х значное число.",, "Ошибка")
                        End If
                    End Try
                Else
                    MsgBox("Вы не задали распределительный номер! Формат: имя/распределительный номер",, "Ошибка")
                End If
            Else
                MsgBox("Поле ввода пустое!",, "Ошибка")
            End If
        End If
    End Sub
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Rand.work_mode = 0
        Rand.ShowDialog()
    End Sub
    Private Sub Label2_MouseDown(sender As Object, e As MouseEventArgs) Handles Label2.MouseDown
        If My.Computer.Keyboard.ShiftKeyDown = True Then
            Members_commands_redacted(True, False)
        End If
    End Sub
#End Region
    Private Structure MC_rand_memory
        Public member, com As Integer
        Public Sub New(participator As Integer, command As Integer)
            member = participator
            com = command
        End Sub
    End Structure
End Class
Public Class Messanger
    Public Shared Sub Renovate_messanger(add_mode As Boolean)
        Main.ListBox4.Items.Clear()
        Dim max_widht As Short = 0
        For i As Integer = 0 To Main.messanges.Count - 1
            Main.ListBox4.Items.Add(Main.messanges(i).text)
            Main.Label9.Text = Main.messanges(i).text
            If Main.Label9.Width > max_widht Then
                max_widht = Main.Label9.Width
            End If
        Next
        Main.Label9.Text = ""
        max_widht = (Int(max_widht / 40) + 1) * 40
        If max_widht <= 120 Then
            Main.ListBox4.Width = 120
        Else
            Main.ListBox4.Width = max_widht
        End If
        If Main.ListBox4.Items.Count <= 24 Then
            Main.ListBox4.Height = 312
        Else
            Main.ListBox4.Height = Main.ListBox4.Items.Count * 13
        End If
        Main.scr_lab.Top = Main.ListBox4.Top + Main.ListBox4.Height - 13
        Main.Label7.Text = "Уведомления (" & Main.messanges.Count & "):"
        Main.Button11.Enabled = False
        Main.Button12.Enabled = False
        If Main.ListBox4.Items.Count > 0 Then
            If add_mode Then
                Main.Panel1.ScrollControlIntoView(Main.scr_lab)
                Main.ListBox4.SelectedIndex = Main.ListBox4.Items.Count - 1
                Main.Button11.Enabled = True
            Else
                Main.Panel1.ScrollControlIntoView(Main.scr_lab)
            End If
        End If
    End Sub
    Public Shared Sub Clear_inactive_messanger()
        Dim temp_arr(-1) As Short, del_pos As Short = 0
        For i As Short = 0 To Main.messanges.Count - 1
            If Main.messanges(i).link = "" Or Main.messanges(i).read_only = True Then
                Array.Resize(temp_arr, temp_arr.Length + 1)
                temp_arr(temp_arr.Length - 1) = i
            End If
        Next
        For i As Short = 0 To temp_arr.Length - 1
            Delete_messange(temp_arr(i) - del_pos)
            del_pos += 1
        Next
        Renovate_messanger(False)
    End Sub
    Public Shared Sub Clear_messanger()
        Main.messanges.Clear()
        Renovate_messanger(False)
    End Sub
    Public Shared Sub Renovate_messange(link As String, text As String)
        Dim position As Short = Get_position_by_link(link)
        If position > -1 Then
            If Main.messanges(position).read_only = False And link <> "" Then
                Exit Sub
            End If
            Main.messanges(position).text = text
            Main.ListBox4.Items(position) = Main.messanges(position).text
            Dim max_widht As Short = 0
            For i As Integer = 0 To Main.messanges.Count - 1
                Main.Label9.Text = Main.messanges(i).text
                If Main.Label9.Width > max_widht Then
                    max_widht = Main.Label9.Width
                End If
            Next
            Main.Label9.Text = ""
            max_widht = Math.Floor(max_widht / 40) * 40
            If max_widht <= 120 Then
                Main.ListBox4.Width = 120
            Else
                Main.ListBox4.Width = max_widht
            End If
        End If
    End Sub
    Public Shared Sub Delete_messange(position As Short)
        If position > -1 Then
            Main.messanges.RemoveAt(position)
            Renovate_messanger(False)
        End If
    End Sub
    Public Shared Function Get_position_by_link(link As String) As Short
        Dim pos As Short = -1
        For i As Short = 0 To Main.messanges.Count - 1
            If Main.messanges(i).link = link And link <> "" And Main.messanges(i).read_only = False Then
                pos = i
                Exit For
            End If
        Next
        Return pos
    End Function
    Public Shared Sub Add_messange(text As String, link As String, read_only As Boolean)
        Dim position As Short = Get_position_by_link(link)
        If position = -1 Then
            If Mes_set.CheckBox3.Checked Then
                My.Computer.Audio.Play(My.Resources.mess_add, AudioPlayMode.Background)
            End If
            If (link = "" Or read_only = True) And (Mes_set.CheckBox1.Checked) Then
                Dim time_temp As Date = My.Computer.Clock.LocalTime
                Main.messanges.Add(New Messange_temp(text & " - " & Format(time_temp.Hour, "00") & ":" & Format(time_temp.Minute, "00") & ":" & Format(time_temp.Second, "00"), link, read_only))
                Renovate_messanger(True)
            Else
                Main.messanges.Add(New Messange_temp(text, link, read_only))
                Renovate_messanger(True)
            End If
        End If
    End Sub
    Shared lib_list As New List(Of Link_lib_part)
    Public Shared Sub Add_link_lib(link_name As String, tab_page As Short, x As Integer, y As Integer, width As Integer, height As Integer)
        If link_name <> "" Then
            lib_list.Add(New Link_lib_part)
            lib_list(lib_list.Count - 1).link_name = link_name
            lib_list(lib_list.Count - 1).tab_page = tab_page
            lib_list(lib_list.Count - 1).x = x
            lib_list(lib_list.Count - 1).y = y
            lib_list(lib_list.Count - 1).width = width
            lib_list(lib_list.Count - 1).height = height
        End If
    End Sub
    Public Shared Function Get_link_lib(link_name As String) As Link_lib_part
        Dim pos As New Link_lib_part, pos_find As Boolean = False
        For i As Short = 0 To lib_list.Count - 1
            If lib_list(i).link_name = link_name Then
                pos = lib_list(i)
                pos_find = True
                Exit For
            End If
        Next
        If pos_find = True Then
            Return pos
        Else
            pos.tab_page = -1
            Return pos
        End If
    End Function
    Public Class Link_lib_part
        Public link_name As String, tab_page As Short, x As Integer, y As Integer, width As Integer, height As Integer
    End Class
End Class
Public Class Messange_temp
    Public text, link As String, read_only As Boolean
    Public Sub New(text As String, link As String, read_only As Boolean)
        Me.text = text
        Me.link = link
        Me.read_only = read_only
    End Sub
End Class
Public Class Timer_temp
    Public name As String = "", new_interval As UInteger = 0, active As Boolean = False, work_messange(3) As Boolean 'work_messange(0) - timer end window, work_messange(1) - timer tick messange, work_messange(2-3) - timer start/stop messange
    Dim interval1 As UInteger = 0, con_type As UShort = 0, trig_action As Trigger_action = Trigger_action.none, trig_data As String = ""
    Dim mem As New List(Of Member)
    Dim com As New List(Of Command)
    Dim tour As New List(Of Tournier)
    Public Property connect_type() As UShort
        Get
            Return con_type
        End Get
        Set(ByVal value As UShort)
            If con_type <> value Then
                mem.Clear()
                com.Clear()
                tour.Clear()
            End If
            con_type = value
        End Set
    End Property
    Public Property interval() As UInteger
        Get
            Return interval1
        End Get
        Set(ByVal value As UInteger)
            If interval1 <> value Then
                new_interval = value
            End If
            interval1 = value
        End Set
    End Property
    Public Property members() As List(Of Member)
        Get
            Return mem
        End Get
        Set(ByVal value As List(Of Member))
            If con_type = 1 Then
                mem = value
            End If
        End Set
    End Property
    Public Property commands() As List(Of Command)
        Get
            Return com
        End Get
        Set(ByVal value As List(Of Command))
            If con_type = 2 Then
                com = value
            End If
        End Set
    End Property
    Public Property tourniers() As List(Of Tournier)
        Get
            Return tour
        End Get
        Set(ByVal value As List(Of Tournier))
            If con_type = 3 Then
                tour = value
            End If
        End Set
    End Property
    Public Property action() As Trigger_action
        Get
            Return trig_action
        End Get
        Set(ByVal value As Trigger_action)
            If trig_action <> value Then
                trig_data = ""
            End If
            trig_action = value
        End Set
    End Property
    Public Property action_data() As String
        Get
            Return trig_data
        End Get
        Set(ByVal value As String)
            If trig_action = 2 Then
                trig_data = value
            End If
        End Set
    End Property
    Public Enum Trigger_action
        none = 0
        Repeat_timer = 1
        Start_over_timer = 2
        Sort_members = 3
        Sort_commands = 4
        Delete_this_select = 5
    End Enum
End Class
Public Class Member
    Public name As String = "", dir As UInteger = 0
End Class
Public Class Command
    Public name As String = "", dir As UInteger = 0, members As New List(Of Member)
End Class
Public Class Tournier
    Public name As String = "", dir As UInteger = 0, commands As New List(Of Command)
End Class