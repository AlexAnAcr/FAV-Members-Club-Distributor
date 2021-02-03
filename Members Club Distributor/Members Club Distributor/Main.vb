Imports System.ComponentModel

Public Class Main
    Public participators As New List(Of Member), dbf_info As db_info_file, dbf_collector As New db_collector_file, db_data As New List(Of db_Participator), db_not_saved As Boolean = False, sdb_keys As New ArrayList
    Public commands As New List(Of Command)
    Public tourniers As New List(Of Tournier)
    Public messanges As New List(Of Messange_temp), settings(8) As String, info_set(2) As String, timer_func As New Timer_page, timer_button_mode As Short = 1, timer_resourse As WMPLib.IWMPMedia
    Dim pic_stage As Short = 0, sorted(1) As Boolean
    Public WithEvents player1 As New WMPLib.WindowsMediaPlayer
    Public Shared TEMP_WAY As String = Environ("Temp"), ME_WAY As String, MY_WAY As String = Application.StartupPath, MY_VERSION As String = "2.1"
    Private Sub Main_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove
        If pic_stage = 1 Then
            PictureBox1.Image = My.Resources.pd
            pic_stage = 0
        End If
    End Sub
    Private Sub PictureBox1_MouseMove(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseMove
        If pic_stage = 0 Then
            PictureBox1.Image = My.Resources.pd_h
            pic_stage = 1
        End If
    End Sub
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        PictureBox1.Image = My.Resources.pd
        About.ShowDialog()
    End Sub
    Private Sub Main_Load(sender As Object, e As EventArgs) Handles Me.Load
        If My.Computer.FileSystem.FileExists(MY_WAY & "\mcdCAnc.exe") = False Then
            MsgBox("Members Club Distributor установлен некорректно!" & Chr(10) & "Не удаётся найти некоторые компоненты программы!" & Chr(10) & "Совет: переустановите программу.", , "Ошибка")
            is_system_closing = True
            Me.Close()
        End If
        Try
            My.Computer.FileSystem.WriteAllText(MY_WAY & "/testing.tmp", "Test data.", False)
            My.Computer.FileSystem.DeleteFile(MY_WAY & "/testing.tmp")
        Catch
            MsgBox("Members Club Distributor установлен некорректно!" & Chr(10) & "Возможная причина ошибки: Members Club Distributor для работы на данном компьютере необходимы права администратора (актуальна для Windows 7 и выше), проблемы совместимости с вашей версией или пакетом обновления Windows (вероятна для Windows 8 и выше)." & Chr(10) & "Совет: Попробуйте воспользоваться средствами исправления неполадок совместимости или предоставьте Members Club Distributor права администратора. Если вышеперечисленное не помогло, попробуйте переустановить программу.", , "Ошибка")
            is_system_closing = True
            Me.Close()
        End Try
        If My.Computer.FileSystem.FileExists(MY_WAY & "\mcd_instruction_19679879038.dat") = False Then
            ME_WAY = Application.ExecutablePath
        Else
            ME_WAY = MY_WAY & "\" & My.Computer.FileSystem.ReadAllText(MY_WAY & "\mcd_instruction_19679879038.dat", System.Text.Encoding.ASCII) & ".exe"
        End If
        Dim mcd_canc() As Process = Process.GetProcessesByName("mcdCAnc")
        If mcd_canc.Length > 0 Then
            Dim finded As Boolean = False
            For Each i As Process In mcd_canc
                Try
                    If My.Computer.FileSystem.GetFileInfo(i.MainModule.FileName).DirectoryName = MY_WAY Then
                        finded = True
                    End If
                Catch
                End Try
                If finded = True Then
                    Exit For
                End If
            Next
            If finded = True Then
                MsgBox("Программа Members Club Distributor не может быть запущена сейчас!" & Chr(10) & "Повторите попытку запуска программы позже.", , "Ошибка")
                is_system_closing = True
                Me.Close()
            End If
        End If
        If My.Computer.FileSystem.DirectoryExists(TEMP_WAY & "\MCD-temp-47591795970905") = False Then
            My.Computer.FileSystem.CreateDirectory(TEMP_WAY & "\MCD-temp-47591795970905")
            My.Computer.FileSystem.WriteAllText(TEMP_WAY & "\MCD-temp-47591795970905\mcd_j.tmp", ME_WAY, False)
        Else
            If My.Computer.FileSystem.FileExists(TEMP_WAY & "\MCD-temp-47591795970905\mcd_j.tmp") = False Then
                My.Computer.FileSystem.WriteAllText(TEMP_WAY & "\MCD-temp-47591795970905\mcd_j.tmp", ME_WAY, False)
            Else
                Dim mcd_j_dat As String = My.Computer.FileSystem.ReadAllText(TEMP_WAY & "\MCD-temp-47591795970905\mcd_j.tmp"), is_duble As Short = 0, proc() As Process = Process.GetProcesses
                For Each i As Process In proc
                    Try
                        If i.MainModule.FileName.ToLower = ME_WAY.ToLower Then
                            is_duble += 1
                        ElseIf i.MainModule.FileName.ToLower = mcd_j_dat.ToLower Then
                            is_duble = 2
                        End If
                        If is_duble = 2 Then
                            Exit For
                        End If
                    Catch
                    End Try
                Next
                If is_duble = 2 Then
                    MsgBox("Members Club Distributor уже запущен!", , "Ошибка")
                    is_system_closing = True
                    Me.Close()
                Else
                    My.Computer.FileSystem.WriteAllText(TEMP_WAY & "\MCD-temp-47591795970905\mcd_j.tmp", ME_WAY, False)
                End If
            End If
        End If
        Label20.Left = 10
        Label20.Top = 10
        Label20.Width = Me.Width - 20
        Label20.Height = Me.Height - 20
        Label24.Left = 0
        Label24.Top = 0
        Label24.Width = TabPage4.Width
        Label24.Height = TabPage4.Height
    End Sub
    Dim is_system_closing As Boolean = False, is_system_closing_nmd As Boolean = False
    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If is_system_closing = False Then
            If is_system_closing_nmd = False Then
                If db_not_saved = True Then
                    Dim result As MsgBoxResult = MsgBox("Изменения в базе данных Участников не сохранены!" & Chr(10) & "Вы действительно хотите закрыть программу без сохранения базы данных Участников?", MsgBoxStyle.YesNo, "Сообщение")
                    If result = MsgBoxResult.No Then
                        e.Cancel = True
                        Exit Sub
                    End If
                End If
                Dim proc() As Process = Process.GetProcessesByName("mcdCAnc.exe")
                If proc.Length = 0 Then
                    cmd_service_wait_proc = Shell(MY_WAY & "\mcdCAnc.exe ""MCDmode:dir-delete,Code:2802"" """ & TEMP_WAY & "\MCD-temp-47591795970905""")
                    e.Cancel = True
                    cmd_service_wait_mode = 1
                    cmd_service_wait.Start()
                    is_system_closing_nmd = True
                    Label20.Text = "Завершение работы..."
                    Label20.Visible = True
                Else
                    Dim is_double As Boolean = False
                    For Each i As Process In proc
                        Try
                            If i.MainModule.FileName.ToLower = (MY_WAY & "\mcdCAnc.exe").ToLower Then
                                is_double = True
                                cmd_service_wait_proc = i.Id
                                Exit For
                            End If
                        Catch
                        End Try
                    Next
                    If is_double = False Then
                        cmd_service_wait_proc = Shell(MY_WAY & "\mcdCAnc.exe ""MCDmode:dir-delete,Code:2802"" """ & TEMP_WAY & "\MCD-temp-47591795970905""")
                        e.Cancel = True
                        cmd_service_wait_mode = 1
                        cmd_service_wait.Start()
                        is_system_closing_nmd = True
                        Label20.Text = "Завершение работы..."
                        Label20.Visible = True
                    Else
                        MsgBox("Программа Members Club Distributor не может быть закрыта сейчас! Пожалуйста, подождите..." & Chr(10) & "Программа закроется автоматически, как только это будет возможно.",, "Сообщение")
                        e.Cancel = True
                        cmd_service_wait_mode = 0
                        cmd_service_wait.Start()
                        is_system_closing_nmd = True
                        Label20.Text = "Завершение работы..."
                        Label20.Visible = True
                    End If
                End If
            Else
                e.Cancel = True
                Me.Hide()
            End If
        End If
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
            Dim names() As String = {"StartMessange", "InactiveMessangeTime", "MessangeSound", "TimerSoundWay", "DistributingStartMessange", "DistributingEndMessange", "SaveMessange", "DBLoad", "DBLoaded"}, types() As Microsoft.Win32.RegistryValueKind = {Microsoft.Win32.RegistryValueKind.DWord, 4, 4, Microsoft.Win32.RegistryValueKind.String, 4, 4, 4, 4, 4}, values() As String = {"1", "1", "1", "", "1", "1", "1", "1", "1"}, reg_values() As String = reg.GetValueNames
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
            settings(7) = reg.GetValue("DBLoad")
            settings(8) = reg.GetValue("DBLoaded")
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
            If settings(7) = "1" Then
                Mes_set.CheckBox7.Checked = True
            Else
                Mes_set.CheckBox7.Checked = False
            End If
            If settings(8) = "1" Then
                Mes_set.CheckBox8.Checked = True
            Else
                Mes_set.CheckBox8.Checked = False
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
            reg.SetValue("DBLoad", "1", Microsoft.Win32.RegistryValueKind.DWord)
            reg.SetValue("DBLoaded", "1", Microsoft.Win32.RegistryValueKind.DWord)
            settings(0) = "1"
            settings(1) = "1"
            settings(2) = "1"
            settings(3) = ""
            settings(4) = "1"
            settings(5) = "1"
            settings(6) = "1"
            settings(7) = "1"
            settings(8) = "1"
            reg.Close()
        End If
        If My.Computer.FileSystem.FileExists(MY_WAY & "\mcdinfo.dat") Then
            Dim search_pattern() As String = {"username", "organization", "programID"}, file_data() As String = IO.File.ReadAllLines(MY_WAY & "\mcdinfo.dat"), re_write As Boolean = False
            For i As Short = 0 To search_pattern.Length - 1
                Dim finded As Boolean = False
                For i1 As Short = 0 To file_data.Length - 1
                    Dim split_data() As String = file_data(i1).Split("=")
                    If split_data.Length = 2 Then
                        If split_data(0) = search_pattern(i) Then
                            info_set(i) = split_data(1)
                            finded = True
                            Exit For
                        End If
                    End If
                Next
                If finded = False Then
                    If i = 0 Or i = 1 Then
                        info_set(i) = ""
                    ElseIf i = 2 Then
                        info_set(2) = Generate_code()
                    End If
                    re_write = True
                End If
            Next
            If info_set(2) = "" Then
                info_set(2) = Generate_code()
                re_write = True
            Else
                Dim simvols_arr() As String = {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}
                For i As UShort = 1 To info_set(2).Length
                    Dim no_simvol As Boolean = False
                    For Each i1 As String In simvols_arr
                        If Mid(info_set(2), i, 1) = simvols_arr(i1) Then
                            no_simvol = True
                        End If
                    Next
                    If no_simvol = False Then
                        info_set(2) = Generate_code()
                        re_write = True
                        Exit For
                    End If
                Next
            End If
            If re_write = True Then
                Info_set_write()
            End If
        Else
            My.Computer.FileSystem.WriteAllText(MY_WAY & "\mcdinfo.dat", "", False)
            info_set(0) = ""
            info_set(1) = ""
            info_set(2) = Generate_code()
            Info_set_write()
        End If
        Messanger.Add_link_lib("mem_dist inactual", 0, 145, 325, 210, 50)
        Messanger.Add_link_lib("com_dist inactual", 1, 145, 325, 210, 50)
        Messanger.Add_link_lib("distr_start command", 0, 515, 355, 157, 97)
        Messanger.Add_link_lib("distr_end command", 0, 15, 355, 515, 180)
        Messanger.Add_link_lib("distr_start tournier", 1, 515, 355, 157, 97)
        Messanger.Add_link_lib("distr_end tournier", 1, 15, 355, 515, 180)
        Messanger.Add_link_lib("reset com", 0, 515, 485, 157, 50)
        Messanger.Add_link_lib("reset tour", 1, 515, 485, 157, 50)
        Messanger.Add_link_lib("save com", 0, 515, 440, 157, 60)
        Messanger.Add_link_lib("save tour", 1, 515, 440, 157, 60)
        Messanger.Add_link_lib("timer progress", 2, 105, 215, 486, 100)
        Messanger.Add_link_lib("timer start/stop", 2, 25, 495, 442, 50)
        Messanger.Add_link_lib("timer reset", 2, 445, 495, 110, 50)
        Messanger.Add_link_lib("db not loaded", 3, 290, 230, 74, 84)
        Messanger.Add_link_lib("db loading", 3, 290, 230, 74, 84)
        Messanger.Add_link_lib("db loaded", 3, 25, 67, 640, 455)
        StartFileCopy.RunWorkerAsync()
    End Sub
    Public Sub Info_set_write()
        Dim write_info() As String = {"username=" & info_set(0), "organization=" & info_set(1), "programID=" & info_set(2)}
        IO.File.WriteAllLines(MY_WAY & "\mcdinfo.dat", write_info)
    End Sub
    Private Sub StartFileCopy_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles StartFileCopy.DoWork
        My.Computer.FileSystem.WriteAllBytes(TEMP_WAY & "\MCD-temp-47591795970905\tds.mp3", My.Resources.timer, False)
        If My.Computer.FileSystem.FileExists(MY_WAY & "\database.dat") And My.Computer.FileSystem.DirectoryExists(TEMP_WAY & "\MCD-temp-47591795970905\database") Then
            My.Computer.FileSystem.DeleteDirectory(TEMP_WAY & "\MCD-temp-47591795970905\database", FileIO.DeleteDirectoryOption.DeleteAllContents)
        End If
        If My.Computer.FileSystem.DirectoryExists(TEMP_WAY & "\MCD-temp-47591795970905\temp") Then
            My.Computer.FileSystem.DeleteDirectory(TEMP_WAY & "\MCD-temp-47591795970905\temp", FileIO.DeleteDirectoryOption.DeleteAllContents)
        End If
    End Sub
    Private Sub StartFileCopy_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles StartFileCopy.RunWorkerCompleted
        If settings(3) = "" Then
            timer_resourse = player1.newMedia(TEMP_WAY & "\MCD-temp-47591795970905\tds.mp3")
        Else
            If My.Computer.FileSystem.FileExists(settings(3)) Then
                If My.Computer.FileSystem.GetFileInfo(settings(3)).Extension.ToLower = ".mp3" Then
                    timer_resourse = player1.newMedia(settings(3))
                Else
                    timer_resourse = player1.newMedia(TEMP_WAY & "\MCD-temp-47591795970905\tds.mp3")
                End If
            Else
                timer_resourse = player1.newMedia(TEMP_WAY & "\MCD-temp-47591795970905\tds.mp3")
            End If
        End If
        timer_func.work_messange(0) = True
        timer_func.work_messange(1) = False
        timer_func.work_messange(2) = False
        timer_func.work_messange(3) = False
        timer_func.work_messange(4) = False
        timer_func.select_time(0) = 10
        timer_func.select_time(1) = 10
        timer_func.select_time(2) = 5
        If Mes_set.CheckBox2.Checked Then
            Messanger.Add_messange("Добро пожаловать в MCD!", "", True)
        End If
        If My.Computer.FileSystem.FileExists(MY_WAY & "\database.dat") Then
            My.Computer.FileSystem.CreateDirectory(TEMP_WAY & "\MCD-temp-47591795970905\database")
            If Mes_set.CheckBox7.Checked Then
                Messanger.Add_messange("Идёт загрузка базы данных Участников...", "db loading", False)
            End If
            cmd_service_wait_proc = Shell(MY_WAY & "\mcdCAnc.exe ""MCDmode:archiver,Code:2801"" ""decompress"" """ & MY_WAY & "\database.dat"" """ & TEMP_WAY & "\MCD-temp-47591795970905\database""")
            cmd_service_wait_mode = 2
            cmd_service_wait.Start()
        Else
            If My.Computer.FileSystem.DirectoryExists(TEMP_WAY & "\MCD-temp-47591795970905\database") And My.Computer.FileSystem.FileExists(MY_WAY & "\database.dat") = False Then
                Dim f_names() As String = IO.Directory.GetFiles(TEMP_WAY & "\MCD-temp-47591795970905\database", "*", IO.SearchOption.AllDirectories)
                If f_names.Length > 0 Then
                    MsgBox("База данных Участников отсутствует, однако, были найдены её остаточные данные." & Chr(10) & "MCD предпримет попытку восстановить базу данных.",, "Сообщение")
                    Messanger.Add_messange("База данных MCD отсутствует, но будет предпринята попытка восстановления.", "", True)
                    If Mes_set.CheckBox7.Checked Then
                        Messanger.Add_messange("Идёт загрузка базы данных Участников...", "db loading", False)
                    End If
                    cmd_service_wait_proc = Shell(MY_WAY & "\mcdCAnc.exe ""MCDmode:archiver,Code:2801"" ""compress"" """ & MY_WAY & "\database.dat"" """ & TEMP_WAY & "\MCD-temp-47591795970905\database""")
                    cmd_service_wait_mode = 2
                    cmd_service_wait.Start()
                End If
            ElseIf My.Computer.FileSystem.DirectoryExists(TEMP_WAY & "\MCD-temp-47591795970905\database") = False Then
                My.Computer.FileSystem.CreateDirectory(TEMP_WAY & "\MCD-temp-47591795970905\database")
                If My.Computer.FileSystem.FileExists(MY_WAY & "\database.dat") = False Then
                    Dim info_file(), collector_file() As String
                    info_file = {"[format]", "LmsI-CLrw-6o2m1i", "", "[versions]", MY_VERSION}
                    collector_file = {"[ADB info]", "LastReadTime=" & Get_date_string(My.Computer.Clock.LocalTime), "LastWriteTime=" & Get_date_string(My.Computer.Clock.LocalTime), "progID=" & info_set(2), "userName=" & info_set(0), "organization=" & info_set(1), "", "[participators]"}
                    IO.File.WriteAllLines(TEMP_WAY & "\MCD-temp-47591795970905\database\info.ini", info_file)
                    IO.File.WriteAllLines(TEMP_WAY & "\MCD-temp-47591795970905\database\collector.ini", collector_file)
                    MsgBox("База данных Участников отсутствует!" & Chr(10) & "MCD пересоздаст базу данных.",, "Сообщение")
                    Messanger.Add_messange("База данных MCD отсутствует!", "", True)
                    If Mes_set.CheckBox7.Checked Then
                        Messanger.Add_messange("Идёт загрузка базы данных Участников...", "db loading", False)
                    End If
                    cmd_service_wait_proc = Shell(MY_WAY & "\mcdCAnc.exe ""MCDmode:archiver,Code:2801"" ""compress"" """ & MY_WAY & "\database.dat"" """ & TEMP_WAY & "\MCD-temp-47591795970905\database""")
                    cmd_service_wait_mode = 2
                    cmd_service_wait.Start()
                End If
            End If
        End If
        If My.Computer.FileSystem.DirectoryExists(TEMP_WAY & "\MCD-temp-47591795970905\temp") = False Then
            My.Computer.FileSystem.CreateDirectory(TEMP_WAY & "\MCD-temp-47591795970905\temp")
        End If
        Label20.Visible = False
        If info_set(0) = "" Or info_set(1) = "" Then
            my_info.ShowDialog()
            If info_set(0) = "" Or info_set(1) = "" Then
                Messanger.Add_messange("Информация о пользователе не задана!", "", True)
            End If
        End If
        If info_set(1) <> "" Then
            Me.Text &= " (" & info_set(1) & ")"
        End If
        DataGridView1.Sort(DataGridView1.Columns(2), ListSortDirection.Ascending)
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
    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
        Timer_redact.ShowDialog()
    End Sub
    Public Sub Renovate_timer(all_renovate As Boolean)
        Dim temp_time As hms_temp = Get_hms(timer_func.new_interval)
        Label17.Text = Format(temp_time.hours, "00") & ":" & Format(temp_time.minutes, "00") & ":" & Format(temp_time.seconds, "00")
        full_screen.Label1.Text = Format(temp_time.hours, "00") & ":" & Format(temp_time.minutes, "00") & ":" & Format(temp_time.seconds, "00")
        temp_time = Get_hms(timer_func.interval - timer_func.new_interval)
        Label23.Text = Format(temp_time.hours, "00") & ":" & Format(temp_time.minutes, "00") & ":" & Format(temp_time.seconds, "00")
        If all_renovate = True Then
            temp_time = Get_hms(timer_func.interval)
            Label18.Text = Format(temp_time.hours, "00") & ":" & Format(temp_time.minutes, "00") & ":" & Format(temp_time.seconds, "00")
        End If
    End Sub
    Private Sub Timer_tick_Tick(sender As Object, e As EventArgs) Handles Timer_timer.Tick
        If timer_button_mode = 2 Then
            If timer_func.new_interval = 1 Then
                timer_func.new_interval = 0
                Renovate_timer(False)
                Timer_timer.Stop()
                timer_button_mode = 3
                Button24.Text = "Ответить"
                Button25.Enabled = False
                Button26.Enabled = False
                full_screen.Button2.Enabled = True
                player1.currentPlaylist.clear()
                player1.currentPlaylist.appendItem(timer_resourse)
                player1.controls.play()
                If timer_func.work_messange(0) Then
                    Messanger.Delete_messange(Messanger.Get_position_by_link("timer progress"))
                End If
                If timer_func.work_messange(1) Then
                    Messanger.Add_messange("Таймер сработал.", "timer start/stop", True)
                End If
            Else
                timer_func.new_interval -= 1
                Renovate_timer(False)
                If timer_func.work_messange(0) Then
                    Dim temp_time As hms_temp = Get_hms(timer_func.new_interval)
                    Messanger.Renovate_messange("timer progress", "Время таймера: " & Format(temp_time.hours, "00") & ":" & Format(temp_time.minutes, "00") & ":" & Format(temp_time.seconds, "00"))
                End If
                If timer_func.work_messange(2) Then
                    If timer_func.new_interval <= timer_func.select_time(0) Then
                        Label17.ForeColor = Color.DarkRed
                        full_screen.Label1.ForeColor = Color.DarkRed
                    End If
                End If
                If timer_func.work_messange(4) Then
                    If timer_func.new_interval <= timer_func.select_time(2) Then
                        My.Computer.Audio.Play(My.Resources.tick_two, AudioPlayMode.Background)
                    ElseIf timer_func.new_interval <= timer_func.select_time(1) Then
                        If timer_func.work_messange(3) Then
                            My.Computer.Audio.Play(My.Resources.tick_one, AudioPlayMode.Background)
                        End If
                    End If
                ElseIf timer_func.work_messange(3) Then
                    If timer_func.new_interval <= timer_func.select_time(1) Then
                        My.Computer.Audio.Play(My.Resources.tick_one, AudioPlayMode.Background)
                    End If
                End If
            End If
        ElseIf timer_button_mode = 3 Then
            Timer_timer.Stop()
            player1.currentPlaylist.clear()
            player1.currentPlaylist.appendItem(timer_resourse)
            player1.controls.play()
        End If
    End Sub
    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        If timer_button_mode = 1 Then
            Timer_timer.Start()
            timer_button_mode = 2
            Button24.Text = "Стоп"
            Button27.Enabled = False
            full_screen.Button2.Enabled = False
            If timer_func.work_messange(1) Then
                Messanger.Add_messange("Таймер запущен.", "timer start/stop", True)
            End If
            If timer_func.work_messange(0) Then
                Dim temp_time As hms_temp = Get_hms(timer_func.new_interval)
                Messanger.Add_messange("Время таймера: " & Format(temp_time.hours, "00") & ":" & Format(temp_time.minutes, "00") & ":" & Format(temp_time.seconds, "00"), "timer progress", False)
            End If
        ElseIf timer_button_mode = 2 Then
            Timer_timer.Stop()
            timer_button_mode = 1
            Button24.Text = "Старт"
            Button27.Enabled = True
            full_screen.Button2.Enabled = False
            Label17.ForeColor = Color.Black
            full_screen.Label1.ForeColor = Color.Black
            If timer_func.work_messange(1) Then
                Messanger.Add_messange("Таймер остановлен.", "timer start/stop", True)
            End If
        ElseIf timer_button_mode = 3 Then
            timer_button_mode = 1
            player1.controls.stop()
            Button24.Text = "Старт"
            Button27.Enabled = True
            Button25.Enabled = True
            Label17.ForeColor = Color.Black
            full_screen.Label1.ForeColor = Color.Black
            Button24.Enabled = False
            full_screen.Button2.Enabled = False
            Button26.Enabled = True
        End If
    End Sub
    Private Sub Button25_Click() Handles player1.PlayStateChange
        If (player1.playState = WMPLib.WMPPlayState.wmppsMediaEnded) And timer_button_mode = 3 Then
            Timer_timer.Start()
        End If
    End Sub
    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        timer_func.new_interval = timer_func.interval
        Label17.ForeColor = Color.Black
        full_screen.Label1.ForeColor = Color.Black
        Button24.Enabled = True
        Renovate_timer(False)
        If timer_func.work_messange(0) Then
            Dim temp_time As hms_temp = Get_hms(timer_func.new_interval)
            Messanger.Renovate_messange("timer progress", "Время таймера: " & Format(temp_time.hours, "00") & ":" & Format(temp_time.minutes, "00") & ":" & Format(temp_time.seconds, "00"))
        End If
        If timer_func.work_messange(1) Then
            Messanger.Add_messange("Таймер сброшен.", "timer reset", True)
        End If
    End Sub
    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        full_screen.ShowDialog()
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
                    Members_commands_redacted(False, True)
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
                    Members_commands_redacted(False, True)
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
                Progress.progress_mode(1) = commands.Count * 2 / 100
            Else
                Progress.progress_mode(0) = 0
                Progress.progress_mode(1) = 100 / commands.Count
            End If
        Else
            If commands.Count >= 100 Then
                Progress.progress_mode(0) = 1
                Progress.progress_mode(1) = commands.Count / 100
            Else
                Progress.progress_mode(0) = 0
                Progress.progress_mode(1) = 100 / commands.Count
            End If
        End If
        Progress.Set_parent_(Me)
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
                Progress.progress_mode(1) = participators.Count * 2 / 100
            Else
                Progress.progress_mode(0) = 0
                Progress.progress_mode(1) = 100 / participators.Count
            End If
        Else
            If participators.Count >= 100 Then
                Progress.progress_mode(0) = 1
                Progress.progress_mode(1) = participators.Count / 100
            Else
                Progress.progress_mode(0) = 0
                Progress.progress_mode(1) = 100 / participators.Count
            End If
        End If
        Progress.Set_parent_(Me)
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
                    Members_commands_redacted(False, True)
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

#Region "Data Base"
    Public search_mode As Boolean = False
    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked And search_mode = True Then
            search_mode = False
            If My.Computer.Keyboard.ShiftKeyDown Then
                Dim selected_keys(DataGridView1.SelectedRows.Count - 1) As String
                For i As Short = 0 To selected_keys.Length - 1
                    selected_keys(i) = DataGridView1.SelectedRows(i).Cells(0).Value
                Next
                Renovate_DataGrid(0)
                If selected_keys.Length > 0 Then
                    For Each i As DataGridViewRow In DataGridView1.Rows
                        For Each i1 As String In selected_keys
                            If i.Cells(0).Value = i1 Then
                                i.Selected = True
                            End If
                        Next
                    Next
                End If
            Else
                Renovate_DataGrid(0)
            End If
        End If
    End Sub
    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked And search_mode = False Then
            search_mode = True
            If My.Computer.Keyboard.ShiftKeyDown Then
                Dim selected_keys(DataGridView1.SelectedRows.Count - 1) As String
                For i As Short = 0 To selected_keys.Length - 1
                    selected_keys(i) = DataGridView1.SelectedRows(i).Cells(0).Value
                Next
                Renovate_DataGrid(1)
                If selected_keys.Length > 0 Then
                    For Each i As DataGridViewRow In DataGridView1.Rows
                        For Each i1 As String In selected_keys
                            If i.Cells(0).Value = i1 Then
                                i.Selected = True
                            End If
                        Next
                    Next
                End If
            Else
                Renovate_DataGrid(1)
            End If
        End If
    End Sub
    Private Sub Button31_Click(sender As Object, e As EventArgs) Handles Button31.Click
        Search_db.ShowDialog()
    End Sub
    Private Sub Button29_Click(sender As Object, e As EventArgs) Handles Button29.Click
        If DataGridView1.SelectedRows.Count = 1 Then
            Add_Redact_db_data.work_mode = 3
            Add_Redact_db_data.ShowDialog()
        ElseIf DataGridView1.SelectedRows.Count = 0 Then
            MsgBox("Вы не выбрали элемент для просмотра!",, "Ошибка")
        Else
            MsgBox("Вы выбрали слишком много элементов для просмотра!",, "Ошибка")
        End If
    End Sub
    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles Button28.Click
        Add_Redact_db_data.work_mode = 1
        Add_Redact_db_data.ShowDialog()
    End Sub
    Private Sub Button33_Click(sender As Object, e As EventArgs) Handles Button33.Click
        Button_save_db()
    End Sub
    Public Sub Button_save_db()
        Button33.Enabled = False
        CheckBox1.Enabled = False
        db_not_saved = True
        Help.SetToolTip(PictureBox2, "Идёт сохранение изменений базы данных.")
        PictureBox2.Image = My.Resources.save_anim
        DB_WriteDataBase(4, False)
    End Sub
    Public Sub Renovate_DataGrid(mode As Short)
        DataSet1.Tables(0).Rows.Clear()
        If mode = 0 Then
            For Each i As db_Participator In db_data
                Dim temp_arr() As Object = {i.img, i.name, i.year, i.school, i.telephone, i.post, i.key}
                DataSet1.Tables(0).Rows.Add(temp_arr)
            Next
        ElseIf mode = 1 Then
            For Each i As db_Participator In db_data
                For Each i1 As String In sdb_keys
                    If i.key = i1 Then
                        Dim temp_arr() As Object = {i.img, i.name, i.year, i.school, i.telephone, i.post, i.key}
                        DataSet1.Tables(0).Rows.Add(temp_arr)
                    End If
                Next
            Next
        End If
        DataGridView1.Columns.Item(0).Visible = False
        DataGridView1.AutoResizeColumn(2)
        DataGridView1.AutoResizeColumn(3)
        DataGridView1.AutoResizeColumn(4)
        DataGridView1.AutoResizeColumn(5)
        DataGridView1.AutoResizeColumn(6)
        For Each i As DataGridViewRow In DataGridView1.SelectedRows
            i.Selected = False
        Next
    End Sub
    Dim DB1Analyser_end_err, DB1Analyser_copy_db As Boolean
    Private Sub ИзменитьToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ИзменитьToolStripMenuItem1.Click
        If DataGridView1.SelectedRows.Count = 1 Then
            Add_Redact_db_data.work_mode = 2
            Add_Redact_db_data.ShowDialog()
        ElseIf DataGridView1.SelectedRows.Count = 0 Then
            MsgBox("Вы не выбрали элемент для изменения!",, "Ошибка")
        Else
            MsgBox("Вы выбрали слишком много элементов для изменения!",, "Ошибка")
        End If
    End Sub
    Private Sub ПросмотрToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПросмотрToolStripMenuItem.Click
        If DataGridView1.SelectedRows.Count = 1 Then
            Add_Redact_db_data.work_mode = 3
            Add_Redact_db_data.ShowDialog()
        ElseIf DataGridView1.SelectedRows.Count = 0 Then
            MsgBox("Вы не выбрали элемент для просмотра!",, "Ошибка")
        Else
            MsgBox("Вы выбрали слишком много элементов для просмотра!",, "Ошибка")
        End If
    End Sub
    Private Sub Button30_Click(sender As Object, e As EventArgs) Handles Button30.Click
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim result As MsgBoxResult = MsgBox("Вы действительно хотите удалить выбранных участников из базы данных?", MsgBoxStyle.YesNo, "Сообщение")
            If result = MsgBoxResult.Yes Then
                Dim indexes(-1) As String
                For Each i As DataGridViewRow In DataGridView1.SelectedRows
                    Array.Resize(indexes, indexes.Length + 1)
                    indexes(indexes.Length - 1) = i.Cells.Item(0).Value
                Next
                For i As Short = 0 To indexes.Length - 1
                    Dim pos As Short = -1
                    For i1 As Short = 0 To db_data.Count - 1
                        If db_data(i1).key = indexes(i) Then
                            pos = i1
                            Exit For
                        End If
                    Next
                    If db_data(pos).photo <> "" Then
                        If My.Computer.FileSystem.FileExists(Main.TEMP_WAY & "\MCD-temp-47591795970905\database\" & db_data(pos).photo & ".jpg") Then
                            My.Computer.FileSystem.DeleteFile(Main.TEMP_WAY & "\MCD-temp-47591795970905\database\" & db_data(pos).photo & ".jpg")
                        End If
                    End If
                    If search_mode = False Then
                        'Renovate datagrid
                        For i1 As Short = 0 To sdb_keys.Count - 1
                            If sdb_keys(i1) = db_data(pos).key Then
                                sdb_keys.RemoveAt(i1)
                                Exit For
                            End If
                        Next
                        DataSet1.Tables(0).Rows.RemoveAt(pos)
                        DataGridView1.Columns.Item(0).Visible = False
                        DataGridView1.AutoResizeColumn(2)
                        DataGridView1.AutoResizeColumn(3)
                        DataGridView1.AutoResizeColumn(4)
                        DataGridView1.AutoResizeColumn(5)
                        DataGridView1.AutoResizeColumn(6)
                    Else
                        sdb_keys.Remove(db_data(pos).key)
                        For i1 As Short = 0 To db_data.Count - 1
                            If db_data(pos).key = DataSet1.Tables(0).Rows(i1).Item(6) Then
                                DataSet1.Tables(0).Rows.RemoveAt(i1)
                                Exit For
                            End If
                        Next
                        DataGridView1.Columns.Item(0).Visible = False
                        DataGridView1.AutoResizeColumn(2)
                        DataGridView1.AutoResizeColumn(3)
                        DataGridView1.AutoResizeColumn(4)
                        DataGridView1.AutoResizeColumn(5)
                        DataGridView1.AutoResizeColumn(6)
                    End If
                    db_data.RemoveAt(pos)
                    dbf_collector.participators.Remove(indexes(i))
                    DB_WriteDataFile(indexes(i), True)
                Next
                DB_WriteCollectorFile()
                If CheckBox1.Checked Then
                    Button_save_db()
                Else
                    db_not_saved = True
                    Help.SetToolTip(PictureBox2, "Изменения в базе данных не сохранены.")
                    PictureBox2.Image = My.Resources.not_saved
                End If
            End If
        Else
            MsgBox("Вы не выбрали элементы для удаления!",, "Ошибка")
        End If
    End Sub
    Private Sub УдалитьToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles УдалитьToolStripMenuItem1.Click
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim result As MsgBoxResult = MsgBox("Вы действительно хотите удалить выбранных участников из базы данных?", MsgBoxStyle.YesNo, "Сообщение")
            If result = MsgBoxResult.Yes Then
                Dim indexes(-1) As String
                For Each i As DataGridViewRow In DataGridView1.SelectedRows
                    Array.Resize(indexes, indexes.Length + 1)
                    indexes(indexes.Length - 1) = i.Cells.Item(0).Value
                Next
                For i As Short = 0 To indexes.Length - 1
                    Dim pos As Short = -1
                    For i1 As Short = 0 To db_data.Count - 1
                        If db_data(i1).key = indexes(i) Then
                            pos = i1
                            Exit For
                        End If
                    Next
                    If db_data(pos).photo <> "" Then
                        If My.Computer.FileSystem.FileExists(Main.TEMP_WAY & "\MCD-temp-47591795970905\database\" & db_data(pos).photo & ".jpg") Then
                            My.Computer.FileSystem.DeleteFile(Main.TEMP_WAY & "\MCD-temp-47591795970905\database\" & db_data(pos).photo & ".jpg")
                        End If
                    End If
                    If search_mode = False Then
                        'Renovate datagrid
                        For i1 As Short = 0 To sdb_keys.Count - 1
                            If sdb_keys(i1) = db_data(pos).key Then
                                sdb_keys.RemoveAt(i1)
                                Exit For
                            End If
                        Next
                        DataSet1.Tables(0).Rows.RemoveAt(pos)
                        DataGridView1.Columns.Item(0).Visible = False
                        DataGridView1.AutoResizeColumn(2)
                        DataGridView1.AutoResizeColumn(3)
                        DataGridView1.AutoResizeColumn(4)
                        DataGridView1.AutoResizeColumn(5)
                        DataGridView1.AutoResizeColumn(6)
                    Else
                        sdb_keys.Remove(db_data(pos).key)
                        For i1 As Short = 0 To db_data.Count - 1
                            If db_data(pos).key = DataSet1.Tables(0).Rows(i1).Item(6) Then
                                DataSet1.Tables(0).Rows.RemoveAt(i1)
                                Exit For
                            End If
                        Next
                        DataGridView1.Columns.Item(0).Visible = False
                        DataGridView1.AutoResizeColumn(2)
                        DataGridView1.AutoResizeColumn(3)
                        DataGridView1.AutoResizeColumn(4)
                        DataGridView1.AutoResizeColumn(5)
                        DataGridView1.AutoResizeColumn(6)
                    End If
                    db_data.RemoveAt(pos)
                    dbf_collector.participators.Remove(indexes(i))
                    DB_WriteDataFile(indexes(i), True)
                Next
                DB_WriteCollectorFile()
                If CheckBox1.Checked Then
                    Button_save_db()
                Else
                    db_not_saved = True
                    Help.SetToolTip(PictureBox2, "Изменения в базе данных не сохранены.")
                    PictureBox2.Image = My.Resources.not_saved
                End If
            End If
        Else
            MsgBox("Вы не выбрали элеметы для удаления!",, "Ошибка")
        End If
    End Sub
    Private Sub DB1Analyser_DoWork(sender As Object, e As DoWorkEventArgs) Handles DB1Analyser.DoWork
        Try
            If DB1Analyser_copy_db = True Then Error 1
            Dim info_file(), collector_file() As String
            info_file = IO.File.ReadAllLines(TEMP_WAY & "\MCD-temp-47591795970905\database\info.ini")
            collector_file = IO.File.ReadAllLines(TEMP_WAY & "\MCD-temp-47591795970905\database\collector.ini")
            dbf_info = read_info_file(info_file)
            If dbf_info.format <> "LmsI-CLrw-6o2m1i" Then Error 1
            'Clear the Data Base from trash
            Dim db_data_files() As String = IO.Directory.GetDirectories(TEMP_WAY & "\MCD-temp-47591795970905\database")
            For Each i As String In db_data_files
                My.Computer.FileSystem.DeleteDirectory(i, FileIO.DeleteDirectoryOption.DeleteAllContents)
            Next
            db_data_files = IO.Directory.GetFiles(TEMP_WAY & "\MCD-temp-47591795970905\database\", "*", IO.SearchOption.TopDirectoryOnly)
            For Each i As String In db_data_files
                Dim file_extension As String = My.Computer.FileSystem.GetFileInfo(i).Extension
                If file_extension <> ".txt" And file_extension <> ".jpg" Then
                    file_extension = My.Computer.FileSystem.GetFileInfo(i).Name
                    If file_extension <> "info.ini" And file_extension <> "collector.ini" Then
                        My.Computer.FileSystem.DeleteFile(i)
                    End If
                End If
            Next
            'End clear the Data Base from trash
            Dim readed_collector As dbs_readed_file = read_file(collector_file)
            If readed_collector.names.Count <> 2 Then Error 1
            Dim finded() As Boolean = {False, False, False, False, False}, s_index As Short = 0
            For i As Short = 0 To 1
                If readed_collector.names(i) = "ADB info" And readed_collector.pages(i).keys.Count = 5 Then
                    For i1 As Short = 0 To 4
                        If readed_collector.pages(i).keys(i1) = "LastReadTime" And readed_collector.pages(i).params(i1).params.Length = 1 Then
                            dbf_collector.LastReadTime = readed_collector.pages(i).params(i1).params(0)
                            finded(0) = True
                        ElseIf readed_collector.pages(i).keys(i1) = "LastWriteTime" And readed_collector.pages(i).params(i1).params.Length = 1 Then
                            dbf_collector.LastWriteTime = readed_collector.pages(i).params(i1).params(0)
                            finded(1) = True
                        ElseIf readed_collector.pages(i).keys(i1) = "progID" And readed_collector.pages(i).params(i1).params.Length = 1 Then
                            dbf_collector.progID = readed_collector.pages(i).params(i1).params(0)
                            finded(2) = True
                        ElseIf readed_collector.pages(i).keys(i1) = "userName" And readed_collector.pages(i).params(i1).params.Length = 1 Then
                            dbf_collector.user_name = readed_collector.pages(i).params(i1).params(0)
                            finded(3) = True
                        ElseIf readed_collector.pages(i).keys(i1) = "organization" And readed_collector.pages(i).params(i1).params.Length = 1 Then
                            dbf_collector.organization = readed_collector.pages(i).params(i1).params(0)
                            finded(4) = True
                        End If
                    Next
                    Exit For
                End If
            Next
            If finded(0) = False Or finded(1) = False Or finded(2) = False Or finded(3) = False Or finded(4) = False Then Error 1
            finded(0) = False
            For i As Short = 0 To 1
                If readed_collector.names(i) = "participators" Then
                    s_index = i
                    finded(0) = True
                    Exit For
                End If
            Next
            If finded(0) = False Then Error 1
            For i As Short = 0 To readed_collector.pages(s_index).keys.Count - 1
                If db_data_collector_exist(readed_collector.pages(s_index).keys(i)) = False Then
                    If My.Computer.FileSystem.FileExists(TEMP_WAY & "\MCD-temp-47591795970905\database\" & readed_collector.pages(s_index).keys(i) & ".txt") Then
                        dbf_collector.participators.Add(readed_collector.pages(s_index).keys(i))
                    End If
                Else
                    Error 1
                End If
            Next
            'Clear the Data Base from trash
            db_data_files = IO.Directory.GetFiles(TEMP_WAY & "\MCD-temp-47591795970905\database\", "*.txt", IO.SearchOption.TopDirectoryOnly)
            For i As Short = 0 To db_data_files.Length - 1
                If db_data_collector_exist(Mid(My.Computer.FileSystem.GetFileInfo(db_data_files(i)).Name, 1, My.Computer.FileSystem.GetFileInfo(db_data_files(i)).Name.Length - 4)) = False Then
                    My.Computer.FileSystem.DeleteFile(db_data_files(i))
                End If
            Next
            'End clear the Data Base from trash
            Dim data_to_delete, img_no_delete As New ArrayList
            For i As Short = 0 To dbf_collector.participators.Count - 1
                db_data_files = IO.File.ReadAllLines(TEMP_WAY & "\MCD-temp-47591795970905\database\" & dbf_collector.participators(i) & ".txt")
                Dim readed_data_file As dbs_readed_file = read_file(db_data_files)
                If readed_data_file.names.Count = 2 Then
                    Dim rtf_participator As New db_Participator, finded_arr() As Boolean = {False, False, False, False, False, False, False, False, False}
                    rtf_participator.key = dbf_collector.participators(i)
                    For i1 As Short = 0 To 1
                        If readed_data_file.names(i1) = "one line" And readed_data_file.pages(i1).keys.Count = 7 Then
                            For i2 As Short = 0 To 6
                                If readed_data_file.pages(i1).keys(i2) = "name" And readed_data_file.pages(i1).params(i2).params.Length = 1 Then
                                    rtf_participator.name = readed_data_file.pages(i1).params(i2).params(0)
                                    finded_arr(0) = True
                                ElseIf readed_data_file.pages(i1).keys(i2) = "year" And readed_data_file.pages(i1).params(i2).params.Length = 1 Then
                                    rtf_participator.year = readed_data_file.pages(i1).params(i2).params(0)
                                    finded_arr(1) = True
                                ElseIf readed_data_file.pages(i1).keys(i2) = "school" And readed_data_file.pages(i1).params(i2).params.Length = 1 Then
                                    rtf_participator.school = readed_data_file.pages(i1).params(i2).params(0)
                                    finded_arr(2) = True
                                ElseIf readed_data_file.pages(i1).keys(i2) = "class" And readed_data_file.pages(i1).params(i2).params.Length = 1 Then
                                    rtf_participator.class_ = readed_data_file.pages(i1).params(i2).params(0)
                                    finded_arr(3) = True
                                ElseIf readed_data_file.pages(i1).keys(i2) = "telephone" And readed_data_file.pages(i1).params(i2).params.Length = 1 Then
                                    rtf_participator.telephone = readed_data_file.pages(i1).params(i2).params(0)
                                    finded_arr(4) = True
                                ElseIf readed_data_file.pages(i1).keys(i2) = "post" And readed_data_file.pages(i1).params(i2).params.Length = 1 Then
                                    rtf_participator.post = readed_data_file.pages(i1).params(i2).params(0)
                                    finded_arr(5) = True
                                ElseIf readed_data_file.pages(i1).keys(i2) = "photo" And readed_data_file.pages(i1).params(i2).params.Length = 1 Then
                                    rtf_participator.photo = readed_data_file.pages(i1).params(i2).params(0)
                                    If rtf_participator.photo <> "" Then
                                        If My.Computer.FileSystem.FileExists(TEMP_WAY & "\MCD-temp-47591795970905\database\" & readed_data_file.pages(i1).params(i2).params(0) & ".jpg") Then
                                            Dim img_stream As IO.Stream = IO.File.OpenRead(TEMP_WAY & "\MCD-temp-47591795970905\database\" & readed_data_file.pages(i1).params(i2).params(0) & ".jpg")
                                            rtf_participator.img = New Bitmap(Image.FromStream(img_stream))
                                            img_stream.Close()
                                            img_no_delete.Add(readed_data_file.pages(i1).params(i2).params(0))
                                        Else
                                            rtf_participator.photo = ""
                                            rtf_participator.img = My.Resources.no_photo
                                        End If
                                    Else
                                        rtf_participator.img = My.Resources.no_photo
                                    End If
                                    finded_arr(6) = True
                                End If
                            Next
                        ElseIf readed_data_file.names(i1) = "multi line" And readed_data_file.pages(i1).keys.Count = 2 Then
                            For i2 As Short = 0 To 1
                                If readed_data_file.pages(i1).keys(i2) = "skills" Then
                                    rtf_participator.skills = readed_data_file.pages(i1).params(i2).params
                                    finded_arr(7) = True
                                ElseIf readed_data_file.pages(i1).keys(i2) = "achievements" Then
                                    rtf_participator.achievements = readed_data_file.pages(i1).params(i2).params
                                    finded_arr(8) = True
                                End If
                            Next
                        End If
                    Next
                    finded(1) = True
                    For i1 As Short = 0 To 8
                        If finded_arr(i1) = False Then
                            finded(1) = False
                        End If
                    Next
                    If finded(1) = True Then
                        db_data.Add(rtf_participator)
                    Else
                        data_to_delete.Add(rtf_participator.key)
                    End If
                End If
            Next
            'Clear the Data Base from trash
            For i As Short = 0 To data_to_delete.Count - 1
                My.Computer.FileSystem.DeleteFile(TEMP_WAY & "\MCD-temp-47591795970905\database\" & data_to_delete(i) & ".txt")
            Next
            db_data_files = IO.Directory.GetFiles(TEMP_WAY & "\MCD-temp-47591795970905\database\", "*.jpg", IO.SearchOption.TopDirectoryOnly)
            For i As Short = 0 To db_data_files.Length - 1
                finded(0) = False
                For i1 As Short = 0 To img_no_delete.Count - 1
                    If Mid(My.Computer.FileSystem.GetFileInfo(db_data_files(i)).Name, 1, My.Computer.FileSystem.GetFileInfo(db_data_files(i)).Name.Length - 4) = img_no_delete(i1) Then finded(0) = True
                Next
                If finded(0) = False Then My.Computer.FileSystem.DeleteFile(db_data_files(i))
            Next
            'End clear the Data Base from trash
            'Analysing ended correctly
            If dbf_info.versions.Length = 0 Then
                Array.Resize(dbf_info.versions, 1)
                dbf_info.versions(0) = MY_VERSION
                DB_WriteInfoFile()
            Else
                If dbf_info.versions(dbf_info.versions.Length - 1) <> MY_VERSION Then
                    Array.Resize(dbf_info.versions, dbf_info.versions.Length + 1)
                    dbf_info.versions(dbf_info.versions.Length - 1) = MY_VERSION
                    DB_WriteInfoFile()
                End If
            End If
            dbf_collector.LastReadTime = Get_date_string(My.Computer.Clock.LocalTime)
            If dbf_collector.progID = info_set(2) Then
                DB_WriteCollectorFile()
            End If
            DB1Analyser_end_err = False
        Catch
            Dim hdb_time As String = Format(My.Computer.Clock.LocalTime.Day, "00") & "." & Format(My.Computer.Clock.LocalTime.Month, "00") & "." & Format(My.Computer.Clock.LocalTime.Year, "0000")
            If My.Computer.FileSystem.DirectoryExists(MY_WAY & "\harmed databases") Then
                If My.Computer.FileSystem.FileExists(MY_WAY & "\harmed databases\" & hdb_time & ".dat") Then
                    Dim selected_num As Short = 1, file_exist As Boolean = True
                    While file_exist = True
                        If My.Computer.FileSystem.FileExists(MY_WAY & "\harmed databases\" & hdb_time & "-" & selected_num & ".dat") = False Then
                            file_exist = False
                        Else
                            selected_num += 1
                        End If
                    End While
                    My.Computer.FileSystem.CopyFile(MY_WAY & "\database.dat", MY_WAY & "\harmed databases\" & hdb_time & "-" & selected_num & ".dat")
                Else
                    My.Computer.FileSystem.CopyFile(MY_WAY & "\database.dat", MY_WAY & "\harmed databases\" & hdb_time & ".dat")
                End If
            Else
                My.Computer.FileSystem.CreateDirectory(MY_WAY & "\harmed databases")
                My.Computer.FileSystem.CopyFile(MY_WAY & "\database.dat", MY_WAY & "\harmed databases\" & hdb_time & ".dat")
            End If
            My.Computer.FileSystem.DeleteDirectory(TEMP_WAY & "\MCD-temp-47591795970905\database", FileIO.DeleteDirectoryOption.DeleteAllContents)
            DB1Analyser_end_err = True
        End Try
    End Sub
    Private Sub DB1Analyser_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles DB1Analyser.RunWorkerCompleted
        If DB1Analyser_end_err = True Then
            If DB1Analyser_copy_db = False Then
                If dbf_info.format <> "LmsI-CLrw-6o2m1i" And dbf_info.format <> "" Then
                    MsgBox("База данных Участников имеет несовместимый с данной версией программы MCD формат!" & Chr(10) & "MCD сделал резервную копию имеющейся базы данных Участников в папке: """ & MY_WAY & "\harmed databases""" & Chr(10) & "База данных Участников будет пересоздана.",, "Ошибка")
                Else
                    MsgBox("База данных Участников повреждена!" & Chr(10) & "MCD сделал резервную копию имеющейся базы данных Участников в папке: """ & MY_WAY & "\harmed databases""" & Chr(10) & "База данных Участников будет пересоздана.",, "Ошибка")
                End If
            End If
            Messanger.Renovate_messange("db loading", "Идёт загрузка базы данных Участников...")
            Messanger.Add_messange("База данных Участников будет пересоздана.", "", True)
            My.Computer.FileSystem.CreateDirectory(TEMP_WAY & "\MCD-temp-47591795970905\database")
            Dim info_file(), collector_file() As String
            info_file = {"[format]", "LmsI-CLrw-6o2m1i", "", "[versions]", MY_VERSION}
            collector_file = {"[ADB info]", "LastReadTime=" & Get_date_string(My.Computer.Clock.LocalTime), "LastWriteTime=" & Get_date_string(My.Computer.Clock.LocalTime), "progID=" & info_set(2), "userName=" & info_set(0), "organization=" & info_set(1), "", "[participators]"}
            IO.File.WriteAllLines(TEMP_WAY & "\MCD-temp-47591795970905\database\info.ini", info_file)
            IO.File.WriteAllLines(TEMP_WAY & "\MCD-temp-47591795970905\database\collector.ini", collector_file)
            dbf_collector = New db_collector_file
            db_data.Clear()
            cmd_service_wait_proc = Shell(MY_WAY & "\mcdCAnc.exe ""MCDmode:archiver,Code:2801"" ""compress"" """ & MY_WAY & "\database.dat"" """ & TEMP_WAY & "\MCD-temp-47591795970905\database""")
            cmd_service_wait_mode = 2
            cmd_service_wait.Start()
        Else
            If dbf_collector.progID <> info_set(2) And dbf_collector.progID <> "" Then
                Dim result As MsgBoxResult = MsgBox("Текущая база данных Участников была создана другой копией программы MCD." & Chr(10) & "Ранее эта база данных принадлежала: " & Chr(10) & "  Пользователю: " & dbf_collector.user_name & Chr(10) & "  Организации: " & dbf_collector.organization & Chr(10) & Chr(10) & "Вы хотите продолжить загрузку этой базы данных?" & Chr(10) & "---- Если Вы откажитесь от загрузки базы данных, она будет скопирована в папку: """ & MY_WAY & "\harmed databases"". Будет создана новая база данных Участников. ----", MsgBoxStyle.YesNo, "Сообщение")
                If result = MsgBoxResult.Yes Then
                    dbf_collector.progID = info_set(2)
                    DB_WriteCollectorFile()
                    DB_WriteDataBase(3, False)
                Else
                    DB1Analyser_copy_db = True
                    DB1Analyser.RunWorkerAsync()
                End If
            Else
                If dbf_collector.progID = "" Then
                    dbf_collector.progID = info_set(2)
                    DB_WriteCollectorFile()
                End If
                DB_WriteDataBase(3, True)
            End If
        End If
    End Sub
    Private Sub DB_WriteInfoFile()
        Dim info_file() As String = {"[format]", dbf_info.format, "", "[versions]"}
        For Each i As String In dbf_info.versions
            Array.Resize(info_file, info_file.Length + 1)
            info_file(info_file.Length - 1) = i
        Next
        IO.File.WriteAllLines(TEMP_WAY & "\MCD-temp-47591795970905\database\info.ini", info_file)
    End Sub
    Private Sub ДобавитьКРаспределениюToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ДобавитьКРаспределениюToolStripMenuItem.Click
        If DataGridView1.SelectedRows.Count > 0 Then
            Member_to_distribute.ShowDialog()
        Else
            MsgBox("Вы не участников для добавления!",, "Ошибка")
        End If
    End Sub
    Private Sub СохранитьАнкетуToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles СохранитьАнкетуToolStripMenuItem.Click
        If DataGridView1.SelectedRows.Count > 0 Then
            Member_to_word.ShowDialog()
        Else
            MsgBox("Вы не выбрали элементы для экспорта в Word!",, "Ошибка")
        End If
    End Sub
    Private Sub ДублироватьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ДублироватьToolStripMenuItem.Click
        If DataGridView1.SelectedRows.Count > 0 Then
            Membars_double.ShowDialog()
        Else
            MsgBox("Вы не выбрали элементы для дублирования!",, "Ошибка")
        End If
    End Sub
    Private Sub DataGridView1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles DataGridView1.MouseDoubleClick
        If DataGridView1.SelectedRows.Count = 1 Then
            Add_Redact_db_data.work_mode = 3
            Add_Redact_db_data.ShowDialog()
        ElseIf DataGridView1.SelectedRows.Count = 0 Then
            MsgBox("Вы не выбрали элемент для просмотра!",, "Ошибка")
        Else
            MsgBox("Вы выбрали слишком много элементов для просмотра!",, "Ошибка")
        End If
    End Sub
    Public Sub DB_WriteCollectorFile()
        Dim collector_file() As String = {"[ADB info]", "LastReadTime=" & dbf_collector.LastReadTime, "LastWriteTime=" & dbf_collector.LastWriteTime, "progID=" & info_set(2), "userName=" & info_set(0), "organization=" & info_set(1), "", "[participators]"}
        For Each i As String In dbf_collector.participators
            Array.Resize(collector_file, collector_file.Length + 1)
            collector_file(collector_file.Length - 1) = i & "=1"
        Next
        IO.File.WriteAllLines(TEMP_WAY & "\MCD-temp-47591795970905\database\collector.ini", collector_file)
    End Sub
    Public Sub DB_WriteDataFile(key As String, delete As Boolean)
        If delete = False Then
            Dim pos As Short = -1
            For i As Short = 0 To db_data.Count - 1
                If db_data(i).key = key Then
                    pos = i
                    Exit For
                End If
            Next
            If pos > -1 Then
                Dim data_file() As String = {"[one line]", "name=" & db_data(pos).name, "year=" & db_data(pos).year, "school=" & db_data(pos).school, "class=" & db_data(pos).class_, "telephone=" & db_data(pos).telephone, "post=" & db_data(pos).post, "photo=" & db_data(pos).photo, "", "[multi line]"}
                If db_data(pos).skills.Length = 1 Then
                    Array.Resize(data_file, data_file.Length + 1)
                    data_file(data_file.Length - 1) = "skills=" & db_data(pos).skills(0)
                Else
                    Array.Resize(data_file, data_file.Length + 1)
                    data_file(data_file.Length - 1) = "skills=<" & db_data(pos).skills(0)
                    For i As Short = 1 To db_data(pos).skills.Length - 1
                        If i = db_data(pos).skills.Length - 1 Then
                            Array.Resize(data_file, data_file.Length + 1)
                            data_file(data_file.Length - 1) = db_data(pos).skills(i) & ">"
                        Else
                            Array.Resize(data_file, data_file.Length + 1)
                            data_file(data_file.Length - 1) = db_data(pos).skills(i)
                        End If
                    Next
                End If
                If db_data(pos).achievements.Length = 1 Then
                    Array.Resize(data_file, data_file.Length + 1)
                    data_file(data_file.Length - 1) = "achievements=" & db_data(pos).achievements(0)
                Else
                    Array.Resize(data_file, data_file.Length + 1)
                    data_file(data_file.Length - 1) = "achievements=<" & db_data(pos).achievements(0)
                    For i As Short = 1 To db_data(pos).achievements.Length - 1
                        If i = db_data(pos).achievements.Length - 1 Then
                            Array.Resize(data_file, data_file.Length + 1)
                            data_file(data_file.Length - 1) = db_data(pos).achievements(i) & ">"
                        Else
                            Array.Resize(data_file, data_file.Length + 1)
                            data_file(data_file.Length - 1) = db_data(pos).achievements(i)
                        End If
                    Next
                End If
                IO.File.WriteAllLines(TEMP_WAY & "\MCD-temp-47591795970905\database\" & key & ".txt", data_file)
            End If
        Else
            My.Computer.FileSystem.DeleteFile(TEMP_WAY & "\MCD-temp-47591795970905\database\" & key & ".txt")
        End If
    End Sub
    Public Sub DB_WriteDataBase(scenary As Short, system_mode As Boolean)
        If system_mode = False Then dbf_collector.LastWriteTime = Get_date_string(My.Computer.Clock.LocalTime)
        cmd_service_wait_mode = scenary
        DB1Background_save.RunWorkerAsync()
    End Sub
    Private Sub DB1Background_save_DoWork(sender As Object, e As DoWorkEventArgs) Handles DB1Background_save.DoWork
        DB_WriteCollectorFile()
    End Sub
    Private Sub DB1Background_save_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles DB1Background_save.RunWorkerCompleted
        cmd_service_wait_proc = Shell(MY_WAY & "\mcdCAnc.exe ""MCDmode:archiver,Code:2801"" ""compress"" """ & MY_WAY & "\database.dat"" """ & TEMP_WAY & "\MCD-temp-47591795970905\database""")
        cmd_service_wait.Start()
    End Sub
    Private Function read_info_file(line_arr() As String) As db_info_file
        Dim ret_readed_file As New db_info_file
        ret_readed_file.format = "null"
        For i As Short = 0 To 1
            If i = 0 Then
                For i1 As Short = 0 To line_arr.Length - 1
                    If line_arr(i1).ToLower = "[format]" Then
                        For i2 As Short = i1 + 1 To line_arr.Length - 1
                            If line_arr(i2).ToLower = "[versions]" Then
                                Exit For
                            ElseIf line_arr(i2).TrimStart(" ") <> "" Then
                                ret_readed_file.format = line_arr(i2)
                                Exit For
                            End If
                        Next
                        Exit For
                    End If
                Next
            ElseIf i = 1 Then
                For i1 As Short = 0 To line_arr.Length - 1
                    If line_arr(i1).ToLower = "[versions]" Then
                        For i2 As Short = i1 + 1 To line_arr.Length - 1
                            If line_arr(i2).ToLower = "[format]" Then
                                Exit For
                            ElseIf line_arr(i2).TrimStart(" ") <> "" Then
                                Array.Resize(ret_readed_file.versions, ret_readed_file.versions.Length + 1)
                                ret_readed_file.versions(ret_readed_file.versions.Length - 1) = line_arr(i2)
                            End If
                        Next
                        Exit For
                    End If
                Next
            End If
        Next
        Return ret_readed_file
    End Function
    Private Function read_file(line_arr() As String) As dbs_readed_file
        Dim ret_readed_file As New dbs_readed_file
        For i As Short = 0 To line_arr.Length - 1
            If line_arr(i) = "[" & line_arr(i).Trim("[", "]") & "]" And Banned_simvols_test(line_arr(i).Trim("[", "]")) Then
                Dim multiline_read As Boolean = False, params_temp(-1) As String
                Dim rfp As New dbs_page, harmed As Boolean = False
                For i1 As Short = i + 1 To line_arr.Length - 1
                    If multiline_read Then
                        If Banned_simvols_test(line_arr(i1)) And i1 = line_arr.Length - 1 Then
                            harmed = True
                            Exit For
                        ElseIf Banned_simvols_test(line_arr(i1)) Then
                            Array.Resize(params_temp, params_temp.Length + 1)
                            params_temp(params_temp.Length - 1) = line_arr(i1)
                        Else
                            If line_arr(i1) = line_arr(i1).TrimEnd(">") & ">" Then
                                If Banned_simvols_test(line_arr(i1).TrimEnd(">")) Then
                                    Array.Resize(params_temp, params_temp.Length + 1)
                                    params_temp(params_temp.Length - 1) = line_arr(i1).TrimEnd(">")
                                    rfp.params.Add(New dbs_param(params_temp))
                                    multiline_read = False
                                Else
                                    harmed = True
                                    Exit For
                                End If
                            Else
                                harmed = True
                                Exit For
                            End If
                        End If
                    Else
                        If line_arr(i1) = "" Then
                            Exit For
                        ElseIf line_arr(i1).IndexOf("=") > -1 Then
                            Dim spl_parts() As String = line_arr(i1).Split("=")
                            If spl_parts.Length = 2 Then
                                If Banned_simvols_test(spl_parts(0)) Then
                                    rfp.keys.Add(spl_parts(0))
                                    If spl_parts(1) = "<" & spl_parts(1).TrimStart("<") Then
                                        If spl_parts(1) = spl_parts(1).TrimEnd(">") & ">" Then
                                            If Banned_simvols_test(spl_parts(1).Trim("<", ">")) Then
                                                rfp.params.Add(New dbs_param({spl_parts(1).Trim("<", ">")}))
                                            Else
                                                harmed = True
                                                Exit For
                                            End If
                                        Else
                                            If Banned_simvols_test(spl_parts(1).TrimStart("<")) Then
                                                Array.Resize(params_temp, 1)
                                                params_temp(0) = spl_parts(1).TrimStart("<")
                                                multiline_read = True
                                            Else
                                                harmed = True
                                                Exit For
                                            End If
                                        End If
                                    Else
                                        If Banned_simvols_test(spl_parts(1)) Then
                                            rfp.params.Add(New dbs_param({spl_parts(1)}))
                                        Else
                                            harmed = True
                                            Exit For
                                        End If
                                    End If
                                Else
                                    harmed = True
                                    Exit For
                                End If
                            Else
                                harmed = True
                                Exit For
                            End If
                        Else
                            harmed = True
                            Exit For
                        End If
                    End If
                Next
                If harmed = False Then
                    ret_readed_file.names.Add(line_arr(i).Trim("[", "]"))
                    ret_readed_file.pages.Add(rfp)
                End If
            End If
        Next
        Return ret_readed_file
    End Function
    Public Function db_data_collector_exist(name As String) As Boolean
        Dim exist As Boolean = False
        For i As Short = 0 To dbf_collector.participators.Count - 1
            If dbf_collector.participators(i) = name Then
                exist = True
                Exit For
            End If
        Next
        Return exist
    End Function
    Private Class dbs_readed_file
        Public names As New ArrayList, pages As New List(Of dbs_page)
    End Class
    Private Class dbs_page
        Public keys As New ArrayList, params As New List(Of dbs_param)
    End Class
    Private Class dbs_param
        Sub New(params() As String)
            Me.params = params
        End Sub
        Public params() As String
    End Class
    Dim textboxes_text(7) As String, radiobuttons_value(41) As Boolean, checkboxes_value(7) As Boolean
    Public Sub Search_start()
        textboxes_text(0) = Search_db.TextBox1.Text
        textboxes_text(1) = Search_db.TextBox2.Text
        textboxes_text(2) = Search_db.TextBox3.Text
        textboxes_text(3) = Search_db.TextBox4.Text
        textboxes_text(4) = Search_db.TextBox5.Text
        textboxes_text(5) = Search_db.TextBox6.Text
        textboxes_text(6) = Search_db.TextBox7.Text
        textboxes_text(7) = Search_db.TextBox8.Text
        checkboxes_value(0) = Search_db.CheckBox1.Checked
        checkboxes_value(1) = Search_db.CheckBox2.Checked
        checkboxes_value(2) = Search_db.CheckBox3.Checked
        checkboxes_value(3) = Search_db.CheckBox4.Checked
        checkboxes_value(4) = Search_db.CheckBox5.Checked
        checkboxes_value(5) = Search_db.CheckBox6.Checked
        checkboxes_value(6) = Search_db.CheckBox7.Checked
        checkboxes_value(7) = Search_db.CheckBox8.Checked
        radiobuttons_value(0) = Search_db.RadioButton1.Checked
        radiobuttons_value(1) = Search_db.RadioButton2.Checked
        radiobuttons_value(2) = Search_db.RadioButton3.Checked
        radiobuttons_value(3) = Search_db.RadioButton4.Checked
        radiobuttons_value(4) = Search_db.RadioButton5.Checked
        radiobuttons_value(5) = Search_db.RadioButton6.Checked
        radiobuttons_value(6) = Search_db.RadioButton7.Checked
        radiobuttons_value(7) = Search_db.RadioButton8.Checked
        radiobuttons_value(8) = Search_db.RadioButton9.Checked
        radiobuttons_value(9) = Search_db.RadioButton10.Checked
        radiobuttons_value(10) = Search_db.RadioButton11.Checked
        radiobuttons_value(11) = Search_db.RadioButton12.Checked
        radiobuttons_value(12) = Search_db.RadioButton13.Checked
        radiobuttons_value(13) = Search_db.RadioButton14.Checked
        radiobuttons_value(14) = Search_db.RadioButton15.Checked
        radiobuttons_value(15) = Search_db.RadioButton16.Checked
        radiobuttons_value(16) = Search_db.RadioButton17.Checked
        radiobuttons_value(17) = Search_db.RadioButton18.Checked
        radiobuttons_value(18) = Search_db.RadioButton19.Checked
        radiobuttons_value(19) = Search_db.RadioButton20.Checked
        radiobuttons_value(20) = Search_db.RadioButton21.Checked
        radiobuttons_value(21) = Search_db.RadioButton22.Checked
        radiobuttons_value(22) = Search_db.RadioButton23.Checked
        radiobuttons_value(23) = Search_db.RadioButton24.Checked
        radiobuttons_value(24) = Search_db.RadioButton25.Checked
        radiobuttons_value(25) = Search_db.RadioButton26.Checked
        radiobuttons_value(26) = Search_db.RadioButton27.Checked
        radiobuttons_value(27) = Search_db.RadioButton28.Checked
        radiobuttons_value(28) = Search_db.RadioButton29.Checked
        radiobuttons_value(29) = Search_db.RadioButton30.Checked
        radiobuttons_value(30) = Search_db.RadioButton31.Checked
        radiobuttons_value(31) = Search_db.RadioButton32.Checked
        radiobuttons_value(32) = Search_db.RadioButton33.Checked
        radiobuttons_value(33) = Search_db.RadioButton34.Checked
        radiobuttons_value(34) = Search_db.RadioButton35.Checked
        radiobuttons_value(35) = Search_db.RadioButton36.Checked
        radiobuttons_value(36) = Search_db.RadioButton37.Checked
        radiobuttons_value(37) = Search_db.RadioButton38.Checked
        radiobuttons_value(38) = Search_db.RadioButton39.Checked
        radiobuttons_value(39) = Search_db.RadioButton40.Checked
        radiobuttons_value(40) = Search_db.RadioButton41.Checked
        radiobuttons_value(41) = Search_db.RadioButton42.Checked
        sdb_keys.Clear()
        Search_Elements.RunWorkerAsync()
    End Sub
    Private Sub Search_Elements_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles Search_Elements.DoWork
        For i As Short = 0 To db_data.Count - 1
            Dim results() As Boolean = {False, False, False, False, False, False, False, False}, no_result() As Boolean = {False, False, False, False, False, False, False, False}
            Dim splitter(-1) As String
            'Step 1
            If textboxes_text(0).IndexOf("\") > -1 Then
                splitter = textboxes_text(0).Split("\")
                Dim temp_next_info As Boolean = False
                For i1 As Short = 0 To splitter.Length - 1
                    If splitter(i1) <> "" Or radiobuttons_value(14) = True Then
                        If splitter(i1) = "" And radiobuttons_value(14) = True Then
                            temp_next_info = True
                            Exit For
                        Else
                            If db_data(i).name <> "" Then
                                temp_next_info = True
                                Exit For
                            End If
                        End If
                    End If
                Next
                If temp_next_info = True Then
                    Dim temp_res(-1) As Boolean
                    For i1 As Short = 0 To splitter.Length - 1
                        If splitter(i1) <> "" Then
                            Array.Resize(temp_res, temp_res.Length + 1)
                            temp_res(temp_res.Length - 1) = False
                            If checkboxes_value(0) Then
                                If radiobuttons_value(0) = True Then
                                    If db_data(i).name.IndexOf(splitter(i1)) > -1 Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                ElseIf radiobuttons_value(1) = True Then
                                    If splitter(i1).IndexOf(db_data(i).name) > -1 Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                ElseIf radiobuttons_value(14) = True Then
                                    If db_data(i).name = splitter(i1) Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                End If
                            Else
                                If radiobuttons_value(0) = True Then
                                    If db_data(i).name.ToLower.IndexOf(splitter(i1).ToLower) > -1 Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                ElseIf radiobuttons_value(1) = True Then
                                    If splitter(i1).ToLower.IndexOf(db_data(i).name.ToLower) > -1 Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                ElseIf radiobuttons_value(14) = True Then
                                    If db_data(i).name.ToLower = splitter(i1).ToLower Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                End If
                            End If
                        End If
                    Next
                    If radiobuttons_value(23) Then
                        results(0) = True
                        For Each i1 As Boolean In temp_res
                            If i1 <> True Then
                                results(0) = False
                                Exit For
                            End If
                        Next
                    Else
                        For Each i1 As Boolean In temp_res
                            If i1 = True Then
                                results(0) = True
                                Exit For
                            End If
                        Next
                    End If
                Else
                    results(0) = True
                    no_result(0) = True
                End If
            Else
                Dim temp_res As Boolean = False
                If textboxes_text(0) <> "" Or radiobuttons_value(14) = True Then
                    If textboxes_text(0) = "" And radiobuttons_value(14) = True Then
                        temp_res = True
                    Else
                        If db_data(i).name <> "" Then
                            temp_res = True
                        End If
                    End If
                End If
                If temp_res Then
                    If checkboxes_value(0) Then
                        If radiobuttons_value(0) = True Then
                            If db_data(i).name.IndexOf(textboxes_text(0)) > -1 Then
                                results(0) = True
                            End If
                        ElseIf radiobuttons_value(1) = True Then
                            If textboxes_text(0).IndexOf(db_data(i).name) > -1 Then
                                results(0) = True
                            End If
                        ElseIf radiobuttons_value(14) = True Then
                            If db_data(i).name = textboxes_text(0) Then
                                results(0) = True
                            End If
                        End If
                    Else
                        If radiobuttons_value(0) = True Then
                            If db_data(i).name.ToLower.IndexOf(textboxes_text(0).ToLower) > -1 Then
                                results(0) = True
                            End If
                        ElseIf radiobuttons_value(1) = True Then
                            If textboxes_text(0).ToLower.IndexOf(db_data(i).name.ToLower) > -1 Then
                                results(0) = True
                            End If
                        ElseIf radiobuttons_value(14) = True Then
                            If db_data(i).name.ToLower = textboxes_text(0).ToLower Then
                                results(0) = True
                            End If
                        End If
                    End If
                Else
                    results(0) = True
                    no_result(0) = True
                End If
            End If
            'Step 2
            If textboxes_text(1).IndexOf("\") > -1 Then
                splitter = textboxes_text(1).Split("\")
                Dim temp_next_info As Boolean = False
                For i1 As Short = 0 To splitter.Length - 1
                    If splitter(i1) <> "" Or radiobuttons_value(15) = True Then
                        If splitter(i1) = "" And radiobuttons_value(15) = True Then
                            temp_next_info = True
                            Exit For
                        Else
                            If db_data(i).year <> "" Then
                                temp_next_info = True
                                Exit For
                            End If
                        End If
                    End If
                Next
                If temp_next_info = True Then
                    Dim temp_res(-1) As Boolean
                    For i1 As Short = 0 To splitter.Length - 1
                        If splitter(i1) <> "" Then
                            Array.Resize(temp_res, temp_res.Length + 1)
                            temp_res(temp_res.Length - 1) = False
                            If checkboxes_value(1) Then
                                If radiobuttons_value(3) = True Then
                                    If db_data(i).year.IndexOf(splitter(i1)) > -1 Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                ElseIf radiobuttons_value(2) = True Then
                                    If splitter(i1).IndexOf(db_data(i).year) > -1 Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                ElseIf radiobuttons_value(15) = True Then
                                    If db_data(i).year = splitter(i1) Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                End If
                            Else
                                If radiobuttons_value(3) = True Then
                                    If db_data(i).year.ToLower.IndexOf(splitter(i1).ToLower) > -1 Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                ElseIf radiobuttons_value(2) = True Then
                                    If splitter(i1).ToLower.IndexOf(db_data(i).year.ToLower) > -1 Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                ElseIf radiobuttons_value(15) = True Then
                                    If db_data(i).year.ToLower = splitter(i1).ToLower Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                End If
                            End If
                        End If
                    Next
                    If radiobuttons_value(26) Then
                        results(1) = True
                        For Each i1 As Boolean In temp_res
                            If i1 <> True Then
                                results(1) = False
                                Exit For
                            End If
                        Next
                    Else
                        For Each i1 As Boolean In temp_res
                            If i1 = True Then
                                results(1) = True
                                Exit For
                            End If
                        Next
                    End If
                Else
                    results(1) = True
                    no_result(1) = True
                End If
            Else
                Dim temp_res As Boolean = False
                If textboxes_text(1) <> "" Or radiobuttons_value(15) = True Then
                    If textboxes_text(1) = "" And radiobuttons_value(15) = True Then
                        temp_res = True
                    Else
                        If db_data(i).year <> "" Then
                            temp_res = True
                        End If
                    End If
                End If
                If temp_res Then
                    If checkboxes_value(1) Then
                        If radiobuttons_value(3) = True Then
                            If db_data(i).year.IndexOf(textboxes_text(1)) > -1 Then
                                results(1) = True
                            End If
                        ElseIf radiobuttons_value(2) = True Then
                            If textboxes_text(1).IndexOf(db_data(i).year) > -1 Then
                                results(1) = True
                            End If
                        ElseIf radiobuttons_value(15) = True Then
                            If db_data(i).year = textboxes_text(1) Then
                                results(1) = True
                            End If
                        End If
                    Else
                        If radiobuttons_value(3) = True Then
                            If db_data(i).year.ToLower.IndexOf(textboxes_text(1).ToLower) > -1 Then
                                results(1) = True
                            End If
                        ElseIf radiobuttons_value(2) = True Then
                            If textboxes_text(1).ToLower.IndexOf(db_data(i).year.ToLower) > -1 Then
                                results(1) = True
                            End If
                        ElseIf radiobuttons_value(15) = True Then
                            If db_data(i).year.ToLower = textboxes_text(1).ToLower Then
                                results(1) = True
                            End If
                        End If
                    End If
                Else
                    results(1) = True
                    no_result(1) = True
                End If
            End If
            'Step 3
            If textboxes_text(2).IndexOf("\") > -1 Then
                splitter = textboxes_text(2).Split("\")
                Dim temp_next_info As Boolean = False
                For i1 As Short = 0 To splitter.Length - 1
                    If splitter(i1) <> "" Or radiobuttons_value(16) = True Then
                        If splitter(i1) = "" And radiobuttons_value(16) = True Then
                            temp_next_info = True
                            Exit For
                        Else
                            If db_data(i).school <> "" Then
                                temp_next_info = True
                                Exit For
                            End If
                        End If
                    End If
                Next
                If temp_next_info = True Then
                    Dim temp_res(-1) As Boolean
                    For i1 As Short = 0 To splitter.Length - 1
                        If splitter(i1) <> "" Then
                            Array.Resize(temp_res, temp_res.Length + 1)
                            temp_res(temp_res.Length - 1) = False
                            If checkboxes_value(2) Then
                                If radiobuttons_value(5) = True Then
                                    If db_data(i).school.IndexOf(splitter(i1)) > -1 Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                ElseIf radiobuttons_value(4) = True Then
                                    If splitter(i1).IndexOf(db_data(i).school) > -1 Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                ElseIf radiobuttons_value(16) = True Then
                                    If db_data(i).school = splitter(i1) Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                End If
                            Else
                                If radiobuttons_value(5) = True Then
                                    If db_data(i).school.ToLower.IndexOf(splitter(i1).ToLower) > -1 Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                ElseIf radiobuttons_value(4) = True Then
                                    If splitter(i1).ToLower.IndexOf(db_data(i).school.ToLower) > -1 Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                ElseIf radiobuttons_value(16) = True Then
                                    If db_data(i).school.ToLower = splitter(i1).ToLower Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                End If
                            End If
                        End If
                    Next
                    If radiobuttons_value(28) Then
                        results(2) = True
                        For Each i1 As Boolean In temp_res
                            If i1 <> True Then
                                results(2) = False
                                Exit For
                            End If
                        Next
                    Else
                        For Each i1 As Boolean In temp_res
                            If i1 = True Then
                                results(2) = True
                                Exit For
                            End If
                        Next
                    End If
                Else
                    results(2) = True
                    no_result(2) = True
                End If
            Else
                Dim temp_res As Boolean = False
                If textboxes_text(2) <> "" Or radiobuttons_value(16) = True Then
                    If textboxes_text(2) = "" And radiobuttons_value(16) = True Then
                        temp_res = True
                    Else
                        If db_data(i).school <> "" Then
                            temp_res = True
                        End If
                    End If
                End If
                If temp_res Then
                    If checkboxes_value(2) Then
                        If radiobuttons_value(5) = True Then
                            If db_data(i).school.IndexOf(textboxes_text(2)) > -1 Then
                                results(2) = True
                            End If
                        ElseIf radiobuttons_value(4) = True Then
                            If textboxes_text(2).IndexOf(db_data(i).school) > -1 Then
                                results(2) = True
                            End If
                        ElseIf radiobuttons_value(16) = True Then
                            If db_data(i).school = textboxes_text(2) Then
                                results(2) = True
                            End If
                        End If
                    Else
                        If radiobuttons_value(5) = True Then
                            If db_data(i).school.ToLower.IndexOf(textboxes_text(2).ToLower) > -1 Then
                                results(2) = True
                            End If
                        ElseIf radiobuttons_value(4) = True Then
                            If textboxes_text(2).ToLower.IndexOf(db_data(i).school.ToLower) > -1 Then
                                results(2) = True
                            End If
                        ElseIf radiobuttons_value(16) = True Then
                            If db_data(i).school.ToLower = textboxes_text(2).ToLower Then
                                results(2) = True
                            End If
                        End If
                    End If
                Else
                    results(2) = True
                    no_result(2) = True
                End If
            End If
            'Step 4
            If textboxes_text(3).IndexOf("\") > -1 Then
                splitter = textboxes_text(3).Split("\")
                Dim temp_next_info As Boolean = False
                For i1 As Short = 0 To splitter.Length - 1
                    If splitter(i1) <> "" Or radiobuttons_value(17) = True Then
                        If splitter(i1) = "" And radiobuttons_value(17) = True Then
                            temp_next_info = True
                            Exit For
                        Else
                            If db_data(i).class_ <> "" Then
                                temp_next_info = True
                                Exit For
                            End If
                        End If
                    End If
                Next
                If temp_next_info = True Then
                    Dim temp_res(-1) As Boolean
                    For i1 As Short = 0 To splitter.Length - 1
                        If splitter(i1) <> "" Then
                            Array.Resize(temp_res, temp_res.Length + 1)
                            temp_res(temp_res.Length - 1) = False
                            If checkboxes_value(3) Then
                                If radiobuttons_value(7) = True Then
                                    If db_data(i).class_.IndexOf(splitter(i1)) > -1 Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                ElseIf radiobuttons_value(6) = True Then
                                    If splitter(i1).IndexOf(db_data(i).class_) > -1 Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                ElseIf radiobuttons_value(17) = True Then
                                    If db_data(i).class_ = splitter(i1) Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                End If
                            Else
                                If radiobuttons_value(7) = True Then
                                    If db_data(i).class_.ToLower.IndexOf(splitter(i1).ToLower) > -1 Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                ElseIf radiobuttons_value(6) = True Then
                                    If splitter(i1).ToLower.IndexOf(db_data(i).class_.ToLower) > -1 Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                ElseIf radiobuttons_value(17) = True Then
                                    If db_data(i).class_.ToLower = splitter(i1).ToLower Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                End If
                            End If
                        End If
                    Next
                    If radiobuttons_value(30) Then
                        results(3) = True
                        For Each i1 As Boolean In temp_res
                            If i1 <> True Then
                                results(3) = False
                                Exit For
                            End If
                        Next
                    Else
                        For Each i1 As Boolean In temp_res
                            If i1 = True Then
                                results(3) = True
                                Exit For
                            End If
                        Next
                    End If
                Else
                    results(3) = True
                    no_result(3) = True
                End If
            Else
                Dim temp_res As Boolean = False
                If textboxes_text(3) <> "" Or radiobuttons_value(17) = True Then
                    If textboxes_text(3) = "" And radiobuttons_value(17) = True Then
                        temp_res = True
                    Else
                        If db_data(i).class_ <> "" Then
                            temp_res = True
                        End If
                    End If
                End If
                If temp_res Then
                    If checkboxes_value(3) Then
                        If radiobuttons_value(7) = True Then
                            If db_data(i).class_.IndexOf(textboxes_text(3)) > -1 Then
                                results(3) = True
                            End If
                        ElseIf radiobuttons_value(6) = True Then
                            If textboxes_text(3).IndexOf(db_data(i).class_) > -1 Then
                                results(3) = True
                            End If
                        ElseIf radiobuttons_value(17) = True Then
                            If db_data(i).class_ = textboxes_text(3) Then
                                results(3) = True
                            End If
                        End If
                    Else
                        If radiobuttons_value(7) = True Then
                            If db_data(i).class_.ToLower.IndexOf(textboxes_text(3).ToLower) > -1 Then
                                results(3) = True
                            End If
                        ElseIf radiobuttons_value(6) = True Then
                            If textboxes_text(3).ToLower.IndexOf(db_data(i).class_.ToLower) > -1 Then
                                results(3) = True
                            End If
                        ElseIf radiobuttons_value(17) = True Then
                            If db_data(i).class_.ToLower = textboxes_text(3).ToLower Then
                                results(3) = True
                            End If
                        End If
                    End If
                Else
                    results(3) = True
                    no_result(3) = True
                End If
            End If
            'Step 5
            If textboxes_text(4).IndexOf("\") > -1 Then
                splitter = textboxes_text(4).Split("\")
                Dim temp_next_info As Boolean = False
                For i1 As Short = 0 To splitter.Length - 1
                    If splitter(i1) <> "" Or radiobuttons_value(18) = True Then
                        If splitter(i1) = "" And radiobuttons_value(18) = True Then
                            temp_next_info = True
                            Exit For
                        Else
                            If db_data(i).telephone <> "" Then
                                temp_next_info = True
                                Exit For
                            End If
                        End If
                    End If
                Next
                If temp_next_info = True Then
                    Dim temp_res(-1) As Boolean
                    For i1 As Short = 0 To splitter.Length - 1
                        If splitter(i1) <> "" Then
                            Array.Resize(temp_res, temp_res.Length + 1)
                            temp_res(temp_res.Length - 1) = False
                            If checkboxes_value(4) Then
                                If radiobuttons_value(9) = True Then
                                    If db_data(i).telephone.IndexOf(splitter(i1)) > -1 Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                ElseIf radiobuttons_value(8) = True Then
                                    If splitter(i1).IndexOf(db_data(i).telephone) > -1 Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                ElseIf radiobuttons_value(18) = True Then
                                    If db_data(i).telephone = splitter(i1) Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                End If
                            Else
                                If radiobuttons_value(9) = True Then
                                    If db_data(i).telephone.ToLower.IndexOf(splitter(i1).ToLower) > -1 Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                ElseIf radiobuttons_value(8) = True Then
                                    If splitter(i1).ToLower.IndexOf(db_data(i).telephone.ToLower) > -1 Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                ElseIf radiobuttons_value(18) = True Then
                                    If db_data(i).telephone.ToLower = splitter(i1).ToLower Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                End If
                            End If
                        End If
                    Next
                    If radiobuttons_value(32) Then
                        results(4) = True
                        For Each i1 As Boolean In temp_res
                            If i1 <> True Then
                                results(4) = False
                                Exit For
                            End If
                        Next
                    Else
                        For Each i1 As Boolean In temp_res
                            If i1 = True Then
                                results(4) = True
                                Exit For
                            End If
                        Next
                    End If
                Else
                    results(4) = True
                    no_result(4) = True
                End If
            Else
                Dim temp_res As Boolean = False
                If textboxes_text(4) <> "" Or radiobuttons_value(18) = True Then
                    If textboxes_text(4) = "" And radiobuttons_value(18) = True Then
                        temp_res = True
                    Else
                        If db_data(i).telephone <> "" Then
                            temp_res = True
                        End If
                    End If
                End If
                If temp_res Then
                    If checkboxes_value(4) Then
                        If radiobuttons_value(9) = True Then
                            If db_data(i).telephone.IndexOf(textboxes_text(4)) > -1 Then
                                results(4) = True
                            End If
                        ElseIf radiobuttons_value(8) = True Then
                            If textboxes_text(4).IndexOf(db_data(i).telephone) > -1 Then
                                results(4) = True
                            End If
                        ElseIf radiobuttons_value(18) = True Then
                            If db_data(i).telephone = textboxes_text(4) Then
                                results(4) = True
                            End If
                        End If
                    Else
                        If radiobuttons_value(9) = True Then
                            If db_data(i).telephone.ToLower.IndexOf(textboxes_text(4).ToLower) > -1 Then
                                results(4) = True
                            End If
                        ElseIf radiobuttons_value(8) = True Then
                            If textboxes_text(4).ToLower.IndexOf(db_data(i).telephone.ToLower) > -1 Then
                                results(4) = True
                            End If
                        ElseIf radiobuttons_value(18) = True Then
                            If db_data(i).telephone.ToLower = textboxes_text(4).ToLower Then
                                results(4) = True
                            End If
                        End If
                    End If
                Else
                    results(4) = True
                    no_result(4) = True
                End If
            End If
            'Step 6
            If textboxes_text(5).IndexOf("\") > -1 Then
                splitter = textboxes_text(5).Split("\")
                Dim temp_next_info As Boolean = False
                For i1 As Short = 0 To splitter.Length - 1
                    If splitter(i1) <> "" Or radiobuttons_value(19) = True Then
                        If splitter(i1) = "" And radiobuttons_value(19) = True Then
                            temp_next_info = True
                            Exit For
                        Else
                            If db_data(i).post <> "" Then
                                temp_next_info = True
                                Exit For
                            End If
                        End If
                    End If
                Next
                If temp_next_info = True Then
                    Dim temp_res(-1) As Boolean
                    For i1 As Short = 0 To splitter.Length - 1
                        If splitter(i1) <> "" Then
                            Array.Resize(temp_res, temp_res.Length + 1)
                            temp_res(temp_res.Length - 1) = False
                            If checkboxes_value(5) Then
                                If radiobuttons_value(11) = True Then
                                    If db_data(i).post.IndexOf(splitter(i1)) > -1 Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                ElseIf radiobuttons_value(10) = True Then
                                    If splitter(i1).IndexOf(db_data(i).post) > -1 Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                ElseIf radiobuttons_value(19) = True Then
                                    If db_data(i).post = splitter(i1) Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                End If
                            Else
                                If radiobuttons_value(11) = True Then
                                    If db_data(i).post.ToLower.IndexOf(splitter(i1).ToLower) > -1 Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                ElseIf radiobuttons_value(10) = True Then
                                    If splitter(i1).ToLower.IndexOf(db_data(i).post.ToLower) > -1 Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                ElseIf radiobuttons_value(19) = True Then
                                    If db_data(i).post.ToLower = splitter(i1).ToLower Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                End If
                            End If
                        End If
                    Next
                    If radiobuttons_value(34) Then
                        results(5) = True
                        For Each i1 As Boolean In temp_res
                            If i1 <> True Then
                                results(5) = False
                                Exit For
                            End If
                        Next
                    Else
                        For Each i1 As Boolean In temp_res
                            If i1 = True Then
                                results(5) = True
                                Exit For
                            End If
                        Next
                    End If
                Else
                    results(5) = True
                    no_result(5) = True
                End If
            Else
                Dim temp_res As Boolean = False
                If textboxes_text(5) <> "" Or radiobuttons_value(19) = True Then
                    If textboxes_text(5) = "" And radiobuttons_value(19) = True Then
                        temp_res = True
                    Else
                        If db_data(i).post <> "" Then
                            temp_res = True
                        End If
                    End If
                End If
                If temp_res Then
                    If checkboxes_value(5) Then
                        If radiobuttons_value(11) = True Then
                            If db_data(i).post.IndexOf(textboxes_text(5)) > -1 Then
                                results(5) = True
                            End If
                        ElseIf radiobuttons_value(10) = True Then
                            If textboxes_text(5).IndexOf(db_data(i).post) > -1 Then
                                results(5) = True
                            End If
                        ElseIf radiobuttons_value(19) = True Then
                            If db_data(i).post = textboxes_text(5) Then
                                results(5) = True
                            End If
                        End If
                    Else
                        If radiobuttons_value(11) = True Then
                            If db_data(i).post.ToLower.IndexOf(textboxes_text(5).ToLower) > -1 Then
                                results(5) = True
                            End If
                        ElseIf radiobuttons_value(10) = True Then
                            If textboxes_text(5).ToLower.IndexOf(db_data(i).post.ToLower) > -1 Then
                                results(5) = True
                            End If
                        ElseIf radiobuttons_value(19) = True Then
                            If db_data(i).post.ToLower = textboxes_text(5).ToLower Then
                                results(5) = True
                            End If
                        End If
                    End If
                Else
                    results(5) = True
                    no_result(5) = True
                End If
            End If
            'Step 7
            Dim str_lines As String = ""
            For i1 As Integer = 0 To db_data(i).skills.Length - 1
                If i1 + 1 = db_data(i).skills.Length Then
                    str_lines &= db_data(i).skills(i1)
                Else
                    str_lines &= db_data(i).skills(i1) & vbNewLine
                End If
            Next
            If textboxes_text(6).IndexOf("\") > -1 Then
                splitter = textboxes_text(6).Split("\")
                Dim temp_next_info As Boolean = False
                For i1 As Short = 0 To splitter.Length - 1
                    If splitter(i1) <> "" Or radiobuttons_value(12) = True Then
                        If splitter(i1) = "" And radiobuttons_value(12) = True Then
                            temp_next_info = True
                            Exit For
                        Else
                            If Add_Redact_db_data.ArraysEqual(db_data(i).skills, {""}) = False Then
                                temp_next_info = True
                                Exit For
                            End If
                        End If
                    End If
                Next
                If temp_next_info = True Then
                    For i1 As Short = 0 To splitter.Length - 1
                        splitter(i1) = splitter(i1).Trim(Chr(10), Chr(13))
                    Next
                    Dim temp_res(-1) As Boolean
                    For i1 As Short = 0 To splitter.Length - 1
                        If splitter(i1) <> "" Then
                            Array.Resize(temp_res, temp_res.Length + 1)
                            temp_res(temp_res.Length - 1) = False
                            If checkboxes_value(6) Then
                                If radiobuttons_value(20) = True Then
                                    If str_lines.IndexOf(splitter(i1)) > -1 Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                ElseIf radiobuttons_value(13) = True Then
                                    If splitter(i1).IndexOf(str_lines) > -1 Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                ElseIf radiobuttons_value(12) = True Then
                                    If str_lines = splitter(i1) Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                End If
                            Else
                                If radiobuttons_value(20) = True Then
                                    If str_lines.ToLower.IndexOf(splitter(i1).ToLower) > -1 Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                ElseIf radiobuttons_value(13) = True Then
                                    If splitter(i1).ToLower.IndexOf(str_lines.ToLower) > -1 Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                ElseIf radiobuttons_value(12) = True Then
                                    If str_lines.ToLower = splitter(i1).ToLower Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                End If
                            End If
                        End If
                    Next
                    If radiobuttons_value(36) Then
                        results(6) = True
                        For Each i1 As Boolean In temp_res
                            If i1 <> True Then
                                results(6) = False
                                Exit For
                            End If
                        Next
                    Else
                        For Each i1 As Boolean In temp_res
                            If i1 = True Then
                                results(6) = True
                                Exit For
                            End If
                        Next
                    End If
                Else
                    results(6) = True
                    no_result(6) = True
                End If
            Else
                Dim temp_res As Boolean = False
                If textboxes_text(6) <> "" Or radiobuttons_value(12) = True Then
                    If textboxes_text(6) = "" And radiobuttons_value(12) = True Then
                        temp_res = True
                    Else
                        If Add_Redact_db_data.ArraysEqual(db_data(i).skills, {""}) = False Then
                            temp_res = True
                        End If
                    End If
                End If
                If temp_res Then
                    If checkboxes_value(6) Then
                        If radiobuttons_value(20) = True Then
                            If str_lines.IndexOf(textboxes_text(6)) > -1 Then
                                results(6) = True
                            End If
                        ElseIf radiobuttons_value(13) = True Then
                            If textboxes_text(6).IndexOf(str_lines) > -1 Then
                                results(6) = True
                            End If
                        ElseIf radiobuttons_value(12) = True Then
                            If str_lines = textboxes_text(6) Then
                                results(6) = True
                            End If
                        End If
                    Else
                        If radiobuttons_value(20) = True Then
                            If str_lines.ToLower.IndexOf(textboxes_text(6).ToLower) > -1 Then
                                results(6) = True
                            End If
                        ElseIf radiobuttons_value(13) = True Then
                            If textboxes_text(6).ToLower.IndexOf(str_lines.ToLower) > -1 Then
                                results(6) = True
                            End If
                        ElseIf radiobuttons_value(12) = True Then
                            If str_lines.ToLower = textboxes_text(6).ToLower Then
                                results(6) = True
                            End If
                        End If
                    End If
                Else
                    results(6) = True
                    no_result(6) = True
                End If
            End If
            'Step 8
            str_lines = ""
            For i1 As Integer = 0 To db_data(i).achievements.Length - 1
                If i1 + 1 = db_data(i).achievements.Length Then
                    str_lines &= db_data(i).achievements(i1)
                Else
                    str_lines &= db_data(i).achievements(i1) & vbNewLine
                End If
            Next
            If textboxes_text(7).IndexOf("\") > -1 Then
                splitter = textboxes_text(7).Split("\")
                Dim temp_next_info As Boolean = False
                For i1 As Short = 0 To splitter.Length - 1
                    If splitter(i1) <> "" Or radiobuttons_value(37) = True Then
                        If splitter(i1) = "" And radiobuttons_value(37) = True Then
                            temp_next_info = True
                            Exit For
                        Else
                            If Add_Redact_db_data.ArraysEqual(db_data(i).achievements, {""}) = False Then
                                temp_next_info = True
                                Exit For
                            End If
                        End If
                    End If
                Next
                If temp_next_info = True Then
                    For i1 As Short = 0 To splitter.Length - 1
                        splitter(i1) = splitter(i1).Trim(Chr(10), Chr(13))
                    Next
                    Dim temp_res(-1) As Boolean
                    For i1 As Short = 0 To splitter.Length - 1
                        If splitter(i1) <> "" Then
                            Array.Resize(temp_res, temp_res.Length + 1)
                            temp_res(temp_res.Length - 1) = False
                            If checkboxes_value(7) Then
                                If radiobuttons_value(39) = True Then
                                    If str_lines.IndexOf(splitter(i1)) > -1 Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                ElseIf radiobuttons_value(38) = True Then
                                    If splitter(i1).IndexOf(str_lines) > -1 Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                ElseIf radiobuttons_value(37) = True Then
                                    If str_lines = splitter(i1) Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                End If
                            Else
                                If radiobuttons_value(39) = True Then
                                    If str_lines.ToLower.IndexOf(splitter(i1).ToLower) > -1 Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                ElseIf radiobuttons_value(38) = True Then
                                    If splitter(i1).ToLower.IndexOf(str_lines.ToLower) > -1 Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                ElseIf radiobuttons_value(37) = True Then
                                    If str_lines.ToLower = splitter(i1).ToLower Then
                                        temp_res(temp_res.Length - 1) = True
                                    End If
                                End If
                            End If
                        End If
                    Next
                    If radiobuttons_value(41) Then
                        results(7) = True
                        For Each i1 As Boolean In temp_res
                            If i1 <> True Then
                                results(7) = False
                                Exit For
                            End If
                        Next
                    Else
                        For Each i1 As Boolean In temp_res
                            If i1 = True Then
                                results(7) = True
                                Exit For
                            End If
                        Next
                    End If
                Else
                    results(7) = True
                    no_result(7) = True
                End If
            Else
                Dim temp_res As Boolean = False
                If textboxes_text(7) <> "" Or radiobuttons_value(37) = True Then
                    If textboxes_text(7) = "" And radiobuttons_value(37) = True Then
                        temp_res = True
                    Else
                        If Add_Redact_db_data.ArraysEqual(db_data(i).achievements, {""}) = False Then
                            temp_res = True
                        End If
                    End If
                End If
                If temp_res Then
                    If checkboxes_value(7) Then
                        If radiobuttons_value(39) = True Then
                            If str_lines.IndexOf(textboxes_text(7)) > -1 Then
                                results(7) = True
                            End If
                        ElseIf radiobuttons_value(38) = True Then
                            If textboxes_text(7).IndexOf(str_lines) > -1 Then
                                results(7) = True
                            End If
                        ElseIf radiobuttons_value(37) = True Then
                            If str_lines = textboxes_text(7) Then
                                results(7) = True
                            End If
                        End If
                    Else
                        If radiobuttons_value(39) = True Then
                            If str_lines.ToLower.IndexOf(textboxes_text(7).ToLower) > -1 Then
                                results(7) = True
                            End If
                        ElseIf radiobuttons_value(38) = True Then
                            If textboxes_text(7).ToLower.IndexOf(str_lines.ToLower) > -1 Then
                                results(7) = True
                            End If
                        ElseIf radiobuttons_value(37) = True Then
                            If str_lines.ToLower = textboxes_text(7).ToLower Then
                                results(7) = True
                            End If
                        End If
                    End If
                Else
                    results(7) = True
                    no_result(7) = True
                End If
            End If
            Dim all_result As Boolean = False, resulted As Boolean = False
            If radiobuttons_value(24) Then
                all_result = True
                For i1 As Short = 0 To 7
                    If results(i1) <> True And no_result(i1) = False Then
                        all_result = False
                        Exit For
                    ElseIf results(i1) = True And no_result(i1) = False Then
                        resulted = True
                    End If
                Next
            Else
                resulted = True
                For i1 As Short = 0 To 7
                    If results(i1) = True And no_result(i1) = False Then
                        all_result = True
                        Exit For
                    End If
                Next
            End If
            If all_result = True And resulted = True Then
                sdb_keys.Add(db_data(i).key)
            End If
        Next
    End Sub
    Private Sub Search_Elements_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles Search_Elements.RunWorkerCompleted
        Search_db.Search_Elements_Completed()
    End Sub
    Dim export_to_word_cb_values(11) As Boolean, export_to_word_tb As String, export_to_word_cb(1) As UShort, export_to_word_indexes As New List(Of export_to_word_link)
    Public export_to_word_progress As UShort
    Public Sub Expor_to_word_start()
        export_to_word_cb_values(0) = Member_to_word.CheckBox1.Checked
        export_to_word_cb_values(1) = Member_to_word.CheckBox2.Checked
        export_to_word_cb_values(2) = Member_to_word.CheckBox3.Checked
        export_to_word_cb_values(3) = Member_to_word.CheckBox4.Checked
        export_to_word_cb_values(4) = Member_to_word.CheckBox5.Checked
        export_to_word_cb_values(5) = Member_to_word.CheckBox6.Checked
        export_to_word_cb_values(6) = Member_to_word.CheckBox7.Checked
        export_to_word_cb_values(7) = Member_to_word.CheckBox8.Checked
        export_to_word_cb_values(8) = Member_to_word.CheckBox9.Checked
        export_to_word_cb_values(9) = Member_to_word.CheckBox10.Checked 'Use object pefrics
        export_to_word_cb_values(10) = Member_to_word.CheckBox11.Checked 'Stationar object pefrics
        export_to_word_cb_values(11) = Member_to_word.CheckBox12.Checked
        export_to_word_tb = Member_to_word.TextBox1.Text
        export_to_word_cb(0) = Member_to_word.ComboBox1.SelectedIndex
        export_to_word_cb(1) = Member_to_word.ComboBox2.SelectedIndex
        Dim doc_num As UShort = 1
        For Each i As DataGridViewRow In DataGridView1.SelectedRows
            export_to_word_indexes.Add(New export_to_word_link(i.Cells.Item(0).Value, 0, doc_num))
            doc_num += 1
        Next
        ExportToWord.RunWorkerAsync()
    End Sub
    Private Sub ExportToWord_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles ExportToWord.DoWork
        Dim temp_list As New List(Of db_Participator), save_names(-1) As String
        For i As Short = 0 To export_to_word_indexes.Count - 1
            Dim pos As Short = -1
            For i1 As Short = 0 To db_data.Count - 1
                If db_data(i1).key = export_to_word_indexes(i).index Then
                    temp_list.Add(Clone_dbPart_Class(db_data(i1)))
                    pos = i1
                    temp_list(temp_list.Count - 1).img = Nothing
                    temp_list(temp_list.Count - 1).skills = {""}
                    temp_list(temp_list.Count - 1).achievements = {""}
                    Exit For
                End If
            Next
            export_to_word_indexes(i).err = Make_word_doc(db_data(pos), TEMP_WAY & "\MCD-temp-47591795970905\temp\ti1.jpg", TEMP_WAY & "\MCD-temp-47591795970905\temp\twd" & export_to_word_indexes(i).doc_index & ".docx")
            If Progress.progress_mode(0) = 1 Then
                If Math.IEEERemainder(i, Progress.progress_mode(1)) = 0 Then
                    export_to_word_progress += 1
                End If
            ElseIf Progress.progress_mode(0) = 0 Then
                export_to_word_progress += 1
            End If
        Next
        Dim file_name As String = Mid(export_to_word_tb, export_to_word_tb.LastIndexOf("\") + 2, export_to_word_tb.Length - export_to_word_tb.LastIndexOf("\") - 6)
        Dim file_dir As String = Mid(export_to_word_tb, 1, export_to_word_tb.LastIndexOf("\"))
        If My.Computer.FileSystem.DirectoryExists(file_dir) Then
            Try
                If export_to_word_cb(1) = 2 Then
                    Dim files() As String = IO.Directory.GetFiles(file_dir, "*.docx", IO.SearchOption.TopDirectoryOnly)
                    For Each i As String In files
                        My.Computer.FileSystem.DeleteFile(i)
                    Next
                End If
                If export_to_word_cb(1) = 1 Then
                    For i As Short = 0 To export_to_word_indexes.Count - 1
                        If export_to_word_indexes(i).err = 0 Then
                            Try
                                If export_to_word_cb_values(9) Then
                                    Dim perfics As String = ""
                                    If export_to_word_cb(0) = 0 Then
                                        perfics = Replace_banned_simvols(temp_list(i).name)
                                    ElseIf export_to_word_cb(0) = 1 Then
                                        perfics = Replace_banned_simvols(temp_list(i).year)
                                    ElseIf export_to_word_cb(0) = 2 Then
                                        perfics = Replace_banned_simvols(temp_list(i).school)
                                    ElseIf export_to_word_cb(0) = 3 Then
                                        perfics = Replace_banned_simvols(temp_list(i).class_)
                                    ElseIf export_to_word_cb(0) = 4 Then
                                        perfics = Replace_banned_simvols(temp_list(i).telephone)
                                    ElseIf export_to_word_cb(0) = 5 Then
                                        perfics = Replace_banned_simvols(temp_list(i).post)
                                    End If
                                    If export_to_word_cb_values(10) Then
                                        If String_exist(save_names, file_dir & "\" & file_name & "-" & perfics & ".docx") = False Then
                                            My.Computer.FileSystem.CopyFile(TEMP_WAY & "\MCD-temp-47591795970905\temp\twd" & export_to_word_indexes(i).doc_index & ".docx", file_dir & "\" & file_name & "-" & perfics & ".docx")
                                            Array.Resize(save_names, save_names.Length + 1)
                                            save_names(save_names.Length - 1) = file_dir & "\" & file_name & "-" & perfics & ".docx"
                                        Else
                                            Dim selected_num As Short = 1, file_exist As Boolean = True
                                            While file_exist = True
                                                If String_exist(save_names, file_dir & "\" & file_name & "-" & perfics & "-" & selected_num & ".docx") = False Then
                                                    file_exist = False
                                                Else
                                                    selected_num += 1
                                                End If
                                            End While
                                            My.Computer.FileSystem.CopyFile(TEMP_WAY & "\MCD-temp-47591795970905\temp\twd" & export_to_word_indexes(i).doc_index & ".docx", file_dir & "\" & file_name & "-" & perfics & "-" & selected_num & ".docx", True)
                                            Array.Resize(save_names, save_names.Length + 1)
                                            save_names(save_names.Length - 1) = file_dir & "\" & file_name & "-" & perfics & "-" & selected_num & ".docx"
                                        End If
                                    Else
                                        If String_exist(save_names, file_dir & "\" & file_name & ".docx") = False Then
                                            My.Computer.FileSystem.CopyFile(TEMP_WAY & "\MCD-temp-47591795970905\temp\twd" & export_to_word_indexes(i).doc_index & ".docx", file_dir & "\" & file_name & ".docx", True)
                                            Array.Resize(save_names, save_names.Length + 1)
                                            save_names(save_names.Length - 1) = file_dir & "\" & file_name & ".docx"
                                        Else
                                            If String_exist(save_names, file_dir & "\" & file_name & "-" & perfics & ".docx") = False Then
                                                My.Computer.FileSystem.CopyFile(TEMP_WAY & "\MCD-temp-47591795970905\temp\twd" & export_to_word_indexes(i).doc_index & ".docx", file_dir & "\" & file_name & "-" & perfics & ".docx", True)
                                                Array.Resize(save_names, save_names.Length + 1)
                                                save_names(save_names.Length - 1) = file_dir & "\" & file_name & "-" & perfics & ".docx"
                                            Else
                                                Dim selected_num As Short = 1, file_exist As Boolean = True
                                                While file_exist = True
                                                    If String_exist(save_names, file_dir & "\" & file_name & "-" & perfics & "-" & selected_num & ".docx") = False Then
                                                        file_exist = False
                                                    Else
                                                        selected_num += 1
                                                    End If
                                                End While
                                                My.Computer.FileSystem.CopyFile(TEMP_WAY & "\MCD-temp-47591795970905\temp\twd" & export_to_word_indexes(i).doc_index & ".docx", file_dir & "\" & file_name & "-" & perfics & "-" & selected_num & ".docx", True)
                                                Array.Resize(save_names, save_names.Length + 1)
                                                save_names(save_names.Length - 1) = file_dir & "\" & file_name & "-" & perfics & "-" & selected_num & ".docx"
                                            End If
                                        End If
                                    End If
                                Else
                                    If String_exist(save_names, file_dir & "\" & file_name & ".docx") = False Then
                                        My.Computer.FileSystem.CopyFile(TEMP_WAY & "\MCD-temp-47591795970905\temp\twd" & export_to_word_indexes(i).doc_index & ".docx", file_dir & "\" & file_name & ".docx", True)
                                        Array.Resize(save_names, save_names.Length + 1)
                                        save_names(save_names.Length - 1) = file_dir & "\" & file_name & ".docx"
                                    Else
                                        Dim selected_num As Short = 1, file_exist As Boolean = True
                                        While file_exist = True
                                            If String_exist(save_names, file_dir & "\" & file_name & "-" & selected_num & ".docx") = False Then
                                                file_exist = False
                                            Else
                                                selected_num += 1
                                            End If
                                        End While
                                        My.Computer.FileSystem.CopyFile(TEMP_WAY & "\MCD-temp-47591795970905\temp\twd" & export_to_word_indexes(i).doc_index & ".docx", file_dir & "\" & file_name & "-" & selected_num & ".docx")
                                        Array.Resize(save_names, save_names.Length + 1)
                                        save_names(save_names.Length - 1) = file_dir & "\" & file_name & "-" & selected_num & ".docx"
                                    End If
                                End If
                            Catch
                                export_to_word_indexes(i).err = 1
                            End Try
                        End If
                    Next
                Else
                    For i As Short = 0 To export_to_word_indexes.Count - 1
                        If export_to_word_indexes(i).err = 0 Then
                            Try
                                If export_to_word_cb_values(9) Then
                                    Dim perfics As String = ""
                                    If export_to_word_cb(0) = 0 Then
                                        perfics = Replace_banned_simvols(temp_list(i).name)
                                    ElseIf export_to_word_cb(0) = 1 Then
                                        perfics = Replace_banned_simvols(temp_list(i).year)
                                    ElseIf export_to_word_cb(0) = 2 Then
                                        perfics = Replace_banned_simvols(temp_list(i).school)
                                    ElseIf export_to_word_cb(0) = 3 Then
                                        perfics = Replace_banned_simvols(temp_list(i).class_)
                                    ElseIf export_to_word_cb(0) = 4 Then
                                        perfics = Replace_banned_simvols(temp_list(i).telephone)
                                    ElseIf export_to_word_cb(0) = 5 Then
                                        perfics = Replace_banned_simvols(temp_list(i).post)
                                    End If
                                    If export_to_word_cb_values(10) Then
                                        If My.Computer.FileSystem.FileExists(file_dir & "\" & file_name & "-" & perfics & ".docx") = False Then
                                            My.Computer.FileSystem.CopyFile(TEMP_WAY & "\MCD-temp-47591795970905\temp\twd" & export_to_word_indexes(i).doc_index & ".docx", file_dir & "\" & file_name & "-" & perfics & ".docx")
                                        Else
                                            Dim selected_num As Short = 1, file_exist As Boolean = True
                                            While file_exist = True
                                                If My.Computer.FileSystem.FileExists(file_dir & "\" & file_name & "-" & perfics & "-" & selected_num & ".docx") = False Then
                                                    file_exist = False
                                                Else
                                                    selected_num += 1
                                                End If
                                            End While
                                            My.Computer.FileSystem.CopyFile(TEMP_WAY & "\MCD-temp-47591795970905\temp\twd" & export_to_word_indexes(i).doc_index & ".docx", file_dir & "\" & file_name & "-" & perfics & "-" & selected_num & ".docx")
                                        End If
                                    Else
                                        If My.Computer.FileSystem.FileExists(file_dir & "\" & file_name & ".docx") = False Then
                                            My.Computer.FileSystem.CopyFile(TEMP_WAY & "\MCD-temp-47591795970905\temp\twd" & export_to_word_indexes(i).doc_index & ".docx", file_dir & "\" & file_name & ".docx")
                                        Else
                                            If My.Computer.FileSystem.FileExists(file_dir & "\" & file_name & "-" & perfics & ".docx") = False Then
                                                My.Computer.FileSystem.CopyFile(TEMP_WAY & "\MCD-temp-47591795970905\temp\twd" & export_to_word_indexes(i).doc_index & ".docx", file_dir & "\" & file_name & "-" & perfics & ".docx")
                                            Else
                                                Dim selected_num As Short = 1, file_exist As Boolean = True
                                                While file_exist = True
                                                    If My.Computer.FileSystem.FileExists(file_dir & "\" & file_name & "-" & perfics & "-" & selected_num & ".docx") = False Then
                                                        file_exist = False
                                                    Else
                                                        selected_num += 1
                                                    End If
                                                End While
                                                My.Computer.FileSystem.CopyFile(TEMP_WAY & "\MCD-temp-47591795970905\temp\twd" & export_to_word_indexes(i).doc_index & ".docx", file_dir & "\" & file_name & "-" & perfics & "-" & selected_num & ".docx")
                                            End If
                                        End If
                                    End If
                                Else
                                    If My.Computer.FileSystem.FileExists(file_dir & "\" & file_name & ".docx") = False Then
                                        My.Computer.FileSystem.CopyFile(TEMP_WAY & "\MCD-temp-47591795970905\temp\twd" & export_to_word_indexes(i).doc_index & ".docx", file_dir & "\" & file_name & ".docx")
                                    Else
                                        Dim selected_num As Short = 1, file_exist As Boolean = True
                                        While file_exist = True
                                            If My.Computer.FileSystem.FileExists(file_dir & "\" & file_name & "-" & selected_num & ".docx") = False Then
                                                file_exist = False
                                            Else
                                                selected_num += 1
                                            End If
                                        End While
                                        My.Computer.FileSystem.CopyFile(TEMP_WAY & "\MCD-temp-47591795970905\temp\twd" & export_to_word_indexes(i).doc_index & ".docx", file_dir & "\" & file_name & "-" & selected_num & ".docx")
                                    End If
                                End If
                            Catch
                                export_to_word_indexes(i).err = 1
                            End Try
                        End If
                    Next
                End If
            Catch
                For Each i As export_to_word_link In export_to_word_indexes
                    i.err = 1
                Next
            End Try
            For i As Short = 0 To export_to_word_indexes.Count - 1
                If export_to_word_indexes(i).err = 0 Or export_to_word_indexes(i).err = 1 Then
                    Try
                        My.Computer.FileSystem.DeleteFile(TEMP_WAY & "\MCD-temp-47591795970905\temp\twd" & export_to_word_indexes(i).doc_index & ".docx")
                    Catch
                    End Try
                End If
            Next
        Else
            For Each i As export_to_word_link In export_to_word_indexes
                i.err = 1
            Next
        End If
        ReDim save_names(-1)
        temp_list.Clear()
    End Sub
    Private Function String_exist(str_arr() As String, search_str As String) As Boolean
        Dim finded As Boolean = False
        For Each i As String In str_arr
            If i = search_str Then
                finded = True
                Exit For
            End If
        Next
        Return finded
    End Function
    Private Sub ExportToWord_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles ExportToWord.RunWorkerCompleted
        Member_to_word.Export_to_word_ended()
        Dim sucess_count As UShort = 0
        For Each i As export_to_word_link In export_to_word_indexes
            If i.err = 0 Then
                sucess_count += 1
            End If
        Next
        MsgBox("Экспортировано элементов в Word: " & sucess_count & " из " & export_to_word_indexes.Count, MsgBoxStyle.OkOnly, "Информация")
        If sucess_count < export_to_word_indexes.Count Then
            Dim result As MsgBoxResult = MsgBox("Вы хотите сохранить отчёт об ошибках?", MsgBoxStyle.YesNo, "Сообщение")
            If result = MsgBoxResult.Yes Then
                Dim error_list(0) As String
                error_list(0) = "---- List of printing module MCD errors. ----"
                For i As Short = 0 To export_to_word_indexes.Count - 1
                    Dim pos As Short = -1
                    For i1 As Short = 0 To db_data.Count - 1
                        If db_data(i1).key = export_to_word_indexes(i).index Then
                            pos = i1
                            Exit For
                        End If
                    Next
                    Array.Resize(error_list, error_list.Length + 1)
                    error_list(error_list.Length - 1) = "document number:" & export_to_word_indexes(i).doc_index & " |error code:" & export_to_word_indexes(i).err
                Next
                Dim file_dir As String = Mid(export_to_word_tb, 1, export_to_word_tb.LastIndexOf("\"))
                Array.Resize(error_list, error_list.Length + 4)
                error_list(error_list.Length - 4) = "---- MCD information ----"
                error_list(error_list.Length - 3) = "Version: " & MY_VERSION
                error_list(error_list.Length - 2) = "Work directory: " & MY_WAY
                error_list(error_list.Length - 1) = "Destination directory: " & file_dir
                Dim hdb_time As String = Format(My.Computer.Clock.LocalTime.Day, "00") & "." & Format(My.Computer.Clock.LocalTime.Month, "00") & "." & Format(My.Computer.Clock.LocalTime.Year, "0000")
                If My.Computer.FileSystem.DirectoryExists(MY_WAY & "\crash reports") Then
                    If My.Computer.FileSystem.FileExists(MY_WAY & "\crash reports\" & hdb_time & "-ExportToWord_CrashReport.txt") Then
                        Dim selected_num As Short = 1, file_exist As Boolean = True
                        While file_exist = True
                            If My.Computer.FileSystem.FileExists(MY_WAY & "\crash reports\" & hdb_time & "-ExportToWord_CrashReport-" & selected_num & ".txt") = False Then
                                file_exist = False
                            Else
                                selected_num += 1
                            End If
                        End While
                        IO.File.WriteAllLines(MY_WAY & "\crash reports\" & hdb_time & "-ExportToWord_CrashReport-" & selected_num & ".txt", error_list)
                    Else
                        IO.File.WriteAllLines(MY_WAY & "\crash reports\" & hdb_time & "-ExportToWord_CrashReport.txt", error_list)
                    End If
                Else
                    My.Computer.FileSystem.CreateDirectory(MY_WAY & "\crash reports")
                    IO.File.WriteAllLines(MY_WAY & "\crash reports\" & hdb_time & "-ExportToWord_CrashReport.txt", error_list)
                End If
                MsgBox("Отчёт об ошибках сохранен в папку: """ & MY_WAY & "\crash reports""", MsgBoxStyle.OkOnly, "Информация")
            End If
        End If
        export_to_word_indexes.Clear()
    End Sub
    Private Class export_to_word_link
        Public index As String = "", err As UShort = 0, doc_index As UShort = 1
        Public Sub New(index As String, err As UShort, doc_index As UShort)
            Me.index = index
            Me.err = err
            Me.doc_index = doc_index
        End Sub
    End Class
    Private Function Replace_banned_simvols(text As String) As String
        Dim exit_str As String = Replace(text, "/", "")
        exit_str = Replace(exit_str, ":", "")
        exit_str = Replace(exit_str, "*", "")
        exit_str = Replace(exit_str, "?", "")
        exit_str = Replace(exit_str, """", "")
        Return exit_str
    End Function
    Private Sub DB_load_err()
        Dim hdb_time As String = Format(My.Computer.Clock.LocalTime.Day, "00") & "." & Format(My.Computer.Clock.LocalTime.Month, "00") & "." & Format(My.Computer.Clock.LocalTime.Year, "0000")
        If My.Computer.FileSystem.DirectoryExists(MY_WAY & "\harmed databases") Then
            If My.Computer.FileSystem.FileExists(MY_WAY & "\harmed databases\" & hdb_time & ".dat") Then
                Dim selected_num As Short = 1, file_exist As Boolean = True
                While file_exist = True
                    If My.Computer.FileSystem.FileExists(MY_WAY & "\harmed databases\" & hdb_time & "-" & selected_num & ".dat") = False Then
                        file_exist = False
                    Else
                        selected_num += 1
                    End If
                End While
                My.Computer.FileSystem.CopyFile(MY_WAY & "\database.dat", MY_WAY & "\harmed databases\" & hdb_time & "-" & selected_num & ".dat")
            Else
                My.Computer.FileSystem.CopyFile(MY_WAY & "\database.dat", MY_WAY & "\harmed databases\" & hdb_time & ".dat")
            End If
        Else
            My.Computer.FileSystem.CreateDirectory(MY_WAY & "\harmed databases")
            My.Computer.FileSystem.CopyFile(MY_WAY & "\database.dat", MY_WAY & "\harmed databases\" & hdb_time & ".dat")
        End If
        My.Computer.FileSystem.DeleteDirectory(TEMP_WAY & "\MCD-temp-47591795970905\database", FileIO.DeleteDirectoryOption.DeleteAllContents)
        MsgBox("База данных Участников повреждена!" & Chr(10) & "Работа с этой базой данных не может быть продолжена!" & Chr(10) & "MCD сделал резервную копию имеющейся базы данных Участников в папке: """ & MY_WAY & "\harmed databases""" & Chr(10) & "База данных Участников будет пересоздана.",, "Ошибка")
        Messanger.Add_messange("База данных Участников будет пересоздана.", "", True)
        My.Computer.FileSystem.CreateDirectory(TEMP_WAY & "\MCD-temp-47591795970905\database")
        Dim info_file(), collector_file() As String
        info_file = {"[format]", "LmsI-CLrw-6o2m1i", "", "[versions]", MY_VERSION}
        collector_file = {"[ADB info]", "LastReadTime=" & Get_date_string(My.Computer.Clock.LocalTime), "LastWriteTime=" & Get_date_string(My.Computer.Clock.LocalTime), "progID=" & info_set(2), "userName=" & info_set(0), "organization=" & info_set(1), "", "[participators]"}
        IO.File.WriteAllLines(TEMP_WAY & "\MCD-temp-47591795970905\database\info.ini", info_file)
        IO.File.WriteAllLines(TEMP_WAY & "\MCD-temp-47591795970905\database\collector.ini", collector_file)
        dbf_collector = New db_collector_file
        db_data.Clear()
        cmd_service_wait_proc = Shell(MY_WAY & "\mcdCAnc.exe ""MCDmode:archiver,Code:2801"" ""compress"" """ & MY_WAY & "\database.dat"" """ & TEMP_WAY & "\MCD-temp-47591795970905\database""")
        cmd_service_wait_mode = 2
        cmd_service_wait.Start()
    End Sub
    Private Function Make_word_doc(part As db_Participator, photo_way As String, doc_way As String) As UShort
        Dim e_code As UShort = 7
        Try
            Dim doc As New Microsoft.Office.Interop.Word.Document, app As New Microsoft.Office.Interop.Word.Application
            doc = app.Documents.Add()
            doc.PageSetup.LeftMargin = 30
            doc.PageSetup.TopMargin = 30
            doc.PageSetup.BottomMargin = 30
            doc.PageSetup.RightMargin = 30
            e_code = 6
            Dim paragraph As Microsoft.Office.Interop.Word.Paragraph, p_size As ULong = 0
            'Name print
            If export_to_word_cb_values(0) Then
                paragraph = doc.Content.Paragraphs.Add
                paragraph.Range.Font.Size = 18
                paragraph.Range.Text = "Имя: " & part.name
                app.Selection.Start = 0
                app.Selection.End = "Имя:".Length
                app.Selection.ParagraphFormat.SpaceBefore = 0
                app.Selection.ParagraphFormat.SpaceAfter = 12
                app.Selection.ParagraphFormat.LineSpacing = 12
                app.Selection.Range.Font.Size = 20
                app.Selection.Range.Font.Bold = True
                app.Selection.Range.Font.Underline = True
                p_size = paragraph.Range.Text.Length
                paragraph.Range.InsertParagraphAfter()
            End If
            'Old
            If export_to_word_cb_values(1) Then
                paragraph = doc.Content.Paragraphs.Add
                paragraph.Range.Font.Size = 18
                paragraph.Range.Text = "Дата рождения: " & part.year
                app.Selection.Start = p_size
                app.Selection.End = p_size + "Дата рождения:".Length
                app.Selection.ParagraphFormat.SpaceBefore = 0
                app.Selection.ParagraphFormat.SpaceAfter = 12
                app.Selection.ParagraphFormat.LineSpacing = 12
                app.Selection.Range.Font.Size = 20
                app.Selection.Range.Font.Bold = True
                app.Selection.Range.Font.Underline = True
                p_size += paragraph.Range.Text.Length
                paragraph.Range.InsertParagraphAfter()
            End If
            'School
            If export_to_word_cb_values(2) Then
                paragraph = doc.Content.Paragraphs.Add
                paragraph.Range.Font.Size = 18
                paragraph.Range.Text = "Школа: " & part.school
                app.Selection.Start = p_size
                app.Selection.End = p_size + "Школа:".Length
                app.Selection.ParagraphFormat.SpaceBefore = 0
                app.Selection.ParagraphFormat.SpaceAfter = 12
                app.Selection.ParagraphFormat.LineSpacing = 12
                app.Selection.Range.Font.Size = 20
                app.Selection.Range.Font.Bold = True
                app.Selection.Range.Font.Underline = True
                p_size += paragraph.Range.Text.Length
                paragraph.Range.InsertParagraphAfter()
            End If
            'Class
            If export_to_word_cb_values(3) Then
                paragraph = doc.Content.Paragraphs.Add
                paragraph.Range.Font.Size = 18
                paragraph.Range.Text = "Класс: " & part.class_
                app.Selection.Start = p_size
                app.Selection.End = p_size + "Класс:".Length
                app.Selection.ParagraphFormat.SpaceBefore = 0
                app.Selection.ParagraphFormat.SpaceAfter = 12
                app.Selection.ParagraphFormat.LineSpacing = 12
                app.Selection.Range.Font.Size = 20
                app.Selection.Range.Font.Bold = True
                app.Selection.Range.Font.Underline = True
                p_size += paragraph.Range.Text.Length
                paragraph.Range.InsertParagraphAfter()
            End If
            'Telephone number
            If export_to_word_cb_values(4) Then
                paragraph = doc.Content.Paragraphs.Add
                paragraph.Range.Font.Size = 18
                paragraph.Range.Text = "Номер телефона: " & part.telephone
                app.Selection.Start = p_size
                app.Selection.End = p_size + "Номер телефона:".Length
                app.Selection.ParagraphFormat.SpaceBefore = 0
                app.Selection.ParagraphFormat.SpaceAfter = 12
                app.Selection.ParagraphFormat.LineSpacing = 12
                app.Selection.Range.Font.Size = 20
                app.Selection.Range.Font.Bold = True
                app.Selection.Range.Font.Underline = True
                p_size += paragraph.Range.Text.Length
                paragraph.Range.InsertParagraphAfter()
            End If
            'Post
            If export_to_word_cb_values(5) Then
                paragraph = doc.Content.Paragraphs.Add
                paragraph.Range.Font.Size = 18
                paragraph.Range.Text = "Должность: " & part.post
                app.Selection.Start = p_size
                app.Selection.End = p_size + "Должность:".Length
                app.Selection.ParagraphFormat.SpaceBefore = 0
                app.Selection.ParagraphFormat.SpaceAfter = 12
                app.Selection.ParagraphFormat.LineSpacing = 12
                app.Selection.Range.Font.Size = 20
                app.Selection.Range.Font.Bold = True
                app.Selection.Range.Font.Underline = True
                p_size += paragraph.Range.Text.Length
                paragraph.Range.InsertParagraphAfter()
            End If
            e_code = 5
            'Skills
            If export_to_word_cb_values(6) Then
                Dim output_str(-1) As String
                If export_to_word_cb_values(8) And part.skills.Length > 0 Then
                    Dim pos As Short = -1
                    For i As Short = 0 To part.skills.Length - 1
                        If part.skills(i).ToLower.IndexOf("/dn ") = 0 Then
                            pos = i
                            Exit For
                        End If
                    Next
                    If pos = -1 Then
                        output_str = part.skills
                    Else
                        For i As Short = 0 To part.skills.Length - 1
                            If i <> pos Then
                                Array.Resize(output_str, output_str.Length + 1)
                                output_str(output_str.Length - 1) = part.skills(i)
                            End If
                        Next
                    End If
                Else
                    output_str = part.skills
                End If
                If output_str.Length = 0 Then
                    paragraph = doc.Content.Paragraphs.Add
                    paragraph.Range.Font.Size = 18
                    paragraph.Range.Text = "Навыки: "
                    app.Selection.Start = p_size
                    app.Selection.End = p_size + "Навыки:".Length
                    app.Selection.ParagraphFormat.SpaceBefore = 0
                    app.Selection.ParagraphFormat.SpaceAfter = 12
                    app.Selection.ParagraphFormat.LineSpacing = 12
                    app.Selection.Range.Font.Size = 20
                    app.Selection.Range.Font.Bold = True
                    app.Selection.Range.Font.Underline = True
                    p_size += paragraph.Range.Text.Length
                    paragraph.Range.InsertParagraphAfter()
                ElseIf output_str.Length = 1 Then
                    paragraph = doc.Content.Paragraphs.Add
                    paragraph.Range.Font.Size = 18
                    paragraph.Range.Text = "Навыки: " & output_str(0)
                    app.Selection.Start = p_size
                    app.Selection.End = p_size + "Навыки:".Length
                    app.Selection.ParagraphFormat.SpaceBefore = 0
                    app.Selection.ParagraphFormat.SpaceAfter = 12
                    app.Selection.ParagraphFormat.LineSpacing = 12
                    app.Selection.Range.Font.Size = 20
                    app.Selection.Range.Font.Bold = True
                    app.Selection.Range.Font.Underline = True
                    p_size += paragraph.Range.Text.Length
                    paragraph.Range.InsertParagraphAfter()
                Else
                    For i As Integer = 0 To output_str.Length - 1
                        If i = 0 Then
                            paragraph = doc.Content.Paragraphs.Add
                            paragraph.Range.Font.Size = 18
                            paragraph.Range.Text = "Навыки: " & output_str(i)
                            app.Selection.Start = p_size
                            app.Selection.End = p_size + "Навыки:".Length
                            app.Selection.ParagraphFormat.SpaceBefore = 0
                            app.Selection.ParagraphFormat.SpaceAfter = 0
                            app.Selection.ParagraphFormat.LineSpacing = 12
                            app.Selection.Range.Font.Size = 20
                            app.Selection.Range.Font.Bold = True
                            app.Selection.Range.Font.Underline = True
                            p_size += paragraph.Range.Text.Length
                            paragraph.Range.InsertParagraphAfter()
                        ElseIf i = output_str.Length - 1 Then
                            paragraph = doc.Content.Paragraphs.Add
                            paragraph.Range.Font.Size = 18
                            paragraph.Range.Text = output_str(i)
                            app.Selection.Start = p_size
                            app.Selection.End = p_size
                            app.Selection.ParagraphFormat.SpaceBefore = 0
                            app.Selection.ParagraphFormat.SpaceAfter = 12
                            app.Selection.ParagraphFormat.LineSpacing = 12
                            p_size += paragraph.Range.Text.Length
                            paragraph.Range.InsertParagraphAfter()
                        Else
                            paragraph = doc.Content.Paragraphs.Add
                            paragraph.Range.Font.Size = 18
                            paragraph.Range.Text = output_str(i)
                            app.Selection.Start = p_size
                            app.Selection.End = p_size
                            app.Selection.ParagraphFormat.SpaceBefore = 0
                            app.Selection.ParagraphFormat.SpaceAfter = 0
                            app.Selection.ParagraphFormat.LineSpacing = 12
                            p_size += paragraph.Range.Text.Length
                            paragraph.Range.InsertParagraphAfter()
                        End If
                    Next
                End If
            End If
            'Achievements
            If export_to_word_cb_values(7) Then
                If part.achievements.Length = 0 Then
                    paragraph = doc.Content.Paragraphs.Add
                    paragraph.Range.Font.Size = 18
                    paragraph.Range.Text = "Достижения: "
                    app.Selection.Start = p_size
                    app.Selection.End = p_size + "Достижения:".Length
                    app.Selection.ParagraphFormat.SpaceBefore = 0
                    app.Selection.ParagraphFormat.SpaceAfter = 12
                    app.Selection.ParagraphFormat.LineSpacing = 12
                    app.Selection.Range.Font.Size = 20
                    app.Selection.Range.Font.Bold = True
                    app.Selection.Range.Font.Underline = True
                    p_size += paragraph.Range.Text.Length
                    paragraph.Range.InsertParagraphAfter()
                ElseIf part.achievements.Length = 1 Then
                    paragraph = doc.Content.Paragraphs.Add
                    paragraph.Range.Font.Size = 18
                    paragraph.Range.Text = "Достижения: " & part.achievements(0)
                    app.Selection.Start = p_size
                    app.Selection.End = p_size + "Достижения:".Length
                    app.Selection.ParagraphFormat.SpaceBefore = 0
                    app.Selection.ParagraphFormat.SpaceAfter = 12
                    app.Selection.ParagraphFormat.LineSpacing = 12
                    app.Selection.Range.Font.Size = 20
                    app.Selection.Range.Font.Bold = True
                    app.Selection.Range.Font.Underline = True
                    p_size += paragraph.Range.Text.Length
                    paragraph.Range.InsertParagraphAfter()
                Else
                    For i As Integer = 0 To part.achievements.Length - 1
                        If i = 0 Then
                            paragraph = doc.Content.Paragraphs.Add
                            paragraph.Range.Font.Size = 18
                            paragraph.Range.Text = "Достижения: " & part.achievements(i)
                            app.Selection.Start = p_size
                            app.Selection.End = p_size + "Достижения:".Length
                            app.Selection.ParagraphFormat.SpaceBefore = 0
                            app.Selection.ParagraphFormat.SpaceAfter = 0
                            app.Selection.ParagraphFormat.LineSpacing = 12
                            app.Selection.Range.Font.Size = 20
                            app.Selection.Range.Font.Bold = True
                            app.Selection.Range.Font.Underline = True
                            p_size += paragraph.Range.Text.Length
                            paragraph.Range.InsertParagraphAfter()
                        ElseIf i = part.achievements.Length - 1 Then
                            paragraph = doc.Content.Paragraphs.Add
                            paragraph.Range.Font.Size = 18
                            paragraph.Range.Text = part.achievements(i)
                            app.Selection.Start = p_size
                            app.Selection.End = p_size
                            app.Selection.ParagraphFormat.SpaceBefore = 0
                            app.Selection.ParagraphFormat.SpaceAfter = 12
                            app.Selection.ParagraphFormat.LineSpacing = 12
                            p_size += paragraph.Range.Text.Length
                            paragraph.Range.InsertParagraphAfter()
                        Else
                            paragraph = doc.Content.Paragraphs.Add
                            paragraph.Range.Font.Size = 18
                            paragraph.Range.Text = part.achievements(i)
                            app.Selection.Start = p_size
                            app.Selection.End = p_size
                            app.Selection.ParagraphFormat.SpaceBefore = 0
                            app.Selection.ParagraphFormat.SpaceAfter = 0
                            app.Selection.ParagraphFormat.LineSpacing = 12
                            p_size += paragraph.Range.Text.Length
                            paragraph.Range.InsertParagraphAfter()
                        End If
                    Next
                End If
            End If
            e_code = 4
            'Image print
            If export_to_word_cb_values(11) Then
                Dim photo As New Bitmap(1000, 1000), gr As Graphics = Graphics.FromImage(photo)
                If part.photo = "" Then
                    gr.FillRectangle(New Drawing.SolidBrush(Color.Gainsboro), 0, 0, 1000, 1000)
                    Dim size_image(1) As UShort
                    size_image(0) = (1000 - My.Resources.no_photo.Width) / 2
                    size_image(1) = (1000 - My.Resources.no_photo.Height) / 2
                    gr.DrawImage(My.Resources.no_photo, size_image(0), size_image(1))
                    gr.DrawRectangle(New Drawing.Pen(Color.Black, 4), 2, 2, 996, 996)
                Else
                    gr.FillRectangle(New Drawing.SolidBrush(Color.Gainsboro), 0, 0, 1000, 1000)
                    Dim size_image(1) As UShort
                    size_image(0) = (1000 - part.img.Width) / 2
                    size_image(1) = (1000 - part.img.Height) / 2
                    gr.DrawImage(part.img, size_image(0), size_image(1))
                    gr.DrawRectangle(New Drawing.Pen(Color.Black, 4), 2, 2, 996, 996)
                End If
                photo.Save(photo_way, Imaging.ImageFormat.Jpeg)
                app.Selection.Start = 0
                app.Selection.End = 0
                e_code = 3
                Dim schape As Microsoft.Office.Interop.Word.Shape
                schape = doc.Shapes.AddPicture(photo_way, False, True)
                schape.WrapFormat.Type = Microsoft.Office.Interop.Word.WdWrapType.wdWrapSquare
                schape.Width = 280
                schape.Height = 280
                schape.Top = 0
                schape.Left = 256
                My.Computer.FileSystem.DeleteFile(photo_way)
            End If
            e_code = 2
            doc.SaveAs2(doc_way)
            app.Quit(False)
            e_code = 0
        Catch
        End Try
        Return e_code
    End Function
#End Region
    Dim cmd_service_wait_mode As Short, cmd_service_wait_proc As Integer
    Private Sub cmd_service_wait_Tick(sender As Object, e As EventArgs) Handles cmd_service_wait.Tick
        If cmd_service_wait_mode = 0 Then
            Try
                Process.GetProcessById(cmd_service_wait_proc)
            Catch
                cmd_service_wait.Stop()
                cmd_service_wait_proc = Shell(MY_WAY & "\mcdCAnc.exe ""MCDmode:dir-delete,Code:2802"" """ & TEMP_WAY & "\MCD-temp-47591795970905""")
                cmd_service_wait_mode = 1
                cmd_service_wait.Start()
            End Try
        ElseIf cmd_service_wait_mode = 1 Then
            Try
                Process.GetProcessById(cmd_service_wait_proc)
            Catch
                cmd_service_wait.Stop()
                is_system_closing = True
                Me.Close()
            End Try
        ElseIf cmd_service_wait_mode = 2 Then
            Try
                Process.GetProcessById(cmd_service_wait_proc)
            Catch
                cmd_service_wait.Stop()
                If My.Computer.FileSystem.FileExists(MY_WAY & "\work_result.tmp") Then
                    If My.Computer.FileSystem.ReadAllText(MY_WAY & "\work_result.tmp") = "sucess" Then
                        Messanger.Renovate_messange("db loading", "Идёт подготовка базы данных Участников...")
                        DB1Analyser_copy_db = False
                        DB1Analyser.RunWorkerAsync()
                        My.Computer.FileSystem.DeleteFile(MY_WAY & "\work_result.tmp")
                    Else
                        Messanger.Delete_messange(Messanger.Get_position_by_link("db loading"))
                        Messanger.Add_messange("База данных Участников не загружена!", "db not loaded", True)
                        My.Computer.FileSystem.DeleteFile(MY_WAY & "\work_result.tmp")
                        DB_load_err()
                    End If

                Else
                    Messanger.Delete_messange(Messanger.Get_position_by_link("db loading"))
                    Messanger.Add_messange("База данных Участников не загружена!", "db not loaded", True)
                    DB_load_err()
                End If
            End Try
        ElseIf cmd_service_wait_mode = 3 Then
            Try
                Process.GetProcessById(cmd_service_wait_proc)
            Catch
                cmd_service_wait.Stop()
                If My.Computer.FileSystem.FileExists(MY_WAY & "\work_result.tmp") Then
                    If My.Computer.FileSystem.ReadAllText(MY_WAY & "\work_result.tmp") = "sucess" Then
                        Dim last_pos As Short = TabControl1.SelectedIndex
                        TabControl1.SelectedIndex = 3
                        Renovate_DataGrid(0)
                        TabControl1.SelectedIndex = last_pos
                        PictureBox3.Visible = False
                        Label24.Visible = False
                        Messanger.Delete_messange(Messanger.Get_position_by_link("db loading"))
                        If Mes_set.CheckBox8.Checked Then
                            Messanger.Add_messange("База данных Участников загружена.", "db loaded", True)
                        End If
                        My.Computer.FileSystem.DeleteFile(MY_WAY & "\work_result.tmp")
                    Else
                        Messanger.Delete_messange(Messanger.Get_position_by_link("db loading"))
                        Messanger.Add_messange("База данных Участников не загружена!", "db not loaded", True)
                        My.Computer.FileSystem.DeleteFile(MY_WAY & "\work_result.tmp")
                        DB_load_err()
                    End If
                Else
                    Messanger.Delete_messange(Messanger.Get_position_by_link("db loading"))
                    Messanger.Add_messange("База данных Участников не загружена!", "db not loaded", True)
                    DB_load_err()
                End If
            End Try
        ElseIf cmd_service_wait_mode = 4 Then
            Try
                Process.GetProcessById(cmd_service_wait_proc)
            Catch
                cmd_service_wait.Stop()
                If My.Computer.FileSystem.FileExists(MY_WAY & "\work_result.tmp") Then
                    If My.Computer.FileSystem.ReadAllText(MY_WAY & "\work_result.tmp") = "sucess" Then
                        db_not_saved = False
                        Help.SetToolTip(PictureBox2, "Изменения в базе данных сохранены.")
                        PictureBox2.Image = My.Resources.saved
                        Button33.Enabled = True
                        CheckBox1.Enabled = True
                        If Mes_set.CheckBox8.Checked Then
                            Messanger.Add_messange("База данных Участников сохранена.", "db loaded", True)
                        End If
                    Else
                        MsgBox("При работе с базой данных Участников произошла ошибка!" & Chr(10) & "Совет: попробуйте перезапустить программу.",, "Ошибка")
                        db_not_saved = True
                        Help.SetToolTip(PictureBox2, "Изменения в базе данных не сохранены.")
                        PictureBox2.Image = My.Resources.not_saved
                        Messanger.Delete_messange(Messanger.Get_position_by_link("db loading"))
                        Messanger.Add_messange("База данных Участников не сохранена!", "db not loaded", True)
                    End If
                    My.Computer.FileSystem.DeleteFile(MY_WAY & "\work_result.tmp")
                Else
                    MsgBox("При работе с базой данных Участников произошла ошибка!" & Chr(10) & "Совет: попробуйте перезапустить программу.",, "Ошибка")
                    db_not_saved = True
                    Help.SetToolTip(PictureBox2, "Изменения в базе данных не сохранены.")
                    PictureBox2.Image = My.Resources.not_saved
                    Messanger.Delete_messange(Messanger.Get_position_by_link("db loading"))
                    Messanger.Add_messange("База данных Участников не загружена!", "db not loaded", True)
                End If
            End Try
        End If
    End Sub
    Private Function Generate_code() As String
        Dim r As New Random, result As String = "", last_num As Short = -1
        For i As Short = 0 To 9
            Dim r_num As Short = r.Next(10)
            While last_num = r_num
                r_num = r.Next(10)
            End While
            result &= r_num
        Next
        Return result
    End Function
    Public Shared Function Get_date_string(t_date As Date) As String
        Dim result As String = t_date.Day & "." & t_date.Month & "." & t_date.Year & " " & t_date.Hour & ":" & t_date.Minute
        Return result
    End Function
    Public Function Get_hms(time) As hms_temp
        Dim left As Integer = time
        Dim hours_left As Integer = Int((left / 60) / 60)
        left -= hours_left * 3600
        Dim minutes_left As Integer = Int(left / 60)
        left -= minutes_left * 60
        Return New hms_temp(hours_left, minutes_left, left)
    End Function
    Public Shared Function Banned_simvols_test(text As String) As Boolean
        If text.IndexOf("<") = -1 And text.IndexOf(">") = -1 And text.IndexOf("\") = -1 And text.IndexOf("=") = -1 And text.IndexOf("[") = -1 And text.IndexOf("]") = -1 And text.IndexOf("|") = -1 And text.IndexOf("&") = -1 Then
            Return True
        Else
            Return False
        End If
    End Function
    Public Shared Function Clone_dbPart_Class(ByVal sourse As db_Participator) As db_Participator
        Dim new_class As New db_Participator
        new_class.key = sourse.key
        new_class.name = sourse.name
        new_class.year = sourse.year
        new_class.school = sourse.school
        new_class.class_ = sourse.class_
        new_class.telephone = sourse.telephone
        new_class.post = sourse.post
        new_class.skills = sourse.skills
        new_class.achievements = sourse.achievements
        new_class.photo = sourse.photo
        new_class.img = sourse.img
        Return new_class
    End Function
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
            If Main.messanges(position).read_only = True Or link = "" Then
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
            max_widht = (Int(max_widht / 40) + 1) * 40
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
Public Class Timer_page
    Public name As String = "", new_interval As UInteger = 0, work_messange(4) As Boolean, select_time(2) As UShort 'work_messange(0) - timer tick messange, work_messange(1) - timer start/stop messange, work_messange(2) - color select, work_messange(3) - sound select, work_messange(4) - double sound select
    Dim interval1 As UInteger = 0
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
End Class
Public Structure hms_temp
    Public hours, minutes, seconds As UShort
    Public Sub New(hours As UShort, minutes As UShort, seconds As UShort)
        Me.hours = hours
        Me.minutes = minutes
        Me.seconds = seconds
    End Sub
End Structure
Public Class Member
    Public name As String = "", dir As UInteger = 0
End Class
Public Class Command
    Public name As String = "", dir As UInteger = 0, members As New List(Of Member)
End Class
Public Class Tournier
    Public name As String = "", dir As UInteger = 0, commands As New List(Of Command)
End Class
Public Class db_info_file
    Public format As String = "", versions(-1) As String
End Class
Public Class db_collector_file
    Public LastReadTime As String = "", LastWriteTime As String = "", progID As String = "", user_name, organization As String, participators As New ArrayList
End Class
Public Class db_Participator
    Public key, name, year, school, class_, telephone, post, skills(), achievements(), photo As String, img As Bitmap
End Class

