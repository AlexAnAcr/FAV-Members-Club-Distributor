Public Class Member_to_word
    Private Sub Member_to_word_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        CheckBox1.Checked = True
        CheckBox2.Checked = True
        CheckBox3.Checked = True
        CheckBox4.Checked = True
        CheckBox5.Checked = True
        CheckBox6.Checked = True
        CheckBox7.Checked = True
        CheckBox8.Checked = True
        CheckBox9.Checked = True
        TextBox1.Text = ""
        CheckBox10.Checked = False
        ComboBox1.Enabled = False
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
        CheckBox11.Enabled = False
        CheckBox11.Checked = False
        Label2.Text = "Экспорт элементов: " & Main.DataGridView1.SelectedRows.Count & " шт."
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
    Dim cb_values() As Boolean
    Private Sub LinkLabel1_Click(sender As Object, e As EventArgs) Handles LinkLabel1.Click
        MsgBox("Примечание:" & Chr(10) & "Если для экспорта в Word Вы выбрали более одного элемента, то для сохранения будет использована папка, в которой находится документ, к которому указан путь в поле ""Путь к сохраняемому файлу"". В этом случае рекомендуем настроить параметры раздела ""Имена сохраняемых документов"".",, "Информация")
    End Sub
    Private Sub LinkLabel2_Click(sender As Object, e As EventArgs) Handles LinkLabel2.Click
        If ComboBox2.SelectedIndex = 0 Then
            MsgBox("Файлы word, имеющиеся в выходной папке, не будут удалены. Сохраняемые файлы не заменят имеющиеся в папке файлы.",, "Информация")
        ElseIf ComboBox2.SelectedIndex = 1 Then
            MsgBox("Файлы word, имеющиеся в выходной папке, имя которых совпадает с именем сохраняемого файла, будут удалены. Сохраняемые файлы заменят имеющиеся в папке word (.docx) с совпадающими именами.",, "Информация")
        ElseIf ComboBox2.SelectedIndex = 2 Then
            MsgBox("Файлы word, имеющиеся в выходной папке, будут удалены. Сохраняемые файлы полностью заменят имеющиеся в папке файлы word (.docx).",, "Информация")
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Then
            MsgBox("Вы не выбрали папку файла!",, "Ошибка")
            Exit Sub
        End If
        If My.Computer.FileSystem.DirectoryExists(Mid(TextBox1.Text, 1, TextBox1.Text.LastIndexOf("\"))) = False Then
            MsgBox("Выбранная папка файла не существует!",, "Ошибка")
            Exit Sub
        End If
        If Mid(TextBox1.Text, TextBox1.Text.LastIndexOf(".") + 2) <> "docx" Then
            MsgBox("Расширение сохраняемого файла не соответствует расширению "".docx""!",, "Ошибка")
            Exit Sub
        End If
        Panel1.Enabled = False
        If Main.DataGridView1.SelectedRows.Count >= 100 Then
            Progress.progress_mode(0) = 1
            Progress.progress_mode(1) = Main.DataGridView1.SelectedRows.Count / 100
        Else
            Progress.progress_mode(0) = 0
            Progress.progress_mode(1) = 100 / Main.DataGridView1.SelectedRows.Count
        End If
        Progress.Reset()
        Progress.Set_parent_(Me)
        Progress.Show()
        Main.export_to_word_progress = 0
        Main.Expor_to_word_start()
        ExportWordTimer.Start()
    End Sub
    Public Sub Export_to_word_ended()
        ExportWordTimer.Stop()
        Progress.Close()
        Panel1.Enabled = True
    End Sub
    Private Sub CheckBox10_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox10.CheckedChanged
        If CheckBox10.Checked Then
            ComboBox1.Enabled = True
            CheckBox11.Enabled = True
        Else
            ComboBox1.Enabled = False
            CheckBox11.Enabled = False
        End If
    End Sub
    Private Class export_to_word_link
        Public index As String = "", err As UShort = 0, doc_index As UShort = 1
        Public Sub New(index As String, err As UShort, doc_index As UShort)
            index = index
            err = err
            doc_index = doc_index
        End Sub
    End Class
    Dim active_cicle As UShort = 0
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim result As DialogResult = SaveFile1.ShowDialog
        If result = DialogResult.OK Then
            TextBox1.Text = SaveFile1.FileName
        End If
    End Sub
    Private Sub ExportWordTimer_Tick(sender As Object, e As EventArgs) Handles ExportWordTimer.Tick
        If 0 < Main.export_to_word_progress Then
            For i As UShort = 0 To Main.export_to_word_progress - 1
                Progress.Progress_tick()
            Next
            Main.export_to_word_progress = 0
        End If
        If active_cicle = 5 Then
            active_cicle = 0
            Try
                If ActiveForm.Name = "Member_to_word" Then
                    Progress.Activate()
                End If
            Catch
            End Try
        Else
            active_cicle += 1
        End If
    End Sub
    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        If CheckBox6.Checked Then
            CheckBox9.Enabled = True
        Else
            CheckBox9.Enabled = False
        End If
    End Sub
End Class
