Public Class Add_Redact_db_data
    Public Shared work_mode As Short = 0
    Public p_temp As db_Participator
    Dim text_boxes() As TextBox
    Private Sub Add_Redact_db_data_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        text_boxes = {TextBox1, TextBox2, TextBox3, TextBox4, TextBox5, TextBox6, TextBox7, TextBox8}
        If work_mode = 1 Then
            p_temp = New db_Participator
            Panel1.Visible = True
            Panel2.Visible = False
            PictureBox1.Image = My.Resources.no_photo
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            TextBox5.Text = ""
            TextBox6.Text = ""
            TextBox7.Text = ""
            TextBox8.Text = ""
            p_temp.photo = ""
            p_temp.key = ""
            Me.Text = "Добавление участника"
            ContextMenu1.Enabled = False
            For Each i As TextBox In text_boxes
                i.ReadOnly = False
            Next
            saved = False
        ElseIf work_mode = 2 Then
            Dim pos As Short = -1
            For i As Short = 0 To Main.db_data.Count - 1
                If Main.db_data(i).key = Main.DataGridView1.SelectedRows.Item(0).Cells.Item(0).Value Then
                    pos = i
                    Exit For
                End If
            Next
            For Each i As TextBox In text_boxes
                i.ReadOnly = False
            Next
            Me.Text = "Изменение данных участника"
            If pos > -1 Then
                p_temp = Main.Clone_dbPart_Class(Main.db_data(pos))
                TextBox1.Text = p_temp.name
                TextBox2.Text = p_temp.year
                TextBox3.Text = p_temp.school
                TextBox4.Text = p_temp.class_
                TextBox5.Text = p_temp.telephone
                TextBox6.Text = p_temp.post
                TextBox7.Lines = p_temp.skills
                TextBox8.Lines = p_temp.achievements
                If p_temp.photo <> "" Then
                    If p_temp.img.Width > 256 Or p_temp.img.Height > 256 Then
                        PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
                    Else
                        PictureBox1.SizeMode = PictureBoxSizeMode.CenterImage
                    End If
                    PictureBox1.Image = p_temp.img
                    ContextMenu1.Enabled = True
                Else
                    PictureBox1.Image = My.Resources.no_photo
                    ContextMenu1.Enabled = False
                End If
                Panel1.Visible = True
                Panel2.Visible = False
            End If
            saved = False
        ElseIf work_mode = 3 Then
            Dim pos As Short = -1
            For i As Short = 0 To Main.db_data.Count - 1
                If Main.db_data(i).key = Main.DataGridView1.SelectedRows.Item(0).Cells.Item(0).Value Then
                    pos = i
                    Exit For
                End If
            Next
            For Each i As TextBox In text_boxes
                i.ReadOnly = True
            Next
            Me.Text = "Просмотр данных участника"
            If pos > -1 Then
                p_temp = Main.Clone_dbPart_Class(Main.db_data(pos))
                TextBox1.Text = p_temp.name
                TextBox2.Text = p_temp.year
                TextBox3.Text = p_temp.school
                TextBox4.Text = p_temp.class_
                TextBox5.Text = p_temp.telephone
                TextBox6.Text = p_temp.post
                TextBox7.Lines = p_temp.skills
                TextBox8.Lines = p_temp.achievements
                If p_temp.photo <> "" Then
                    If p_temp.img.Width > 256 Or p_temp.img.Height > 256 Then
                        PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
                    Else
                        PictureBox1.SizeMode = PictureBoxSizeMode.CenterImage
                    End If
                    PictureBox1.Image = p_temp.img
                    ContextMenu1.Enabled = True
                Else
                    PictureBox1.Image = My.Resources.no_photo
                    ContextMenu1.Enabled = False
                End If
                Panel1.Visible = False
                Panel2.Visible = True
            End If
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        saved = True
        Me.Close()
    End Sub
    Dim saved As Boolean = False
    Private Sub Add_Redact_db_data_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If saved = False Then
            If work_mode = 1 Then
                If TextBox1.Text <> "" Or TextBox2.Text <> "" Or TextBox3.Text <> "" Or TextBox4.Text <> "" Or TextBox5.Text <> "" Or TextBox6.Text <> "" Or TextBox7.Text <> "" Or TextBox8.Text <> "" Or p_temp.photo <> "" Then
                    Dim result As MsgBoxResult = MsgBox("Вы уверены, что хотите выйти без сохранения?", MsgBoxStyle.YesNo, "Сообщение")
                    If result = MsgBoxResult.No Then
                        e.Cancel = True
                    End If
                End If
            ElseIf work_mode = 2 Then
                Dim pos As Short = -1
                For i As Short = 0 To Main.db_data.Count - 1
                    If Main.db_data(i).key = Main.DataGridView1.SelectedRows.Item(0).Cells.Item(0).Value Then
                        pos = i
                        Exit For
                    End If
                Next
                If p_temp.name <> TextBox1.Text Or p_temp.year <> TextBox2.Text Or p_temp.school <> TextBox3.Text Or p_temp.class_ <> TextBox4.Text Or p_temp.telephone <> TextBox5.Text Or p_temp.post <> TextBox6.Text Or ArraysEqual(p_temp.skills, TextBox7.Lines) = False Or ArraysEqual(p_temp.achievements, TextBox8.Lines) = False Or Main.db_data(pos).photo <> p_temp.photo Or Main.db_data(pos).img IsNot p_temp.img Then
                    Dim result As MsgBoxResult = MsgBox("Вы уверены, что хотите выйти без сохранения?", MsgBoxStyle.YesNo, "Сообщение")
                    If result = MsgBoxResult.No Then
                        e.Cancel = True
                    End If
                End If
            End If
        End If
    End Sub
    Function ArraysEqual(primArray() As String, secondArray() As String) As Boolean
        Dim diff As Boolean = True
        If primArray.Length = secondArray.Length Then
            For i As Integer = 0 To primArray.Length - 1
                If primArray(i) <> secondArray(i) Then
                    diff = False
                    Exit For
                End If
            Next
        Else
            If Math.Max(primArray.Length, secondArray.Length) - Math.Min(primArray.Length, secondArray.Length) = 1 Then
                If primArray.Length = 0 Then
                    If secondArray(0) <> "" Then
                        diff = False
                    End If
                ElseIf secondArray.Length = 0 Then
                    If primArray(0) <> "" Then
                        diff = False
                    End If
                Else
                    diff = False
                End If
            Else
                diff = False
            End If
        End If
        Return diff
    End Function
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If work_mode = 1 Then
            If Main.Banned_simvols_test(TextBox1.Text) = False Or TextBox1.Text = "" Then
                MsgBox("Поле ""Имя"" содержит запрещённые символы или пустое!" & Chr(10) & "К запрещённым символам относятся: <>\=[]|&", MsgBoxStyle.OkOnly, "Сообщение")
                Exit Sub
            End If
            If Main.Banned_simvols_test(TextBox2.Text) = False Or TextBox2.Text = "" Then
                MsgBox("Поле ""Дата рождения"" содержит запрещённые символы или пустое!" & Chr(10) & "К запрещённым символам относятся: <>\=[]|&", MsgBoxStyle.OkOnly, "Сообщение")
                Exit Sub
            End If
            If Main.Banned_simvols_test(TextBox3.Text) = False Or TextBox3.Text = "" Then
                MsgBox("Поле ""Школа"" содержит запрещённые символы или пустое!" & Chr(10) & "К запрещённым символам относятся: <>\=[]|&", MsgBoxStyle.OkOnly, "Сообщение")
                Exit Sub
            End If
            If Main.Banned_simvols_test(TextBox4.Text) = False Or TextBox4.Text = "" Then
                MsgBox("Поле ""Класс"" содержит запрещённые символы или пустое!" & Chr(10) & "К запрещённым символам относятся: <>\=[]|&", MsgBoxStyle.OkOnly, "Сообщение")
                Exit Sub
            End If
            If Main.Banned_simvols_test(TextBox5.Text) = False Then
                MsgBox("Поле ""Номер телефона"" содержит запрещённые символы!" & Chr(10) & "К запрещённым символам относятся: <>\=[]|&", MsgBoxStyle.OkOnly, "Сообщение")
                Exit Sub
            End If
            If Main.Banned_simvols_test(TextBox6.Text) = False Then
                MsgBox("Поле ""Должность"" содержит запрещённые символы!" & Chr(10) & "К запрещённым символам относятся: <>\=[]|&", MsgBoxStyle.OkOnly, "Сообщение")
                Exit Sub
            End If
            If Main.Banned_simvols_test(TextBox7.Text) = False Then
                MsgBox("Поле ""Навыки"" содержит запрещённые символы!" & Chr(10) & "К запрещённым символам относятся: <>\=[]|&", MsgBoxStyle.OkOnly, "Сообщение")
                Exit Sub
            End If
            If Main.Banned_simvols_test(TextBox8.Text) = False Then
                MsgBox("Поле ""Достижения"" содержит запрещённые символы!" & Chr(10) & "К запрещённым символам относятся: <>\=[]|&", MsgBoxStyle.OkOnly, "Сообщение")
                Exit Sub
            End If
            Dim index As Short = 1, finded As Boolean = True
            While finded = True
                If index.ToString.Length < 3 Then
                    finded = Main.db_data_collector_exist(Format(index, "000"))
                    If finded = True Then
                        index += 1
                    End If
                Else
                    finded = Main.db_data_collector_exist(index)
                    If finded = True Then
                        index += 1
                    End If
                End If
            End While
            If index.ToString.Length < 3 Then
                p_temp.key = Format(index, "000")
            Else
                p_temp.key = index.ToString
            End If
            p_temp.name = TextBox1.Text
            p_temp.year = TextBox2.Text
            p_temp.school = TextBox3.Text
            p_temp.class_ = TextBox4.Text
            p_temp.telephone = TextBox5.Text
            p_temp.post = TextBox6.Text
            If TextBox7.Lines.Length = 0 Then
                p_temp.skills = {""}
            Else
                p_temp.skills = TextBox7.Lines
            End If
            If TextBox8.Lines.Length = 0 Then
                p_temp.achievements = {""}
            Else
                p_temp.achievements = TextBox8.Lines
            End If
            If p_temp.photo <> "" Then
                p_temp.photo = p_temp.key
                p_temp.img.Save(Main.TEMP_WAY & "\MCD-temp-47591795970905\database\" & p_temp.photo & ".jpg", Imaging.ImageFormat.Jpeg)
            Else
                p_temp.img = My.Resources.no_photo
            End If
            Main.db_data.Add(p_temp)
            Main.dbf_collector.participators.Add(p_temp.key)
            Main.DB_WriteDataFile(p_temp.key, False)
            Main.DB_WriteCollectorFile()
            If Main.search_mode = False Then
                'Renovate datagrid
                Main.DataSet1.Tables(0).Rows.Add(p_temp.img, p_temp.name, p_temp.year, p_temp.school, p_temp.telephone, p_temp.post, p_temp.key)
                Main.DataGridView1.Columns.Item(0).Visible = False
                Main.DataGridView1.AutoResizeColumn(2)
                Main.DataGridView1.AutoResizeColumn(3)
                Main.DataGridView1.AutoResizeColumn(4)
                Main.DataGridView1.AutoResizeColumn(5)
                Main.DataGridView1.AutoResizeColumn(6)
            End If
            If Main.CheckBox1.Checked Then
                Main.Button_save_db()
            Else
                Main.db_not_saved = True
                Main.Help.SetToolTip(Main.PictureBox2, "Изменения в базе данных не сохранены.")
                Main.PictureBox2.Image = My.Resources.not_saved
            End If
        ElseIf work_mode = 2 Then
            If Main.Banned_simvols_test(TextBox1.Text) = False Or TextBox1.Text = "" Then
                MsgBox("Поле ""Имя"" содержит запрещённые символы или пустое!" & Chr(10) & "К запрещённым символам относятся: <>\=[]|&", MsgBoxStyle.OkOnly, "Сообщение")
                Exit Sub
            End If
            If Main.Banned_simvols_test(TextBox2.Text) = False Or TextBox2.Text = "" Then
                MsgBox("Поле ""Дата рождения"" содержит запрещённые символы или пустое!" & Chr(10) & "К запрещённым символам относятся: <>\=[]|&", MsgBoxStyle.OkOnly, "Сообщение")
                Exit Sub
            End If
            If Main.Banned_simvols_test(TextBox3.Text) = False Or TextBox3.Text = "" Then
                MsgBox("Поле ""Школа"" содержит запрещённые символы или пустое!" & Chr(10) & "К запрещённым символам относятся: <>\=[]|&", MsgBoxStyle.OkOnly, "Сообщение")
                Exit Sub
            End If
            If Main.Banned_simvols_test(TextBox4.Text) = False Or TextBox4.Text = "" Then
                MsgBox("Поле ""Класс"" содержит запрещённые символы или пустое!" & Chr(10) & "К запрещённым символам относятся: <>\=[]|&", MsgBoxStyle.OkOnly, "Сообщение")
                Exit Sub
            End If
            If Main.Banned_simvols_test(TextBox5.Text) = False Then
                MsgBox("Поле ""Номер телефона"" содержит запрещённые символы!" & Chr(10) & "К запрещённым символам относятся: <>\=[]|&", MsgBoxStyle.OkOnly, "Сообщение")
                Exit Sub
            End If
            If Main.Banned_simvols_test(TextBox6.Text) = False Then
                MsgBox("Поле ""Должность"" содержит запрещённые символы!" & Chr(10) & "К запрещённым символам относятся: <>\=[]|&", MsgBoxStyle.OkOnly, "Сообщение")
                Exit Sub
            End If
            If Main.Banned_simvols_test(TextBox7.Text) = False Then
                MsgBox("Поле ""Навыки"" содержит запрещённые символы!" & Chr(10) & "К запрещённым символам относятся: <>\=[]|&", MsgBoxStyle.OkOnly, "Сообщение")
                Exit Sub
            End If
            If Main.Banned_simvols_test(TextBox8.Text) = False Then
                MsgBox("Поле ""Достижения"" содержит запрещённые символы!" & Chr(10) & "К запрещённым символам относятся: <>\=[]|&", MsgBoxStyle.OkOnly, "Сообщение")
                Exit Sub
            End If
            Dim pos As Short = -1
            For i As Short = 0 To Main.db_data.Count - 1
                If Main.db_data(i).key = Main.DataGridView1.SelectedRows.Item(0).Cells.Item(0).Value Then
                    pos = i
                    Exit For
                End If
            Next
            If p_temp.name <> TextBox1.Text Or p_temp.year <> TextBox2.Text Or p_temp.school <> TextBox3.Text Or p_temp.class_ <> TextBox4.Text Or p_temp.telephone <> TextBox5.Text Or p_temp.post <> TextBox6.Text Or ArraysEqual(p_temp.skills, TextBox7.Lines) = False Or ArraysEqual(p_temp.achievements, TextBox8.Lines) = False Or Main.db_data(pos).photo <> p_temp.photo Or Main.db_data(pos).img IsNot p_temp.img Then
                p_temp.name = TextBox1.Text
                p_temp.year = TextBox2.Text
                p_temp.school = TextBox3.Text
                p_temp.class_ = TextBox4.Text
                p_temp.telephone = TextBox5.Text
                p_temp.post = TextBox6.Text
                If TextBox7.Lines.Length = 0 Then
                    p_temp.skills = {""}
                Else
                    p_temp.skills = TextBox7.Lines
                End If
                If TextBox8.Lines.Length = 0 Then
                    p_temp.achievements = {""}
                Else
                    p_temp.achievements = TextBox8.Lines
                End If
                If p_temp.photo <> "" Then
                    p_temp.img.Save(Main.TEMP_WAY & "\MCD-temp-47591795970905\database\" & p_temp.photo & ".jpg", Imaging.ImageFormat.Jpeg)
                Else
                    If Main.db_data(pos).photo <> p_temp.photo Then
                        If My.Computer.FileSystem.FileExists(Main.TEMP_WAY & "\MCD-temp-47591795970905\database\" & Main.db_data(pos).photo & ".jpg") Then
                            My.Computer.FileSystem.DeleteFile(Main.TEMP_WAY & "\MCD-temp-47591795970905\database\" & Main.db_data(pos).photo & ".jpg")
                        End If
                    End If
                    p_temp.img = My.Resources.no_photo
                End If
                Main.db_data(pos) = p_temp
                Main.DB_WriteDataFile(p_temp.key, False)
                If Main.search_mode = False Then
                    'Renovate datagrid
                    For i As Short = 0 To Main.sdb_keys.Count - 1
                        If Main.sdb_keys(i) = p_temp.key Then
                            Main.sdb_keys.RemoveAt(i)
                            Exit For
                        End If
                    Next
                    Main.DataSet1.Tables(0).Rows.Item(pos).ItemArray = {p_temp.img, p_temp.name, p_temp.year, p_temp.school, p_temp.telephone, p_temp.post, p_temp.key}
                    Main.DataGridView1.Columns.Item(0).Visible = False
                    Main.DataGridView1.AutoResizeColumn(2)
                    Main.DataGridView1.AutoResizeColumn(3)
                    Main.DataGridView1.AutoResizeColumn(4)
                    Main.DataGridView1.AutoResizeColumn(5)
                    Main.DataGridView1.AutoResizeColumn(6)
                Else
                    Main.sdb_keys.Remove(p_temp.key)
                    For i As Short = 0 To Main.db_data.Count - 1
                        If p_temp.key = Main.DataSet1.Tables(0).Rows(i).Item(6) Then
                            Main.DataSet1.Tables(0).Rows.RemoveAt(i)
                            Exit For
                        End If
                    Next
                    Main.DataGridView1.Columns.Item(0).Visible = False
                    Main.DataGridView1.AutoResizeColumn(2)
                    Main.DataGridView1.AutoResizeColumn(3)
                    Main.DataGridView1.AutoResizeColumn(4)
                    Main.DataGridView1.AutoResizeColumn(5)
                    Main.DataGridView1.AutoResizeColumn(6)
                End If
                If Main.CheckBox1.Checked Then
                    Main.Button_save_db()
                Else
                    Main.db_not_saved = True
                    Main.Help.SetToolTip(Main.PictureBox2, "Изменения в базе данных не сохранены.")
                    Main.PictureBox2.Image = My.Resources.not_saved
                End If
            End If
        End If
        saved = True
        Me.Close()
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        saved = True
        Me.Close()
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        work_mode = 2
        Dim pos As Short = -1
        For i As Short = 0 To Main.db_data.Count - 1
            If Main.db_data(i).key = Main.DataGridView1.SelectedRows.Item(0).Cells.Item(0).Value Then
                pos = i
                Exit For
            End If
        Next
        For Each i As TextBox In text_boxes
            i.ReadOnly = False
        Next
        Me.Text = "Изменение данных участника"
        p_temp = Main.Clone_dbPart_Class(Main.db_data(pos))
        TextBox1.Text = p_temp.name
        TextBox2.Text = p_temp.year
        TextBox3.Text = p_temp.school
        TextBox4.Text = p_temp.class_
        TextBox5.Text = p_temp.telephone
        TextBox6.Text = p_temp.post
        TextBox7.Lines = p_temp.skills
        TextBox8.Lines = p_temp.achievements
        If p_temp.photo <> "" Then
            If p_temp.img.Width > 256 Or p_temp.img.Height > 256 Then
                PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
            Else
                PictureBox1.SizeMode = PictureBoxSizeMode.CenterImage
            End If
            PictureBox1.Image = p_temp.img
        Else
            PictureBox1.Image = My.Resources.no_photo
        End If
        Panel1.Visible = True
        Panel2.Visible = False
        saved = False
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Image_Redact.ShowDialog()
        If p_temp.photo <> "" Then
            If p_temp.img.Width > 256 Or p_temp.img.Height > 256 Then
                PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
            Else
                PictureBox1.SizeMode = PictureBoxSizeMode.CenterImage
            End If
            PictureBox1.Image = p_temp.img
            ContextMenu1.Enabled = True
        Else
            PictureBox1.Image = My.Resources.no_photo
            ContextMenu1.Enabled = False
        End If
    End Sub
    Private Sub СохранитьИзображениеКакToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles СохранитьИзображениеКакToolStripMenuItem.Click
        Dim result As DialogResult = SaveFile1.ShowDialog
        If result = DialogResult.OK Then
            Try
                p_temp.img.Save(SaveFile1.FileName, Imaging.ImageFormat.Jpeg)
                MsgBox("Изображение сохранено успешно!",, "Сообщение")
            Catch
                MsgBox("Не удалось сохранить изображение!",, "Ошибка")
            End Try
        End If
    End Sub
End Class
