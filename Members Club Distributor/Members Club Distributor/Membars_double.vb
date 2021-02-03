Public Class Membars_double
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
    Private Sub Membars_double_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        CheckBox1.Checked = False
        ComboBox1.SelectedIndex = 0
        ComboBox1.Enabled = False
        TextBox1.Text = ""
        TextBox1.Enabled = False
        Label1.Text = "Дублирование элементов: " & Main.DataGridView1.SelectedRows.Count & " шт."
    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            ComboBox1.Enabled = True
            TextBox1.Enabled = True
        Else
            ComboBox1.Enabled = False
            TextBox1.Enabled = False
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim indexes(-1) As String, added As Short = 0
        For Each i As DataGridViewRow In Main.DataGridView1.SelectedRows
            Array.Resize(indexes, indexes.Length + 1)
            indexes(indexes.Length - 1) = i.Cells.Item(0).Value
        Next
        For i As Short = 0 To indexes.Length - 1
            Dim pos As Short = -1
            For i1 As Short = 0 To Main.db_data.Count - 1
                If Main.db_data(i1).key = indexes(i) Then
                    pos = i1
                    Exit For
                End If
            Next
            Dim index As Short = 1, finded As Boolean = True, p_temp As New db_Participator
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
            If CheckBox1.Checked = False Then
                p_temp.name = Main.db_data(pos).name
                p_temp.year = Main.db_data(pos).year
                p_temp.school = Main.db_data(pos).school
                p_temp.class_ = Main.db_data(pos).class_
                p_temp.telephone = Main.db_data(pos).telephone
                p_temp.post = Main.db_data(pos).post
                p_temp.skills = Main.db_data(pos).skills
                p_temp.achievements = Main.db_data(pos).achievements
            Else
                If ComboBox1.SelectedIndex = 0 Then
                    p_temp.name = Main.db_data(pos).name & TextBox1.Text
                    p_temp.year = Main.db_data(pos).year
                    p_temp.school = Main.db_data(pos).school
                    p_temp.class_ = Main.db_data(pos).class_
                    p_temp.telephone = Main.db_data(pos).telephone
                    p_temp.post = Main.db_data(pos).post
                    p_temp.skills = Main.db_data(pos).skills
                    p_temp.achievements = Main.db_data(pos).achievements
                ElseIf ComboBox1.SelectedIndex = 1 Then
                    p_temp.name = Main.db_data(pos).name
                    p_temp.year = Main.db_data(pos).year & TextBox1.Text
                    p_temp.school = Main.db_data(pos).school
                    p_temp.class_ = Main.db_data(pos).class_
                    p_temp.telephone = Main.db_data(pos).telephone
                    p_temp.post = Main.db_data(pos).post
                    p_temp.skills = Main.db_data(pos).skills
                    p_temp.achievements = Main.db_data(pos).achievements
                ElseIf ComboBox1.SelectedIndex = 2 Then
                    p_temp.name = Main.db_data(pos).name
                    p_temp.year = Main.db_data(pos).year
                    p_temp.school = Main.db_data(pos).school & TextBox1.Text
                    p_temp.class_ = Main.db_data(pos).class_
                    p_temp.telephone = Main.db_data(pos).telephone
                    p_temp.post = Main.db_data(pos).post
                    p_temp.skills = Main.db_data(pos).skills
                    p_temp.achievements = Main.db_data(pos).achievements
                ElseIf ComboBox1.SelectedIndex = 3 Then
                    p_temp.name = Main.db_data(pos).name
                    p_temp.year = Main.db_data(pos).year
                    p_temp.school = Main.db_data(pos).school
                    p_temp.class_ = Main.db_data(pos).class_ & TextBox1.Text
                    p_temp.telephone = Main.db_data(pos).telephone
                    p_temp.post = Main.db_data(pos).post
                    p_temp.skills = Main.db_data(pos).skills
                    p_temp.achievements = Main.db_data(pos).achievements
                ElseIf ComboBox1.SelectedIndex = 4 Then
                    p_temp.name = Main.db_data(pos).name
                    p_temp.year = Main.db_data(pos).year
                    p_temp.school = Main.db_data(pos).school
                    p_temp.class_ = Main.db_data(pos).class_
                    p_temp.telephone = Main.db_data(pos).telephone & TextBox1.Text
                    p_temp.post = Main.db_data(pos).post
                    p_temp.skills = Main.db_data(pos).skills
                    p_temp.achievements = Main.db_data(pos).achievements
                ElseIf ComboBox1.SelectedIndex = 5 Then
                    p_temp.name = Main.db_data(pos).name
                    p_temp.year = Main.db_data(pos).year
                    p_temp.school = Main.db_data(pos).school
                    p_temp.class_ = Main.db_data(pos).class_
                    p_temp.telephone = Main.db_data(pos).telephone
                    p_temp.post = Main.db_data(pos).post & TextBox1.Text
                    p_temp.skills = Main.db_data(pos).skills
                    p_temp.achievements = Main.db_data(pos).achievements
                ElseIf ComboBox1.SelectedIndex = 6 Then
                    p_temp.name = Main.db_data(pos).name
                    p_temp.year = Main.db_data(pos).year
                    p_temp.school = Main.db_data(pos).school
                    p_temp.class_ = Main.db_data(pos).class_
                    p_temp.telephone = Main.db_data(pos).telephone
                    p_temp.post = Main.db_data(pos).post
                    If Main.db_data(pos).skills.Length = 1 AndAlso Main.db_data(pos).skills(0) = "" Then
                        p_temp.skills = {TextBox1.Text}
                    Else
                        Dim temp_arr() As String = Main.db_data(pos).skills
                        Array.Resize(temp_arr, temp_arr.Length + 1)
                        temp_arr(temp_arr.Length - 1) = TextBox1.Text
                        p_temp.skills = temp_arr
                    End If
                    p_temp.achievements = Main.db_data(pos).achievements
                ElseIf ComboBox1.SelectedIndex = 7 Then
                    p_temp.name = Main.db_data(pos).name
                    p_temp.year = Main.db_data(pos).year
                    p_temp.school = Main.db_data(pos).school
                    p_temp.class_ = Main.db_data(pos).class_
                    p_temp.telephone = Main.db_data(pos).telephone
                    p_temp.post = Main.db_data(pos).post
                    p_temp.skills = Main.db_data(pos).skills
                    If Main.db_data(pos).achievements.Length = 1 AndAlso Main.db_data(pos).achievements(0) = "" Then
                        p_temp.achievements = {TextBox1.Text}
                    Else
                        Dim temp_arr() As String = Main.db_data(pos).achievements
                        Array.Resize(temp_arr, temp_arr.Length + 1)
                        temp_arr(temp_arr.Length - 1) = TextBox1.Text
                        p_temp.achievements = temp_arr
                    End If
                End If
            End If
            If Main.db_data(pos).photo <> "" Then
                p_temp.photo = p_temp.key
                p_temp.img = Main.db_data(pos).img
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
        Next
        If Main.CheckBox1.Checked Then
            Main.Button_save_db()
        Else
            Main.db_not_saved = True
            Main.Help.SetToolTip(Main.PictureBox2, "Изменения в базе данных не сохранены.")
            Main.PictureBox2.Image = My.Resources.not_saved
        End If
        MsgBox("Дублировано " & indexes.Length & " элемент(-а)(-ов).",, "Сообщение")
        Me.Close()
    End Sub
End Class
