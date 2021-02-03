Public Class Search_db
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'Name - TextBox
        TextBox1.Text = ""
        'Name - CheckBox
        CheckBox1.Checked = False
        'Name - have, self, full
        RadioButton1.Checked = True
        RadioButton2.Checked = False
        RadioButton15.Checked = False
        RadioButton15.Enabled = False
        'Name - and, or
        RadioButton24.Checked = True
        RadioButton23.Checked = False

        'Year - TextBox
        TextBox2.Text = ""
        'Year - CheckBox
        CheckBox2.Checked = False
        'Year - have, self, full
        RadioButton4.Checked = True
        RadioButton3.Checked = False
        RadioButton16.Checked = False
        RadioButton16.Enabled = False
        'Year - and, or
        RadioButton27.Checked = True
        RadioButton22.Checked = False

        'School - TextBox
        TextBox3.Text = ""
        'School - CheckBox
        CheckBox3.Checked = False
        'School - have, self, full
        RadioButton6.Checked = True
        RadioButton5.Checked = False
        RadioButton17.Checked = False
        RadioButton17.Enabled = False
        'School - and, or
        RadioButton29.Checked = True
        RadioButton28.Checked = False

        'Class - TextBox
        TextBox4.Text = ""
        'Class - CheckBox
        CheckBox4.Checked = False
        'Class - have, self, full
        RadioButton8.Checked = True
        RadioButton7.Checked = False
        RadioButton18.Checked = False
        RadioButton18.Enabled = False
        'Class - and, or
        RadioButton31.Checked = True
        RadioButton30.Checked = False

        'Telephone - TextBox
        TextBox5.Text = ""
        'Telephone - CheckBox
        CheckBox5.Checked = False
        'Telephone - have, self, full
        RadioButton10.Checked = True
        RadioButton9.Checked = False
        RadioButton19.Checked = False
        RadioButton19.Enabled = False
        'Telephone - and, or
        RadioButton33.Checked = True
        RadioButton32.Checked = False

        'Post - TextBox
        TextBox6.Text = ""
        'Post - CheckBox
        CheckBox6.Checked = False
        'Post - have, self, full
        RadioButton12.Checked = True
        RadioButton11.Checked = False
        RadioButton20.Checked = False
        RadioButton20.Enabled = False
        'Post - and, or
        RadioButton35.Checked = True
        RadioButton34.Checked = False

        'Skill - TextBox
        TextBox7.Text = ""
        'Skill - CheckBox
        CheckBox7.Checked = False
        'Skill - have, self, full
        RadioButton21.Checked = True
        RadioButton14.Checked = False
        RadioButton13.Checked = False
        RadioButton13.Enabled = False
        'Skill - and, or
        RadioButton37.Checked = True
        RadioButton36.Checked = False

        'Achievements - TextBox
        TextBox8.Text = ""
        'Achievements - CheckBox
        CheckBox8.Checked = False
        'Achievements - have, self, full
        RadioButton40.Checked = True
        RadioButton39.Checked = False
        RadioButton38.Checked = False
        RadioButton38.Enabled = False
        'Achievements - and, or
        RadioButton42.Checked = True
        RadioButton41.Checked = False

        'Output factor - and, or
        RadioButton25.Checked = True
        RadioButton26.Checked = False
    End Sub
    Private Sub RadioButton24_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton24.CheckedChanged
        If RadioButton24.Checked Then
            If RadioButton15.Checked Then
                RadioButton1.Checked = True
                RadioButton15.Checked = False
            End If
            RadioButton15.Enabled = False
        Else
            RadioButton15.Enabled = True
        End If
    End Sub
    Private Sub RadioButton23_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton23.CheckedChanged
        If RadioButton24.Checked Then
            If RadioButton15.Checked Then
                RadioButton1.Checked = True
                RadioButton15.Checked = False
            End If
            RadioButton15.Enabled = False
        Else
            RadioButton15.Enabled = True
        End If
    End Sub
    Private Sub RadioButton27_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton27.CheckedChanged
        If RadioButton27.Checked Then
            If RadioButton16.Checked Then
                RadioButton4.Checked = True
                RadioButton16.Checked = False
            End If
            RadioButton16.Enabled = False
        Else
            RadioButton16.Enabled = True
        End If
    End Sub
    Private Sub RadioButton22_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton22.CheckedChanged
        If RadioButton27.Checked Then
            If RadioButton16.Checked Then
                RadioButton4.Checked = True
                RadioButton16.Checked = False
            End If
            RadioButton16.Enabled = False
        Else
            RadioButton16.Enabled = True
        End If
    End Sub
    Private Sub RadioButton29_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton29.CheckedChanged
        If RadioButton29.Checked Then
            If RadioButton17.Checked Then
                RadioButton6.Checked = True
                RadioButton17.Checked = False
            End If
            RadioButton17.Enabled = False
        Else
            RadioButton17.Enabled = True
        End If
    End Sub
    Private Sub RadioButton28_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton28.CheckedChanged
        If RadioButton29.Checked Then
            If RadioButton17.Checked Then
                RadioButton6.Checked = True
                RadioButton17.Checked = False
            End If
            RadioButton17.Enabled = False
        Else
            RadioButton17.Enabled = True
        End If
    End Sub
    Private Sub RadioButton31_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton31.CheckedChanged
        If RadioButton31.Checked Then
            If RadioButton18.Checked Then
                RadioButton8.Checked = True
                RadioButton18.Checked = False
            End If
            RadioButton18.Enabled = False
        Else
            RadioButton18.Enabled = True
        End If
    End Sub
    Private Sub RadioButton30_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton30.CheckedChanged
        If RadioButton31.Checked Then
            If RadioButton18.Checked Then
                RadioButton8.Checked = True
                RadioButton18.Checked = False
            End If
            RadioButton18.Enabled = False
        Else
            RadioButton18.Enabled = True
        End If
    End Sub
    Private Sub RadioButton33_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton33.CheckedChanged
        If RadioButton33.Checked Then
            If RadioButton19.Checked Then
                RadioButton10.Checked = True
                RadioButton19.Checked = False
            End If
            RadioButton19.Enabled = False
        Else
            RadioButton19.Enabled = True
        End If
    End Sub
    Private Sub RadioButton32_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton32.CheckedChanged
        If RadioButton33.Checked Then
            If RadioButton19.Checked Then
                RadioButton10.Checked = True
                RadioButton19.Checked = False
            End If
            RadioButton19.Enabled = False
        Else
            RadioButton19.Enabled = True
        End If
    End Sub
    Private Sub RadioButton35_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton35.CheckedChanged
        If RadioButton35.Checked Then
            If RadioButton20.Checked Then
                RadioButton12.Checked = True
                RadioButton20.Checked = False
            End If
            RadioButton20.Enabled = False
        Else
            RadioButton20.Enabled = True
        End If
    End Sub
    Private Sub RadioButton34_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton34.CheckedChanged
        If RadioButton35.Checked Then
            If RadioButton20.Checked Then
                RadioButton12.Checked = True
                RadioButton20.Checked = False
            End If
            RadioButton20.Enabled = False
        Else
            RadioButton20.Enabled = True
        End If
    End Sub
    Private Sub RadioButton37_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton37.CheckedChanged
        If RadioButton37.Checked Then
            If RadioButton13.Checked Then
                RadioButton21.Checked = True
                RadioButton13.Checked = False
            End If
            RadioButton13.Enabled = False
        Else
            RadioButton13.Enabled = True
        End If
    End Sub
    Private Sub RadioButton36_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton36.CheckedChanged
        If RadioButton37.Checked Then
            If RadioButton13.Checked Then
                RadioButton21.Checked = True
                RadioButton13.Checked = False
            End If
            RadioButton13.Enabled = False
        Else
            RadioButton13.Enabled = True
        End If
    End Sub
    Private Sub RadioButton42_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton42.CheckedChanged
        If RadioButton42.Checked Then
            If RadioButton38.Checked Then
                RadioButton40.Checked = True
                RadioButton38.Checked = False
            End If
            RadioButton38.Enabled = False
        Else
            RadioButton38.Enabled = True
        End If
    End Sub
    Private Sub RadioButton41_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton41.CheckedChanged
        If RadioButton42.Checked Then
            If RadioButton38.Checked Then
                RadioButton40.Checked = True
                RadioButton38.Checked = False
            End If
            RadioButton38.Enabled = False
        Else
            RadioButton38.Enabled = True
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Bst_search(TextBox1.Text) And Bst_search(TextBox2.Text) And Bst_search(TextBox3.Text) And Bst_search(TextBox4.Text) And Bst_search(TextBox5.Text) And Bst_search(TextBox6.Text) And Bst_search(TextBox7.Text) And Bst_search(TextBox8.Text) Then
            For Each i As Control In Me.Controls
                If i.Tag = "ToDisable" Then
                    i.Enabled = False
                End If
            Next
            Main.Search_start()
        Else
            MsgBox("Поле(я) ввода содержит(ат) запрещённые символы!" & Chr(10) & "К запрещённым символам поиска относятся: <>=[]|&", MsgBoxStyle.OkOnly, "Сообщение")
        End If
    End Sub
    Private Function Bst_search(text As String) As Boolean
        If text.IndexOf("<") = -1 And text.IndexOf(">") = -1 And text.IndexOf("=") = -1 And text.IndexOf("[") = -1 And text.IndexOf("]") = -1 And text.IndexOf("|") = -1 And text.IndexOf("&") = -1 Then
            Return True
        Else
            Return False
        End If
    End Function
    Public Sub Search_Elements_Completed()
        If Main.sdb_keys.Count = 0 Then
            MsgBox("Поиск не дал результатов." & vbNewLine & "Найдено элементов: 0",, "Сообщение")
        Else
            MsgBox("Поиск завершён." & vbNewLine & "Найдено элементов: " & Main.sdb_keys.Count,, "Сообщение")
        End If
        For Each i As Control In Me.Controls
            If i.Tag = "ToDisable" Then
                i.Enabled = True
            End If
        Next
        If Main.search_mode = False Then
            Main.RadioButton2.Checked = True
            Main.RadioButton1.Checked = False
        Else
            Main.Renovate_DataGrid(1)
        End If
    End Sub
    Private Sub Search_db_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If Main.Search_Elements.IsBusy = True Then
            MsgBox("Идёт поиск..." & vbNewLine & "Пожалуйста, подождите...",, "Сообщение")
            e.Cancel = True
        End If
    End Sub
End Class
