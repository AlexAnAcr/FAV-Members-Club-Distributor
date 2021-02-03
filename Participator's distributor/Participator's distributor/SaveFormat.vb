Imports System.Windows.Forms
Public Class SaveFormat
    Public work_mode As Short
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim result As DialogResult = SaveFile.ShowDialog
        If result = DialogResult.OK Then
            Me.DialogResult = DialogResult.OK
            Me.Close()
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
    Private Sub SaveFormat_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        If work_mode = 0 Then
            CheckBox1.Text = "Сохранять распределительный номер команд"
            CheckBox2.Text = "Сохранять распределительный номер участников"
            CheckBox3.Text = "Сохранять кол-во участников в командах"
        Else
            CheckBox1.Text = "Сохранять распределительный номер турниров"
            CheckBox2.Text = "Сохранять распределительный номер команд"
            CheckBox3.Text = "Сохранять кол-во команд в турнире"
        End If
    End Sub
End Class
