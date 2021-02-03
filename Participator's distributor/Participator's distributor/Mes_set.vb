Public Class Mes_set
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
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
            Dim names() As String = {"StartMessange", "InactiveMessangeTime", "MessangeSound", "DistributingStartMessange", "DistributingEndMessange", "SaveMessange"}, reg_values() As String = reg.GetValueNames
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
                    Exit Sub
                End If
            Next
            If CheckBox2.Checked = True Then
                Main.settings(0) = "1"
            Else
                Main.settings(0) = "0"
            End If
            If CheckBox1.Checked = True Then
                Main.settings(1) = "1"
            Else
                Main.settings(1) = "0"
            End If
            If CheckBox3.Checked = True Then
                Main.settings(2) = "1"
            Else
                Main.settings(2) = "0"
            End If
            If CheckBox4.Checked = True Then
                Main.settings(4) = "1"
            Else
                Main.settings(4) = "0"
            End If
            If CheckBox5.Checked = True Then
                Main.settings(5) = "1"
            Else
                Main.settings(5) = "0"
            End If
            If CheckBox6.Checked = True Then
                Main.settings(6) = "1"
            Else
                Main.settings(6) = "0"
            End If
            reg.SetValue("StartMessange", Main.settings(0), Microsoft.Win32.RegistryValueKind.DWord)
            reg.SetValue("InactiveMessangeTime", Main.settings(1), Microsoft.Win32.RegistryValueKind.DWord)
            reg.SetValue("MessangeSound", Main.settings(2), Microsoft.Win32.RegistryValueKind.DWord)
            reg.SetValue("DistributingStartMessange", Main.settings(4), Microsoft.Win32.RegistryValueKind.DWord)
            reg.SetValue("DistributingEndMessange", Main.settings(5), Microsoft.Win32.RegistryValueKind.DWord)
            reg.SetValue("SaveMessange", Main.settings(6), Microsoft.Win32.RegistryValueKind.DWord)
            reg.Close()
            Me.Close()
        Else
            reg.Close()
            MsgBox("Нет доступа к настройкам программы!",, "Ошибка")
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class
