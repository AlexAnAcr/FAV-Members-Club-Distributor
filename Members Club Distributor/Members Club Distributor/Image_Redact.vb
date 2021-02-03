Public Class Image_Redact
    Dim photo_exist As String, image As Image, image_way As String = ""
    Private Sub Image_Redact_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Add_Redact_db_data.p_temp.photo = "" Then
            RadioButton1.Enabled = False
            RadioButton2.Enabled = True
            RadioButton1.Checked = False
            RadioButton2.Checked = True
            TextBox5.Enabled = True
            Button1.Enabled = True
            Button5.Enabled = False
        Else
            RadioButton1.Enabled = True
            RadioButton2.Enabled = True
            RadioButton2.Checked = False
            RadioButton1.Checked = True
            TextBox5.Enabled = False
            Button1.Enabled = False
            Button5.Enabled = True
            photo_exist = Add_Redact_db_data.p_temp.photo
            image = Add_Redact_db_data.p_temp.img
        End If
        TextBox5.Text = ""
        Image_Cut1.CutImagePrewiev.Image = Nothing
        PictureBox1.Image = Nothing
        TextBox1.Enabled = False
        TextBox1.Text = "0"
        TextBox2.Enabled = False
        TextBox2.Text = "0"
        TextBox3.Enabled = False
        TextBox3.Text = "0"
        TextBox4.Enabled = False
        TextBox4.Text = "0"
        Image_Cut1.Enabled = False
        Label3.Visible = False
    End Sub
    Private Sub Image_Redact_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Image_Cut1.outside_rect.X = 0
        Image_Cut1.outside_rect.Y = 0
        Image_Cut1.outside_rect.Width = 400
        Image_Cut1.outside_rect.Height = 400
        Image_Cut1.rect.X = 50
        Image_Cut1.rect.Y = 50
        Image_Cut1.rect.Width = 290
        Image_Cut1.rect.Height = 290
        Image_Cut1.Draw_Markers()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim result As DialogResult = OpenFile1.ShowDialog
        If result = DialogResult.OK Then
            TextBox5.Text = OpenFile1.FileName
        End If
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        photo_exist = ""
        RadioButton1.Enabled = False
        RadioButton2.Enabled = True
        RadioButton1.Checked = False
        RadioButton2.Checked = True
        TextBox5.Enabled = True
        PictureBox1.Image = Nothing
        Button1.Enabled = True
        Button5.Enabled = False
        TextBox1.Enabled = False
        TextBox1.Text = "0"
        TextBox2.Enabled = False
        TextBox2.Text = "0"
        TextBox3.Enabled = False
        TextBox3.Text = "0"
        TextBox4.Enabled = False
        TextBox4.Text = "0"
        Image_Cut1.Enabled = False
        Label3.Visible = False

        Image_Cut1.outside_rect.X = 0
        Image_Cut1.outside_rect.Y = 0
        Image_Cut1.outside_rect.Width = 400
        Image_Cut1.outside_rect.Height = 400
        Image_Cut1.rect.X = 50
        Image_Cut1.rect.Y = 50
        Image_Cut1.rect.Width = 290
        Image_Cut1.rect.Height = 290
        Image_Cut1.CutImagePrewiev.Image = Nothing
        Image_Cut1.Draw_Markers()
        Image_Cut1.Enabled = False
        r_start = False
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If RadioButton1.Checked = True Then
            TextBox5.Text = ""
            sourse_img = image
            Initial_ImageRedact()
        Else
            If My.Computer.FileSystem.FileExists(TextBox5.Text) Then
                If (My.Computer.FileSystem.GetFileInfo(TextBox5.Text).Extension.ToLower = ".jpg" Or My.Computer.FileSystem.GetFileInfo(TextBox5.Text).Extension.ToLower = ".jpeg") Then
                    image_way = TextBox5.Text
                    Dim img_stream As IO.Stream = IO.File.OpenRead(image_way)
                    sourse_img = New Bitmap(Image.FromStream(img_stream))
                    img_stream.Close()
                    Initial_ImageRedact()
                Else
                    MsgBox("Выбранные файлы не имеют расширение '.jpg' или '.jpeg'!",, "Ошибка")
                End If
            Else
                MsgBox("Выбранный файл не существует!",, "Ошибка")
            End If
        End If
    End Sub
    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton2.Checked Then
            TextBox5.Enabled = True
            Button1.Enabled = True
        Else
            TextBox5.Enabled = False
            Button1.Enabled = False
            image = Add_Redact_db_data.p_temp.img
        End If
        Image_Cut1.outside_rect.X = 0
        Image_Cut1.outside_rect.Y = 0
        Image_Cut1.outside_rect.Width = 400
        Image_Cut1.outside_rect.Height = 400
        Image_Cut1.rect.X = 50
        Image_Cut1.rect.Y = 50
        Image_Cut1.rect.Width = 290
        Image_Cut1.rect.Height = 290
        Image_Cut1.CutImagePrewiev.Image = Nothing
        PictureBox1.Image = Nothing
        Image_Cut1.Draw_Markers()
        TextBox1.Enabled = False
        TextBox1.Text = "0"
        TextBox2.Enabled = False
        TextBox2.Text = "0"
        TextBox3.Enabled = False
        TextBox3.Text = "0"
        TextBox4.Enabled = False
        TextBox4.Text = "0"
        Image_Cut1.Enabled = False
        Label3.Visible = False
    End Sub
    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked Then
            TextBox5.Enabled = True
            Button1.Enabled = True
        Else
            TextBox5.Enabled = False
            Button1.Enabled = False
            image = Add_Redact_db_data.p_temp.img
        End If
        Image_Cut1.outside_rect.X = 0
        Image_Cut1.outside_rect.Y = 0
        Image_Cut1.outside_rect.Width = 400
        Image_Cut1.outside_rect.Height = 400
        Image_Cut1.rect.X = 50
        Image_Cut1.rect.Y = 50
        Image_Cut1.rect.Width = 290
        Image_Cut1.rect.Height = 290
        Image_Cut1.CutImagePrewiev.Image = Nothing
        PictureBox1.Image = Nothing
        Image_Cut1.Draw_Markers()
        TextBox1.Enabled = False
        TextBox1.Text = "0"
        TextBox2.Enabled = False
        TextBox2.Text = "0"
        TextBox3.Enabled = False
        TextBox3.Text = "0"
        TextBox4.Enabled = False
        TextBox4.Text = "0"
        Image_Cut1.Enabled = False
        Label3.Visible = False
    End Sub
    Dim sourse_img As Image
    Private Sub Initial_ImageRedact()
        If sourse_img.Width <= 64 Or sourse_img.Height <= 64 Then
            Image_Cut1.Enabled = False
            TextBox1.Text = 0
            TextBox2.Text = 0
            TextBox3.Text = sourse_img.Width
            TextBox4.Text = sourse_img.Height
            image = sourse_img
            If Add_Redact_db_data.p_temp.key = "" Then
                photo_exist = "1"
            Else
                photo_exist = Add_Redact_db_data.p_temp.key
            End If
            PictureBox1.Image = sourse_img
            Label3.Text = "Изображение имеет очень малые размеры."
            Label3.Visible = True
        Else
            TextBox1.Enabled = True
            TextBox2.Enabled = True
            TextBox3.Enabled = True
            TextBox4.Enabled = True
            image = sourse_img
            If Add_Redact_db_data.p_temp.key = "" Then
                photo_exist = "1"
            Else
                photo_exist = Add_Redact_db_data.p_temp.key
            End If
            PictureBox1.Image = sourse_img
            If sourse_img.Width > 1000 Or sourse_img.Height > 1000 Then
                Label3.Text = "Будет применено масштабирование."
                Label3.Visible = True
            Else
                Label3.Visible = False
            End If
            Image_Cut1.Enabled = True
            Initial_CutRectanglle()
        End If
    End Sub
    Dim prew_img As Bitmap, size_image(3) As Integer, r_start As Boolean = False, dark_brush As Drawing.SolidBrush, proportion(1) As Single
    Private Sub Initial_CutRectanglle()
        dark_brush = New Drawing.SolidBrush(Color.FromArgb(125, Color.Black))
        Dim bmp As New Bitmap(sourse_img), bmp1 As New Bitmap(Image_Cut1.CutImagePrewiev.Width, Image_Cut1.CutImagePrewiev.Height, Imaging.PixelFormat.Format24bppRgb)

        If bmp.Width > bmp1.Width Or bmp.Height > bmp1.Height Then
            Dim scale_size As Single = bmp.Width / bmp.Height
            If scale_size >= 1 Then
                size_image(0) = bmp1.Width
                size_image(1) = Math.Round(bmp1.Height / scale_size)
            Else
                size_image(0) = Math.Round(bmp1.Width * scale_size)
                size_image(1) = bmp1.Height
            End If
        Else
            size_image(0) = bmp.Width
            size_image(1) = bmp.Height
        End If
        prew_img = New Bitmap(size_image(0), size_image(1), Imaging.PixelFormat.Format24bppRgb)
        Dim r As New Rectangle(0, 0, size_image(0), size_image(1)), gr1 As Graphics = Graphics.FromImage(prew_img)
        gr1.CompositingQuality = Drawing2D.CompositingQuality.HighSpeed
        gr1.CompositingMode = Drawing2D.CompositingMode.SourceOver
        gr1.DrawImage(bmp, r)

        proportion(0) = sourse_img.Width / prew_img.Width
        proportion(1) = sourse_img.Height / prew_img.Height

        size_image(2) = (bmp1.Width - size_image(0)) / 2
        size_image(3) = (bmp1.Height - size_image(1)) / 2
        'Set cut bounds
        Image_Cut1.outside_rect.X = size_image(2)
        Image_Cut1.outside_rect.Y = size_image(3)
        Image_Cut1.outside_rect.Width = size_image(0)
        Image_Cut1.outside_rect.Height = size_image(1)
        Image_Cut1.rect.X = size_image(2)
        Image_Cut1.rect.Y = size_image(3)
        Image_Cut1.rect.Width = size_image(0)
        Image_Cut1.rect.Height = size_image(1)
        Image_Cut1.Draw_Markers()
        TextBox1.Text = Math.Round((Image_Cut1.rect.X - size_image(2)) * proportion(0))
        TextBox2.Text = Math.Round((Image_Cut1.rect.Y - size_image(3)) * proportion(1))
        TextBox3.Text = Math.Round(Image_Cut1.rect.Width * proportion(0))
        TextBox4.Text = Math.Round(Image_Cut1.rect.Height * proportion(1))
        'Set cut bounds end
        Dim r1 As New Rectangle(size_image(2), size_image(3), size_image(0), size_image(1)), gr As Graphics = Graphics.FromImage(bmp1)
        gr.CompositingQuality = Drawing2D.CompositingQuality.HighSpeed
        gr.CompositingMode = Drawing2D.CompositingMode.SourceOver
        gr.DrawImage(prew_img, r1)

        Image_Cut1.CutImagePrewiev.Image = bmp1
        r_start = True
    End Sub
    Private Sub Image_Cut1_CutRectangleChangeEnd() Handles Image_Cut1.CutRectangleChangeEnd
        If r_start = True Then
            Dim bmp As New Bitmap(Image_Cut1.CutImagePrewiev.Width, Image_Cut1.CutImagePrewiev.Height, Imaging.PixelFormat.Format24bppRgb)
            Dim gr As Graphics = Graphics.FromImage(bmp)
            gr.CompositingQuality = Drawing2D.CompositingQuality.HighSpeed
            gr.CompositingMode = Drawing2D.CompositingMode.SourceOver
            Dim r As New Rectangle(size_image(2), size_image(3), prew_img.Width, prew_img.Height), r1 As New Rectangle(Image_Cut1.rect.X - size_image(2), Image_Cut1.rect.Y - size_image(3), Image_Cut1.rect.Width, Image_Cut1.rect.Height)
            gr.DrawImage(prew_img, r)
            r = New Rectangle(size_image(2) + Image_Cut1.rect.X - size_image(2), size_image(3) + Image_Cut1.rect.Y - size_image(3), Image_Cut1.rect.Width, Image_Cut1.rect.Height)
            gr.FillRectangle(dark_brush, size_image(2), size_image(3), bmp.Width - size_image(2), bmp.Height - size_image(3))
            gr.DrawImage(prew_img, r, r1, GraphicsUnit.Pixel)
            Image_Cut1.CutImagePrewiev.Image = bmp
            Renov_dest_img()
        End If
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If photo_exist <> "" Then
            If image.Width > 1000 Or image.Height > 1000 Then
                Dim scale_size As Single = image.Width / image.Height
                Dim new_size(1) As UShort
                If scale_size > 1 Then
                    new_size(0) = 1000
                    new_size(1) = Math.Round(1000 / scale_size)
                Else
                    new_size(0) = Math.Round(1000 * scale_size)
                    new_size(1) = 1000
                End If
                Dim bmp_exit As New Bitmap(new_size(0), new_size(1), Imaging.PixelFormat.Format24bppRgb), r As New Rectangle(0, 0, new_size(0), new_size(1))
                Dim gr As Graphics = Graphics.FromImage(bmp_exit)
                gr.DrawImage(image, r)
                Add_Redact_db_data.p_temp.img = bmp_exit
            Else
                Add_Redact_db_data.p_temp.img = image
            End If
            Add_Redact_db_data.p_temp.photo = photo_exist
        Else
            Add_Redact_db_data.p_temp.photo = ""
        End If
        Me.Close()
    End Sub
    Private Sub Renov_dest_img()
        Dim bmp1 As New Bitmap(Math.Round(Image_Cut1.rect.Width * proportion(0)), Math.Round(Image_Cut1.rect.Height * proportion(1)), Imaging.PixelFormat.Format24bppRgb), gr As Graphics = Graphics.FromImage(bmp1)
        Dim r As New Rectangle(0, 0, Math.Round(Image_Cut1.rect.Width * proportion(0)), Math.Round(Image_Cut1.rect.Height * proportion(1))), r1 As New Rectangle(Math.Round((Image_Cut1.rect.X - size_image(2)) * proportion(0)), Math.Round((Image_Cut1.rect.Y - size_image(3)) * proportion(1)), Math.Round(Image_Cut1.rect.Width * proportion(0)), Math.Round(Image_Cut1.rect.Height * proportion(1)))
        gr.CompositingQuality = Drawing2D.CompositingQuality.Default
        gr.CompositingMode = Drawing2D.CompositingMode.SourceOver
        gr.DrawImage(sourse_img, r, r1, GraphicsUnit.Pixel)
        PictureBox1.Image = bmp1
        image = bmp1
        If image.Width > 1000 Or image.Height > 1000 Then
            Label3.Text = "Будет применено масштабирование."
            Label3.Visible = True
        Else
            If image.Width <= 64 Or image.Height <= 64 Then
                Label3.Text = "Изображение имеет очень малые размеры."
                Label3.Visible = True
            Else
                Label3.Visible = False
            End If
        End If
    End Sub
    Private Sub Image_Cut1_CutRectangleChanged() Handles Image_Cut1.CutRectangleChanged
        TextBox1.Text = Math.Round((Image_Cut1.rect.X - size_image(2)) * proportion(0))
        TextBox2.Text = Math.Round((Image_Cut1.rect.Y - size_image(3)) * proportion(1))
        TextBox3.Text = Math.Round(Image_Cut1.rect.Width * proportion(0))
        TextBox4.Text = Math.Round(Image_Cut1.rect.Height * proportion(1))
    End Sub
    Private Sub TextBox1_LostFocus(sender As Object, e As EventArgs) Handles TextBox1.LostFocus
        Try
            Dim temp As UShort = TextBox1.Text
            temp = Math.Round(temp / proportion(0))
            If temp < 0 Then
                Image_Cut1.rect.X = Image_Cut1.outside_rect.X
                TextBox1.Text = 0
            ElseIf temp + Image_Cut1.rect.Width > Image_Cut1.outside_rect.Width Then
                Image_Cut1.rect.X = Image_Cut1.outside_rect.X + Image_Cut1.outside_rect.Width - Image_Cut1.rect.Width
                TextBox1.Text = Math.Round((Image_Cut1.rect.X - Image_Cut1.outside_rect.X) * proportion(0))
            Else
                Image_Cut1.rect.X = Image_Cut1.outside_rect.X + temp
                TextBox1.Text = Math.Round(temp * proportion(0))
            End If
            Image_Cut1.Draw_Markers()
            If r_start = True Then
                Dim bmp As New Bitmap(Image_Cut1.CutImagePrewiev.Width, Image_Cut1.CutImagePrewiev.Height, Imaging.PixelFormat.Format24bppRgb)
                Dim gr As Graphics = Graphics.FromImage(bmp)
                gr.CompositingQuality = Drawing2D.CompositingQuality.HighSpeed
                gr.CompositingMode = Drawing2D.CompositingMode.SourceOver
                Dim r As New Rectangle(size_image(2), size_image(3), prew_img.Width, prew_img.Height), r1 As New Rectangle(Image_Cut1.rect.X - size_image(2), Image_Cut1.rect.Y - size_image(3), Image_Cut1.rect.Width, Image_Cut1.rect.Height)
                gr.DrawImage(prew_img, r)
                r = New Rectangle(size_image(2) + Image_Cut1.rect.X - size_image(2), size_image(3) + Image_Cut1.rect.Y - size_image(3), Image_Cut1.rect.Width, Image_Cut1.rect.Height)
                gr.FillRectangle(dark_brush, size_image(2), size_image(3), bmp.Width - size_image(2), bmp.Height - size_image(3))
                gr.DrawImage(prew_img, r, r1, GraphicsUnit.Pixel)
                Image_Cut1.CutImagePrewiev.Image = bmp
                Renov_dest_img()
            End If
        Catch
            MsgBox("Неверно заданы координаты!",, "Сообщение")
        End Try
    End Sub
    Private Sub TextBox2_LostFocus(sender As Object, e As EventArgs) Handles TextBox2.LostFocus
        Try
            Dim temp As UShort = TextBox2.Text
            temp = Math.Round(temp / proportion(1))
            If temp < 0 Then
                Image_Cut1.rect.Y = Image_Cut1.outside_rect.Y
                TextBox2.Text = 0
            ElseIf temp + Image_Cut1.rect.Height > Image_Cut1.outside_rect.Height Then
                Image_Cut1.rect.Y = Image_Cut1.outside_rect.Y + Image_Cut1.outside_rect.Height - Image_Cut1.rect.Height
                TextBox2.Text = Math.Round((Image_Cut1.rect.Y - Image_Cut1.outside_rect.Y) * proportion(1))
            Else
                Image_Cut1.rect.Y = Image_Cut1.outside_rect.Y + temp
                TextBox2.Text = Math.Round(temp * proportion(1))
            End If
            Image_Cut1.Draw_Markers()
            If r_start = True Then
                Dim bmp As New Bitmap(Image_Cut1.CutImagePrewiev.Width, Image_Cut1.CutImagePrewiev.Height, Imaging.PixelFormat.Format24bppRgb)
                Dim gr As Graphics = Graphics.FromImage(bmp)
                gr.CompositingQuality = Drawing2D.CompositingQuality.HighSpeed
                gr.CompositingMode = Drawing2D.CompositingMode.SourceOver
                Dim r As New Rectangle(size_image(2), size_image(3), prew_img.Width, prew_img.Height), r1 As New Rectangle(Image_Cut1.rect.X - size_image(2), Image_Cut1.rect.Y - size_image(3), Image_Cut1.rect.Width, Image_Cut1.rect.Height)
                gr.DrawImage(prew_img, r)
                r = New Rectangle(size_image(2) + Image_Cut1.rect.X - size_image(2), size_image(3) + Image_Cut1.rect.Y - size_image(3), Image_Cut1.rect.Width, Image_Cut1.rect.Height)
                gr.FillRectangle(dark_brush, size_image(2), size_image(3), bmp.Width - size_image(2), bmp.Height - size_image(3))
                gr.DrawImage(prew_img, r, r1, GraphicsUnit.Pixel)
                Image_Cut1.CutImagePrewiev.Image = bmp
                Renov_dest_img()
            End If
        Catch
            MsgBox("Неверно заданы координаты!",, "Сообщение")
        End Try
    End Sub
    Private Sub TextBox3_LostFocus(sender As Object, e As EventArgs) Handles TextBox3.LostFocus
        Try
            Dim temp As UShort = TextBox3.Text
            temp = Math.Round(temp / proportion(0))
            If temp <= 64 Then
                Image_Cut1.rect.Width = 64
                TextBox3.Text = Math.Round(64 * proportion(0))
            ElseIf temp >= Image_Cut1.outside_rect.Width Then
                Image_Cut1.rect.Width = Image_Cut1.outside_rect.Width
                TextBox3.Text = Math.Round(Image_Cut1.outside_rect.Width * proportion(0))
            Else
                Image_Cut1.rect.Width = temp
                TextBox3.Text = Math.Round(temp * proportion(1))
            End If
            Image_Cut1.Draw_Markers()
            If r_start = True Then
                Dim bmp As New Bitmap(Image_Cut1.CutImagePrewiev.Width, Image_Cut1.CutImagePrewiev.Height, Imaging.PixelFormat.Format24bppRgb)
                Dim gr As Graphics = Graphics.FromImage(bmp)
                gr.CompositingQuality = Drawing2D.CompositingQuality.HighSpeed
                gr.CompositingMode = Drawing2D.CompositingMode.SourceOver
                Dim r As New Rectangle(size_image(2), size_image(3), prew_img.Width, prew_img.Height), r1 As New Rectangle(Image_Cut1.rect.X - size_image(2), Image_Cut1.rect.Y - size_image(3), Image_Cut1.rect.Width, Image_Cut1.rect.Height)
                gr.DrawImage(prew_img, r)
                r = New Rectangle(size_image(2) + Image_Cut1.rect.X - size_image(2), size_image(3) + Image_Cut1.rect.Y - size_image(3), Image_Cut1.rect.Width, Image_Cut1.rect.Height)
                gr.FillRectangle(dark_brush, size_image(2), size_image(3), bmp.Width - size_image(2), bmp.Height - size_image(3))
                gr.DrawImage(prew_img, r, r1, GraphicsUnit.Pixel)
                Image_Cut1.CutImagePrewiev.Image = bmp
                Renov_dest_img()
            End If
        Catch
            MsgBox("Неверно заданы координаты!",, "Сообщение")
        End Try
    End Sub
    Private Sub TextBox4_LostFocus(sender As Object, e As EventArgs) Handles TextBox4.LostFocus
        Try
            Dim temp As UShort = TextBox4.Text
            temp = Math.Round(temp / proportion(1))
            If temp <= 64 Then
                Image_Cut1.rect.Height = 64
                TextBox4.Text = Math.Round(64 * proportion(1))
            ElseIf temp >= Image_Cut1.outside_rect.Height Then
                Image_Cut1.rect.Height = Image_Cut1.outside_rect.Height
                TextBox4.Text = Math.Round(Image_Cut1.outside_rect.Height * proportion(1))
            Else
                Image_Cut1.rect.Height = temp
                TextBox4.Text = Math.Round(temp * proportion(1))
            End If
            Image_Cut1.Draw_Markers()
            If r_start = True Then
                Dim bmp As New Bitmap(Image_Cut1.CutImagePrewiev.Width, Image_Cut1.CutImagePrewiev.Height, Imaging.PixelFormat.Format24bppRgb)
                Dim gr As Graphics = Graphics.FromImage(bmp)
                gr.CompositingQuality = Drawing2D.CompositingQuality.HighSpeed
                gr.CompositingMode = Drawing2D.CompositingMode.SourceOver
                Dim r As New Rectangle(size_image(2), size_image(3), prew_img.Width, prew_img.Height), r1 As New Rectangle(Image_Cut1.rect.X - size_image(2), Image_Cut1.rect.Y - size_image(3), Image_Cut1.rect.Width, Image_Cut1.rect.Height)
                gr.DrawImage(prew_img, r)
                r = New Rectangle(size_image(2) + Image_Cut1.rect.X - size_image(2), size_image(3) + Image_Cut1.rect.Y - size_image(3), Image_Cut1.rect.Width, Image_Cut1.rect.Height)
                gr.FillRectangle(dark_brush, size_image(2), size_image(3), bmp.Width - size_image(2), bmp.Height - size_image(3))
                gr.DrawImage(prew_img, r, r1, GraphicsUnit.Pixel)
                Image_Cut1.CutImagePrewiev.Image = bmp
                Renov_dest_img()
            End If
        Catch
            MsgBox("Неверно заданы координаты!",, "Сообщение")
        End Try
    End Sub
End Class
