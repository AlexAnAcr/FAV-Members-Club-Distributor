Public Class Image_Cut
    Dim new_pos_mode As Boolean = False, dcl As Point
    Public rect As New Rect_temp, outside_rect As New Rect_temp
    Public Event CutRectangleChanged()
    Public Event CutRectangleChangeEnd()
    Private Sub Image_Cut_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        rect.X = 10
        rect.Y = 10
        rect.Width = 100
        rect.Height = 100
        outside_rect.X = 0
        outside_rect.Y = 0
        outside_rect.Width = 200
        outside_rect.Height = 200
        Draw_Markers()
    End Sub
    Public Sub Draw_Markers()
        Cut_UpBarier.Left = rect.X - 2
        Cut_UpBarier.Top = rect.Y - 2
        Cut_UpBarier.Width = rect.Width + 4

        Cut_LeftBarier.Left = rect.X - 2
        Cut_LeftBarier.Top = rect.Y - 2
        Cut_LeftBarier.Height = rect.Height + 4

        Cut_DownBarier.Left = rect.X - 2
        Cut_DownBarier.Top = rect.Y + rect.Height - 2
        Cut_DownBarier.Width = rect.Width + 4

        Cut_RightBarier.Left = rect.X + rect.Width - 2
        Cut_RightBarier.Top = rect.Y - 2
        Cut_RightBarier.Height = rect.Height + 4

        Tap_Up.Top = rect.Y - 4
        Tap_Up.Left = rect.X + rect.Width / 2 - 7

        Tap_Left.Left = rect.X - 4
        Tap_Left.Top = rect.Y + rect.Height / 2 - 7

        Tap_Down.Left = rect.X + rect.Width / 2 - 7
        Tap_Down.Top = rect.Y + rect.Height - 4

        Tap_Right.Left = rect.X + rect.Width - 4
        Tap_Right.Top = rect.Y + rect.Height / 2 - 7

        Tap_LeftUp.Left = rect.X - 7
        Tap_LeftUp.Top = rect.Y - 7

        Tap_RightUp.Left = rect.X + rect.Width - 7
        Tap_RightUp.Top = rect.Y - 7

        Tap_LeftDown.Left = rect.X - 7
        Tap_LeftDown.Top = rect.Y + rect.Height - 7

        Tap_RightDown.Left = rect.X + rect.Width - 7
        Tap_RightDown.Top = rect.Y + rect.Height - 7
    End Sub
    Private Sub Tap_Up_MouseDown(sender As Object, e As MouseEventArgs) Handles Tap_Up.MouseDown
        dcl.Y = e.Y
        new_pos_mode = True
    End Sub
    Private Sub Tap_Up_MouseMove(sender As Object, e As MouseEventArgs) Handles Tap_Up.MouseMove
        If new_pos_mode = True Then
            If Tap_Up.Top > Tap_Up.Top + e.Y - dcl.Y Then
                If Tap_Up.Top + e.Y - dcl.Y > outside_rect.Y - 4 Then
                    Dim pos As Integer = Tap_Up.Top + e.Y - dcl.Y
                    rect.Height = rect.Height + (rect.Y - (pos + 4))
                    rect.Y = pos + 4
                Else
                    rect.Height = rect.Height + (rect.Y - outside_rect.Y)
                    rect.Y = outside_rect.Y
                End If
            Else
                If rect.Height + (rect.Y - (Tap_Up.Top + 4)) > 64 Then
                    Dim pos As Integer = Tap_Up.Top + e.Y - dcl.Y
                    rect.Height = rect.Height + (rect.Y - (pos + 4))
                    rect.Y = pos + 4
                Else
                    rect.Y = rect.Y + (rect.Height - 64)
                    rect.Height = 64
                End If
            End If
            Draw_Markers()
            RaiseEvent CutRectangleChanged()
        End If
    End Sub
    Private Sub Tap_Up_MouseUp(sender As Object, e As MouseEventArgs) Handles Tap_Up.MouseUp
        new_pos_mode = False
        RaiseEvent CutRectangleChangeEnd()
    End Sub
    Private Sub Tap_Left_MouseDown(sender As Object, e As MouseEventArgs) Handles Tap_Left.MouseDown
        dcl.X = e.X
        new_pos_mode = True
    End Sub
    Private Sub Tap_Left_MouseUp(sender As Object, e As MouseEventArgs) Handles Tap_Left.MouseUp
        new_pos_mode = False
        RaiseEvent CutRectangleChangeEnd()
    End Sub
    Private Sub Tap_Left_MouseMove(sender As Object, e As MouseEventArgs) Handles Tap_Left.MouseMove
        If new_pos_mode = True Then
            If Tap_Left.Left > Tap_Left.Left + e.X - dcl.X Then
                If Tap_Left.Left + e.X - dcl.X > outside_rect.X - 4 Then
                    Dim pos As Integer = Tap_Left.Left + e.X - dcl.X
                    rect.Width = rect.Width + (rect.X - (pos + 4))
                    rect.X = pos + 4
                Else
                    rect.Width = rect.Width + (rect.X - outside_rect.X)
                    rect.X = outside_rect.X
                End If
            Else
                If rect.Width + (rect.X - (Tap_Left.Left + 4)) > 64 Then
                    Dim pos As Integer = Tap_Left.Left + e.X - dcl.X
                    rect.Width = rect.Width + (rect.X - (pos + 4))
                    rect.X = pos + 4
                Else
                    rect.X = rect.X + (rect.Width - 64)
                    rect.Width = 64
                End If
            End If
            Draw_Markers()
            RaiseEvent CutRectangleChanged()
        End If
    End Sub
    Private Sub Tap_Right_MouseDown(sender As Object, e As MouseEventArgs) Handles Tap_Right.MouseDown
        dcl.X = e.X
        new_pos_mode = True
    End Sub
    Private Sub Tap_Right_MouseMove(sender As Object, e As MouseEventArgs) Handles Tap_Right.MouseMove
        If new_pos_mode = True Then
            If Tap_Right.Left < Tap_Right.Left + e.X - dcl.X Then
                If Tap_Right.Left + e.X - dcl.X < outside_rect.X + outside_rect.Width - 4 Then
                    Dim pos As Integer = Tap_Right.Left + e.X - dcl.X
                    rect.Width = pos - rect.X + 4
                Else
                    rect.Width = outside_rect.X + outside_rect.Width - rect.X
                End If
            Else
                If Tap_Right.Left - rect.X + 4 > 64 Then
                    Dim pos As Integer = Tap_Right.Left + e.X - dcl.X
                    rect.Width = pos - rect.X + 4
                Else
                    rect.Width = 64
                End If
            End If
            Draw_Markers()
            RaiseEvent CutRectangleChanged()
        End If
    End Sub
    Private Sub Tap_Right_MouseUp(sender As Object, e As MouseEventArgs) Handles Tap_Right.MouseUp
        new_pos_mode = False
        RaiseEvent CutRectangleChangeEnd()
    End Sub
    Private Sub Tap_Down_MouseDown(sender As Object, e As MouseEventArgs) Handles Tap_Down.MouseDown
        dcl.Y = e.Y
        new_pos_mode = True
    End Sub
    Private Sub Tap_Down_MouseMove(sender As Object, e As MouseEventArgs) Handles Tap_Down.MouseMove
        If new_pos_mode = True Then
            If Tap_Down.Top < Tap_Down.Top + e.Y - dcl.Y Then
                If Tap_Down.Top + e.Y - dcl.Y < outside_rect.Y + outside_rect.Height - 4 Then
                    Dim pos As Integer = Tap_Down.Top + e.Y - dcl.Y
                    rect.Height = pos - rect.Y + 4
                Else
                    rect.Height = outside_rect.Y + outside_rect.Height - rect.Y
                End If
            Else
                If Tap_Down.Top - rect.Y + 4 > 64 Then
                    Dim pos As Integer = Tap_Down.Top + e.Y - dcl.Y
                    rect.Height = pos - rect.Y + 4
                Else
                    rect.Height = 64
                End If
            End If
            Draw_Markers()
            RaiseEvent CutRectangleChanged()
        End If
    End Sub
    Private Sub Tap_Down_MouseUp(sender As Object, e As MouseEventArgs) Handles Tap_Down.MouseUp
        new_pos_mode = False
        RaiseEvent CutRectangleChangeEnd()
    End Sub
    Private Sub Tap_LeftUp_MouseDown(sender As Object, e As MouseEventArgs) Handles Tap_LeftUp.MouseDown
        dcl.X = e.X
        dcl.Y = e.Y
        new_pos_mode = True
    End Sub
    Private Sub Tap_LeftUp_MouseMove(sender As Object, e As MouseEventArgs) Handles Tap_LeftUp.MouseMove
        If new_pos_mode = True Then
            If Tap_LeftUp.Left > Tap_LeftUp.Left + e.X - dcl.X Then
                If Tap_LeftUp.Left + e.X - dcl.X > outside_rect.X - 7 Then
                    Dim pos As Integer = Tap_LeftUp.Left + e.X - dcl.X
                    rect.Width = rect.Width + (rect.X - (pos + 7))
                    rect.X = pos + 7
                Else
                    rect.Width = rect.Width + (rect.X - outside_rect.X)
                    rect.X = outside_rect.X
                End If
            Else
                If rect.Width + (rect.X - (Tap_LeftUp.Left + 7)) > 64 Then
                    Dim pos As Integer = Tap_LeftUp.Left + e.X - dcl.X
                    rect.Width = rect.Width + (rect.X - (pos + 7))
                    rect.X = pos + 7
                Else
                    rect.X = rect.X + (rect.Width - 64)
                    rect.Width = 64
                End If
            End If
            If Tap_LeftUp.Top > Tap_LeftUp.Top + e.Y - dcl.Y Then
                If Tap_LeftUp.Top + e.Y - dcl.Y > outside_rect.Y - 7 Then
                    Dim pos As Integer = Tap_LeftUp.Top + e.Y - dcl.Y
                    rect.Height = rect.Height + (rect.Y - (pos + 7))
                    rect.Y = pos + 7
                Else
                    rect.Height = rect.Height + (rect.Y - outside_rect.Y)
                    rect.Y = outside_rect.Y
                End If
            Else
                If rect.Height + (rect.Y - (Tap_LeftUp.Top + 7)) > 64 Then
                    Dim pos As Integer = Tap_LeftUp.Top + e.Y - dcl.Y
                    rect.Height = rect.Height + (rect.Y - (pos + 7))
                    rect.Y = pos + 7
                Else
                    rect.Y = rect.Y + (rect.Height - 64)
                    rect.Height = 64
                End If
            End If
            Draw_Markers()
            RaiseEvent CutRectangleChanged()
        End If
    End Sub
    Private Sub Tap_LeftUp_MouseUp(sender As Object, e As MouseEventArgs) Handles Tap_LeftUp.MouseUp
        new_pos_mode = False
        RaiseEvent CutRectangleChangeEnd()
    End Sub
    Private Sub Tap_LeftDown_MouseDown(sender As Object, e As MouseEventArgs) Handles Tap_LeftDown.MouseDown
        dcl.X = e.X
        dcl.Y = e.Y
        new_pos_mode = True
    End Sub
    Private Sub Tap_LeftDown_MouseMove(sender As Object, e As MouseEventArgs) Handles Tap_LeftDown.MouseMove
        If new_pos_mode = True Then
            If Tap_LeftDown.Left > Tap_LeftDown.Left + e.X - dcl.X Then
                If Tap_LeftDown.Left + e.X - dcl.X > outside_rect.X - 7 Then
                    Dim pos As Integer = Tap_LeftDown.Left + e.X - dcl.X
                    rect.Width = rect.Width + (rect.X - (pos + 7))
                    rect.X = pos + 7
                Else
                    rect.Width = rect.Width + (rect.X - outside_rect.X)
                    rect.X = outside_rect.X
                End If
            Else
                If rect.Width + (rect.X - (Tap_LeftDown.Left + 7)) > 64 Then
                    Dim pos As Integer = Tap_LeftDown.Left + e.X - dcl.X
                    rect.Width = rect.Width + (rect.X - (pos + 7))
                    rect.X = pos + 7
                Else
                    rect.X = rect.X + (rect.Width - 64)
                    rect.Width = 64
                End If
            End If
            If Tap_LeftDown.Top < Tap_LeftDown.Top + e.Y - dcl.Y Then
                If Tap_LeftDown.Top + e.Y - dcl.Y < outside_rect.Y + outside_rect.Height - 7 Then
                    Dim pos As Integer = Tap_LeftDown.Top + e.Y - dcl.Y
                    rect.Height = pos - rect.Y + 7
                Else
                    rect.Height = outside_rect.Y + outside_rect.Height - rect.Y
                End If
            Else
                If Tap_LeftDown.Top - rect.Y + 4 > 64 Then
                    Dim pos As Integer = Tap_LeftDown.Top + e.Y - dcl.Y
                    rect.Height = pos - rect.Y + 7
                Else
                    rect.Height = 64
                End If
            End If
            Draw_Markers()
            RaiseEvent CutRectangleChanged()
        End If
    End Sub
    Private Sub Tap_LeftDown_MouseUp(sender As Object, e As MouseEventArgs) Handles Tap_LeftDown.MouseUp
        new_pos_mode = False
        RaiseEvent CutRectangleChangeEnd()
    End Sub
    Private Sub Tap_RightUp_MouseDown(sender As Object, e As MouseEventArgs) Handles Tap_RightUp.MouseDown
        dcl.X = e.X
        dcl.Y = e.Y
        new_pos_mode = True
    End Sub
    Private Sub Tap_RightUp_MouseMove(sender As Object, e As MouseEventArgs) Handles Tap_RightUp.MouseMove
        If new_pos_mode = True Then
            If Tap_RightUp.Left < Tap_RightUp.Left + e.X - dcl.X Then
                If Tap_RightUp.Left + e.X - dcl.X < outside_rect.X + outside_rect.Width - 7 Then
                    Dim pos As Integer = Tap_RightUp.Left + e.X - dcl.X
                    rect.Width = pos - rect.X + 7
                Else
                    rect.Width = outside_rect.X + outside_rect.Width - rect.X
                End If
            Else
                If Tap_RightUp.Left - rect.X + 7 > 64 Then
                    Dim pos As Integer = Tap_RightUp.Left + e.X - dcl.X
                    rect.Width = pos - rect.X + 7
                Else
                    rect.Width = 64
                End If
            End If
            If Tap_RightUp.Top > Tap_RightUp.Top + e.Y - dcl.Y Then
                If Tap_RightUp.Top + e.Y - dcl.Y > outside_rect.Y - 7 Then
                    Dim pos As Integer = Tap_RightUp.Top + e.Y - dcl.Y
                    rect.Height = rect.Height + (rect.Y - (pos + 7))
                    rect.Y = pos + 7
                Else
                    rect.Height = rect.Height + (rect.Y - outside_rect.Y)
                    rect.Y = outside_rect.Y
                End If
            Else
                If rect.Height + (rect.Y - (Tap_RightUp.Top + 7)) > 64 Then
                    Dim pos As Integer = Tap_RightUp.Top + e.Y - dcl.Y
                    rect.Height = rect.Height + (rect.Y - (pos + 7))
                    rect.Y = pos + 7
                Else
                    rect.Y = rect.Y + (rect.Height - 64)
                    rect.Height = 64
                End If
            End If
            Draw_Markers()
            RaiseEvent CutRectangleChanged()
        End If
    End Sub
    Private Sub Tap_RightUp_MouseUp(sender As Object, e As MouseEventArgs) Handles Tap_RightUp.MouseUp
        new_pos_mode = False
        RaiseEvent CutRectangleChangeEnd()
    End Sub
    Private Sub Tap_RightDown_MouseDown(sender As Object, e As MouseEventArgs) Handles Tap_RightDown.MouseDown
        dcl.X = e.X
        dcl.Y = e.Y
        new_pos_mode = True
    End Sub
    Private Sub Tap_RightDown_MouseMove(sender As Object, e As MouseEventArgs) Handles Tap_RightDown.MouseMove
        If new_pos_mode = True Then
            If Tap_RightDown.Left < Tap_RightDown.Left + e.X - dcl.X Then
                If Tap_RightDown.Left + e.X - dcl.X < outside_rect.X + outside_rect.Width - 7 Then
                    Dim pos As Integer = Tap_RightDown.Left + e.X - dcl.X
                    rect.Width = pos - rect.X + 7
                Else
                    rect.Width = outside_rect.X + outside_rect.Width - rect.X
                End If
            Else
                If Tap_RightDown.Left - rect.X + 7 > 64 Then
                    Dim pos As Integer = Tap_RightDown.Left + e.X - dcl.X
                    rect.Width = pos - rect.X + 7
                Else
                    rect.Width = 64
                End If
            End If
            If Tap_RightDown.Top < Tap_RightDown.Top + e.Y - dcl.Y Then
                If Tap_RightDown.Top + e.Y - dcl.Y < outside_rect.Y + outside_rect.Height - 7 Then
                    Dim pos As Integer = Tap_RightDown.Top + e.Y - dcl.Y
                    rect.Height = pos - rect.Y + 7
                Else
                    rect.Height = outside_rect.Y + outside_rect.Height - rect.Y
                End If
            Else
                If Tap_RightDown.Top - rect.Y + 7 > 64 Then
                    Dim pos As Integer = Tap_RightDown.Top + e.Y - dcl.Y
                    rect.Height = pos - rect.Y + 7
                Else
                    rect.Height = 64
                End If
            End If
            Draw_Markers()
            RaiseEvent CutRectangleChanged()
        End If
    End Sub
    Private Sub Tap_RightDown_MouseUp(sender As Object, e As MouseEventArgs) Handles Tap_RightDown.MouseUp
        new_pos_mode = False
        RaiseEvent CutRectangleChangeEnd()
    End Sub
    Private Sub Cut_UpBarier_MouseDown(sender As Object, e As MouseEventArgs) Handles Cut_UpBarier.MouseDown
        dcl.X = e.X
        dcl.Y = e.Y
        new_pos_mode = True
    End Sub
    Private Sub Cut_UpBarier_MouseMove(sender As Object, e As MouseEventArgs) Handles Cut_UpBarier.MouseMove
        If new_pos_mode = True Then
            If rect.X < Cut_UpBarier.Left + e.X - dcl.X Then
                If Cut_UpBarier.Left + e.X - dcl.X < outside_rect.X + outside_rect.Width - rect.Width - 1 Then
                    rect.X = Cut_UpBarier.Left + e.X - dcl.X
                Else
                    rect.X = outside_rect.X + outside_rect.Width - rect.Width
                End If
            Else
                If Cut_UpBarier.Left + e.X - dcl.X > outside_rect.X - 1 Then
                    rect.X = Cut_UpBarier.Left + e.X - dcl.X
                Else
                    rect.X = outside_rect.X
                End If
            End If
            If rect.Y < Cut_UpBarier.Top + e.Y - dcl.Y Then
                If Cut_UpBarier.Top + e.Y - dcl.Y < outside_rect.Y + outside_rect.Height - rect.Height - 1 Then
                    rect.Y = Cut_UpBarier.Top + e.Y - dcl.Y
                Else
                    rect.Y = outside_rect.Y + outside_rect.Height - rect.Height
                End If
            Else
                If Cut_UpBarier.Top + e.Y - dcl.Y > outside_rect.Y - 1 Then
                    rect.Y = Cut_UpBarier.Top + e.Y - dcl.Y
                Else
                    rect.Y = outside_rect.Y
                End If
            End If
            Draw_Markers()
            RaiseEvent CutRectangleChanged()
        End If
    End Sub
    Private Sub Cut_UpBarier_MouseUp(sender As Object, e As MouseEventArgs) Handles Cut_UpBarier.MouseUp
        new_pos_mode = False
        RaiseEvent CutRectangleChangeEnd()
    End Sub
    Private Sub Cut_LeftBarier_MouseDown(sender As Object, e As MouseEventArgs) Handles Cut_LeftBarier.MouseDown
        dcl.X = e.X
        dcl.Y = e.Y
        new_pos_mode = True
    End Sub
    Private Sub Cut_LeftBarier_MouseMove(sender As Object, e As MouseEventArgs) Handles Cut_LeftBarier.MouseMove
        If new_pos_mode = True Then
            If rect.X < Cut_LeftBarier.Left + e.X - dcl.X Then
                If Cut_LeftBarier.Left + e.X - dcl.X < outside_rect.X + outside_rect.Width - rect.Width - 1 Then
                    rect.X = Cut_LeftBarier.Left + e.X - dcl.X
                Else
                    rect.X = outside_rect.X + outside_rect.Width - rect.Width
                End If
            Else
                If Cut_LeftBarier.Left + e.X - dcl.X > outside_rect.X - 1 Then
                    rect.X = Cut_LeftBarier.Left + e.X - dcl.X
                Else
                    rect.X = outside_rect.X
                End If
            End If
            If rect.Y < Cut_LeftBarier.Top + e.Y - dcl.Y Then
                If Cut_LeftBarier.Top + e.Y - dcl.Y < outside_rect.Y + outside_rect.Height - rect.Height - 1 Then
                    rect.Y = Cut_LeftBarier.Top + e.Y - dcl.Y
                Else
                    rect.Y = outside_rect.Y + outside_rect.Height - rect.Height
                End If
            Else
                If Cut_LeftBarier.Top + e.Y - dcl.Y > outside_rect.Y - 1 Then
                    rect.Y = Cut_LeftBarier.Top + e.Y - dcl.Y
                Else
                    rect.Y = outside_rect.Y
                End If
            End If
            Draw_Markers()
            RaiseEvent CutRectangleChanged()
        End If
    End Sub
    Private Sub Cut_LeftBarier_MouseUp(sender As Object, e As MouseEventArgs) Handles Cut_LeftBarier.MouseUp
        new_pos_mode = False
        RaiseEvent CutRectangleChangeEnd()
    End Sub
    Private Sub Cut_RightBarier_MouseDown(sender As Object, e As MouseEventArgs) Handles Cut_RightBarier.MouseDown
        dcl.X = e.X
        dcl.Y = e.Y
        new_pos_mode = True
    End Sub
    Private Sub Cut_RightBarier_MouseMove(sender As Object, e As MouseEventArgs) Handles Cut_RightBarier.MouseMove
        If new_pos_mode = True Then
            If rect.X < Cut_RightBarier.Left + e.X - dcl.X - rect.Width Then
                If Cut_RightBarier.Left + e.X - dcl.X < outside_rect.X + outside_rect.Width - 1 Then
                    rect.X = Cut_RightBarier.Left + e.X - dcl.X - rect.Width
                Else
                    rect.X = outside_rect.X + outside_rect.Width - rect.Width
                End If
            Else
                If Cut_RightBarier.Left + e.X - dcl.X - rect.Width > outside_rect.X - 1 Then
                    rect.X = Cut_RightBarier.Left + e.X - dcl.X - rect.Width
                Else
                    rect.X = outside_rect.X
                End If
            End If
            If rect.Y < Cut_RightBarier.Top + e.Y - dcl.Y Then
                If Cut_RightBarier.Top + e.Y - dcl.Y < outside_rect.Y + outside_rect.Height - rect.Height - 1 Then
                    rect.Y = Cut_RightBarier.Top + e.Y - dcl.Y
                Else
                    rect.Y = outside_rect.Y + outside_rect.Height - rect.Height
                End If
            Else
                If Cut_RightBarier.Top + e.Y - dcl.Y > outside_rect.Y - 1 Then
                    rect.Y = Cut_RightBarier.Top + e.Y - dcl.Y
                Else
                    rect.Y = outside_rect.Y
                End If
            End If
            Draw_Markers()
            RaiseEvent CutRectangleChanged()
        End If
    End Sub
    Private Sub Cut_RightBarier_MouseUp(sender As Object, e As MouseEventArgs) Handles Cut_RightBarier.MouseUp
        new_pos_mode = False
        RaiseEvent CutRectangleChangeEnd()
    End Sub
    Private Sub Cut_DownBarier_MouseDown(sender As Object, e As MouseEventArgs) Handles Cut_DownBarier.MouseDown
        dcl.X = e.X
        dcl.Y = e.Y
        new_pos_mode = True
    End Sub
    Private Sub Cut_DownBarier_MouseMove(sender As Object, e As MouseEventArgs) Handles Cut_DownBarier.MouseMove
        If new_pos_mode = True Then
            If rect.X < Cut_DownBarier.Left + e.X - dcl.X Then
                If Cut_DownBarier.Left + e.X - dcl.X < outside_rect.X + outside_rect.Width - rect.Width - 1 Then
                    rect.X = Cut_DownBarier.Left + e.X - dcl.X
                Else
                    rect.X = outside_rect.X + outside_rect.Width - rect.Width
                End If
            Else
                If Cut_DownBarier.Left + e.X - dcl.X > outside_rect.X - 1 Then
                    rect.X = Cut_DownBarier.Left + e.X - dcl.X
                Else
                    rect.X = outside_rect.X
                End If
            End If
            If rect.Y < Cut_DownBarier.Top + e.Y - dcl.Y - rect.Height Then
                If Cut_DownBarier.Top + e.Y - dcl.Y < outside_rect.Y + outside_rect.Height - 1 Then
                    rect.Y = Cut_DownBarier.Top + e.Y - dcl.Y - rect.Height
                Else
                    rect.Y = outside_rect.Y + outside_rect.Height - rect.Height
                End If
            Else
                If Cut_DownBarier.Top + e.Y - dcl.Y - rect.Height > outside_rect.Y - 1 Then
                    rect.Y = Cut_DownBarier.Top + e.Y - dcl.Y - rect.Height
                Else
                    rect.Y = outside_rect.Y
                End If
            End If
            Draw_Markers()
            RaiseEvent CutRectangleChanged()
        End If
    End Sub
    Private Sub Cut_DownBarier_MouseUp(sender As Object, e As MouseEventArgs) Handles Cut_DownBarier.MouseUp
        new_pos_mode = False
        RaiseEvent CutRectangleChangeEnd()
    End Sub
    Private Sub Me_EnabledChanged(sender As Object, e As EventArgs) Handles MyBase.EnabledChanged
        If Me.Enabled = True Then
            Cut_UpBarier.BackColor = Color.DarkBlue
            Cut_LeftBarier.BackColor = Color.DarkBlue
            Cut_DownBarier.BackColor = Color.DarkBlue
            Cut_RightBarier.BackColor = Color.DarkBlue
            Tap_Up.BackColor = Color.DarkBlue
            Tap_Left.BackColor = Color.DarkBlue
            Tap_Down.BackColor = Color.DarkBlue
            Tap_Right.BackColor = Color.DarkBlue
            Tap_LeftUp.BackColor = Color.DarkBlue
            Tap_RightUp.BackColor = Color.DarkBlue
            Tap_LeftDown.BackColor = Color.DarkBlue
            Tap_RightDown.BackColor = Color.DarkBlue
        Else
            Cut_UpBarier.BackColor = Color.DimGray
            Cut_LeftBarier.BackColor = Color.DimGray
            Cut_DownBarier.BackColor = Color.DimGray
            Cut_RightBarier.BackColor = Color.DimGray
            Tap_Up.BackColor = Color.DimGray
            Tap_Left.BackColor = Color.DimGray
            Tap_Down.BackColor = Color.DimGray
            Tap_Right.BackColor = Color.DimGray
            Tap_LeftUp.BackColor = Color.DimGray
            Tap_RightUp.BackColor = Color.DimGray
            Tap_LeftDown.BackColor = Color.DimGray
            Tap_RightDown.BackColor = Color.DimGray
        End If
    End Sub
    Public Class Rect_temp
        Public X, Y, Width, Height As Integer
    End Class
End Class
