<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Image_Cut
    Inherits System.Windows.Forms.UserControl

    'Пользовательский элемент управления (UserControl) переопределяет метод Dispose для очистки списка компонентов.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Является обязательной для конструктора форм Windows Forms
    Private components As System.ComponentModel.IContainer

    'Примечание: следующая процедура является обязательной для конструктора форм Windows Forms
    'Для ее изменения используйте конструктор форм Windows Form.  
    'Не изменяйте ее в редакторе исходного кода.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.CutImagePrewiev = New System.Windows.Forms.PictureBox()
        Me.Tap_LeftDown = New System.Windows.Forms.Label()
        Me.Tap_LeftUp = New System.Windows.Forms.Label()
        Me.Tap_RightUp = New System.Windows.Forms.Label()
        Me.Tap_RightDown = New System.Windows.Forms.Label()
        Me.Tap_Up = New System.Windows.Forms.Label()
        Me.Tap_Left = New System.Windows.Forms.Label()
        Me.Tap_Down = New System.Windows.Forms.Label()
        Me.Tap_Right = New System.Windows.Forms.Label()
        Me.Cut_UpBarier = New System.Windows.Forms.Label()
        Me.Cut_RightBarier = New System.Windows.Forms.Label()
        Me.Cut_LeftBarier = New System.Windows.Forms.Label()
        Me.Cut_DownBarier = New System.Windows.Forms.Label()
        CType(Me.CutImagePrewiev, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'CutImagePrewiev
        '
        Me.CutImagePrewiev.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CutImagePrewiev.Location = New System.Drawing.Point(0, 0)
        Me.CutImagePrewiev.Name = "CutImagePrewiev"
        Me.CutImagePrewiev.Size = New System.Drawing.Size(298, 298)
        Me.CutImagePrewiev.TabIndex = 0
        Me.CutImagePrewiev.TabStop = False
        '
        'Tap_LeftDown
        '
        Me.Tap_LeftDown.BackColor = System.Drawing.Color.DarkBlue
        Me.Tap_LeftDown.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Tap_LeftDown.Cursor = System.Windows.Forms.Cursors.SizeNESW
        Me.Tap_LeftDown.Location = New System.Drawing.Point(29, 219)
        Me.Tap_LeftDown.Name = "Tap_LeftDown"
        Me.Tap_LeftDown.Size = New System.Drawing.Size(15, 15)
        Me.Tap_LeftDown.TabIndex = 25
        '
        'Tap_LeftUp
        '
        Me.Tap_LeftUp.BackColor = System.Drawing.Color.DarkBlue
        Me.Tap_LeftUp.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Tap_LeftUp.Cursor = System.Windows.Forms.Cursors.SizeNWSE
        Me.Tap_LeftUp.Location = New System.Drawing.Point(29, 24)
        Me.Tap_LeftUp.Name = "Tap_LeftUp"
        Me.Tap_LeftUp.Size = New System.Drawing.Size(15, 15)
        Me.Tap_LeftUp.TabIndex = 24
        '
        'Tap_RightUp
        '
        Me.Tap_RightUp.BackColor = System.Drawing.Color.DarkBlue
        Me.Tap_RightUp.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Tap_RightUp.Cursor = System.Windows.Forms.Cursors.SizeNESW
        Me.Tap_RightUp.Location = New System.Drawing.Point(197, 24)
        Me.Tap_RightUp.Name = "Tap_RightUp"
        Me.Tap_RightUp.Size = New System.Drawing.Size(15, 15)
        Me.Tap_RightUp.TabIndex = 23
        '
        'Tap_RightDown
        '
        Me.Tap_RightDown.BackColor = System.Drawing.Color.DarkBlue
        Me.Tap_RightDown.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Tap_RightDown.Cursor = System.Windows.Forms.Cursors.SizeNWSE
        Me.Tap_RightDown.Location = New System.Drawing.Point(197, 219)
        Me.Tap_RightDown.Name = "Tap_RightDown"
        Me.Tap_RightDown.Size = New System.Drawing.Size(15, 15)
        Me.Tap_RightDown.TabIndex = 22
        '
        'Tap_Up
        '
        Me.Tap_Up.BackColor = System.Drawing.Color.DarkBlue
        Me.Tap_Up.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Tap_Up.Cursor = System.Windows.Forms.Cursors.SizeNS
        Me.Tap_Up.Location = New System.Drawing.Point(108, 24)
        Me.Tap_Up.Name = "Tap_Up"
        Me.Tap_Up.Size = New System.Drawing.Size(15, 10)
        Me.Tap_Up.TabIndex = 14
        '
        'Tap_Left
        '
        Me.Tap_Left.BackColor = System.Drawing.Color.DarkBlue
        Me.Tap_Left.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Tap_Left.Cursor = System.Windows.Forms.Cursors.SizeWE
        Me.Tap_Left.Location = New System.Drawing.Point(29, 116)
        Me.Tap_Left.Name = "Tap_Left"
        Me.Tap_Left.Size = New System.Drawing.Size(10, 15)
        Me.Tap_Left.TabIndex = 21
        '
        'Tap_Down
        '
        Me.Tap_Down.BackColor = System.Drawing.Color.DarkBlue
        Me.Tap_Down.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Tap_Down.Cursor = System.Windows.Forms.Cursors.SizeNS
        Me.Tap_Down.Location = New System.Drawing.Point(108, 224)
        Me.Tap_Down.Name = "Tap_Down"
        Me.Tap_Down.Size = New System.Drawing.Size(15, 10)
        Me.Tap_Down.TabIndex = 19
        '
        'Tap_Right
        '
        Me.Tap_Right.BackColor = System.Drawing.Color.DarkBlue
        Me.Tap_Right.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Tap_Right.Cursor = System.Windows.Forms.Cursors.SizeWE
        Me.Tap_Right.Location = New System.Drawing.Point(202, 116)
        Me.Tap_Right.Name = "Tap_Right"
        Me.Tap_Right.Size = New System.Drawing.Size(10, 15)
        Me.Tap_Right.TabIndex = 20
        '
        'Cut_UpBarier
        '
        Me.Cut_UpBarier.BackColor = System.Drawing.Color.DarkBlue
        Me.Cut_UpBarier.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Cut_UpBarier.Cursor = System.Windows.Forms.Cursors.SizeAll
        Me.Cut_UpBarier.Location = New System.Drawing.Point(34, 29)
        Me.Cut_UpBarier.Name = "Cut_UpBarier"
        Me.Cut_UpBarier.Size = New System.Drawing.Size(174, 5)
        Me.Cut_UpBarier.TabIndex = 15
        '
        'Cut_RightBarier
        '
        Me.Cut_RightBarier.BackColor = System.Drawing.Color.DarkBlue
        Me.Cut_RightBarier.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Cut_RightBarier.Cursor = System.Windows.Forms.Cursors.SizeAll
        Me.Cut_RightBarier.Location = New System.Drawing.Point(202, 29)
        Me.Cut_RightBarier.Name = "Cut_RightBarier"
        Me.Cut_RightBarier.Size = New System.Drawing.Size(5, 200)
        Me.Cut_RightBarier.TabIndex = 16
        '
        'Cut_LeftBarier
        '
        Me.Cut_LeftBarier.BackColor = System.Drawing.Color.DarkBlue
        Me.Cut_LeftBarier.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Cut_LeftBarier.Cursor = System.Windows.Forms.Cursors.SizeAll
        Me.Cut_LeftBarier.Location = New System.Drawing.Point(34, 29)
        Me.Cut_LeftBarier.Name = "Cut_LeftBarier"
        Me.Cut_LeftBarier.Size = New System.Drawing.Size(5, 200)
        Me.Cut_LeftBarier.TabIndex = 17
        '
        'Cut_DownBarier
        '
        Me.Cut_DownBarier.BackColor = System.Drawing.Color.DarkBlue
        Me.Cut_DownBarier.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Cut_DownBarier.Cursor = System.Windows.Forms.Cursors.SizeAll
        Me.Cut_DownBarier.Location = New System.Drawing.Point(34, 224)
        Me.Cut_DownBarier.Name = "Cut_DownBarier"
        Me.Cut_DownBarier.Size = New System.Drawing.Size(174, 5)
        Me.Cut_DownBarier.TabIndex = 18
        '
        'Image_Cut
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Controls.Add(Me.Tap_LeftDown)
        Me.Controls.Add(Me.Tap_LeftUp)
        Me.Controls.Add(Me.Tap_RightUp)
        Me.Controls.Add(Me.Tap_RightDown)
        Me.Controls.Add(Me.Tap_Up)
        Me.Controls.Add(Me.Tap_Left)
        Me.Controls.Add(Me.Tap_Down)
        Me.Controls.Add(Me.Tap_Right)
        Me.Controls.Add(Me.Cut_UpBarier)
        Me.Controls.Add(Me.Cut_RightBarier)
        Me.Controls.Add(Me.Cut_LeftBarier)
        Me.Controls.Add(Me.Cut_DownBarier)
        Me.Controls.Add(Me.CutImagePrewiev)
        Me.Name = "Image_Cut"
        Me.Size = New System.Drawing.Size(298, 298)
        CType(Me.CutImagePrewiev, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents CutImagePrewiev As PictureBox
    Friend WithEvents Tap_LeftDown As Label
    Friend WithEvents Tap_LeftUp As Label
    Friend WithEvents Tap_RightUp As Label
    Friend WithEvents Tap_RightDown As Label
    Friend WithEvents Tap_Up As Label
    Friend WithEvents Tap_Left As Label
    Friend WithEvents Tap_Down As Label
    Friend WithEvents Tap_Right As Label
    Friend WithEvents Cut_UpBarier As Label
    Friend WithEvents Cut_RightBarier As Label
    Friend WithEvents Cut_LeftBarier As Label
    Friend WithEvents Cut_DownBarier As Label
End Class
