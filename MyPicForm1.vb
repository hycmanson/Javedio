Public Class MyPicForm1
    Private blnDragging As Boolean = False
    Private OriginalLocation As Point
    Private MoveToPoint As Point
    Public Filepath(1) As String
    Public num As Integer

    Dim MovBoll As Boolean
    Dim CurrX As Integer
    Dim CurrY As Integer
    Dim MousX As Integer
    Dim MousY As Integer



    Private Sub Form1_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown
        If Me.WindowState <> FormWindowState.Maximized Then


            Me.blnDragging = True
            Me.OriginalLocation = New Point(e.X, e.Y)   '获取鼠标按下后在窗体上的位置坐标
        End If
    End Sub

    Private Sub Form1_MouseUp(sender As Object, e As MouseEventArgs) Handles Me.MouseUp
        Me.blnDragging = False
    End Sub

    Private Sub Form1_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove
        PictureBox2.Visible = False
        PictureBox3.Visible = False
        PictureBox4.Visible = False
        If Me.blnDragging Then
            Me.MoveToPoint = Me.PointToScreen(New Point(e.X, e.Y))  '获取鼠标相对于屏幕的位置坐标
            Me.MoveToPoint.Offset(Me.OriginalLocation.X  -1, Me.OriginalLocation.Y  -1)
            Me.Location = Me.MoveToPoint
            MyPicForm2.Left = Me.Left
            MyPicForm2.Top = Me.Top
        End If
    End Sub

    Private Sub Form1_MouseLeave(sender As Object, e As EventArgs) Handles Me.MouseLeave
        Me.blnDragging = False
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.FormBorderStyle = FormBorderStyle.None
        'Me.WindowState = FormWindowState.Maximized
        Me.BackColor = Color.Blue
        Me.TransparencyKey = Me.BackColor
        'MyPicForm2.Show()
        'MyPicForm2.FormBorderStyle = FormBorderStyle.None
        'MyPicForm2.Left = Me.Location.X
        'MyPicForm2.Top = Me.Location.Y
        'MyPicForm2.Height = Me.Height
        'MyPicForm2.Width = Me.Width
        'MyPicForm2.BackColor = Color.DarkGray
        'MyPicForm2.Opacity = 0.8
        Button1_Click(sender, e)
        'Me.BringToFront()
        Me.Height = PictureBox1.Height
        Me.Width = PictureBox1.Width
        PictureBox1.Left = 0
        PictureBox1.Top = 0
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Dim PicPath As String = Application.StartupPath & "\Pic\HK-416\extra\"
        'Filepath(0) = ""
        'num = 0
        'If Dir(PicPath) <> "" Then
        '    Debug.Print(PicPath)
        '    For Each foundFile In My.Computer.FileSystem.GetFiles(PicPath, FileIO.SearchOption.SearchAllSubDirectories, {".jpg", ".bmp", ".gif"})
        '        Filepath(num) = foundFile
        '        num = num + 1
        '        ReDim Preserve Filepath(num)
        '    Next
        'End If
        'ReDim Preserve Filepath(num - 1)
        'num = 0
        PictureBox5.SizeMode = PictureBoxSizeMode.AutoSize
        PictureBox5.Image = PictureBox1.Image
        Dim PicH As Double
        Dim PicW As Double
        PicW = PictureBox5.Width
        PicH = PictureBox5.Height
        If PicH = 0 Then PicH = 200
        PictureBox1.Width = PictureBox1.Height * (PicW / PicH)
        PictureBox1.SizeMode = PictureBoxSizeMode.CenterImage
        PictureBox1.Image = PictureBox5.Image
        PictureBox1.Left = (Me.Width - PictureBox1.Width) / 2
        PictureBox1.Top = (Me.Height - PictureBox1.Height) / 2
        'Debug.Print(Filepath(0))
    End Sub



    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If ExtraPicIndex >= 1 Then
            ExtraPicIndex = ExtraPicIndex - 1
            PictureBox5.SizeMode = PictureBoxSizeMode.AutoSize
            PictureBox5.Image = myBitmap(ExtraPicIndex)
            Dim PicH As Double
            Dim PicW As Double
            PicW = PictureBox5.Width
            PicH = PictureBox5.Height
            If PicH = 0 Then PicH = 200
            PictureBox1.Width = PictureBox1.Height * (PicW / PicH)
            PictureBox1.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox1.Image = myBitmap(ExtraPicIndex)
            PictureBox1.Left = (Me.Width - PictureBox1.Width) / 2
            PictureBox1.Top = (Me.Height - PictureBox1.Height) / 2
        ElseIf ExtraPicIndex = 0 Then
            ExtraPicIndex = UBound(myBitmap)
            PictureBox5.SizeMode = PictureBoxSizeMode.AutoSize
            PictureBox5.Image = myBitmap(ExtraPicIndex)
            Dim PicH As Double
            Dim PicW As Double
            PicW = PictureBox5.Width
            PicH = PictureBox5.Height
            If PicH = 0 Then PicH = 200
            PictureBox1.Width = PictureBox1.Height * (PicW / PicH)
            PictureBox1.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox1.Image = myBitmap(ExtraPicIndex)
            PictureBox1.Left = (Me.Width - PictureBox1.Width) / 2
            PictureBox1.Top = (Me.Height - PictureBox1.Height) / 2
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If ExtraPicIndex < UBound(myBitmap) Then
            ExtraPicIndex = ExtraPicIndex + 1
            PictureBox5.SizeMode = PictureBoxSizeMode.AutoSize
            PictureBox5.Image = myBitmap(ExtraPicIndex)
            Dim PicH As Double
            Dim PicW As Double
            PicW = PictureBox5.Width
            PicH = PictureBox5.Height
            If PicH = 0 Then PicH = 200
            PictureBox1.Width = PictureBox1.Height * (PicW / PicH)
            PictureBox1.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox1.Image = myBitmap(ExtraPicIndex)
            PictureBox1.Left = (Me.Width - PictureBox1.Width) / 2
            PictureBox1.Top = (Me.Height - PictureBox1.Height) / 2
        ElseIf ExtraPicIndex = UBound(myBitmap) Then
            ExtraPicIndex = 0
            PictureBox5.SizeMode = PictureBoxSizeMode.AutoSize
            PictureBox5.Image = myBitmap(ExtraPicIndex)
            Dim PicH As Double
            Dim PicW As Double
            PicW = PictureBox5.Width
            PicH = PictureBox5.Height
            If PicH = 0 Then PicH = 200
            PictureBox1.Width = PictureBox1.Height * (PicW / PicH)
            PictureBox1.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox1.Image = myBitmap(ExtraPicIndex)
            PictureBox1.Left = (Me.Width - PictureBox1.Width) / 2
            PictureBox1.Top = (Me.Height - PictureBox1.Height) / 2
        End If
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs)
        Me.WindowState = FormWindowState.Minimized

    End Sub



    Private Sub Panel2_Click(sender As Object, e As EventArgs) Handles Panel2.Click
        Button3_Click(sender, e)
    End Sub

    Private Sub Panel1_Click(sender As Object, e As EventArgs) Handles Panel1.Click
        Button2_Click(sender, e)
    End Sub

    Private Sub Panel1_MouseMove(sender As Object, e As MouseEventArgs) Handles Panel1.MouseMove
        Panel1.BackColor = Color.AliceBlue

    End Sub

    Private Sub Panel1_MouseLeave(sender As Object, e As EventArgs) Handles Panel1.MouseLeave
        Panel1.BackColor = Color.Transparent
    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub Panel2_Paint(sender As Object, e As PaintEventArgs) Handles Panel2.Paint

    End Sub

    Private Sub Panel2_MouseMove(sender As Object, e As MouseEventArgs) Handles Panel2.MouseMove
        Panel2.BackColor = Color.AliceBlue
    End Sub

    Private Sub Panel2_MouseLeave(sender As Object, e As EventArgs) Handles Panel2.MouseLeave
        Panel2.BackColor = Color.Transparent
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Debug.Print("click")
    End Sub

    Private Sub PictureBox1_MouseWheel(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseWheel
        Debug.Print(e.Delta)

    End Sub

    Private Sub Form1_MouseWheel(sender As Object, e As MouseEventArgs) Handles Me.MouseWheel
        'Debug.Print(e.Delta)
        Dim p As New Point()
        p.X = System.Windows.Forms.Control.MousePosition.X
        p.Y = System.Windows.Forms.Control.MousePosition.Y
        If p.X < Me.Left + Me.Width And p.X > Me.Left Then
            If p.Y < Me.Top + Me.Height And p.Y > Me.Top Then
                'Debug.Print(e.Delta)

                PictureBox5.SizeMode = PictureBoxSizeMode.AutoSize
                PictureBox5.Image = PictureBox1.Image
                Dim PicH As Double
                Dim PicW As Double
                PicH = PictureBox5.Height
                PicW = PictureBox5.Width
                If PicH = 0 Then PicH = 100
                PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage
                If e.Delta > 0 Then
                    If PictureBox1.Width > My.Computer.Screen.Bounds.Width Or PictureBox1.Height > My.Computer.Screen.Bounds.Height Then
                        Exit Sub
                    End If
                    PictureBox1.Height = PictureBox1.Height * 1.2
                    PictureBox1.Width = PictureBox1.Height * (PicW / PicH)
                    Me.Height = PictureBox1.Height
                    Me.Width = PictureBox1.Width

                    Me.Top = (My.Computer.Screen.Bounds.Height - Me.Height) / 2
                    Me.Left = (My.Computer.Screen.Bounds.Width - Me.Width) / 2
                    PictureBox1.Left = 0
                    PictureBox1.Top = 0

                Else
                    If PictureBox1.Width < 50 Or PictureBox1.Height < 50 Then
                        Exit Sub
                    End If

                    PictureBox1.Height = PictureBox1.Height * 0.8
                    PictureBox1.Width = PictureBox1.Height * (PicW / PicH)

                    Me.Height = PictureBox1.Height
                    Me.Width = PictureBox1.Width

                    Me.Top = (My.Computer.Screen.Bounds.Height - Me.Height) / 2
                    Me.Left = (My.Computer.Screen.Bounds.Width - Me.Width) / 2
                    PictureBox1.Left = 0
                    PictureBox1.Top = 0
                End If

            End If
        End If
    End Sub

    Private Sub PictureBox1_MouseDown(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseDown
        'Form1_MouseDown(sender, e)
        Me.blnDragging = True
        MousX = e.X
        MousY = e.Y
        'MovBoll = True
    End Sub

    Private Sub PictureBox1_MouseMove(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseMove
        If Me.blnDragging Then
            Debug.Print(sender.GetType.ToString)
            'Me.MoveToPoint = Me.PointToScreen(New Point(e.X, e.Y))  '获取鼠标相对于屏幕的位置坐标
            'Me.MoveToPoint.Offset(Me.OriginalLocation.X - 1, Me.OriginalLocation.Y - 1)
            'Me.Location = Me.MoveToPoint

            CurrX = Me.Left - MousX + e.X
            CurrY = Me.Top - MousY + e.Y
            Me.Location = New Point(CurrX, CurrY)
        End If

        'If MovBoll = True Then
        '    CurrX = PictureBox1.Left - MousX + e.X
        '    CurrY = PictureBox1.Top - MousY + e.Y
        '    PictureBox1.Location = New Point(CurrX, CurrY)
        'End If
    End Sub

    Private Sub Form1_MouseCaptureChanged(sender As Object, e As EventArgs) Handles Me.MouseCaptureChanged

    End Sub

    Private Sub PictureBox1_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox1.MouseLeave
        Me.blnDragging = False
    End Sub

    Private Sub PictureBox1_MouseUp(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseUp
        Me.blnDragging = False
        'MovBoll = False
    End Sub

    Private Sub PictureBox3_MouseMove(sender As Object, e As MouseEventArgs)

    End Sub

    Private Sub Panel3_Paint(sender As Object, e As PaintEventArgs) Handles Panel3.Paint

    End Sub

    Private Sub Panel3_MouseMove(sender As Object, e As MouseEventArgs) Handles Panel3.MouseMove
        PictureBox2.Visible = True
        PictureBox3.Visible = True
        PictureBox4.Visible = True
    End Sub

    Private Sub PictureBox4_Click_1(sender As Object, e As EventArgs) Handles PictureBox4.Click
        Me.Close()
        MyPicForm2.Close()
    End Sub

    Private Sub PictureBox2_Click_1(sender As Object, e As EventArgs) Handles PictureBox2.Click
        If Me.WindowState = FormWindowState.Maximized Then
            MyPicForm2.WindowState = FormWindowState.Normal
            Me.WindowState = FormWindowState.Normal
            Me.Top = (My.Computer.Screen.Bounds.Height - Me.Height) / 2
            Me.Left = (My.Computer.Screen.Bounds.Width - Me.Width) / 2
            MyPicForm2.Top = Me.Top
            MyPicForm2.Left = Me.Left

            PictureBox1.Left = (Me.Width - PictureBox1.Width) / 2
            PictureBox1.Top = (Me.Height - PictureBox1.Height) / 2
        Else
            MyPicForm2.WindowState = FormWindowState.Maximized
            Me.WindowState = FormWindowState.Maximized
            PictureBox1.Left = (Me.Width - PictureBox1.Width) / 2
            PictureBox1.Top = (Me.Height - PictureBox1.Height) / 2
        End If
    End Sub

    Private Sub PictureBox3_Click_1(sender As Object, e As EventArgs) Handles PictureBox3.Click
        Me.WindowState = FormWindowState.Minimized
        MyPicForm2.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub Panel3_MouseLeave(sender As Object, e As EventArgs) Handles Panel3.MouseLeave

    End Sub

    Private Sub PictureBox3_SizeChanged(sender As Object, e As EventArgs) Handles PictureBox3.SizeChanged

    End Sub

    Private Sub Form1_Activated(sender As Object, e As EventArgs) Handles Me.Activated


    End Sub

    Private Sub PictureBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox1.KeyDown
        If e.KeyCode = Keys.Left Then
            Button2_Click(sender, e)
        ElseIf e.KeyCode = Keys.Right Then
            Button3_Click(sender, e)
        End If
    End Sub

    Private Sub PictureBox1_BindingContextChanged(sender As Object, e As EventArgs) Handles PictureBox1.BindingContextChanged

    End Sub

    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Left Then
            Button2_Click(sender, e)
        ElseIf e.KeyCode = Keys.Right Then
            Button3_Click(sender, e)
        ElseIf e.KeyCode = Keys.Escape Then
            Me.Close()
            MyPicForm2.Close()
        End If
    End Sub

    Private Sub MyPicForm1_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        If Me.WindowState = FormWindowState.Normal Then
            MyPicForm2.WindowState = FormWindowState.Normal
        End If
    End Sub

    Private Sub MyPicForm1_LostFocus(sender As Object, e As EventArgs) Handles Me.LostFocus
        Me.Close()
    End Sub

    Private Sub Panel2_MouseDown(sender As Object, e As MouseEventArgs) Handles Panel2.MouseDown

    End Sub
End Class
