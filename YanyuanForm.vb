Imports System.ComponentModel
Imports System.Data.OleDb
Imports System.IO
Public Class YanyuanForm
    Private Sub ToolStripComboBox1_Click(sender As Object, e As EventArgs) Handles ToolStripComboBox1.Click

    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Dim mylp As FlowLayoutPanel
        Dim mylabel As Label
        For Each lb In FlowLayoutPanel1.Controls
            mylp = lb
            For Each item In mylp.Controls
                If InStr(item.GetType.ToString, "Label") > 0 Then
                    mylabel = item
                    mylabel.BackColor = Color.FromArgb(224, 224, 224)
                    If InStr(mylabel.Text, ToolStripTextBox1.Text) > 0 Then
                        mylabel.BackColor = Color.Yellow
                    End If
                End If
            Next
        Next

    End Sub

    Private Sub YanyuanForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ToolStripComboBox1.SelectedIndex = 1
        'ToolStripButton2_Click(sender, e)
        'FlowLayoutPanel1.Focus()
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        FlowLayoutPanel1.Controls.Clear()
        Dim con As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dr As OleDbDataReader
        Dim sql As String
        Dim num As Integer
        con.ConnectionString = con_ConnectionString
        con.Open()
        cmd.Connection = con '初始化OLEDB命令的连接属性为con
        sql = "select yanyuan from Main where shipinleixing=1"
        Select Case ToolStripComboBox1.Text
            Case "步兵"
                sql = "select yanyuan from Main where shipinleixing=1"
            Case "骑兵"
                sql = "select yanyuan from Main where shipinleixing=2"
            Case "欧美"
                sql = "select yanyuan from Main where shipinleixing=3"
            Case "所有"
                sql = "select yanyuan from Main"
        End Select
        cmd.CommandText = sql
        dr = cmd.ExecuteReader() '执行OLEDB命令以ExecuteReader()方式，并返回一个OLEDBReader，赋值给dr
        num = 0
        Dim myYanyuan(0) As String
        Dim IsYanyuanExist As Boolean
        While (dr.Read)
            IsYanyuanExist = False
            For Each s In Split(dr(0).ToString, " ")
                IsYanyuanExist = False
                For Each Yanyuan In myYanyuan
                    If Yanyuan = s Then IsYanyuanExist = True
                Next
                If IsYanyuanExist = False Then
                    myYanyuan(num) = s
                    num = num + 1
                    ReDim Preserve myYanyuan(num)
                End If
            Next
        End While
        dr.Close()
        con.Close()
        ReDim Preserve myYanyuan(num - 1)
        For Each s In myYanyuan

            Dim myFlowLayoutPanel As New FlowLayoutPanel
            Dim myPicturebox As New PictureBox
            Dim myLabel As New Label
            'myFlowLayoutPanel.Cursor = Windows.Forms.Cursors.Hand
            myFlowLayoutPanel.BackColor = Color.White
            'myFlowLayoutPanel.AutoSize = True
            myFlowLayoutPanel.Width = 145
            myFlowLayoutPanel.Height = 180
            s = Juncode(s)
            'Debug.Print(s)
            myPicturebox.Cursor = Windows.Forms.Cursors.Hand
            myPicturebox.Margin = New Padding(10, 10, 3, 3)
            myPicturebox.Width = 125
            myPicturebox.Height = 125
            'myPicturebox.Image = PictureBox1.Image
            myPicturebox.Name = s
            '设置图片
            Dim PicPath As String = ActressesPicSavePath & "\" & s & ".jpg"
            If Dir(PicPath) <> "" Then
                If My.Computer.FileSystem.GetFileInfo(PicPath).Length > 0 Then '如果图片错误会显示内存不足
                    Dim pFileStream As New FileStream(PicPath, FileMode.Open, FileAccess.Read)
                    myPicturebox.Image = Image.FromStream(pFileStream)
                    pFileStream.Close()
                    pFileStream.Dispose()
                End If
            End If


            AddHandler myPicturebox.Click, AddressOf myLabel_Click

            myLabel.Name = s
            myLabel.Text = s
            myLabel.Width = 145
            myLabel.Height = 30
            myLabel.Margin = New Padding(2, 10, 0, 0)
            myLabel.TextAlign = ContentAlignment.MiddleCenter
            myLabel.AutoSize = False
            myLabel.Font = New System.Drawing.Font("微软雅黑", 12)
            myLabel.BackColor = Color.FromArgb(224, 224, 224)
            myLabel.ForeColor = Color.Black
            myLabel.Cursor = Windows.Forms.Cursors.Hand
            AddHandler myLabel.Click, AddressOf myLabel_Click
            myFlowLayoutPanel.Controls.Add(myPicturebox)
            myFlowLayoutPanel.SetFlowBreak(myPicturebox, True)
            myFlowLayoutPanel.Controls.Add(myLabel)
            FlowLayoutPanel1.Controls.Add(myFlowLayoutPanel)
        Next
        FlowLayoutPanel1.Focus()
    End Sub

    Private Sub myLabel_Click(sender As Object, e As EventArgs)
        Debug.Print(sender.name)
        Clickindex = 11
        myYanyuan = Jencode(sender.name)
        Debug.Print(myYanyuan)
        Try
            Form1.SelectFromDatabase("where yanyuan like '%" & myYanyuan & "%'", "Main", 1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Me.Close()
        Form1.Activate()
        Form1.ToolStripLabel6.Text = "演员为：" & Juncode(myYanyuan)
    End Sub


    Private Sub ToolStripComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        'ToolStripButton2_Click(sender, e)
        'FlowLayoutPanel1.Focus()
    End Sub

    Private Sub ToolStripTextBox1_Click(sender As Object, e As EventArgs) Handles ToolStripTextBox1.Click

    End Sub

    Private Sub ToolStripTextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ToolStripTextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            ToolStripButton1_Click(sender, e)
        End If
    End Sub

    Private Sub FlowLayoutPanel1_Paint(sender As Object, e As PaintEventArgs) Handles FlowLayoutPanel1.Paint

    End Sub

    Private Sub FlowLayoutPanel1_Click(sender As Object, e As EventArgs) Handles FlowLayoutPanel1.Click
        FlowLayoutPanel1.Focus()
    End Sub

    Private Sub YanyuanForm_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        e.Cancel = -1
        Me.Hide()
    End Sub

    Private Sub ToolStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles ToolStrip1.ItemClicked

    End Sub
End Class