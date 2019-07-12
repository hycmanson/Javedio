Imports System.Data.OleDb
Public Class NewVedio
    Private Sub NewVedio_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.InitialDirectory = "c:\"
        OpenFileDialog1.Title = "选择一个视频"
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Filter = "视频文件(*.avi, *.mp4, *.mkv, *.mpg, *.rmvb, *.rm, *.mov, *.mpeg, *.flv, *.wmv, *.m4v)|*.avi; *.mp4; *.mkv; *.mpg; *.rmvb; *.rm; *.mov; *.mpeg; *.flv; *.wmv; *.m4v|All files (*.*)|*.*"
        OpenFileDialog1.FilterIndex = 1
        OpenFileDialog1.RestoreDirectory = True
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            Debug.Print(OpenFileDialog1.FileName)
            If OpenFileDialog1.FileName <> "" Then
                TextBox1.Text = OpenFileDialog1.FileName

            End If
        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Select Case getshipinLeixing(TextBox2.Text)
            Case 0
                ComboBox1.SelectedIndex = 0
            Case 1
                ComboBox1.SelectedIndex = 1
            Case 2
                ComboBox1.SelectedIndex = 2
            Case 3
                ComboBox1.SelectedIndex = 3
            Case 4
                ComboBox1.SelectedIndex = 4
            Case Else
                ComboBox1.SelectedIndex = 0
        End Select
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Label2.Text = GetFileName(TextBox1.Text)
        TextBox2.Text = getFanhao(TextBox1.Text)
        Label6.Text = getFileSize(My.Computer.FileSystem.GetFileInfo(TextBox1.Text).Length)
    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label4_MouseMove(sender As Object, e As MouseEventArgs)

    End Sub

    Private Sub Label4_MouseLeave(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim con As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dr As OleDbDataReader
        Dim FileSize As Long
        Dim mySql As String
        Dim str1 As String
        Dim fanhao As String
        Dim sql1 As String
        Dim FanhaoExist As Boolean
        FanhaoExist = False
        Dim yuanFileSize As Long
        Dim yuanWeizhi As String
        Dim myshipinleixing As Integer
        yuanWeizhi = ""
        yuanFileSize = 0
        myshipinleixing = 0

        If TextBox2.Text = "" Then
            MsgBox("番号不能为空！")
            Exit Sub
            TextBox2.Focus()
        End If
        Select Case ComboBox1.SelectedIndex
            Case 0
                myshipinleixing = 0
            Case 1
                myshipinleixing = 1
            Case 2
                myshipinleixing = 2
            Case 3
                myshipinleixing = 3
            Case 4
                myshipinleixing = 4
            Case Else
                MsgBox("请选择视频类型！")
                Exit Sub
        End Select
        Try

            con.ConnectionString = con_ConnectionString
            con.Open()

        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try
        cmd.Connection = con '初始化OLEDB命令的连接属性为con
        mySql = "select * from Main"
        cmd.CommandText = mySql
        dr = cmd.ExecuteReader() '执行OLEDB命令以ExecuteReader()方式，并返回一个OLEDBReader，赋值给dr
        str1 = TextBox1.Text
        FileSize = My.Computer.FileSystem.GetFileInfo(str1).Length
        fanhao = TextBox2.Text
        Dim myDataAdapter As New OleDbDataAdapter()
        While (dr.Read)
            If dr.Item("fanhao").ToString = fanhao Then '存在的话则更新
                'Debug.Print("该番号存在")
                yuanFileSize = CType(dr.Item("wenjiandaxiao").ToString, Long)
                yuanWeizhi = dr.Item("weizhi").ToString
                FanhaoExist = True
            End If
        End While
        dr.Close()
        con.Close()
        con.ConnectionString = con_ConnectionString
        con.Open()
        cmd.Connection = con '初始化OLEDB命令的连接属性为con
        If FanhaoExist = True Then '数据库里存在
            If yuanFileSize < FileSize Then  '如果新加入的文件较大，则更新位置和大小
                sql1 = "update Main set weizhi='" & str1 & "',wenjiandaxiao='" & FileSize & "' where fanhao='" & fanhao & "'"
                cmd.CommandText = sql1
                myDataAdapter.UpdateCommand = cmd
                cmd.ExecuteNonQuery()
                'Debug.Print("新加入的文件较大，则更新位置和大小")
                Form1.ToolStripLabel6.Text = "新加入的文件较大，更新位置和大小 "
            ElseIf yuanFileSize >= FileSize And Dir(yuanWeizhi) = "" Then  '如果新加入的文件较小，但是原文件被删除了，也更新位置和大小
                sql1 = "update Main set weizhi='" & str1 & "',wenjiandaxiao='" & FileSize & "' where fanhao='" & fanhao & "'"
                cmd.CommandText = sql1
                myDataAdapter.UpdateCommand = cmd
                cmd.ExecuteNonQuery()
                'Debug.Print("新加入的文件较小，但是原文件被删除了，也更新位置和大小")
                Form1.ToolStripLabel6.Text = "原文件被删除了，更新位置和大小 "
            Else
                'Debug.Print("什么也不改变")
                Form1.ToolStripLabel6.Text = "该番号已存在！ "
                MsgBox("该番号已存在！ ")
            End If
        Else '数据库里不存在
            Dim thetime As String = Now.ToString("yyyyMMddHHmmss")
            mySql = "insert into Main (fanhao,weizhi,wenjiandaxiao,shipinleixing,daorushijian) values('" & fanhao & "','" & str1 & "','" & FileSize & "','" & myshipinleixing & "','" & thetime & "')"
            cmd.CommandText = mySql
            myDataAdapter.UpdateCommand = cmd
            cmd.ExecuteNonQuery()
            Form1.ToolStripLabel6.Text = "该番号不存在，已经添加入数据库 "
        End If
        con.Close()
    End Sub
End Class