Imports System.Data.OleDb
Public Class daoruxiangqing
    Private Sub Daoruxiangqing_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ListBox1.Items.Clear()
        For Each s In myChongfufanhaoIndex
            ListBox1.Items.Add(myScanForm.ListView3.Items.Item(s).Text)
        Next
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        For i = UBound(myChongfufanhaoIndex) To 0 Step -1
            myScanForm.ListBox3.Items.Add(myScanForm.ListBox1.Items.Item(myChongfufanhaoIndex(i)))
            myScanForm.ListView3.Items.RemoveAt(myChongfufanhaoIndex(i))
            myScanForm.ListBox1.Items.RemoveAt(myChongfufanhaoIndex(i))
        Next
        CanImport(2) = True
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        myScanForm.PictureBox3.Image = My.Resources.Resource1.waiting
        myScanForm.PictureBox3.Visible = False
        myScanForm.PictureBox4.Image = My.Resources.Resource1.waiting
        myScanForm.PictureBox4.Visible = False
        CanImport(0) = False
        CanImport(1) = False
        CanImport(2) = False
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim con As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim myDataAdapter As New OleDbDataAdapter
        Try
            con.ConnectionString = con_ConnectionString
            con.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try


        Dim sql1 As String
        Dim FileSize As Long
        Dim myVedioWeizhi As String
        myVedioWeizhi = ""
        FileSize = 0
        cmd.Connection = con '初始化OLEDB命令的连接属性为con
        Dim myshipinleixing As Integer
        Dim mylove As Integer
        Select Case myScanForm.ComboBox1.Text
            Case "自动识别"
            Case "日本骑兵"
                myshipinleixing = 2
            Case "日本步兵"
                myshipinleixing = 1
            Case "欧美"
                myshipinleixing = 3
            Case "国产"
                myshipinleixing = 4
            Case "动漫"
                myshipinleixing = 5
            Case "其他"
                myshipinleixing = 0
            Case Else
                myshipinleixing = 0
        End Select

        Select Case myScanForm.ComboBox2.Text
            Case "0-不喜欢"
                mylove = 0
            Case 1
                mylove = 1
            Case 2
                mylove = 2
            Case 3
                mylove = 3
            Case 4
                mylove = 4
            Case 5
                mylove = 5
            Case Else
                mylove = 0
        End Select
        Dim myfanhao As String

        For i = UBound(myChongfufanhaoIndex) To 0 Step -1
            'index出错？
            myfanhao = myScanForm.ListView3.Items.Item(myChongfufanhaoIndex(i)).Text
            myVedioWeizhi = myScanForm.ListBox1.Items.Item(myChongfufanhaoIndex(i))
            FileSize = My.Computer.FileSystem.GetFileInfo(myVedioWeizhi).Length
            cmd.Connection = con '初始化OLEDB命令的连接属性为con
            If myScanForm.ComboBox1.Text = "自动识别" Then
                sql1 = "update Main set weizhi='" & myVedioWeizhi & "',wenjiandaxiao='" & FileSize & "',shipinleixing='" & getshipinLeixing(myfanhao) & "',love='" & mylove & "' where fanhao='" & myfanhao & "'"
            Else
                sql1 = "update Main set weizhi='" & myVedioWeizhi & "',wenjiandaxiao='" & FileSize & "',shipinleixing='" & myshipinleixing & "',love='" & mylove & "' where fanhao='" & myfanhao & "'"
            End If
            cmd.CommandText = sql1
            myDataAdapter.UpdateCommand = cmd
            cmd.ExecuteNonQuery()
            myScanForm.ListBox1.Items.RemoveAt(myChongfufanhaoIndex(i))
            myScanForm.ListView3.Items.RemoveAt(myChongfufanhaoIndex(i))
        Next
        con.Close()
        CanImport(2) = True
        MsgBox("番号更新成功！")

        Me.Close()
    End Sub
End Class