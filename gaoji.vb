Imports System.Data.OleDb
Public Class gaoji
    Private Sub Gaoji_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
        ComboBox3.SelectedIndex = 0
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim str1, str2 As String
        str1 = UCase(TextBox1.Text)
        str2 = UCase(TextBox2.Text)
        If str1 = "" Or str2 = "" Then Exit Sub
        Dim sql As String
        sql = ""
        If IsFileInUse(DataBaseSavePath & "\javbus.mdb") = False Then
            Select Case ComboBox1.Text
                Case "番号"
                    sql = "select biaoqian,fanhao from Main where fanhao like '%" & str1 & "%'"
                Case "名称"
                    sql = "select biaoqian,fanhao from Main where mingcheng like '%" & Jencode(str1) & "%'"
            End Select
            Dim con As New OleDbConnection
            Dim cmd As New OleDbCommand
            Dim dr As OleDbDataReader
            Dim mybiaoqian(0) As String
            Dim myfanhao(0) As String
            myfanhao(0) = ""
            mybiaoqian(0) = ""
            Try
                con.ConnectionString = con_ConnectionString
                con.Open()
            Catch ex As Exception
                MsgBox(ex.Message)
                Exit Sub
            End Try

            cmd.Connection = con '初始化OLEDB命令的连接属性为con
            PaixuSql = sql
            cmd.CommandText = PaixuSql
            dr = cmd.ExecuteReader() '执行OLEDB命令以ExecuteReader()方式，并返回一个OLEDBReader，赋值给dr
            Dim i As Int16
            i = 0
            While (dr.Read)
                myfanhao(i) = dr.Item("fanhao").ToString
                mybiaoqian(i) = dr.Item("biaoqian").ToString
                i = i + 1
                ReDim Preserve myfanhao(i)
                ReDim Preserve mybiaoqian(i)
            End While
            If UBound(myfanhao) > 0 And myfanhao(0) <> "" Then
                ReDim Preserve myfanhao(i - 1)
                ReDim Preserve mybiaoqian(i - 1)
            End If
            dr.Close()
            ' con.Close()
            Dim myDataAdapter As New OleDbDataAdapter()
            cmd.Connection = con '初始化OLEDB命令的连接属性为con

            For i = 0 To UBound(myfanhao) Step 1
                If InStr(mybiaoqian(i), Jencode(str2)) > 0 Then
                Else
                    mybiaoqian(i) = mybiaoqian(i) & " " & Jencode(str2)
                End If
                cmd.CommandText = "update Main set biaoqian='" & mybiaoqian(i) & "' where fanhao='" & myfanhao(i) & "'"
                myDataAdapter.UpdateCommand = cmd
                cmd.ExecuteNonQuery()
            Next
            con.Close()

            MsgBox("成功！")

        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim str1, str2 As String
        str1 = UCase(TextBox4.Text)
        If str1 = "" Then Exit Sub
        Dim sql As String
        sql = ""
        If IsFileInUse(DataBaseSavePath & "\javbus.mdb") = False Then
            Select Case ComboBox1.Text
                Case "番号"
                    sql = "select fanhao from Main where fanhao like '%" & str1 & "%'"
            End Select
            Dim con As New OleDbConnection
            Dim cmd As New OleDbCommand
            Dim dr As OleDbDataReader
            Dim myfanhao(0) As String
            myfanhao(0) = ""
            Dim myshipinleixing As Integer
            Select Case ComboBox3.Text
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
            Try
                con.ConnectionString = con_ConnectionString
                con.Open()
            Catch ex As Exception
                MsgBox(ex.Message)
                Exit Sub
            End Try

            cmd.Connection = con '初始化OLEDB命令的连接属性为con
            PaixuSql = sql
            cmd.CommandText = PaixuSql
            dr = cmd.ExecuteReader() '执行OLEDB命令以ExecuteReader()方式，并返回一个OLEDBReader，赋值给dr
            Dim i As Int16
            i = 0
            While (dr.Read)
                myfanhao(i) = dr.Item("fanhao").ToString
                i = i + 1
                ReDim Preserve myfanhao(i)
            End While
            If UBound(myfanhao) > 0 And myfanhao(0) <> "" Then
                ReDim Preserve myfanhao(i - 1)
            End If
            dr.Close()
            ' con.Close()
            Dim myDataAdapter As New OleDbDataAdapter()
            cmd.Connection = con '初始化OLEDB命令的连接属性为con

            For i = 0 To UBound(myfanhao) Step 1
                cmd.CommandText = "update Main set shipinleixing='" & myshipinleixing & "' where fanhao='" & myfanhao(i) & "'"
                myDataAdapter.UpdateCommand = cmd
                cmd.ExecuteNonQuery()
            Next
            con.Close()

            MsgBox("成功！")

        End If
    End Sub
End Class