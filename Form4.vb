Imports System.Data.OleDb
Imports System.Data.SqlClient
Public Class myDatabaseForm
    Dim myOleDbConnection As OleDbConnection
    Dim myOleDbDataAdapter As OleDbDataAdapter
    Dim myDataSet As DataSet
    Dim myDataBaseTable As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try


            Dim cnStr As String = con_ConnectionString
            Dim sql As String = "Select *  from Main"
            myOleDbConnection = New OleDbConnection(cnStr)
            myOleDbDataAdapter = New OleDbDataAdapter(sql, myOleDbConnection)
            myDataSet = New DataSet
            myOleDbDataAdapter.Fill(myDataSet, "Main")
            BindingSource1.DataSource = myDataSet.Tables(0)
            BindingNavigator1.BindingSource = BindingSource1
            DataGridView1.DataSource = BindingSource1
            myOleDbConnection.Close()
            Button2.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            Dim myOleDbCommandBuilder As OleDbCommandBuilder = New OleDbCommandBuilder(myOleDbDataAdapter)
            myOleDbDataAdapter.Update(myDataSet, myDataBaseTable)
            Button2.Enabled = False
            MsgBox("保存成功")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub MyDatabaseForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.Add("所有视频")
        ComboBox1.Items.Add("我的喜爱")
        ComboBox1.Items.Add("骑兵")
        ComboBox1.Items.Add("步兵")
        myDataBaseTable = "Main"
        ComboBox1.Text = "所有视频"
        Button2.Enabled = False
        Button1_Click(sender, e)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub



    Private Sub ToolStripTextBox1_Click(sender As Object, e As EventArgs) Handles ToolStripTextBox1.Click
        If ToolStripTextBox1.Text = "搜索番号" Then
            ToolStripTextBox1.Text = ""
            ToolStripTextBox1.ForeColor = Color.Black
        End If
    End Sub

    Private Sub ToolStripTextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ToolStripTextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            ToolStripButton1_Click(sender, e)
        End If
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Try
            Dim cnStr As String = con_ConnectionString
            Dim sql As String = "Select *  from " & myDataBaseTable & " where fanhao like '%" & ToolStripTextBox1.Text & "%'"
            myOleDbConnection = New OleDbConnection(cnStr)
            myOleDbDataAdapter = New OleDbDataAdapter(sql, myOleDbConnection)
            myDataSet = New DataSet
            myOleDbDataAdapter.Fill(myDataSet, myDataBaseTable)
            BindingSource1.DataSource = myDataSet.Tables(0)
            BindingNavigator1.BindingSource = BindingSource1
            DataGridView1.DataSource = BindingSource1
            myOleDbConnection.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub ToolStripTextBox1_LostFocus(sender As Object, e As EventArgs) Handles ToolStripTextBox1.LostFocus
        If ToolStripTextBox1.Text = "" Then
            ToolStripTextBox1.Text = "搜索番号"
            ToolStripTextBox1.ForeColor = Color.Gray
        End If
    End Sub
End Class