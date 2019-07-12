Imports System.Data.OleDb
Imports System.ComponentModel

Public Class myScanForm
    Public minFileSize, maxFileSize As Long
    Public myScanThread As System.Threading.Thread
    Public MyImportThread As System.Threading.Thread
    Delegate Sub importWT()
    Public Insertnum As Integer
    Public totalInsertnum As Integer

    Private Sub MyScanForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        PictureBox1.Visible = False
        PictureBox2.Visible = False
        PictureBox3.Visible = False
        PictureBox4.Visible = False
        PictureBox5.Visible = False
        'ImportInfo.Show()
        'ImportInfo.Left = Me.Left + Me.Width
        'ImportInfo.Top = Me.Top
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        ListView1.SmallImageList = ImageList1
        MaskedTextBox1.Text = Split(Scan_WenjianDaxiao, "-")(0)
        MaskedTextBox2.Text = Split(Scan_WenjianDaxiao, "-")(1)
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
        ComboBox3.SelectedIndex = 0
        ToolStripLabel2.Text = "欢迎使用扫描导入工具"
        For i = 0 To 2 Step 1
            CanImport(i) = False
        Next

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs)

        If FolderBrowserDialog1.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
            ListView1.Items.Add(FolderBrowserDialog1.SelectedPath)
            ListView1.Items.Item(ListView1.Items.Count - 1).ImageIndex = 0

        End If
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs)
        For i As Integer = ListView1.SelectedIndices.Count - 1 To 0 Step -1
            ListView1.Items.RemoveAt(ListView1.SelectedIndices(i))
        Next
    End Sub





    Private Sub SearchFile()
        'On Error Resume Next
        ToolStripLabel2.Text = "扫描中………"
        For Each filepath In ListView1.Items
            Dim folderArr = My.Computer.FileSystem.GetDirectoryInfo(filepath.Text)
            '先获取主目录文件
            For Each foundFile In My.Computer.FileSystem.GetFiles(folderArr.FullName, FileIO.SearchOption.SearchTopLevelOnly, mySearchPattern)
                ListBox1.Items.Add(foundFile)
            Next
            '再获取子目录文件
            For Each folder In folderArr.GetDirectories()
                ' ToolStripStatusLabel4.Text = folder.ToString
                Try

                    For Each foundFile In My.Computer.FileSystem.GetFiles(folder.FullName, FileIO.SearchOption.SearchAllSubDirectories, mySearchPattern)
                        '！！！这里不能给toolStripstatuslabel4设置，不然会弹出索引超出范围
                        '获取文件大小
                        If My.Computer.FileSystem.GetFileInfo(foundFile).Length >= minFileSize * 1024 * 1024 And My.Computer.FileSystem.GetFileInfo(foundFile).Length <= maxFileSize * 1024 * 1024 Then
                            ListBox1.Items.Add(foundFile)
                        End If
                    Next
                Catch ex As Exception
                    Continue For
                    Debug.Print(ex.Message)
                End Try
            Next
        Next

        Dim mystr As String()
        For i = 0 To ListBox1.Items.Count - 1
            mystr = Split(ListBox1.Items(i), "\")
            ListView3.Items.Add(getFanhao(mystr(UBound(mystr))))
            'ListView3.Items.
            ListView3.Items.Item(ListView3.Items.Count - 1).ImageIndex = 1
        Next


        ToolStripLabel2.Text = "扫描出" & ListBox1.Items.Count & "个"
    End Sub







    Private Sub ListBox1_SelectedIndexChanged_1(sender As Object, e As EventArgs)
        On Error Resume Next
        ToolStripStatusLabel3.Text = ListBox1.Items(ListBox1.SelectedIndex)
    End Sub

    Private Sub PictureBox1_Click_1(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs)
        'For i As Integer = ListView3.SelectedIndices.Count - 1 To 0 Step -1
        '    ListView3.Items.RemoveAt(ListView3.SelectedIndices(i))
        '    ToolStripStatusLabel3.Text = ListView3.SelectedItems.Item(i).Index
        '    'ListBox1.Items.RemoveAt(ListView3.SelectedItems.Item(i).Index)
        'Next

        ToolStripStatusLabel4.Text = "扫描出" & ListBox1.Items.Count & "个"
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs)

    End Sub


    Private Sub weituo()
        Me.Invoke(New importWT(AddressOf importintodatabase))
    End Sub

    Private Sub importintodatabase()
        '导入数据库
        Dim con As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim myDataAdapter As New OleDbDataAdapter
        Dim sql As String
        Dim str1 As String
        Try


            con.ConnectionString = con_ConnectionString
            con.Open()
            cmd.Connection = con '初始化OLEDB命令的连接属性为con
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try
        Dim FileSize As Long
        Dim myVedioWeizhi As String
        myVedioWeizhi = ""
        Label8.Text = "0%"
        totalInsertnum = ListBox1.Items.Count
        Insertnum = 0
        Timer3.Enabled = True
        Dim thetime As String
        totalInsertnum = ListBox1.Items.Count
        Dim myshipinleixing As String
        myshipinleixing = 0
        Select Case ComboBox1.Text
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
        Dim mylove As Integer
        Select Case ComboBox2.Text
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
        If ComboBox1.Text = "自动识别" Then
            'Try
            For i = 0 To ListBox1.Items.Count - 1
                    Application.DoEvents()
                    Insertnum = Insertnum + 1
                    str1 = ListBox1.Items.Item(i)
                    FileSize = My.Computer.FileSystem.GetFileInfo(str1).Length
                    thetime = Now.ToString("yyyyMMddHHmmss")
                    str1 = Replace(str1, "'", "’")
                sql = "insert into Main (fanhao,weizhi,wenjiandaxiao,shipinleixing,daorushijian,love) values('" & ListView3.Items.Item(i).Text & "','" & str1 & "','" & FileSize & "','" & getshipinLeixing(ListView3.Items.Item(i).Text) & "','" & thetime & "','" & mylove & "')"
                cmd.CommandText = sql
                    myDataAdapter.UpdateCommand = cmd
                    cmd.ExecuteNonQuery()
                Next i
            'Catch ex As Exception
            '    MsgBox(ex.Message)
            'End Try

        Else
            Try
                For i = 0 To ListBox1.Items.Count - 1
                    Application.DoEvents()
                    Insertnum = Insertnum + 1
                    str1 = ListBox1.Items.Item(i)
                    FileSize = My.Computer.FileSystem.GetFileInfo(str1).Length
                    thetime = Now.ToString("yyyyMMddHHmmss")
                    str1 = Replace(str1, "'", "’")
                    sql = "insert into Main (fanhao,weizhi,wenjiandaxiao,shipinleixing,daorushijian,love) values('" & ListView3.Items.Item(i).Text & "','" & str1 & "','" & FileSize & "','" & myshipinleixing & "','" & thetime & "','" & mylove & "')"
                    cmd.CommandText = sql
                    myDataAdapter.UpdateCommand = cmd
                    cmd.ExecuteNonQuery()
                Next i
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If


        con.Close()
        ToolStripLabel2.Text = "导入成功！"
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs)

        'Button7.Enabled = True
        ToolStripStatusLabel4.Text = "总共有 " & ListView3.Items.Count & "个"
    End Sub






    Private Sub myScanForm_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If myScanThread IsNot Nothing And MyImportThread IsNot Nothing Then


            Try
                myScanThread.Abort()
                myScanThread.Join()
                MyImportThread.Abort()
                MyImportThread.Join()
            Catch ex As Exception

            End Try
        End If
    End Sub




    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If myScanThread IsNot Nothing Then
            If myScanThread.IsAlive = False Then
                PictureBox1.Image = My.Resources.Resource1.complete
                PictureBox1.Visible = True
                Timer1.Enabled = False
                ToolStripLabel2.Text = "扫描完成！"
            End If
        End If
        'If ListView3.Items.Count > 0 Then
        '    PictureBox1.Image = My.Resources.Resource1.complete
        '    PictureBox1.Visible = True
        '    Timer1.Enabled = False
        '    ToolStripLabel2.Text = "扫描完成！"
        'End If
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        If MyImportThread IsNot Nothing Then
            If MyImportThread.IsAlive = False Then
                PictureBox2.Image = My.Resources.Resource1.complete
                PictureBox2.Visible = True
                Timer2.Enabled = False
            End If
        End If
    End Sub


    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick
        Dim max As Double
        'max = Maxmyfanhao - 1
        max = totalInsertnum
        If max = 0 Then Exit Sub
        Dim num As Double
        num = Insertnum
        'Debug.Print(num)
        'Debug.Print(num / max)
        If num / max >= 1 Then
            ProgressBar1.Value = ProgressBar1.Maximum
            Label8.Text = "100%"
            'ReDim myThreadIsCompleted(0)
            'myThreadIsCompleted(0) = True
            Timer3.Enabled = False
            ToolStripLabel2.Text = "导入成功！"
        Else
            ProgressBar1.Value = ProgressBar1.Maximum * (num / max)
            Label8.Text = 100 * Math.Round(num / max, 3) & "%"
        End If
    End Sub

    Private Sub Button5_Click_1(sender As Object, e As EventArgs) Handles Button5.Click
        If FolderBrowserDialog1.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
            ListView1.Items.Add(FolderBrowserDialog1.SelectedPath)
            ListView1.Items.Item(ListView1.Items.Count - 1).ImageIndex = 0

        End If
    End Sub

    Private Sub Button6_Click_1(sender As Object, e As EventArgs) Handles Button6.Click
        For i As Integer = ListView1.SelectedIndices.Count - 1 To 0 Step -1
            ListView1.Items.RemoveAt(ListView1.SelectedIndices(i))
        Next
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        ListView1.Clear()
    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub myScanForm_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        'ImportInfo.TopMost = True
    End Sub

    Private Sub TextBox1_Click(sender As Object, e As EventArgs) Handles TextBox1.Click
        If TextBox1.Text = "查找" Then
            TextBox1.Text = ""
            TextBox1.ForeColor = Color.Black
        End If
    End Sub

    Private Sub TabPage2_Click(sender As Object, e As EventArgs) Handles TabPage2.Click
        TabPage2.Focus()
    End Sub

    Private Sub TextBox1_LostFocus(sender As Object, e As EventArgs) Handles TextBox1.LostFocus
        If TextBox1.Text = "" Then
            TextBox1.Text = "查找"
            TextBox1.ForeColor = Color.Gray
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        If ListBox2.Items.Count = 0 And ListBox3.Items.Count = 0 Then
            Exit Sub
        End If


        Dim myLogPath As String = Application.StartupPath & "\Log"
        Dim txtPath As String = myLogPath & "\" & Format(DateTime.Now, "yyyy年MM月dd日 HH：mm：ss") & ".txt"
        If Dir(myLogPath) = "" Then
            My.Computer.FileSystem.CreateDirectory(myLogPath)
        End If
        Dim sw As System.IO.StreamWriter = Nothing
        If Not System.IO.File.Exists(txtPath) Then
            sw = System.IO.File.CreateText(txtPath)
        End If
        Try


            sw.Write("扫描日期：" & Format(DateTime.Now, "yyyy/MM/dd hh:mm:ss"))
            sw.WriteLine()
            sw.Write("--------------------------------")
            sw.WriteLine()
            sw.Write("未识别出番号的视频：")
            sw.WriteLine()
            sw.Write("--------------------------------")
            sw.WriteLine()
            For i = 0 To ListBox2.Items.Count - 1
                sw.Write(ListBox2.Items.Item(i))
                sw.WriteLine()
            Next
            sw.Write("--------------------------------")
            sw.WriteLine()
            sw.Write("重复的视频：")
            sw.WriteLine()
            sw.Write("--------------------------------")
            sw.WriteLine()
            For i = 0 To ListBox3.Items.Count - 1
                sw.Write(ListBox3.Items.Item(i))
                sw.WriteLine()
            Next
            sw.Write("--------------------------------")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        If sw IsNot Nothing Then
            sw.Close()
        End If
        MsgBox("成功！")
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs)
        'Debug.Print(TabControl1.TabPages.tabindex)
        'Debug.Print(TabPage1)
    End Sub

    Private Sub Button2_Click_2(sender As Object, e As EventArgs) Handles Button2.Click
        If ListView1.Items.Count = 0 Then Exit Sub
        '判断是否在扫描和导入
        If myScanThread IsNot Nothing Then
            If myScanThread.IsAlive = True Then
                MsgBox("请等待本次扫描停止或取消扫描！")
                Exit Sub
            End If
        End If
        If MyImportThread IsNot Nothing Then
            If MyImportThread.IsAlive = True Then
                MsgBox("数据库正在导入！请等待")
                Exit Sub
            End If
        End If

        If ListBox1.Items.Count > 0 Then
            If MsgBox("扫描将清空上次结果，是否继续？", vbYesNo) = vbNo Then
                Exit Sub
            End If
        End If
        For i = 0 To 2
            CanImport(i) = False
        Next
        PictureBox2.Image = My.Resources.Resource1.waiting
        PictureBox2.Visible = False
        PictureBox3.Image = My.Resources.Resource1.waiting
        PictureBox3.Visible = False
        PictureBox4.Image = My.Resources.Resource1.waiting
        PictureBox4.Visible = False
        PictureBox5.Image = My.Resources.Resource1.waiting
        PictureBox5.Visible = False

        ToolStripLabel2.Text = "扫描中……"
        Timer1.Enabled = True
        PictureBox1.Image = My.Resources.Resource1.waiting
        PictureBox1.Visible = True
        minFileSize = CType(MaskedTextBox1.Text, Long)
        maxFileSize = CType(MaskedTextBox2.Text, Long)
        Scan_Shipinleixing = ""


        Scan_WenjianDaxiao = Replace(MaskedTextBox1.Text, " ", "") & "-" & Replace(MaskedTextBox2.Text, " ", "")
        Dim myXml As New CSysXML("config.xml")
        myXml.SaveElement("ScanSetup", "Scan_Shipinleixing", Scan_Shipinleixing)
        myXml.SaveElement("ScanSetup", "Scan_WenjianDaxiao", Scan_WenjianDaxiao)
        myXml.SaveElement("ScanSetup", "BubingFanhao", BubingFanhao)
        ListBox1.Items.Clear()
        ListView3.Items.Clear()
        ListBox2.Items.Clear()
        ListBox3.Items.Clear()

        '启动扫描
        myScanThread = New System.Threading.Thread(AddressOf SearchFile)
        myScanThread.Start()
    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        If MsgBox("取消扫描将清除当前扫描结果，是否继续？", vbYesNo) = vbYes Then
            If myScanThread IsNot Nothing Then
                myScanThread.Abort()
                myScanThread.Join()
                ToolStripLabel2.Text = "扫描已终止！"
                'ToolStripLabel2.BackColor = Color.Red
                Timer1.Enabled = False
                PictureBox1.Image = My.Resources.Resource1.waiting
                PictureBox1.Visible = False
                PictureBox2.Image = My.Resources.Resource1.waiting
                PictureBox2.Visible = False
                PictureBox3.Image = My.Resources.Resource1.waiting
                PictureBox3.Visible = False
                PictureBox4.Image = My.Resources.Resource1.waiting
                PictureBox4.Visible = False
                PictureBox5.Image = My.Resources.Resource1.waiting
                PictureBox5.Visible = False
            End If
            ListBox1.Items.Clear()
            ListView3.Items.Clear()
        End If
        'CanImport = False
    End Sub

    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles Button4.Click
        If ListBox1.Items.Count = 0 Then Exit Sub
        PictureBox4.Visible = True
        PictureBox4.Image = My.Resources.Resource1.waiting
        For i = 0 To ListView3.Items.Count - 1
            If ListView3.Items.Item(i).Text = "" Then
                MsgBox("清先清除空番号")
                PictureBox4.Image = My.Resources.Resource1.waiting
                PictureBox4.Visible = False
                Exit Sub
            End If
        Next
        Dim myarr(0) As Integer
        Dim num As Integer
        num = 0
        myarr(0) = -1
        For i = 0 To ListView3.Items.Count - 2 Step 1
            For j = i + 1 To ListView3.Items.Count - 1 Step 1
                If ListView3.Items.Item(i).Text = ListView3.Items.Item(j).Text Then
                    '看看哪个文件大
                    '待修改
                    Application.DoEvents()
                    If My.Computer.FileSystem.GetFileInfo(ListBox1.Items.Item(i)).Length > My.Computer.FileSystem.GetFileInfo(ListBox1.Items.Item(j)).Length Then
                        myarr(num) = j
                        num = num + 1
                        ReDim Preserve myarr(num)
                        Debug.Print(ListBox1.Items.Item(i) & "大于" & ListBox1.Items.Item(j))
                        'Debug.Print(My.Computer.FileSystem.GetFileInfo(ListBox1.Items.Item(i)).Length)
                        'Debug.Print(My.Computer.FileSystem.GetFileInfo(ListBox1.Items.Item(j)).Length)
                    Else
                        myarr(num) = i
                        num = num + 1
                        ReDim Preserve myarr(num)
                        Debug.Print(ListBox1.Items.Item(j) & "大于" & ListBox1.Items.Item(i))
                    End If

                End If
            Next
        Next
        If myarr(0) <> -1 And UBound(myarr) > 0 Then
            ReDim Preserve myarr(num - 1)
        End If
        If myarr(0) <> -1 Then
            Array.Sort(myarr)
            For i = 0 To UBound(myarr) - 1 Step 1
                'Debug.Print(myarr(i))
                For j = i + 1 To UBound(myarr)
                    If myarr(j) = myarr(i) Then
                        myarr(i) = -1
                    End If
                Next
            Next
            For i = 0 To UBound(myarr)
                'Debug.Print(myarr(i))
            Next
            For i = UBound(myarr) To 0 Step -1
                If myarr(i) <> -1 Then
                    ListBox3.Items.Add(ListBox1.Items.Item(myarr(i)))
                    ListBox1.Items.RemoveAt(myarr(i))
                    ListView3.Items.RemoveAt(myarr(i))
                End If
            Next
        End If
        PictureBox4.Image = My.Resources.Resource1.complete
        CanImport(1) = True
    End Sub



    Private Sub ListView3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView3.SelectedIndexChanged

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then

            For i = 0 To ListView3.Items.Count - 1
                ListView3.Items(i).Selected = False
            Next


            For i = 0 To ListView3.Items.Count - 1
                If InStr(ListView3.Items.Item(i).Text, TextBox1.Text, CompareMethod.Text) > 0 Then
                    Debug.Print(ListView3.Items.Item(i).Text)
                    ListView3.Items(i).Selected = True
                    ListView3.Focus()
                End If
            Next
        End If
    End Sub

    Private Sub Button7_Click_1(sender As Object, e As EventArgs) Handles Button7.Click
        If ListBox1.Items.Count = 0 Then Exit Sub
        '判断是否在扫描和导入
        If myScanThread IsNot Nothing Then
            If myScanThread.IsAlive = True Then
                MsgBox("请等待本次扫描停止或取消扫描！")
                Exit Sub
            End If
        End If
        If MyImportThread IsNot Nothing Then
            If MyImportThread.IsAlive = True Then
                MsgBox("数据库正在导入！请等待")
                Exit Sub
            End If
        End If


        For Each s In CanImport
            If s = False Then
                MsgBox("请完成所有检查！")
                Exit Sub
            End If

        Next


        Timer2.Enabled = True
        PictureBox2.Image = My.Resources.Resource1.waiting
        PictureBox2.Visible = True
        'ToolStripStatusLabel4.Text = "导入中……"
        ToolStripLabel2.Text = "导入中"
        MyImportThread = New System.Threading.Thread(AddressOf weituo)
        MyImportThread.Start()
    End Sub

    Private Sub Button10_Click_1(sender As Object, e As EventArgs) Handles Button10.Click
        If ListBox1.Items.Count = 0 Then Exit Sub
        PictureBox3.Visible = True
        PictureBox3.Image = My.Resources.Resource1.waiting
        For i = ListView3.Items.Count - 1 To 0 Step -1
            If ListView3.Items.Item(i).Text = "" Then
                ListBox2.Items.Add(ListBox1.Items.Item(i))
                ListBox1.Items.RemoveAt(i)
                ListView3.Items.RemoveAt(i)
            End If
        Next
        PictureBox3.Image = My.Resources.Resource1.complete
        CanImport(0) = True
    End Sub

    Private Sub Button9_Click_1(sender As Object, e As EventArgs) Handles Button9.Click
        ListBox1.Items.Clear()
        ListView3.Items.Clear()

    End Sub

    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView1.SelectedIndexChanged

    End Sub

    Private Sub ListView3_Click(sender As Object, e As EventArgs) Handles ListView3.Click
        'Me.Text = ListView3.SelectedItems(0).Index
        'ListView2.Items(ListView3.SelectedItems(0).Index).Selected = True
        ListBox1.SelectedIndex = ListView3.SelectedItems(0).Index
        'ListView3.
    End Sub

    Private Sub ListView1_DoubleClick(sender As Object, e As EventArgs) Handles ListView1.DoubleClick

    End Sub

    Private Sub ListView1_DragDrop(sender As Object, e As DragEventArgs) Handles ListView1.DragDrop
        For Each s As String In e.Data.GetData(DataFormats.FileDrop) '循环枚举数据
            If System.IO.Directory.Exists(s) = True Then
                ListView1.Items.Add(s) '添加到表
                ListView1.Items.Item(ListView1.Items.Count - 1).ImageIndex = 0
            End If
        Next
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Me.Close()

    End Sub

    Private Sub 删除ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 删除ToolStripMenuItem.Click
        For i = ListView3.SelectedIndices.Count - 1 To 0 Step -1
            ListBox1.Items.RemoveAt(ListView3.SelectedItems.Item(i).Index)
            ListView3.Items.RemoveAt(ListView3.SelectedIndices(i))
            '    ToolStripStatusLabel3.Text = ListView3.SelectedItems.Item(i).Index
        Next
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        If ListBox1.Items.Count = 0 Then Exit Sub
        If CanImport(0) = False Or CanImport(1) = False Then
            Exit Sub
        End If
        PictureBox5.Visible = True
        PictureBox5.Image = My.Resources.Resource1.waiting
        Dim con As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dr As OleDbDataReader
        Dim myDataAdapter As New OleDbDataAdapter
        Dim sql As String
        Try
            con.ConnectionString = con_ConnectionString
            con.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try
        ReDim myChongfufanhaoIndex(0)
        myChongfufanhaoIndex(0) = -1
        Dim num As Integer
        num = 0
        For i = 0 To ListView3.Items.Count - 1
            cmd.Connection = con '初始化OLEDB命令的连接属性为con
            sql = "select ID from Main where fanhao='" & ListView3.Items.Item(i).Text & "'"
            cmd.CommandText = sql
            dr = cmd.ExecuteReader() '执行OLEDB命令以ExecuteReader()方式，并返回一个OLEDBReader，赋值给dr
            'Debug.Print(dr.)
            While (dr.Read)
                If dr.Item(0).ToString <> "" Then
                    'ListBox3.Items.Add(ListView3.Items.Item(i).Text)
                    myChongfufanhaoIndex(num) = i
                    num = num + 1
                    ReDim Preserve myChongfufanhaoIndex(num)
                    Exit While
                End If
            End While
            dr.Close()
        Next
        con.Close()
        If myChongfufanhaoIndex(0) <> -1 And UBound(myChongfufanhaoIndex) > 0 Then
            ReDim Preserve myChongfufanhaoIndex(num - 1)
        End If

        If myChongfufanhaoIndex(0) <> -1 Then

            daoruxiangqing.ShowDialog()
        Else
            '无重复
            CanImport(2) = True
        End If
        If CanImport(2) = True Then
            PictureBox5.Image = My.Resources.Resource1.complete
        Else
            PictureBox5.Image = My.Resources.Resource1.waiting
            PictureBox5.Visible = False
        End If

    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        For i = 0 To UBound(CanImport)
            Debug.Print(CanImport(i))
        Next

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub ListView1_DragEnter(sender As Object, e As DragEventArgs) Handles ListView1.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub TextBox2_Click(sender As Object, e As EventArgs) Handles TextBox2.Click
        If TextBox2.Text = "查询" Then
            TextBox2.Text = ""
            TextBox2.ForeColor = Color.Black

        End If
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        ListBox4.Items.Clear()
        If ComboBox3.Text = "日本步兵" Then
            For Each s In Split(BubingFanhao, ",")
                ListBox4.Items.Add(s)
            Next
        End If
    End Sub

    Private Sub TextBox2_LostFocus(sender As Object, e As EventArgs) Handles TextBox2.LostFocus
        If TextBox2.Text = "" Then
            TextBox2.Text = "查询"
            TextBox2.ForeColor = Color.Gray

        End If
    End Sub

    Private Sub Button14_Click_1(sender As Object, e As EventArgs) Handles Button14.Click
        ListBox4.Items.Clear()
        BubingFanhao = "MCBD,CWPBD,LAFBD,HEYDOUGA,SMD,BT,S2MBD,S2M,HEYZO,HEY,CCDV,LLDV,SSDV,MXX,NIP,SKY,SKYHD,MKD,MKBD,CWP,CWPD,RHJ,RED,PT,LAF,DSAM,DSAMD,DSAMBD,DRG,DRGBD,DRC,KP,SSKP,CYC,CWDV,KSC"
        If ComboBox3.Text = "日本步兵" Then
            For Each s In Split(BubingFanhao, ",")
                ListBox4.Items.Add(s)
            Next
        End If
    End Sub

    Private Sub 删除ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles 删除ToolStripMenuItem1.Click
        For i = ListBox4.SelectedIndices.Count - 1 To 0 Step -1
            ListBox4.Items.Remove(ListBox4.SelectedItems.Item(i))
        Next

        BubingFanhao = ""
        For i = 0 To ListBox4.Items.Count - 1 Step 1

            If i = ListBox4.Items.Count - 1 Then
                BubingFanhao = BubingFanhao & ListBox4.Items.Item(i)
            Else
                BubingFanhao = BubingFanhao & ListBox4.Items.Item(i) & ","
            End If

        Next
        ListBox4.Items.Clear()
        For Each s In Split(BubingFanhao, ",")
            ListBox4.Items.Add(s)
        Next
    End Sub

    Private Sub Button8_Click_1(sender As Object, e As EventArgs) Handles Button8.Click
        Dim str1
        str1 = UCase(InputBox("新的番号"))
        If str1 = "" Then Exit Sub
        For Each item In ListBox4.Items
            If item = str1 Then
                MsgBox("不允许重复！")
                Exit Sub
            End If
        Next
        BubingFanhao = ""
        For i = 0 To ListBox4.Items.Count - 1 Step 1
            BubingFanhao = BubingFanhao & ListBox4.Items.Item(i) & ","
        Next
        BubingFanhao = BubingFanhao & str1
        ListBox4.Items.Clear()
        For Each s In Split(BubingFanhao, ",")
            ListBox4.Items.Add(s)
        Next
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            For i = 0 To ListBox4.Items.Count - 1 Step 1
                If ListBox4.Items.Item(i) = UCase(TextBox2.Text) Then
                    ListBox4.SelectedIndex = i
                    Exit Sub
                End If
            Next
        End If
    End Sub
End Class