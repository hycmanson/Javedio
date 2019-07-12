Imports System.Net.NetworkInformation
Public Class MySetupForm
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TreeView1.SelectedNode = TreeView1.TopNode.Nodes(0)
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Button4.Enabled = False
    End Sub

    Private Sub TreeView1_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView1.AfterSelect
        If TreeView1.SelectedNode.Text = "小图界面字体和颜色" Or TreeView1.SelectedNode.Text = "大图界面字体和颜色" Then
            FlowLayoutPanel1.Controls.Clear()
            Dim myLabel1 As New Label
            FlowLayoutPanel1.Controls.Add(myLabel1)
            myLabel1.Text = "选择一个更改："
            myLabel1.AutoSize = True
            myLabel1.Margin = New Padding(2, 10, 0, 0)
            myLabel1.Font = New System.Drawing.Font("微软雅黑", 12)

            Dim myComboBox1 As New ComboBox
            FlowLayoutPanel1.Controls.Add(myComboBox1)
            If TreeView1.SelectedNode.Text = "小图界面字体和颜色" Then
                myComboBox1.Items.Add("番号")
                myComboBox1.Items.Add("名称")
                myComboBox1.Items.Add("日期")
                myComboBox1.Items.Add("其他")
            Else
                myComboBox1.Items.Add("番号")
                myComboBox1.Items.Add("名称")
                myComboBox1.Items.Add("发行日期")
                myComboBox1.Items.Add("长度")
                myComboBox1.Items.Add("导演")
                myComboBox1.Items.Add("制作商")
                myComboBox1.Items.Add("发行商")
                myComboBox1.Items.Add("类别")
                myComboBox1.Items.Add("演员")
                myComboBox1.Items.Add("标签")
                myComboBox1.Items.Add("系列")
            End If
            myComboBox1.Text = "番号"
            myComboBox1.Name = "myComboBox1"
            myComboBox1.Font = New System.Drawing.Font("微软雅黑", 10)
            myComboBox1.Margin = New Padding(6, 8, 0, 0)
            myComboBox1.DropDownStyle = ComboBoxStyle.DropDownList
            FlowLayoutPanel1.SetFlowBreak(myComboBox1, True)
            AddHandler myComboBox1.SelectedIndexChanged, AddressOf FlowLayoutPanel1_ComboBox1_SelectedIndexChanged

            Dim myLabel2 As New Label
            FlowLayoutPanel1.Controls.Add(myLabel2)
            myLabel2.Text = "字体："
            myLabel2.AutoSize = True
            myLabel2.Margin = New Padding(2, 10, 0, 0)
            myLabel2.Font = New System.Drawing.Font("微软雅黑", 12)

            Dim myComboBox2 As New ComboBox
            FlowLayoutPanel1.Controls.Add(myComboBox2)
            myComboBox2.IntegralHeight = False
            myComboBox2.MaxDropDownItems = 10
            myComboBox2.Width = 200
            myComboBox2.Margin = New Padding(2, 8, 0, 0)
            myComboBox2.Items.AddRange(FontFamily.Families.Select(Function(font) font.Name).ToArray())
            If TreeView1.SelectedNode.Text = "小图界面字体和颜色" Then
                myComboBox2.Text = SmallPicFanhao.myFontString(0)
            Else
                myComboBox2.Text = BigpicFanhao.myFontString(0)
            End If
            myComboBox2.DropDownStyle = ComboBoxStyle.DropDownList
            myComboBox2.Font = New System.Drawing.Font("微软雅黑", 10)
            FlowLayoutPanel1.SetFlowBreak(myComboBox2, True)
            AddHandler myComboBox2.SelectedIndexChanged, AddressOf FlowLayoutPanel1_ComboBox2_SelectedIndexChanged

            Dim myLabel3 As New Label
            FlowLayoutPanel1.Controls.Add(myLabel3)
            myLabel3.Text = "大小："
            myLabel3.AutoSize = True
            myLabel3.Margin = New Padding(2, 10, 0, 0)
            myLabel3.Font = New System.Drawing.Font("微软雅黑", 12)

            Dim myComboBox3 As New ComboBox
            FlowLayoutPanel1.Controls.Add(myComboBox3)
            myComboBox3.IntegralHeight = False
            myComboBox3.MaxDropDownItems = 10
            myComboBox3.Width = 200
            myComboBox3.Margin = New Padding(2, 8, 0, 0)
            For i = 10 To 25
                myComboBox3.Items.Add(i)
            Next
            If TreeView1.SelectedNode.Text = "小图界面字体和颜色" Then
                myComboBox3.Text = SmallPicFanhao.myFontString(1)
            Else
                myComboBox3.Text = BigpicFanhao.myFontString(1)
            End If
            myComboBox3.DropDownStyle = ComboBoxStyle.DropDownList
            myComboBox3.Font = New System.Drawing.Font("微软雅黑", 10)
            AddHandler myComboBox3.SelectedIndexChanged, AddressOf FlowLayoutPanel1_ComboBox3_SelectedIndexChanged


            Dim myLabel4 As New Label
            FlowLayoutPanel1.Controls.Add(myLabel4)
            myLabel4.Text = "前景色："
            myLabel4.AutoSize = True
            myLabel4.Margin = New Padding(2, 10, 0, 0)
            myLabel4.Font = New System.Drawing.Font("微软雅黑", 12)

            Dim PicLabel1 As New Label
            FlowLayoutPanel1.Controls.Add(PicLabel1)
            PicLabel1.Cursor = Cursors.Hand
            PicLabel1.Name = "PicLabel1"
            PicLabel1.Text = ""
            PicLabel1.Width = 20
            PicLabel1.Height = 20
            'PicLabel1.BackColor = Color.Black
            If TreeView1.SelectedNode.Text = "小图界面字体和颜色" Then
                PicLabel1.BackColor = StringToColor(SmallPicFanhao.myFontString(2))
            Else
                PicLabel1.BackColor = StringToColor(BigpicFanhao.myFontString(2))
            End If
            PicLabel1.Margin = New Padding(2, 10, 0, 0)
            AddHandler PicLabel1.Click, AddressOf FlowLayoutPanel1_PicLabel1_Click
            FlowLayoutPanel1.SetFlowBreak(PicLabel1, True)

            Dim myLabel5 As New Label
            FlowLayoutPanel1.Controls.Add(myLabel5)
            myLabel5.Text = "背景色："
            myLabel5.AutoSize = True
            myLabel5.Margin = New Padding(2, 10, 0, 0)
            myLabel5.Font = New System.Drawing.Font("微软雅黑", 12)

            Dim PicLabel2 As New Label
            FlowLayoutPanel1.Controls.Add(PicLabel2)
            PicLabel2.Cursor = Cursors.Hand
            PicLabel2.Name = "PicLabel2"
            PicLabel2.Text = ""
            PicLabel2.Width = 20
            PicLabel2.Height = 20
            If TreeView1.SelectedNode.Text = "小图界面字体和颜色" Then
                PicLabel2.BackColor = StringToColor(SmallPicFanhao.myFontString(3))
            Else
                PicLabel2.BackColor = StringToColor(BigpicFanhao.myFontString(3))
            End If
            PicLabel2.Margin = New Padding(2, 10, 0, 0)
            AddHandler PicLabel2.Click, AddressOf FlowLayoutPanel1_PicLabel2_Click
            FlowLayoutPanel1.SetFlowBreak(PicLabel2, True)

            Dim myCheckBox1 As New CheckBox
            Dim myCheckBox2 As New CheckBox
            myCheckBox1.AutoSize = True
            myCheckBox2.AutoSize = True
            If TreeView1.SelectedNode.Text = "小图界面字体和颜色" Then
                myCheckBox1.Checked = IIf(SmallPicFanhao.myFontString(4) = "True", True, False)
                myCheckBox2.Checked = IIf(SmallPicFanhao.myFontString(5) = "True", True, False)
            Else
                myCheckBox1.Checked = IIf(BigpicFanhao.myFontString(4) = "True", True, False)
                myCheckBox2.Checked = IIf(BigpicFanhao.myFontString(5) = "True", True, False)
            End If
            myCheckBox1.Margin = New Padding(8, 8, 0, 0)
            myCheckBox2.Margin = New Padding(8, 8, 0, 0)
            myCheckBox1.Text = "粗体"
            myCheckBox2.Text = "斜体"
            FlowLayoutPanel1.Controls.Add(myCheckBox1)
            FlowLayoutPanel1.Controls.Add(myCheckBox2)
            AddHandler myCheckBox1.CheckedChanged, AddressOf FlowLayoutPanel1_CheckBox1_CheckedChanged
            AddHandler myCheckBox2.CheckedChanged, AddressOf FlowLayoutPanel1_CheckBox2_CheckedChanged



        ElseIf TreeView1.SelectedNode.Text = "功能设置" Then
            FlowLayoutPanel1.Controls.Clear()
            Dim myCheckBox1 As New CheckBox
            myCheckBox1.AutoSize = True
            myCheckBox1.Checked = IsClickSmalPicShowBigPic
            myCheckBox1.Text = "点击主界面图片显示大图"
            'myCheckBox1.Enabled = False
            AddHandler myCheckBox1.CheckedChanged, AddressOf FlowLayoutPanel1_CheckBoxView_CheckedChanged
            FlowLayoutPanel1.Controls.Add(myCheckBox1)

            FlowLayoutPanel1.SetFlowBreak(myCheckBox1, True)


        ElseIf TreeView1.SelectedNode.Text = "外观设置" Then
            FlowLayoutPanel1.Controls.Clear()

            Dim myLabel1 As New Label
            myLabel1.Text = "颜色主题："
            myLabel1.AutoSize = True
            myLabel1.Margin = New Padding(2, 10, 0, 0)
            myLabel1.Font = New System.Drawing.Font("微软雅黑", 12)

            Dim myComboBox1 As New ComboBox
            myComboBox1.Items.Add("默认")
            myComboBox1.Items.Add("灰色")
            myComboBox1.Items.Add("黑色")


            myComboBox1.Text = ColorZhuti
            myComboBox1.Name = "myComboBox1"
            myComboBox1.Font = New System.Drawing.Font("微软雅黑", 10)
            myComboBox1.Margin = New Padding(6, 8, 0, 0)
            myComboBox1.DropDownStyle = ComboBoxStyle.DropDownList
            FlowLayoutPanel1.SetFlowBreak(myComboBox1, True)
            AddHandler myComboBox1.SelectedIndexChanged, AddressOf FlowLayoutPanel1_ColorComboBox_SelectedIndexChanged

            FlowLayoutPanel1.Controls.Add(myLabel1)
            FlowLayoutPanel1.Controls.Add(myComboBox1)

        ElseIf TreeView1.SelectedNode.Text = "显示设置" Then
            FlowLayoutPanel1.Controls.Clear()
            Dim myCheckBox1 As New CheckBox
            myCheckBox1.AutoSize = True
            myCheckBox1.Checked = IsCloseNow
            myCheckBox1.Text = "点击关闭后直接退出"
            'myCheckBox1.Enabled = False
            AddHandler myCheckBox1.CheckedChanged, AddressOf FlowLayoutPanel1_CheckBoxView_CheckedChanged

            Dim myCheckBox2 As New CheckBox
            myCheckBox2.AutoSize = True
            myCheckBox2.Checked = IsSmallPicAutoSize
            myCheckBox2.Text = "小图图片自适应"
            AddHandler myCheckBox2.CheckedChanged, AddressOf FlowLayoutPanel1_CheckBoxView_CheckedChanged

            Dim myLabel1 As New Label
            myLabel1.Text = "每页显示番号的数量："
            myLabel1.AutoSize = True
            myLabel1.Margin = New Padding(2, 8, 0, 0)

            Dim myTextBox1 As New TextBox
            myTextBox1.Text = FlowlayoutPanel_PageNum
            myTextBox1.Width = 50
            AddHandler myTextBox1.TextChanged, AddressOf View_myTextBox1_TextChanged



            Dim myCheckBox3 As New CheckBox
            myCheckBox3.AutoSize = True
            myCheckBox3.Checked = IsFanhaoShow
            myCheckBox3.Text = "主页面显示番号"
            AddHandler myCheckBox3.CheckedChanged, AddressOf FlowLayoutPanel1_CheckBoxView_CheckedChanged

            Dim myCheckBox4 As New CheckBox
            myCheckBox4.AutoSize = True
            myCheckBox4.Checked = IsMingchengShow
            myCheckBox4.Text = "主页面显示名称"
            AddHandler myCheckBox4.CheckedChanged, AddressOf FlowLayoutPanel1_CheckBoxView_CheckedChanged

            Dim myCheckBox5 As New CheckBox
            myCheckBox5.AutoSize = True
            myCheckBox5.Checked = IsRiqiShow
            myCheckBox5.Text = "主页面显示日期"
            AddHandler myCheckBox5.CheckedChanged, AddressOf FlowLayoutPanel1_CheckBoxView_CheckedChanged

            Dim myCheckBox6 As New CheckBox
            myCheckBox6.AutoSize = True
            myCheckBox6.Checked = IsShoucangShow
            myCheckBox6.Text = "主页面显示收藏"
            AddHandler myCheckBox6.CheckedChanged, AddressOf FlowLayoutPanel1_CheckBoxView_CheckedChanged

            FlowLayoutPanel1.Controls.Add(myCheckBox1)
            FlowLayoutPanel1.Controls.Add(myCheckBox2)
            FlowLayoutPanel1.Controls.Add(myCheckBox3)
            FlowLayoutPanel1.Controls.Add(myCheckBox4)
            FlowLayoutPanel1.Controls.Add(myCheckBox5)
            FlowLayoutPanel1.Controls.Add(myCheckBox6)
            FlowLayoutPanel1.Controls.Add(myLabel1)
            FlowLayoutPanel1.Controls.Add(myTextBox1)
            FlowLayoutPanel1.SetFlowBreak(myCheckBox1, True)
            FlowLayoutPanel1.SetFlowBreak(myCheckBox2, True)
            FlowLayoutPanel1.SetFlowBreak(myCheckBox3, True)
            FlowLayoutPanel1.SetFlowBreak(myCheckBox4, True)
            FlowLayoutPanel1.SetFlowBreak(myCheckBox5, True)
            FlowLayoutPanel1.SetFlowBreak(myCheckBox6, True)
        ElseIf TreeView1.SelectedNode.Text = "设置网站" Then
            FlowLayoutPanel1.Controls.Clear()
            Dim myLabel1 As New Label
            FlowLayoutPanel1.Controls.Add(myLabel1)
            myLabel1.Text = "设置网址（www.XX.com）："
            myLabel1.AutoSize = True
            myLabel1.Margin = New Padding(2, 10, 0, 0)
            myLabel1.Font = New System.Drawing.Font("微软雅黑", 12)
            FlowLayoutPanel1.SetFlowBreak(myLabel1, True)

            Dim myTextBox1 As New TextBox
            FlowLayoutPanel1.Controls.Add(myTextBox1)
            myTextBox1.Text = JavWebSite
            myTextBox1.Width = 150
            myTextBox1.Margin = New Padding(2, 10, 0, 0)
            'FlowLayoutPanel1.Controls.SetChildIndex(myTextBox1, 1)

            Dim myButton1 As New Button
            FlowLayoutPanel1.Controls.Add(myButton1)
            myButton1.Text = "Ping"
            myButton1.Margin = New Padding(2, 10, 0, 0)
            FlowLayoutPanel1.SetFlowBreak(myButton1, True)
            AddHandler myButton1.Click, AddressOf FlowLayoutPanel1_myButton1_Click

            Dim myLabel2 As New Label
            FlowLayoutPanel1.Controls.Add(myLabel2)
            myLabel2.Name = "myLabel2"
            myLabel2.AutoSize = True
            myLabel2.Margin = New Padding(2, 10, 0, 0)
            myLabel2.Font = New System.Drawing.Font("微软雅黑", 12)

        ElseIf TreeView1.SelectedNode.Text = "存储位置" Then
            FlowLayoutPanel1.Controls.Clear()
            Dim myButton As New Button
            FlowLayoutPanel1.Controls.Add(myButton)
            myButton.Text = "全部设置为根目录"
            myButton.AutoSize = True
            myButton.Margin = New Padding(2, 10, 0, 0)
            FlowLayoutPanel1.SetFlowBreak(myButton, True)
            AddHandler myButton.Click, AddressOf FlowLayoutPanel1_myButton_yijianshezhi_Click



            Dim myLabel1 As New Label
            FlowLayoutPanel1.Controls.Add(myLabel1)
            myLabel1.Text = "设置数据库保存位置："
            myLabel1.AutoSize = True
            myLabel1.Margin = New Padding(2, 10, 0, 0)
            myLabel1.Font = New System.Drawing.Font("微软雅黑", 12)
            FlowLayoutPanel1.SetFlowBreak(myLabel1, True)

            Dim myTextBox1 As New TextBox
            FlowLayoutPanel1.Controls.Add(myTextBox1)
            myTextBox1.Text = DataBaseSavePath
            myTextBox1.Width = 200
            myTextBox1.Margin = New Padding(2, 10, 0, 0)


            Dim myButton1 As New Button
            FlowLayoutPanel1.Controls.Add(myButton1)
            myButton1.Text = "浏览"
            myButton1.Margin = New Padding(2, 10, 0, 0)
            FlowLayoutPanel1.SetFlowBreak(myButton1, True)
            AddHandler myButton1.Click, AddressOf FlowLayoutPanel1_myButton1_liulan_Click

            Dim myLabel2 As New Label
            FlowLayoutPanel1.Controls.Add(myLabel2)
            myLabel2.Text = "设置小图保存位置："
            myLabel2.AutoSize = True
            myLabel2.Margin = New Padding(2, 10, 0, 0)
            myLabel2.Font = New System.Drawing.Font("微软雅黑", 12)
            FlowLayoutPanel1.SetFlowBreak(myLabel2, True)

            Dim myTextBox2 As New TextBox
            FlowLayoutPanel1.Controls.Add(myTextBox2)
            myTextBox2.Text = SmallPicSavePath
            myTextBox2.Width = 200
            myTextBox2.Margin = New Padding(2, 10, 0, 0)


            Dim myButton2 As New Button
            FlowLayoutPanel1.Controls.Add(myButton2)
            myButton2.Text = "浏览"
            myButton2.Margin = New Padding(2, 10, 0, 0)
            AddHandler myButton2.Click, AddressOf FlowLayoutPanel1_myButton2_liulan_Click

            Dim myLabel3 As New Label
            FlowLayoutPanel1.Controls.Add(myLabel3)
            myLabel3.Text = "设置大图保存位置："
            myLabel3.AutoSize = True
            myLabel3.Margin = New Padding(2, 10, 0, 0)
            myLabel3.Font = New System.Drawing.Font("微软雅黑", 12)
            FlowLayoutPanel1.SetFlowBreak(myLabel3, True)

            Dim myTextBox3 As New TextBox
            FlowLayoutPanel1.Controls.Add(myTextBox3)
            myTextBox3.Text = BigPicSavePath
            myTextBox3.Width = 200
            myTextBox3.Margin = New Padding(2, 10, 0, 0)


            Dim myButton3 As New Button
            FlowLayoutPanel1.Controls.Add(myButton3)
            myButton3.Text = "浏览"
            myButton3.Margin = New Padding(2, 10, 0, 0)
            AddHandler myButton3.Click, AddressOf FlowLayoutPanel1_myButton3_liulan_Click


            Dim myLabel4 As New Label
            FlowLayoutPanel1.Controls.Add(myLabel4)
            myLabel4.Text = "设置演员头像保存位置："
            myLabel4.AutoSize = True
            myLabel4.Margin = New Padding(2, 10, 0, 0)
            myLabel4.Font = New System.Drawing.Font("微软雅黑", 12)
            FlowLayoutPanel1.SetFlowBreak(myLabel4, True)

            Dim myTextBox4 As New TextBox
            FlowLayoutPanel1.Controls.Add(myTextBox4)
            myTextBox4.Text = ActressesPicSavePath
            myTextBox4.Width = 200
            myTextBox4.Margin = New Padding(2, 10, 0, 0)


            Dim myButton4 As New Button
            FlowLayoutPanel1.Controls.Add(myButton4)
            myButton4.Text = "浏览"
            myButton4.Margin = New Padding(2, 10, 0, 0)
            AddHandler myButton4.Click, AddressOf FlowLayoutPanel1_myButton4_liulan_Click

            Dim myLabel5 As New Label
            FlowLayoutPanel1.Controls.Add(myLabel5)
            myLabel5.Text = "设置额外图片保存位置："
            myLabel5.AutoSize = True
            myLabel5.Margin = New Padding(2, 10, 0, 0)
            myLabel5.Font = New System.Drawing.Font("微软雅黑", 12)
            FlowLayoutPanel1.SetFlowBreak(myLabel5, True)

            Dim myTextBox5 As New TextBox
            FlowLayoutPanel1.Controls.Add(myTextBox5)
            myTextBox5.Text = ExtraPicSavePath
            myTextBox5.Width = 200
            myTextBox5.Margin = New Padding(2, 10, 0, 0)


            Dim myButton5 As New Button
            FlowLayoutPanel1.Controls.Add(myButton5)
            myButton5.Text = "浏览"
            myButton5.Margin = New Padding(2, 10, 0, 0)
            AddHandler myButton5.Click, AddressOf FlowLayoutPanel1_myButton5_liulan_Click
        End If
        'FlowLayoutPanel1.Focus()
    End Sub

    Private Sub FlowLayoutPanel1_myButton_yijianshezhi_Click()
        Dim myTextBox As TextBox
        myTextBox = FlowLayoutPanel1.Controls(2)
        myTextBox.Text = Application.StartupPath & "\mdb"
        myTextBox = FlowLayoutPanel1.Controls(5)
        myTextBox.Text = Application.StartupPath & "\Pic\SmallPic"
        myTextBox = FlowLayoutPanel1.Controls(8)
        myTextBox.Text = Application.StartupPath & "\Pic\BigPic"
        myTextBox = FlowLayoutPanel1.Controls(11)
        myTextBox.Text = Application.StartupPath & "\Pic\Actresses"
        myTextBox = FlowLayoutPanel1.Controls(14)
        myTextBox.Text = Application.StartupPath & "\Pic\Extra"
        Button4.Enabled = True
    End Sub
    Private Sub FlowLayoutPanel1_myButton1_liulan_Click()
        If FolderBrowserDialog1.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
            Dim myTextBox As TextBox
            myTextBox = FlowLayoutPanel1.Controls(2)
            myTextBox.Text = FolderBrowserDialog1.SelectedPath
            Button4.Enabled = True
        End If
    End Sub
    Private Sub FlowLayoutPanel1_myButton2_liulan_Click()
        If FolderBrowserDialog1.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
            Dim myTextBox As TextBox
            myTextBox = FlowLayoutPanel1.Controls(5)
            myTextBox.Text = FolderBrowserDialog1.SelectedPath
            Button4.Enabled = True
        End If
    End Sub

    Private Sub FlowLayoutPanel1_myButton3_liulan_Click()
        If FolderBrowserDialog1.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
            Dim myTextBox As TextBox
            myTextBox = FlowLayoutPanel1.Controls(8)
            myTextBox.Text = FolderBrowserDialog1.SelectedPath
            Button4.Enabled = True
        End If
    End Sub

    Private Sub FlowLayoutPanel1_myButton4_liulan_Click()
        If FolderBrowserDialog1.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
            Dim myTextBox As TextBox
            myTextBox = FlowLayoutPanel1.Controls(11)
            myTextBox.Text = FolderBrowserDialog1.SelectedPath
            Button4.Enabled = True
        End If
    End Sub

    Private Sub FlowLayoutPanel1_myButton5_liulan_Click()
        If FolderBrowserDialog1.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
            Dim myTextBox As TextBox
            myTextBox = FlowLayoutPanel1.Controls(14)
            myTextBox.Text = FolderBrowserDialog1.SelectedPath
            Button4.Enabled = True
        End If
    End Sub

    Private Sub View_myTextBox1_TextChanged()
        Button4.Enabled = True
    End Sub

    Private Sub FlowLayoutPanel1_PicLabel1_Click()
        If ColorDialog1.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
            For i = 0 To FlowLayoutPanel1.Controls.Count - 1
                If FlowLayoutPanel1.Controls(i).Name = "PicLabel1" Then
                    FlowLayoutPanel1.Controls(i).BackColor = ColorDialog1.Color
                    Button4.Enabled = True
                    'MsgBox(FlowLayoutPanel1.Controls(i).BackColor.ToString)
                End If
            Next
        End If
    End Sub

    Private Sub FlowLayoutPanel1_PicLabel2_Click()
        If ColorDialog1.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
            For i = 0 To FlowLayoutPanel1.Controls.Count - 1
                If FlowLayoutPanel1.Controls(i).Name = "PicLabel2" Then
                    FlowLayoutPanel1.Controls(i).BackColor = ColorDialog1.Color
                    Button4.Enabled = True
                    'MsgBox(FlowLayoutPanel1.Controls(i).BackColor.ToString)
                End If
            Next
        End If
    End Sub

    Private Sub FlowLayoutPanel1_myButton1_Click()
        Try
            Dim WebUrl As String
            Dim returnPing
            Dim myPing As New System.Net.NetworkInformation.Ping
            WebUrl = FlowLayoutPanel1.Controls(1).Text '
            Dim PingYes As PingReply = New Ping().Send(WebUrl, 1000)
            If PingYes.Status = IPStatus.Success Then
                returnPing = myPing.Send(WebUrl, 1000).RoundtripTime
                FlowLayoutPanel1.Controls(3).Text = "网址的延时为： " & returnPing & " ms"
                Button4.Enabled = True
            Else
                FlowLayoutPanel1.Controls(3).Text = "无法Ping通"
            End If

        Catch ex As Exception
            FlowLayoutPanel1.Controls(3).Text = "无法Ping通"
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Button4.Enabled = True Then Call Button4_Click(sender, e)
        Me.Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        '存
        Dim myXml As New CSysXML("config.xml")
        Select Case (TreeView1.SelectedNode.Text)
            Case "小图界面字体和颜色"
                If FlowLayoutPanel1.Controls(1).Text = "番号" Then
                    SmallPicFanhao.myFontString(0) = FlowLayoutPanel1.Controls(3).Text
                    SmallPicFanhao.myFontString(1) = FlowLayoutPanel1.Controls(5).Text
                    SmallPicFanhao.myFontString(2) = ColorToString(FlowLayoutPanel1.Controls(7).BackColor)
                    SmallPicFanhao.myFontString(3) = ColorToString(FlowLayoutPanel1.Controls(9).BackColor)
                    Dim TempCheckBox1 As CheckBox = FlowLayoutPanel1.Controls(10)
                    Dim TempCheckBox2 As CheckBox = FlowLayoutPanel1.Controls(11)
                    SmallPicFanhao.myFontString(4) = TempCheckBox1.Checked
                    SmallPicFanhao.myFontString(5) = TempCheckBox2.Checked
                    myXml.SaveElement2("SmallPicForm", "Fanhao", SmallPicFanhao.myFontString)
                ElseIf FlowLayoutPanel1.Controls(1).Text = "名称" Then
                    SmallPicMingcheng.myFontString(0) = FlowLayoutPanel1.Controls(3).Text
                    SmallPicMingcheng.myFontString(1) = FlowLayoutPanel1.Controls(5).Text
                    SmallPicMingcheng.myFontString(2) = ColorToString(FlowLayoutPanel1.Controls(7).BackColor)
                    SmallPicMingcheng.myFontString(3) = ColorToString(FlowLayoutPanel1.Controls(9).BackColor)
                    Dim TempCheckBox1 As CheckBox = FlowLayoutPanel1.Controls(10)
                    Dim TempCheckBox2 As CheckBox = FlowLayoutPanel1.Controls(11)
                    SmallPicMingcheng.myFontString(4) = TempCheckBox1.Checked
                    SmallPicMingcheng.myFontString(5) = TempCheckBox2.Checked
                    myXml.SaveElement2("SmallPicForm", "Mingcheng", SmallPicMingcheng.myFontString)
                ElseIf FlowLayoutPanel1.Controls(1).Text = "日期" Then
                    SmallPicRiqi.myFontString(0) = FlowLayoutPanel1.Controls(3).Text
                    SmallPicRiqi.myFontString(1) = FlowLayoutPanel1.Controls(5).Text
                    SmallPicRiqi.myFontString(2) = ColorToString(FlowLayoutPanel1.Controls(7).BackColor)
                    SmallPicRiqi.myFontString(3) = ColorToString(FlowLayoutPanel1.Controls(9).BackColor)
                    Dim TempCheckBox1 As CheckBox = FlowLayoutPanel1.Controls(10)
                    Dim TempCheckBox2 As CheckBox = FlowLayoutPanel1.Controls(11)
                    SmallPicRiqi.myFontString(4) = TempCheckBox1.Checked
                    SmallPicRiqi.myFontString(5) = TempCheckBox2.Checked
                    myXml.SaveElement2("SmallPicForm", "Riqi", SmallPicRiqi.myFontString)
                ElseIf FlowLayoutPanel1.Controls(1).Text = "其他" Then
                    SmallPicQita.myFontString(0) = FlowLayoutPanel1.Controls(3).Text
                    SmallPicQita.myFontString(1) = FlowLayoutPanel1.Controls(5).Text
                    SmallPicQita.myFontString(2) = ColorToString(FlowLayoutPanel1.Controls(7).BackColor)
                    SmallPicQita.myFontString(3) = ColorToString(FlowLayoutPanel1.Controls(9).BackColor)
                    Dim TempCheckBox1 As CheckBox = FlowLayoutPanel1.Controls(10)
                    Dim TempCheckBox2 As CheckBox = FlowLayoutPanel1.Controls(11)
                    SmallPicQita.myFontString(4) = TempCheckBox1.Checked
                    SmallPicQita.myFontString(5) = TempCheckBox2.Checked
                    myXml.SaveElement2("SmallPicForm", "Qita", SmallPicQita.myFontString)
                End If
            Case "大图界面字体和颜色"
                If FlowLayoutPanel1.Controls(1).Text = "番号" Then
                    BigpicFanhao.myFontString(0) = FlowLayoutPanel1.Controls(3).Text
                    BigpicFanhao.myFontString(1) = FlowLayoutPanel1.Controls(5).Text
                    BigpicFanhao.myFontString(2) = ColorToString(FlowLayoutPanel1.Controls(7).BackColor)
                    BigpicFanhao.myFontString(3) = ColorToString(FlowLayoutPanel1.Controls(9).BackColor)
                    Dim TempCheckBox1 As CheckBox = FlowLayoutPanel1.Controls(10)
                    Dim TempCheckBox2 As CheckBox = FlowLayoutPanel1.Controls(11)
                    BigpicFanhao.myFontString(4) = TempCheckBox1.Checked
                    BigpicFanhao.myFontString(5) = TempCheckBox2.Checked
                    myXml.SaveElement2("BigPicForm", "Fanhao", BigpicFanhao.myFontString)
                ElseIf FlowLayoutPanel1.Controls(1).Text = "名称" Then
                    BigpicMingcheng.myFontString(0) = FlowLayoutPanel1.Controls(3).Text
                    BigpicMingcheng.myFontString(1) = FlowLayoutPanel1.Controls(5).Text
                    BigpicMingcheng.myFontString(2) = ColorToString(FlowLayoutPanel1.Controls(7).BackColor)
                    BigpicMingcheng.myFontString(3) = ColorToString(FlowLayoutPanel1.Controls(9).BackColor)
                    Dim TempCheckBox1 As CheckBox = FlowLayoutPanel1.Controls(10)
                    Dim TempCheckBox2 As CheckBox = FlowLayoutPanel1.Controls(11)
                    BigpicMingcheng.myFontString(4) = TempCheckBox1.Checked
                    BigpicMingcheng.myFontString(5) = TempCheckBox2.Checked
                    myXml.SaveElement2("BigPicForm", "Mingcheng", BigpicMingcheng.myFontString)
                ElseIf FlowLayoutPanel1.Controls(1).Text = "发行日期" Then
                    BigpicFaxingriqi.myFontString(0) = FlowLayoutPanel1.Controls(3).Text
                    BigpicFaxingriqi.myFontString(1) = FlowLayoutPanel1.Controls(5).Text
                    BigpicFaxingriqi.myFontString(2) = ColorToString(FlowLayoutPanel1.Controls(7).BackColor)
                    BigpicFaxingriqi.myFontString(3) = ColorToString(FlowLayoutPanel1.Controls(9).BackColor)
                    Dim TempCheckBox1 As CheckBox = FlowLayoutPanel1.Controls(10)
                    Dim TempCheckBox2 As CheckBox = FlowLayoutPanel1.Controls(11)
                    BigpicFaxingriqi.myFontString(4) = TempCheckBox1.Checked
                    BigpicFaxingriqi.myFontString(5) = TempCheckBox2.Checked
                    myXml.SaveElement2("BigPicForm", "Faxingriqi", BigpicFaxingriqi.myFontString)
                ElseIf FlowLayoutPanel1.Controls(1).Text = "长度" Then
                    BigpicChangdu.myFontString(0) = FlowLayoutPanel1.Controls(3).Text
                    BigpicChangdu.myFontString(1) = FlowLayoutPanel1.Controls(5).Text
                    BigpicChangdu.myFontString(2) = ColorToString(FlowLayoutPanel1.Controls(7).BackColor)
                    BigpicChangdu.myFontString(3) = ColorToString(FlowLayoutPanel1.Controls(9).BackColor)
                    Dim TempCheckBox1 As CheckBox = FlowLayoutPanel1.Controls(10)
                    Dim TempCheckBox2 As CheckBox = FlowLayoutPanel1.Controls(11)
                    BigpicChangdu.myFontString(4) = TempCheckBox1.Checked
                    BigpicChangdu.myFontString(5) = TempCheckBox2.Checked
                    myXml.SaveElement2("BigPicForm", "Changdu", BigpicChangdu.myFontString)
                ElseIf FlowLayoutPanel1.Controls(1).Text = "导演" Then
                    BigpicDaoyan.myFontString(0) = FlowLayoutPanel1.Controls(3).Text
                    BigpicDaoyan.myFontString(1) = FlowLayoutPanel1.Controls(5).Text
                    BigpicDaoyan.myFontString(2) = ColorToString(FlowLayoutPanel1.Controls(7).BackColor)
                    BigpicDaoyan.myFontString(3) = ColorToString(FlowLayoutPanel1.Controls(9).BackColor)
                    Dim TempCheckBox1 As CheckBox = FlowLayoutPanel1.Controls(10)
                    Dim TempCheckBox2 As CheckBox = FlowLayoutPanel1.Controls(11)
                    BigpicDaoyan.myFontString(4) = TempCheckBox1.Checked
                    BigpicDaoyan.myFontString(5) = TempCheckBox2.Checked
                    myXml.SaveElement2("BigPicForm", "Daoyan", BigpicDaoyan.myFontString)
                ElseIf FlowLayoutPanel1.Controls(1).Text = "制作商" Then
                    BigpicZhizuoshang.myFontString(0) = FlowLayoutPanel1.Controls(3).Text
                    BigpicZhizuoshang.myFontString(1) = FlowLayoutPanel1.Controls(5).Text
                    BigpicZhizuoshang.myFontString(2) = ColorToString(FlowLayoutPanel1.Controls(7).BackColor)
                    BigpicZhizuoshang.myFontString(3) = ColorToString(FlowLayoutPanel1.Controls(9).BackColor)
                    Dim TempCheckBox1 As CheckBox = FlowLayoutPanel1.Controls(10)
                    Dim TempCheckBox2 As CheckBox = FlowLayoutPanel1.Controls(11)
                    BigpicZhizuoshang.myFontString(4) = TempCheckBox1.Checked
                    BigpicZhizuoshang.myFontString(5) = TempCheckBox2.Checked
                    myXml.SaveElement2("BigPicForm", "Zhizuoshang", BigpicZhizuoshang.myFontString)
                ElseIf FlowLayoutPanel1.Controls(1).Text = "发行商" Then
                    BigpicFaxingshang.myFontString(0) = FlowLayoutPanel1.Controls(3).Text
                    BigpicFaxingshang.myFontString(1) = FlowLayoutPanel1.Controls(5).Text
                    BigpicFaxingshang.myFontString(2) = ColorToString(FlowLayoutPanel1.Controls(7).BackColor)
                    BigpicFaxingshang.myFontString(3) = ColorToString(FlowLayoutPanel1.Controls(9).BackColor)
                    Dim TempCheckBox1 As CheckBox = FlowLayoutPanel1.Controls(10)
                    Dim TempCheckBox2 As CheckBox = FlowLayoutPanel1.Controls(11)
                    BigpicFaxingshang.myFontString(4) = TempCheckBox1.Checked
                    BigpicFaxingshang.myFontString(5) = TempCheckBox2.Checked
                    myXml.SaveElement2("BigPicForm", "Zhizuoshang", BigpicFaxingshang.myFontString)
                ElseIf FlowLayoutPanel1.Controls(1).Text = "类别" Then
                    BigpicLeibie.myFontString(0) = FlowLayoutPanel1.Controls(3).Text
                    BigpicLeibie.myFontString(1) = FlowLayoutPanel1.Controls(5).Text
                    BigpicLeibie.myFontString(2) = ColorToString(FlowLayoutPanel1.Controls(7).BackColor)
                    BigpicLeibie.myFontString(3) = ColorToString(FlowLayoutPanel1.Controls(9).BackColor)
                    Dim TempCheckBox1 As CheckBox = FlowLayoutPanel1.Controls(10)
                    Dim TempCheckBox2 As CheckBox = FlowLayoutPanel1.Controls(11)
                    BigpicLeibie.myFontString(4) = TempCheckBox1.Checked
                    BigpicLeibie.myFontString(5) = TempCheckBox2.Checked
                    myXml.SaveElement2("BigPicForm", "Leibie", BigpicLeibie.myFontString)
                ElseIf FlowLayoutPanel1.Controls(1).Text = "演员" Then
                    BigpicYanyuan.myFontString(0) = FlowLayoutPanel1.Controls(3).Text
                    BigpicYanyuan.myFontString(1) = FlowLayoutPanel1.Controls(5).Text
                    BigpicYanyuan.myFontString(2) = ColorToString(FlowLayoutPanel1.Controls(7).BackColor)
                    BigpicYanyuan.myFontString(3) = ColorToString(FlowLayoutPanel1.Controls(9).BackColor)
                    Dim TempCheckBox1 As CheckBox = FlowLayoutPanel1.Controls(10)
                    Dim TempCheckBox2 As CheckBox = FlowLayoutPanel1.Controls(11)
                    BigpicYanyuan.myFontString(4) = TempCheckBox1.Checked
                    BigpicYanyuan.myFontString(5) = TempCheckBox2.Checked
                    myXml.SaveElement2("BigPicForm", "Yanyuan", BigpicYanyuan.myFontString)
                ElseIf FlowLayoutPanel1.Controls(1).Text = "标签" Then
                    BigpicBiaoqian.myFontString(0) = FlowLayoutPanel1.Controls(3).Text
                    BigpicBiaoqian.myFontString(1) = FlowLayoutPanel1.Controls(5).Text
                    BigpicBiaoqian.myFontString(2) = ColorToString(FlowLayoutPanel1.Controls(7).BackColor)
                    BigpicBiaoqian.myFontString(3) = ColorToString(FlowLayoutPanel1.Controls(9).BackColor)
                    Dim TempCheckBox1 As CheckBox = FlowLayoutPanel1.Controls(10)
                    Dim TempCheckBox2 As CheckBox = FlowLayoutPanel1.Controls(11)
                    BigpicBiaoqian.myFontString(4) = TempCheckBox1.Checked
                    BigpicBiaoqian.myFontString(5) = TempCheckBox2.Checked
                    myXml.SaveElement2("BigPicForm", "Biaoqian", BigpicBiaoqian.myFontString)
                ElseIf FlowLayoutPanel1.Controls(1).Text = "系列" Then
                    BigpicXilie.myFontString(0) = FlowLayoutPanel1.Controls(3).Text
                    BigpicXilie.myFontString(1) = FlowLayoutPanel1.Controls(5).Text
                    BigpicXilie.myFontString(2) = ColorToString(FlowLayoutPanel1.Controls(7).BackColor)
                    BigpicXilie.myFontString(3) = ColorToString(FlowLayoutPanel1.Controls(9).BackColor)
                    Dim TempCheckBox1 As CheckBox = FlowLayoutPanel1.Controls(10)
                    Dim TempCheckBox2 As CheckBox = FlowLayoutPanel1.Controls(11)
                    BigpicXilie.myFontString(4) = TempCheckBox1.Checked
                    BigpicXilie.myFontString(5) = TempCheckBox2.Checked
                    myXml.SaveElement2("BigPicForm", "Xilie", BigpicXilie.myFontString)
                End If

            Case "功能设置"
                Dim TempCheckBox1 As CheckBox = FlowLayoutPanel1.Controls(0)
                IsClickSmalPicShowBigPic = TempCheckBox1.Checked
                myXml.SaveElement("GongNeng", "IsClickSmalPicShowBigPic", IsClickSmalPicShowBigPic)


            Case "外观设置"
                Dim myComboBox As ComboBox = FlowLayoutPanel1.Controls(1)
                ColorZhuti = myComboBox.Text
                myXml.SaveElement("SoftView", "ColorZhuti", ColorZhuti)

                '设置主题
                Select Case ColorZhuti
                    Case "默认"
                        Form1.FlowLayoutPanel1.BackColor = Color.FromArgb(255, 240, 240, 240)
                        Form1.MenuStrip1.BackColor = Color.FromArgb(255, 240, 240, 240)
                        Form1.ToolStrip1.BackColor = Color.FromArgb(255, 240, 240, 240)
                        Form1.ToolStrip3.BackColor = Color.SkyBlue
                        Form1.ToolStripComboBox1.BackColor = Color.White
                        Form1.ToolStripTextBox1.BackColor = Color.White
                        Form1.ToolStripTextBox2.BackColor = Color.White

                        Form1.ToolStripLabel4.ForeColor = Color.Black
                        Form1.ToolStripTextBox1.ForeColor = Color.Black

                        Form1.ToolStripComboBox1.ForeColor = Color.Black
                        Form1.ToolStripTextBox2.ForeColor = Color.Black

                        Form1.ToolStripButton2.BackColor = Color.Transparent
                        Form1.ToolStripButton4.BackColor = Color.Transparent
                        Form1.ToolStripButton5.BackColor = Color.Transparent
                        Form1.ToolStripButton13.BackColor = Color.Transparent

                        Form1.文件ToolStripMenuItem.ForeColor = Color.Black
                        Form1.编辑ToolStripMenuItem.ForeColor = Color.Black
                        Form1.下载ToolStripMenuItem.ForeColor = Color.Black
                        Form1.关于ToolStripMenuItem.ForeColor = Color.Black
                        Form1.工具ToolStripMenuItem.ForeColor = Color.Black

                        Dim mylayout As FlowLayoutPanel
                        Dim mytextbox As TextBox
                        For Each item In Form1.FlowLayoutPanel1.Controls
                            mylayout = item
                            mylayout.BackColor = Color.White
                            For Each item1 In mylayout.Controls
                                If item1.GetType.ToString = "System.Windows.Forms.TextBox" Then
                                    mytextbox = item1
                                    mytextbox.BackColor = Color.White
                                    mytextbox.ForeColor = Color.Black
                                End If
                            Next
                        Next

                        Dim mytoolstripbtn As ToolStripButton
                        For Each item In Form1.ToolStrip3.Items
                            If item.GetType.ToString = "System.Windows.Forms.ToolStripButton" Then
                                mytoolstripbtn = item
                                mytoolstripbtn.ForeColor = Color.Black
                            End If
                        Next




                    Case "灰色"
                        Form1.FlowLayoutPanel1.BackColor = Color.DimGray
                        Form1.MenuStrip1.BackColor = Color.DimGray
                        Form1.ToolStrip1.BackColor = Color.DimGray
                        Form1.ToolStrip3.BackColor = Color.SlateGray
                        Form1.ToolStripComboBox1.BackColor = Color.DimGray
                        Form1.ToolStripTextBox1.BackColor = Color.DimGray
                        Form1.ToolStripTextBox2.BackColor = Color.DimGray

                        Form1.ToolStripLabel4.ForeColor = Color.White
                        Form1.ToolStripTextBox1.ForeColor = Color.White

                        Form1.ToolStripComboBox1.ForeColor = Color.White
                        Form1.ToolStripTextBox2.ForeColor = Color.White

                        Form1.ToolStripButton2.BackColor = Color.White
                        Form1.ToolStripButton4.BackColor = Color.Transparent
                        Form1.ToolStripButton5.BackColor = Color.Transparent
                        Form1.ToolStripButton13.BackColor = Color.Transparent

                        Form1.文件ToolStripMenuItem.ForeColor = Color.White
                        Form1.编辑ToolStripMenuItem.ForeColor = Color.White
                        Form1.下载ToolStripMenuItem.ForeColor = Color.White
                        Form1.关于ToolStripMenuItem.ForeColor = Color.White
                        Form1.工具ToolStripMenuItem.ForeColor = Color.White

                        Dim mylayout As FlowLayoutPanel
                        Dim mytextbox As TextBox
                        For Each item In Form1.FlowLayoutPanel1.Controls
                            mylayout = item
                            mylayout.BackColor = Color.FromArgb(64, 64, 64)
                            For Each item1 In mylayout.Controls
                                If item1.GetType.ToString = "System.Windows.Forms.TextBox" Then
                                    mytextbox = item1
                                    mytextbox.BackColor = Color.FromArgb(64, 64, 64)
                                    mytextbox.ForeColor = Color.White
                                End If
                            Next
                        Next

                        Dim mytoolstripbtn As ToolStripButton
                        For Each item In Form1.ToolStrip3.Items
                            Debug.Print(item.GetType.ToString)
                            If item.GetType.ToString = "System.Windows.Forms.ToolStripButton" Then
                                mytoolstripbtn = item
                                mytoolstripbtn.ForeColor = Color.White
                            End If
                        Next


                    Case "黑色"

                        Form1.FlowLayoutPanel1.BackColor = Color.Black
                        Form1.MenuStrip1.BackColor = Color.Black
                        Form1.ToolStrip1.BackColor = Color.Black
                        Form1.ToolStrip3.BackColor = Color.Black
                        Form1.ToolStripComboBox1.BackColor = Color.Black
                        Form1.ToolStripTextBox1.BackColor = Color.Black
                        Form1.ToolStripTextBox2.BackColor = Color.Black

                        Form1.ToolStripLabel4.ForeColor = Color.White
                        Form1.ToolStripTextBox1.ForeColor = Color.White

                        Form1.ToolStripComboBox1.ForeColor = Color.White
                        Form1.ToolStripTextBox2.ForeColor = Color.White

                        Form1.ToolStripButton2.BackColor = Color.White
                        Form1.ToolStripButton4.BackColor = Color.White
                        Form1.ToolStripButton5.BackColor = Color.White
                        Form1.ToolStripButton13.BackColor = Color.White

                        Form1.文件ToolStripMenuItem.ForeColor = Color.White
                        Form1.编辑ToolStripMenuItem.ForeColor = Color.White
                        Form1.下载ToolStripMenuItem.ForeColor = Color.White
                        Form1.关于ToolStripMenuItem.ForeColor = Color.White
                        Form1.工具ToolStripMenuItem.ForeColor = Color.White

                        Dim mylayout As FlowLayoutPanel
                        Dim mytextbox As TextBox
                        For Each item In Form1.FlowLayoutPanel1.Controls
                            mylayout = item
                            mylayout.BackColor = Color.Black
                            For Each item1 In mylayout.Controls
                                If item1.GetType.ToString = "System.Windows.Forms.TextBox" Then
                                    mytextbox = item1
                                    mytextbox.BackColor = Color.Black
                                    mytextbox.ForeColor = Color.White
                                End If
                            Next
                        Next

                        Dim mytoolstripbtn As ToolStripButton
                        For Each item In Form1.ToolStrip3.Items
                            Debug.Print(item.GetType.ToString)
                            If item.GetType.ToString = "System.Windows.Forms.ToolStripButton" Then
                                mytoolstripbtn = item
                                mytoolstripbtn.ForeColor = Color.White
                            End If
                        Next


                End Select

            Case "显示设置"
                Dim TempCheckBox1 As CheckBox = FlowLayoutPanel1.Controls(0)
                IsCloseNow = TempCheckBox1.Checked
                myXml.SaveElement("SoftView", "IsCloseNow", IsCloseNow)
                Dim TempCheckBox2 As CheckBox = FlowLayoutPanel1.Controls(1)
                IsSmallPicAutoSize = TempCheckBox2.Checked
                myXml.SaveElement("SoftView", "IsSmallPicAutoSize", IsSmallPicAutoSize)

                Dim TempCheckBox3 As CheckBox = FlowLayoutPanel1.Controls(2)
                IsFanhaoShow = TempCheckBox3.Checked
                myXml.SaveElement("SoftView", "IsFanhaoShow", IsFanhaoShow)

                Dim TempCheckBox4 As CheckBox = FlowLayoutPanel1.Controls(3)
                IsMingchengShow = TempCheckBox4.Checked
                myXml.SaveElement("SoftView", "IsMingchengShow", IsMingchengShow)

                Dim TempCheckBox5 As CheckBox = FlowLayoutPanel1.Controls(4)
                IsRiqiShow = TempCheckBox5.Checked
                myXml.SaveElement("SoftView", "IsRiqiShow", IsRiqiShow)

                Dim TempCheckBox6 As CheckBox = FlowLayoutPanel1.Controls(5)
                IsShoucangShow = TempCheckBox6.Checked
                myXml.SaveElement("SoftView", "IsShoucangShow", IsShoucangShow)


                Dim TempTextBox1 As TextBox = FlowLayoutPanel1.Controls(7)
                If IsNumeric(TempTextBox1.Text) Then
                    If Int(TempTextBox1.Text) >= 20 And Int(TempTextBox1.Text) <= 200 Then
                        FlowlayoutPanel_PageNum = TempTextBox1.Text
                        myXml.SaveElement("SoftView", "FlowlayoutPanel_PageNum", FlowlayoutPanel_PageNum)
                    Else
                        MsgBox("限定在20-200之间！")
                    End If

                Else
                    MsgBox("请输入数字！")
                End If



            Case "设置网站"
                JavWebSite = FlowLayoutPanel1.Controls(1).Text
                myXml.SaveElement("myWebSite", "JavWebSite", JavWebSite)
            Case "存储位置"
                DataBaseSavePath = FlowLayoutPanel1.Controls(2).Text
                myXml.SaveElement("SaveSetup", "DataBaseSavePath", DataBaseSavePath)
                SmallPicSavePath = FlowLayoutPanel1.Controls(5).Text
                myXml.SaveElement("SaveSetup", "SmallPicSavePath", SmallPicSavePath)
                BigPicSavePath = FlowLayoutPanel1.Controls(8).Text
                myXml.SaveElement("SaveSetup", "BigPicSavePath", BigPicSavePath)
                ActressesPicSavePath = FlowLayoutPanel1.Controls(11).Text
                myXml.SaveElement("SaveSetup", "ActressesPicSavePath", ActressesPicSavePath)
                ExtraPicSavePath = FlowLayoutPanel1.Controls(14).Text
                myXml.SaveElement("SaveSetup", "ExtraPicSavePath", ExtraPicSavePath)
                con_ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DataBaseSavePath & "\javbus.mdb"
                con_ConnectionString_read = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DataBaseSavePath & "\javbus.mdb;Mode=Read"

        End Select




        Button4.Enabled = False
    End Sub

    Sub FlowLayoutPanel1_ColorComboBox_SelectedIndexChanged()
        Button4.Enabled = True
    End Sub

    Public Sub FlowLayoutPanel1_ComboBox1_SelectedIndexChanged()

        '读
        Select Case FlowLayoutPanel1.Controls(1).Text
            Case "番号"
                If TreeView1.SelectedNode.Text = "小图界面字体和颜色" Then
                    FlowLayoutPanel1.Controls(3).Text = SmallPicFanhao.myFontString(0)
                    FlowLayoutPanel1.Controls(5).Text = SmallPicFanhao.myFontString(1)
                    FlowLayoutPanel1.Controls(7).BackColor = StringToColor(SmallPicFanhao.myFontString(2))
                    FlowLayoutPanel1.Controls(9).BackColor = StringToColor(SmallPicFanhao.myFontString(3))
                    Dim TempCheckBox1 As CheckBox = FlowLayoutPanel1.Controls(10)
                    Dim TempCheckBox2 As CheckBox = FlowLayoutPanel1.Controls(11)
                    TempCheckBox1.Checked = IIf(SmallPicFanhao.myFontString(4) = "True", True, False)
                    TempCheckBox2.Checked = IIf(SmallPicFanhao.myFontString(5) = "True", True, False)
                Else
                    FlowLayoutPanel1.Controls(3).Text = BigpicFanhao.myFontString(0)
                    FlowLayoutPanel1.Controls(5).Text = BigpicFanhao.myFontString(1)
                    FlowLayoutPanel1.Controls(7).BackColor = StringToColor(BigpicFanhao.myFontString(2))
                    FlowLayoutPanel1.Controls(9).BackColor = StringToColor(BigpicFanhao.myFontString(3))
                    Dim TempCheckBox1 As CheckBox = FlowLayoutPanel1.Controls(10)
                    Dim TempCheckBox2 As CheckBox = FlowLayoutPanel1.Controls(11)
                    TempCheckBox1.Checked = IIf(BigpicFanhao.myFontString(4) = "True", True, False)
                    TempCheckBox2.Checked = IIf(BigpicFanhao.myFontString(5) = "True", True, False)
                End If
            Case "名称"
                If TreeView1.SelectedNode.Text = "小图界面字体和颜色" Then
                    FlowLayoutPanel1.Controls(3).Text = SmallPicMingcheng.myFontString(0)
                    FlowLayoutPanel1.Controls(5).Text = SmallPicMingcheng.myFontString(1)
                    FlowLayoutPanel1.Controls(7).BackColor = StringToColor(SmallPicMingcheng.myFontString(2))
                    FlowLayoutPanel1.Controls(9).BackColor = StringToColor(SmallPicMingcheng.myFontString(3))
                    Dim TempCheckBox1 As CheckBox = FlowLayoutPanel1.Controls(10)
                    Dim TempCheckBox2 As CheckBox = FlowLayoutPanel1.Controls(11)
                    TempCheckBox1.Checked = IIf(SmallPicMingcheng.myFontString(4) = "True", True, False)
                    TempCheckBox2.Checked = IIf(SmallPicMingcheng.myFontString(5) = "True", True, False)
                Else
                    FlowLayoutPanel1.Controls(3).Text = BigpicMingcheng.myFontString(0)
                    FlowLayoutPanel1.Controls(5).Text = BigpicMingcheng.myFontString(1)
                    FlowLayoutPanel1.Controls(7).BackColor = StringToColor(BigpicMingcheng.myFontString(2))
                    FlowLayoutPanel1.Controls(9).BackColor = StringToColor(BigpicMingcheng.myFontString(3))
                    Dim TempCheckBox1 As CheckBox = FlowLayoutPanel1.Controls(10)
                    Dim TempCheckBox2 As CheckBox = FlowLayoutPanel1.Controls(11)
                    TempCheckBox1.Checked = IIf(BigpicMingcheng.myFontString(4) = "True", True, False)
                    TempCheckBox2.Checked = IIf(BigpicMingcheng.myFontString(5) = "True", True, False)
                End If
            Case "日期"
                If TreeView1.SelectedNode.Text = "小图界面字体和颜色" Then
                    FlowLayoutPanel1.Controls(3).Text = SmallPicRiqi.myFontString(0)
                    FlowLayoutPanel1.Controls(5).Text = SmallPicRiqi.myFontString(1)
                    FlowLayoutPanel1.Controls(7).BackColor = StringToColor(SmallPicRiqi.myFontString(2))
                    FlowLayoutPanel1.Controls(9).BackColor = StringToColor(SmallPicRiqi.myFontString(3))
                    Dim TempCheckBox1 As CheckBox = FlowLayoutPanel1.Controls(10)
                    Dim TempCheckBox2 As CheckBox = FlowLayoutPanel1.Controls(11)
                    TempCheckBox1.Checked = IIf(SmallPicRiqi.myFontString(4) = "True", True, False)
                    TempCheckBox2.Checked = IIf(SmallPicRiqi.myFontString(5) = "True", True, False)
                End If
            Case "其他"
                If TreeView1.SelectedNode.Text = "小图界面字体和颜色" Then
                    FlowLayoutPanel1.Controls(3).Text = SmallPicQita.myFontString(0)
                    FlowLayoutPanel1.Controls(5).Text = SmallPicQita.myFontString(1)
                    FlowLayoutPanel1.Controls(7).BackColor = StringToColor(SmallPicQita.myFontString(2))
                    FlowLayoutPanel1.Controls(9).BackColor = StringToColor(SmallPicQita.myFontString(3))
                    Dim TempCheckBox1 As CheckBox = FlowLayoutPanel1.Controls(10)
                    Dim TempCheckBox2 As CheckBox = FlowLayoutPanel1.Controls(11)
                    TempCheckBox1.Checked = IIf(SmallPicQita.myFontString(4) = "True", True, False)
                    TempCheckBox2.Checked = IIf(SmallPicQita.myFontString(5) = "True", True, False)
                End If
            Case "发行日期"
                FlowLayoutPanel1.Controls(3).Text = BigpicFaxingriqi.myFontString(0)
                FlowLayoutPanel1.Controls(5).Text = BigpicFaxingriqi.myFontString(1)
                FlowLayoutPanel1.Controls(7).BackColor = StringToColor(BigpicFaxingriqi.myFontString(2))
                FlowLayoutPanel1.Controls(9).BackColor = StringToColor(BigpicFaxingriqi.myFontString(3))
                Dim TempCheckBox1 As CheckBox = FlowLayoutPanel1.Controls(10)
                Dim TempCheckBox2 As CheckBox = FlowLayoutPanel1.Controls(11)
                TempCheckBox1.Checked = IIf(BigpicFaxingriqi.myFontString(4) = "True", True, False)
                TempCheckBox2.Checked = IIf(BigpicFaxingriqi.myFontString(5) = "True", True, False)
            Case "长度"
                FlowLayoutPanel1.Controls(3).Text = BigpicChangdu.myFontString(0)
                FlowLayoutPanel1.Controls(5).Text = BigpicChangdu.myFontString(1)
                FlowLayoutPanel1.Controls(7).BackColor = StringToColor(BigpicChangdu.myFontString(2))
                FlowLayoutPanel1.Controls(9).BackColor = StringToColor(BigpicChangdu.myFontString(3))
                Dim TempCheckBox1 As CheckBox = FlowLayoutPanel1.Controls(10)
                Dim TempCheckBox2 As CheckBox = FlowLayoutPanel1.Controls(11)
                TempCheckBox1.Checked = IIf(BigpicChangdu.myFontString(4) = "True", True, False)
                TempCheckBox2.Checked = IIf(BigpicChangdu.myFontString(5) = "True", True, False)
            Case "导演"
                FlowLayoutPanel1.Controls(3).Text = BigpicDaoyan.myFontString(0)
                FlowLayoutPanel1.Controls(5).Text = BigpicDaoyan.myFontString(1)
                FlowLayoutPanel1.Controls(7).BackColor = StringToColor(BigpicDaoyan.myFontString(2))
                FlowLayoutPanel1.Controls(9).BackColor = StringToColor(BigpicDaoyan.myFontString(3))
                Dim TempCheckBox1 As CheckBox = FlowLayoutPanel1.Controls(10)
                Dim TempCheckBox2 As CheckBox = FlowLayoutPanel1.Controls(11)
                TempCheckBox1.Checked = IIf(BigpicDaoyan.myFontString(4) = "True", True, False)
                TempCheckBox2.Checked = IIf(BigpicDaoyan.myFontString(5) = "True", True, False)
            Case "制作商"
                FlowLayoutPanel1.Controls(3).Text = BigpicZhizuoshang.myFontString(0)
                FlowLayoutPanel1.Controls(5).Text = BigpicZhizuoshang.myFontString(1)
                FlowLayoutPanel1.Controls(7).BackColor = StringToColor(BigpicZhizuoshang.myFontString(2))
                FlowLayoutPanel1.Controls(9).BackColor = StringToColor(BigpicZhizuoshang.myFontString(3))
                Dim TempCheckBox1 As CheckBox = FlowLayoutPanel1.Controls(10)
                Dim TempCheckBox2 As CheckBox = FlowLayoutPanel1.Controls(11)
                TempCheckBox1.Checked = IIf(BigpicZhizuoshang.myFontString(4) = "True", True, False)
                TempCheckBox2.Checked = IIf(BigpicZhizuoshang.myFontString(5) = "True", True, False)
            Case "发行商"
                FlowLayoutPanel1.Controls(3).Text = BigpicFaxingshang.myFontString(0)
                FlowLayoutPanel1.Controls(5).Text = BigpicFaxingshang.myFontString(1)
                FlowLayoutPanel1.Controls(7).BackColor = StringToColor(BigpicFaxingshang.myFontString(2))
                FlowLayoutPanel1.Controls(9).BackColor = StringToColor(BigpicFaxingshang.myFontString(3))
                Dim TempCheckBox1 As CheckBox = FlowLayoutPanel1.Controls(10)
                Dim TempCheckBox2 As CheckBox = FlowLayoutPanel1.Controls(11)
                TempCheckBox1.Checked = IIf(BigpicFaxingshang.myFontString(4) = "True", True, False)
                TempCheckBox2.Checked = IIf(BigpicFaxingshang.myFontString(5) = "True", True, False)
            Case "类别"
                FlowLayoutPanel1.Controls(3).Text = BigpicLeibie.myFontString(0)
                FlowLayoutPanel1.Controls(5).Text = BigpicLeibie.myFontString(1)
                FlowLayoutPanel1.Controls(7).BackColor = StringToColor(BigpicLeibie.myFontString(2))
                FlowLayoutPanel1.Controls(9).BackColor = StringToColor(BigpicLeibie.myFontString(3))
                Dim TempCheckBox1 As CheckBox = FlowLayoutPanel1.Controls(10)
                Dim TempCheckBox2 As CheckBox = FlowLayoutPanel1.Controls(11)
                TempCheckBox1.Checked = IIf(BigpicLeibie.myFontString(4) = "True", True, False)
                TempCheckBox2.Checked = IIf(BigpicLeibie.myFontString(5) = "True", True, False)
            Case "演员"
                FlowLayoutPanel1.Controls(3).Text = BigpicYanyuan.myFontString(0)
                FlowLayoutPanel1.Controls(5).Text = BigpicYanyuan.myFontString(1)
                FlowLayoutPanel1.Controls(7).BackColor = StringToColor(BigpicYanyuan.myFontString(2))
                FlowLayoutPanel1.Controls(9).BackColor = StringToColor(BigpicYanyuan.myFontString(3))
                Dim TempCheckBox1 As CheckBox = FlowLayoutPanel1.Controls(10)
                Dim TempCheckBox2 As CheckBox = FlowLayoutPanel1.Controls(11)
                TempCheckBox1.Checked = IIf(BigpicYanyuan.myFontString(4) = "True", True, False)
                TempCheckBox2.Checked = IIf(BigpicYanyuan.myFontString(5) = "True", True, False)
            Case "标签"
                FlowLayoutPanel1.Controls(3).Text = BigpicBiaoqian.myFontString(0)
                FlowLayoutPanel1.Controls(5).Text = BigpicBiaoqian.myFontString(1)
                FlowLayoutPanel1.Controls(7).BackColor = StringToColor(BigpicBiaoqian.myFontString(2))
                FlowLayoutPanel1.Controls(9).BackColor = StringToColor(BigpicBiaoqian.myFontString(3))
                Dim TempCheckBox1 As CheckBox = FlowLayoutPanel1.Controls(10)
                Dim TempCheckBox2 As CheckBox = FlowLayoutPanel1.Controls(11)
                TempCheckBox1.Checked = IIf(BigpicBiaoqian.myFontString(4) = "True", True, False)
                TempCheckBox2.Checked = IIf(BigpicBiaoqian.myFontString(5) = "True", True, False)
            Case "系列"
                FlowLayoutPanel1.Controls(3).Text = BigpicXilie.myFontString(0)
                FlowLayoutPanel1.Controls(5).Text = BigpicXilie.myFontString(1)
                FlowLayoutPanel1.Controls(7).BackColor = StringToColor(BigpicXilie.myFontString(2))
                FlowLayoutPanel1.Controls(9).BackColor = StringToColor(BigpicXilie.myFontString(3))
                Dim TempCheckBox1 As CheckBox = FlowLayoutPanel1.Controls(10)
                Dim TempCheckBox2 As CheckBox = FlowLayoutPanel1.Controls(11)
                TempCheckBox1.Checked = IIf(BigpicXilie.myFontString(4) = "True", True, False)
                TempCheckBox2.Checked = IIf(BigpicXilie.myFontString(5) = "True", True, False)
        End Select
    End Sub
    Public Sub FlowLayoutPanel1_ComboBox2_SelectedIndexChanged()
        Button4.Enabled = True
    End Sub

    Public Sub FlowLayoutPanel1_ComboBox3_SelectedIndexChanged()
        Button4.Enabled = True
    End Sub

    Private Sub FlowLayoutPanel1_CheckBox1_CheckedChanged()
        Button4.Enabled = True
    End Sub
    Private Sub FlowLayoutPanel1_CheckBox2_CheckedChanged()
        Button4.Enabled = True
    End Sub
    Private Sub FlowLayoutPanel1_CheckBoxView_CheckedChanged()
        Button4.Enabled = True
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim myXml As New CSysXML("config.xml")
        Select Case (TreeView1.SelectedNode.Text)
            Case "小图界面字体和颜色"
                SmallPicFanhao.myFontString(0) = "Times New Roman"
                SmallPicFanhao.myFontString(1) = 12
                SmallPicFanhao.myFontString(2) = "255 255 0 0 "
                SmallPicFanhao.myFontString(3) = "255 255 255 255 "
                SmallPicFanhao.myFontString(4) = False
                SmallPicFanhao.myFontString(5) = False
                myXml.SaveElement2("SmallPicForm", "Fanhao", SmallPicFanhao.myFontString)
                SmallPicMingcheng.myFontString(0) = "微软雅黑"
                SmallPicMingcheng.myFontString(1) = 10
                SmallPicMingcheng.myFontString(2) = "255 0 0 0 "
                SmallPicMingcheng.myFontString(3) = "255 255 255 255 "
                SmallPicMingcheng.myFontString(4) = False
                SmallPicMingcheng.myFontString(5) = False
                myXml.SaveElement2("SmallPicForm", "Mingcheng", SmallPicMingcheng.myFontString)
                SmallPicRiqi.myFontString(0) = "Times New Roman"
                SmallPicRiqi.myFontString(1) = 12
                SmallPicRiqi.myFontString(2) = "255 255 0 0 "
                SmallPicRiqi.myFontString(3) = "255 255 255 255 "
                SmallPicRiqi.myFontString(4) = False
                SmallPicRiqi.myFontString(5) = False
                myXml.SaveElement2("SmallPicForm", "Riqi", SmallPicRiqi.myFontString)
                SmallPicQita.myFontString(0) = "微软雅黑"
                SmallPicQita.myFontString(1) = 10
                SmallPicQita.myFontString(2) = "255 0 0 0 "
                SmallPicQita.myFontString(3) = "255 255 255 255 "
                SmallPicQita.myFontString(4) = False
                SmallPicQita.myFontString(5) = False
                myXml.SaveElement2("SmallPicForm", "Qita", SmallPicQita.myFontString)
            Case "大图界面字体和颜色"
                BigpicFanhao.myFontString(0) = "Times New Roman"
                BigpicFanhao.myFontString(1) = 16
                BigpicFanhao.myFontString(2) = "255 255 0 0 "
                BigpicFanhao.myFontString(3) = "255 255 255 255 "
                BigpicFanhao.myFontString(4) = False
                BigpicFanhao.myFontString(5) = False
                myXml.SaveElement2("BigPicForm", "Fanhao", BigpicFanhao.myFontString)
                BigpicMingcheng.myFontString(0) = "微软雅黑"
                BigpicMingcheng.myFontString(1) = 16
                BigpicMingcheng.myFontString(2) = "255 0 0 0 "
                BigpicMingcheng.myFontString(3) = "255 240 240 240 "
                BigpicMingcheng.myFontString(4) = False
                BigpicMingcheng.myFontString(5) = False
                myXml.SaveElement2("BigPicForm", "Mingcheng", BigpicMingcheng.myFontString)
                BigpicFaxingriqi.myFontString(0) = "Times New Roman"
                BigpicFaxingriqi.myFontString(1) = 16
                BigpicFaxingriqi.myFontString(2) = "255 0 0 0 "
                BigpicFaxingriqi.myFontString(3) = "255 255 255 255 "
                BigpicFaxingriqi.myFontString(4) = False
                BigpicFaxingriqi.myFontString(5) = False
                myXml.SaveElement2("BigPicForm", "Faxingriqi", BigpicFaxingriqi.myFontString)
                BigpicChangdu.myFontString(0) = "Times New Roman"
                BigpicChangdu.myFontString(1) = 16
                BigpicChangdu.myFontString(2) = "255 0 0 0 "
                BigpicChangdu.myFontString(3) = "255 255 255 255 "
                BigpicChangdu.myFontString(4) = False
                BigpicChangdu.myFontString(5) = False
                myXml.SaveElement2("BigPicForm", "Changdu", BigpicChangdu.myFontString)
                BigpicDaoyan.myFontString(0) = "微软雅黑"
                BigpicDaoyan.myFontString(1) = 12
                BigpicDaoyan.myFontString(2) = "255 0 0 0 "
                BigpicDaoyan.myFontString(3) = "255 255 255 255 "
                BigpicDaoyan.myFontString(4) = False
                BigpicDaoyan.myFontString(5) = False
                myXml.SaveElement2("BigPicForm", "Daoyan", BigpicDaoyan.myFontString)
                BigpicZhizuoshang.myFontString(0) = "微软雅黑"
                BigpicZhizuoshang.myFontString(1) = 12
                BigpicZhizuoshang.myFontString(2) = "255 0 0 0 "
                BigpicZhizuoshang.myFontString(3) = "255 255 255 255 "
                BigpicZhizuoshang.myFontString(4) = False
                BigpicZhizuoshang.myFontString(5) = False
                myXml.SaveElement2("BigPicForm", "Zhizuoshang", BigpicZhizuoshang.myFontString)
                BigpicFaxingshang.myFontString(0) = "微软雅黑"
                BigpicFaxingshang.myFontString(1) = 12
                BigpicFaxingshang.myFontString(2) = "255 0 0 0 "
                BigpicFaxingshang.myFontString(3) = "255 255 255 255 "
                BigpicFaxingshang.myFontString(4) = False
                BigpicFaxingshang.myFontString(5) = False
                myXml.SaveElement2("BigPicForm", "Faxingshang", BigpicFaxingshang.myFontString)
                BigpicLeibie.myFontString(0) = "微软雅黑"
                BigpicLeibie.myFontString(1) = 12
                BigpicLeibie.myFontString(2) = "255 0 0 0 "
                BigpicLeibie.myFontString(3) = "255 255 255 255 "
                BigpicLeibie.myFontString(4) = False
                BigpicLeibie.myFontString(5) = False
                myXml.SaveElement2("BigPicForm", "Leibie", BigpicLeibie.myFontString)
                BigpicYanyuan.myFontString(0) = "微软雅黑"
                BigpicYanyuan.myFontString(1) = 12
                BigpicYanyuan.myFontString(2) = "255 0 0 0 "
                BigpicYanyuan.myFontString(3) = "255 255 255 255 "
                BigpicYanyuan.myFontString(4) = False
                BigpicYanyuan.myFontString(5) = False
                myXml.SaveElement2("BigPicForm", "Yanyuan", BigpicYanyuan.myFontString)
                BigpicBiaoqian.myFontString(0) = "微软雅黑"
                BigpicBiaoqian.myFontString(1) = 12
                BigpicBiaoqian.myFontString(2) = "255 0 0 0 "
                BigpicBiaoqian.myFontString(3) = "255 255 255 255 "
                BigpicBiaoqian.myFontString(4) = False
                BigpicBiaoqian.myFontString(5) = False
                myXml.SaveElement2("BigPicForm", "Biaoqian", BigpicBiaoqian.myFontString)
                BigpicXilie.myFontString(0) = "微软雅黑"
                BigpicXilie.myFontString(1) = 12
                BigpicXilie.myFontString(2) = "255 0 0 0 "
                BigpicXilie.myFontString(3) = "255 255 255 255 "
                BigpicXilie.myFontString(4) = False
                BigpicXilie.myFontString(5) = False
                myXml.SaveElement2("BigPicForm", "Xilie", BigpicXilie.myFontString)
                'TreeView1.SelectedNode = TreeView1.Nodes(0).Nodes(1)
            Case "功能设置"
                IsClickSmalPicShowBigPic = False
                myXml.SaveElement("GongNeng", "IsClickSmalPicShowBigPic", IsClickSmalPicShowBigPic)
            Case "显示设置"
                IsCloseNow = False
                IsMingchengShow = True
                IsFanhaoShow = True
                IsRiqiShow = True
                IsShoucangShow = True
                myXml.SaveElement("SoftView", "IsCloseNow", IsCloseNow)
                myXml.SaveElement("SoftView", "IsMingchengShow", IsMingchengShow)
                myXml.SaveElement("SoftView", "IsFanhaoShow", IsFanhaoShow)
                myXml.SaveElement("SoftView", "IsRiqiShow", IsRiqiShow)
                myXml.SaveElement("SoftView", "IsShoucangShow", IsShoucangShow)
                IsSmallPicAutoSize = False
                myXml.SaveElement("SoftView", "IsSmallPicAutoSize", IsSmallPicAutoSize)
                FlowlayoutPanel_PageNum = 50
                myXml.SaveElement("SoftView", "FlowlayoutPanel_PageNum", FlowlayoutPanel_PageNum)
                'TreeView1.SelectedNode = TreeView1.Nodes(1)
            Case "设置网站"
                JavWebSite = "www.javbus.com"
                myXml.SaveElement("myWebSite", "JavWebSite", JavWebSite)
                'TreeView1.SelectedNode = TreeView1.Nodes(2)
            Case "存储位置"
                BigPicSavePath = Application.StartupPath & "\Pic\BigPic"
                myXml.SaveElement("SaveSetup", "BigPicSavePath", BigPicSavePath)
                DataBaseSavePath = Application.StartupPath & "\Mdb"
                myXml.SaveElement("SaveSetup", "DataBaseSavePath", DataBaseSavePath)
                SmallPicSavePath = Application.StartupPath & "\Pic\SmallPic"
                myXml.SaveElement("SaveSetup", "SmallPicSavePath", SmallPicSavePath)
                ActressesPicSavePath = Application.StartupPath & "\Pic\Actresses"
                myXml.SaveElement("SaveSetup", "ActressesPicSavePath", ActressesPicSavePath)
                ExtraPicSavePath = Application.StartupPath & "\Pic\Extra"
                myXml.SaveElement("SaveSetup", "ExtraPicSavePath", ExtraPicSavePath)
        End Select
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs)

    End Sub
End Class