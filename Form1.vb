Imports System.ComponentModel
Imports System.Text.RegularExpressions
Imports System.Configuration
Imports System.Data.OleDb
Imports System.IO
Public Class Form1
    'Private IsForm1Click As Boolean = False
    'Private Form1MouseClickPoint As Point
    'Private MouseWeizhi As Point
    Public ZidingyiClickIndex As Integer
    Public Maxmyfanhao As Integer
    Public zidingyiName As String
    '从数据库中读取载入


    Public Sub SelectFromDatabase(str As String, database As String, PageNum As Integer)
        FlowLayoutPanel1.Controls.Clear()
        Dim con As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dr As OleDbDataReader
        Try


            con.ConnectionString = con_ConnectionString
            con.Open()
            cmd.Connection = con '初始化OLEDB命令的连接属性为con
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try
        TotalNum = 0


        Select Case ToolStripComboBox1.Text
            Case "番号"
                If Zhengpaixu = True Then
                    PaixuSql = "select * from " & database & " " & str & " order by fanhao"
                Else
                    PaixuSql = "select * from " & database & " " & str & " order by fanhao desc"
                End If
            Case "名称"
                If Zhengpaixu = True Then
                    PaixuSql = "select * from " & database & " " & str & " order by mingcheng"
                Else
                    PaixuSql = "select * from " & database & " " & str & " order by mingcheng desc"
                End If
            Case "访问次数"
                If Zhengpaixu = True Then
                    PaixuSql = "select * from " & database & " " & str & " order by fangwencishu"
                Else
                    PaixuSql = "select * from " & database & " " & str & " order by fangwencishu desc"
                End If
            Case "文件大小"
                If Zhengpaixu = True Then
                    PaixuSql = "select * from " & database & " " & str & " order by wenjiandaxiao"
                Else
                    PaixuSql = "select * from " & database & " " & str & " order by wenjiandaxiao desc"

                End If
            Case "发行日期"
                If Zhengpaixu = True Then
                    PaixuSql = "select * from " & database & " " & str & " order by faxingriqi"
                Else
                    PaixuSql = "select * from " & database & " " & str & " order by faxingriqi desc"
                End If
            Case "导入时间"
                If Zhengpaixu = True Then
                    PaixuSql = "select * from " & database & " " & str & " order by daorushijian"
                Else
                    PaixuSql = "select * from " & database & " " & str & " order by daorushijian desc"
                End If
            Case "喜爱程度"
                If Zhengpaixu = True Then
                    PaixuSql = "select * from " & database & " " & str & " order by love"
                Else
                    PaixuSql = "select * from " & database & " " & str & " order by love desc"
                End If
            Case Else
                PaixuSql = "select * from  " & database
        End Select
        'MsgBox(PaixuSql)
        cmd.CommandText = PaixuSql
        Debug.Print(PaixuSql)
        dr = cmd.ExecuteReader() '执行OLEDB命令以ExecuteReader()方式，并返回一个OLEDBReader，赋值给dr
        While (dr.Read)
            TotalNum = TotalNum + 1
        End While

        dr.Close()
        con.Close()
        '再打开一次
        con.ConnectionString = con_ConnectionString
        con.Open()
        cmd.Connection = con '初始化OLEDB命令的连接属性为con
        cmd.CommandText = PaixuSql
        dr = cmd.ExecuteReader() '执行OLEDB命令以ExecuteReader()方式，并返回一个OLEDBReader，赋值给dr
        Dim PicPath As String
        Dim TooltipText As String
        Dim TooltipText1 As String
        Dim TooltipText2 As String
        Dim num As Integer
        num = 1
        While (dr.Read)

            If num = PageNum * FlowlayoutPanel_PageNum + 1 Then Exit While
            If num >= (PageNum - 1) * FlowlayoutPanel_PageNum + 1 And num <= TotalNum Then
                Dim mySmallPanel1 As New smallFrame
                mySmallPanel1.NameText = Juncode(dr.Item("mingcheng").ToString)
                PicPath = SmallPicSavePath & "\" & dr.Item("fanhao").ToString & ".jpg"
                mySmallPanel1.IsLove = IIf(Int(dr.Item("love")) > 0, True, False)
                If Dir(PicPath) <> "" Then
                    If My.Computer.FileSystem.GetFileInfo(PicPath).Length > 0 Then '如果图片错误会显示内存不足
                        Try


                            Dim pFileStream As New FileStream(PicPath, FileMode.Open, FileAccess.Read)
                            mySmallPanel1.Pic = Image.FromStream(pFileStream)
                            pFileStream.Close()
                            pFileStream.Dispose()
                        Catch ex As Exception
                            MsgBox(ex.Message)
                        End Try
                    End If
                End If
                Dim leixingstr As String
                leixingstr = ""
                Select Case dr.Item("shipinleixing")
                    Case 0
                        leixingstr = "其他"
                    Case 1
                        leixingstr = "日本步兵"
                    Case 2
                        leixingstr = "日本骑兵"
                    Case 3
                        leixingstr = "欧美"
                    Case 4
                        leixingstr = "国产"
                End Select
                'ToolTip1.RemoveAll()
                TooltipText = "大小：" & getFileSize(dr.Item("wenjiandaxiao").ToString) & vbCrLf & "位置：" &
                dr.Item("weizhi").ToString & vbCrLf & "访问次数：" & dr.Item("fangwencishu").ToString &
            vbCrLf & "长度：" & IIf(dr.Item("changdu").ToString <> "", dr.Item("changdu").ToString & "分钟", "") &
            vbCrLf & "视频类型：" & leixingstr
                ToolTip1.SetToolTip(mySmallPanel1.PicBox, TooltipText)

                TooltipText1 = "喜欢"
                ToolTip1.SetToolTip(mySmallPanel1.LovePicbox, TooltipText1)
                TooltipText2 = "添加到标签"
                ToolTip1.SetToolTip(mySmallPanel1.addToListPicbox, TooltipText2)
                mySmallPanel1.RiqiText = dr.Item("faxingriqi").ToString
                mySmallPanel1.FanhaoText = dr.Item("fanhao").ToString
                mySmallPanel1.CMS = ContextMenuStrip1
                mySmallPanel1.CMS2 = ContextMenuStrip2
                mySmallPanel1.CMSLove = ContextMenuStrip6
                mySmallPanel1.initial()
                FlowLayoutPanel1.Controls.Add(mySmallPanel1.Panel)

                Application.DoEvents()
            End If
            num = num + 1
        End While
        dr.Close()
        con.Close()
        ToolStripLabel2.Text = "本页有 " & FlowLayoutPanel1.Controls.Count & "个"
        ToolStripLabel5.Text = "总计有 " & TotalNum & "个"
        ToolStripLabel4.Text = "/" & Math.Ceiling(TotalNum / FlowlayoutPanel_PageNum) & " 页"
        ToolStripTextBox2.Text = PageNum
    End Sub

    Public Sub SelectFromDatabaseByInstr(str As String, strValue As String, PageNum As Integer)
        FlowLayoutPanel1.Controls.Clear()
        Dim con As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dr As OleDbDataReader
        Try
            con.ConnectionString = con_ConnectionString
            con.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try
        cmd.Connection = con '初始化OLEDB命令的连接属性为con

        TotalNum = 0
        PaixuSql = "select * from Main"

        'MsgBox(PaixuSql)
        cmd.CommandText = PaixuSql
        Debug.Print(PaixuSql)
        dr = cmd.ExecuteReader() '执行OLEDB命令以ExecuteReader()方式，并返回一个OLEDBReader，赋值给dr
        Dim myID(0) As Int16
        Dim myIDnum As Integer
        myIDnum = 0
        myID(0) = 0
        While (dr.Read)
            If InStr(dr(str).ToString, strValue) > 0 Then
                myID(myIDnum) = Int(dr("ID"))
                myIDnum = myIDnum + 1
                ReDim Preserve myID(myIDnum)
            End If
        End While
        ReDim Preserve myID(myIDnum - 1)
        TotalNum = myIDnum
        dr.Close()
        con.Close()
        '再打开一次
        con.ConnectionString = con_ConnectionString
        con.Open()
        cmd.Connection = con '初始化OLEDB命令的连接属性为con
        cmd.CommandText = PaixuSql
        dr = cmd.ExecuteReader() '执行OLEDB命令以ExecuteReader()方式，并返回一个OLEDBReader，赋值给dr
        Dim PicPath As String
        Dim TooltipText As String
        Dim TooltipText1 As String
        Dim TooltipText2 As String
        Dim num As Integer
        num = 1
        While (dr.Read)
            If num = PageNum * FlowlayoutPanel_PageNum + 1 Then Exit While
            If num >= (PageNum - 1) * FlowlayoutPanel_PageNum + 1 And num <= TotalNum Then
                Dim mySmallPanel1 As New smallFrame
                mySmallPanel1.NameText = dr.Item("mingcheng").ToString
                PicPath = SmallPicSavePath & "\" & dr.Item("fanhao").ToString & ".jpg"
                mySmallPanel1.IsLove = IIf(dr.Item("love").ToString = "1", True, False)
                If Dir(PicPath) <> "" Then
                    If My.Computer.FileSystem.GetFileInfo(PicPath).Length > 0 Then '如果图片错误会显示内存不足
                        Try

                            Dim pFileStream As New FileStream(PicPath, FileMode.Open, FileAccess.Read)
                            mySmallPanel1.Pic = Image.FromStream(pFileStream)
                            pFileStream.Close()
                            pFileStream.Dispose()
                        Catch ex As Exception
                            MsgBox(ex.Message)
                        End Try
                    End If
                End If
                Dim leixingstr As String
                leixingstr = ""
                Select Case dr.Item("shipinleixing")
                    Case 0
                        leixingstr = "其他"
                    Case 1
                        leixingstr = "日本步兵"
                    Case 2
                        leixingstr = "日本骑兵"
                    Case 3
                        leixingstr = "欧美"
                    Case 4
                        leixingstr = "国产"
                End Select
                'ToolTip1.RemoveAll()
                TooltipText = "大小：" & getFileSize(dr.Item("wenjiandaxiao").ToString) & vbCrLf & "位置：" &
                dr.Item("weizhi").ToString & vbCrLf & "访问次数：" & dr.Item("fangwencishu").ToString &
            vbCrLf & "长度：" & IIf(dr.Item("changdu").ToString <> "", dr.Item("changdu").ToString & "分钟", "") &
            vbCrLf & "视频类型：" & leixingstr
                ToolTip1.SetToolTip(mySmallPanel1.PicBox, TooltipText)

                TooltipText1 = "喜欢"
                ToolTip1.SetToolTip(mySmallPanel1.LovePicbox, TooltipText1)
                TooltipText2 = "添加到标签"
                ToolTip1.SetToolTip(mySmallPanel1.addToListPicbox, TooltipText2)
                mySmallPanel1.RiqiText = dr.Item("faxingriqi").ToString
                mySmallPanel1.FanhaoText = dr.Item("fanhao").ToString
                mySmallPanel1.CMS = ContextMenuStrip1
                mySmallPanel1.CMS2 = ContextMenuStrip2
                mySmallPanel1.initial()
                FlowLayoutPanel1.Controls.Add(mySmallPanel1.Panel)

            End If
            num = num + 1
        End While
        dr.Close()
        con.Close()
        ToolStripLabel2.Text = "本页有 " & FlowLayoutPanel1.Controls.Count & "个"
        ToolStripLabel5.Text = "总计有 " & TotalNum & "个"
        ToolStripLabel4.Text = "/" & Math.Ceiling(TotalNum / FlowlayoutPanel_PageNum) & " 页"
        ToolStripTextBox2.Text = PageNum
    End Sub




    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '避免重复打开
        If UBound(Diagnostics.Process.GetProcessesByName(Diagnostics.Process.GetCurrentProcess.ProcessName)) > 0 Then
            MsgBox("只允许打开一个！")
            End
        End If
        Dim myXml As New CSysXML("config.xml")
        Me_X = IIf(IsNumeric(myXml.GetElement("AppInitial", "Me_X")), myXml.GetElement("AppInitial", "Me_X"), 0)
        Me_Y = IIf(IsNumeric(myXml.GetElement("AppInitial", "Me_Y")), myXml.GetElement("AppInitial", "Me_Y"), 0)
        Me_Height = IIf(IsNumeric(myXml.GetElement("AppInitial", "Me_Height")), myXml.GetElement("AppInitial", "Me_Height"), 0)
        Me_Width = IIf(IsNumeric(myXml.GetElement("AppInitial", "Me_Width")), myXml.GetElement("AppInitial", "Me_Width"), 0)
        JavWebSite = IIf(myXml.GetElement("myWebSite", "JavWebSite") <> "$", myXml.GetElement("myWebSite", "JavWebSite"), "www.javbus.com")
        IsCloseNow = myXml.GetElement("SoftView", "IsCloseNow")
        IsFanhaoShow = myXml.GetElement("SoftView", "IsFanhaoShow")
        IsMingchengShow = myXml.GetElement("SoftView", "IsMingchengShow")
        IsRiqiShow = myXml.GetElement("SoftView", "IsRiqiShow")
        IsShoucangShow = myXml.GetElement("SoftView", "IsShoucangShow")
        IsSmallPicAutoSize = myXml.GetElement("SoftView", "IsSmallPicAutoSize")
        IsClickSmalPicShowBigPic = myXml.GetElement("GongNeng", "IsClickSmalPicShowBigPic")
        FlowlayoutPanel_PageNum = myXml.GetElement("SoftView", "FlowlayoutPanel_PageNum")
        Zhengpaixu = myXml.GetElement("DataBase", "Zhengpaixu")
        Paixuleixing = IIf(IsNumeric(myXml.GetElement("DataBase", "Paixuleixing")), myXml.GetElement("DataBase", "Paixuleixing"), 0)
        NewItem = myXml.GetElement("Item", "NewItem")
        BigPicSavePath = myXml.GetElement("SaveSetup", "BigPicSavePath")
        SmallPicSavePath = myXml.GetElement("SaveSetup", "SmallPicSavePath")
        ActressesPicSavePath = myXml.GetElement("SaveSetup", "ActressesPicSavePath")
        ExtraPicSavePath = myXml.GetElement("SaveSetup", "ExtraPicSavePath")
        DataBaseSavePath = myXml.GetElement("SaveSetup", "DataBaseSavePath")
        IsFirstStart = myXml.GetElement("SoftView", "IsFirstStart")
        ColorZhuti = myXml.GetElement("SoftView", "ColorZhuti")

        con_ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DataBaseSavePath & "\javbus.mdb"
        con_ConnectionString_read = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DataBaseSavePath & "\javbus.mdb;Mode=Read"
        If IsFirstStart = True Then
            BigPicSavePath = Application.StartupPath & "\Pic\BigPic"
            SmallPicSavePath = Application.StartupPath & "\Pic\SmallPic"
            ExtraPicSavePath = Application.StartupPath & "\Pic\Extra"
            ActressesPicSavePath = Application.StartupPath & "\Pic\Actresses"
            DataBaseSavePath = Application.StartupPath & "\mdb"
            IsFirstStart = False
            myXml.SaveElement("SoftView", "IsFirstStart", IsFirstStart)
        End If
        '读取字体设置
        For i = 0 To 4
            myXml.GetElement2("SmallPicForm", "Fanhao", SmallPicFanhao.myFontString)
            myXml.GetElement2("SmallPicForm", "Mingcheng", SmallPicMingcheng.myFontString)
            myXml.GetElement2("SmallPicForm", "Riqi", SmallPicRiqi.myFontString)
            myXml.GetElement2("SmallPicForm", "Qita", SmallPicQita.myFontString)
            myXml.GetElement2("BigPicForm", "Fanhao", BigpicFanhao.myFontString)
            myXml.GetElement2("BigPicForm", "Mingcheng", BigpicMingcheng.myFontString)
            myXml.GetElement2("BigPicForm", "Faxingriqi", BigpicFaxingriqi.myFontString)
            myXml.GetElement2("BigPicForm", "Changdu", BigpicChangdu.myFontString)
            myXml.GetElement2("BigPicForm", "Daoyan", BigpicDaoyan.myFontString)
            myXml.GetElement2("BigPicForm", "Zhizuoshang", BigpicZhizuoshang.myFontString)
            myXml.GetElement2("BigPicForm", "Zhizuoshang", BigpicFaxingshang.myFontString)
            myXml.GetElement2("BigPicForm", "Leibie", BigpicLeibie.myFontString)
            myXml.GetElement2("BigPicForm", "Yanyuan", BigpicYanyuan.myFontString)
            myXml.GetElement2("BigPicForm", "Biaoqian", BigpicBiaoqian.myFontString)
            myXml.GetElement2("BigPicForm", "Biaoqian", BigpicXilie.myFontString)
        Next

        '扫描设置
        Scan_Shipinleixing = myXml.GetElement("ScanSetup", "Scan_Shipinleixing")
        Scan_WenjianDaxiao = myXml.GetElement("ScanSetup", "Scan_WenjianDaxiao")
        BubingFanhao = myXml.GetElement("ScanSetup", "BubingFanhao")
        'ToolStripStatusLabel2.Text = Scan_WenjianDaxiao
        '默认设置
        If Me_X < -500 Or Me_Y < 0 Then
            Me_X = 400
            Me_Y = 200
        End If

        If Me_Height <= 0 Or Me_Width <= 0 Then
            Me.Height = Screen.PrimaryScreen.WorkingArea.Height / 2
            Me.Width = Screen.PrimaryScreen.WorkingArea.Width / 2
        End If
        Me.Left = Me_X
        Me.Top = Me_Y
        Me.Height = Me_Height
        Me.Width = Me_Width

        If Me_Width >= Screen.PrimaryScreen.WorkingArea.Width - 100 And Me_Height >= Screen.PrimaryScreen.WorkingArea.Height - 100 Then
            Me.Left = 400
            Me.Top = 200
            Me.Height = Screen.PrimaryScreen.WorkingArea.Height / 2
            Me.Width = Screen.PrimaryScreen.WorkingArea.Width / 2
            Me.WindowState = FormWindowState.Maximized
        End If

        For Each s In Split(NewItem, " ")
            If s <> "" Then
                Dim ToolTipbtn As New ToolStripButton
                ToolTipbtn.DisplayStyle = ToolStripItemDisplayStyle.Text
                ToolTipbtn.Text = s
                ToolTipbtn.Name = "ToolTipbtn_" & s
                ToolTipbtn.Font = ToolStripButton7.Font
                Select Case ColorZhuti
                    Case "默认"
                        ToolTipbtn.ForeColor = Color.Black
                    Case "灰色"
                        ToolTipbtn.ForeColor = Color.White
                    Case "黑色"
                        ToolTipbtn.ForeColor = Color.White
                End Select
                AddHandler ToolTipbtn.MouseDown, AddressOf ToolTipbtn_MouseDown
                ToolStrip3.Items.Add(ToolTipbtn)
                ContextMenuStrip2.Items.Add(s) '添加到ContextMenuStrip2中
            End If
        Next

        Dim myToolStripMenuItem As ToolStripMenuItem
        For Each item In ContextMenuStrip2.Items
            myToolStripMenuItem = item
            AddHandler myToolStripMenuItem.Click, AddressOf myToolStripMenuItem_Click
        Next

        ToolStripComboBox1.SelectedIndex = Paixuleixing

        CheckForIllegalCrossThreadCalls = False

        mySearchPattern = {"*.ts", "*.wma", "*.avi", "*.mp4", "*.mkv", "*.mpg", "*.rmvb", "*.rm", "*.mov", "*.mpeg", "*.flv", "*.wmv", "*.m4v"}


        '设置主题
        Select Case ColorZhuti
            Case "默认"
                FlowLayoutPanel1.BackColor = Color.FromArgb(255, 240, 240, 240)
                MenuStrip1.BackColor = Color.FromArgb(255, 240, 240, 240)
                ToolStrip1.BackColor = Color.FromArgb(255, 240, 240, 240)
                ToolStrip3.BackColor = Color.SkyBlue
                ToolStripComboBox1.BackColor = Color.White
                ToolStripTextBox1.BackColor = Color.White

                ToolStripTextBox2.BackColor = Color.White
                ToolStripLabel4.ForeColor = Color.Black
                ToolStripTextBox1.ForeColor = Color.Black

                ToolStripComboBox1.ForeColor = Color.Black
                ToolStripTextBox2.ForeColor = Color.Black

                ToolStripButton2.BackColor = Color.Transparent
                ToolStripButton4.BackColor = Color.Transparent
                ToolStripButton5.BackColor = Color.Transparent
                ToolStripButton13.BackColor = Color.Transparent


                文件ToolStripMenuItem.ForeColor = Color.Black
                编辑ToolStripMenuItem.ForeColor = Color.Black
                下载ToolStripMenuItem.ForeColor = Color.Black
                关于ToolStripMenuItem.ForeColor = Color.Black
                工具ToolStripMenuItem.ForeColor = Color.Black


                Dim mytoolstripbtn As ToolStripButton
                For Each item In ToolStrip3.Items
                    Debug.Print(item.GetType.ToString)
                    If item.GetType.ToString = "System.Windows.Forms.ToolStripButton" Then
                        mytoolstripbtn = item
                        mytoolstripbtn.ForeColor = Color.Black
                    End If
                Next


            Case "灰色"
                FlowLayoutPanel1.BackColor = Color.DimGray
                MenuStrip1.BackColor = Color.DimGray
                ToolStrip1.BackColor = Color.DimGray
                ToolStrip3.BackColor = Color.SlateGray
                ToolStripComboBox1.BackColor = Color.DimGray
                ToolStripTextBox1.BackColor = Color.DimGray

                ToolStripTextBox2.BackColor = Color.DimGray
                ToolStripLabel4.ForeColor = Color.White
                ToolStripTextBox1.ForeColor = Color.White

                ToolStripComboBox1.ForeColor = Color.White
                ToolStripTextBox2.ForeColor = Color.White

                ToolStripButton2.BackColor = Color.White
                ToolStripButton4.BackColor = Color.Transparent
                ToolStripButton5.BackColor = Color.Transparent
                ToolStripButton13.BackColor = Color.Transparent

                Dim mytoolstripbtn As ToolStripButton
                For Each item In ToolStrip3.Items
                    Debug.Print(item.GetType.ToString)
                    If item.GetType.ToString = "System.Windows.Forms.ToolStripButton" Then
                        mytoolstripbtn = item
                        mytoolstripbtn.ForeColor = Color.White
                    End If
                Next


            Case "黑色"
                FlowLayoutPanel1.BackColor = Color.Black
                MenuStrip1.BackColor = Color.Black
                ToolStrip1.BackColor = Color.Black
                ToolStrip3.BackColor = Color.Black
                ToolStripComboBox1.BackColor = Color.Black
                ToolStripTextBox1.BackColor = Color.Black

                ToolStripTextBox2.BackColor = Color.Black
                ToolStripLabel4.ForeColor = Color.White
                ToolStripTextBox1.ForeColor = Color.White

                ToolStripComboBox1.ForeColor = Color.White
                ToolStripTextBox2.ForeColor = Color.White

                ToolStripButton2.BackColor = Color.White
                ToolStripButton4.BackColor = Color.White
                ToolStripButton5.BackColor = Color.White
                ToolStripButton13.BackColor = Color.White

                文件ToolStripMenuItem.ForeColor = Color.White
                编辑ToolStripMenuItem.ForeColor = Color.White
                下载ToolStripMenuItem.ForeColor = Color.White
                关于ToolStripMenuItem.ForeColor = Color.White
                工具ToolStripMenuItem.ForeColor = Color.White


                Dim mytoolstripbtn As ToolStripButton
                For Each item In ToolStrip3.Items
                    Debug.Print(item.GetType.ToString)
                    If item.GetType.ToString = "System.Windows.Forms.ToolStripButton" Then
                        mytoolstripbtn = item
                        mytoolstripbtn.ForeColor = Color.White
                    End If
                Next

        End Select



    End Sub

    Private Sub myToolStripMenuItem_Click(sender As Object, e As EventArgs)
        'Select Case CType(sender, ContextMenuStrip).Text
        '    Case "菜单1"

        '    Case "菜单2"
        '        '菜单2的代码
        'End Select
        Dim ListName As String = sender.text
        '判断有没有重复
        Dim con As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dr As OleDbDataReader
        Try
            con.ConnectionString = con_ConnectionString
            con.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try
        cmd.Connection = con '初始化OLEDB命令的连接属性为con
        cmd.CommandText = "select * from Main where fanhao='" & GlobalFanhao & "'"
        dr = cmd.ExecuteReader() '执行OLEDB命令以ExecuteReader()方式，并返回一个OLEDBReader，赋值给dr
        Dim mybiaoqian As String
        mybiaoqian = ""
        While (dr.Read)
            mybiaoqian = Juncode(dr.Item("biaoqian").ToString)
        End While
        dr.Close()
        con.Close()
        For Each s In Split(mybiaoqian, " ")
            If ListName = s Then
                MsgBox("已存在该项！")
                Exit Sub
            End If
        Next
        '不存在的话则更新
        For Each s In Split(mybiaoqian, " ")
            If s <> "" Then
                ListName = ListName & " " & s
            End If
        Next
        Dim myDataAdapter As New OleDbDataAdapter()
        con.Open()
        cmd.Connection = con '初始化OLEDB命令的连接属性为con
        cmd.CommandText = "update Main set biaoqian='" & Jencode(ListName) & "' where fanhao='" & GlobalFanhao & "'"
        myDataAdapter.UpdateCommand = cmd
        cmd.ExecuteNonQuery()
        con.Close()
        'MsgBox(GlobalFanhao & "保存成功！")
    End Sub


    Private Sub Form1_MouseClick(sender As Object, e As MouseEventArgs) Handles Me.MouseClick

    End Sub

    Private Sub 导入ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 导入ToolStripMenuItem.Click
        myScanForm.ShowDialog()
    End Sub

    Private Sub 关于本软件ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 关于本软件ToolStripMenuItem.Click
        AboutBox1.Show()
    End Sub

    Private Sub 导入单个文件ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 导入单个文件ToolStripMenuItem.Click
        ToolStripButton4_Click(sender, e)
    End Sub

    Private Sub 设置ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 设置ToolStripMenuItem.Click
        MySetupForm.Show()
        MySetupForm.Activate()
    End Sub

    Private Sub ToolStripButton5_Click(sender As Object, e As EventArgs)
        Dim p As New Point()
        p.X = System.Windows.Forms.Control.MousePosition.X
        p.Y = System.Windows.Forms.Control.MousePosition.Y
        ContextMenuStrip2.Show(p)

    End Sub

    Private Sub Form1_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        Me_X = Me.Left
        Me_Y = Me.Top
        Me_Height = Me.Height
        Me_Width = Me.Width
        Dim myXml As New CSysXML("config.xml")
        myXml.SaveElement("AppInitial", "Me_X", Me_X)
        myXml.SaveElement("AppInitial", "Me_Y", Me_Y)
        myXml.SaveElement("AppInitial", "Me_Height", Me_Height)
        myXml.SaveElement("AppInitial", "Me_Width", Me_Width)
        myXml.SaveElement("Item", "NewItem", NewItem)
        If Zhengpaixu = True Then
            Zhengpaixu = False
        Else
            Zhengpaixu = True
        End If
        myXml.SaveElement("DataBase", "Zhengpaixu", Zhengpaixu)
    End Sub

    Private Sub Form1_Activated(sender As Object, e As EventArgs) Handles Me.Activated

    End Sub

    Private Sub 数据库管理ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 数据库管理ToolStripMenuItem.Click
        myDatabaseForm.Show()
    End Sub

    Private Sub 下载管理ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 下载管理ToolStripMenuItem.Click
        'myDownLoadForm.Show()
    End Sub

    Private Sub ToolStripButton6_Click(sender As Object, e As EventArgs) Handles ToolStripButton6.Click
        下载管理ToolStripMenuItem_Click(sender, e)
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        myScanForm.ShowDialog()
    End Sub

    Private Sub ToolStripButton_clickchangeColor(num As Integer)
        For i = 0 To 4
            If i = num - 1 Then
                ToolStrip3.Items(i).BackColor = Color.AliceBlue
            Else
                ToolStrip3.Items(i).BackColor = Color.Transparent
            End If
        Next
    End Sub

    Private Sub ToolStripButton7_Click(sender As Object, e As EventArgs) Handles ToolStripButton7.Click
        Clickindex = 1
        ToolStripButton_clickchangeColor(Clickindex)
        SelectFromDatabase("", "Main", 1)
        ToolStripLabel6.Text = "载入所有视频…… "
        'MsgBox(ColorToString(FlowLayoutPanel1.BackColor))
        移出自定义选项卡ToolStripMenuItem.Enabled = False
    End Sub

    Private Sub ToolStripButton8_Click(sender As Object, e As EventArgs) Handles ToolStripButton8.Click
        Clickindex = 2
        ToolStripButton_clickchangeColor(Clickindex)
        SelectFromDatabase("where love>0", "Main", 1)
        ToolStripLabel6.Text = "载入我的喜爱…… "
        移出自定义选项卡ToolStripMenuItem.Enabled = False
    End Sub

    Private Sub ToolStripButton9_Click(sender As Object, e As EventArgs) Handles ToolStripButton9.Click
        Clickindex = 3
        ToolStripButton_clickchangeColor(Clickindex)
        SelectFromDatabase("where shipinleixing=2", "Main", 1)
        ToolStripLabel6.Text = "载入骑兵…… "
        移出自定义选项卡ToolStripMenuItem.Enabled = False
    End Sub

    Private Sub ToolStripButton10_Click(sender As Object, e As EventArgs) Handles ToolStripButton10.Click
        Clickindex = 4
        ToolStripButton_clickchangeColor(Clickindex)
        SelectFromDatabase("where shipinleixing=1", "Main", 1)
        ToolStripLabel6.Text = "载入步兵…… "
        移出自定义选项卡ToolStripMenuItem.Enabled = False
    End Sub

    Private Sub ToolStripButton14_Click(sender As Object, e As EventArgs) Handles ToolStripButton14.Click
        Clickindex = 5
        ToolStripButton_clickchangeColor(Clickindex)
        SelectFromDatabase("where shipinleixing=3", "Main", 1)
        ToolStripLabel6.Text = "载入欧美…… "
        移出自定义选项卡ToolStripMenuItem.Enabled = False
    End Sub

    Private Sub MenuStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub

    Private Sub MenuStrip1_ItemAdded(sender As Object, e As ToolStripItemEventArgs) Handles MenuStrip1.ItemAdded
        'If (e.Item.Text.Length = 0) Then
        '    e.Item.Visible = False
        'End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub ToolStripButton15_Click(sender As Object, e As EventArgs)
        FlowLayoutPanel2.AutoSize = False
    End Sub

    Private Sub TextBox1_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub ToolStripButton15_Click_1(sender As Object, e As EventArgs)

    End Sub
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub FlowLayoutPanel1_Paint(sender As Object, e As PaintEventArgs) Handles FlowLayoutPanel1.Paint

    End Sub

    Private Sub FlowLayoutPanel1_MouseLeave(sender As Object, e As EventArgs) Handles FlowLayoutPanel1.MouseLeave

    End Sub

    Private Sub FlowLayoutPanel1_MouseWheel(sender As Object, e As MouseEventArgs) Handles FlowLayoutPanel1.MouseWheel
        Form1_MouseWheel(sender, e)
        'Call Me.Form1_MouseWheel(sender, e)
    End Sub

    Private Sub Form1_MouseWheel(sender As Object, e As MouseEventArgs) Handles Me.MouseWheel
        'e.X e.Y以窗体左上角为原点，aPoint为鼠标滚动时的坐标
        Dim aPoint As New Point(e.X, e.Y)
        'this.Location.X,this.Location.Y为窗体左上角相对于screen的坐标,得出的结果是鼠标相对于电脑screen的坐标
        aPoint.Offset(Me.Location.X, Me.Location.Y)
        Dim r As New Rectangle(FlowLayoutPanel1.Location.X, FlowLayoutPanel1.Location.Y, FlowLayoutPanel1.Width, FlowLayoutPanel1.Height)
        '判断鼠标是不是在flowLayoutPanel1区域内
        If (RectangleToScreen(r).Contains(aPoint)) Then
            '设置鼠标滚动幅度的大小
            FlowLayoutPanel1.AutoScrollPosition = New Point(0, FlowLayoutPanel1.VerticalScroll.Value - e.Delta / 2)
        End If
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        'MsgBox(sender.name)

    End Sub

    Private Sub PictureBox1_MouseWheel(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseWheel
        Form1_MouseWheel(sender, e)
    End Sub

    Private Sub FlowLayoutPanel1_Click(sender As Object, e As EventArgs) Handles FlowLayoutPanel1.Click
        'ToolStripTextBox1.Focus()
        FlowLayoutPanel1.Focus()
    End Sub

    Private Sub PictureBox1_MouseUp(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseUp
        'If e.Button = Windows.Forms.MouseButtons.Left Or e.Button = Windows.Forms.MouseButtons.Right Then

        'End If
    End Sub

    Private Sub 打开照片位置ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 打开照片位置ToolStripMenuItem.Click
        '判断文件夹是否存在
        Dim PicPath As String = SmallPicSavePath
        If System.IO.Directory.Exists(PicPath) = True Then
            'Process.Start(Application.StartupPath & "\pic\" & GlobalFanhao)
            Process.Start("Explorer.exe", "/select, """ & PicPath & "\" & GlobalFanhao & ".jpg" & """")
        End If


    End Sub

    Private Sub 打开影片位置ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 打开影片位置ToolStripMenuItem.Click
        Dim con As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dr As OleDbDataReader
        Try
            con.ConnectionString = con_ConnectionString
            con.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try
        cmd.Connection = con '初始化OLEDB命令的连接属性为con
        cmd.CommandText = "select * from Main where fanhao='" & GlobalFanhao & "'"
        dr = cmd.ExecuteReader() '执行OLEDB命令以ExecuteReader()方式，并返回一个OLEDBReader，赋值给dr

        'Dim PicPath As String
        While (dr.Read)
            If Dir(dr.Item("weizhi").ToString) <> "" Then
                ' System.Diagnostics.Process.Start(GetFilePath(dr.Item("weizhi").ToString))
                Process.Start("Explorer.exe", "/select, """ & dr.Item("weizhi").ToString & """")
                Exit While
            End If
        End While
        dr.Close()
        con.Close()
    End Sub

    Private Sub 打开其网址ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 打开其网址ToolStripMenuItem.Click
        If getshipinLeixingFromDatebase(GlobalFanhao) = 1 Then
            Process.Start("http://" & JavWebSite & "/uncensored/search/" & GlobalFanhao)
        ElseIf getshipinLeixingFromDatebase(GlobalFanhao) = 2 Then
            Process.Start("http://" & JavWebSite & "/search/" & GlobalFanhao)
        End If


    End Sub

    Private Sub 查看其信息ToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ToolTip1_Popup(sender As Object, e As PopupEventArgs) Handles ToolTip1.Popup

    End Sub

    Private Sub ToolStripComboBox1_Click(sender As Object, e As EventArgs) Handles ToolStripComboBox1.Click

    End Sub

    Private Sub ToolStripComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        'MsgBox(sender.text)
        If Zhengpaixu = True Then
            Zhengpaixu = False
        Else
            Zhengpaixu = True
        End If
        Select Case Clickindex
            Case 1
                ToolStripButton7_Click(sender, e)
            Case 2
                ToolStripButton8_Click(sender, e)
            Case 3
                ToolStripButton9_Click(sender, e)
            Case 4
                ToolStripButton10_Click(sender, e)
            Case 5
                ToolStripButton14_Click(sender, e)
            Case 10
                SelectFromDatabase("where leibie like '%" & myLeibie & "%'", "Main", ToolStripTextBox2.Text)
                ToolStripLabel6.Text = "载入【" & myLeibie & "】…… "
            Case 11
                SelectFromDatabase("where yanyuan like '%" & myYanyuan & "%'", "Main", ToolStripTextBox2.Text)
                ToolStripLabel6.Text = "载入【" & myYanyuan & "】…… "
            Case 12
                SelectFromDatabase("where biaoqian like '%" & myBiaoqian & "%'", "Main", ToolStripTextBox2.Text)
                ToolStripLabel6.Text = "载入【" & myBiaoqian & "】…… "
            Case 13
                SelectFromDatabase("where faxingshang like '%" & myFaxingshang & "%'", "Main", ToolStripTextBox2.Text)
                ToolStripLabel6.Text = "载入【" & myFaxingshang & "】…… "
            Case 14
                SelectFromDatabase("where zhizuoshang like '%" & myZhizuoshang & "%'", "Main", ToolStripTextBox2.Text)
                ToolStripLabel6.Text = "载入【" & myZhizuoshang & "】…… "
            Case 15
                SelectFromDatabase("where daoyan like '%" & myDaoyan & "%'", "Main", ToolStripTextBox2.Text)
                ToolStripLabel6.Text = "载入【" & myDaoyan & "】…… "
            Case 16
                SelectFromDatabase("where xilie like '%" & myXilie & "%'", "Main", ToolStripTextBox2.Text)
                ToolStripLabel6.Text = "载入【" & myXilie & "】…… "
            Case 20
                SelectFromDatabase("where fanhao like '%" & mySearch & "%' or mingcheng like '%" & mySearch & "%' ", "Main", ToolStripTextBox2.Text)
                ToolStripLabel6.Text = "载入【" & mySearch & "】…… "
        End Select
        Paixuleixing = ToolStripComboBox1.SelectedIndex
        Dim myXml As New CSysXML("config.xml")
        myXml.SaveElement("DataBase", "Paixuleixing", Paixuleixing)
        FlowLayoutPanel1.Focus()
    End Sub

    Private Sub ToolStrip3_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles ToolStrip3.ItemClicked

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ToolStripButton5_MouseDown(sender As Object, e As MouseEventArgs)

    End Sub

    Private Sub ToolStripButton17_Click(sender As Object, e As EventArgs) Handles ToolStripButton17.Click
        Dim str1
        str1 = InputBox("新的项")
        If str1 = "" Then Exit Sub
        If InStr(str1, " ") > 0 Then
            MsgBox("不允许有空格！")
            Exit Sub
        End If
        For Each item In ToolStrip3.Items
            If item.text = str1 Then
                MsgBox("不允许重复！")
                Exit Sub
            End If
        Next


        If str1 <> "" Then
            'MsgBox(Len(str1))
            str1 = UCase(str1)
            Label1.Text = str1
            Label1.Font = ToolStripButton7.Font
            Label1.AutoSize = True
            If Label1.Width > ToolStrip3.Width - 10 Then
                MsgBox("长度过长！")
                Exit Sub
            End If
            Dim ToolTipbtn As New ToolStripButton
            ToolTipbtn.DisplayStyle = ToolStripItemDisplayStyle.Text
            ToolTipbtn.Text = str1
            ToolTipbtn.Name = "ToolTipbtn_" & str1
            ToolTipbtn.Font = ToolStripButton7.Font
            Select Case ColorZhuti
                Case "默认"
                    ToolTipbtn.ForeColor = Color.Black
                Case "灰色"
                    ToolTipbtn.ForeColor = Color.White
                Case "黑色"
                    ToolTipbtn.ForeColor = Color.White
            End Select
            AddHandler ToolTipbtn.MouseDown, AddressOf ToolTipbtn_MouseDown
            ToolStrip3.Items.Add(ToolTipbtn)
            ContextMenuStrip2.Items.Add(str1) '添加到ContextMenuStrip2中
            Dim myToolStripMenuItem As ToolStripMenuItem
            For Each item In ContextMenuStrip2.Items
                myToolStripMenuItem = item
                AddHandler myToolStripMenuItem.Click, AddressOf myToolStripMenuItem_Click
            Next
            'AddHandler ToolTipbtn.Click, AddressOf ToolTipbtn_Click
            Dim myXml As New CSysXML("config.xml")
            NewItem = NewItem & " " & str1
            myXml.SaveElement("Item", "NewItem", NewItem)
        End If


    End Sub




    Private Sub ToolTipbtn_MouseDown(sender As Object, e As MouseEventArgs)
        zidingyiName = sender.text
        ToolTipbtnName = sender.name
        If e.Button = MouseButtons.Right Then
            Dim p As New Point()
            p.X = System.Windows.Forms.Control.MousePosition.X
            p.Y = System.Windows.Forms.Control.MousePosition.Y
            ContextMenuStrip3.Show(p)
        ElseIf e.Button = MouseButtons.Left Then
            '从数据库中载入
            Clickindex = 12
            myBiaoqian = Jencode(UCase(sender.text))
            SelectFromDatabase("where biaoqian like '%" & myBiaoqian & "%'", "Main", 1)
        End If

        移出自定义选项卡ToolStripMenuItem.Enabled = True


    End Sub

    Private Sub ContextMenuStrip3_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip3.Opening

    End Sub

    Private Sub 删除ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 删除ToolStripMenuItem.Click
        For Each item In ToolStrip3.Items
            If item.name = ToolTipbtnName Then
                NewItem = Replace(NewItem, Mid(ToolTipbtnName, 12), "")
                'MsgBox(ToolTipbtnName)
                Dim myXml As New CSysXML("config.xml")
                myXml.SaveElement("Item", "NewItem", NewItem)
                ToolStrip3.Items.Remove(item)
                Exit For
            End If
        Next
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        If ToolStripTextBox1.Text <> "" Then
            Clickindex = 20
            mySearch = Jencode(ToolStripTextBox1.Text)
            SelectFromDatabase("where fanhao like '%" & mySearch & "%' or mingcheng like '%" & mySearch & "%' ", "Main", 1)
        End If



        'Dim myFlowLayoutPanel As FlowLayoutPanel
        'Dim IsSearchFind As Boolean
        ''MsgBox(FlowLayoutPanel1.Controls.Count)
        'For i = FlowLayoutPanel1.Controls.Count - 1 To 0 Step -1
        '    myFlowLayoutPanel = FlowLayoutPanel1.Controls.Item(i)
        '    'myFlowLayoutPanel = flow
        '    IsSearchFind = False
        '    For Each item In myFlowLayoutPanel.Controls
        '        If item.GetType.ToString = "System.Windows.Forms.TextBox" Then
        '            If InStr(item.Text, ToolStripTextBox1.Text, CompareMethod.Text) > 0 Then
        '                Debug.Print("在" & item.text & "中找到了 " & ToolStripTextBox1.Text)
        '                IsSearchFind = True
        '            End If
        '        End If
        '    Next
        '    If IsSearchFind = False Then
        '        FlowLayoutPanel1.Controls.RemoveAt(i)
        '    End If
        'Next





        ToolStripLabel2.Text = "总计 " & FlowLayoutPanel1.Controls.Count & "个"

    End Sub

    Private Sub ToolStripTextBox1_Click(sender As Object, e As EventArgs) Handles ToolStripTextBox1.Click
        If ToolStripTextBox1.Text = "模糊查找番号、名称" Then
            ToolStripTextBox1.Text = ""
            Select Case ColorZhuti
                Case "默认"
                    ToolStripTextBox1.ForeColor = Color.Black
                Case "灰色"
                    ToolStripTextBox1.ForeColor = Color.White
                Case "黑色"
                    ToolStripTextBox1.ForeColor = Color.White
            End Select

        End If
    End Sub

    Private Sub ToolStripTextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles ToolStripTextBox1.KeyPress

    End Sub

    Private Sub ToolStripTextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ToolStripTextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            ToolStripButton1_Click(sender, e)
        End If
    End Sub





    Private Sub 数据库去重ToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ToolStripButton13_Click(sender As Object, e As EventArgs)
        'Debug.Print(getshipinLeixing("CWDV"))

    End Sub

    Private Sub ToolStripButton15_Click_2(sender As Object, e As EventArgs)

    End Sub

    Private Sub ToolStripButton13_Click_1(sender As Object, e As EventArgs) Handles ToolStripButton13.Click
        If Not IsNumeric(ToolStripTextBox2.Text) Then
            'MsgBox("页数错误")
            Exit Sub
        Else
            If ToolStripTextBox2.Text <= 0 Then
                'MsgBox("页数错误")
                Exit Sub
            ElseIf ToolStripTextBox2.Text >= Math.Ceiling(TotalNum / FlowlayoutPanel_PageNum) Then
                ToolStripTextBox2.Text = 1
            Else
                ToolStripTextBox2.Text = Int(ToolStripTextBox2.Text) + 1
            End If
        End If

        Select Case Clickindex
            Case 1
                SelectFromDatabase("", "Main", ToolStripTextBox2.Text)
            Case 2
                SelectFromDatabase("where love=1", "Main", ToolStripTextBox2.Text)
            Case 3
                SelectFromDatabase("where shipinleixing=2", "Main", ToolStripTextBox2.Text)
            Case 4
                SelectFromDatabase("where shipinleixing=1", "Main", ToolStripTextBox2.Text)
            Case 5
            Case 10
                SelectFromDatabase("where leibie like '%" & myLeibie & "%'", "Main", ToolStripTextBox2.Text)
            Case 11
                SelectFromDatabase("where yanyuan like '%" & myYanyuan & "%'", "Main", ToolStripTextBox2.Text)
            Case 12
                SelectFromDatabase("where biaoqian like '%" & myBiaoqian & "%'", "Main", ToolStripTextBox2.Text)
            Case 13
                SelectFromDatabase("where faxingshang like '%" & myFaxingshang & "%'", "Main", ToolStripTextBox2.Text)
            Case 14
                SelectFromDatabase("where zhizuoshang like '%" & myZhizuoshang & "%'", "Main", ToolStripTextBox2.Text)
            Case 15
                SelectFromDatabase("where daoyan like '%" & myDaoyan & "%'", "Main", ToolStripTextBox2.Text)
            Case 16
                SelectFromDatabase("where xilie like '%" & myXilie & "%'", "Main", ToolStripTextBox2.Text)
            Case 20
                SelectFromDatabase("where fanhao like '%" & mySearch & "%' or mingcheng like '%" & mySearch & "%' ", "Main", ToolStripTextBox2.Text)
        End Select
    End Sub

    Private Sub ToolStripTextBox2_Click(sender As Object, e As EventArgs) Handles ToolStripTextBox2.Click

    End Sub

    Private Sub ToolStripTextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ToolStripTextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            If Not IsNumeric(ToolStripTextBox2.Text) Then
                'MsgBox("页数错误")
                Exit Sub
            Else
                If ToolStripTextBox2.Text <= 0 Or ToolStripTextBox2.Text > Math.Ceiling(TotalNum / FlowlayoutPanel_PageNum) Then
                    'MsgBox("页数错误")
                    Exit Sub
                End If
            End If
            Select Case Clickindex
                Case 1
                    SelectFromDatabase("", "Main", ToolStripTextBox2.Text)
                Case 2
                    SelectFromDatabase("where love=1", "Main", ToolStripTextBox2.Text)
                Case 3
                    SelectFromDatabase("where shipinleixing=2", "Main", ToolStripTextBox2.Text)
                Case 4
                    SelectFromDatabase("where shipinleixing=1", "Main", ToolStripTextBox2.Text)
                Case 5
            End Select
        End If
    End Sub

    Private Sub ToolStripButton5_Click_1(sender As Object, e As EventArgs) Handles ToolStripButton5.Click
        If Not IsNumeric(ToolStripTextBox2.Text) Then
            'MsgBox("页数错误")
            Exit Sub
        Else
            If ToolStripTextBox2.Text > Math.Ceiling(TotalNum / FlowlayoutPanel_PageNum) Then
                'MsgBox("页数错误")
                Exit Sub
            ElseIf ToolStripTextBox2.Text <= 1 Then
                ToolStripTextBox2.Text = Math.Ceiling(TotalNum / FlowlayoutPanel_PageNum)
            Else
                ToolStripTextBox2.Text = Int(ToolStripTextBox2.Text) - 1
            End If
        End If

        Select Case Clickindex
            Case 1
                SelectFromDatabase("", "Main", ToolStripTextBox2.Text)
            Case 2
                SelectFromDatabase("where love=1", "Main", ToolStripTextBox2.Text)
            Case 3
                SelectFromDatabase("where shipinleixing=2", "Main", ToolStripTextBox2.Text)
            Case 4
                SelectFromDatabase("where shipinleixing=1", "Main", ToolStripTextBox2.Text)
            Case 5
            Case 10
                SelectFromDatabase("where leibie like '%" & myLeibie & "%'", "Main", ToolStripTextBox2.Text)
            Case 11
                SelectFromDatabase("where yanyuan like '%" & myYanyuan & "%'", "Main", ToolStripTextBox2.Text)
            Case 12
                SelectFromDatabase("where biaoqian like '%" & myBiaoqian & "%'", "Main", ToolStripTextBox2.Text)
            Case 13
                SelectFromDatabase("where faxingshang like '%" & myFaxingshang & "%'", "Main", ToolStripTextBox2.Text)
            Case 14
                SelectFromDatabase("where zhizuoshang like '%" & myZhizuoshang & "%'", "Main", ToolStripTextBox2.Text)
            Case 15
                SelectFromDatabase("where daoyan like '%" & myDaoyan & "%'", "Main", ToolStripTextBox2.Text)
            Case 16
                SelectFromDatabase("where xilie like '%" & myXilie & "%'", "Main", ToolStripTextBox2.Text)
            Case 20
                SelectFromDatabase("where fanhao like '%" & mySearch & "%' or mingcheng like '%" & mySearch & "%' ", "Main", ToolStripTextBox2.Text)
        End Select
    End Sub

    Private Sub ContextMenuStrip1_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip1.Opening

    End Sub

    Private Sub FlowLayoutPanel1_MouseDown(sender As Object, e As MouseEventArgs) Handles FlowLayoutPanel1.MouseDown


    End Sub

    Private Sub FlowLayoutPanel2_Paint(sender As Object, e As PaintEventArgs) Handles FlowLayoutPanel2.Paint

    End Sub

    Private Sub FlowLayoutPanel2_MouseMove(sender As Object, e As MouseEventArgs) Handles FlowLayoutPanel2.MouseMove

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub FlowLayoutPanel2_MouseLeave(sender As Object, e As EventArgs) Handles FlowLayoutPanel2.MouseLeave

    End Sub

    Private Sub ToolStripButton11_Click(sender As Object, e As EventArgs) Handles ToolStripButton11.Click
        LeibieForm.Show()
        移出自定义选项卡ToolStripMenuItem.Enabled = False
    End Sub

    Private Sub ToolStripButton7_MouseMove(sender As Object, e As MouseEventArgs) Handles ToolStripButton7.MouseMove

    End Sub

    Private Sub ToolStripButton7_MouseLeave(sender As Object, e As EventArgs) Handles ToolStripButton7.MouseLeave

    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ToolStripButton12_Click(sender As Object, e As EventArgs) Handles ToolStripButton12.Click
        YanyuanForm.Show()
        移出自定义选项卡ToolStripMenuItem.Enabled = False
    End Sub

    Private Sub 下载其小图ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 下载其小图ToolStripMenuItem.Click

    End Sub




    Private Sub 下载其大图和信息ToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub 小图ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 小图ToolStripMenuItem.Click
        Dim max As Double
        max = UBound(MyThread)
        For k = 0 To max Step 1
            If MyThread(k) IsNot Nothing Then
                If MyThread(k).IsAlive = True Then
                    MsgBox("其他线程正在下载！")
                    Exit Sub
                End If
            End If
        Next

        Dim con As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dr As OleDbDataReader
        Dim myShipinleixing As Integer
        Try
            con.ConnectionString = con_ConnectionString
            con.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try
        cmd.Connection = con '初始化OLEDB命令的连接属性为con
        PaixuSql = "select shipinleixing from Main where fanhao='" & GlobalFanhao & "'"
        cmd.CommandText = PaixuSql
        dr = cmd.ExecuteReader() '执行OLEDB命令以ExecuteReader()方式，并返回一个OLEDBReader，赋值给dr
        myShipinleixing = 0
        While (dr.Read)
            myShipinleixing = Int(dr.Item("shipinleixing"))
        End While
        dr.Close()
        con.Close()
        MultiThreadDownLoadSmallPicByFanhao(GlobalFanhao, True, myShipinleixing)
        ToolStripLabel6.Text = "下载番号  " & GlobalFanhao & " 的小图中"
    End Sub

    Private Sub 大图ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 大图ToolStripMenuItem.Click
        Dim max As Double
        max = UBound(MyThread)
        For k = 0 To max Step 1
            If MyThread(k) IsNot Nothing Then
                If MyThread(k).IsAlive = True Then
                    MsgBox("其他线程正在下载！")
                    Exit Sub
                End If
            End If
        Next
        Dim con As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dr As OleDbDataReader
        Dim myShipinleixing As Integer
        Try
            con.ConnectionString = con_ConnectionString
            con.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try
        cmd.Connection = con '初始化OLEDB命令的连接属性为con
        PaixuSql = "select shipinleixing from Main where fanhao='" & GlobalFanhao & "'"
        cmd.CommandText = PaixuSql
        dr = cmd.ExecuteReader() '执行OLEDB命令以ExecuteReader()方式，并返回一个OLEDBReader，赋值给dr
        myShipinleixing = 0
        While (dr.Read)
            myShipinleixing = Int(dr.Item("shipinleixing"))
        End While
        dr.Close()
        con.Close()
        MultiThreadDownLoadSmallPicByFanhao(GlobalFanhao, False, myShipinleixing)
        ToolStripLabel6.Text = "下载番号  " & GlobalFanhao & " 的大图中"
    End Sub

    Private Sub 信息ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 信息ToolStripMenuItem.Click
        Dim max As Double
        max = UBound(MyThread)
        For k = 0 To max Step 1
            If MyThread(k) IsNot Nothing Then
                If MyThread(k).IsAlive = True Then
                    MsgBox("其他线程正在下载！")
                    Exit Sub
                End If
            End If
        Next

        MultiThreadGetInfoByFanhao(GlobalFanhao)
        ToolStripLabel6.Text = "下载番号  " & GlobalFanhao & " 的信息中"
    End Sub



    Private Sub 欧美ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 欧美ToolStripMenuItem.Click
        Dim con As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim myDataAdapter As New OleDbDataAdapter()
        Try
            con.ConnectionString = con_ConnectionString
            con.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try
        cmd.Connection = con '初始化OLEDB命令的连接属性为con
        cmd.CommandText = "update Main set shipinleixing=3 where fanhao='" & GlobalFanhao & "'"
        myDataAdapter.UpdateCommand = cmd
        cmd.ExecuteNonQuery()
        con.Close()
        ToolStripLabel6.Text = "视频类型已更改为欧美"
    End Sub

    Private Sub 日本骑兵ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 日本骑兵ToolStripMenuItem.Click
        Dim con As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim myDataAdapter As New OleDbDataAdapter()
        Try
            con.ConnectionString = con_ConnectionString
            con.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try
        cmd.Connection = con '初始化OLEDB命令的连接属性为con
        cmd.CommandText = "update Main set shipinleixing=2 where fanhao='" & GlobalFanhao & "'"
        myDataAdapter.UpdateCommand = cmd
        cmd.ExecuteNonQuery()
        con.Close()
        ToolStripLabel6.Text = "视频类型已更改为日本骑兵"
    End Sub

    Private Sub 日本步兵ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 日本步兵ToolStripMenuItem.Click
        Dim con As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim myDataAdapter As New OleDbDataAdapter()
        Try
            con.ConnectionString = con_ConnectionString
            con.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try
        cmd.Connection = con '初始化OLEDB命令的连接属性为con
        cmd.CommandText = "update Main set shipinleixing=1 where fanhao='" & GlobalFanhao & "'"
        myDataAdapter.UpdateCommand = cmd
        cmd.ExecuteNonQuery()
        con.Close()
        ToolStripLabel6.Text = "视频类型已更改为日本步兵"
    End Sub

    Private Sub 国产ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 国产ToolStripMenuItem.Click
        Dim con As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim myDataAdapter As New OleDbDataAdapter()
        Try
            con.ConnectionString = con_ConnectionString
            con.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try
        cmd.Connection = con '初始化OLEDB命令的连接属性为con
        cmd.CommandText = "update Main set shipinleixing=4 where fanhao='" & GlobalFanhao & "'"
        myDataAdapter.UpdateCommand = cmd
        cmd.ExecuteNonQuery()
        con.Close()
        ToolStripLabel6.Text = "视频类型已更改为国产"
    End Sub

    Private Sub 其他ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 其他ToolStripMenuItem.Click
        Dim con As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim myDataAdapter As New OleDbDataAdapter()
        Try
            con.ConnectionString = con_ConnectionString
            con.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try
        cmd.Connection = con '初始化OLEDB命令的连接属性为con
        cmd.CommandText = "update Main set shipinleixing=0 where fanhao='" & GlobalFanhao & "'"
        myDataAdapter.UpdateCommand = cmd
        cmd.ExecuteNonQuery()
        con.Close()
        ToolStripLabel6.Text = "视频类型已更改为其他"
    End Sub

    Private Sub 修改信息ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 修改信息ToolStripMenuItem.Click

    End Sub

    Private Sub 刷新ToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub FlowLayoutPanel1_DragDrop(sender As Object, e As DragEventArgs) Handles FlowLayoutPanel1.DragDrop
        'Debug.Print(e.Data.GetData(DataFormats.FileDrop).count)
        Dim num As Integer = 1
        For Each s As String In e.Data.GetData(DataFormats.FileDrop) '循环枚举数据
            If num >= 2 Then Exit For
            Debug.Print(s)
            If InStr(".avi, .mp4, .mkv, .mpg, .rmvb, .rm, .mov, .mpeg,.flv, .wmv, .m4v", GetFileExtName(s)) > 0 Then
                'Debug.Print(s) '添加到表
                NewVedio.Show()
                NewVedio.TextBox1.Text = s
            End If
            num = num + 1
        Next
    End Sub

    Private Sub FlowLayoutPanel1_DragEnter(sender As Object, e As DragEventArgs) Handles FlowLayoutPanel1.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub FlowLayoutPanel1_QueryAccessibilityHelp(sender As Object, e As QueryAccessibilityHelpEventArgs) Handles FlowLayoutPanel1.QueryAccessibilityHelp

    End Sub

    Private Sub FlowLayoutPanel1_ControlAdded(sender As Object, e As ControlEventArgs) Handles FlowLayoutPanel1.ControlAdded

    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        NewVedio.Show()
    End Sub

    Private Sub ToolStripButton18_Click(sender As Object, e As EventArgs) Handles ToolStripButton18.Click
        'MsgBox("暂未开发")
        'ToolStripButton_clickchangeColor(Clickindex)
        SelectFromDatabase("where shipinleixing=4", "Main", 1)
        ToolStripLabel6.Text = "载入国产 "
        移出自定义选项卡ToolStripMenuItem.Enabled = False
    End Sub

    Private Sub ToolStripButton15_Click_3(sender As Object, e As EventArgs) Handles ToolStripButton15.Click
        MsgBox("暂未开发")
        移出自定义选项卡ToolStripMenuItem.Enabled = False
    End Sub

    Private Sub ToolStripButton3_Click_1(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        'MsgBox("暂未开发")
        SelectFromDatabase("where shipinleixing=0", "Main", 1)
        ToolStripLabel6.Text = "载入国产 "
        移出自定义选项卡ToolStripMenuItem.Enabled = False
    End Sub

    Private Sub 删除该番号ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 删除该番号ToolStripMenuItem.Click
        If MsgBox("确认删除番号：" & GlobalFanhao & "  ？", vbYesNo) = vbYes Then
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
            cmd.Connection = con '初始化OLEDB命令的连接属性为con
            cmd.CommandText = "delete from Main where fanhao='" & GlobalFanhao & "'"
            myDataAdapter.UpdateCommand = cmd
            cmd.ExecuteNonQuery()
            con.Close()
            ToolStripLabel6.Text = GlobalFanhao & "番号已删除！ "
        End If
    End Sub

    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If IsCloseNow = False Then
            For index = Application.OpenForms.Count - 1 To 0 Step -1
                If Application.OpenForms(index).Name <> "Form1" Then
                    Application.OpenForms(index).Close()
                End If
            Next
            e.Cancel = -1
            Me.Hide()
            NotifyIcon1.Visible = True

        Else '直接退出
            'For i = 0 To UBound(MyThread)
            '    If MyThread(i) IsNot Nothing Then

            '        MyThread(i).Abort()
            '        MyThread(i).Join()
            '    End If
            'Next
            End
            'Application.Exit()
        End If
    End Sub

    Private Sub NotifyIcon1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        Me.Show()
    End Sub

    Private Sub 退出ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 退出ToolStripMenuItem.Click
        NotifyIcon1.Visible = False
        End
    End Sub

    Private Sub 打开主窗体ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 打开主窗体ToolStripMenuItem.Click
        Me.Show()
        'NotifyIcon1.Visible = False
    End Sub

    Private Sub 检查更新ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 检查更新ToolStripMenuItem.Click
        'Debug.Print("https://" & JavWebSite & "/forum/home.php?mod=space&uid=85988&do=profile&from=space")
        'CheckForUpdate.Show()
        Process.Start("https://" & JavWebSite & "/forum/home.php?mod=space&uid=85988&do=profile&from=space")
    End Sub

    Private Sub ToolStripButton19_Click(sender As Object, e As EventArgs)
        Debug.Print(My.Computer.FileSystem.GetFileInfo("G:\hky\fileDownload\[Thz.la]fc2ppv_662049.mp4").Length)
    End Sub

    Private Sub ToolStripButton20_Click(sender As Object, e As EventArgs)
        ToolStripButton_clickchangeColor(100)
    End Sub

    Private Sub 刷新ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles 刷新ToolStripMenuItem1.Click
        Select Case Clickindex
            Case 1
                ToolStripButton7_Click(sender, e)
            Case 2
                ToolStripButton8_Click(sender, e)
            Case 3
                ToolStripButton9_Click(sender, e)
            Case 4
                ToolStripButton10_Click(sender, e)
            Case 5
                ToolStripButton14_Click(sender, e)
            Case Else
                ToolStripButton7_Click(sender, e)
        End Select
    End Sub

    Private Sub 数据库去重ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles 数据库去重ToolStripMenuItem1.Click
        Dim con As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dr As OleDbDataReader
        Dim myfanhao(0) As String
        Dim myID(0) As Integer
        Try


            con.ConnectionString = con_ConnectionString
            con.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try
        cmd.Connection = con '初始化OLEDB命令的连接属性为con
        PaixuSql = "select * from Main"
        cmd.CommandText = PaixuSql
        dr = cmd.ExecuteReader() '执行OLEDB命令以ExecuteReader()方式，并返回一个OLEDBReader，赋值给dr
        Dim i As Int16
        i = 0
        myfanhao(0) = ""
        myID(0) = 0
        While (dr.Read)
            i = i + 1
            ReDim Preserve myfanhao(i)
            ReDim Preserve myID(i)
            myfanhao(i) = dr.Item("fanhao").ToString
            myID(i) = Int(dr.Item("ID").ToString)
        End While
        dr.Close()
        con.Close()
        '将重复的项的ID记录到另一个数组
        Dim chongfu(0) As String
        chongfu(0) = ""
        Dim num As Integer
        num = 0
        For i = 1 To UBound(myfanhao) - 1
            For j = i + 1 To UBound(myfanhao)
                If myfanhao(i) = myfanhao(j) Then
                    num = num + 1
                    ReDim Preserve chongfu(num)
                    chongfu(num) = myID(j)
                End If
            Next
        Next

        Dim myDataAdapter As New OleDbDataAdapter()
        con.ConnectionString = con_ConnectionString
        con.Open()
        cmd.Connection = con '初始化OLEDB命令的连接属性为con
        For i = 1 To UBound(chongfu)
            cmd.CommandText = "delete from Main where ID=" & chongfu(i)
            myDataAdapter.UpdateCommand = cmd
            cmd.ExecuteNonQuery()
        Next
        con.Close()
        MsgBox("成功")
    End Sub

    Private Sub ToolStripTextBox1_LostFocus(sender As Object, e As EventArgs) Handles ToolStripTextBox1.LostFocus
        If ToolStripTextBox1.Text = "" Then
            ToolStripTextBox1.Text = "模糊查找番号、名称"
            Select Case ColorZhuti
                Case "默认"
                    ToolStripTextBox1.ForeColor = Color.Gray
                Case "灰色"
                    ToolStripTextBox1.ForeColor = Color.Gray
                Case "黑色"
                    ToolStripTextBox1.ForeColor = Color.Gray
            End Select
        End If
    End Sub

    Private Sub 下载数据库中所有小图ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 下载数据库中所有小图ToolStripMenuItem.Click
        Dim max As Double
        max = UBound(MyThread)
        For k = 0 To max Step 1
            If MyThread(k) IsNot Nothing Then
                If MyThread(k).IsAlive = True Then
                    MsgBox("其他线程正在下载！")
                    Exit Sub
                End If
            End If
        Next

        ToolStripProgressBar2.Value = 0
        Dim con As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dr As OleDbDataReader
        Dim myfanhao(0) As String
        Dim myShipinleixing(0) As Integer
        Try
            con.ConnectionString = con_ConnectionString
            con.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try

        cmd.Connection = con '初始化OLEDB命令的连接属性为con
        PaixuSql = "select fanhao,shipinleixing from Main"
        cmd.CommandText = PaixuSql
        dr = cmd.ExecuteReader() '执行OLEDB命令以ExecuteReader()方式，并返回一个OLEDBReader，赋值给dr
        Dim i As Int16
        i = 0
        myfanhao(0) = ""
        myShipinleixing(0) = 0
        While (dr.Read)
            myfanhao(i) = dr.Item("fanhao").ToString
            myShipinleixing(i) = Int(dr.Item("shipinleixing"))
            i = i + 1
            ReDim Preserve myfanhao(i)
            ReDim Preserve myShipinleixing(i)
        End While
        ReDim Preserve myfanhao(i - 1)
        ReDim Preserve myShipinleixing(i - 1)
        dr.Close()
        con.Close()
        Dim fanhao As String
        Dim vedioType As Integer


        '先判断要下载哪几个
        Dim totalthreadnum As Integer
        totalthreadnum = 0
        Dim myfanhao2(0) As String
        myfanhao2(0) = ""
        Dim myShipinleixing2(0) As Integer
        myShipinleixing2(0) = 0
        For i = 0 To UBound(myfanhao)
            fanhao = myfanhao(i)
            vedioType = myShipinleixing(i)
            Debug.Print(fanhao)
            Dim PicPath As String = SmallPicSavePath & "\" & fanhao & ".jpg"
            If Not File.Exists(PicPath) Then
                myfanhao2(totalthreadnum) = fanhao
                myShipinleixing2(totalthreadnum) = vedioType
                totalthreadnum = totalthreadnum + 1
                ReDim Preserve myfanhao2(totalthreadnum)
                ReDim Preserve myShipinleixing2(totalthreadnum)
            Else
                If My.Computer.FileSystem.GetFileInfo(PicPath).Length = 0 Then
                    IO.File.Delete(PicPath) '删除文件
                    myfanhao2(totalthreadnum) = fanhao
                    myShipinleixing2(totalthreadnum) = vedioType
                    totalthreadnum = totalthreadnum + 1
                    ReDim Preserve myfanhao2(totalthreadnum)
                    ReDim Preserve myShipinleixing2(totalthreadnum)
                End If
            End If
        Next

        Maxmyfanhao = UBound(myfanhao2)
        ToolStripLabel8.Text = "0%"
        'myThreadNum = 0
        Timer1.Enabled = True

        Dim MyDownClass(Maxmyfanhao) As DownClass
        ReDim MyThread(Maxmyfanhao)
        Dim mypicpath As String
        mypicpath = SmallPicSavePath
        For i = 0 To UBound(myfanhao2)
            fanhao = myfanhao2(i)
            vedioType = myShipinleixing2(i)
            Debug.Print(fanhao)
            If fanhao <> "" Then

                MyDownClass(i) = New DownClass
                MyDownClass(i).getWebResponse_fanhao = fanhao
                MyDownClass(i).myDownType = True
                MyDownClass(i).VedioType = vedioType
                'MyDownClass(i).index = myThreadNum

                If System.IO.Directory.Exists(mypicpath) = False Then
                    My.Computer.FileSystem.CreateDirectory(mypicpath)
                End If
                MyDownClass(i).FilePath = mypicpath
                'ReDim Preserve MyThread(i)
                MyThread(i) = New System.Threading.Thread(AddressOf MyDownClass(i).getWebResponseAndDownload)
                MyThread(i).Start()

            End If
            MySleep(200)
        Next
    End Sub

    Public Shared Sub MySleep(ByVal sleep_ms As Integer)
        Dim tick As Integer = Environment.TickCount
        While Environment.TickCount - tick < sleep_ms
            Application.DoEvents()
        End While
    End Sub

    Private Sub 下载数据库中所有大图ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 下载数据库中所有大图ToolStripMenuItem.Click
        Dim max As Double
        max = UBound(MyThread)
        For k = 0 To max Step 1
            If MyThread(k) IsNot Nothing Then
                If MyThread(k).IsAlive = True Then
                    MsgBox("其他线程正在下载！")
                    Exit Sub
                End If
            End If
        Next


        ToolStripProgressBar2.Value = 0
        Dim con As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dr As OleDbDataReader
        Dim myfanhao(0) As String
        Dim myShipinleixing(0) As Integer
        Try
            con.ConnectionString = con_ConnectionString
            con.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try
        cmd.Connection = con '初始化OLEDB命令的连接属性为con
        PaixuSql = "select fanhao,shipinleixing from Main"
        cmd.CommandText = PaixuSql
        dr = cmd.ExecuteReader() '执行OLEDB命令以ExecuteReader()方式，并返回一个OLEDBReader，赋值给dr
        Dim i As Int16
        i = 0
        myfanhao(0) = ""
        myShipinleixing(0) = 0
        While (dr.Read)
            myfanhao(i) = dr.Item("fanhao").ToString
            myShipinleixing(i) = Int(dr.Item("shipinleixing"))
            i = i + 1
            ReDim Preserve myfanhao(i)
            ReDim Preserve myShipinleixing(i)
        End While
        ReDim Preserve myfanhao(i - 1)
        ReDim Preserve myShipinleixing(i - 1)
        dr.Close()
        con.Close()
        Dim fanhao As String
        Dim vedioType As Integer

        '先判断要下载哪几个
        Dim totalthreadnum As Integer
        totalthreadnum = 0
        Dim myfanhao2(0) As String
        myfanhao2(0) = ""
        Dim myShipinleixing2(0) As Integer
        myShipinleixing2(0) = 0
        For i = 0 To UBound(myfanhao)
            fanhao = myfanhao(i)
            vedioType = myShipinleixing(i)
            Debug.Print(fanhao)
            Dim PicPath As String = BigPicSavePath & "\" & fanhao & ".jpg"
            If Not File.Exists(PicPath) Then
                myfanhao2(totalthreadnum) = fanhao
                myShipinleixing2(totalthreadnum) = vedioType
                totalthreadnum = totalthreadnum + 1
                ReDim Preserve myfanhao2(totalthreadnum)
                ReDim Preserve myShipinleixing2(totalthreadnum)
            Else
                If My.Computer.FileSystem.GetFileInfo(PicPath).Length = 0 Then
                    IO.File.Delete(PicPath) '删除文件
                    myfanhao2(totalthreadnum) = fanhao
                    myShipinleixing2(totalthreadnum) = vedioType
                    totalthreadnum = totalthreadnum + 1
                    ReDim Preserve myfanhao2(totalthreadnum)
                    ReDim Preserve myShipinleixing2(totalthreadnum)
                End If
            End If
        Next

        Maxmyfanhao = UBound(myfanhao2)
        ToolStripLabel8.Text = "0%"
        'myThreadNum = 0
        Timer1.Enabled = True

        Dim MyDownClass(Maxmyfanhao) As DownClass
        ReDim MyThread(Maxmyfanhao)
        Dim mypicpath As String
        mypicpath = BigPicSavePath
        For i = 0 To UBound(myfanhao2)
            fanhao = myfanhao2(i)
            vedioType = myShipinleixing2(i)
            Debug.Print(fanhao)
            If fanhao <> "" Then

                MyDownClass(i) = New DownClass
                MyDownClass(i).getWebResponse_fanhao = fanhao
                MyDownClass(i).myDownType = False
                MyDownClass(i).VedioType = vedioType
                'MyDownClass(i).index = myThreadNum

                If System.IO.Directory.Exists(mypicpath) = False Then
                    My.Computer.FileSystem.CreateDirectory(mypicpath)
                End If
                MyDownClass(i).FilePath = mypicpath
                'ReDim Preserve MyThread(i)
                MyThread(i) = New System.Threading.Thread(AddressOf MyDownClass(i).getWebResponseAndDownload)
                MyThread(i).Start()

            End If
            MySleep(800)
        Next
    End Sub

    Private Sub 下载数据库中所有番号信息ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 下载数据库中所有番号信息ToolStripMenuItem.Click
        Dim max As Double
        max = UBound(MyThread)
        For k = 0 To max Step 1
            If MyThread(k) IsNot Nothing Then
                If MyThread(k).IsAlive = True Then
                    MsgBox("其他线程正在下载！")
                    Exit Sub
                End If
            End If
        Next

        ToolStripProgressBar2.Value = 0
        Dim con As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dr As OleDbDataReader
        Dim myfanhao(0) As String
        Dim myMingcheng(0) As String
        Try
            con.ConnectionString = con_ConnectionString
            con.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try
        cmd.Connection = con '初始化OLEDB命令的连接属性为con
        PaixuSql = "select fanhao,mingcheng from Main"
        cmd.CommandText = PaixuSql
        dr = cmd.ExecuteReader() '执行OLEDB命令以ExecuteReader()方式，并返回一个OLEDBReader，赋值给dr
        Dim i As Int16
        i = 0
        myfanhao(0) = ""
        myMingcheng(0) = 0
        While (dr.Read)
            myfanhao(i) = dr.Item("fanhao").ToString
            myMingcheng(i) = dr.Item("mingcheng").ToString
            i = i + 1
            ReDim Preserve myfanhao(i)
            ReDim Preserve myMingcheng(i)
        End While
        ReDim Preserve myfanhao(i - 1)
        ReDim Preserve myMingcheng(i - 1)
        dr.Close()
        con.Close()

        Dim fanhao As String
        'Dim mingcheng As String


        '先判断要下载哪几个
        Dim totalthreadnum As Integer
        totalthreadnum = 0
        Dim myfanhao2(0) As String
        myfanhao2(0) = ""
        For i = 0 To UBound(myfanhao)
            fanhao = myfanhao(i)
            If myMingcheng(i) = "" Then
                myfanhao2(totalthreadnum) = fanhao
                totalthreadnum = totalthreadnum + 1
                ReDim Preserve myfanhao2(totalthreadnum)
            End If
        Next

        Maxmyfanhao = UBound(myfanhao2)
        ToolStripLabel8.Text = "0%"
        'myThreadNum = 0
        Timer1.Enabled = True

        Dim MyGetInfoClass(UBound(myfanhao2)) As GetInfoClass
        ReDim MyThread(UBound(myfanhao2))

        For i = 0 To UBound(myfanhao2)
            fanhao = myfanhao2(i)
            'MultiThreadGetInfoByFanhao(fanhao)

            If fanhao <> "" Then
                MyGetInfoClass(i) = New GetInfoClass
                MyGetInfoClass(i).getWebResponse_fanhao = fanhao

                MyThread(i) = New System.Threading.Thread(AddressOf MyGetInfoClass(i).getWebResponseAndGetInfo)
                MyThread(i).Start()
                Do
                    Application.DoEvents()
                    If MyGetInfoClass(i).GetInfoCompleted = True Then
                        Dim myDataAdapter As New OleDbDataAdapter()
                        Dim sql As String
                        con.ConnectionString = con_ConnectionString
                        con.Open()
                        cmd.Connection = con '初始化OLEDB命令的连接属性为con
                        Dim str1, str2 As String
                        str1 = ""
                        str2 = ""
                        For Each s In MyGetInfoClass(i).leibie
                            str1 = str1 & s & " "
                        Next
                        For Each s In MyGetInfoClass(i).yanyuan
                            str2 = str2 & s & " "
                        Next
                        str1 = Replace(str1, "'", "’")
                        str2 = Replace(str2, "'", "’")

                        sql = "update Main set faxingriqi = '" & MyGetInfoClass(i).faxingriqi & "', changdu = '" & MyGetInfoClass(i).changdu & "', mingcheng = '" & Replace(Jencode(MyGetInfoClass(i).mingcheng), "'", "’") & "', daoyan = '" & Replace(Jencode(MyGetInfoClass(i).daoyan), "'", "’") & "', zhizuoshang = '" & Replace(Jencode(MyGetInfoClass(i).zhizuoshang), "'", "’") & "', faxingshang = '" & Replace(Jencode(MyGetInfoClass(i).faxingshang), "'", "’") & "', xilie = '" & Replace(Jencode(MyGetInfoClass(i).xilie), "'", "’") & "', leibie = '" & Jencode(str1) & "', yanyuan = '" & Jencode(str2) & "'"
                        cmd.CommandText = sql & " where fanhao = '" & fanhao & "'"
                        myDataAdapter.UpdateCommand = cmd
                        cmd.ExecuteNonQuery()
                        con.Close()


                        Exit Do
                    End If
                Loop
                'ToolStripLabel6.Text = "正在下载  " & fanhao & " 的信息"
            End If
        Next
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim max As Double
        'max = Maxmyfanhao - 1
        max = UBound(MyThread)
        If max = 0 Then Exit Sub
        Dim num As Double
        num = 0
        For i = 0 To UBound(MyThread) Step 1
            If MyThread(i) IsNot Nothing Then
                'Debug.Print(myThreadIsCompleted(i))
                If MyThread(i).IsAlive = False Then
                    num = num + 1
                End If
            End If
        Next
        'Debug.Print(num)
        'Debug.Print(num / max)
        If num / max >= 1 Then
            ToolStripProgressBar2.Value = ToolStripProgressBar2.Maximum
            ToolStripLabel8.Text = "100%"
            'ReDim myThreadIsCompleted(0)
            'myThreadIsCompleted(0) = True
            Timer1.Enabled = False
            ToolStripLabel6.Text = "下载完成！"
        Else
            ToolStripProgressBar2.Value = ToolStripProgressBar2.Maximum * (num / max)
            ToolStripLabel8.Text = 100 * Math.Round(num / max, 3) & "%"
        End If
    End Sub

    Private Sub 喜欢等级ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 喜欢等级ToolStripMenuItem.Click
        Dim myflow As FlowLayoutPanel
        Dim PicB As PictureBox
        For Each flow In FlowLayoutPanel1.Controls
            myflow = flow
            For Each item In myflow.Controls
                If item.GetType.ToString = "System.Windows.Forms.PictureBox" Then
                    If InStr(item.name, GlobalFanhao) > 0 And InStr(item.name, "LovePicbox_") > 0 Then
                        PicB = item
                        GoTo mygoto1
                    End If
                End If
            Next
        Next

mygoto1:
        Debug.Print(PicB.Name)


        If InStr(PicB.Name, "LovePicbox_Love") > 0 Then
            PicB.Image = My.Resources.Resource1.xihuan
            PicB.Name = Replace(PicB.Name, "LovePicbox_Love_", "LovePicbox_")
            Dim con As New OleDbConnection
            Dim cmd As New OleDbCommand
            Dim myDataAdapter As New OleDbDataAdapter()
            con.ConnectionString = con_ConnectionString
            con.Open()
            cmd.Connection = con '初始化OLEDB命令的连接属性为con
            cmd.CommandText = "update Main set love=0 where fanhao='" & GlobalFanhao & "'"
            myDataAdapter.UpdateCommand = cmd
            cmd.ExecuteNonQuery()
            con.Close()
        End If

    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        Dim myflow As FlowLayoutPanel
        Dim PicB As PictureBox
        For Each flow In FlowLayoutPanel1.Controls
            myflow = flow
            For Each item In myflow.Controls
                If item.GetType.ToString = "System.Windows.Forms.PictureBox" Then
                    If InStr(item.name, GlobalFanhao) > 0 And InStr(item.name, "LovePicbox_") > 0 Then
                        PicB = item
                        GoTo mygoto1
                    End If
                End If
            Next
        Next

mygoto1:
        Debug.Print(PicB.Name)


        PicB.Image = My.Resources.Resource1.xihuan_fill
        If InStr(PicB.Name, "LovePicbox_NotLove_") > 0 Then
            PicB.Name = Replace(PicB.Name, "LovePicbox_NotLove_", "LovePicbox_Love_")
        End If
        Dim con As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim myDataAdapter As New OleDbDataAdapter()
        con.ConnectionString = con_ConnectionString
        con.Open()
        cmd.Connection = con '初始化OLEDB命令的连接属性为con
        cmd.CommandText = "update Main set love=1 where fanhao='" & GlobalFanhao & "'"
        myDataAdapter.UpdateCommand = cmd
        cmd.ExecuteNonQuery()
        con.Close()

    End Sub

    Private Sub ToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem3.Click
        Dim myflow As FlowLayoutPanel
        Dim PicB As PictureBox
        For Each flow In FlowLayoutPanel1.Controls
            myflow = flow
            For Each item In myflow.Controls
                If item.GetType.ToString = "System.Windows.Forms.PictureBox" Then
                    If InStr(item.name, GlobalFanhao) > 0 And InStr(item.name, "LovePicbox_") > 0 Then
                        PicB = item
                        GoTo mygoto1
                    End If
                End If
            Next
        Next

mygoto1:
        Debug.Print(PicB.Name)


        PicB.Image = My.Resources.Resource1.xihuan_fill
        If InStr(PicB.Name, "LovePicbox_NotLove_") > 0 Then
            PicB.Name = Replace(PicB.Name, "LovePicbox_NotLove_", "LovePicbox_Love_")
        End If
        Dim con As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim myDataAdapter As New OleDbDataAdapter()
        con.ConnectionString = con_ConnectionString
        con.Open()
        cmd.Connection = con '初始化OLEDB命令的连接属性为con
        cmd.CommandText = "update Main set love=2 where fanhao='" & GlobalFanhao & "'"
        myDataAdapter.UpdateCommand = cmd
        cmd.ExecuteNonQuery()
        con.Close()

    End Sub

    Private Sub ToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem4.Click
        Dim myflow As FlowLayoutPanel
        Dim PicB As PictureBox
        For Each flow In FlowLayoutPanel1.Controls
            myflow = flow
            For Each item In myflow.Controls
                If item.GetType.ToString = "System.Windows.Forms.PictureBox" Then
                    If InStr(item.name, GlobalFanhao) > 0 And InStr(item.name, "LovePicbox_") > 0 Then
                        PicB = item
                        GoTo mygoto1
                    End If
                End If
            Next
        Next

mygoto1:
        Debug.Print(PicB.Name)

        PicB.Image = My.Resources.Resource1.xihuan_fill
        If InStr(PicB.Name, "LovePicbox_NotLove_") > 0 Then
            PicB.Name = Replace(PicB.Name, "LovePicbox_NotLove_", "LovePicbox_Love_")
        End If
        Dim con As New OleDbConnection
            Dim cmd As New OleDbCommand
            Dim myDataAdapter As New OleDbDataAdapter()
            con.ConnectionString = con_ConnectionString
            con.Open()
            cmd.Connection = con '初始化OLEDB命令的连接属性为con
            cmd.CommandText = "update Main set love=3 where fanhao='" & GlobalFanhao & "'"
            myDataAdapter.UpdateCommand = cmd
            cmd.ExecuteNonQuery()
            con.Close()


    End Sub

    Private Sub ToolStripMenuItem5_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem5.Click
        Dim myflow As FlowLayoutPanel
        Dim PicB As PictureBox
        For Each flow In FlowLayoutPanel1.Controls
            myflow = flow
            For Each item In myflow.Controls
                If item.GetType.ToString = "System.Windows.Forms.PictureBox" Then
                    If InStr(item.name, GlobalFanhao) > 0 And InStr(item.name, "LovePicbox_") > 0 Then
                        PicB = item
                        GoTo mygoto1
                    End If
                End If
            Next
        Next

mygoto1:
        Debug.Print(PicB.Name)


        PicB.Image = My.Resources.Resource1.xihuan_fill
        If InStr(PicB.Name, "LovePicbox_NotLove_") > 0 Then
            PicB.Name = Replace(PicB.Name, "LovePicbox_NotLove_", "LovePicbox_Love_")
        End If
        Dim con As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim myDataAdapter As New OleDbDataAdapter()
        con.ConnectionString = con_ConnectionString
        con.Open()
        cmd.Connection = con '初始化OLEDB命令的连接属性为con
        cmd.CommandText = "update Main set love=4 where fanhao='" & GlobalFanhao & "'"
        myDataAdapter.UpdateCommand = cmd
        cmd.ExecuteNonQuery()
        con.Close()

    End Sub

    Private Sub ToolStripMenuItem6_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem6.Click
        Dim myflow As FlowLayoutPanel
        Dim PicB As PictureBox
        For Each flow In FlowLayoutPanel1.Controls
            myflow = flow
            For Each item In myflow.Controls
                If item.GetType.ToString = "System.Windows.Forms.PictureBox" Then
                    If InStr(item.name, GlobalFanhao) > 0 And InStr(item.name, "LovePicbox_") > 0 Then
                        PicB = item
                        GoTo mygoto1
                    End If
                End If
            Next
        Next

mygoto1:
        Debug.Print(PicB.Name)


        PicB.Image = My.Resources.Resource1.xihuan_fill
        If InStr(PicB.Name, "LovePicbox_NotLove_") > 0 Then
            PicB.Name = Replace(PicB.Name, "LovePicbox_NotLove_", "LovePicbox_Love_")
        End If
        Dim con As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim myDataAdapter As New OleDbDataAdapter()
        con.ConnectionString = con_ConnectionString
        con.Open()
        cmd.Connection = con '初始化OLEDB命令的连接属性为con
        cmd.CommandText = "update Main set love=5 where fanhao='" & GlobalFanhao & "'"
        myDataAdapter.UpdateCommand = cmd
        cmd.ExecuteNonQuery()
        con.Close()

    End Sub

    Private Sub 额外图片ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 额外图片ToolStripMenuItem.Click

    End Sub

    Private Sub 复制ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 复制ToolStripMenuItem.Click
        Dim con As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dr As OleDbDataReader
        Dim myWeizhi As String
        Try
            con.ConnectionString = con_ConnectionString
            con.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try
        cmd.Connection = con '初始化OLEDB命令的连接属性为con
        PaixuSql = "select weizhi from Main where fanhao='" & GlobalFanhao & "'"
        cmd.CommandText = PaixuSql
        dr = cmd.ExecuteReader() '执行OLEDB命令以ExecuteReader()方式，并返回一个OLEDBReader，赋值给dr
        myWeizhi = ""
        While (dr.Read)
            myWeizhi = dr.Item("weizhi").ToString
        End While
        dr.Close()
        con.Close()

        Dim myFileList = New Specialized.StringCollection
        myFileList.Add(myWeizhi)
        My.Computer.Clipboard.SetFileDropList(myFileList)
        MsgBox("复制成功")
    End Sub

    Private Sub 高级操作ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 高级操作ToolStripMenuItem.Click


        gaoji.ShowDialog()
    End Sub

    Private Sub 播放ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 播放ToolStripMenuItem.Click
        Dim con As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dr As OleDbDataReader
        Try
            con.ConnectionString = con_ConnectionString
            con.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try
        cmd.Connection = con '初始化OLEDB命令的连接属性为con
        cmd.CommandText = "select * from Main where fanhao='" & GlobalFanhao & "'"
        dr = cmd.ExecuteReader() '执行OLEDB命令以ExecuteReader()方式，并返回一个OLEDBReader，赋值给dr

        'Dim PicPath As String
        While (dr.Read)
            If Dir(dr.Item("weizhi").ToString) <> "" Then
                ' System.Diagnostics.Process.Start(GetFilePath(dr.Item("weizhi").ToString))
                'Process.Start("Explorer.exe", "/select, """ & dr.Item("weizhi").ToString & """")
                Shell("explorer " & dr.Item("weizhi").ToString)
                Exit While
            End If
        End While
        dr.Close()
        con.Close()
    End Sub

    Private Sub 删除该文件ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 删除该文件ToolStripMenuItem.Click
        Dim con As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dr As OleDbDataReader
        Try
            con.ConnectionString = con_ConnectionString
            con.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try
        cmd.Connection = con '初始化OLEDB命令的连接属性为con
        cmd.CommandText = "select * from Main where fanhao='" & GlobalFanhao & "'"
        dr = cmd.ExecuteReader() '执行OLEDB命令以ExecuteReader()方式，并返回一个OLEDBReader，赋值给dr

        'Dim PicPath As String
        While (dr.Read)
            If Dir(dr.Item("weizhi").ToString) <> "" Then
                Dim fpath As String = dr.Item("weizhi").ToString
                ' System.Diagnostics.Process.Start(GetFilePath(dr.Item("weizhi").ToString))
                'Process.Start("Explorer.exe", "/select, """ & dr.Item("weizhi").ToString & """")

                '删除文件file的方法1:删除到回收站里面。
                My.Computer.FileSystem.DeleteFile(fpath, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin, FileIO.UICancelOption.DoNothing)
                '删除文件file的方法2:直接从硬盘上删除。
                'IO.File.Delete(fpath)

                '删除该panel


                Exit While
            End If
        End While
        dr.Close()
        con.Close()
        MsgBox("删除成功！")
    End Sub

    Private Sub ToolStripButton19_Click_1(sender As Object, e As EventArgs) Handles ToolStripButton19.Click
        Debug.Print(ToolStrip3.Items.Item(0).GetType.ToString)
    End Sub

    Private Sub 番号ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 番号ToolStripMenuItem.Click
        Dim str1
        str1 = UCase(InputBox("新番号", "新番号", GlobalFanhao))

        If str1 <> "" Then
            '查看是否重复
            Dim con As New OleDbConnection
            Dim cmd As New OleDbCommand
            Dim dr As OleDbDataReader
            Dim sql As String
            Try
                con.ConnectionString = con_ConnectionString
                con.Open()
            Catch ex As Exception
                MsgBox(ex.Message)
                Exit Sub
            End Try
            cmd.Connection = con '初始化OLEDB命令的连接属性为con
            sql = "select fanhao from Main"
            cmd.CommandText = sql
            dr = cmd.ExecuteReader() '执行OLEDB命令以ExecuteReader()方式，并返回一个OLEDBReader，赋值给dr
            While (dr.Read)
                If UCase(dr(0).ToString) = str1 Then
                    MsgBox("数据库里已有该番号！")
                    con.Close()
                    Exit Sub
                End If
            End While
            Dim myDataAdapter As New OleDbDataAdapter()
            con.Close()
            con.ConnectionString = con_ConnectionString
            con.Open()
            cmd.Connection = con '初始化OLEDB命令的连接属性为con
            cmd.CommandText = "update Main set fanhao='" & str1 & "' where fanhao='" & GlobalFanhao & "'"
            myDataAdapter.UpdateCommand = cmd
            cmd.ExecuteNonQuery()
            con.Close()
            ToolStripLabel6.Text = "番号已更改为 " & str1
        End If
    End Sub

    Private Sub 其他信息ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 其他信息ToolStripMenuItem.Click

    End Sub

    Private Sub 使用说明ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 使用说明ToolStripMenuItem.Click
        shuoming.Show()
    End Sub

    Private Sub 移出自定义选项卡ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 移出自定义选项卡ToolStripMenuItem.Click
        If MsgBox("确认移出" & zidingyiName & "  ？", vbYesNo) = vbYes Then
            Dim con As New OleDbConnection
            Dim cmd As New OleDbCommand
            Dim dr As OleDbDataReader
            Dim sql As String
            Try
                con.ConnectionString = con_ConnectionString
                con.Open()
            Catch ex As Exception
                MsgBox(ex.Message)
                Exit Sub
            End Try
            cmd.Connection = con '初始化OLEDB命令的连接属性为con
            sql = "select biaoqian from Main where fanhao='" & GlobalFanhao & "'"
            cmd.CommandText = sql
            dr = cmd.ExecuteReader() '执行OLEDB命令以ExecuteReader()方式，并返回一个OLEDBReader，赋值给dr
            Dim mybiaoqian As String
            mybiaoqian = ""
            While (dr.Read)
                mybiaoqian = Juncode(dr(0).ToString)
            End While
            Debug.Print(mybiaoqian)
            Dim str1 As String() = Split(mybiaoqian, " ")

            For i = 0 To UBound(str1)
                If zidingyiName = str1(i) Then
                    str1(i) = ""
                End If
            Next
            Dim t As String
            t = ""
            For i = 0 To UBound(str1)
                If str1(i) <> "" Then
                    t = t & " " & str1(i)
                End If
            Next
            Debug.Print(t)
            Dim myDataAdapter As New OleDbDataAdapter()
            con.Close()
            con.ConnectionString = con_ConnectionString
            con.Open()
            cmd.Connection = con '初始化OLEDB命令的连接属性为con
            cmd.CommandText = "update Main set biaoqian='" & t & "' where fanhao='" & GlobalFanhao & "'"
            myDataAdapter.UpdateCommand = cmd
            cmd.ExecuteNonQuery()
            con.Close()
            ToolStripLabel6.Text = "已从该选项卡删除！"
        End If
    End Sub
End Class







