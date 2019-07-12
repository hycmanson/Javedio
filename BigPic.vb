Imports System.ComponentModel
Imports System.Data.OleDb
Imports System.IO
Public Class BigPic

    Public YingpianWeizhi As String
    Public myPicboxName As String
    Delegate Sub WT()




    Private Sub myLabel_Leibie_Click(sender As Object, e As EventArgs)
        Clickindex = 10
        myLeibie = Jencode(sender.text)
        Form1.SelectFromDatabase("where leibie like '%" & myLeibie & "%'", "Main", 1)
        Me.Close()
        Form1.Activate()
    End Sub

    Private Sub myLabel_Yanyuan_Click(sender As Object, e As EventArgs)
        Clickindex = 11
        myYanyuan = Jencode(sender.text)
        Form1.SelectFromDatabase("where yanyuan like '%" & myYanyuan & "%'", "Main", 1)
        Me.Close()
        Form1.Activate()
    End Sub

    Private Sub myLabel_Biaoqian_Click(sender As Object, e As EventArgs)
        Clickindex = 12
        myBiaoqian = Jencode(sender.text)
        Form1.SelectFromDatabase("where biaoqian like '%" & myBiaoqian & "%'", "Main", 1)
        Me.Close()
        Form1.Activate()
    End Sub

    Public Sub ToolStripButton5_Click(sender As Object, e As EventArgs) Handles ToolStripButton5.Click
        Dim mySingleThread As New System.Threading.Thread(AddressOf weituo)
        mySingleThread.Start()
    End Sub

    Private Sub weituo()
        Me.Invoke(New WT(AddressOf LoadInfo))
    End Sub

    Private Sub LoadInfo()
        Me.Text = GlobalFanhao
        Dim fangwencishu As Integer
        Dim con As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dr As OleDbDataReader
        con.ConnectionString = con_ConnectionString
        con.Open()
        cmd.Connection = con '初始化OLEDB命令的连接属性为con
        cmd.CommandText = "select *  from Main where fanhao='" & GlobalFanhao & "'"
        dr = cmd.ExecuteReader() '执行OLEDB命令以ExecuteReader()方式，并返回一个OLEDBReader，赋值给dr
        Dim leibie As String
        Dim yanyuan As String
        Dim biaoqian As String
        leibie = ""
        yanyuan = ""
        biaoqian = ""
        'Dim str1 As String




        While (dr.Read)
            TextBox1.Text = Juncode(dr.Item("mingcheng").ToString)
            Label14.Text = dr.Item("fanhao").ToString
            Label15.Text = dr.Item("faxingriqi").ToString
            Label16.Text = IIf(dr.Item("changdu").ToString <> "", dr.Item("changdu").ToString & "分钟", "")

            Label17.Text = Juncode(dr.Item("daoyan").ToString)
            Label18.Text = Juncode(dr.Item("zhizuoshang").ToString)
            Label19.Text = Juncode(dr.Item("faxingshang").ToString)
            Label21.Text = Juncode(dr.Item("xilie").ToString)

            leibie = dr.Item("leibie").ToString
            yanyuan = dr.Item("yanyuan").ToString
            biaoqian = dr.Item("biaoqian").ToString
            YingpianWeizhi = dr.Item("weizhi").ToString
            fangwencishu = Int(dr.Item("fangwencishu").ToString)
        End While
        dr.Close()
        con.Close()
        PictureBox1.Image = Nothing
        '加载大图
        Dim PicPath As String
        PicPath = BigPicSavePath & "\" & GlobalFanhao & ".jpg"
        If Dir(PicPath) <> "" Then
            If My.Computer.FileSystem.GetFileInfo(PicPath).Length > 0 Then '如果图片错误会显示内存不足
                Try
                    Dim pFileStream As New FileStream(PicPath, FileMode.Open, FileAccess.Read)
                    PictureBox1.Image = Image.FromStream(pFileStream)
                    pFileStream.Close()
                    pFileStream.Dispose()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End If
        End If

        '加载额外图片
        FlowLayoutPanel4.Controls.Clear()
        Dim picW As Double
        Dim picH As Double
        PicPath = ExtraPicSavePath & "\" & GlobalFanhao
        Debug.Print(PicPath)
        If System.IO.Directory.Exists(PicPath) = True Then

            Dim foundFile2 As String() = My.Computer.FileSystem.GetFiles(PicPath, FileIO.SearchOption.SearchAllSubDirectories, {"*.jpg", "*.bmp", "*.gif"}).ToArray




            For Each foundFile In foundFile2
                'foundFile = t.ToString
                Debug.Print(foundFile)
                If File.Exists(foundFile) Then

                    If My.Computer.FileSystem.GetFileInfo(foundFile).Length > 0 Then '如果图片错误会显示内存不足
                        PictureBox2.SizeMode = PictureBoxSizeMode.AutoSize

                        Try
                            Dim pFileStream As New FileStream(foundFile, FileMode.Open, FileAccess.Read)
                            PictureBox2.Image = Image.FromStream(pFileStream)
                            pFileStream.Close()
                            pFileStream.Dispose()
                        Catch ex As Exception
                        End Try


                        picW = PictureBox2.Width
                        picH = PictureBox2.Height
                        If picH = 0 Then picH = 200
                        Dim myPicBox As New PictureBox
                        myPicBox.Name = "myPicBox_" & GetFileName(foundFile)
                        myPicBox.Cursor = Windows.Forms.Cursors.Hand
                        myPicBox.Height = 200
                        myPicBox.Width = 200 * picW / picH
                        myPicBox.SizeMode = PictureBoxSizeMode.StretchImage

                        '从流中读取图片
                        Try
                            Dim pFileStream2 As New FileStream(foundFile, FileMode.Open, FileAccess.Read)
                            myPicBox.Image = Image.FromStream(pFileStream2)
                            pFileStream2.Close()
                            pFileStream2.Dispose()
                        Catch ex As Exception
                            MsgBox(ex.Message)
                        End Try
                        myPicBox.ContextMenuStrip = ContextMenuStrip2

                        FlowLayoutPanel4.Controls.Add(myPicBox)
                        AddHandler myPicBox.MouseDown, AddressOf myPicBox_MouseDown
                    End If

                End If
            Next
        End If

        '加载类别
        FlowLayoutPanel3.Controls.Clear()
        'Label9.Text = leibie
        'Label9.AutoSize = True
        'Label9.Visible = True
        For Each s In Split(leibie, " ")
            Dim myLabel As New Label
            myLabel.Text = Juncode(s)
            myLabel.AutoSize = True
            myLabel.Cursor = Windows.Forms.Cursors.Hand

            If BigpicLeibie.myFontString(4) = "True" And BigpicLeibie.myFontString(5) = "True" Then
                myLabel.Font = New System.Drawing.Font(BigpicLeibie.myFontString(0), Int(BigpicLeibie.myFontString(1)), FontStyle.Bold Or FontStyle.Italic)
            ElseIf BigpicLeibie.myFontString(4) = "True" And BigpicLeibie.myFontString(5) <> "True" Then
                myLabel.Font = New System.Drawing.Font(BigpicLeibie.myFontString(0), Int(BigpicLeibie.myFontString(1)), FontStyle.Bold)
            ElseIf BigpicLeibie.myFontString(4) <> "True" And BigpicLeibie.myFontString(5) = "True" Then
                myLabel.Font = New System.Drawing.Font(BigpicLeibie.myFontString(0), Int(BigpicLeibie.myFontString(1)), FontStyle.Italic)
            ElseIf BigpicLeibie.myFontString(4) <> "True" And BigpicLeibie.myFontString(5) <> "True" Then
                myLabel.Font = New System.Drawing.Font(BigpicLeibie.myFontString(0), Int(BigpicLeibie.myFontString(1)))
            End If
            'myLabel.ForeColor = StringToColor(BigpicLeibie.myFontString(2))
            'myLabel.BackColor = StringToColor(BigpicLeibie.myFontString(3))

            Select Case ColorZhuti
                Case "默认"
                    myLabel.ForeColor = Color.Black
                Case "灰色"
                    myLabel.ForeColor = Color.White
                Case "黑色"
                    myLabel.ForeColor = Color.White
            End Select

            myLabel.BackColor = Color.Transparent
            AddHandler myLabel.Click, AddressOf myLabel_Leibie_Click


            myLabel.Margin = New Padding(3, 3, 3, 0)
            FlowLayoutPanel3.Controls.Add(myLabel)
        Next

        '加载演员
        FlowLayoutPanel6.Controls.Clear()
        'Label10.Text = yanyuan
        'Label10.AutoSize = True
        'Label10.Visible = True
        For Each s In Split(yanyuan, " ")
            Dim myLabel As New Label
            myLabel.Text = Juncode(s)
            myLabel.AutoSize = True
            myLabel.Cursor = Windows.Forms.Cursors.Hand

            If BigpicYanyuan.myFontString(4) = "True" And BigpicYanyuan.myFontString(5) = "True" Then
                myLabel.Font = New System.Drawing.Font(BigpicYanyuan.myFontString(0), Int(BigpicYanyuan.myFontString(1)), FontStyle.Bold Or FontStyle.Italic)
            ElseIf BigpicYanyuan.myFontString(4) = "True" And BigpicYanyuan.myFontString(5) <> "True" Then
                myLabel.Font = New System.Drawing.Font(BigpicYanyuan.myFontString(0), Int(BigpicYanyuan.myFontString(1)), FontStyle.Bold)
            ElseIf BigpicYanyuan.myFontString(4) <> "True" And BigpicYanyuan.myFontString(5) = "True" Then
                myLabel.Font = New System.Drawing.Font(BigpicYanyuan.myFontString(0), Int(BigpicYanyuan.myFontString(1)), FontStyle.Italic)
            ElseIf BigpicYanyuan.myFontString(4) <> "True" And BigpicYanyuan.myFontString(5) <> "True" Then
                myLabel.Font = New System.Drawing.Font(BigpicYanyuan.myFontString(0), Int(BigpicYanyuan.myFontString(1)))
            End If
            'myLabel.ForeColor = StringToColor(BigpicYanyuan.myFontString(2))
            'myLabel.BackColor = StringToColor(BigpicYanyuan.myFontString(3))

            Select Case ColorZhuti
                Case "默认"
                    myLabel.ForeColor = Color.Black
                Case "灰色"
                    myLabel.ForeColor = Color.White
                Case "黑色"
                    myLabel.ForeColor = Color.White
            End Select
            myLabel.BackColor = Color.Transparent

            AddHandler myLabel.Click, AddressOf myLabel_Yanyuan_Click

            myLabel.Margin = New Padding(3, 3, 3, 0)
            FlowLayoutPanel6.Controls.Add(myLabel)
        Next

        '加载标签
        FlowLayoutPanel5.Controls.Clear()
        'Label11.Text = biaoqian
        'Label11.AutoSize = True
        'Label11.Visible = True

        '标签增加号
        Dim myjia As New Label
        myjia.AutoSize = True
        myjia.Text = "+"
        myjia.Cursor = Windows.Forms.Cursors.Hand

        Select Case ColorZhuti
            Case "默认"
                myjia.BackColor = Color.Transparent
            Case "灰色"
                myjia.BackColor = Color.Transparent
            Case "黑色"
                myjia.BackColor = Color.White
        End Select

        myjia.Font = New System.Drawing.Font("Times New Roman", 12)
        FlowLayoutPanel5.Controls.Add(myjia)
        AddHandler myjia.Click, AddressOf myjia_Click
        For Each s In Split(biaoqian, " ")
            Dim myLabel As New Label
            myLabel.Text = Juncode(s)
            myLabel.AutoSize = True
            myLabel.Cursor = Windows.Forms.Cursors.Hand

            If BigpicBiaoqian.myFontString(4) = "True" And BigpicBiaoqian.myFontString(5) = "True" Then
                myLabel.Font = New System.Drawing.Font(BigpicBiaoqian.myFontString(0), Int(BigpicBiaoqian.myFontString(1)), FontStyle.Bold Or FontStyle.Italic)
            ElseIf BigpicBiaoqian.myFontString(4) = "True" And BigpicBiaoqian.myFontString(5) <> "True" Then
                myLabel.Font = New System.Drawing.Font(BigpicBiaoqian.myFontString(0), Int(BigpicBiaoqian.myFontString(1)), FontStyle.Bold)
            ElseIf BigpicBiaoqian.myFontString(4) <> "True" And BigpicBiaoqian.myFontString(5) = "True" Then
                myLabel.Font = New System.Drawing.Font(BigpicBiaoqian.myFontString(0), Int(BigpicBiaoqian.myFontString(1)), FontStyle.Italic)
            ElseIf BigpicBiaoqian.myFontString(4) <> "True" And BigpicBiaoqian.myFontString(5) <> "True" Then
                myLabel.Font = New System.Drawing.Font(BigpicBiaoqian.myFontString(0), Int(BigpicBiaoqian.myFontString(1)))
            End If
            'myLabel.ForeColor = StringToColor(BigpicBiaoqian.myFontString(2))
            'myLabel.BackColor = StringToColor(BigpicBiaoqian.myFontString(3))

            Select Case ColorZhuti
                Case "默认"
                    myLabel.ForeColor = Color.Black
                Case "灰色"
                    myLabel.ForeColor = Color.White
                Case "黑色"
                    myLabel.ForeColor = Color.White
            End Select
            myLabel.BackColor = Color.Transparent

            AddHandler myLabel.Click, AddressOf myLabel_Biaoqian_Click

            myLabel.Margin = New Padding(3, 3, 3, 0)
            FlowLayoutPanel5.Controls.Add(myLabel)
        Next



        '设置字体、背景

        Label14.BackColor = Color.Transparent
        Label15.BackColor = Color.Transparent
        Label16.BackColor = Color.Transparent
        Label17.BackColor = Color.Transparent
        Label18.BackColor = Color.Transparent
        Label19.BackColor = Color.Transparent
        Label21.BackColor = Color.Transparent

        Select Case ColorZhuti
            Case "默认"
                Label14.ForeColor = Color.Black
                Label15.ForeColor = Color.Black
                Label16.ForeColor = Color.Black
                Label17.ForeColor = Color.Black
                Label18.ForeColor = Color.Black
                Label19.ForeColor = Color.Black
                Label21.ForeColor = Color.Black
            Case "灰色"
                Label14.ForeColor = Color.White
                Label15.ForeColor = Color.White
                Label16.ForeColor = Color.White
                Label17.ForeColor = Color.White
                Label18.ForeColor = Color.White
                Label19.ForeColor = Color.White
                Label21.ForeColor = Color.White
            Case "黑色"
                Label14.ForeColor = Color.White
                Label15.ForeColor = Color.White
                Label16.ForeColor = Color.White
                Label17.ForeColor = Color.White
                Label18.ForeColor = Color.White
                Label19.ForeColor = Color.White
                Label21.ForeColor = Color.White
        End Select

        If BigpicMingcheng.myFontString(4) = "True" And BigpicMingcheng.myFontString(5) = "True" Then
            TextBox1.Font = New System.Drawing.Font(BigpicMingcheng.myFontString(0), Int(BigpicMingcheng.myFontString(1)), FontStyle.Bold Or FontStyle.Italic)
        ElseIf BigpicMingcheng.myFontString(4) = "True" And BigpicMingcheng.myFontString(5) <> "True" Then
            TextBox1.Font = New System.Drawing.Font(BigpicMingcheng.myFontString(0), Int(BigpicMingcheng.myFontString(1)), FontStyle.Bold)
        ElseIf BigpicMingcheng.myFontString(4) <> "True" And BigpicMingcheng.myFontString(5) = "True" Then
            TextBox1.Font = New System.Drawing.Font(BigpicMingcheng.myFontString(0), Int(BigpicMingcheng.myFontString(1)), FontStyle.Italic)
        ElseIf BigpicMingcheng.myFontString(4) <> "True" And BigpicMingcheng.myFontString(5) <> "True" Then
            TextBox1.Font = New System.Drawing.Font(BigpicMingcheng.myFontString(0), Int(BigpicMingcheng.myFontString(1)))
        End If
        'TextBox1.ForeColor = StringToColor(BigpicMingcheng.myFontString(2))
        'TextBox1.BackColor = StringToColor(BigpicMingcheng.myFontString(3))


        If BigpicFanhao.myFontString(4) = "True" And BigpicFanhao.myFontString(5) = "True" Then
            Label14.Font = New System.Drawing.Font(BigpicFanhao.myFontString(0), Int(BigpicFanhao.myFontString(1)), FontStyle.Bold Or FontStyle.Italic)
        ElseIf BigpicFanhao.myFontString(4) = "True" And BigpicFanhao.myFontString(5) <> "True" Then
            Label14.Font = New System.Drawing.Font(BigpicFanhao.myFontString(0), Int(BigpicFanhao.myFontString(1)), FontStyle.Bold)
        ElseIf BigpicFanhao.myFontString(4) <> "True" And BigpicFanhao.myFontString(5) = "True" Then
            Label14.Font = New System.Drawing.Font(BigpicFanhao.myFontString(0), Int(BigpicFanhao.myFontString(1)), FontStyle.Italic)
        ElseIf BigpicFanhao.myFontString(4) <> "True" And BigpicFanhao.myFontString(5) <> "True" Then
            Label14.Font = New System.Drawing.Font(BigpicFanhao.myFontString(0), Int(BigpicFanhao.myFontString(1)))
        End If
        'Label14.ForeColor = StringToColor(BigpicFanhao.myFontString(2))
        'Label14.BackColor = StringToColor(BigpicFanhao.myFontString(3))

        If BigpicFaxingriqi.myFontString(4) = "True" And BigpicFaxingriqi.myFontString(5) = "True" Then
            Label15.Font = New System.Drawing.Font(BigpicFaxingriqi.myFontString(0), Int(BigpicFaxingriqi.myFontString(1)), FontStyle.Bold Or FontStyle.Italic)
        ElseIf BigpicFaxingriqi.myFontString(4) = "True" And BigpicFaxingriqi.myFontString(5) <> "True" Then
            Label15.Font = New System.Drawing.Font(BigpicFaxingriqi.myFontString(0), Int(BigpicFaxingriqi.myFontString(1)), FontStyle.Bold)
        ElseIf BigpicFaxingriqi.myFontString(4) <> "True" And BigpicFaxingriqi.myFontString(5) = "True" Then
            Label15.Font = New System.Drawing.Font(BigpicFaxingriqi.myFontString(0), Int(BigpicFaxingriqi.myFontString(1)), FontStyle.Italic)
        ElseIf BigpicFaxingriqi.myFontString(4) <> "True" And BigpicFaxingriqi.myFontString(5) <> "True" Then
            Label15.Font = New System.Drawing.Font(BigpicFaxingriqi.myFontString(0), Int(BigpicFaxingriqi.myFontString(1)))
        End If
        'Label15.ForeColor = StringToColor(BigpicFaxingriqi.myFontString(2))
        'Label15.BackColor = StringToColor(BigpicFaxingriqi.myFontString(3))

        If BigpicChangdu.myFontString(4) = "True" And BigpicChangdu.myFontString(5) = "True" Then
            Label16.Font = New System.Drawing.Font(BigpicChangdu.myFontString(0), Int(BigpicChangdu.myFontString(1)), FontStyle.Bold Or FontStyle.Italic)
        ElseIf BigpicChangdu.myFontString(4) = "True" And BigpicChangdu.myFontString(5) <> "True" Then
            Label16.Font = New System.Drawing.Font(BigpicChangdu.myFontString(0), Int(BigpicChangdu.myFontString(1)), FontStyle.Bold)
        ElseIf BigpicChangdu.myFontString(4) <> "True" And BigpicChangdu.myFontString(5) = "True" Then
            Label16.Font = New System.Drawing.Font(BigpicChangdu.myFontString(0), Int(BigpicChangdu.myFontString(1)), FontStyle.Italic)
        ElseIf BigpicChangdu.myFontString(4) <> "True" And BigpicChangdu.myFontString(5) <> "True" Then
            Label16.Font = New System.Drawing.Font(BigpicChangdu.myFontString(0), Int(BigpicChangdu.myFontString(1)))
        End If
        'Label16.ForeColor = StringToColor(BigpicChangdu.myFontString(2))
        'Label16.BackColor = StringToColor(BigpicChangdu.myFontString(3))

        If BigpicDaoyan.myFontString(4) = "True" And BigpicDaoyan.myFontString(5) = "True" Then
            Label17.Font = New System.Drawing.Font(BigpicDaoyan.myFontString(0), Int(BigpicDaoyan.myFontString(1)), FontStyle.Bold Or FontStyle.Italic)
        ElseIf BigpicDaoyan.myFontString(4) = "True" And BigpicDaoyan.myFontString(5) <> "True" Then
            Label17.Font = New System.Drawing.Font(BigpicDaoyan.myFontString(0), Int(BigpicDaoyan.myFontString(1)), FontStyle.Bold)
        ElseIf BigpicDaoyan.myFontString(4) <> "True" And BigpicDaoyan.myFontString(5) = "True" Then
            Label17.Font = New System.Drawing.Font(BigpicDaoyan.myFontString(0), Int(BigpicDaoyan.myFontString(1)), FontStyle.Italic)
        ElseIf BigpicDaoyan.myFontString(4) <> "True" And BigpicDaoyan.myFontString(5) <> "True" Then
            Label17.Font = New System.Drawing.Font(BigpicDaoyan.myFontString(0), Int(BigpicDaoyan.myFontString(1)))
        End If
        'Label17.ForeColor = StringToColor(BigpicDaoyan.myFontString(2))
        'Label17.BackColor = StringToColor(BigpicDaoyan.myFontString(3))

        If BigpicZhizuoshang.myFontString(4) = "True" And BigpicZhizuoshang.myFontString(5) = "True" Then
            Label18.Font = New System.Drawing.Font(BigpicZhizuoshang.myFontString(0), Int(BigpicZhizuoshang.myFontString(1)), FontStyle.Bold Or FontStyle.Italic)
        ElseIf BigpicZhizuoshang.myFontString(4) = "True" And BigpicZhizuoshang.myFontString(5) <> "True" Then
            Label18.Font = New System.Drawing.Font(BigpicZhizuoshang.myFontString(0), Int(BigpicZhizuoshang.myFontString(1)), FontStyle.Bold)
        ElseIf BigpicZhizuoshang.myFontString(4) <> "True" And BigpicZhizuoshang.myFontString(5) = "True" Then
            Label18.Font = New System.Drawing.Font(BigpicZhizuoshang.myFontString(0), Int(BigpicZhizuoshang.myFontString(1)), FontStyle.Italic)
        ElseIf BigpicZhizuoshang.myFontString(4) <> "True" And BigpicZhizuoshang.myFontString(5) <> "True" Then
            Label18.Font = New System.Drawing.Font(BigpicZhizuoshang.myFontString(0), Int(BigpicZhizuoshang.myFontString(1)))
        End If
        'Label18.ForeColor = StringToColor(BigpicZhizuoshang.myFontString(2))
        'Label18.BackColor = StringToColor(BigpicZhizuoshang.myFontString(3))

        If BigpicFaxingshang.myFontString(4) = "True" And BigpicFaxingshang.myFontString(5) = "True" Then
            Label19.Font = New System.Drawing.Font(BigpicFaxingshang.myFontString(0), Int(BigpicFaxingshang.myFontString(1)), FontStyle.Bold Or FontStyle.Italic)
        ElseIf BigpicFaxingshang.myFontString(4) = "True" And BigpicFaxingshang.myFontString(5) <> "True" Then
            Label19.Font = New System.Drawing.Font(BigpicFaxingshang.myFontString(0), Int(BigpicFaxingshang.myFontString(1)), FontStyle.Bold)
        ElseIf BigpicFaxingshang.myFontString(4) <> "True" And BigpicFaxingshang.myFontString(5) = "True" Then
            Label19.Font = New System.Drawing.Font(BigpicFaxingshang.myFontString(0), Int(BigpicFaxingshang.myFontString(1)), FontStyle.Italic)
        ElseIf BigpicFaxingshang.myFontString(4) <> "True" And BigpicFaxingshang.myFontString(5) <> "True" Then
            Label19.Font = New System.Drawing.Font(BigpicFaxingshang.myFontString(0), Int(BigpicFaxingshang.myFontString(1)))
        End If
        'Label19.ForeColor = StringToColor(BigpicFaxingshang.myFontString(2))
        'Label19.BackColor = StringToColor(BigpicFaxingshang.myFontString(3))


        If BigpicXilie.myFontString(4) = "True" And BigpicXilie.myFontString(5) = "True" Then
            Label21.Font = New System.Drawing.Font(BigpicXilie.myFontString(0), Int(BigpicXilie.myFontString(1)), FontStyle.Bold Or FontStyle.Italic)
        ElseIf BigpicXilie.myFontString(4) = "True" And BigpicXilie.myFontString(5) <> "True" Then
            Label21.Font = New System.Drawing.Font(BigpicXilie.myFontString(0), Int(BigpicXilie.myFontString(1)), FontStyle.Bold)
        ElseIf BigpicXilie.myFontString(4) <> "True" And BigpicXilie.myFontString(5) = "True" Then
            Label21.Font = New System.Drawing.Font(BigpicXilie.myFontString(0), Int(BigpicXilie.myFontString(1)), FontStyle.Italic)
        ElseIf BigpicXilie.myFontString(4) <> "True" And BigpicXilie.myFontString(5) <> "True" Then
            Label21.Font = New System.Drawing.Font(BigpicXilie.myFontString(0), Int(BigpicXilie.myFontString(1)))
        End If
        'Label21.ForeColor = StringToColor(BigpicXilie.myFontString(2))
        'Label21.BackColor = StringToColor(BigpicXilie.myFontString(3))
        FlowLayoutPanel1.Focus()


        '访问次数+1
        Dim myDataAdapter As New OleDbDataAdapter()
        con.ConnectionString = con_ConnectionString
        con.Open()
        cmd.Connection = con '初始化OLEDB命令的连接属性为con
        fangwencishu = fangwencishu + 1
        cmd.CommandText = "update Main set fangwencishu=" & fangwencishu & " where fanhao='" & GlobalFanhao & "'"
        myDataAdapter.UpdateCommand = cmd
        cmd.ExecuteNonQuery()
        con.Close()






    End Sub

    Private Sub myjia_Click()
        Dim str1
        str1 = InputBox("新的标签")
        If str1 = "" Then Exit Sub
        If InStr(str1, " ") > 0 Then
            MsgBox("不允许有空格！")
            Exit Sub
        End If
        For Each item In FlowLayoutPanel5.Controls
            If item.text = str1 Then
                MsgBox("不允许重复！")
                Exit Sub
            End If
        Next
        Dim con As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dr As OleDbDataReader
        Dim mybiaoqian As String
        mybiaoqian = ""
        Try
            con.ConnectionString = con_ConnectionString
            con.Open()
            cmd.Connection = con '初始化OLEDB命令的连接属性为con
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try
        cmd.CommandText = "select biaoqian from Main where fanhao='" & GlobalFanhao & "'"
        dr = cmd.ExecuteReader() '执行OLEDB命令以ExecuteReader()方式，并返回一个OLEDBReader，赋值给dr
        While (dr.Read)
            mybiaoqian = dr.Item("biaoqian").ToString
        End While
        dr.Close()
        Debug.Print(mybiaoqian)
        Dim myDataAdapter As New OleDbDataAdapter()
        If str1 <> "" Then
            mybiaoqian = mybiaoqian & " " & str1
            cmd.Connection = con '初始化OLEDB命令的连接属性为con
            cmd.CommandText = "update Main set biaoqian='" & mybiaoqian & "' where fanhao='" & GlobalFanhao & "'"
            myDataAdapter.UpdateCommand = cmd
        End If
        cmd.ExecuteNonQuery()
        con.Close()
        '载入标签

        FlowLayoutPanel5.Controls.Clear()
        Dim myjia As New Label
        myjia.AutoSize = True
        myjia.Text = "+"
        myjia.Cursor = Windows.Forms.Cursors.Hand
        Select Case ColorZhuti
            Case "默认"
                myjia.BackColor = Color.Transparent
            Case "灰色"
                myjia.BackColor = Color.Transparent
            Case "黑色"
                myjia.BackColor = Color.White
        End Select

        myjia.Font = New System.Drawing.Font("Times New Roman", 12)
        FlowLayoutPanel5.Controls.Add(myjia)
        AddHandler myjia.Click, AddressOf myjia_Click
        For Each s In Split(mybiaoqian, " ")
            Dim myLabel As New Label
            myLabel.Text = Juncode(s)
            myLabel.AutoSize = True
            myLabel.Cursor = Windows.Forms.Cursors.Hand

            If BigpicBiaoqian.myFontString(4) = "True" And BigpicBiaoqian.myFontString(5) = "True" Then
                myLabel.Font = New System.Drawing.Font(BigpicBiaoqian.myFontString(0), Int(BigpicBiaoqian.myFontString(1)), FontStyle.Bold Or FontStyle.Italic)
            ElseIf BigpicBiaoqian.myFontString(4) = "True" And BigpicBiaoqian.myFontString(5) <> "True" Then
                myLabel.Font = New System.Drawing.Font(BigpicBiaoqian.myFontString(0), Int(BigpicBiaoqian.myFontString(1)), FontStyle.Bold)
            ElseIf BigpicBiaoqian.myFontString(4) <> "True" And BigpicBiaoqian.myFontString(5) = "True" Then
                myLabel.Font = New System.Drawing.Font(BigpicBiaoqian.myFontString(0), Int(BigpicBiaoqian.myFontString(1)), FontStyle.Italic)
            ElseIf BigpicBiaoqian.myFontString(4) <> "True" And BigpicBiaoqian.myFontString(5) <> "True" Then
                myLabel.Font = New System.Drawing.Font(BigpicBiaoqian.myFontString(0), Int(BigpicBiaoqian.myFontString(1)))
            End If
            'myLabel.ForeColor = StringToColor(BigpicBiaoqian.myFontString(2))
            'myLabel.BackColor = StringToColor(BigpicBiaoqian.myFontString(3))

            Select Case ColorZhuti
                Case "默认"
                    myLabel.ForeColor = Color.Black
                Case "灰色"
                    myLabel.ForeColor = Color.White
                Case "黑色"
                    myLabel.ForeColor = Color.White
            End Select
            myLabel.BackColor = Color.Transparent

            AddHandler myLabel.Click, AddressOf myLabel_Biaoqian_Click

            myLabel.Margin = New Padding(3, 3, 3, 0)
            FlowLayoutPanel5.Controls.Add(myLabel)
        Next

    End Sub

    Private Sub myPicBox_MouseDown(sender As Object, e As MouseEventArgs)
        myPicboxName = Mid(sender.name, 10, Len(sender.name))
        If e.Button = MouseButtons.Left Then
            ReDim myBitmap(0)
            ExtraPicIndex = 0
            Dim picnum As String
            picnum = 0
            For Each picb In FlowLayoutPanel4.Controls
                Dim mypicb As PictureBox
                mypicb = picb
                myBitmap(picnum) = mypicb.Image
                If mypicb.Name = sender.name Then
                    ExtraPicIndex = picnum
                End If
                picnum = picnum + 1
                ReDim Preserve myBitmap(picnum)
            Next
            ReDim Preserve myBitmap(picnum - 1)
            MyPicForm1.Show()
            MyPicForm1.PictureBox1.Image = myBitmap(ExtraPicIndex)
        End If
    End Sub

    Private Sub BigPic_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '设置主题
        Select Case ColorZhuti
            Case "默认"
                Me.BackColor = Color.FromArgb(255, 240, 240, 240)
                FlowLayoutPanel1.BackColor = Color.FromArgb(255, 240, 240, 240)
                ToolStrip1.BackColor = Color.FromArgb(255, 240, 240, 240)
                TextBox1.BackColor = Color.FromArgb(255, 240, 240, 240)
                TextBox1.ForeColor = Color.Black
                ToolStripButton1.ForeColor = Color.Black
                ToolStripButton5.ForeColor = Color.Black
                ToolStripLabel1.ForeColor = Color.Black

                Label1.ForeColor = Color.Black
                Label2.ForeColor = Color.Black
                Label3.ForeColor = Color.Black
                Label4.ForeColor = Color.Black
                Label5.ForeColor = Color.Black
                Label6.ForeColor = Color.Black
                Label7.ForeColor = Color.Black
                Label8.ForeColor = Color.Black
                Label12.ForeColor = Color.Black
                Label20.ForeColor = Color.Black

            Case "灰色"
                Me.BackColor = Color.DimGray
                FlowLayoutPanel1.BackColor = Color.DimGray
                ToolStrip1.BackColor = Color.DimGray
                TextBox1.BackColor = Color.DimGray
                TextBox1.ForeColor = Color.White
                ToolStripButton1.ForeColor = Color.White
                ToolStripButton5.ForeColor = Color.White
                ToolStripLabel1.ForeColor = Color.White

                Label1.ForeColor = Color.Black
                Label2.ForeColor = Color.Black
                Label3.ForeColor = Color.Black
                Label4.ForeColor = Color.Black
                Label5.ForeColor = Color.Black
                Label6.ForeColor = Color.Black
                Label7.ForeColor = Color.Black
                Label8.ForeColor = Color.Black
                Label12.ForeColor = Color.Black
                Label20.ForeColor = Color.Black

            Case "黑色"
                Me.BackColor = Color.Black
                FlowLayoutPanel1.BackColor = Color.Black
                ToolStrip1.BackColor = Color.Black
                TextBox1.BackColor = Color.Black
                TextBox1.ForeColor = Color.White
                ToolStripButton1.ForeColor = Color.White
                ToolStripButton5.ForeColor = Color.White
                ToolStripLabel1.ForeColor = Color.White

                Label1.ForeColor = Color.White
                Label2.ForeColor = Color.White
                Label3.ForeColor = Color.White
                Label4.ForeColor = Color.White
                Label5.ForeColor = Color.White
                Label6.ForeColor = Color.White
                Label7.ForeColor = Color.White
                Label8.ForeColor = Color.White
                Label12.ForeColor = Color.White
                Label20.ForeColor = Color.White


        End Select
    End Sub

    Private Sub ToolStripButton1_Click_1(sender As Object, e As EventArgs)
        Debug.Print(ColorToString(Me.BackColor))
    End Sub

    Private Sub Label13_Click(sender As Object, e As EventArgs) Handles Label13.Click
        Debug.Print(YingpianWeizhi)
        Shell("explorer " & YingpianWeizhi)

    End Sub

    Private Sub Label13_MouseMove(sender As Object, e As MouseEventArgs) Handles Label13.MouseMove
        Panel1.BackColor = Color.DarkOrange
    End Sub

    Private Sub Label13_MouseLeave(sender As Object, e As EventArgs) Handles Label13.MouseLeave
        Panel1.BackColor = Color.Bisque
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        Shell("explorer " & YingpianWeizhi)
    End Sub

    Private Sub PictureBox3_MouseMove(sender As Object, e As MouseEventArgs) Handles PictureBox3.MouseMove
        Panel1.BackColor = Color.DarkOrange
    End Sub

    Private Sub PictureBox3_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox3.MouseLeave
        Panel1.BackColor = Color.Bisque
    End Sub

    Private Sub 打开其网址ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 打开其网址ToolStripMenuItem.Click
        Process.Start("http://" & JavWebSite & "/" & GlobalFanhao)
    End Sub

    Private Sub 打开照片位置ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 打开照片位置ToolStripMenuItem.Click
        '判断文件夹是否存在
        Dim PicPath As String = BigPicSavePath
        If System.IO.Directory.Exists(PicPath) = True Then
            'Process.Start(Application.StartupPath & "\pic\" & GlobalFanhao)
            Process.Start("Explorer.exe", "/select, """ & PicPath & "\" & GlobalFanhao & ".jpg" & """")
        End If
    End Sub

    Private Sub 打开影片位置ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 打开影片位置ToolStripMenuItem.Click
        If Dir(YingpianWeizhi) <> "" Then
            Process.Start("Explorer.exe", "/select, """ & YingpianWeizhi & """")
        End If
    End Sub

    Private Sub Label17_Click(sender As Object, e As EventArgs) Handles Label17.Click
        Clickindex = 15
        myDaoyan = Jencode(sender.text)
        Form1.SelectFromDatabase("where daoyan like '%" & myDaoyan & "%'", "Main", 1)
        Me.Close()
        Form1.Activate()
    End Sub

    Private Sub Label18_Click(sender As Object, e As EventArgs) Handles Label18.Click
        Clickindex = 14
        myZhizuoshang = Jencode(sender.text)
        Form1.SelectFromDatabase("where zhizuoshang like '%" & myZhizuoshang & "%'", "Main", 1)
        Me.Close()
        Form1.Activate()
    End Sub

    Private Sub Label19_Click(sender As Object, e As EventArgs) Handles Label19.Click
        Clickindex = 13
        myFaxingshang = Jencode(sender.text)
        Form1.SelectFromDatabase("where faxingshang like '%" & myFaxingshang & "%'", "Main", 1)
        Me.Close()
        Form1.Activate()
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        MyPicForm1.Show()
        MyPicForm1.PictureBox1.Image = PictureBox1.Image

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Label21_Click(sender As Object, e As EventArgs) Handles Label21.Click
        Clickindex = 16
        myXilie = Jencode(sender.text)
        Form1.SelectFromDatabase("where xilie like '%" & myXilie & "%'", "Main", 1)
        Me.Close()
        Form1.Activate()
    End Sub

    Private Sub ToolStripButton1_Click_2(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        MultiThreadDownLoadExtraPicByFanhao(GlobalFanhao)
    End Sub

    Private Sub BigPic_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        'e.Cancel = -1
        'Me.Hide()
    End Sub

    Private Sub 打开照片位置ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles 打开照片位置ToolStripMenuItem1.Click
        Dim PicPath As String = ExtraPicSavePath
        If System.IO.Directory.Exists(PicPath) = True Then
            'Process.Start(Application.StartupPath & "\pic\" & GlobalFanhao)
            'Process.Start(PicPath & "\" & GlobalFanhao)

            Process.Start("Explorer.exe", "/select, """ & PicPath & "\" & GlobalFanhao & "\" & myPicboxName & """")
            'Debug.Print(PicPath & "\" & GlobalFanhao & "\" & myPicboxName & ".jpg")
        End If
    End Sub

    Private Sub ToolStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles ToolStrip1.ItemClicked

    End Sub

    Private Sub 刷新ToolStripMenuItem_Click(sender As Object, e As EventArgs)
        ToolStripButton5_Click(sender, e)
    End Sub

    Private Sub 下载额外图片ToolStripMenuItem_Click(sender As Object, e As EventArgs)
        ToolStripButton1_Click_2(sender, e)
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim max As Double
        max = UBound(MyExtraPicThread)
        If max = 0 Then Exit Sub
        Dim num As Double
        num = 0
        For i = 0 To UBound(MyExtraPicThread) Step 1
            If MyExtraPicThread(i) IsNot Nothing Then
                If MyExtraPicThread(i).IsAlive = False Then
                    num = num + 1
                End If
            End If
        Next
        If num / max >= 1 Then
            ToolStripProgressBar1.Value = ToolStripProgressBar1.Maximum
            ToolStripLabel1.Text = "100%"
            Timer1.Enabled = False
        Else
            ToolStripProgressBar1.Value = ToolStripProgressBar1.Maximum * (num / max)
            ToolStripLabel1.Text = 100 * Math.Round(num / max, 3) & "%"
        End If
    End Sub

    Private Sub Label11_Click(sender As Object, e As EventArgs) Handles Label11.Click

    End Sub
End Class