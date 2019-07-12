Imports System.IO
Imports System.Net
Imports System.Text.RegularExpressions
Imports System.Data.OleDb
Module myOwnClass
    '输入网址、下载地址、正则表达式
    Public Class ExtraPicDownClass
        Public FilePath As String
        Public DownLoadUrl As String

        Public Sub ExtraPicDownload()
            Dim filename As String()
            filename = Split(DownLoadUrl, "/")
            Dim DownLoadFilePath As String
            DownLoadFilePath = FilePath & "\" & filename(UBound(filename))
            Debug.Print(DownLoadFilePath)
            'System.Net.ServicePointManager.DefaultConnectionLimit = 50     'keepalive的http连接增加到50
            Dim SPosition As Long = 0 '当前文件流的位置
            Dim FStream As FileStream '本地文件流
            Dim myRequest As HttpWebRequest 'Http
            Dim myStream As Stream '服务器文件流
            Dim btContent(512) As Byte '缓冲区
            Dim FileSize As Int64 '文件大小
            Dim CurrentSize As Int64 '当前文件大小
            Dim intSize As Integer = 0 '每次写入多少字节



            '首先获得文件大小
            myRequest = CType(HttpWebRequest.Create(DownLoadUrl), HttpWebRequest) '建立http连接

            '错误：超时
            Try
                FileSize = myRequest.GetResponse.Headers.Get("Content-Length") '获得文件头中文件大小
                myRequest.KeepAlive = False
                myRequest.GetResponse.Close() '关闭该http连接

            Catch ex As Exception
                Debug.Print(ex.Message)
            End Try
            'Try
            If File.Exists(DownLoadFilePath) Then '如果文件已存在则判断其大小是否等于头文件大小
                If My.Computer.FileSystem.GetFileInfo(DownLoadFilePath).Length < FileSize Then
                    IO.File.Delete(DownLoadFilePath) '删除文件
                Else
                    Exit Sub
                End If
            End If

            If Not File.Exists(DownLoadFilePath) Then
                FStream = New FileStream(DownLoadFilePath, FileMode.Create) '文件不存在则新建一个文件
                SPosition = 0
            End If



            ''首先获得文件大小
            'myRequest = CType(HttpWebRequest.Create(DownLoadUrl), HttpWebRequest) '建立http连接
            'FileSize = myRequest.GetResponse.Headers.Get("Content-Length") '获得文件头中文件大小
            'myRequest.KeepAlive = False
            'myRequest.GetResponse.Close() '关闭该http连接
            ''Try

            'If File.Exists(DownLoadFilePath) Then '如果文件已存在则跳过

            '    Exit Sub
            'Else
            '    FStream = New FileStream(DownLoadFilePath, FileMode.Create)  '文件不存在则新建一个文件
            '    SPosition = 0
            'End If
            'Catch ex As Exception
            '    Debug.Print(ex.Message)
            'End Try
            CurrentSize = SPosition '当前文件大小
            If CurrentSize >= FileSize Then FStream.Close() : Exit Sub '如果文件大小一致则不下载
            myRequest = CType(HttpWebRequest.Create(DownLoadUrl), HttpWebRequest) '建立http连接
            myRequest.KeepAlive = True '持续连接
            myRequest.AddRange(CurrentSize, FileSize) '向服务器获得指定区间大小的文件流
            myStream = myRequest.GetResponse.GetResponseStream '获得返回的文件流
            intSize = myStream.Read(btContent, 0, 512) '把文件流读到缓冲区里，然后流的位置自动增加512
            While intSize > 0 'And StopDown = False
                FStream.Write(btContent, 0, intSize) '将缓冲区里开头为0，长度为intSize的流写入文件
                CurrentSize = CurrentSize + intSize '当前文件大小
                Application.DoEvents() '刷新界面，防止卡死
                intSize = myStream.Read(btContent, 0, 512) '把文件流读到缓冲区里，然后流的位置自动增加512
            End While
            FStream.Close() '关闭流文件
            myStream.Close() '关闭http返回的流文件
            myRequest.GetResponse.Close() '断开Http连接
        End Sub

    End Class


    '输入网址、下载地址、正则表达式
    Public Class GetInfoClass
        Public getWebResponse_fanhao As String
        Public GetInfoCompleted As Boolean
        Public mingcheng As String
        Public changdu As Integer
        Public faxingriqi As String
        Public daoyan As String
        Public zhizuoshang As String
        Public faxingshang As String
        Public xilie As String
        Public leibie(0) As String
        Public yanyuan(0) As String
        Sub getWebResponseAndGetInfo()


            GetInfoCompleted = False
            'On Error Resume Next
            Dim webRequest As HttpWebRequest
            Dim mySearchUrl As String
            mySearchUrl = "https://" & JavWebSite & "/" & getWebResponse_fanhao
            Debug.Print(mySearchUrl)
            webRequest = CType(Net.WebRequest.Create(mySearchUrl), HttpWebRequest)
            'Dim myResponseStream As respon
            Dim responseReader As StreamReader
            Dim responseData As String
            Try
                responseReader = New StreamReader(webRequest.GetResponse().GetResponseStream())
                responseData = responseReader.ReadToEnd()
                responseReader.Close()
                webRequest.GetResponse.Close()
            Catch ex As Exception
                responseData = ""
                Debug.Print("获得网页源码出错")
            End Try
            Dim mc As MatchCollection
            Dim mc2 As MatchCollection
            'Debug.Print("网页源码为：" & responseData.ToString)
            '用正则表达式爬取
            '判断下载小图还是大图
            '发行日期长度，名称
            Dim str As String
            '<meta name="description" content="【發行日期】2012-11-30，【長度】29分鐘，(n0802)アナル処女子宮処女W強奪">
            mc = Regex.Matches(responseData, "<meta name=""description"".+>")
            If mc.Count > 0 Then
                mc2 = Regex.Matches(mc(0).ToString, "【發行日期】.+，")
                faxingriqi = Regex.Matches(mc2(0).ToString, "\d\d\d\d-\d\d-\d\d")(0).ToString
                mc2 = Regex.Matches(mc(0).ToString, "【長度】\d+分鐘")
                changdu = Int(Regex.Matches(mc2(0).ToString, "\d+")(0).ToString)
                str = Split(mc(0).ToString, "，")(2)
                mingcheng = Mid(str, 1, Len(str) - 2)
                mc2 = Regex.Matches(mc(0).ToString, "【長度】\d+分鐘")
                changdu = Regex.Matches(mc2(0).ToString, "\d+")(0).ToString
            Else
                faxingriqi = ""
                changdu = 0
                mingcheng = ""
            End If

            '<span class="header">導演:</span> <a href="https://www.dmmsee.net/studio/ok">キチックス/妄想族</a>'
            mc = Regex.Matches(responseData, "<span class=""header"">導演.+</a>")
            If mc.Count > 0 Then
                str = Split(mc(0).ToString, """")(4)
                mc2 = Regex.Matches(str, ">.+<")
                str = mc2(0).ToString
                daoyan = Mid(str, 2, Len(str) - 2)
            Else
                daoyan = ""
            End If
            '<span class="header">製作商:</span> <a href="https://www.dmmsee.net/studio/ok">キチックス/妄想族</a>
            mc = Regex.Matches(responseData, "<span class=""header"">製作商.+</a>")
            If mc.Count > 0 Then
                str = Split(mc(0).ToString, """")(4)
                mc2 = Regex.Matches(str, ">.+<")
                str = mc2(0).ToString
                zhizuoshang = Mid(str, 2, Len(str) - 2)
            Else
                zhizuoshang = ""
            End If

            ' <span Class="header">發行商:</span> <a href="https://www.dmmsee.net/studio/ok">キチックス/妄想族</a>
            mc = Regex.Matches(responseData, "<span class=""header"">發行商.+</a>")
            If mc.Count > 0 Then
                str = Split(mc(0).ToString, """")(4)
                mc2 = Regex.Matches(str, ">.+<")
                str = mc2(0).ToString
                faxingshang = Mid(str, 2, Len(str) - 2)
            Else
                faxingshang = ""
            End If

            ' <span class="header">系列:</span> <a href="https://www.dmmsee.net/studio/ok">キチックス/妄想族</a>
            mc = Regex.Matches(responseData, "<span class=""header"">系列.+</a>")
            If mc.Count > 0 Then
                str = Split(mc(0).ToString, """")(4)
                mc2 = Regex.Matches(str, ">.+<")
                str = mc2(0).ToString
                xilie = Mid(str, 2, Len(str) - 2)
            Else
                xilie = ""
            End If


            ' 类别
            ' <span class="genre"><a href="https://www.dmmsee.net/genre/g">DMM獨家</a></span>
            mc = Regex.Matches(responseData, "<span class=""genre"">.+</span>")
            Dim tempstr() As String
            Dim num As Integer
            num = 0
            If mc.Count > 0 Then
                For Each m In mc
                    tempstr = Split(m.ToString, """")
                    str = tempstr(UBound(tempstr))
                    mc2 = Regex.Matches(str, ">.+</a>")
                    str = mc2(0).ToString
                    leibie(num) = Mid(str, 2, Len(str) - 5)
                    num = num + 1
                    ReDim Preserve leibie(num)
                Next
                ReDim Preserve leibie(num - 1)
            Else
                leibie(0) = ""
            End If


            ' 演员
            ' <div class="star-name"><a href="https://www.dmmsee.net/star/ukf" title="やしきれな">やしきれな</a></div>
            mc = Regex.Matches(responseData, "<div class=""star-name"".+</div>")
            num = 0
            If mc.Count > 0 Then
                For Each m In mc
                    tempstr = Split(m.ToString, """")
                    str = tempstr(UBound(tempstr))
                    mc2 = Regex.Matches(str, ">.+</a>")
                    str = mc2(0).ToString
                    yanyuan(num) = Mid(str, 2, Len(str) - 5)
                    num = num + 1
                    ReDim Preserve yanyuan(num)
                Next
                ReDim Preserve yanyuan(num - 1)
            Else
                yanyuan(0) = ""
            End If
            GetInfoCompleted = True
        End Sub
    End Class




    '输入网址、下载地址、正则表达式
    Public Class DownClass
        Public FilePath As String
        'Public ProB As ProgressBar
        Public getWebResponse_fanhao As String
        'Public mylistbox As ListBox
        Public myDownType As Boolean
        Public VedioType As Integer
        Public index As Integer
        Sub getWebResponseAndDownload()
            'On Error Resume Next
            Dim webRequest As HttpWebRequest
            Dim mySearchUrl As String

            '判断下载小图还是大图
            If myDownType = True Then
                'If getshipinLeixingFromDatebase(getWebResponse_fanhao) = 1 Then
                '    mySearchUrl = "https://" & JavWebSite & "/uncensored/search/" & getWebResponse_fanhao
                'ElseIf getshipinLeixingFromDatebase(getWebResponse_fanhao) = 2 Then
                '    mySearchUrl = "https://" & JavWebSite & "/search/" & getWebResponse_fanhao
                'Else
                '    mySearchUrl = "https://" & JavWebSite & "/search/" & getWebResponse_fanhao
                'End If

                If VedioType = 1 Then
                    mySearchUrl = "https://" & JavWebSite & "/uncensored/search/" & getWebResponse_fanhao
                ElseIf VedioType = 2 Then
                    mySearchUrl = "https://" & JavWebSite & "/search/" & getWebResponse_fanhao
                Else
                    mySearchUrl = "https://" & JavWebSite & "/search/" & getWebResponse_fanhao
                End If

            Else '下载大图模式
                mySearchUrl = "https://" & JavWebSite & "/" & getWebResponse_fanhao
            End If
            Debug.Print(mySearchUrl)
            webRequest = CType(Net.WebRequest.Create(mySearchUrl), HttpWebRequest)
            'Dim myResponseStream As respon
            Dim responseReader As StreamReader
            Dim responseData As String
            Try
                responseReader = New StreamReader(webRequest.GetResponse().GetResponseStream())
                responseData = responseReader.ReadToEnd()
                responseReader.Close()
                webRequest.GetResponse.Close()
            Catch ex As Exception
                responseData = ""
                Debug.Print("获得网页源码出错")
            End Try
            Dim mc As MatchCollection
            'Debug.Print("网页源码为：" & responseData.ToString)
            '用正则表达式爬取
            '判断下载小图还是大图
            If myDownType = True Then
                'If getshipinLeixingFromDatebase(getWebResponse_fanhao) = 1 Then '步兵
                '    mc = Regex.Matches(responseData, "https://images\.[a-z]+\.[a-z]+/thumbs/\S+\.jpg")
                'Else
                '    mc = Regex.Matches(responseData, "https://pics\.[a-z]+\.[a-z]+/thumb/\S+\.jpg")
                'End If
                If VedioType = 1 Then '步兵
                    mc = Regex.Matches(responseData, "https://images\.[a-z]+\.[a-z]+/thumbs/\S+\.jpg")
                Else
                    mc = Regex.Matches(responseData, "https://pics\.[a-z]+\.[a-z]+/thumb/\S+\.jpg")
                End If



            Else '下载大图
                If VedioType = 1 Then '无码
                    mc = Regex.Matches(responseData, "https://images\.[a-z]+\.[a-z]+/cover/\S+\.jpg")
                Else
                    mc = Regex.Matches(responseData, "https://pics\.[a-z]+\.[a-z]+/cover/\S+\.jpg")
                End If


                'If getshipinLeixingFromDatebase(getWebResponse_fanhao) = 1 Then '无码
                '    mc = Regex.Matches(responseData, "https://images\.[a-z]+\.[a-z]+/cover/\S+\.jpg")
                'Else
                '    mc = Regex.Matches(responseData, "https://pics\.[a-z]+\.[a-z]+/cover/\S+\.jpg")
                'End If
            End If

            If mc.Count = 0 Then
                Debug.Print("正则表达无结果")
                '线程完成了
                myThreadIsCompleted(index) = True
                'Debug.Print("线程" & index & "完成")
                Exit Sub
            End If
            Dim url As String = mc(0).ToString
            'mylistbox.Items.Add(url)
            '结束爬取图片地址----------------------------
            Dim filename As String()
            filename = Split(url, ".")
            If myDownType = True Then
                FilePath = SmallPicSavePath & "\" & getWebResponse_fanhao & "." & filename(UBound(filename))
            Else
                FilePath = BigPicSavePath & "\" & getWebResponse_fanhao & "." & filename(UBound(filename))
            End If



            'System.Net.ServicePointManager.DefaultConnectionLimit = 50     'keepalive的http连接增加到50
            Dim SPosition As Long = 0 '当前文件流的位置
            Dim FStream As FileStream '本地文件流
            Dim myRequest As HttpWebRequest 'Http
            Dim myStream As Stream '服务器文件流
            Dim btContent(512) As Byte '缓冲区
            Dim FileSize As Int64 '文件大小
            Dim CurrentSize As Int64 '当前文件大小
            Dim intSize As Integer = 0 '每次写入多少字节

            '首先获得文件大小
            myRequest = CType(HttpWebRequest.Create(url), HttpWebRequest) '建立http连接

            '错误：超时
            Try
                FileSize = myRequest.GetResponse.Headers.Get("Content-Length") '获得文件头中文件大小
                myRequest.KeepAlive = False
                myRequest.GetResponse.Close() '关闭该http连接

            Catch ex As Exception
                Debug.Print(ex.Message)
            End Try
            'Try
            If File.Exists(FilePath) Then '如果文件已存在则判断其大小是否等于头文件大小
                If My.Computer.FileSystem.GetFileInfo(FilePath).Length < FileSize Then
                    IO.File.Delete(FilePath) '删除文件
                Else
                    Exit Sub
                End If
            End If

            If Not File.Exists(FilePath) Then
                FStream = New FileStream(FilePath, FileMode.Create) '文件不存在则新建一个文件
                SPosition = 0
            End If
            'Catch ex As Exception
            '    Debug.Print(ex.Message)
            'End Try
            CurrentSize = SPosition '当前文件大小
            If CurrentSize >= FileSize Then FStream.Close() : Exit Sub '如果文件大小一致则不下载

            Try


                myRequest = CType(HttpWebRequest.Create(url), HttpWebRequest) '建立http连接
                myRequest.KeepAlive = True '持续连接
                'myRequest.Timeout = 1000
                myRequest.AddRange(CurrentSize, FileSize) '向服务器获得指定区间大小的文件流
                '错误：这里会弹出操作超时
                myStream = myRequest.GetResponse.GetResponseStream '获得返回的文件流
                intSize = myStream.Read(btContent, 0, 512) '把文件流读到缓冲区里，然后流的位置自动增加512
                While intSize > 0 'And StopDown = False
                    FStream.Write(btContent, 0, intSize) '将缓冲区里开头为0，长度为intSize的流写入文件
                    CurrentSize = CurrentSize + intSize '当前文件大小
                    '显示进度
                    'ProB.Value = IIf(CurrentSize / FileSize < 1, CurrentSize / FileSize  ProB.Maximum, ProB.Maximum)
                    Application.DoEvents() '刷新界面，防止卡死

                    '错误：这里弹出请求被中止，请求已被取消
                    intSize = myStream.Read(btContent, 0, 512) '把文件流读到缓冲区里，然后流的位置自动增加512
                End While
                FStream.Close() '关闭流文件
                myStream.Close() '关闭http返回的流文件
                myRequest.GetResponse.Close() '断开Http连接
            Catch ex As Exception
                Debug.Print("myRequest下载不成功！")
            End Try
            '线程完成了
            'System.Threading.Thread.Sleep(500)
        End Sub
    End Class





    Public Class smallFrame
        Public Panel As New FlowLayoutPanel
        Public NameText As String
        Public FanhaoText As String
        Public RiqiText As String
        Public Pic As Image
        Public CMS As ContextMenuStrip
        Public CMS2 As ContextMenuStrip
        Public CMSLove As ContextMenuStrip
        Public PicBox As New PictureBox
        Public IsLove As Boolean
        Public LovePicbox As New PictureBox
        Public addToListPicbox As New PictureBox

        Dim myTextbox1 As New TextBox
        Dim myTextbox2 As New TextBox
        Dim myTextbox3 As New TextBox
        Dim label1 As New Label
        Dim myPlayButton As New PictureBox
        Sub initial()
            PicBox.Image = Pic
            PicBox.Width = 136
            PicBox.Height = 194
            'PicBox.ErrorImage = Image.FromFile(Application.StartupPath & "\icon\noimageps.gif")
            If IsSmallPicAutoSize = True Then
                PicBox.SizeMode = PictureBoxSizeMode.AutoSize
            Else
                PicBox.SizeMode = PictureBoxSizeMode.CenterImage
            End If
            PicBox.Cursor = Windows.Forms.Cursors.Hand
            PicBox.Location = New Point(8, 8)
            PicBox.ContextMenuStrip = CMS
            AddHandler PicBox.MouseUp, AddressOf Picbox_MouseUp
            PicBox.Name = "PicBox_" & FanhaoText


            If SmallPicMingcheng.myFontString(4) = "True" And SmallPicMingcheng.myFontString(5) = "True" Then
                myTextbox1.Font = New System.Drawing.Font(SmallPicMingcheng.myFontString(0), Int(SmallPicMingcheng.myFontString(1)), FontStyle.Bold Or FontStyle.Italic)
            ElseIf SmallPicMingcheng.myFontString(4) = "True" And SmallPicMingcheng.myFontString(5) <> "True" Then
                myTextbox1.Font = New System.Drawing.Font(SmallPicMingcheng.myFontString(0), Int(SmallPicMingcheng.myFontString(1)), FontStyle.Bold)
            ElseIf SmallPicMingcheng.myFontString(4) <> "True" And SmallPicMingcheng.myFontString(5) = "True" Then
                myTextbox1.Font = New System.Drawing.Font(SmallPicMingcheng.myFontString(0), Int(SmallPicMingcheng.myFontString(1)), FontStyle.Italic)
            ElseIf SmallPicMingcheng.myFontString(4) <> "True" And SmallPicMingcheng.myFontString(5) <> "True" Then
                myTextbox1.Font = New System.Drawing.Font(SmallPicMingcheng.myFontString(0), Int(SmallPicMingcheng.myFontString(1)))
            End If
            myTextbox1.ForeColor = StringToColor(SmallPicMingcheng.myFontString(2))
            myTextbox1.BackColor = StringToColor(SmallPicMingcheng.myFontString(3))
            myTextbox1.BorderStyle = BorderStyle.None
            myTextbox1.ReadOnly = True
            myTextbox1.Multiline = True
            myTextbox1.Text = NameText
            myTextbox1.Width = PicBox.Width
            label1.Font = myTextbox1.Font
            label1.Text = myTextbox1.Text
            label1.AutoSize = True
            myTextbox1.Height = label1.Height * (myTextbox1.GetLineFromCharIndex(myTextbox1.TextLength) + 1)

            If SmallPicFanhao.myFontString(4) = "True" And SmallPicFanhao.myFontString(5) = "True" Then
                myTextbox2.Font = New System.Drawing.Font(SmallPicFanhao.myFontString(0), Int(SmallPicFanhao.myFontString(1)), FontStyle.Bold Or FontStyle.Italic)
            ElseIf SmallPicFanhao.myFontString(4) = "True" And SmallPicFanhao.myFontString(5) <> "True" Then
                myTextbox2.Font = New System.Drawing.Font(SmallPicFanhao.myFontString(0), Int(SmallPicFanhao.myFontString(1)), FontStyle.Bold)
            ElseIf SmallPicFanhao.myFontString(4) <> "True" And SmallPicFanhao.myFontString(5) = "True" Then
                myTextbox2.Font = New System.Drawing.Font(SmallPicFanhao.myFontString(0), Int(SmallPicFanhao.myFontString(1)), FontStyle.Italic)
            ElseIf SmallPicFanhao.myFontString(4) <> "True" And SmallPicFanhao.myFontString(5) <> "True" Then
                myTextbox2.Font = New System.Drawing.Font(SmallPicFanhao.myFontString(0), Int(SmallPicFanhao.myFontString(1)))
            End If
            myTextbox2.ForeColor = StringToColor(SmallPicFanhao.myFontString(2))
            myTextbox2.BackColor = StringToColor(SmallPicFanhao.myFontString(3))
            myTextbox2.BorderStyle = BorderStyle.None
            myTextbox2.ReadOnly = True
            myTextbox2.Multiline = True
            myTextbox2.Text = FanhaoText
            myTextbox2.Width = PicBox.Width
            label1.Font = myTextbox2.Font
            label1.Text = myTextbox2.Text
            label1.AutoSize = True
            '错误收集：
            '描述，当数据很多的时候点击所有视频会出错
            myTextbox2.Height = label1.Height * (myTextbox2.GetLineFromCharIndex(myTextbox2.TextLength) + 1)

            If SmallPicRiqi.myFontString(4) = "True" And SmallPicRiqi.myFontString(5) = "True" Then
                myTextbox3.Font = New System.Drawing.Font(SmallPicRiqi.myFontString(0), Int(SmallPicRiqi.myFontString(1)), FontStyle.Bold Or FontStyle.Italic)
            ElseIf SmallPicRiqi.myFontString(4) = "True" And SmallPicRiqi.myFontString(5) <> "True" Then
                myTextbox3.Font = New System.Drawing.Font(SmallPicRiqi.myFontString(0), Int(SmallPicRiqi.myFontString(1)), FontStyle.Bold)
            ElseIf SmallPicRiqi.myFontString(4) <> "True" And SmallPicRiqi.myFontString(5) = "True" Then
                myTextbox3.Font = New System.Drawing.Font(SmallPicRiqi.myFontString(0), Int(SmallPicRiqi.myFontString(1)), FontStyle.Italic)
            ElseIf SmallPicRiqi.myFontString(4) <> "True" And SmallPicRiqi.myFontString(5) <> "True" Then
                myTextbox3.Font = New System.Drawing.Font(SmallPicRiqi.myFontString(0), Int(SmallPicRiqi.myFontString(1)))
            End If
            myTextbox3.ForeColor = StringToColor(SmallPicRiqi.myFontString(2))
            myTextbox3.BackColor = StringToColor(SmallPicRiqi.myFontString(3))
            myTextbox3.BorderStyle = BorderStyle.None
            myTextbox3.ReadOnly = True
            myTextbox3.Multiline = True
            myTextbox3.Text = RiqiText
            myTextbox3.Width = PicBox.Width
            label1.Font = myTextbox3.Font
            label1.Text = myTextbox3.Text
            label1.AutoSize = True
            myTextbox3.Height = label1.Height * (myTextbox3.GetLineFromCharIndex(myTextbox3.TextLength) + 1)



            If IsLove = True Then
                LovePicbox.Image = My.Resources.Resource1.xihuan_fill
                LovePicbox.Name = "LovePicbox_Love_" & FanhaoText
            Else
                LovePicbox.Image = My.Resources.Resource1.xihuan
                LovePicbox.Name = "LovePicbox_NotLove_" & FanhaoText
            End If
            LovePicbox.SizeMode = PictureBoxSizeMode.AutoSize
            LovePicbox.Cursor = Windows.Forms.Cursors.Hand
            AddHandler LovePicbox.MouseDown, AddressOf LovePicbox_MouseDown

            addToListPicbox.Image = My.Resources.Resource1.nav_list
            addToListPicbox.SizeMode = PictureBoxSizeMode.AutoSize
            addToListPicbox.Cursor = Windows.Forms.Cursors.Hand
            addToListPicbox.Name = "addToListPicbox_" & FanhaoText
            AddHandler addToListPicbox.MouseDown, AddressOf addToListPicbox_MouseDown


            myPlayButton.Image = My.Resources.Resource1.bofang
            myPlayButton.SizeMode = PictureBoxSizeMode.AutoSize
            myPlayButton.Cursor = Windows.Forms.Cursors.Hand
            myPlayButton.Name = "myPlayButton_" & FanhaoText
            AddHandler myPlayButton.MouseDown, AddressOf myPlayButton_MouseDown


            Select Case ColorZhuti
                Case "默认"
                    Panel.BackColor = Color.White
                    label1.BackColor = Color.White
                    myTextbox1.BackColor = Color.White
                    myTextbox2.BackColor = Color.White
                    myTextbox3.BackColor = Color.White
                    LovePicbox.BackColor = Color.White
                    addToListPicbox.BackColor = Color.White

                    myTextbox1.ForeColor = Color.Black
                    myTextbox2.ForeColor = Color.Black
                    myTextbox3.ForeColor = Color.Black
                Case "灰色"
                    Panel.BackColor = Color.FromArgb(64, 64, 64)
                    label1.BackColor = Color.FromArgb(64, 64, 64)
                    myTextbox1.BackColor = Color.FromArgb(64, 64, 64)
                    myTextbox2.BackColor = Color.FromArgb(64, 64, 64)
                    myTextbox3.BackColor = Color.FromArgb(64, 64, 64)
                    LovePicbox.BackColor = Color.White
                    addToListPicbox.BackColor = Color.White

                    myTextbox1.ForeColor = Color.White
                    myTextbox2.ForeColor = Color.White
                    myTextbox3.ForeColor = Color.White
                Case "黑色"
                    Panel.BackColor = Color.Black
                    label1.BackColor = Color.Black
                    myTextbox1.BackColor = Color.Black
                    myTextbox2.BackColor = Color.Black
                    myTextbox3.BackColor = Color.Black
                    LovePicbox.BackColor = Color.White
                    addToListPicbox.BackColor = Color.White
                    myTextbox1.ForeColor = Color.White
                    myTextbox2.ForeColor = Color.White
                    myTextbox3.ForeColor = Color.White

            End Select


            Panel.Padding = New Padding(5, 5, 5, 5)
            Panel.Controls.Add(PicBox)
            Panel.Name = "Panel_" & FanhaoText
            Panel.Cursor = Windows.Forms.Cursors.Hand

            Panel.Controls.Add(myTextbox1)
            Panel.Controls.Add(myTextbox2)
            Panel.Controls.Add(myTextbox3)
            Panel.Controls.Add(LovePicbox)
            Panel.Controls.Add(addToListPicbox)
            Panel.Controls.Add(myPlayButton)
            Panel.SetFlowBreak(PicBox, True)
            Panel.SetFlowBreak(myTextbox1, True)
            Panel.SetFlowBreak(myTextbox2, True)
            Panel.SetFlowBreak(myTextbox3, True)
            If Not IsMingchengShow Then
                myTextbox1.Visible = False
            End If
            If Not IsFanhaoShow Then
                myTextbox2.Visible = False
            End If
            If Not IsRiqiShow Then
                myTextbox3.Visible = False
            End If
            If Not IsShoucangShow Then
                LovePicbox.Visible = False
                addToListPicbox.Visible = False
            End If

            Panel.Width = PicBox.Width + 16
            'Panel.Height = PicBox.Height + myTextbox1.Height + myTextbox2.Height + myTextbox3.Height + addToListPicbox.Height
            Panel.Height = addToListPicbox.Top + addToListPicbox.Height + 10

            AddHandler Panel.MouseUp, AddressOf Panel_MouseUp
        End Sub


        Sub Panel_MouseUp(sender As Object, e As MouseEventArgs)
            GlobalFanhao = Mid(sender.name, 7)
            'Dim myFlowLayoutPanel As FlowLayoutPanel
            'myFlowLayoutPanel = sender.getparent
            'MsgBox(myFlowLayoutPanel.Name)
            If e.Button = MouseButtons.Left Then
                BigPic.Show()
                BigPic.ToolStripButton5_Click(sender, e)
                BigPic.Activate()
            End If
        End Sub

        Sub Picbox_MouseUp(sender As Object, e As MouseEventArgs)
            GlobalFanhao = Mid(sender.name, 8)
            'Dim myFlowLayoutPanel As FlowLayoutPanel
            'myFlowLayoutPanel = sender.getparent
            'MsgBox(myFlowLayoutPanel.Name)
            If e.Button = MouseButtons.Left Then
                If IsClickSmalPicShowBigPic = False Then
                    BigPic.Show()
                    BigPic.ToolStripButton5_Click(sender, e)
                    BigPic.Activate()
                Else
                    MyPicForm1.Show()
                    Dim PicPath As String
                    PicPath = BigPicSavePath & "\" & GlobalFanhao & ".jpg"
                    If Dir(PicPath) <> "" Then
                        If My.Computer.FileSystem.GetFileInfo(PicPath).Length > 0 Then '如果图片错误会显示内存不足
                            Try
                                Dim pFileStream As New FileStream(PicPath, FileMode.Open, FileAccess.Read)
                                MyPicForm1.PictureBox1.Image = Image.FromStream(pFileStream)
                                pFileStream.Close()
                                pFileStream.Dispose()
                            Catch ex As Exception
                                MsgBox(ex.Message)
                            End Try
                        End If
                    End If
                End If
            End If
        End Sub

        Sub LovePicbox_MouseDown(sender As Object, e As MouseEventArgs)
            Dim PicB As PictureBox
            PicB = sender

            If InStr(PicB.Name, "LovePicbox_Love") > 0 Then
                GlobalFanhao = Mid(sender.name, 17)
            ElseIf InStr(PicB.Name, "LovePicbox_NotLove") > 0 Then
                GlobalFanhao = Mid(sender.name, 20)
            End If
            Dim p As New Point()
            p.X = System.Windows.Forms.Control.MousePosition.X
            p.Y = System.Windows.Forms.Control.MousePosition.Y
            CMSLove.Show(p)

        End Sub

        Sub addToListPicbox_MouseDown(sender As Object, e As MouseEventArgs)
            GlobalFanhao = Mid(sender.name, 17)
            Dim p As New Point()
            p.X = System.Windows.Forms.Control.MousePosition.X
            p.Y = System.Windows.Forms.Control.MousePosition.Y
            CMS2.Show(p)
        End Sub

        Sub myPlayButton_MouseDown(sender As Object, e As MouseEventArgs)
            GlobalFanhao = Mid(sender.name, 14)
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

    End Class






    Public Class Panel_DownItem
        Public myPanel As New FlowLayoutPanel
        Public NameText As String
        Public Prob As New ProgressBar
        Dim Pic As New PictureBox
        Dim Name As New Label


        Sub Panel_DownItem_initial()
            myPanel.Height = 40
            myPanel.Width = 500
            'myPanel.BackColor = Color.Transparent
            myPanel.BackColor = Color.DarkGray
            myPanel.FlowDirection = FlowDirection.LeftToRight
            myPanel.Margin = New Padding(0, 0, 0, 0)
            Pic.Image = My.Resources.Resource1.图片
            Pic.SizeMode = PictureBoxSizeMode.AutoSize
            Pic.Margin = New Padding(5, 5, 3, 3)
            Name.Margin = New Padding(3, 10, 3, 0)
            'Name.BackColor = Color.Transparent
            Name.Text = NameText
            Name.AutoSize = True
            Name.Font = New System.Drawing.Font("微软雅黑", 10.5)
            Prob.Margin = New Padding(3, 10, 3, 0)
            Prob.Width = 154

            myPanel.Controls.Add(Pic)
            myPanel.Controls.Add(Name)
            myPanel.Controls.Add(Prob)
            myPanel.AutoSize = True
        End Sub





    End Class






    Public Class myFontClass
        Public myFontString(5) As String
    End Class


    Public Class CSysXML
        Dim mXmlDoc As New System.Xml.XmlDocument
        Public XmlFile As String

        Public Sub New(ByVal File As String)
            MyClass.XmlFile = File
            MyClass.mXmlDoc.Load(MyClass.XmlFile)       '加载配置文件  
        End Sub

        '功能：取得元素值  
        '参数：node--节点       element--元素名          
        '返回：元素值   字符型  
        '             $--表示出错误  
        Public Function GetElement(ByVal node As String, ByVal element As String) As String
            On Error GoTo Err
            Dim mXmlNode As System.Xml.XmlNode = mXmlDoc.SelectSingleNode("//" + node)

            '读数据  
            Dim xmlNode As System.Xml.XmlNode = mXmlNode.SelectSingleNode(element)
            Return xmlNode.InnerText.ToString
Err:
            Return "$"
        End Function
        '  
        '功能：保存元素值  
        '参数：node--节点名称     element--元素名       val--值  
        '返回：True--保存成功     False--保存失败  
        Public Function SaveElement(ByVal node As String, ByVal element As String, ByVal val As String) As Boolean
            On Error GoTo err
            Dim mXmlNode As System.Xml.XmlNode = mXmlDoc.SelectSingleNode("//" + node)
            Dim xmlNodeNew As System.Xml.XmlNode

            xmlNodeNew = mXmlNode.SelectSingleNode(element)
            xmlNodeNew.InnerText = val
            mXmlDoc.Save(MyClass.XmlFile)
            Return True
err:
            Return False
        End Function


        '  

        Public Function GetElement2(ByVal node As String, ByVal element As String, myFontArray() As String)
            On Error GoTo Err
            Dim mXmlNode As System.Xml.XmlNode = mXmlDoc.SelectSingleNode("//" + node)
            '读数据  
            Dim xmlNode As System.Xml.XmlNode = mXmlNode.SelectSingleNode(element)
            Dim n As Integer = 0
            For Each nodeTemp As System.Xml.XmlNode In xmlNode
                myFontArray(n) = nodeTemp.InnerText
                n = n + 1
            Next

Err:
            Return "$"
        End Function

        Public Function SaveElement2(ByVal node As String, ByVal element As String, myFontArray() As String) As Boolean
            On Error GoTo err
            Dim mXmlNode As System.Xml.XmlNode = mXmlDoc.SelectSingleNode("//" + node)
            Dim xmlNodeNew As System.Xml.XmlNode

            xmlNodeNew = mXmlNode.SelectSingleNode(element)
            Dim n As Integer = 0
            For Each nodeTemp As System.Xml.XmlNode In xmlNodeNew
                nodeTemp.InnerText = myFontArray(n)
                n = n + 1
            Next
            mXmlDoc.Save(MyClass.XmlFile)
            Return True
err:
            Return False
        End Function
    End Class



End Module
