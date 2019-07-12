Imports System.Text.RegularExpressions
Imports System.Net
Imports System.IO
Imports System.Text.Encoding
Imports System.Text
Imports System.Data.OleDb
Module mySub

    Public Sub ChangeFormColor()

    End Sub



    Public Function GetByDiv2(ByVal code As String, ByVal divBegin As String, ByVal divEnd As String)  '获取分隔符所夹的内容[完成，未测试]
        '仅用于获取编码数据
        Dim lgStart As Integer
        Dim lens As Integer
        Dim lgEnd As Integer
        lens = Len(divBegin)
        If InStr(1, code, divBegin) = 0 Then GetByDiv2 = "" : Exit Function
        lgStart = InStr(1, code, divBegin) + CInt(lens)

        lgEnd = InStr(lgStart + 1, code, divEnd)
        If lgEnd = 0 Then GetByDiv2 = "" : Exit Function
        GetByDiv2 = Mid(code, lgStart, lgEnd - lgStart)
    End Function

    Public Function GetWebCode(ByVal strURL As String) As String
        Dim httpReq As System.Net.HttpWebRequest
        Dim httpResp As System.Net.HttpWebResponse
        Dim httpURL As New System.Uri(strURL)
        Dim ioS As System.IO.Stream, charSet As String, tCode As String
        Dim k() As Byte
        ReDim k(0)
        Dim dataQue As New Queue(Of Byte)
        httpReq = CType(WebRequest.Create(httpURL), HttpWebRequest)
        Dim sTime As Date = CDate("1990-09-21 00:00:00")
        httpReq.IfModifiedSince = sTime
        httpReq.Method = "GET"
        httpReq.Timeout = 7000

        Try
            httpResp = CType(httpReq.GetResponse(), HttpWebResponse)
        Catch
            'Debug.Print("weberror")
            GetWebCode = "<title>no thing found</title>" : Exit Function
        End Try
        '以上为网络数据获取
        ioS = CType(httpResp.GetResponseStream, Stream)
        Do While ioS.CanRead = True
            Try
                dataQue.Enqueue(ioS.ReadByte)
            Catch
                'Debug.Print("read error")
                Exit Do
            End Try
        Loop
        ReDim k(dataQue.Count - 1)
        For j As Integer = 0 To dataQue.Count - 1
            k(j) = dataQue.Dequeue
        Next
        '以上，为获取流中的的二进制数据
        tCode = Encoding.GetEncoding("UTF-8").GetString(k) '获取特定编码下的情况，毕竟UTF-8支持英文正常的显示
        charSet = Replace(GetByDiv2(tCode, "charset=", """"), """", "") '进行编码类型识别
        '以上，获取编码类型
        If charSet = "" Then 'defalt
            If httpResp.CharacterSet = "" Then
                tCode = Encoding.GetEncoding("UTF-8").GetString(k)
            Else
                tCode = Encoding.GetEncoding(httpResp.CharacterSet).GetString(k)
            End If
        Else
            tCode = Encoding.GetEncoding(charSet).GetString(k)
        End If
        'Debug.Print(charSet)
        'Stop
        '以上，按照获得的编码类型进行数据转换
        '将得到的内容进行最后处理，比如判断是不是有出现字符串为空的情况
        GetWebCode = tCode
        If tCode = "" Then GetWebCode = "<title>no thing found</title>"
    End Function




    '错误收集




    Public Sub MultiThreadGetInfoByFanhao(fanhao As String)
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


        If fanhao <> "" Then
            Dim MyGetInfoClass As New GetInfoClass
            Dim MySingleThread As System.Threading.Thread
            MyGetInfoClass.getWebResponse_fanhao = fanhao

            MySingleThread = New System.Threading.Thread(AddressOf MyGetInfoClass.getWebResponseAndGetInfo)
            MySingleThread.Start()
            Do
                Application.DoEvents()
                If MyGetInfoClass.GetInfoCompleted = True Then
                    Dim con As New OleDbConnection
                    Dim cmd As New OleDbCommand
                    Dim myDataAdapter As New OleDbDataAdapter()
                    Dim sql As String
                    con.ConnectionString = con_ConnectionString
                    con.Open()
                    cmd.Connection = con '初始化OLEDB命令的连接属性为con
                    Dim str1, str2 As String
                    str1 = ""
                    str2 = ""
                    For Each s In MyGetInfoClass.leibie
                        str1 = str1 & s & " "
                    Next
                    For Each s In MyGetInfoClass.yanyuan
                        str2 = str2 & s & " "
                    Next
                    str1 = Replace(str1, "'", "’")
                    str2 = Replace(str2, "'", "’")

                    sql = "update Main set faxingriqi = '" & MyGetInfoClass.faxingriqi & "', changdu = '" & MyGetInfoClass.changdu & "', mingcheng = '" & Replace(Jencode(MyGetInfoClass.mingcheng), "'", "’") & "', daoyan = '" & Replace(Jencode(MyGetInfoClass.daoyan), "'", "’") & "', zhizuoshang = '" & Replace(Jencode(MyGetInfoClass.zhizuoshang), "'", "’") & "', faxingshang = '" & Replace(Jencode(MyGetInfoClass.faxingshang), "'", "’") & "', xilie = '" & Replace(Jencode(MyGetInfoClass.xilie), "'", "’") & "', leibie = '" & Jencode(str1) & "', yanyuan = '" & Jencode(str2) & "'"
                    cmd.CommandText = sql & " where fanhao = '" & fanhao & "'"
                    myDataAdapter.UpdateCommand = cmd
                    cmd.ExecuteNonQuery()
                    con.Close()


                    Exit Do
                End If
            Loop
        End If
    End Sub


    Public Sub MultiThreadDownLoadExtraPicByFanhao(fanhao As String)
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

        If fanhao <> "" Then
            Dim webRequest As HttpWebRequest
            Dim myUrl As String
            Dim url(0) As String
            myUrl = "https://" & JavWebSite & "/" & fanhao
            Debug.Print(myUrl)
            webRequest = CType(Net.WebRequest.Create(myUrl), HttpWebRequest)
            webRequest.Timeout = 3000
            Dim responseReader As StreamReader
            Dim responseData As String
            Try
                responseReader = New StreamReader(webRequest.GetResponse().GetResponseStream())
                responseData = responseReader.ReadToEnd()
                responseReader.Close()
                webRequest.GetResponse.Close()
            Catch ex As Exception
                responseData = ""
                MsgBox("获得网页源码出错")
                Exit Sub
            End Try
            Dim mc As MatchCollection
            'Debug.Print("网页源码为：" & responseData.ToString)
            '用正则表达式爬取
            'https://pics.dmm.co.jp/digital/video/hnd00669/hnd00669jp-1.jpg
            'mc = Regex.Matches(responseData, "https://pics.dmm.co.jp/digital/video/\S+/\S+.jpg")
            mc = Regex.Matches(responseData, "https://pics.dmm.co.jp/digital/video/\S+/\S+jp\S+.jpg")

            If mc.Count = 0 Then
                MsgBox("正则表达无结果")
                Exit Sub
            End If
            Dim num As Integer = 0
            For Each m In mc
                url(num) = m.ToString
                num = num + 1
                ReDim Preserve url(num)
            Next
            ReDim Preserve url(num - 1)
            Dim mypicpath As String
            mypicpath = (ExtraPicSavePath & "\" & fanhao)
            If System.IO.Directory.Exists(mypicpath) = False Then
                My.Computer.FileSystem.CreateDirectory(mypicpath)
            End If
            ReDim MyExtraPicThread(UBound(url))
            Dim u As String
            BigPic.Timer1.Enabled = True
            BigPic.ToolStripLabel1.Text = "0%"
            For i = 0 To UBound(url)
                u = url(i)
                Debug.Print(u)
                Dim MyDownClass As New ExtraPicDownClass
                MyDownClass.FilePath = mypicpath

                MyDownClass.DownLoadUrl = u
                MyExtraPicThread(i) = New System.Threading.Thread(AddressOf MyDownClass.ExtraPicDownload)
                MyExtraPicThread(i).Start()

            Next

        End If


    End Sub

    Public Sub MySleep2(ByVal sleep_ms As Integer)
        Dim tick As Integer = Environment.TickCount
        While Environment.TickCount - tick < sleep_ms
            Application.DoEvents()
        End While
    End Sub


    Public Sub MultiThreadDownLoadSmallPicByFanhao(fanhao As String, myDownType As Boolean, Optional VT As Integer = 0)
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



        If fanhao <> "" Then
            Dim MyDownClass As DownClass
            Dim MySingleThread As System.Threading.Thread
            Dim mypicpath As String
            MyDownClass = New DownClass
            MyDownClass.getWebResponse_fanhao = fanhao
            MyDownClass.myDownType = myDownType
            MyDownClass.VedioType = VT
            MyDownClass.index = myThreadNum
            If myDownType = True Then '小图
                mypicpath = SmallPicSavePath
            Else
                mypicpath = BigPicSavePath
            End If
            If System.IO.Directory.Exists(mypicpath) = False Then
                My.Computer.FileSystem.CreateDirectory(mypicpath)
            End If
            MyDownClass.FilePath = mypicpath
            MySingleThread = New System.Threading.Thread(AddressOf MyDownClass.getWebResponseAndDownload)
            MySingleThread.Start()

        End If
    End Sub


    Public Function getFileSize(str1 As String) As String
        getFileSize = ""
        If str1 = "" Then Exit Function
        Dim num As Double
        num = Val(str1)
        num = num / 1024 / 1024
        If num < 1024 And num > 0 Then
            getFileSize = FormatNumber(num, 2) & " Mb"
        ElseIf num >= 1024 Then
            getFileSize = FormatNumber(num / 1024, 2) & " Gb"
        End If
    End Function




    Public Function GetFileExtName(FilePathFileName As String) As String   '获取扩展名  .txt
        On Error Resume Next
        Dim i As Integer, j As Integer
        i = Len(FilePathFileName)
        j = InStrRev(FilePathFileName, ".")
        If j = 0 Then
            GetFileExtName = ".txt"
        Else
            GetFileExtName = Mid(FilePathFileName, j, i)
        End If
    End Function


    '====================获取路径名各部分:  如： c:\dir1001\aaa.txt

    Public Function GetFileName(FilePathFileName As String) As String   '获取文件名  aaa.txt
        On Error Resume Next
        Dim i As Integer, j As Integer
        i = Len(FilePathFileName)
        j = InStrRev(FilePathFileName, "\")
        GetFileName = Mid(FilePathFileName, j + 1, i)
    End Function

    ''===========获取路径路径   c:\dir1001\
    Public Function GetFilePath(FilePathFileName As String) As String '获取路径路径   c:\dir1001\
        On Error Resume Next
        Dim j As Integer
        j = InStrRev(FilePathFileName, "\")
        GetFilePath = Mid(FilePathFileName, 1, j)
    End Function

    '===========获取文件名但不包括扩展名  aaa
    Public Function GetFileNameNoExt(FilePathFileName As String) As String  '获取文件名但不包括扩展名  aaa
        On Error Resume Next
        Dim i As Integer, j As Integer, k As Integer
        i = Len(FilePathFileName)
        j = InStrRev(FilePathFileName, "\")
        k = InStrRev(FilePathFileName, ".")
        If k = 0 Then
            GetFileNameNoExt = Mid(FilePathFileName, j + 1, i - j)
        Else
            GetFileNameNoExt = Mid(FilePathFileName, j + 1, k - j - 1)
        End If

    End Function

    Public Function ColorToString(myColor As Color) As String
        ColorToString = myColor.A.ToString & " " & myColor.R.ToString & " " & myColor.G.ToString & " " & myColor.B.ToString & " "
    End Function

    Public Function StringToColor(myString As String) As Color
        Dim str1() As String
        str1 = Split(myString, " ")
        StringToColor = Color.FromArgb(str1(0), str1(1), str1(2), str1(3))
    End Function

    '从数据库中获取视频类型
    Public Function getshipinLeixingFromDatebase(fanhao As String) As Integer
        getshipinLeixingFromDatebase = 0
        Dim con As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dr As OleDbDataReader
        Dim sql As String
        con.ConnectionString = con_ConnectionString_read
        con.Open()
        cmd.Connection = con '初始化OLEDB命令的连接属性为con
        sql = "select shipinleixing from Main where fanhao='" & fanhao & "'"
        cmd.CommandText = sql
        dr = cmd.ExecuteReader() '执行OLEDB命令以ExecuteReader()方式，并返回一个OLEDBReader，赋值给dr
        While (dr.Read)
            getshipinLeixingFromDatebase = Int(dr(0).ToString)
        End While
        con.Close()
    End Function


    '判断是步兵还是骑兵
    Public Function getshipinLeixing(str1 As String) As Integer
        'Dim BubingFanhao As String
        Dim str2 As String
        str1 = UCase(str1)

        If InStr(str1, "S2M") > 0 Then 'S2M
            getshipinLeixing = 1
        ElseIf InStr(str1, "FC2PPV") > 0 Then 'FC2PPV
            getshipinLeixing = 1
        ElseIf InStr(str1, "GACHI") > 0 Then 'GACHI
            getshipinLeixing = 1
        ElseIf GetFanhaoByRegExp(str1, "(K|N)\d+") <> "" Then 'Tokyo-Hot
            getshipinLeixing = 1
        ElseIf GetFanhaoByRegExp(str1, "[A-Z][A-Z]+") <> "" Then
            str2 = GetFanhaoByRegExp(str1, "[A-Z]+")
            Dim isbubing As Boolean
            For Each s In Split(BubingFanhao, ",")
                isbubing = False
                If str2 = s Then
                    isbubing = True
                    Exit For
                End If
            Next

            If isbubing Then '步兵
                getshipinLeixing = 1
            Else '骑兵
                getshipinLeixing = 2
            End If
        ElseIf GetFanhaoByRegExp(str1, "\d{5,}(-|_)\d+") <> "" Then
            getshipinLeixing = 1

        Else '未识别番号
            getshipinLeixing = 0
        End If
    End Function




    Public Function GetFanhaoByRegExp(SourceName As String, myPattern As String) As String
        Dim mc As MatchCollection = Regex.Matches(SourceName, myPattern)
        If mc.Count > 0 Then
            GetFanhaoByRegExp = mc(0).Value
        Else
            GetFanhaoByRegExp = ""
        End If

    End Function

    '判断n0124的n前面是否有英文
    Public Function judgeyingwenExist(str1 As String, str2 As String) As Boolean
        Dim snum As Integer
        snum = InStr(1, str1, str2)
        If snum > 1 Then
            If (Mid(str1, snum - 1, 1) < "z" And Mid(str1, snum - 1, 1) > "a") Or (Mid(str1, snum - 1, 1) < "Z" And Mid(str1, snum - 1, 1) > "A") Then
                judgeyingwenExist = True
            Else
                judgeyingwenExist = False
            End If
        Else
            judgeyingwenExist = False
        End If
    End Function

    Public Function getFanhao(str1 As String) As String
        '骑兵的正则表达式
        '正则表达式：[A-Za-z]+-?\d\d+   识别到“任意长度英文-至少2个数字”其中-可有可无
        '正则表达式：T28,sqte
        getFanhao = ""
        If str1 = "" Then
            getFanhao = ""
            Exit Function
        End If
        Dim stt As String
        If InStr(1, UCase(str1), "T28") > 0 Then
            stt = UCase(GetFanhaoByRegExp(str1, "T28-?\d\d+"))
            If InStr(1, stt, "-") > 0 Then
                getFanhao = stt
            ElseIf Len(stt) > 0 Then
                getFanhao = "T28-" & Mid(stt, 4, Len(stt) - 1)
            End If
        ElseIf InStr(1, str1, "sqte") > 0 Then
            getFanhao = UCase(GetFanhaoByRegExp(str1, "sqte-?\d\d+"))
            '步兵的正则表达式
        ElseIf InStr(1, UCase(str1), "HEYZO") > 0 Then     'HEYZO
            stt = Replace(GetFanhaoByRegExp(LCase(str1), "heyzo\s?\)?\(?_?(hd|lt)?\+?-?_?\d\d+"), "hd", "")
            stt = Replace(stt, "lt", "")
            getFanhao = addGang(UCase((stt)))
        ElseIf InStr(1, UCase(str1), "HEYDOUGA") > 0 Then      'HEYDOUGA
            getFanhao = UCase(GetFanhaoByRegExp(LCase(str1), "heydouga_?-?\d\d+"))
        ElseIf InStr(1, UCase(str1), "FC2") > 0 Then      'FC2PPV
            stt = UCase(GetFanhaoByRegExp(LCase(str1), "\d{5,}"))
            If stt <> "" Then
                getFanhao = "FC2PPV-" & stt
            Else
                getFanhao = stt
            End If
        ElseIf InStr(1, UCase(str1), "MKD-S") > 0 Then      'MKD-S
            getFanhao = UCase(GetFanhaoByRegExp(LCase(str1), "mkd-s\d\d+"))
        ElseIf InStr(1, UCase(str1), "MKBD-S") > 0 Then      'MKBD-S
            getFanhao = UCase(GetFanhaoByRegExp(LCase(str1), "mkbd-s\d\d+"))
        ElseIf InStr(1, UCase(str1), "S2M") > 0 Then      's2m
            getFanhao = UCase(GetFanhaoByRegExp(LCase(str1), "s2m-?\d\d+"))
        ElseIf InStr(1, UCase(str1), "S2MBD") > 0 Then      's2mbd
            getFanhao = UCase(GetFanhaoByRegExp(LCase(str1), "s2mbd-?\d\d+"))
        ElseIf InStr(1, UCase(str1), "S2MCR") > 0 Then      's2mcr
            getFanhao = UCase(GetFanhaoByRegExp(LCase(str1), "s2mcr-?\d\d+"))
        ElseIf Len(GetFanhaoByRegExp(str1, "\d{5,}(_|-)\d\d+")) > 0 Then        '加勒比，缺点，如果有时间如2012-01就会提取出时间
            getFanhao = UCase(GetFanhaoByRegExp(str1, "\d{5,}(_|-)\d\d+"))
        ElseIf Len(GetFanhaoByRegExp(str1, "(k|n)_?\d\d+")) > 0 Then  'Tokyo-Hot
            stt = GetFanhaoByRegExp(str1, "(k|n)_?\d\d+")
            If judgeyingwenExist(str1, stt) = False Then
                getFanhao = UCase(stt)
            Else
                getFanhao = addGang(UCase(GetFanhaoByRegExp(str1, "[A-Za-z][A-Za-z]+-?\d\d+")))
            End If
        Else        '类骑兵
            getFanhao = addGang(UCase(GetFanhaoByRegExp(str1, "[A-Za-z][A-Za-z]+-?\d\d+"))) '缺点：匹配不到C-1023
        End If
    End Function

    Function addGang(tFh As String) As String '添加"-"
        'red番号在javbus格式如下
        '有码：red-000
        '无码：red000
        If InStr(1, tFh, "-") > 0 Then
            addGang = tFh
        ElseIf tFh <> "" Then
            addGang = GetFanhaoByRegExp(tFh, "[A-Z]+") & "-" & GetFanhaoByRegExp(tFh, "\d+")
        Else
            addGang = ""
        End If
    End Function
End Module
