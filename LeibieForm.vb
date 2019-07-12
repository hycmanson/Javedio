Imports System.ComponentModel
Imports System.Data.OleDb


Public Class LeibieForm
    Dim Leibie_zhuti As String
    Dim Leibie_juese As String
    Dim Leibie_fuzhuang As String
    Dim Leibie_tixing As String
    Dim Leibie_xingwei As String
    Dim Leibie_wanfa As String
    Dim Leibie_qita As String

    Dim Leibie_bubing_zhuti As String
    Dim Leibie_bubing_juese As String
    Dim Leibie_bubing_fuzhuang As String
    Dim Leibie_bubing_tixing As String
    Dim Leibie_bubing_xingwei As String
    Dim Leibie_bubing_wanfa As String
    Dim Leibie_bubing_qita As String
    Dim Leibie_bubing_changjing As String




    Private Sub LeibieForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '骑兵
        Leibie_zhuti = "折磨_嘔吐_觸手_蠻橫嬌羞_處男_正太控_出軌_瘙癢_運動_女同接吻_性感的_美容院_處女_爛醉如泥的_殘忍畫面_妄想_惡作劇_學校作品_粗暴_通姦_姐妹_雙性人_跳舞_性奴_倒追_性騷擾_其他_戀腿癖_偷窥_花癡_男同性恋_情侶_戀乳癖_亂倫_其他戀物癖_偶像藝人_野外・露出_獵豔_女同性戀_企畫_10枚組_性感的_性感的_科幻_女優ベスト・総集編_温泉_M男_原作コラボ_16時間以上作品_デカチン・巨根_ファン感謝・訪問_動画_巨尻_ハーレム_日焼け_早漏_キス・接吻_汗だく_スマホ専用縦動画_Vシネマ_Don Cipote's choice_アニメ_アクション_イメージビデオ（男性）_孕ませ_ボーイズラブ_ビッチ_特典あり（AVベースボール）_コミック雑誌_時間停止"
        Leibie_juese = "童年朋友_公主_亞洲女演員_伴侶_講師_婆婆_格鬥家_女檢察官_明星臉_女主人、女老板_模特兒_秘書_美少女_新娘、年輕妻子_姐姐_格鬥家_車掌小姐_寡婦_千金小姐_白人_已婚婦女_女醫生_各種職業_妓女_賽車女郎_女大學生_展場女孩_女教師_母親_家教_护士_蕩婦_黑人演員_女生_女主播_高中女生_服務生_魔法少女_學生（其他）_動畫人物_遊戲的真人版_超級女英雄"
        Leibie_fuzhuang = "角色扮演_制服_女戰士_及膝襪_娃娃_女忍者_女裝人妖_內衣_猥褻穿著_兔女郎_貓耳女_女祭司_泡泡襪_制服_緊身衣_裸體圍裙_迷你裙警察_空中小姐_連褲襪_身體意識_OL_和服・喪服_體育服_內衣_水手服_學校泳裝_旗袍_女傭_迷你裙_校服_泳裝_眼鏡_角色扮演_哥德蘿莉_和服・浴衣"
        Leibie_tixing = "超乳_肌肉_乳房_嬌小的_屁股_高_變性者_無毛_胖女人_苗條_孕婦_成熟的女人_蘿莉塔_貧乳・微乳_巨乳"
        Leibie_xingwei = "顏面騎乘_食糞_足交_母乳_手指插入_按摩_女上位_舔陰_拳交_深喉_69_淫語_潮吹_乳交_排便_飲尿_口交_濫交_放尿_打手槍_吞精_肛交_顏射_自慰_顏射_中出_肛内中出"
        Leibie_wanfa = "立即口交_女優按摩棒_子宮頸_催眠_乳液_羞恥_凌辱_拘束_輪姦_插入異物_鴨嘴_灌腸_監禁_紧缚_強姦_藥物_汽車性愛_SM_糞便_玩具_跳蛋_緊縛_按摩棒_多P_性愛_假陽具_逆強姦"
        Leibie_qita = "合作作品_恐怖_給女性觀眾_教學_DMM專屬_R-15_R-18_戲劇_3D_特效_故事集_限時降價_複刻版_戲劇_戀愛_高畫質_主觀視角_介紹影片_4小時以上作品_薄馬賽克_經典_首次亮相_數位馬賽克_投稿_纪录片_國外進口_第一人稱攝影_業餘_局部特寫_獨立製作_DMM獨家_單體作品_合集_高清_字幕_天堂TV_DVD多士爐_AV OPEN 2014 スーパーヘビー_AV OPEN 2014 ヘビー級_AV OPEN 2014 ミドル級_AV OPEN 2015 マニア/フェチ部門_AV OPEN 2015 熟女部門_AV OPEN 2015 企画部門_AV OPEN 2015 乙女部門_AV OPEN 2015 素人部門_AV OPEN 2015 SM/ハード部門_AV OPEN 2015 女優部門_AVOPEN2016人妻・熟女部門_AVOPEN2016企画部門_AVOPEN2016ハード部門_AVOPEN2016マニア・フェチ部門_AVOPEN2016乙女部門_AVOPEN2016女優部門_AVOPEN2016ドラマ・ドキュメンタリー部門_AVOPEN2016素人部門_AVOPEN2016バラエティ部門_VR専用"

        '步兵
        Leibie_bubing_zhuti = "高清_字幕_推薦作品_通姦_淋浴_舌頭_下流_敏感_變態_願望_慾求不滿_服侍_外遇_訪問_性伴侶_保守_購物_誘惑_出差_煩惱_主動_再會_戀物癖_問題_騙奸_鬼混_高手_順從_密會_做家務_秘密_送貨上門_壓力_處女作_淫語_問卷_住一宿_眼淚_跪求_求職_婚禮_第一視角_洗澡_首次_劇情_約會_實拍_同性戀_幻想_淫蕩_旅行_面試_喝酒_尖叫_新年_借款_不忠_檢查_羞恥_勾引_新人_推銷_ブルマ"
        Leibie_bubing_juese = "AV女優_情人_丈夫_辣妹_S級女優_白領_偶像_兒子_女僕_老師_夫婦_保健室_朋友_工作人員_明星_同事_面具男_上司_睡眠系_奶奶_播音員_鄰居_親人_店員_魔女_視訊小姐_大學生_寡婦_小姐_秘書_人妖_啦啦隊_美容師_岳母_警察_熟女_素人_人妻_痴女_角色扮演_蘿莉_姐姐_模特_教師_學生_少女_新手_男友_護士_媽媽_主婦_孕婦_女教師_年輕人妻_職員_看護_外觀相似_色狼_醫生_新婚_黑人_空中小姐"
        Leibie_bubing_fuzhuang = "制服_內衣_休閒裝_水手服_全裸_不穿內褲_和服_不戴胸罩_連衣裙_打底褲_緊身衣_客人_晚禮服_治癒系_大衣_裸體襪子_絲帶_睡衣_面具_牛仔褲_喪服_極小比基尼_混血_毛衣_頸鏈_短褲_美人_連褲襪_裙子_浴衣和服_泳衣_網襪_眼罩_圍裙_比基尼_情趣內衣_迷你裙_套裝_眼鏡_丁字褲_陽具腰帶"
        Leibie_bubing_tixing = "美肌_屁股_美穴_黑髮_嬌小_曬痕_F罩杯_E罩杯_D罩杯_素顏_貓眼_捲髮_虎牙_C罩杯_I罩杯_小麥色_大陰蒂_美乳_巨乳_豐滿_苗條_美臀_美腿_無毛_美白_微乳_性感_高個子_爆乳_G罩杯_多毛_巨臀_軟體_巨大陽具_長發_H罩杯"
        Leibie_bubing_xingwei = "舔陰_電動陽具_淫亂_射在外陰_猛烈_後入內射_足交_射在胸部_側位內射_射在腹部_騎乘內射_射在頭髮_母乳_站立姿勢_肛射_陰道擴張_內射觀察_射在大腿_精液流出_射在屁股_內射潮吹_首次肛交_射在衣服上_首次內射_早洩_翻白眼_舔腳_喝尿_口交_內射_自慰_後入_騎乘位_顏射_口內射精_手淫_潮吹_輪姦_亂交_乳交_小便_吸精_深膚色_指法_騎在臉上_連續內射_打樁機_肛交_吞精_鴨嘴_打飛機_剃毛_站立位_高潮_二穴同入_舔肛_多人口交_痙攣_玩弄肛門_立即口交_舔蛋蛋_口射_陰屁_失禁_大量潮吹_69"
        Leibie_bubing_wanfa = "振動_淫語_搭訕_奴役_打屁股_潤滑油_按摩_散步_扯破連褲襪_手銬_束縛_調教_假陽具_變態遊戲_注視_蠟燭_電鑽_亂搞_摩擦_項圈_繩子_灌腸_監禁_車震_鞭打_懸掛_喝口水_精液塗抹_舔耳朵_女體盛_便利店_插兩根_開口器_暴露_陰道放入食物_大便_經期_惡作劇_電動按摩器_凌辱_玩具_露出_肛門_拘束_多P_潤滑劑_攝影_野外_陰道觀察_SM_灌入精液_受虐_綁縛_偷拍_異物插入_電話_公寓_遠程操作_偷窺_踩踏"
        Leibie_bubing_qita = "企劃物_獨佔動畫_10代_1080p_人氣系列_60fps_超VIP_投稿_VIP_椅子_風格出眾_首次作品_更衣室_下午_KTV_白天_最佳合集"
        Leibie_bubing_changjing = "酒店_密室_車_床_陽台_公園_家中_公交車_公司_門口_附近_學校_辦公室_樓梯_住宅_公共廁所_旅館_教室_廚房_桌子_大街_農村_和室_地下室_牢籠_屋頂_游泳池_電梯_拍攝現場_別墅_房間_愛情旅館_車內_沙發_浴室_廁所_溫泉_醫院_榻榻米"

        ToolStripComboBox1.SelectedIndex = 1

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label1_MouseMove(sender As Object, e As MouseEventArgs)
        Label1.BackColor = Color.Red
        Label1.ForeColor = Color.White

    End Sub

    Private Sub Label1_MouseLeave(sender As Object, e As EventArgs)
        Label1.BackColor = Color.White
        Label1.ForeColor = Color.Black
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click

        FlowLayoutPanel1.Controls.Clear()
        FlowLayoutPanel3.Controls.Clear()
        FlowLayoutPanel4.Controls.Clear()
        FlowLayoutPanel5.Controls.Clear()
        FlowLayoutPanel6.Controls.Clear()
        FlowLayoutPanel7.Controls.Clear()
        FlowLayoutPanel8.Controls.Clear()
        FlowLayoutPanel9.Controls.Clear()

        FlowLayoutPanel1.AutoScroll = False
        FlowLayoutPanel3.AutoScroll = False
        FlowLayoutPanel4.AutoScroll = False
        FlowLayoutPanel5.AutoScroll = False
        FlowLayoutPanel6.AutoScroll = False
        FlowLayoutPanel7.AutoScroll = False
        FlowLayoutPanel8.AutoScroll = False
        FlowLayoutPanel9.AutoScroll = False

        Dim con As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dr As OleDbDataReader
        Dim sql As String
        Dim num As Integer
        con.ConnectionString = con_ConnectionString
        con.Open()
        cmd.Connection = con '初始化OLEDB命令的连接属性为con
        sql = "select leibie from Main where shipinleixing=1"
        Select Case ToolStripComboBox1.Text
            Case "步兵"
                sql = "select leibie from Main where shipinleixing=1"
            Case "骑兵"
                sql = "select leibie from Main where shipinleixing=2"
            Case "欧美"
                sql = "select leibie from Main where shipinleixing=3"
            Case "所有"
                sql = "select leibie from Main"
        End Select
        cmd.CommandText = sql
        dr = cmd.ExecuteReader() '执行OLEDB命令以ExecuteReader()方式，并返回一个OLEDBReader，赋值给dr
        num = 0
        Dim myLeibie(0) As String
        Dim IsLeibieExist As Boolean
        While (dr.Read)
            IsLeibieExist = False
            For Each s In Split(dr(0).ToString, " ")
                IsLeibieExist = False
                For Each Leibie In myLeibie
                    If Leibie = s Then IsLeibieExist = True
                Next
                If IsLeibieExist = False Then
                    myLeibie(num) = s
                    num = num + 1
                    ReDim Preserve myLeibie(num)
                End If
            Next
        End While
        dr.Close()
        con.Close()
        ReDim Preserve myLeibie(num - 1)

        If ToolStripComboBox1.Text = "所有" Then
            'FlowLayoutPanel1.Visible = False
            FlowLayoutPanel3.Visible = False
            FlowLayoutPanel4.Visible = False
            FlowLayoutPanel5.Visible = False
            FlowLayoutPanel6.Visible = False
            FlowLayoutPanel7.Visible = False
            FlowLayoutPanel8.Visible = False
            FlowLayoutPanel9.Visible = False

            Label2.Text = "所有类别"
            Label3.Visible = False
            Label5.Visible = False
            Label7.Visible = False
            Label9.Visible = False
            Label11.Visible = False
            Label13.Visible = False
            Label15.Visible = False

            For Each s In myLeibie
                'Debug.Print(s)
                Dim myLabel As New Label
                myLabel.Text = Juncode(s)

                myLabel.Width = 150
                myLabel.Height = 40
                myLabel.Margin = New Padding(0, 0, 0, 0)
                myLabel.TextAlign = ContentAlignment.MiddleCenter
                myLabel.AutoSize = False
                myLabel.Font = New System.Drawing.Font("微软雅黑", 14)
                myLabel.BackColor = Color.White
                myLabel.ForeColor = Color.Black
                myLabel.Cursor = Windows.Forms.Cursors.Hand
                AddHandler myLabel.MouseMove, AddressOf myLabel_MouseMove
                AddHandler myLabel.MouseLeave, AddressOf myLabel_MouseLeave
                AddHandler myLabel.Click, AddressOf myLabel_Click

                FlowLayoutPanel1.Controls.Add(myLabel)


            Next

        ElseIf ToolStripComboBox1.Text = "欧美" Then

        ElseIf ToolStripComboBox1.Text = "步兵" Then
            FlowLayoutPanel3.Visible = True
            FlowLayoutPanel4.Visible = True
            FlowLayoutPanel5.Visible = True
            FlowLayoutPanel6.Visible = True
            FlowLayoutPanel7.Visible = True
            FlowLayoutPanel8.Visible = True
            FlowLayoutPanel9.Visible = True

            Label2.Text = "主题"
            Label3.Visible = True
            Label5.Visible = True
            Label7.Visible = True
            Label9.Visible = True
            Label11.Visible = True
            Label13.Visible = True
            Label15.Visible = True

            For Each s In myLeibie
                'Debug.Print(s)
                Dim myLabel As New Label
                myLabel.Text = Juncode(s)

                myLabel.Width = 150
                myLabel.Height = 40
                myLabel.Margin = New Padding(0, 0, 0, 0)
                myLabel.TextAlign = ContentAlignment.MiddleCenter
                myLabel.AutoSize = False
                myLabel.Font = New System.Drawing.Font("微软雅黑", 14)
                myLabel.BackColor = Color.White
                myLabel.ForeColor = Color.Black
                myLabel.Cursor = Windows.Forms.Cursors.Hand
                AddHandler myLabel.MouseMove, AddressOf myLabel_MouseMove
                AddHandler myLabel.MouseLeave, AddressOf myLabel_MouseLeave
                AddHandler myLabel.Click, AddressOf myLabel_Click


                If InStr(Leibie_bubing_zhuti, Juncode(s)) > 0 Then
                    FlowLayoutPanel1.Controls.Add(myLabel)
                ElseIf InStr(Leibie_bubing_juese, Juncode(s)) > 0 Then
                    FlowLayoutPanel3.Controls.Add(myLabel)
                ElseIf InStr(Leibie_bubing_fuzhuang, Juncode(s)) > 0 Then
                    FlowLayoutPanel4.Controls.Add(myLabel)
                ElseIf InStr(Leibie_bubing_tixing, Juncode(s)) > 0 Then
                    FlowLayoutPanel5.Controls.Add(myLabel)
                ElseIf InStr(Leibie_bubing_xingwei, Juncode(s)) > 0 Then
                    FlowLayoutPanel6.Controls.Add(myLabel)
                ElseIf InStr(Leibie_bubing_wanfa, Juncode(s)) > 0 Then
                    FlowLayoutPanel7.Controls.Add(myLabel)
                ElseIf InStr(Leibie_bubing_changjing, Juncode(s)) > 0 Then
                    FlowLayoutPanel9.Controls.Add(myLabel)
                Else
                    FlowLayoutPanel8.Controls.Add(myLabel)
                End If

            Next

        ElseIf ToolStripComboBox1.Text = "骑兵" Then

            FlowLayoutPanel3.Visible = True
            FlowLayoutPanel4.Visible = True
            FlowLayoutPanel5.Visible = True
            FlowLayoutPanel6.Visible = True
            FlowLayoutPanel7.Visible = True
            FlowLayoutPanel8.Visible = True
            FlowLayoutPanel9.Visible = False

            Label2.Text = "主题"
            Label3.Visible = True
            Label5.Visible = True
            Label7.Visible = True
            Label9.Visible = True
            Label11.Visible = True
            Label13.Visible = True
            Label15.Visible = False

            For Each s In myLeibie
                'Debug.Print(s)
                Dim myLabel As New Label
                myLabel.Text = Juncode(s)

                myLabel.Width = 150
                myLabel.Height = 40
                myLabel.Margin = New Padding(0, 0, 0, 0)
                myLabel.TextAlign = ContentAlignment.MiddleCenter
                myLabel.AutoSize = False
                myLabel.Font = New System.Drawing.Font("微软雅黑", 14)
                myLabel.BackColor = Color.White
                myLabel.ForeColor = Color.Black
                myLabel.Cursor = Windows.Forms.Cursors.Hand
                AddHandler myLabel.MouseMove, AddressOf myLabel_MouseMove
                AddHandler myLabel.MouseLeave, AddressOf myLabel_MouseLeave
                AddHandler myLabel.Click, AddressOf myLabel_Click


                If InStr(Leibie_zhuti, Juncode(s)) > 0 Then
                    FlowLayoutPanel1.Controls.Add(myLabel)
                ElseIf InStr(Leibie_juese, Juncode(s)) > 0 Then
                    FlowLayoutPanel3.Controls.Add(myLabel)
                ElseIf InStr(Leibie_fuzhuang, Juncode(s)) > 0 Then
                    FlowLayoutPanel4.Controls.Add(myLabel)
                ElseIf InStr(Leibie_tixing, Juncode(s)) > 0 Then
                    FlowLayoutPanel5.Controls.Add(myLabel)
                ElseIf InStr(Leibie_xingwei, Juncode(s)) > 0 Then
                    FlowLayoutPanel6.Controls.Add(myLabel)
                ElseIf InStr(Leibie_wanfa, Juncode(s)) > 0 Then
                    FlowLayoutPanel7.Controls.Add(myLabel)
                Else
                    FlowLayoutPanel8.Controls.Add(myLabel)
                End If

            Next

        End If
    End Sub

    Private Sub myLabel_MouseMove(sender As Object, e As MouseEventArgs)
        Dim mylabel As Label
        mylabel = sender
        mylabel.BackColor = Color.Red
        mylabel.ForeColor = Color.White
    End Sub

    Private Sub myLabel_MouseLeave(sender As Object, e As EventArgs)
        Dim mylabel As Label
        mylabel = sender
        mylabel.BackColor = Color.White
        mylabel.ForeColor = Color.Black
    End Sub

    Private Sub myLabel_Click(sender As Object, e As EventArgs)
        Debug.Print(sender.text)
        Debug.Print(Jencode(sender.text))
        Clickindex = 10
        myLeibie = Jencode(sender.text)
        Try


            Select Case ToolStripComboBox1.Text
                Case "步兵"
                    Form1.SelectFromDatabase("where leibie like '%" & myLeibie & "%' and shipinleixing=1", "Main", 1)
                Case "骑兵"
                    Form1.SelectFromDatabase("where leibie like '%" & myLeibie & "%' and shipinleixing=2", "Main", 1)
                Case "欧美"
                    Form1.SelectFromDatabase("where leibie like '%" & myLeibie & "%' and shipinleixing=3", "Main", 1)
                Case "所有"
                    Form1.SelectFromDatabase("where leibie like '%" & myLeibie & "%'", "Main", 1)
            End Select
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Me.Close()
        Form1.Activate()
        Form1.ToolStripLabel6.Text = "类别为：" & Juncode(myLeibie)
    End Sub

    Private Sub ToolStripComboBox1_Click(sender As Object, e As EventArgs) Handles ToolStripComboBox1.Click

    End Sub

    Private Sub ToolStripComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        ToolStripButton2_Click(sender, e)
        FlowLayoutPanel1.Focus()
    End Sub

    Private Sub ToolStripTextBox1_Click(sender As Object, e As EventArgs) Handles ToolStripTextBox1.Click

    End Sub

    Private Sub ToolStripTextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ToolStripTextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            ToolStripButton1_Click(sender, e)
        End If
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        For Each lb In FlowLayoutPanel1.Controls
            Dim mylabel As Label
            mylabel = lb
            mylabel.BackColor = Color.White
            If InStr(mylabel.Text, ToolStripTextBox1.Text) > 0 Then
                mylabel.BackColor = Color.Yellow
            End If
        Next
    End Sub

    Private Sub FlowLayoutPanel1_Paint(sender As Object, e As PaintEventArgs)

    End Sub

    Private Sub FlowLayoutPanel1_Click(sender As Object, e As EventArgs)
        FlowLayoutPanel1.Focus()
    End Sub

    Private Sub LeibieForm_MaximizedBoundsChanged(sender As Object, e As EventArgs) Handles Me.MaximizedBoundsChanged

    End Sub

    Private Sub LeibieForm_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        e.Cancel = -1
        Me.Hide()
    End Sub

    Private Sub FlowLayoutPanel2_Paint(sender As Object, e As PaintEventArgs) Handles FlowLayoutPanel2.Paint

    End Sub
End Class