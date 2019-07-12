Imports System.ComponentModel

Public Class myDownLoadForm




    Private Sub MyDownLoadForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Me.FormBorderStyle = FormBorderStyle.FixedSingle

    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs)

        'ToolStrip2.Visible = True
    End Sub


    Private Sub myDownLoadForm_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        'If MsgBox("是否确定要退出？", vbYesNo) = vbNo Then e.Cancel = -1
    End Sub

    Private Sub FlowLayoutPanel1_Paint(sender As Object, e As PaintEventArgs) Handles FlowLayoutPanel1.Paint

    End Sub

    Private Sub myDownLoadForm_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove

    End Sub

    Private Sub myDownLoadForm_MouseWheel(sender As Object, e As MouseEventArgs) Handles Me.MouseWheel
        'e.X e.Y以窗体左上角为原点，aPoint为鼠标滚动时的坐标
        Dim aPoint As New Point(e.X, e.Y)
        'this.Location.X,this.Location.Y为窗体左上角相对于screen的坐标,得出的结果是鼠标相对于电脑screen的坐标
        aPoint.Offset(Me.Location.X, Me.Location.Y)
        Dim r As New Rectangle(FlowLayoutPanel1.Location.X, FlowLayoutPanel1.Location.Y, FlowLayoutPanel1.Width, FlowLayoutPanel1.Height)
        ' MessageBox.Show(flowLayoutPanel1.Width+"  "+flowLayoutPanel1.Height)
        '判断鼠标是不是在flowLayoutPanel1区域内
        If (RectangleToScreen(r).Contains(aPoint)) Then
            '设置鼠标滚动幅度的大小
            FlowLayoutPanel1.AutoScrollPosition = New Point(0, FlowLayoutPanel1.VerticalScroll.Value - e.Delta / 2)
        End If
    End Sub

    Private Sub ToolStripButton6_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'MsgBox(RichTextBox1.Text)
        Dim MyDownClass(UBound(Split(RichTextBox1.Text, "|"))) As DownClass
        Dim MyThread(UBound(Split(RichTextBox1.Text, "|"))) As System.Threading.Thread
        Dim i As Integer
        i = 0
        ListBox1.Items.Clear()
        For Each fanhao In Split(RichTextBox1.Text, "|")
            If fanhao <> "" Then
                MyDownClass(i) = New DownClass
                MyDownClass(i).getWebResponse_fanhao = fanhao
                'MyDownClass(i).mylistbox = ListBox1
                Dim myPanel1 As New Panel_DownItem
                myPanel1.NameText = fanhao & ".jpg"
                myPanel1.Panel_DownItem_initial()
                FlowLayoutPanel1.Controls.Add(myPanel1.myPanel)

                MyDownClass(i).FilePath = Application.StartupPath
                'MyDownClass(i).ProB = myPanel1.Prob
                MyThread(i) = New System.Threading.Thread(AddressOf MyDownClass(i).getWebResponseAndDownload)
                MyThread(i).Start()
                i = i + 1
            End If
        Next
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        RichTextBox1.Clear()
        For i = 0 To 10
            RichTextBox1.Text = RichTextBox1.Text & "|" & "RKI-" & i + 400
        Next
    End Sub
End Class