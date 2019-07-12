Module myVariable
    Public GlobalFanhao As String
    Public Zhengpaixu As Boolean
    Public PaixuSql As String
    Public Clickindex As Integer
    Public myLeibie As String
    Public myYanyuan As String
    Public myBiaoqian As String
    Public myFaxingshang As String
    Public myZhizuoshang As String
    Public myDaoyan As String
    Public myXilie As String
    Public mySearch As String
    Public myThreadIsCompleted(0) As Boolean
    Public myThreadNum As Integer
    Public MyThread(0) As System.Threading.Thread
    Public MyExtraPicThread(0) As System.Threading.Thread
    '1所有视频
    '2所有视频
    '3所有视频
    '4所有视频
    '5所有视频
    '10类别
    '11演员
    '12标签
    '13发行商
    '14制作商
    '15导演
    '16系列
    '20查找
    Public IsSmallPicAutoSize As Boolean
    Public IsCloseNow As Boolean '是否开启了夜间模式
    Public JavWebSite As String 'javbus网页
    Public JavLibrarySite As String 'JavLibrary网页
    Public FlowlayoutPanel_PageNum As Integer
    Public BigPicSavePath As String
    Public SmallPicSavePath As String
    Public ActressesPicSavePath As String
    Public ExtraPicSavePath As String
    Public DataBaseSavePath As String
    Public con_ConnectionString As String
    Public con_ConnectionString_read As String
    Public IsFirstStart As Boolean
    Public Paixuleixing As Integer
    Public IsFanhaoShow As Boolean
    Public IsMingchengShow As Boolean
    Public IsRiqiShow As Boolean
    Public IsShoucangShow As Boolean
    Public IsClickSmalPicShowBigPic As Boolean
    Public ColorZhuti As String

    '初始窗体的位置和大小
    Public Me_X As Double
    Public Me_Y As Double
    Public Me_Width As Double
    Public Me_Height As Double
    Public NewItem As String
    Public ToolTipbtnName As String
    Public SmallPicFanhao As New myFontClass
    Public SmallPicMingcheng As New myFontClass
    Public SmallPicRiqi As New myFontClass
    Public SmallPicQita As New myFontClass

    Public BigpicFanhao As New myFontClass
    Public BigpicMingcheng As New myFontClass
    Public BigpicFaxingriqi As New myFontClass
    Public BigpicChangdu As New myFontClass
    Public BigpicDaoyan As New myFontClass
    Public BigpicZhizuoshang As New myFontClass
    Public BigpicFaxingshang As New myFontClass
    Public BigpicLeibie As New myFontClass
    Public BigpicYanyuan As New myFontClass
    Public BigpicBiaoqian As New myFontClass
    Public BigpicXilie As New myFontClass
    Public TotalNum As Int16
    '扫描设置
    Public Scan_Shipinleixing As String
    Public Scan_WenjianDaxiao As String
    Public myChongfufanhaoIndex(0) As Integer
    Public CanImport(2) As Boolean
    Public BubingFanhao As String


    '其他
    Public myBitmap(0) As Bitmap '额外照片的image地址
    Public ExtraPicIndex As Integer

    Public mySearchPattern As String()

End Module
