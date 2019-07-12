Imports System.Text.RegularExpressions
Imports System.IO
Module OtherFunction


    Public Function IsFileInUse(fileName As String) As Boolean
        Dim inUse As Boolean = True
        Dim fs As FileStream
        Try
            fs = New FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.None)
            inUse = False
        Catch ex As Exception
        Finally
            If fs IsNot Nothing Then
                fs.Close()
            End If
        End Try
        IsFileInUse = inUse
    End Function






    '//vb将unicode转成汉字，如：\u8033\u9EA6，转后为：耳麦UnicodeDecode
    Public Function Juncode(strCode As String) As String
        Juncode = strCode
        If InStr(Juncode, "\u") <= 0 Then
            Exit Function
        End If
        strCode = LCase(strCode)
        Dim mc As MatchCollection
        mc = Regex.Matches(strCode, "\\u\S{1,4}")
        For Each m In mc
            strCode = Replace(strCode, m.ToString, ChrW("&H" & Mid(CStr(m.ToString), 3, 6)))
        Next
        Juncode = strCode
    End Function


    '//将中文转为unicode编码，如：耳麦，转后为：\u8033\u9EA6 UnicodeEncode
    Function Jencode(strCode As String) As String
        'Jencode = strCode
        Dim a() As String
        Dim str As String
        Dim i As Integer
        'StrTemp = strCode

        For i = 0 To Len(strCode) - 1
            On Error Resume Next
            str = Mid(strCode, i + 1, 1)
            If isChinese(str) = True Then '//是中文
                Jencode = Jencode & "\u" & StrDup(4 - Len(Hex(AscW(str))), "0") & Hex(AscW(str))
            Else '//不是中文
                Jencode = Jencode & str
            End If

        Next

    End Function

    '//是否为中文
    Public Function isChinese(Text As String) As Boolean

        Dim l As Long
        Dim i As Long
        l = Len(Text)
        isChinese = False

        For i = 1 To l
            If Asc(Mid(Text, i, 1)) < 0 Or Asc(Mid(Text, i, 1)) < 0 Then
                isChinese = True
                Exit Function
            End If
        Next

    End Function














    '转码
    'Public Function Jencode(ByVal iStr)
    '    If IsDBNull(iStr) Or IsNothing(iStr) Then
    '        Jencode = ""
    '        Exit Function
    '    End If
    '    Jencode = iStr
    '    For i = 0 To 25
    '        Jencode = Replace(Jencode, F(i), E(i))
    '    Next
    'End Function

    ''解码
    'Function Juncode(ByVal iStr)
    '    If IsDBNull(iStr) Or IsNothing(iStr) Then
    '        Juncode = ""
    '        Exit Function
    '    End If
    '    Juncode = iStr
    '    For i = 0 To 25
    '        Juncode = Replace(Juncode, E(i), F(i))
    '    Next
    'End Function

End Module
