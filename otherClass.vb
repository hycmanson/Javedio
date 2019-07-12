Imports System.IO
Module otherClass
    Public Class MyDateSorter
        Implements IComparer

        Public Function Compare(x As Object, y As Object) As Integer Implements System.Collections.IComparer.Compare
            If x Is Nothing AndAlso y Is Nothing Then
                Return 0
            End If
            If x Is Nothing Then
                Return -1
            End If
            If y Is Nothing Then
                Return 1
            End If
            Dim xInfo As FileInfo = DirectCast(x, FileInfo)
            Dim yInfo As FileInfo = DirectCast(y, FileInfo)


            '依名稱排序   
            Return xInfo.FullName.CompareTo(yInfo.FullName) '遞增   
            'return yInfo.FullName.CompareTo(xInfo.FullName);//遞減   
            '依修改日期排序   
            'Return xInfo.LastWriteTime.CompareTo(yInfo.LastWriteTime) '遞增   
            'return yInfo.LastWriteTime.CompareTo(xInfo.LastWriteTime) '遞減   
        End Function

    End Class
End Module
