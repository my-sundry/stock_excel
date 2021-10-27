Function FillOneRow(url As String, r As Integer) As Integer
    With CreateObject("msxml2.xmlhttp")
        .Open "GET", url, False
        .send
        sp = Split(.responsetext, "~")
        If UBound(sp) > 3 Then
            FillOneRow = 1
            Cells(r, 2).Value = sp(1) '名称
            Cells(r, 3).Value = sp(3) '当前价格
            Rem Cells(r, 4).Value = sp(4) '昨日收盘价
            Dim zhangDie As Double
            zhangDie = sp(32)
            Rem Cells(r, 5).Value = zhangDie
            If zhangDie > 0 Then
                '上涨使用红色
                Cells(r, 5).Font.Color = vbRed
                Cells(r, 3).Font.Color = vbRed
            Else
                '下跌使用绿色
                Cells(r, 5).Font.Color = &H228B22
                Cells(r, 3).Font.Color = &H228B22
            End If
        Else
            FillOneRow = 0
        End If
    End With
End Function
 
Sub GetData()
Attribute GetData.VB_ProcData.VB_Invoke_Func = "X \n14"
    Dim succeeded As Integer
    Dim url As String
    Dim row As Integer
    Dim code As String
    Dim dateStr As String
    Dim cash As String
    Dim current As String
    Dim firstCode As String
    Dim secondCode As String
    current = Date
    Dim currentRow As Integer
    currentRow = 0
    Dim zhangDie As Double
    Dim isSet As Boolean
            
    For row = 2 To Range("A1").CurrentRegion.Rows.Count
        code = Cells(row, 1).Value
        succeeded = 0
        
        If code <> "" Then
            firstCode = LCase(Mid(code, 1, 1))
            secondCode = LCase(Mid(code, 2, 1))
            
            If firstCode = "s" And secondCode = "h" Then
                url = "http://qt.gtimg.cn/q=" & Cells(row, 1).Value
                succeeded = FillOneRow(url, row)
            ElseIf firstCode = "s" And secondCode = "z" Then
                url = "http://qt.gtimg.cn/q=" & Cells(row, 1).Value
                succeeded = FillOneRow(url, row)
            Else
                If firstCode <> "0" Then
                    url = "http://qt.gtimg.cn/q=sh" & Cells(row, 1).Value
                    succeeded = FillOneRow(url, row)
                End If
                
                If succeeded = 0 Then
                    url = "http://qt.gtimg.cn/q=sz" & Cells(row, 1).Value
                    succeeded = FillOneRow(url, row)
                End If
            End If
            
            If succeeded = 0 Then
                MsgBox ("获取失败")
            End If
        End If
    Next
End Sub
