VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' SUMMARY:
'Baidu-Express-Query-VBA
'百度快递运单号查询接口-Excel通过VBA调用
'------
'参数来源:https://github.com/mo10/Baidu-Express-Query
'
'用Excel表格来管理订单物流数据，还是比较简单便捷普遍的日常应用场景。
'这里通过宏实现实时更新物流信息到Excel中，方便跟踪订单物流最新状态，简化仓管及售后人员对订单物流数据跟踪工作。
'借助百度搜索的统一物流查询接口，实现小批量、小并发、整合多个物流公司的统一的物流信息更新功能。
'
'同时借助Excel强大的宏处理能力，实现对物流的分类筛选分组导出到独立Sheet或独立Excel文件的管理功能。对有类似分组分类导出需求的，提供一个小小的参考。
'
'AUTHOR:
'logintan@foxmail.com
'
'http://blog.tanyucheng.cn/
'mailto:logintan@foxmail.com?subject=your+code+stinks
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    '快递单号列号
    Dim expressNumColNum As Integer
    expressNumColNum = 1
    If Target.Column = expressNumColNum Then '单条快递信息查询
        Call queryExpressInfo
    ElseIf Target.Address = "$G$1" Then '批量快递信息查询
        Call batchQueryExpressInfo
    ElseIf Target.Address = "$H$1" Then '分组导出到新的Sheet
        Call expSheets
    ElseIf Target.Address = "$I$1" Then '分组导出到对应的文件
        Call expFiles
    End If

End Sub

'查询单条物流信息（当前选中）
Sub queryExpressInfo()
    Dim actRowNum As Integer
    actRowNum = ActiveCell.Row
    
    Set winHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    queryPrevious winHttp
    queryDo actRowNum, winHttp
End Sub

'批量物流信息（限制最多不超过50次每次）
Sub batchQueryExpressInfo()
    Dim beginRow, endRow As Integer
    On Error Resume Next
    beginRow = Application.InputBox("起始行", "物流更新范围输入:begin", , , , , , 2)
    
    endRow = Application.InputBox("结束行", "物流更新范围输入:" & beginRow & "-?", , , , , , 2)
    
    If beginRow <= 1 Then
        MsgBox "起始行输入有误，必须是数字且大于1"
    ElseIf endRow <= 1 Or endRow - beginRow < 0 Then
        MsgBox "结束行输入有误，必须是数字且大于起始行"
    ElseIf endRow - beginRow > 50 Then
        MsgBox "每次批量更新范围不能超过50条"
    Else
        
        '快递单号列号
        Dim expressNumColNum As Integer
        expressNumColNum = 1
        Set masterSheet = Sheets("Master")
        Set winHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
        queryPrevious winHttp
        For i = beginRow To endRow
            Dim actRowIndex As Integer
            actRowIndex = i
            queryDo actRowIndex, winHttp
        Next
        MsgBox endRow - beginRow + 1 & " 条物流记录更新完毕"
    End If
    
End Sub

Sub expSheets()
    Dim rowCounts As Integer
    Dim expressCompanyNameColNum As Integer
    expressCompanyNameColNum = 2 '快递公司名称列
    
    Set masterSheet = Sheets("Master")
    rowCounts = masterSheet.Cells(2, expressCompanyNameColNum).CurrentRegion.Rows.Count '获取B2所在单元格所在数据块有多少行（含表头）
    
    '统计B列（快递公司列，从B2开始）：B[rowCounts]中不重复数据的个数,下面用了字典对象的key唯一的方法达到统计效果
    Set tempDict = CreateObject("scripting.dictionary")
    tempArr = masterSheet.Range(masterSheet.Cells(2, expressCompanyNameColNum), masterSheet.Cells(rowCounts, expressCompanyNameColNum)) '从第2行开始，排除表头
    For Each c In tempArr
        If c <> "" Then
            tempDict(c) = ""
        Else
            '考虑快递公司列为空的情况。
            tempDict("无") = ""
        End If
    Next
    
    '这里要考虑当对应快递公司已经存在的情况，示例：如果"顺丰速运"已经存在，则添加新Sheet名称为 "顺丰速运(1)"，依此类推
    Set tempSheetDict = CreateObject("scripting.dictionary")
    
    '把所有Sheets的名字提取出来
    For i = 1 To Sheets.Count
        tempSheetDict(Sheets(i).Name) = ""
    Next

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    '根据 B 列 关键字处理数据开始
    For Each Key In tempDict
        Dim sName As String
        
        If tempSheetDict.Exists(Key) Then
            Dim num As Integer
            num = 1
            Do While tempSheetDict.Exists(Key & "(" & num & ")")
                num = num + 1
            Loop
            sName = Key & "(" & num & ")"
        Else
            sName = Key
        End If
        
        Sheets.Add After:=Sheets(Sheets.Count)
        Set lastSheet = Sheets(Sheets.Count)
        lastSheet.Name = sName

        '复制表头 在外面处理第一行表头数据
        lastSheet.Range("A1:E1").Value = masterSheet.Range("A1:E1").Value
        
        Dim tempRow As Integer
        tempRow = 2 '处理复制新表的行号记录
        
        For i = 2 To rowCounts Step 1 '从第二行开始遍历
            If (Key = "无" And masterSheet.Cells(i, expressCompanyNameColNum).Value = "") Or masterSheet.Cells(i, expressCompanyNameColNum).Value = Key Then
                'Copy数据：
                lastSheet.Cells(tempRow, 1).NumberFormatLocal = "@"  '快递单号所在列需要设置为文本格式
                lastSheet.Range("A" & tempRow & ":E" & tempRow).Value = masterSheet.Range("A" & i & ":E" & i).Value
                tempRow = tempRow + 1
            End If
        Next
    Next
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub


Sub expFiles()
    Dim xPath As String '导出文件的路径
    xPath = Application.ActiveWorkbook.Path '和当前文件同一目录
   
    Dim rowCounts As Integer
    Dim expressCompanyNameColNum As Integer
    expressCompanyNameColNum = 2 '快递公司名称列
    
    Set masterSheet = Sheets("Master")
    rowCounts = masterSheet.Cells(2, expressCompanyNameColNum).CurrentRegion.Rows.Count '获取B2所在单元格所在数据块有多少行（含表头）
    
    '统计B列（快递公司列，从B2开始）：B[rowCounts]中不重复数据的个数,下面用了字典对象的key唯一的方法达到统计效果
    Set tempDict = CreateObject("scripting.dictionary")
    tempArr = masterSheet.Range(masterSheet.Cells(2, expressCompanyNameColNum), masterSheet.Cells(rowCounts, expressCompanyNameColNum)) '从第2行开始，排除表头
    For Each c In tempArr
        If c <> "" Then
            tempDict(c) = ""
        Else
            '考虑快递公司列为空的情况。
            tempDict("无") = ""
        End If
    Next
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    '根据 B 列 关键字处理数据开始
    For Each Key In tempDict
        Set tempXl = New Excel.Application
        Set tempXl = CreateObject("Excel.Application")
        Set tempXlB = tempXl.Workbooks.Add
        Set tempXLBS = tempXlB.Worksheets.Add
        
        tempXLBS.Name = Key
        
        '复制表头 在外面处理第一行表头数据
        tempXLBS.Range("A1:E1").Value = masterSheet.Range("A1:E1").Value
              
        Dim tempRow As Integer
        tempRow = 2 '处理复制新表的行号记录
        
        For i = 2 To rowCounts Step 1 '从第二行开始遍历
            If (Key = "无" And masterSheet.Cells(i, expressCompanyNameColNum).Value = "") Or masterSheet.Cells(i, expressCompanyNameColNum).Value = Key Then
                'Copy数据：
                tempXLBS.Cells(tempRow, 1).NumberFormatLocal = "@"  '快递单号所在列需要设置为文本格式
                tempXLBS.Range("A" & tempRow & ":E" & tempRow).Value = masterSheet.Range("A" & i & ":E" & i).Value
                tempRow = tempRow + 1
            End If
        Next
        
        tempXlB.SaveAs Filename:=xPath & "\" & Key & ".xlsx"
        tempXlB.Close False
        tempXl.Quit
    Next
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub


'在线查询前置步骤
Function queryPrevious(winHttp As Variant)
    With winHttp
        On Error Resume Next
        '访问一次百度首页，winHttp 获取对应Cookies信息
        URL_BAIDU = "https://www.baidu.com/"
        .Open "GET", URL_BAIDU, False
        .Option(6) = False '.Option(6) = False这句是不能少的，禁止网页重定向，否则的话登陆之后网页会跳转，那就得不到我们需要的cookie了
        .setRequestHeader "Connection", "Keep-Alive"
        .setRequestHeader "Cache-Control", "no-cache"
        .setRequestHeader "Accept", "text/html, application/xhtml+xml, */*"
        .setRequestHeader "Accept-Encoding", "deflate"
        .setRequestHeader "Accept-Language", "zh-CN,zh;q=0.8,de-CH;q=0.6,de;q=0.4"
        .setRequestHeader "Pragma", "no-cache"
        .setRequestHeader "Upgrade-Insecure-Requests", "1"
        .setRequestHeader "Host", "www.baidu.com"
        .setRequestHeader "Content-Type", "text/html;"
        .setRequestHeader "Referer", "https://www.baidu.com/"
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/57.0.2987.98 Safari/537.36"
        .send
    End With
    
End Function

'请求在线查询接口，执行具体物流信息查询（执行此函数前需要先调用 在线查询前置步骤：queryPrevious）
Function queryDo(actRowNum As Integer, winHttp As Variant)

    Set masterSheet = Sheets("Master")
    Dim URL_BAIDU, URL_BAIDU_EXPRESS, rspText, lastUpdateTime
    
    Dim expressNum As String
    '快递单号列号
    Dim expressNumColNum As Integer
    
    expressNumColNum = 1
    
    expressNum = masterSheet.Cells(actRowNum, expressNumColNum).Value '读取快递单号
    If expressNum <> "" And actRowNum > 1 Then '快递单号不为空（排除第一行表头不处理）
        URL_BAIDU_EXPRESS = "https://sp0.baidu.com/9_Q4sjW91Qh3otqbppnN2DJv/pae/channel/data/asyncqury?cb=&com=&appid=4001&nu=" & expressNum & "&vcode=&token=&_=" & DateDiff("s", "01/01/1970 00:00:00", Now()) & "021"
        
        With winHttp
            On Error Resume Next
            .Open "GET", URL_BAIDU_EXPRESS, False
            .Option(6) = False
            .setRequestHeader "Connection", "Keep-Alive"
            .setRequestHeader "Cache-Control", "no-cache"
            .setRequestHeader "Accept", "text/html, application/xhtml+xml, */*"
            .setRequestHeader "Accept-Encoding", "deflate"
            .setRequestHeader "Accept-Language", "zh-CN,zh;q=0.8,de-CH;q=0.6,de;q=0.4"
            .setRequestHeader "Pragma", "no-cache"
            .setRequestHeader "Upgrade-Insecure-Requests", "1"
            .setRequestHeader "Host", "sp0.baidu.com"
            .setRequestHeader "Content-Type", "text/html;"
            .setRequestHeader "Referer", "https://www.baidu.com/"
            .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/57.0.2987.98 Safari/537.36"
            .send
            
            rspText = .responsetext
            'Debug.Print rspText
            
        End With
        
         '通过引入htmlfile对象，对物流接口返回的Json数据进行解释
        Dim objHTML, objJSy, objWin As Object
        Set objHTML = CreateObject("htmlfile")
        Set objWin = objHTML.parentWindow
        
        Dim expressCompanyNameColNum, expressStateColNum, lastUpdateTimeColNum, expressContextColNum As Integer
        
        expressCompanyNameColNum = 2 '快递公司名称列
        expressStateColNum = 3 '物流状态列
        lastUpdateTimeColNum = 4 '最后更新时间列
        expressContextColNum = 5 '最新物流动态列
            
        lastUpdateTime = Format(Date, "yyyy年m月d日") & " " & Format(time, "HH:mm:ss") '以当前系统时间为最后更新时间
        
        
        '物流接口返回结果格式
        
        '{
        '  "msg":"",                  #查询状态信息,查询成功时为空
        '  "status":"0",              #查询状态,0为成功
        '  "error_code":"0",          #查询状态,0为成功
        '  "data":{
        '    "notice":"物流信息与官网实时同步，已耗时X天X时XX分",
        '     "com":"yunda",          #快递公司英文
        '     "company":{...},        #快递公司信息
        '     "source":{...},         #数据来源
        '     "kuaidiSource":{...},   #快递接口来源
        '     "info":{
        '      "status":"1",           #查询状态?
        '      "com":"yunda",          #快递公司英文
        '      "state":"0",            #签收状态?
        '       "context":{
        '        {
        '          "time":"1482000831",      #快递动态发生 时间戳
        '          "desc":"在X,即将发往：X"   #发生事件描述
        '        },
        '        {...}
        '      },
        '      "_source_com":"yunda",       #数据来源公司
        '      "_support_from":"partner"    #没几把用
        '    }
        '  }
        '}
        
        objWin.execScript "var json = " & rspText, "JScript"     '定义一个JS变量，将JSON数据引入
        objWin.execScript "var code = json.status;", "JScript"
        objWin.execScript "var msg = json.msg;", "JScript"
        objWin.execScript "var companyName = json.data.company.fullname;", "JScript"
        masterSheet.Cells(actRowNum, expressCompanyNameColNum).Value = objWin.companyName
        masterSheet.Cells(actRowNum, lastUpdateTimeColNum).Value = lastUpdateTime
        If objWin.code = 0 Then
            objWin.execScript "var desc = json.data.info.context[0].desc;", "JScript"
            objWin.execScript "var time = json.data.info.context[0].time;", "JScript"
            objWin.execScript "var infoStatus = json.data.info.status;", "JScript"
            
            objWin.execScript "var infoState = json.data.info.state;", "JScript"
           
            If objWin.infoState = 3 Then
                masterSheet.Cells(actRowNum, expressStateColNum).Value = "已经签收"
            ElseIf objWin.infoState = 0 Then
                masterSheet.Cells(actRowNum, expressStateColNum).Value = "物流中"
            Else
                masterSheet.Cells(actRowNum, expressStateColNum).Value = objWin.infoState
            End If
            
            masterSheet.Cells(actRowNum, expressContextColNum).Value = objWin.companyName & ":(" & Format(19 + 70 * 365 + (8 * 3600 + Left(objWin.time, 10)) / 86400, "yyyyMMDD HH:mm:ss") & ")" & objWin.desc
            
        Else
            If objWin.code = -3 Or objWin.code = -2 Then
                masterSheet.Cells(actRowNum, expressStateColNum).Value = "无物流信息"
            Else
                masterSheet.Cells(actRowNum, expressStateColNum).Value = objWin.code
            End If
            masterSheet.Cells(actRowNum, expressContextColNum).Value = objWin.msg
        End If
        'If masterSheet.Cells(actRowNum, expressStateColNum).Value <> "" And masterSheet.Cells(actRowNum, expressContextColNum).Value <> "" Then
        '    MsgBox masterSheet.Cells(actRowNum, expressContextColNum).Value
        'End If
    End If
End Function