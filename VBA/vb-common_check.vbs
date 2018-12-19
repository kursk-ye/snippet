Attribute VB_Name = "common_check"
Option Explicit

'检查数据文件的列是否与模板一致
'并增加用于数据库字段名一行，用于数据导入
'match_wb               对应表文件workbook
'province_city          数据文件来自于省还是地市，省1，地市0
'data_ws                数据文件的worksheet
'table_name             所检查的表名
'start_check_row        表头起始行
'dbcol_code_row         插入数据文件导入数据库识别字段名称，如果原有则覆盖
'log_path               日志文件路径
'return                 数据文件数据部分的宽度
Function checkDataFileCol(match_wb As Excel.Workbook _
                    , province_city As Integer _
                    , data_ws As Worksheet _
                    , table_name As String _
                    , start_check_row As Integer _
                    , dbcol_code_row As Integer _
                    , log_path As String)

Dim match_ws As Worksheet                               '省级对应表或地市对应表
Dim match_tab_row, match_tab_col As Integer             '对应表里的行号和列号
Dim data_tab_row, data_tab_col As Integer               '数据文件里的行号和列号
Dim data_tab_wide As Integer                            '数据文件数据部分的宽度

Dim row As Integer

If province_city = 1 Then
    Set match_ws = match_wb.Worksheets("网省对应表")
ElseIf province_city = 0 Then
    Set match_ws = match_wb.Worksheets("地市对应表")
Else
    MsgBox "错误！未指定对应表使用省或地市"
End If

'取消整表数据有效性
data_ws.cells.Validation.Delete

'因为已有代码不正确，删除再插入一行，因为下发工程项目表没有此行，因此不删除
If InStr(data_ws.Name, "工程") = 0 Then
    data_ws.Rows(dbcol_code_row).Delete
End If
data_ws.Rows(dbcol_code_row).Insert shift:=xlShiftDown

data_tab_wide = 0

row = 1
data_tab_col = 0
Do While match_ws.cells(row, 1).Text <> ""
    If match_ws.cells(row, 1).Text = table_name Then
        '统计数据文件数据部分的宽度
        data_tab_wide = data_tab_wide + 1
    
        match_tab_row = row
        match_tab_col = 2
        
        data_tab_row = start_check_row
        data_tab_col = data_tab_col + 1
        
        Do While match_ws.cells(match_tab_row, match_tab_col).MergeArea.Item(1).Text <> ""
        
            If (match_ws.cells(match_tab_row, match_tab_col).MergeArea.Item(1).Text = _
                        data_ws.cells(data_tab_row, data_tab_col).MergeArea.Item(1).Text) _
                    Or (match_ws.cells(match_tab_row, match_tab_col).MergeArea.Item(1).Text = "h") Then
                match_tab_col = match_tab_col + 1
                data_tab_row = data_tab_row + 1
                
                '插入数据库导入的识别字段名
                If match_ws.cells(match_tab_row, match_tab_col).MergeArea.Item(1).Text = "" Then
                        data_ws.cells(dbcol_code_row, data_tab_col).Value2 = _
                            match_ws.cells(match_tab_row, 14)
                End If
            Else
                Call log.error(log_path _
                                , "数据文件与模板不符。请检查表:" & table_name & "第" _
                                    & data_tab_row & "行，第" & data_tab_col & "列是否与模板一致")
                data_ws.Columns(data_tab_col).Interior.ColorIndex = 6
                GoTo Next_Data_Table_Column
            End If
            
        Loop
    End If
    
Next_Data_Table_Column:

    row = row + 1

Loop

checkDataFileCol = data_tab_wide

End Function

'修改有错误单元格的行为黄色，并批注错误说明，修改单元格为蓝色
Sub comment_cell(cell As Excel.Range, comment As String)

cell.Parent.Rows(cell.row).Interior.ColorIndex = 6
cell.Interior.ColorIndex = 33
If cell.comment Is Nothing Then
    cell.AddComment
End If
cell.comment.Text Text:=comment & Chr(10) & cell.comment.Text

error_count = error_count + 1

End Sub

'判断字符串是否含有字母,，空格，如果是，返回true
Function IsCounter(value As String) As Boolean
    Dim reg As Object
    Set reg = CreateObject("vbscript.RegExp")
    With reg
     .Global = True
     '.Pattern = "[a-zA-Z_，=~!@#$%^&()-{}\[\]\|:;'“?<>/ ]"
     .Pattern = "[a-zA-Z_，= :;：；]"
    End With
    IsCounter = reg.Test(value)
    Set reg = Nothing
End Function


'进行数字或字符检查
'num_str_code       该列进行数字检查还是字符检查，1，进行数字检查；2，进行字符检查
'data_table_range   被检查的数据集
'data_table_name    被检查的数据文件的表名
'log_path           日志文件路径
'update_num_method  修改为数字型的处理方式，1：使用VAL方式转换，对应于10kV这种类型；2：使用公式转换，如容量里填写成容量组成
Sub check_num_str(num_str_code As Integer _
                    , data_table_range As Excel.Range _
                    , data_table_name As String _
                    , log_path As String _
                    , update_num_method As Integer)
                    
Dim data_table_cell As Excel.Range
Dim try_flag As Boolean

If num_str_code = 1 Then
    For Each data_table_cell In data_table_range
    try_flag = True
Try_again:
        If Not (Application.WorksheetFunction.IsNumber(data_table_cell.Value2)) And data_table_cell.Text <> "" Then
            If update_num_method = 1 And try_flag = True Then
                data_table_cell.Value2 = Val(data_table_cell.Text)
                try_flag = False
                GoTo Try_again
            ElseIf update_num_method = 2 And try_flag = True And IsCounter(data_table_cell.Text) = False Then
                data_table_cell.Value2 = Replace(data_table_cell.Value2, "×", "*")
                data_table_cell.Formula = "=" & data_table_cell.Text
                try_flag = False
                GoTo Try_again
            Else
                'Call log.error(log_path:=log_path _
                                , msg:="不是数字型错误。请检查表:" & data_table_name & "第" _
                                        & data_table_cell.row & "行，第" & data_table_cell.Column & "列是否的确是数字")
                comment_cell data_table_cell, "该单元格应为数字型"
            End If
        End If
    Next
ElseIf num_str_code = 2 Then

Else
    MsgBox "use error num_str_code"
End If

End Sub

'修改为指定的精度
'accuracy               指定小数点后位数,为0,1，2，3，4,6
'                       5 文本
'data_table_range       被修改的数据集
Sub updat_accuracy(accuracy As Integer _
                    , data_table_range As Excel.Range)
Select Case accuracy
    Case 0
        data_table_range.NumberFormat = "#;[red]#"
    Case 1
        data_table_range.NumberFormat = "#.#;[red]#.#"
    Case 2
        data_table_range.NumberFormat = "#.##;[red]#.##"
    Case 3
        data_table_range.NumberFormat = "#.###;[red]#.###"
    Case 4
        data_table_range.NumberFormat = "#.####;[red]#.####"
    Case 5
        data_table_range.NumberFormat = "@"
    Case 6
        data_table_range.NumberFormat = "#.######;[red]#.######"
    Case Else
        MsgBox "指定了错误的小数点精度"
End Select
End Sub

'修改日期类型为指定格式
'date_type              1,修改为2010/09/28;2,修改为YYYY-MM-DD HH:MIN;3,yyyy;4,hh:mm
'data_table_range       被修改的数据集
Sub update_date(date_type As Integer _
                , data_table_range As Excel.Range _
                , log_path As String)

Dim data_cell As Excel.Range
Dim data_table_name As String

data_table_name = data_table_range.Parent.Name

    Select Case date_type
        Case 1
            For Each data_cell In data_table_range
                If IsDate(data_cell.Text) And data_cell.Text <> "" Then
                    data_cell.Value2 = CDate(data_cell.Value2)
                    data_cell.NumberFormatLocal = "yyyy/m/d"
                    data_cell.Value2 = "'" & data_cell.Text
                    data_cell.NumberFormat = "@"
                ElseIf data_cell.Text <> "" Then
                    comment_cell cell:=data_cell _
                                , comment:="不是日期。请按 年/月/日 格式填写"
                End If
            Next
        Case 2
            
            For Each data_cell In data_table_range
                If IsDate(data_cell) And data_cell.Text <> "" Then
                    data_cell.Value2 = CDate(data_cell.Value2)
                    data_cell.NumberFormat = "yyyy/m/d hh:mm"
                    data_cell.Value2 = "'" & data_cell.Text
                    data_cell.NumberFormat = "@"
                ElseIf data_cell.Text <> "" Then
                    comment_cell cell:=data_cell _
                                , comment:="不是日期。请按 年/月/日 小时：分钟 格式填写"
                End If
            Next
              
        Case 3
        
            For Each data_cell In data_table_range
                    data_cell.NumberFormatLocal = "yyyy"
                    data_cell.Value2 = Val(data_cell.Value2)
            Next
            
        Case 4
            For Each data_cell In data_table_range
                If data_cell.Text <> "" Then
                    'data_cell.Value2 = CDate(data_cell.Value2)
                    data_cell.NumberFormatLocal = "hh:mm"
                    data_cell.Value2 = "'" & data_cell.Text
                    data_cell.NumberFormat = "@"
                ElseIf data_cell.Text <> "" Then
                    comment_cell cell:=data_cell _
                                , comment:="未按时间格式要求填写，请按 小时：分钟 格式填写"

                End If
            Next
            
        Case Else
            MsgBox "指定了错误的日期格式"
            
    End Select




End Sub


'在公用数据词典中查找指定的字符串，返回相应的编码
'如果找到，返回一个整型数
'如果没有找到.
Function array_substr_multi(array_searched() As String, str_searched As String, type_code As Long)
Dim i, j As Integer
Dim type_values As String

If str_searched = "" Then
    GoTo End_Function
End If

For i = 0 To UBound(array_searched)
    If array_searched(i, 0) = type_code Then
    type_values = type_values & array_searched(i, 3) & " 值为" & array_searched(i, 2) & "," & Chr(10)
    End If
Next

For i = 0 To UBound(array_searched)
    If array_searched(i, 0) = type_code Then
        If StrComp(StrConv(array_searched(i, 3), vbLowerCase), StrConv(str_searched, vbLowerCase), vbTextCompare) = 0 Then
            array_substr_multi = array_searched(i, 2)
            GoTo End_Function
        End If
    End If
Next

array_substr_multi = type_values

End_Function:
End Function

'在行政区划或公司组织机构数组中查询指定字符串的对应编码
'如果找到，返回一个整型数
'如果没有找到，返回 20122012
'array_searched()       被查询的数组
'str_searched()           被查询的字符串
Function array_substr_single(array_searched() As String, str_searched() As String)
Dim i As Integer, j As Integer


For j = 0 To UBound(str_searched)
    For i = 0 To UBound(array_searched)
        If array_searched(i, 1) = str_searched(j) Then
            array_substr_single = array_searched(i, 0)
            GoTo End_Function
        End If
    Next
Next


array_substr_single = 20122012

End_Function:
End Function

'转换编码
'exchange_array             存储数据标志,1:国网公司组织机构;2:行政区划（空则填充）;3:公用编码;4:行政区划（则不填充）
'data_table_range           被转换的数据文件数据集
'province_name              省名
'city_name                  地市名
'data_table_name            被转换的数据文件表名
'log_path                   日志文件路径
'type_code                  公用数据字典查询用的主编码，不用设为0
Sub exchange(exchange_array As Integer _
                , data_table_range As Excel.Range _
                , province_name As String _
                , city_name As String _
                , data_table_name As String _
                , log_path As String _
                , type_code As Long)
                
Dim unit_name() As String             '省名+地市名+市区/县名,或者市区/县名
Dim data_cell As Excel.Range        '被检查单元格
Dim result_search
Dim exchanged_value As String       '要被公用数据字典查询后替换的字符串

Dim i As Integer
                
Select Case exchange_array
    Case 1 'company_code
    
        For Each data_cell In data_table_range
        
            ReDim unit_name(2)
            unit_name(0) = data_cell.Text
            unit_name(1) = province_organ_name
            
            result_search = array_substr_single(company_code(), unit_name())
            
            If result_search = 20122012 Then
                'Call log.error(log_path:=log_path _
                                , msg:="所填公司名称与数据字典不对应错误。表：" & data_table_name & " 第" & _
                                data_cell.row & " 行，第" & data_cell.Column & _
                                " 列，所填公司名称与数据字典（供电公司.xls）不符")
                comment_cell cell:=data_cell _
                            , comment:="所填公司名称与数据字典不对应错误。表：" & data_table_name & " 第" & _
                            data_cell.row & " 行，第" & data_cell.Column & "列，所填公司名称与数据字典(供电公司.xls)不符"
            Else
                data_cell.NumberFormat = "@"
                data_cell.Value2 = result_search
            End If
        Next
        
    Case 2 'gov_code 如果为行政区划的单元格为空，则用编码填充
        
        For Each data_cell In data_table_range
        
            ReDim unit_name(5)
            
            If data_cell.Text = "" Then
                unit_name(0) = province_name + city_name
            Else
                unit_name(0) = data_cell.Text
                unit_name(1) = province_name + city_name + data_cell.Text
                unit_name(2) = province_name + data_cell.Text
                unit_name(3) = city_name + data_cell.Text
                'unit_name(4) = province_name + city_name       如果没有填写，不用修改
                
                
            End If
                 
            
            result_search = array_substr_single(gov_code(), unit_name())
            
            If result_search = 20122012 Then
                'Call log.error(log_path:=log_path _
                                , msg:="所填行政区划名称与数据字典不对应错误。表：" & data_table_name & " 第" & _
                                data_cell.row & " 行，第" & data_cell.Column & _
                                " 列，所填行政区划名称与数据字典（行政区划.xls）不符")
                comment_cell cell:=data_cell _
                            , comment:="所填行政区划名称与数据字典不对应错误。表：" & data_table_name & " 第" & _
                                data_cell.row & " 行，第" & data_cell.Column & _
                                "列，所填行政区划名称与数据字典(行政区划.xls)不符"
            Else
                data_cell.NumberFormat = "@"
                data_cell.Value2 = result_search
            End If
        Next
        
    Case 3 'pub_domain_code
        
        For Each data_cell In data_table_range
            exchanged_value = data_cell.Text
            
            result_search = array_substr_multi(pub_domain_code(), exchanged_value, type_code)
            
            If IsNumeric(result_search) Then
                data_cell.Value2 = result_search
            Else
                'Call log.error(log_path:=log_path, msg:="枚举值不对应。" & _
                                city_name & "表：" & data_table_name & "第" & _
                                data_cell.row & " 行，第" & data_cell.Column & "应填：" & result_search)
                comment_cell cell:=data_cell _
                            , comment:="枚举值不对应。" & _
                                city_name & "表：" & data_table_name & "第" & _
                                data_cell.row & " 行，第" & data_cell.Column & "应填：" & result_search
            End If
        Next
        
    Case 4 'gov_code 如果为行政区划的单元格为空，则不处理
        ReDim unit_name(1)
        
        For Each data_cell In data_table_range
            
            If data_cell.Text = "" Then
                
                GoTo NextCellExchange
                
            Else
            
                If InStr(data_cell.Text, ",") > 0 Or InStr(data_cell.Text, "，") > 0 Then
                
                    If InStr(data_cell.Text, ",") > 0 Then
                        unit_name = Split(data_cell.Text, ",")
                    Else
                        unit_name = Split(data_cell.Text, "，")
                    End If
                    
                    i = 0
                    Do While i <= UBound(unit_name)
                        result_search = array_substr_single(gov_code(), unit_name()) & "," & result_search
                        i = i + 1
                    Loop
                
                Else
                
                    unit_name(0) = data_cell.Text
                    result_search = array_substr_single(gov_code(), unit_name())
                
                End If
                
            End If
            
            
            If result_search = 20122012 Then
                'Call log.error(log_path:=log_path _
                                , msg:="所填行政区划名称与数据字典不对应错误。表：" & data_table_name & " 第" & _
                                data_cell.row & " 行，第" & data_cell.Column & _
                                " 列，所填行政区划名称与数据字典（行政区划.xls）不符")
                comment_cell cell:=data_cell _
                            , comment:="所填行政区划名称与数据字典不对应错误。表：" & data_table_name & " 第" & _
                                data_cell.row & " 行，第" & data_cell.Column & _
                                "列，所填行政区划名称与数据字典(行政区划.xls)不符"
            Else
                data_cell.Value2 = result_search
            End If
        
NextCellExchange:
        Next
        
    Case Else
        MsgBox "错误的存储数组名称"
End Select


End Sub

'处理电压值
'如果没有带kV,加上kV;
'data_table_range           被转换的数据文件数据集
'province_name              省名
'city_name                  地市名
'data_table_name            被转换的数据文件表名
'log_path                   日志文件路径
Sub update_voltage_value(data_table_range As Excel.Range _
                        , province_name As String _
                        , city_name As String _
                        , data_table_name As String _
                        , log_path As String)
Dim cells As Excel.Range
Dim cell_value As String
Dim reg As Object
Dim reg_nokv As Object

Set reg = CreateObject("vbscript.regexp")
With reg
    .Global = True
    .Pattern = "(^[±][0-9.]+[k][v]$)|(^[0-9.]+[k][v]$)|(^[±][0-9.]+$)|(^[0-9.]+$)"
End With

Set reg_nokv = CreateObject("vbscript.regexp")
With reg_nokv
    .Global = True
    .Pattern = "(^[±][0-9.]+$)|(^[0-9.]+$)"
End With

For Each cells In data_table_range
    
    If Not (cells.comment Is Nothing) Then
        If InStr(cells.comment.Text, Cell_RESOLVED) > 0 Then
            GoTo Next_cell_update_voltage_value
        End If
    End If
    
    cell_value = cells.Text
    cell_value = StrConv(cell_value, vbLowerCase)
    
    If reg.Test(cell_value) Then
        If reg_nokv.Test(cell_value) Then
            cells.Value2 = cells.Text & "kV"
            
            If cells.comment Is Nothing Then
                With cells.AddComment
                    .Visible = False
                    .Text Cell_RESOLVED
                End With
            End If
            
            cells.comment.Text Cell_RESOLVED

        End If
    ElseIf cell_value <> "" Then
        comment_cell cell:=cells _
                            , comment:="电压值格式不正确.列必须是以下四种格式之一:±数字kv,数字kv,±数字,数字。kv大小写都可以，不能带有其他字符"
    End If
Next_cell_update_voltage_value:
Next



Set reg = Nothing
Set reg_nokv = Nothing

End Sub

'生成不重复的序列
Sub update_serial(data_table_range As Excel.Range _
                        , province_name As String _
                        , city_name As String _
                        , data_table_name As String _
                        , log_path As String)
Dim serial As String
Dim city_code As String
Dim data_cell As Excel.Range
Dim i As Integer

Dim u As Double, su As Double
Dim hh As Integer, mm As Integer, ss As Integer, sss As Integer
Dim tmp_str As String
    
data_table_range.NumberFormat = "@"

For Each data_cell In data_table_range
    
    If Not (data_cell.comment Is Nothing) Then
        If InStr(data_cell.comment.Text, Cell_RESOLVED) > 0 Then
            GoTo Next_cell_update_serial
        End If
    End If

    u = Timer
    mm = Int(u / 60)
    su = u - mm * 60#
    ss = Int(su)
    sss = CLng((su - ss) * 1000)
    hh = Int(mm / 60)
    mm = Int(mm / 60)
    serial = CLng(Date) & hh & mm & ss & sss
    
    
    Randomize
    tmp_str = serial & Int(10 * Rnd)
    Randomize
    tmp_str = tmp_str & Int(10 * Rnd)
    Randomize
    tmp_str = tmp_str & Int(10 * Rnd)
    Randomize
    tmp_str = tmp_str & Int(10 * Rnd)
    Randomize
    data_cell.Value2 = tmp_str & Int(10 * Rnd)
    
    If data_cell.comment Is Nothing Then
        With data_cell.AddComment
            .Visible = False
            .Text Cell_RESOLVED
        End With
    Else
        data_cell.Text = Cell_RESOLVED
    End If
    
      
Next_cell_update_serial:
Next

End Sub

'分析字符串中是否有的电压等级,如果有，返回该电压等级的code
Function get_voltage(str As String)
Dim i As Integer
Dim reg As Object

Set reg = CreateObject("vbscript.regexp")
    
    '1000kV-----------------------------
    With reg
    .Global = True
    .Pattern = "(1000)"
    End With
    
    If reg.Test(str) Then
        get_voltage = 12
        GoTo end_get_voltage
    End If
    
    '750kV-----------------------------
    With reg
    .Global = True
    .Pattern = "(750)"
    End With
    
    If reg.Test(str) Then
        get_voltage = 17
        GoTo end_get_voltage
    End If
    
    '500kV-----------------------------
    With reg
    .Global = True
    .Pattern = "(500)"
    End With
    
    If reg.Test(str) Then
        get_voltage = 7
        GoTo end_get_voltage
    End If
    
    '330kV-----------------------------
    With reg
    .Global = True
    .Pattern = "(330)"
    End With
    
    If reg.Test(str) Then
        get_voltage = 13
        GoTo end_get_voltage
    End If
    
    '220kV-----------------------------
    With reg
    .Global = True
    .Pattern = "(220)"
    End With
    
    If reg.Test(str) Then
        get_voltage = 6
        GoTo end_get_voltage
    End If
    
    '110kV-----------------------------
    With reg
    .Global = True
    .Pattern = "(110)"
    End With
    
    If reg.Test(str) Then
        get_voltage = 5
        GoTo end_get_voltage
    End If
    
    '66kV-----------------------------
    With reg
    .Global = True
    .Pattern = "(66)"
    End With
    
    If reg.Test(str) Then
        get_voltage = 16
        GoTo end_get_voltage
    End If
    
    '35kV-----------------------------
    With reg
    .Global = True
    .Pattern = "(35)"
    End With
    
    If reg.Test(str) Then
        get_voltage = 4
        GoTo end_get_voltage
    End If
    
    '10kV-----------------------------
    With reg
    .Global = True
    .Pattern = "(10)"
    End With
    
    If reg.Test(str) Then
        get_voltage = 3
        GoTo end_get_voltage
    End If
    
    '±1100kV-----------------------------
    With reg
    .Global = True
    .Pattern = "(±1100)"
    End With
    
    If reg.Test(str) Then
        get_voltage = 30
        GoTo end_get_voltage
    End If
    
    '±800kV-----------------------------
    With reg
    .Global = True
    .Pattern = "(±800)"
    End With
    
    If reg.Test(str) Then
        get_voltage = 31
        GoTo end_get_voltage
    End If
    
    '±660kV-----------------------------
    With reg
    .Global = True
    .Pattern = "(±660)"
    End With
    
    If reg.Test(str) Then
        get_voltage = 32
        GoTo end_get_voltage
    End If
    
    '±500kV-----------------------------
    With reg
    .Global = True
    .Pattern = "(±500)"
    End With
    
    If reg.Test(str) Then
        get_voltage = 33
        GoTo end_get_voltage
    End If
    
    '±400kV-----------------------------
    With reg
    .Global = True
    .Pattern = "(±400)"
    End With
    
    If reg.Test(str) Then
        get_voltage = 34
        GoTo end_get_voltage
    End If
    
    '±250kV-----------------------------
    With reg
    .Global = True
    .Pattern = "(±250)"
    End With
    
    If reg.Test(str) Then
        get_voltage = 38
        GoTo end_get_voltage
    End If
    
    '±225kV-----------------------------
    With reg
    .Global = True
    .Pattern = "(±225)"
    End With
    
    If reg.Test(str) Then
        get_voltage = 37
        GoTo end_get_voltage
    End If
    
    '±125kV-----------------------------
    With reg
    .Global = True
    .Pattern = "(±125)"
    End With
    
    If reg.Test(str) Then
        get_voltage = 35
        GoTo end_get_voltage
    End If
    
    '±120kV-----------------------------
    With reg
    .Global = True
    .Pattern = "(±120)"
    End With
    
    If reg.Test(str) Then
        get_voltage = 36
        GoTo end_get_voltage
    End If


get_voltage = 2012      '2012 说明没有找到

end_get_voltage:
End Function

'分析字符串中是否有的变压器编号,如果有，返回该变压器编号
Function get_BYQ_NO(str As String)
Dim i As Integer
Dim reg As Object

Set reg = CreateObject("vbscript.regexp")
    
    '1
    With reg
    .Global = True
    .Pattern = "(1号)|(1主)|(1#)|(#1)|一"
    End With
    
    If reg.Test(str) Then
        get_BYQ_NO = 1
        GoTo end_get_BYQ_NO
    End If
    
    '2
    With reg
    .Global = True
    .Pattern = "(2号)|(2主)|(2#)|二"
    End With
    
    If reg.Test(str) Then
        get_BYQ_NO = 2
        GoTo end_get_BYQ_NO
    End If
    
    '3
    With reg
    .Global = True
    .Pattern = "(3号)|(3主)|(3#)|三"
    End With
    
    If reg.Test(str) Then
        get_BYQ_NO = 3
        GoTo end_get_BYQ_NO
    End If
    
    '4
    With reg
    .Global = True
    .Pattern = "(4号)|(4主)|(4#)|四"
    End With
    
    If reg.Test(str) Then
        get_BYQ_NO = 4
        GoTo end_get_BYQ_NO
    End If
    
get_BYQ_NO = 2012
    
end_get_BYQ_NO:
End Function



'获得变电站名称，线路名称中有关地点的信息
'str        输入为小写
Function get_addressinfo(str As String)
Dim i As Integer


'str = Replace(str, "330", "")
'str = Replace(str, "750", "")
'str = Replace(str, "500", "")
'str = Replace(str, "220", "")
'str = Replace(str, "110", "")
'str = Replace(str, "66", "")
'str = Replace(str, "35", "")

str = Replace(str, "kv", "")
str = Replace(str, "变电", "")

str = Replace(str, "风力", "")
str = Replace(str, "风", "")
str = Replace(str, "发电", "")
str = Replace(str, "热电", "")
str = Replace(str, "水电", "")
str = Replace(str, "火电", "")
str = Replace(str, "电", "")
str = Replace(str, "厂", "")
str = Replace(str, "场", "")


str = Replace(str, "站", "")
str = Replace(str, "主变", "")
str = Replace(str, "变", "")
str = Replace(str, "主", "")
str = Replace(str, "压器", "")
str = Replace(str, "路", "")
str = Replace(str, "线", "")
str = Replace(str, "回", "")
str = Replace(str, "号", "")
str = Replace(str, "-", "")
str = Replace(str, "一", "")
str = Replace(str, "二", "")
str = Replace(str, "三", "")
str = Replace(str, "i", "")

For i = 0 To 9
    str = Replace(str, i, "")
Next

get_addressinfo = str

End Function

'获得数据字典的编码
Function get_code(str As String, code_type As Integer)

End Function

'判断10kV以下线路开关号
Function get_SWITCH_NO(str As String)
Dim i As Integer

For i = 1 To Len(str)
    If IsNumeric(Mid(str, i, 1)) Then
        get_SWITCH_NO = Val(Right(str, Len(str) - i + 1))
        GoTo end_get_SWITCH_NO
    End If
Next

get_SWITCH_NO = CInt(2012)

end_get_SWITCH_NO:
End Function


'根据子表的值在父表中查询，将子表的外键替换为父表的主键值
'slave_table_range                  子表数据集
'slave_foreign_col                  子表外键列
'slave_table_name                   子表名
'master_table_name                  父表名
'master_start_row                   父表记录起始行
'master_mainkey_col                 父表主键列
'slave_verify_col()                 0 子表区域列号，1 子表电压列号
'master_verify_col()                0 父表区域列号，1 父表电压列号,2 父表设备名称列号
'province_name                      省名
'city_name                          地市名
'log_path                           日志文件路径
Sub master_to_slave(slave_table_range As Excel.Range _
                    , slave_forkey_range As Excel.Range _
                    , slave_table_name As String _
                    , master_table_name As String _
                    , master_start_row As Integer _
                    , master_mainkey_col As Integer _
                    , slave_verify_col() As Integer _
                    , master_verify_col() As Integer _
                    , province_name As String _
                    , city_name As String _
                    , log_path As String)
                    
Dim slave_cell As Excel.Range
Dim master_cell As Excel.Range

Dim master_wb As Excel.Workbook
Dim master_ws As Excel.Worksheet
Dim master_mainkey_range As Excel.Range
Dim master_table_range As Excel.Range
Dim master_table_bottom As Integer

Dim search_value As String

Dim m_count As Long
Dim s_count As Long

Dim finded_count As Long                '最后发现主表记录匹配的位置
Dim finded_flag As Boolean              '

finded_count = 1
finded_flag = False

Dim master_str As String
Dim i As Integer

Dim search_value_voltage As Integer, master_str_voltage As Integer
Dim search_address_info As String, master_address_info As String
Dim search_no As Integer, master_no As Integer

Dim slave_verify_area(1) As String, slave_verify_voltage(1) As String
Dim master_verify_area, master_verify_voltage


Set master_wb = Workbooks.Open(slave_table_range.Parent.Parent.path & "/" & master_table_name & ".xls")
Set master_ws = master_wb.Worksheets(1)

Dim temp_array() As Variant

master_table_bottom = compare_bottom_row(master_ws)

Set master_table_range = _
            master_ws.Range(master_ws.cells(master_start_row, master_verify_col(2)), master_ws.cells(master_table_bottom, master_verify_col(2)))

Set master_mainkey_range = _
            master_ws.Range(master_ws.cells(master_start_row, master_mainkey_col), master_ws.cells(master_table_bottom, master_mainkey_col))

s_count = 1
Do While s_count <= slave_table_range.Count
    
    Set slave_cell = slave_table_range.Item(s_count)
    slave_cell.NumberFormatLocal = "@"
    
    
    '从上次循环发现的位置开始
    m_count = finded_count
    finded_flag = True
    Do While m_count <= master_table_range.Count
    
        search_value = StrConv(slave_cell.Text, vbLowerCase)
        
        Set master_cell = master_table_range.Item(m_count)
        master_str = StrConv(master_cell.Text, vbLowerCase)
        
        search_value_voltage = get_voltage(search_value)
        master_str_voltage = get_voltage(master_str)
        
        search_address_info = Trim(get_addressinfo(search_value))
        master_address_info = Trim(get_addressinfo(master_str))
        
        '变压器运行数据需要判断变压器编号
        If slave_table_range.Parent.Name = "变压器运行数据" And _
            slave_table_range.Column = 7 Then
            search_no = Trim(get_BYQ_NO(StrConv(slave_cell.Text, vbLowerCase)))
            master_no = Trim(get_BYQ_NO(StrConv(master_cell.Text, vbLowerCase)))
        End If
        
        '10kV线路运行数据需要判断出口开关的开关号
        If slave_table_range.Parent.Parent.Name = "10(20、6)kV线路运行数据.xls" And _
            slave_table_range.Column = 7 Then
            search_no = get_SWITCH_NO(StrConv(slave_cell.Text, vbLowerCase))
            master_no = get_SWITCH_NO(StrConv(master_cell.Text, vbLowerCase))
        End If
        
        slave_verify_area(0) = slave_cell.Parent.cells(slave_cell.row, slave_verify_col(0)).Text
        master_verify_area = master_cell.Parent.cells(master_cell.row, master_verify_col(0)).Text
        
        slave_verify_voltage(0) = slave_cell.Parent.cells(slave_cell.row, slave_verify_col(1)).Text
        master_verify_voltage = master_cell.Parent.cells(master_cell.row, master_verify_col(1)).Text
        
        '字符串里是否有电压信息
        '如果有电压信息，则判断是否相同；再判断地理信息是否相同；再判断区域信息是否相同；如果三者都相同，匹配
        '如果没有电压信息，则从电压列调信息
        If search_value_voltage <> 2012 And master_str_voltage <> 2012 Then
            
            '判断电压是否相等
            If master_str_voltage = search_value_voltage Then
            
verify_voltage_equal:
                '判断地理信息是否相同,如果是变压器匹配，还要判断变压器编号是否相同
                If (StrComp(search_address_info, master_address_info) = 0) And _
                    search_no = master_no Then
                    
                    '判断区域是否相同
                    If ((master_verify_area = slave_verify_area(0)) Or _
                    (master_verify_area = array_substr_single(gov_code(), slave_verify_area()))) Then
                        
                        If slave_forkey_range.cells(s_count).comment Is Nothing Then
                            slave_forkey_range.cells(s_count).AddComment
                            slave_forkey_range.cells(s_count).comment.Visible = False
                        End If
                        
                        slave_forkey_range.cells(s_count).comment.Text _
                            Chr(10) & "原值为:" & slave_cell.Text & ",匹配值为:" & master_cell.Text
                        
                        
                        slave_forkey_range.cells(s_count).Value2 = _
                            CStr(master_mainkey_range.cells(m_count).Value2)
                            
                        '记录主表发现的位置，下次循环优先使用这个位置
                        finded_count = m_count
                        
                        slave_forkey_range.cells(s_count).NumberFormat = "@"
                            
                        GoTo NextSerach
                    '如果不相同，但如果是包含关系，则也替换
                    '如果区域信息的后两位为0，说明是一个较宽泛层次区域，则替换为较狭窄层次区域
                    ElseIf Right(master_verify_area, 2) = "00" Or _
                            Right(slave_verify_area(0), 2) = "00" Or _
                            Right(array_substr_single(gov_code(), slave_verify_area()), 2) = "00" Then
                                
                        slave_forkey_range.cells(s_count).NumberFormatLocal = "@"
                        
                        If slave_forkey_range.cells(s_count).comment Is Nothing Then
                            slave_forkey_range.cells(s_count).AddComment
                            slave_forkey_range.cells(s_count).comment.Visible = False
                        End If
                        
                        slave_forkey_range.cells(s_count).comment.Text _
                            Chr(10) & "原值为:" & slave_cell.Text & ",匹配值为:" & master_cell.Text
                        
                        
                        slave_forkey_range.cells(s_count).Value2 = _
                            master_mainkey_range.cells(m_count).Value2
                            
                        GoTo NextSerach

                    End If
                Else
                    '如果finded_flag为true，说明是使用最后一次主表位置查找没有找到而跳到下次循环
                    If finded_flag Then
                        m_count = 0
                        finded_flag = False
                    End If
                    GoTo next_master_search
                End If
                
            Else
                If finded_flag Then
                        m_count = 0
                        finded_flag = False
                End If
                GoTo next_master_search
            End If
        
        ElseIf ( _
                (master_verify_voltage = slave_verify_voltage(0)) Or _
                (master_verify_voltage = CDbl(get_voltage(slave_verify_voltage(0)))) _
                ) Then
                
                '设备名称里没有电压信息，但是确认用电压等级相同,跳到电压相同处理方式
                GoTo verify_voltage_equal
                
        '这里的slave_verify_voltage处理的是发电厂类型信息
        ElseIf ( _
                (master_verify_voltage = slave_verify_voltage(0)) Or _
                (master_verify_voltage = array_substr_multi(pub_domain_code(), slave_verify_voltage(0), 4194537)) _
                ) Then
                
                '设备名称里没有电压信息，但是确认用电压等级相同,跳到电压相同处理方式
                GoTo verify_voltage_equal
        
        '站内母线第10列-所属变电站
        '开关台帐第10列-所属变电站
        '并联电容电抗器第5列-所属变电站
        '不需要匹配电压信息，因此不用电压信息匹配
        ElseIf (slave_table_range.Parent.Name = "站内母线" And slave_table_range.Column = 10) Or _
                (slave_table_range.Parent.Name = "开关台账" And slave_table_range.Column = 10) Or _
                (slave_table_range.Parent.Name = "并联电容电抗器" And slave_table_range.Column = 5) Then
            GoTo verify_voltage_equal
            
        End If
        
next_master_search:
    m_count = m_count + 1
    Loop
    
    If m_count > master_table_range.Count Then
        comment_cell cell:=slave_cell _
                        , comment:=Chr(10) & "变电站/线路不存在。"
    End If
    
NextSerach:
s_count = s_count + 1
Loop

master_wb.Close saveChanges:=False
End Sub

'非空检查
Sub check_No_Null(data_table_range As Excel.Range _
                        , province_name As String _
                        , city_name As String _
                        , data_table_name As String _
                        , log_path As String)

Dim data_cell As Excel.Range

For Each data_cell In data_table_range

    If data_cell.Text = "" Then
        Call log.error(log_path:=log_path _
                                , msg:="该列必填." & city_name & "表" & data_table_name & "第" & data_cell.Column & "列必填")
            comment_cell cell:=data_cell _
                                , comment:="该列必填." & city_name & "表" & data_table_name & "第" & data_cell.Column & "列必填"
    End If

Next

End Sub

'清空所有单元格
Sub value_to_null(data_table_range As Excel.Range _
                        , province_name As String _
                        , city_name As String _
                        , data_table_name As String _
                        , log_path As String)
                        
Dim data_cell As Excel.Range

For Each data_cell In data_table_range

    data_cell.Value2 = ""

Next

End Sub

'填入指定的值,
'用于填入area_type,省填1，地市填2
Sub input_special_value(data_table_range As Excel.Range _
                        , spec_value As String)
                        
Dim data_cell As Excel.Range

For Each data_cell In data_table_range

    data_cell.Value2 = spec_value
Next
End Sub


'输入DEPTID
Sub input_deptid(data_table_range As Excel.Range _
                        , unit_name As String)

Dim org_id As String
Dim i As Integer
Dim main_ws As Excel.Worksheet

Set main_ws = Workbooks("checkscript.xlsm").Worksheets(2)

i = 2
Do While main_ws.cells(i, 1).Text <> ""

    If StrComp(main_ws.cells(i, 1).Text, unit_name, vbTextCompare) = 0 Then
        org_id = main_ws.cells(i, 2).Text
        Exit Do
    End If

    i = i + 1
Loop

input_special_value data_table_range, org_id

End Sub

'根据确认列，将子表的值替换为父表的值,使用：电抗器
'slave_foreign_col                  子表外键列
'slave_table_name                   子表名
'master_table_name                  父表名
'master_start_row                   父表记录起始行
'master_mainkey_col                 父表主键列
'slave_verify_col()                 0 子表所属变电站编号，1 子表所连母线名称
'master_verify_col()                0 父表变电站编号，1 父表母线名称,2 未使用
'province_name                      省名
'city_name                          地市名
'log_path                           日志文件路径
Sub master_to_slave_dkq(slave_forkey_range As Excel.Range _
                    , slave_table_name As String _
                    , master_table_name As String _
                    , master_start_row As Integer _
                    , master_mainkey_col As Integer _
                    , slave_verify_col() As Integer _
                    , master_verify_col() As Integer _
                    , province_name As String _
                    , city_name As String _
                    , log_path As String)
                    
'Dim slave_cell As Excel.Range
'Dim master_cell As Excel.Range

Dim master_wb As Excel.Workbook
Dim master_ws As Excel.Worksheet
Dim master_mainkey_range As Excel.Range
Dim master_table_bottom As Integer

Dim slave_ws As Excel.Worksheet
Dim slave_table_start As Integer
Dim slave_table_bottom As Long


Dim slave_verify_range_zero As Excel.Range, slave_verify_range_one As Excel.Range
Dim master_verify_range_zero As Excel.Range, master_verify_range_one As Excel.Range

Dim s_count As Integer, m_count As Integer


Set master_wb = Workbooks.Open(slave_forkey_range.Parent.Parent.path & "/" & master_table_name & ".xls")
Set master_ws = master_wb.Worksheets(1)

master_table_bottom = compare_bottom_row(master_ws)

Set master_mainkey_range = _
            master_ws.Range(master_ws.cells(master_start_row, master_mainkey_col), _
            master_ws.cells(master_table_bottom, master_mainkey_col))

Set master_verify_range_zero = _
            master_ws.Range(master_ws.cells(master_start_row, master_verify_col(0)), _
            master_ws.cells(master_table_bottom, master_verify_col(0)))

If master_verify_col(1) <> 0 Then
    Set master_verify_range_one = _
                master_ws.Range(master_ws.cells(master_start_row, master_verify_col(1)), _
                master_ws.cells(master_table_bottom, master_verify_col(1)))
End If
            
Set slave_ws = slave_forkey_range.Parent

slave_table_start = slave_forkey_range.Item(0).row
slave_table_bottom = slave_forkey_range.Item(slave_forkey_range.Count).row

Set slave_verify_range_zero = _
            slave_ws.Range(slave_ws.cells(4, slave_verify_col(0)), _
                            slave_ws.cells(slave_table_bottom, slave_verify_col(0)))
                            
If slave_verify_col(1) <> 0 Then
    Set slave_verify_range_one = _
                slave_ws.Range(slave_ws.cells(4, slave_verify_col(1)), _
                                slave_ws.cells(slave_table_bottom, slave_verify_col(1)))
End If

s_count = 1
Do While s_count <= slave_forkey_range.Count

    
    m_count = 1
    Do While m_count <= master_mainkey_range.Count

        If StrComp(CStr(slave_verify_range_one(s_count)), _
                    CStr(master_verify_range_one(m_count))) = 0 Then
VERIFY_ONE_NO_NEED:
            If StrComp(CStr(slave_verify_range_zero(s_count).Value2), _
                        CStr(master_verify_range_zero(m_count).Value2)) = 0 Then
                
                slave_forkey_range.cells(s_count).NumberFormatLocal = "@"
                        
                        If slave_forkey_range.cells(s_count).comment Is Nothing Then
                            slave_forkey_range.cells(s_count).AddComment
                            slave_forkey_range.cells(s_count).comment.Visible = False
                        End If
                        
                        slave_forkey_range.cells(s_count).comment.Text _
                            Chr(10) & "原值为:" & slave_verify_range_one(s_count).Text & _
                            ",匹配值为:" & master_verify_range_one(m_count).Text
                        
                        
                        slave_forkey_range.cells(s_count).Value2 = _
                            master_mainkey_range.cells(m_count).Value2
                            
                        GoTo NextSerach
                        
            End If
        '站内母线-第23列投运日期不需要确认第一项。
        ElseIf (slave_forkey_range.Parent.Name = "站内母线" And slave_forkey_range.Column = 23) Then
            GoTo VERIFY_ONE_NO_NEED
        End If
    
    m_count = m_count + 1
    Loop
    
    If m_count > slave_forkey_range.Count Then
        comment_cell cell:=slave_verify_range_one(s_count) _
                        , comment:=Chr(10) & "找不到所属母线/变电站。"
    End If
    
NextSerach:
s_count = s_count + 1
Loop

master_wb.Close saveChanges:=False
End Sub