Attribute VB_Name = "QueryRunner"
Public QRDBIndex%

'***********************************
' Запросы чувствительны к регистру!
'***********************************

' константы видов запросов
    ' ОБЯЗАТЕЛЬНО 3 ЗНАКА
Public Const sAdd$ = "Add"
Public Const sDel$ = "Del"
Public Const sSort$ = "Srt"
Public Const sOut$ = "Out"
Public Const sSwap$ = "Swp"
Public Const sChange$ = "Chg"

' константы подтипов запросов
Public Const sCol$ = "Col"
Public Const sRow$ = "Row"
Public Const sTable$ = "Tbl"    '   только для использования в запросе Вывод
Public Const sAZ$ = "AZ"
Public Const sZA$ = "ZA"
Public Const sEqual$ = "?="
Public Const sAbove$ = "?>"
Public Const sBelow$ = "?<"
Public Const sCountEqual$ = "+="
Public Const sCountAbove$ = "+>"
Public Const sCountBelow$ = "+<"
Public Const sI$ = "i"
Public Const sS$ = "s"
Public Const sYes$ = "yes"
Public Const sNo$ = "no"
Public Const sType$ = "Type"
Public Const sName$ = "Name"

' остальные константы
Public Const sSep$ = ";"

'************************ Формирует строку добавления 'What' ************************
Public Function Generate_Add(ByVal what$) As String
    If (what = sCol) Then
        s$ = AddColForm.AddColDlg(QRDBIndex)
        If (s <> "") Then
            Generate_Add = sAdd + sCol + "( " + s + " )"
        Else
            Generate_Add = ""
        End If
    Else
        Generate_Add = sAdd + sRow + "()"
    End If
End Function

'************************ Формирует строку удаления 'What' ************************
Public Function Generate_Del(ByVal what$) As String
    With SelectForm.CheckConfirm
        .value = 1
        .Visible = True
    End With
    Dim conf$
    
    If (what = sCol) Then
        s$ = SelectForm.SelectDlg(QRDBIndex, "Выберите удаляемое поле", sCol)
        If (s <> -1) Then
            If (SelectForm.CheckConfirm.value = 1) Then
                conf = sYes
            Else
                conf = sNo
            End If
            Generate_Del = sDel + sCol + "( " + s + " , " + conf + " )"
        Else
            Generate_Del = ""
        End If
    Else
        s$ = SelectForm.SelectDlg(QRDBIndex, "Выберите удаляемую запись", sRow)
        If (s <> -1) Then
            If (SelectForm.CheckConfirm.value = 1) Then
                conf = sYes
            Else
                conf = sNo
            End If
            Generate_Del = sDel + sRow + "( " + s + " , " + conf + " )"
        Else
            Generate_Del = ""
        End If
    End If
End Function

'************************ Формирует строку сортировки по 'What' ************************
Public Function Generate_Sort(ByVal what$) As String
    SelectForm.CheckConfirm.Visible = False

    s$ = SelectForm.SelectDlg(QRDBIndex, "Выберите поле сортировки", sCol)
    If (s <> -1) Then
        Generate_Sort = sSort + "( " + s + " , " + what + " )"
    Else
        Generate_Sort = ""
    End If
End Function

'************************ Формирует строку вывода по 'What' ************************
Public Function Generate_Out(ByVal what$) As String
    Generate_Out = ""
    SelectForm.CheckConfirm.Visible = False
    Dim str$
    
    s$ = SelectForm.SelectDlg(QRDBIndex, "Выберите поле", sCol)
    If (s <> "-1") Then
        str = Trim(InputForm.InputVal("Введите относительное значение"))
        If (str <> "") Then
            Dim CreateNewTab As Boolean
            CreateNewTab = (MsgForm.QuestMsg("Выводить в новую таблицу? Нет для вывода в уже существующую.") = resOk)
            If (Not CreateNewTab) Then
                Table$ = SelectForm.SelectDlg(QRDBIndex, "Выберите таблицу", sTable)
                If (Table = "-1") Then Exit Function
                Generate_Out = sOut + "( " + s + " , " + what + str + " , " + Table + " )"
            Else
                Generate_Out = sOut + "( " + s + " , " + what + str + " )"
            End If
        Else
            Call MsgForm.ErrorMsg("Не задано относительное значение!")
        End If
    End If
End Function

'************************ Формирует строку обмена по 'What' ************************
Public Function Generate_Swap(ByVal what$) As String
    If (what = sCol) Then
        s$ = SelectForm.MultiSelectDlg(QRDBIndex, "Выберите 2 обмениваемых поля", sCol)
        If (s <> "") Then
            p% = InStr(1, s, ",")
            Generate_Swap = sSwap + sCol + "( " + Left(s, p - 1) + " , " + Mid(s, p + 1) + " )"
        Else
            Generate_Swap = ""
        End If
    Else
        s$ = SelectForm.MultiSelectDlg(QRDBIndex, "Выберите 2 обмениваемые записи", sRow)
        If (s <> "") Then
            p% = InStr(1, s, ",")
            Generate_Swap = sSwap + sRow + "( " + Left(s, p - 1) + " , " + Mid(s, p + 1) + " )"
        Else
            Generate_Swap = ""
        End If
    End If
End Function

'************************ Формирует строку изменения 'What' ************************
Public Function Generate_Change(ByVal what$) As String
    Generate_Change = ""
    SelectForm.CheckConfirm.Visible = False
    
    s$ = SelectForm.SelectDlg(QRDBIndex, "Выберите изменяемое поле", sCol)
    If (s = "-1") Then Exit Function
    Select Case what
        Case sType  '   Изменение типа поля
            Generate_Change = sChange + sType + "( " + s + " )"
        Case sName  '   Изменение названия столбца
            Name$ = InputForm.InputVal("Введите новое название поля")
            If (Name = "") Then Exit Function
            Generate_Change = sChange + sName + "( " + s + " , " + Name + " )"
    End Select
End Function

Sub ErrorInQuery()
    Call MsgForm.ErrorMsg("Ошибка в запросе!")
End Sub

Function TestZero(i%)
    If (i = 0) Then
        Call ErrorInQuery
        TestZero = True
    Else
        TestZero = False
    End If
End Function

Sub AddRun(what$, str$)
    Select Case what
        Case sCol
            ' заголовок
            p% = InStr(1, str, ",")
            If TestZero(p) Then Exit Sub
            title$ = Trim(Left(str, p - 1))
            str = Mid(str, p + 1)
            ' тип
            p = InStr(1, str, ",")
            If TestZero(p) Then Exit Sub
            ColType$ = Trim(Left(str, p - 1))
            str = Mid(str, p + 1)

            ' начальное значение
            p = InStr(1, str, ",")
            If TestZero(p) Then Exit Sub
            StValStr$ = Trim(Left(str, p - 1))
            str = Mid(str, p + 1)
            
            ' позиция
            ColPosStr$ = str
            If (Not IsNumeric(ColPosStr)) Then
                Call ErrorInQuery
                Exit Sub
            End If
            ColPos% = CInt(ColPosStr)
            
            If ItColAlreadyCreate(QRDBIndex, title) Then
                Call MsgForm.ErrorMsg("Добавляемое поле уже существует!")
                Exit Sub
            End If
            
            ' в зависимости от типа определяю значение
            Select Case ColType
                Case sI
                    If (Not IsInteger(StValStr)) Then
                        Call ErrorInQuery
                        Exit Sub
                    End If
                    stval = CInt(StValStr)
                    Call AddCol(QRDBIndex, ccInteger, title, stval, ColPos)
                Case sS
                    stval = CStr(StValStr)
                    Call AddCol(QRDBIndex, ccString, title, stval, ColPos)
                Case Default
                    Call ErrorInQuery
                    Exit Sub
            End Select

        Case sRow
            If (DB(QRDBIndex).Header.ColCount > 0) Then
                Dim row() As Variant
                ReDim row(DB(QRDBIndex).Header.ColCount - 1)
                For i = 0 To DB(QRDBIndex).Header.ColCount - 1
                    row(i) = DB(QRDBIndex).Cols(i).DefValue
                Next i
                If (Not FindRow(QRDBIndex, row)) Then
                    Call AddField(QRDBIndex, row)
                Else
                    Call MsgForm.ErrorMsg("Добавляемый столбец дублируется!")
                End If
            Else
                Call MsgForm.ErrorMsg("Нельзя добавлять записи в БД без полей!")
            End If
    End Select
    
End Sub

Sub DelRun(what$, str$)
    p% = InStr(1, str, ",")
    If TestZero(p) Then Exit Sub
    IndexStr$ = Trim(Left(str, p - 1))
    If (Not IsInteger(IndexStr)) Then
        Call ErrorInQuery
        Exit Sub
    End If
    Index% = CInt(IndexStr)
    str = Mid(str, p + 1)
    ConfirmStr$ = Trim(str)
    Dim Confirm As Boolean
    Select Case ConfirmStr
        Case sYes
            Confirm = True
        Case sNo
            Confirm = False
        Case Default
            Call ErrorInQuery
            Exit Sub
    End Select
    
    Select Case what
        Case sCol
            If (DB(QRDBIndex).Header.ColCount > 0) Then
                Call DelCol_(QRDBIndex, Index, Confirm)
            Else
                Call MsgForm.ErrorMsg("В БД нет полей!")
                Exit Sub
            End If
        Case sRow
            If (DB(QRDBIndex).Header.RowCount > 0) Then
                Call DelRow_(QRDBIndex, Index, Confirm)
            Else
                Call MsgForm.ErrorMsg("В БД нет записей!")
                Exit Sub
            End If
    End Select
End Sub

Sub SortRun(str$)
    If (DB(QRDBIndex).Header.ColCount = 0) Or (DB(QRDBIndex).Header.RowCount = 0) Then
        Call MsgForm.ErrorMsg("Нечего сортировать!")
        Exit Sub
    End If
    
    p% = InStr(1, str, ",")
    If TestZero(p) Then Exit Sub
    what$ = Trim(Left(str, p - 1))
    
    If (Not IsInteger(what)) Then
        Call ErrorInQuery
        Exit Sub
    End If
    
    whatint% = CInt(what)
    
    If (whatint < 0) Or (whatint > DB(QRDBIndex).Header.ColCount - 1) Then
        Call ErrorInQuery
        Exit Sub
    End If
    
    Mode$ = Trim(Mid(str, p + 1))
    
    Select Case Mode
        Case sAZ
            s$ = "А->Я"
        Case sZA
            s$ = "Я->А"
        Case Default
            Call ErrorInQuery
            Exit Sub
    End Select
    
    Count% = MainForm.TabStrip.Tabs.Count
    ReDim Preserve DB(Count)
    
    DB(Count) = DB(QRDBIndex)
    
    MainForm.TabStrip.Tabs.Add pvCaption:=s, pvImage:=1
    
    Dim find As Boolean, needswap As Boolean
    Dim tmp As TDBElem
    With DB(Count)
        Do
            find = False
            For R% = 1 To .Header.RowCount - 1
                If (Mode = sZA) Then
                    needswap = (.Rows(R).Fields(whatint) > .Rows(R - 1).Fields(whatint))
                Else
                    needswap = (.Rows(R).Fields(whatint) < .Rows(R - 1).Fields(whatint))
                End If
                If (needswap) Then
                    tmp = .Rows(R)
                    .Rows(R) = .Rows(R - 1)
                    .Rows(R - 1) = tmp
                    find = True
                End If
            Next R
        Loop While (find)
    End With
End Sub

Function Equal(ByVal col%, ByVal row%, ByVal cmpstr$) As Long
    If (DB(QRDBIndex).Cols(col).Class = ccInteger) Then
        Rval = CLng(DB(QRDBIndex).Rows(row).Fields(col))
        Equal = (Rval - CLng(cmpstr))
    Else
        Rval = CStr(DB(QRDBIndex).Rows(row).Fields(col))
        If (Rval = cmpstr) Then
            Equal = 0
        Else
            If (Rval > cmpstr) Then
                Equal = 1
            Else
                Equal = -1
            End If
        End If
    End If
End Function

Function CalcCount(Index%, c%, value$) As Integer
    Count% = 0
    For i% = 0 To (DB(Index).Header.RowCount - 1)
        If (CStr(DB(Index).Rows(i).Fields(c)) = value) Then Count = Count + 1
    Next i
    CalcCount = Count
End Function

Function EarlierDontFind(Index%, c%, R%, value$) As Boolean
    For i% = 0 To (R - 1)
        If (CStr(DB(Index).Rows(i).Fields(c)) = value) Then
            EarlierDontFind = False
            Exit Function
        End If
    Next i
    EarlierDontFind = True
End Function

Public Function FindRow(Index%, row())
    For R% = 0 To DB(Index).Header.RowCount - 1
        Sum% = 0
        For c% = 0 To DB(Index).Header.ColCount - 1
            If (CStr(DB(Index).Rows(R).Fields(c)) = row(c)) Then Sum = Sum + 1
        Next c
        If (Sum = DB(Index).Header.ColCount) Then
            FindRow = True
            Exit Function
        End If
    Next R
    FindRow = False
End Function

Sub OutRun(str$)
    If (DB(QRDBIndex).Header.ColCount = 0) Or (DB(QRDBIndex).Header.RowCount = 0) Then
        Call MsgForm.ErrorMsg("Не с чем сравнивать!")
        Exit Sub
    End If
    
    p% = InStr(1, str, ",")
    what$ = Trim(Left(str, p - 1))
    
    If (Not IsInteger(what)) Then
        Call ErrorInQuery
        Exit Sub
    End If
    
    whatint% = CInt(what)
    
    If (whatint < 0) Or (whatint > DB(QRDBIndex).Header.ColCount - 1) Then
        Call ErrorInQuery
        Exit Sub
    End If
    
    pi% = p + 1
    Do
        Mode$ = Trim(Mid(str, pi, 1))
        pi = pi + 1
    Loop While (Mode = "")
    Mode = Mode + Mid(str, pi, 1)
    
    If (Mode <> sEqual) And (Mode <> sAbove) And (Mode <> sBelow) And (Mode <> sCountEqual) And (Mode <> sCountAbove) And (Mode <> sCountBelow) Then
        Call ErrorInQuery
        Exit Sub
    End If
    
    Dim CalcMode As Boolean
    CalcMode = (Mode = sCountEqual) Or (Mode = sCountAbove) Or (Mode = sCountBelow)
    
    str = Trim(Mid(str, pi + 1))
    
    If (str = "") Then
        Call ErrorInQuery
        Exit Sub
    End If
    
    ' проверка на наличие индекса таблицы
    p = InStr(1, str, ",")
    tableindex% = -1
    If (p <> 0) Then
        tableindexstr$ = Trim(Mid(str, p + 1))
        If Not IsInteger(tableindexstr) Then
           Call ErrorInQuery
           Exit Sub
        End If
        tableindex% = CLng(tableindexstr)
        If (tableindex < 0) Or (tableindex > MainForm.TabStrip.Tabs.Count - 1) Then
           Call ErrorInQuery
           Exit Sub
        End If
        str = Trim(Left(str, p - 1))
    End If
    
    Dim GlobEqual As Boolean
    If (Not IsInteger(str)) And (DB(QRDBIndex).Cols(whatint).Class = ccInteger) Then
        Call MsgForm.ErrorMsg("Эквивалентом вывода целочисленного столбца не является целое число!" + vbCrLf + _
                              "Условие всегда истинно!")
        GlobEqual = True
    Else
        GlobEqual = False
    End If
    
    Count% = MainForm.TabStrip.Tabs.Count
    If (tableindex = -1) Then
        ReDim Preserve DB(Count)
        
        DB(Count).Header = DB(QRDBIndex).Header
        DB(Count).Header.RowCount = 0
        DB(Count).Cols = DB(QRDBIndex).Cols
    
        MainForm.TabStrip.Tabs.Add pvCaption:="Вывод " + Mode + str, pvImage:=1
    Else
        Count = tableindex
    End If
    
    Dim NeedAdd As Boolean
    With DB(Count)
        Dim Rval
        For R% = 0 To DB(QRDBIndex).Header.RowCount - 1
            If (Not GlobEqual) Then
                Select Case Mode
                    Case sEqual
                        NeedAdd = (Equal(whatint, R, str) = 0)
                    Case sAbove
                        NeedAdd = (Equal(whatint, R, str) > 0)
                    Case sBelow
                        NeedAdd = (Equal(whatint, R, str) < 0)
                    Case sCountEqual
                        value$ = CStr(DB(QRDBIndex).Rows(R).Fields(whatint))
                        NeedAdd = ((CStr(CalcCount(QRDBIndex, whatint, value)) = str) And (EarlierDontFind(QRDBIndex, whatint, R, value)))
                    Case sCountAbove
                        value$ = CStr(DB(QRDBIndex).Rows(R).Fields(whatint))
                        NeedAdd = ((CStr(CalcCount(QRDBIndex, whatint, value)) > str) And (EarlierDontFind(QRDBIndex, whatint, R, value)))
                    Case sCountBelow
                        value$ = CStr(DB(QRDBIndex).Rows(R).Fields(whatint))
                        NeedAdd = ((CStr(CalcCount(QRDBIndex, whatint, value)) < str) And (EarlierDontFind(QRDBIndex, whatint, R, value)))
                End Select
            Else
                NeedAdd = True
            End If
            If (NeedAdd) Then
                ReDim tmparr(DB(QRDBIndex).Header.ColCount)
                tmparr = DB(QRDBIndex).Rows(R).Fields
                If (Not FindRow(Count, tmparr)) Then
                    addindex% = DB(Count).Header.RowCount
                    ReDim Preserve DB(Count).Rows(addindex)
                    ReDim DB(Count).Rows(addindex).Fields(DB(Count).Header.ColCount - 1)
                    DB(Count).Rows(addindex).Fields = DB(QRDBIndex).Rows(R).Fields
                    DB(Count).Header.RowCount = DB(Count).Header.RowCount + 1
                Else
                    Call MsgForm.ErrorMsg("Добавляемая запись уже существует!")
                End If
            End If
        Next R
    End With
End Sub

Sub SwapRun(what$, str$)
    p% = InStr(1, str, ",")
    If TestZero(p) Then Exit Sub
    index1str$ = Trim(Left(str, p - 1))
    index2str$ = Trim(Mid(str, p + 1))
    
    If (Not IsInteger(index1str)) Then
        Call ErrorInQuery
        Exit Sub
    End If
    
    index1% = CInt(index1str)
    index2% = CInt(index2str)
    
    If (index1 < 0) Or (index2 < 0) Or (index1 = index2) Then
        Call ErrorInQuery
        Exit Sub
    End If
    
    Select Case what
        Case sCol
            With DB(QRDBIndex)
                If (index1 > .Header.ColCount - 1) Or (index2 > .Header.ColCount - 1) Then
                    Call ErrorInQuery
                    Exit Sub
                End If
                ' обмен полей
                Dim tmpcol As TDBElemData
                tmpcol = .Cols(index1)
                .Cols(index1) = .Cols(index2)
                .Cols(index2) = tmpcol
                ' обмен полей записей
                Dim tmpcell As Variant
                For R% = 0 To .Header.RowCount - 1
                    tmpcell = .Rows(R).Fields(index1)
                    .Rows(R).Fields(index1) = .Rows(R).Fields(index2)
                    .Rows(R).Fields(index2) = tmpcell
                Next R
                
            End With
        Case sRow
            With DB(QRDBIndex)
                If (index1 > .Header.RowCount - 1) Or (index2 > .Header.RowCount - 1) Then
                    Call ErrorInQuery
                    Exit Sub
                End If
                Dim tmprow As TDBElem
                tmprow = .Rows(index1)
                .Rows(index1) = .Rows(index2)
                .Rows(index2) = tmprow
            End With
    End Select
End Sub

Sub ChangeRun(what$, param$)
    Select Case what
        Case sType      '   ************** ...::: Type :::... ***************
            If Not IsInteger(param) Then
                Call ErrorInQuery
                Exit Sub
            End If
            colindex% = CLng(param)
            If (colindex < 0) Or (colindex > DB(QRDBIndex).Header.ColCount - 1) Then
                Call ErrorInQuery
                Exit Sub
            End If
            If (DB(QRDBIndex).Cols(colindex).Class = ccString) Then
                If (MsgForm.QuestMsg("Поле строкового типа преобразуется в числовой тип. " + _
                                      "Все нечисловые значения будут преобразованы в 0. " + _
                                      "Продолжить?") <> resOk) Then Exit Sub
                
            End If
            For i% = 0 To (DB(QRDBIndex).Header.RowCount - 1)
                Select Case DB(QRDBIndex).Cols(colindex).Class
                    Case ccInteger
                        DB(QRDBIndex).Rows(i).Fields(colindex) = CStr(DB(QRDBIndex).Rows(i).Fields(colindex))
                    Case ccString
                        If Not IsInteger(DB(QRDBIndex).Rows(i).Fields(colindex)) Then
                            DB(QRDBIndex).Rows(i).Fields(colindex) = 0
                        Else
                            DB(QRDBIndex).Rows(i).Fields(colindex) = CLng(DB(QRDBIndex).Rows(i).Fields(colindex))
                        End If
                End Select
            Next i
            Select Case DB(QRDBIndex).Cols(colindex).Class
                Case ccInteger
                    DB(QRDBIndex).Cols(colindex).Class = ccString
                Case ccString
                    DB(QRDBIndex).Cols(colindex).Class = ccInteger
                End Select
                
        Case sName      '   ************** ...::: Name :::... ***************
            p% = InStr(1, param, ",")
            If TestZero(p) Then Exit Sub
            colindexstr$ = Trim(Left(param, p - 1))
            If Not IsInteger(colindexstr) Then
                Call ErrorInQuery
                Exit Sub
            End If
            colindex% = CLng(colindexstr)
            param = Trim(Mid(param, p + 1))
            If (param = "") Then
                Call ErrorInQuery
                Exit Sub
            End If
            ' поиск на дубликат
            For i% = 0 To DB(QRDBIndex).Header.ColCount - 1
                If (DB(QRDBIndex).Cols(i).title = param) And (i <> colindex) Then
                    Call MsgForm.ErrorMsg("Поле с названием " + param + " уже существует!")
                    Exit Sub
                End If
            Next i
            DB(QRDBIndex).Cols(colindex).title = param
            DB(QRDBIndex).Cols(colindex).TitleLen = Len(param)
        Case Default    '   **************  !!  ***************
            Call ErrorInQuery
    End Select
End Sub

Public Sub RunQuery(DBIndex_%, query$)
    Dim s1$, p%

    s1 = Mid(query, 4)
    query = Left(query, 3)
    
    QRDBIndex = DBIndex_
    
    Select Case query
        Case sAdd
            query = Left(s1, 3)
            s1 = Mid(s1, InStr(1, s1, "("))
            If (Left(s1, 1) <> "(") Or (Right(s1, 1) <> ")") Or ((Len(s1) < 8) And (query = sCol)) Then
                Call ErrorInQuery
            Else
                Call AddRun(query, Trim(Mid(s1, 2, Len(s1) - 2)))
            End If
        Case sDel
            query = Left(s1, 3)
            s1 = Mid(s1, InStr(1, s1, "("))
            If (Left(s1, 1) <> "(") Or (Right(s1, 1) <> ")") Or (Len(s1) < 5) Then
                Call ErrorInQuery
            Else
                Call DelRun(query, Trim(Mid(s1, 2, Len(s1) - 2)))
            End If
        Case sSort
            If (Left(s1, 1) <> "(") Or (Right(s1, 1) <> ")") Or (Len(s1) < 5) Then
                Call ErrorInQuery
            Else
                Call SortRun(Trim(Mid(s1, 2, Len(s1) - 2)))
            End If
        Case sOut
            If (Left(s1, 1) <> "(") Or (Right(s1, 1) <> ")") Or (Len(s1) < 5) Then
                Call ErrorInQuery
            Else
                Call OutRun(Trim(Mid(s1, 2, Len(s1) - 2)))
            End If
        Case sSwap
            query = Left(s1, 3)
            s1 = Mid(s1, InStr(1, s1, "("))
            If (Left(s1, 1) <> "(") Or (Right(s1, 1) <> ")") Or ((Len(s1) < 5) And (query = sCol)) Then
                Call ErrorInQuery
            Else
                Call SwapRun(query, Trim(Mid(s1, 2, Len(s1) - 2)))
            End If
        Case sChange
            query = Left(s1, 4)
            s1 = Mid(s1, InStr(1, s1, "("))
            If (Left(s1, 1) <> "(") Or (Right(s1, 1) <> ")") Or (Len(s1) < 3) Then
                Call ErrorInQuery
            Else
                Call ChangeRun(query, Trim(Mid(s1, 2, Len(s1) - 2)))
            End If
    End Select
    
End Sub
