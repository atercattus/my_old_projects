Attribute VB_Name = "DBTypes"
'************************************
'          модуль DBTypes.bas
'        вся работа с файлом БД
'************************************

'************************************** Описание типов **************************************

' заголовок файла
Type TDBHeader
 ' "DBX" - проверка файла
 Header As String * 3
 ' флаги
 Flags As Byte
 ' количество полей
 ColCount As Long
 ' количество записей
 RowCount As Long
End Type

' имеет ли пользователь права на редактирование
Public UserIsAdmin As Boolean

' данные о столбце
Type TDBElemData
 ' тип данных
 Class As Byte
 ' длина заголовка
 TitleLen As Byte
 ' заголовок, длины TitleLen
 title As String
 ' значение по-умолчанию
 DefValue As Variant
End Type

' запись
Type TDBElem
 ' поля записи
 Fields() As Variant
End Type

' элемент в массиве DB
Type TDBCell
    Header As TDBHeader
    Cols() As TDBElemData
    Rows() As TDBElem
    Password As String
End Type

'************************************** Описание констант **************************************

' контрольный байт
Public Const ValidateByte As Byte = &H7F

'************************************** Описание переменных **************************************

' путь к БД
Public DBPath$
' флаг изменения БД
Public DBChanged As Boolean
' данные таблиц: каждый элемент - это копия некоторой таблицы
Public DB() As TDBCell

'************************************** Процедуры и функции **************************************

' удаление поля
Public Sub DelCol_(DBIndex%, Optional ByVal Index% = -1, Optional ByVal conf As Boolean = True)
 With DB(DBIndex).Header
  If (.ColCount = 0) Then Exit Sub
  If (Index = -1) Then Index = .ColCount - 1
  If (Index > .ColCount - 1) Or (Index < -1) Then
   Call MsgForm.ErrorMsg("Ошибка удаления столбца!")
   Exit Sub
  End If
 
  If conf Then
    If (MsgForm.QuestMsg("Удалить столбец?") <> resOk) Then Exit Sub
  End If
  ' вырезаю из полей
  For i% = Index To (.ColCount - 2)
    DB(DBIndex).Cols(i) = DB(DBIndex).Cols(i + 1)
  Next i
  ' вырезаю из записей
  For R% = 0 To (.RowCount - 1)
    For c% = Index To (.ColCount - 2)
        DB(DBIndex).Rows(R).Fields(c) = DB(DBIndex).Rows(R).Fields(c + 1)
    Next c
  Next R
  
  .ColCount = .ColCount - 1
  ReDim Preserve DB(DBIndex).Cols(.ColCount)
  DBChanged = True
End With
End Sub

' удаление записи
Public Sub DelRow_(DBIndex%, Optional ByVal Index% = -1, Optional ByVal conf As Boolean = True)
 With DB(DBIndex).Header
  If (.RowCount = 0) Then Exit Sub
  If (Index = -1) Then Index = .RowCount - 1
  If (Index > .RowCount - 1) Then
   Call MsgForm.ErrorMsg("Ошибка удаления записи!")
   Exit Sub
  End If
 
  If conf Then
    If (MsgForm.QuestMsg("Удалить запись?") = resNo) Then Exit Sub
  End If
  For i% = Index To (.RowCount - 2)
    DB(DBIndex).Rows(i) = DB(DBIndex).Rows(i + 1)
  Next i
  .RowCount = .RowCount - 1
  ReDim Preserve DB(DBIndex).Rows(.RowCount)
  DBChanged = True
End With
End Sub

Public Sub TestDBChanged()
    If DBChanged Then
        MainForm.SB.Panels(1).Picture = MainForm.ImageList1.ListImages(2).Picture
    Else
        Set MainForm.SB.Panels(1).Picture = Nothing
    End If
End Sub

' отображение таблицы
Public Sub ShowTable(DBIndex%)
    MainForm.ListView.ListItems.Clear
    MainForm.ListView.ColumnHeaders.Clear
    If (DBIndex = -1) Then
        DBPath = ""
        MainForm.SB.Panels(3).Text = ""
        GoTo exit_
    End If
    If (DB(DBIndex).Header.ColCount = 0) Then GoTo exit_
    For c% = 0 To DB(DBIndex).Header.ColCount - 1
        Call MainForm.ListView.ColumnHeaders.Add( _
                                                    MainForm.ListView.ColumnHeaders.Count + 1, _
                                                    "col_key_" + CStr(c), _
                                                    DB(DBIndex).Cols(c).title, _
                                                    1440, _
                                                    lvwColumnLeft, _
                                                    0 _
                                                )

    Next c
    For R% = 0 To DB(DBIndex).Header.RowCount - 1
        With MainForm.ListView.ListItems.Add
            .Key = "row_key_" + CStr(R)
            .Text = DB(DBIndex).Rows(R).Fields(0)
            For i% = 1 To DB(DBIndex).Header.ColCount - 1
                .SubItems(i) = DB(DBIndex).Rows(R).Fields(i)
            Next i
        End With
    Next R
exit_:
    MainForm.TabStrip.Visible = (DBPath <> "")
    MainForm.ListView.Visible = MainForm.TabStrip.Visible
    If (DBIndex <> -1) Then
        MainForm.SB.Panels(2).Text = CStr(DB(DBIndex).Header.RowCount)
    Else
        MainForm.SB.Panels(2).Text = ""
    End If
    Call TestDBChanged
End Sub

' поиск поля *************************************************
Public Function ItColAlreadyCreate(QRDBIndex%, title$) As Boolean
    With DB(QRDBIndex)
        For i% = 0 To (DB(QRDBIndex).Header.ColCount - 1)
            If (.Cols(i).title = title) Then
                ItColAlreadyCreate = True
                Exit Function
            End If
        Next i
    End With
    ItColAlreadyCreate = False
End Function

' добавление поля *************************************************
Public Sub AddCol(DBIndex%, ByVal Class%, ByVal title$, ByVal defval, Optional ByVal pos% = -1)
 With DB(DBIndex).Header
  ReDim Preserve DB(DBIndex).Cols(.ColCount)
  If (pos = -1) Then
   pos = .ColCount
  Else
   For i% = 1 To (.ColCount - pos)
    DB(DBIndex).Cols(.ColCount - i + 1) = DB(DBIndex).Cols(.ColCount - i)
   Next i
  End If
  With DB(DBIndex).Cols(pos)
   .Class = Class
   .title = title
   .TitleLen = Len(title)
   .DefValue = defval
  End With
 
  ' увеличиваю размерность записей
  For R% = 0 To DB(DBIndex).Header.RowCount - 1
   ReDim Preserve DB(DBIndex).Rows(R).Fields(.ColCount)
   For i% = 1 To (.ColCount - pos)
    DB(DBIndex).Rows(R).Fields(.ColCount - i + 1) = DB(DBIndex).Rows(R).Fields(.ColCount - i)
   Next i
   DB(DBIndex).Rows(R).Fields(pos) = DB(DBIndex).Cols(pos).DefValue
  Next R
 
  .ColCount = .ColCount + 1
 
  DBChanged = True
 End With
End Sub

' добавление записи *************************************************
Public Sub AddField(DBIndex%, row)
 With DB(DBIndex).Header
  ReDim Preserve DB(DBIndex).Rows(.RowCount)
  DB(DBIndex).Rows(.RowCount).Fields = row
  .RowCount = .RowCount + 1
  DBChanged = True
 End With
End Sub

' удаление таблицы *************************************************
Public Sub DelTable(Index%)
    For i% = Index To (UBound(DB) - 1)
        DB(i) = DB(i + 1)
    Next i
    If (UBound(DB) > 0) Then ReDim Preserve DB(UBound(DB) - 1)
End Sub

' если нужно то строка шифруется по паролю, иначе не изменяется
Function CodeDecode(Index%, str$, col%, row%, Optional pass$ = "", Optional usepass As Boolean = False) As String
        If Not usepass Then pass$ = DB(Index).Password
        If (pass = "") Then
            CodeDecode = str
            Exit Function
        End If
        CodeDecode = ""
        p% = 1
        Dim ch As Byte
        For i% = 1 To Len(str)
            ch = Asc(Mid(str, i, 1)) Xor Asc(Mid(pass, p, 1)) Xor col Xor row
            CodeDecode = CodeDecode + Chr(ch)
            p = p + 1: If p > Len(pass) Then p = 1
        Next i
End Function

' сохранение БД в файле *************************************************
Public Sub FlushDB(DBIndex%)
 Dim s$, W%
 If Not UserIsAdmin Then
    Call ProtectedMsg
    Exit Sub
 End If
 If (DBPath <> "") Then
  Call DeleteFile(DBPath)
  DBI% = FreeFile
  Open DBPath For Binary As DBI
  
  ' заголовок - 12
  Put DBI, , DB(DBIndex).Header
  
  ' если надо, то сохраняю пароль
  If (DB(DBIndex).Header.Flags And flPasswordNeed) Then
   Dim str$, ch1 As Byte, ch2 As Byte
   Dim lng As Byte, lng2 As Byte
   lng = Len(DB(DBIndex).Password)
   lng2 = lng / 2
   Put DBI, , lng
   
   For i% = 1 To lng2
    ch1 = Asc(Mid(DB(DBIndex).Password, i, 1))
    ch2 = Asc(Mid(DB(DBIndex).Password, lng - i + 1, 1))
    str = Chr(ch1 Xor ch2) + str
   Next i
   For i = lng2 To 1 Step -1
    Put DBI, , CByte(Asc(Mid(str, i, 1)))
   Next i
  End If ' сохранение пароля
   
  ' данные полей
  Dim l As Long
  For i% = 0 To DB(DBIndex).Header.ColCount - 1
   Put DBI, , DB(DBIndex).Cols(i).Class
   Put DBI, , DB(DBIndex).Cols(i).TitleLen
   If (DB(Index).Header.Flags And flCoded) Then
    Put DBI, , CodeDecode(DBIndex, DB(DBIndex).Cols(i).title, i, 0)
   Else
    Put DBI, , DB(DBIndex).Cols(i).title
   End If
   Select Case DB(DBIndex).Cols(i).Class
    Case ccString
     If (DB(Index).Header.Flags And flCoded) Then
      s = CodeDecode(DBIndex, CStr(DB(DBIndex).Cols(i).DefValue), i, 0)
     Else
      s = CStr(DB(DBIndex).Cols(i).DefValue)
     End If
     W = Len(s)
     Put DBI, , W
     Put DBI, , s
     Case ccInteger
     If (Not IsInteger(DB(DBIndex).Cols(i).DefValue)) Then
        l = 0
     Else
        l = CLng(DB(DBIndex).Cols(i).DefValue)
     End If
     Put DBI, , l
   End Select
  Next i
  
  ' запись контрольного байта
  Put DBI, , ValidateByte
  
  ' записи
  Dim f As TDBElem
  Dim col As TDBElemData
  For R% = 0 To DB(DBIndex).Header.RowCount - 1
   f = DB(DBIndex).Rows(R)
   For c% = 0 To DB(DBIndex).Header.ColCount - 1
    col = DB(DBIndex).Cols(c)
    ' в зависимости от типа данных колонки пишу в файл определённый тип данных
    Select Case col.Class
     ' если число - записываю как long
     Case ccInteger
      l = CLng(f.Fields(c))
      Put DBI, , l
     ' если строка - то байт длины и сама строка
     Case ccString
      If (DB(Index).Header.Flags And flCoded) Then
       s = CodeDecode(DBIndex, CStr(f.Fields(c)), c, R)
      Else
       s = CStr(f.Fields(c))
      End If
      ' Len возвращает 4 байта, а мне нужно 2
      W = Len(s)
      Put DBI, , W
      Put DBI, , s
    End Select
   Next c
  Next R
  
  MainForm.SB.Panels(3).Text = DBPath
  Call MsgForm.InfoMsg("БД сохранена!")
  
  ' закрытие файла
  Close
  DBChanged = False
  Call TestDBChanged
 End If
End Sub

' загрузка БД *************************************************
Public Function LoadDB(DBIndex%, ByVal Path$) As Boolean
 Dim DBH As TDBHeader
 pwrd$ = ""
 LoadDB = False
 DBI% = FreeFile
 DBP$ = Path
 ' открываю БД
 Open DBP For Binary As DBI
 ' считываю заголовок
 Get DBI, , DBH
 With DBH
  If (.Header <> "DBX") Then
   Call MsgForm.ErrorMsg("БД повреждена!")
   Close
   Exit Function
  End If

  ' если надо, то загружаю пароль
  If (DBH.Flags And flPasswordNeed) Then
   Dim lng As Byte
   Get DBI, , lng
   Dim str$, ch1 As Byte, ch2 As Byte, ch3 As Byte
   str = ""
   For i% = 1 To lng \ 2
    Get DBI, , ch1
    str = str + Chr(ch1)
   Next i
'********************************************************
   With PasswordForm
       .PassText = ""
        
       .CaptionLabel = "Защита БД"
       .TextLabel = "Открываемая БД защищена паролем. Для работы с БД необходимо ввести пароль."
       .Frame2.Visible = False
       .Frame1.Visible = True
       
       Dim ROE As Boolean
       
       ROE = Not ((DBH.Flags And flReadOnlyEnable) = flReadOnlyEnable)
       
       .Frame3.Visible = ROE
       .NoFullLabel.Visible = Not ROE
       .Show vbModal
       If (.res) Then
            ' допустимый тип доступа
            Mode% = 0
            ' введёный пароль
            str2$ = Trim(.PassText)
            
            ' проверка пароля
                lng_2 = Len(str2)
                If (lng_2 <> lng) Then
                    Mode = -1
                    GoTo bad
                End If
                For i% = 1 To lng \ 2
                    ch1 = Asc(Mid(str2, i, 1))
                    ch2 = Asc(Mid(str2, lng - i + 1, 1))
                    ch3 = Asc(Mid(str, i, 1))
                    If ((ch1 Xor ch2) <> ch3) Then
                        Mode = -1
                        GoTo bad
                    End If
                Next i
            
bad:
            ' обработка правильности пароля и уровня доступа
            If (Mode = 0) And (.Check1 = 0) Then
                Call MsgForm.InfoMsg("Пароль принят!")
                pwrd = str2
                UserIsAdmin = True
            Else
                If ROE And (.Check1 = 1) Then
                    Call MsgForm.InfoMsg("Только чтение!")
                    UserIsAdmin = False
                Else
                    Call MsgForm.ErrorMsg("Пароль не принят! Доступ запрещён!")
                    Unload PasswordForm
                    Close
                    Exit Function
                End If
            End If
       Else
            Unload PasswordForm
            Close
            Exit Function
       End If   ' if (.res)
       Unload PasswordForm
   End With
'********************************************************
  End If

  ' выделение нужной памяти
    If (.ColCount > 0) Then
        ReDim DB(DBIndex).Cols(.ColCount - 1)
        If (.RowCount > 0) Then
            ReDim DB(DBIndex).Rows(.RowCount - 1)
            For R% = 0 To .RowCount - 1
                ReDim DB(DBIndex).Rows(R).Fields(.ColCount - 1)
            Next R
        End If
    End If
  
  ' считывание данных полей
  For i% = 0 To DBH.ColCount - 1
   ' получение класса
   Get DBI, , DB(DBIndex).Cols(i).Class
   ' получение длины заголовка
   Get DBI, , DB(DBIndex).Cols(i).TitleLen
   ' получение заголовка
   s$ = ""
   Dim B As Byte
   For j% = 1 To DB(DBIndex).Cols(i).TitleLen
    Get DBI, , B
    s = s + Chr(B)
   Next j
   s = CodeDecode(DBIndex, s, i, 0, pwrd, True)
   DB(DBIndex).Cols(i).title = s
   ' получение значения по-умолчанию
   Dim l As Long
   Dim W%
   Select Case DB(DBIndex).Cols(i).Class
    Case ccInteger
     Get DBI, , l
     DB(DBIndex).Cols(i).DefValue = l
    Case ccString
     Get DBI, , W
     s = ""
     For j% = 1 To W
      Get DBI, , B
      s = s + Chr(B)
     Next j
     s = CodeDecode(DBIndex, s, i, 0, pwrd, True)
     DB(DBIndex).Cols(i).DefValue = s
   End Select
  Next i
  
  ' чтение контрольного байта
  Dim VB As Byte
  Get DBI, , VB
  If (VB <> ValidateByte) Then
   Call MsgForm.ErrorMsg("БД повреждена!")
   Close
   Exit Function
  End If

  ' считывание записей
  Dim col As TDBElemData
  For R% = 0 To .RowCount - 1
   For c% = 0 To .ColCount - 1
    col = DB(DBIndex).Cols(c)
   ' в зависимости от типа данных колонки пишу в файл определённый тип данных
   Select Case col.Class
    ' если число - считываю как long
    Case ccInteger
     Get DBI, , l
     DB(DBIndex).Rows(R).Fields(c) = l
    ' если строка - то байт длины и сама строка
    Case ccString
     Get DBI, , W
     s = ""
     For j% = 1 To W
      Get DBI, , B
      s = s + Chr(B)
     Next j
     s = CodeDecode(DBIndex, s, c, R, pwrd, True)
     DB(DBIndex).Rows(R).Fields(c) = s
   End Select
   Next c
  Next R
  
 End With
 LoadDB = True
 
 DB(DBIndex).Header = DBH
 DBPath = DBP
 DBChanged = False
 DB(DBIndex).Password = pwrd
 
 Call MsgForm.InfoMsg("БД загружена!")
 
 ' закрытие файла
 Close
End Function

' создание новой БД *************************************************
Public Function NewDB(Path$)
 DBI% = FreeFile
 ' удаляю БД
 Call DeleteFile(Path)
 ' открываю БД
 Open Path For Binary As DBI
 ' применяю стандартный заголовок к БД
 Call ClearAll
 DBPath = Path
 ' записываю заголовок БД
 Put DBI, , DB(0).Header
 ' запись контрольного байта
 Put DBI, , ValidateByte
 Close
 Call MsgForm.InfoMsg("БД создана с настройками по-умолчанию!")
End Function

' очистка ВСЕГО
Public Sub ClearAll()
 ReDim DB(0)
 Call ClearHeader(DB(0).Header)
 DBChanged = False
 DBPath = ""
End Sub

' установка полей в начальные значения *************************************************
Public Sub ClearHeader(H As TDBHeader)
 H.Header = "DBX"
 H.Flags = 0
 H.ColCount = 0
 H.RowCount = 0
End Sub
