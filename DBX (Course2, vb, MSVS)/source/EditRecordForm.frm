VERSION 5.00
Begin VB.Form EditRecordForm 
   Appearance      =   0  'Плоска
   BackColor       =   &H00DCDCDC&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "DBX Редактирование"
   ClientHeight    =   6615
   ClientLeft      =   4920
   ClientTop       =   3105
   ClientWidth     =   8040
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Плоска
      BackColor       =   &H00DCDCDC&
      Caption         =   "Изменение значения поля"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1392
      Left            =   2400
      TabIndex        =   13
      Top             =   4500
      Width           =   5472
      Begin VB.TextBox Text1 
         Appearance      =   0  'Плоска
         BackColor       =   &H00FFFFFF&
         Height          =   336
         Left            =   120
         TabIndex        =   14
         Top             =   300
         Width           =   5232
      End
      Begin VB.Label FlipBut 
         Appearance      =   0  'Плоска
         BackColor       =   &H80000005&
         BackStyle       =   0  'Прозрачно
         ForeColor       =   &H80000008&
         Height          =   492
         Left            =   3720
         MousePointer    =   1  'Указатель
         TabIndex        =   18
         Top             =   780
         Width           =   1572
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Центровка
         Appearance      =   0  'Плоска
         BackColor       =   &H80000005&
         BackStyle       =   0  'Прозрачно
         Caption         =   "Применить"
         ForeColor       =   &H80000008&
         Height          =   348
         Left            =   3780
         MousePointer    =   1  'Указатель
         TabIndex        =   17
         Top             =   876
         Width           =   1476
      End
      Begin VB.Image FlipImg 
         Height          =   480
         Left            =   3720
         MousePointer    =   10  'Стрелка вверх
         Stretch         =   -1  'True
         Top             =   780
         Width           =   1572
      End
      Begin VB.Label EditorBut 
         Appearance      =   0  'Плоска
         BackColor       =   &H80000005&
         BackStyle       =   0  'Прозрачно
         ForeColor       =   &H80000008&
         Height          =   492
         Left            =   180
         MousePointer    =   1  'Указатель
         TabIndex        =   15
         Top             =   780
         Width           =   1572
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Центровка
         Appearance      =   0  'Плоска
         BackColor       =   &H80000005&
         BackStyle       =   0  'Прозрачно
         Caption         =   "Редактор"
         ForeColor       =   &H80000008&
         Height          =   348
         Left            =   240
         MousePointer    =   1  'Указатель
         TabIndex        =   16
         Top             =   876
         Width           =   1476
      End
      Begin VB.Image EditorImg 
         Appearance      =   0  'Плоска
         Height          =   480
         Left            =   180
         MousePointer    =   10  'Стрелка вверх
         Stretch         =   -1  'True
         Top             =   780
         Width           =   1572
      End
   End
   Begin VB.ListBox CellList 
      Appearance      =   0  'Плоска
      BackColor       =   &H00F4F4F4&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1824
      IntegralHeight  =   0   'False
      ItemData        =   "EditRecordForm.frx":0000
      Left            =   2400
      List            =   "EditRecordForm.frx":0002
      TabIndex        =   10
      Top             =   2040
      Width           =   5472
   End
   Begin VB.Frame Frame0 
      Appearance      =   0  'Плоска
      BackColor       =   &H00CCCCCC&
      BorderStyle     =   0  'Нет
      ForeColor       =   &H80000008&
      Height          =   5952
      Left            =   0
      TabIndex        =   0
      Top             =   660
      Width           =   2214
      Begin VB.Label Label7 
         Alignment       =   2  'Центровка
         Appearance      =   0  'Плоска
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Прозрачно
         Caption         =   "Change yourself"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   228
         Left            =   60
         TabIndex        =   8
         Top             =   5640
         Width           =   2088
         WordWrap        =   -1  'True
      End
      Begin VB.Label MainLabel 
         Alignment       =   2  'Центровка
         Appearance      =   0  'Плоска
         BackColor       =   &H80000005&
         BackStyle       =   0  'Прозрачно
         Caption         =   "Редактор полей записей"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   204
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   876
         Left            =   120
         TabIndex        =   5
         Top             =   900
         Width           =   1980
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image6 
         Appearance      =   0  'Плоска
         Height          =   1440
         Left            =   420
         Picture         =   "EditRecordForm.frx":0004
         Top             =   1920
         Width           =   1440
      End
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Плоска
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Это X поле"
      ForeColor       =   &H80000008&
      Height          =   264
      Left            =   2580
      TabIndex        =   19
      Top             =   4020
      Width           =   3612
      WordWrap        =   -1  'True
   End
   Begin VB.Label ReturnBut 
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      ForeColor       =   &H80000008&
      Height          =   492
      Left            =   6240
      MousePointer    =   1  'Указатель
      TabIndex        =   11
      Top             =   3960
      Width           =   1572
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Плоска
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Список всех полей выбранной записи"
      ForeColor       =   &H00400000&
      Height          =   264
      Left            =   2400
      TabIndex        =   9
      Top             =   1740
      Width           =   4692
      WordWrap        =   -1  'True
   End
   Begin VB.Label CaptionLabel 
      Alignment       =   2  'Центровка
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Редактирование записей"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   13.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   432
      Left            =   2280
      TabIndex        =   7
      Top             =   120
      Width           =   5712
   End
   Begin VB.Label TextLabel 
      Appearance      =   0  'Плоска
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      Caption         =   $"EditRecordForm.frx":70CE
      ForeColor       =   &H00400000&
      Height          =   792
      Left            =   2280
      TabIndex        =   6
      Top             =   720
      Width           =   5712
      WordWrap        =   -1  'True
   End
   Begin VB.Label CancelBut 
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      ForeColor       =   &H80000008&
      Height          =   492
      Left            =   6420
      MousePointer    =   1  'Указатель
      TabIndex        =   4
      Top             =   6060
      Width           =   1572
   End
   Begin VB.Label SelectBut 
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      ForeColor       =   &H80000008&
      Height          =   492
      Left            =   2340
      MousePointer    =   1  'Указатель
      TabIndex        =   3
      Top             =   6060
      Width           =   1572
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Центровка
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Принять"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   13.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   348
      Left            =   2400
      MousePointer    =   1  'Указатель
      TabIndex        =   2
      Top             =   6144
      Width           =   1476
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Центровка
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Отмена"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   13.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   348
      Left            =   6480
      MousePointer    =   1  'Указатель
      TabIndex        =   1
      Top             =   6144
      Width           =   1476
   End
   Begin VB.Image CancelImg 
      Height          =   480
      Left            =   6420
      MousePointer    =   10  'Стрелка вверх
      Stretch         =   -1  'True
      Top             =   6060
      Width           =   1572
   End
   Begin VB.Image SelectImg 
      Height          =   480
      Left            =   2340
      MousePointer    =   10  'Стрелка вверх
      Stretch         =   -1  'True
      Top             =   6060
      Width           =   1572
   End
   Begin VB.Image Image3 
      Height          =   660
      Left            =   2280
      Picture         =   "EditRecordForm.frx":7156
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5772
   End
   Begin VB.Image Image1 
      Height          =   660
      Left            =   0
      Picture         =   "EditRecordForm.frx":7274
      Stretch         =   -1  'True
      Top             =   0
      Width           =   672
   End
   Begin VB.Image TopImg 
      Height          =   660
      Left            =   660
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1620
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   2220
      X2              =   2220
      Y1              =   660
      Y2              =   6600
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Центровка
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Вернуть"
      ForeColor       =   &H80000008&
      Height          =   348
      Left            =   6300
      MousePointer    =   1  'Указатель
      TabIndex        =   12
      Top             =   4044
      Width           =   1476
   End
   Begin VB.Image ReturnImg 
      Height          =   480
      Left            =   6240
      MousePointer    =   10  'Стрелка вверх
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   1572
   End
End
Attribute VB_Name = "EditRecordForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ERFDBIndex%
Dim RowIndexSave%
Dim protect As Boolean
Dim Arr()

Public Sub LoadData(RowIndex%)
 RowIndexSave = RowIndex
 With DB(ERFDBIndex).Header
    ReDim Arr(.ColCount, 1)
    For i% = 0 To .ColCount - 1
        Arr(i, 0) = DB(ERFDBIndex).Rows(RowIndex).Fields(i)
        Arr(i, 1) = DB(ERFDBIndex).Cols(i).Class
    Next i
 End With
End Sub

Private Sub CellList_Click()
 i% = CellList.ListIndex
 Select Case Arr(i, 1)
    Case ccInteger
        Label6.Caption = "Поле числового типа"
        Call ButEnabled(EditorImg, EditorBut, False)
    Case ccString
        Label6.Caption = "Поле строкового типа"
        Call ButEnabled(EditorImg, EditorBut, True)
 End Select
 With Text1
    .Text = CStr(Arr(i, 0))
    .SelStart = 0
    .SelLength = Len(.Text)
 End With
End Sub

Public Sub OverloadList()
 CellList.Clear
 For i% = 0 To DB(ERFDBIndex).Header.ColCount - 1
    CellList.AddItem CStr(Arr(i, 0))
 Next i
 CellList.ListIndex = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If (KeyCode = 27) Then CancelBut_Click
End Sub

Private Sub Form_Load()
 protect = False
 Call ButEnabled(ReturnImg, ReturnBut, True)
 Call ButEnabled(EditorImg, EditorBut, False)
 Call ButEnabled(FlipImg, FlipBut, True)
 Call ButEnabled(SelectImg, SelectBut, True)
 Call ButEnabled(CancelImg, CancelBut, True)
 TopImg.Picture = MainForm.TopImageList.ListImages(1).Picture
 
' If (Not protect) Then
'    Call OverloadList
' Else
'    protect = False
' End If
 
End Sub

Private Sub ReturnBut_Click()
 Call SoundClick
 If (MsgForm.QuestMsg("Восстановить поля из БД?") = resOk) Then
    Call LoadData(RowIndexSave)
    Call OverloadList
    Call MsgForm.InfoMsg("Поля были восстановлены!")
 End If
End Sub

Private Sub EditorBut_Click()
    If (EditorBut.Tag = 0) Then Exit Sub
    Call SoundClick
    i% = CellList.ListIndex
    If (Arr(i, 1) = ccInteger) Then
        Call MsgForm.InfoMsg("Для редактирования чисел редактор не исспользуется.")
        Exit Sub
    End If
    If IsDate(Text1.Text) And (MonthForm.Check1.value = 0) Then
        s$ = Text1.Text
        p% = InStr(1, s, ".")
        MonthForm.MonthView1.Day = CInt(Left(s, p - 1))
        s = Mid(s, p + 1)
        p% = InStr(1, s, ".")
        MonthForm.MonthView1.Month = CInt(Left(s, p - 1))
        s = Mid(s, p + 1)
        MonthForm.MonthView1.Year = CInt(s)

        MonthForm.Show vbModal
        Select Case MonthForm.res
            Case 1
                Text1.Text = CStr(MonthForm.MonthView1.Day) + "." + CStr(MonthForm.MonthView1.Month) + "." + CStr(MonthForm.MonthView1.Year)
            Case -1
                GoTo text_
        End Select
    Else
text_:
        With TextEditForm
            .TextEdit.Text = Text1.Text
            protect = True
            .Show vbModal
            If (.res = 1) Then Text1.Text = .TextEdit.Text
            Unload TextEditForm
        End With
    End If
End Sub

Private Sub SelectBut_Click()
Call SoundClick
If UserIsAdmin Then
 If (MsgForm.QuestMsg("Сохранить поля в БД?") = resOk) Then
    With DB(ERFDBIndex)
        Dim tmparr()
        ReDim tmparr(.Header.ColCount)
        For i% = 0 To .Header.ColCount - 1
            tmparr(i) = Arr(i, 0)
        Next i
        If (Not FindRow(ERFDBIndex, tmparr)) Then
            For i% = 0 To .Header.ColCount - 1
                .Rows(RowIndexSave).Fields(i) = Arr(i, 0)
            Next i
            DBChanged = True
            Call MsgForm.InfoMsg("Поля были сохранены в БД!")
            Call ShowTable(ERFDBIndex)
            Unload Me
        Else
            Call MsgForm.ErrorMsg("Изменённое поле перекрывает уже существующее! Измените данные.")
        End If
    End With
 End If
Else
 Call ProtectedMsg
End If
End Sub

Private Sub CancelBut_Click()
 Call SoundClick
 Unload Me
End Sub

' Посимвольное сравнение str с '2147483647' - максимальным значением Long
Function isVeryLong(str$) As Boolean
    If (Left(str, 1) = "-") Then str = Mid(str, 2)
    For i% = 1 To (10 - Len(str))
        str = "0" + str
    Next i
    
    maxval$ = "2147483647"
    For i% = 1 To 10
        ch1$ = Mid(maxval, i, 1)
        ch2$ = Mid(str, i, 1)
        If (Asc(ch2) > Asc(ch1)) Then
            isVeryLong = True
            GoTo exit_
        ElseIf (ch2 <> ch1) Then
         isVeryLong = False
         GoTo exit_
        End If
    Next i
    isVeryLong = False
exit_:
End Function

Private Sub FlipBut_Click()
Call SoundClick
If UserIsAdmin Then
 tmp = Null
 i% = CellList.ListIndex
 mln% = 10
 If (Left(Text1.Text, 1) = "-") Then mln = mln + 1
 If (Arr(i, 1) = ccInteger) Then
    If (Len(Trim(Text1.Text)) > mln) Or (isVeryLong(Trim(Text1.Text))) Then
        Call MsgForm.ErrorMsg("Числовое значение превышает разрядную сетку!")
        With Text1
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
        GoTo exit_
    End If
    
    If IsInteger(Trim(Text1.Text)) Then
        tmp = CLng(Text1.Text)
    Else
        Call MsgForm.ErrorMsg("Значение не является целым числом!")
        With Text1
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If
 Else
    If (Trim(Text1.Text) = "") Then
        If (MsgForm.QuestMsg("Строка пуста. Продолжить?") = resOk) Then
            tmp = Text1.Text
            GoTo exit_
        Else
            With Text1
                .SelStart = 0
                .SelLength = Len(.Text)
            End With
        End If
    Else
        tmp = Text1.Text
    End If
 End If
 
 ' Введёное значение прошло контроль
 If (Not IsNull(tmp)) Then
    Select Case Arr(i, 1)
        Case ccInteger: Arr(i, 0) = CLng(tmp)
        Case ccString: Arr(i, 0) = CStr(tmp)
    End Select
    curpos% = CellList.ListIndex
    Call OverloadList
    CellList.ListIndex = curpos
 End If
exit_:
Else
    Call ProtectedMsg
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 13) Then FlipBut_Click
End Sub
