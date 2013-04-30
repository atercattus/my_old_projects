VERSION 5.00
Begin VB.Form QueryMasterForm 
   Appearance      =   0  'Плоска
   BackColor       =   &H00DCDCDC&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "DBX"
   ClientHeight    =   6600
   ClientLeft      =   4710
   ClientTop       =   2850
   ClientWidth     =   8055
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
   ScaleHeight     =   6600
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox QueryList 
      Appearance      =   0  'Плоска
      Height          =   2760
      IntegralHeight  =   0   'False
      ItemData        =   "QueryMasterForm.frx":0000
      Left            =   1980
      List            =   "QueryMasterForm.frx":0002
      TabIndex        =   15
      Top             =   2580
      Width           =   5472
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Плоска
      Height          =   336
      Left            =   1980
      TabIndex        =   14
      Top             =   5520
      Width           =   5472
   End
   Begin VB.ComboBox QuerySubtypeCombo 
      Appearance      =   0  'Плоска
      Height          =   360
      ItemData        =   "QueryMasterForm.frx":0004
      Left            =   4920
      List            =   "QueryMasterForm.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1800
      Width           =   2952
   End
   Begin VB.ComboBox QueryTypeCombo 
      Appearance      =   0  'Плоска
      Height          =   360
      ItemData        =   "QueryMasterForm.frx":0008
      Left            =   2040
      List            =   "QueryMasterForm.frx":001E
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1800
      Width           =   2832
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
      Width           =   1740
      Begin VB.Label Label7 
         Alignment       =   2  'Центровка
         Appearance      =   0  'Плоска
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Прозрачно
         Caption         =   "Ask yourself"
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
         Width           =   1608
         WordWrap        =   -1  'True
      End
      Begin VB.Label MainLabel 
         Alignment       =   2  'Центровка
         Appearance      =   0  'Плоска
         BackColor       =   &H80000005&
         BackStyle       =   0  'Прозрачно
         Caption         =   "Мастер запросов"
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
         Height          =   636
         Left            =   60
         TabIndex        =   5
         Top             =   780
         Width           =   1560
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image6 
         Height          =   1440
         Left            =   180
         Picture         =   "QueryMasterForm.frx":0059
         Top             =   1860
         Width           =   1440
      End
   End
   Begin VB.Image ChangeImage 
      Height          =   480
      Left            =   7560
      Picture         =   "QueryMasterForm.frx":7123
      Top             =   5460
      Width           =   480
   End
   Begin VB.Image ClearImage 
      Height          =   480
      Left            =   7560
      Picture         =   "QueryMasterForm.frx":7965
      Top             =   4140
      Width           =   480
   End
   Begin VB.Image DelImage 
      Height          =   480
      Left            =   7560
      Picture         =   "QueryMasterForm.frx":81A7
      Top             =   3480
      Width           =   480
   End
   Begin VB.Image AddImage 
      Height          =   480
      Left            =   7560
      Picture         =   "QueryMasterForm.frx":89E9
      Top             =   2820
      Width           =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   8040
      X2              =   1740
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Центровка
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Список подготовленных запросов"
      ForeColor       =   &H00400000&
      Height          =   264
      Left            =   1860
      TabIndex        =   13
      Top             =   2280
      Width           =   6132
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Центровка
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Подтип запроса"
      ForeColor       =   &H00400000&
      Height          =   264
      Left            =   4920
      TabIndex        =   11
      Top             =   1500
      Width           =   2964
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Центровка
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Тип запроса"
      ForeColor       =   &H00400000&
      Height          =   264
      Left            =   2040
      TabIndex        =   9
      Top             =   1500
      Width           =   2820
   End
   Begin VB.Label CaptionLabel 
      Alignment       =   2  'Центровка
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Работа с БД"
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
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Мастер помогает формировать запросы к БД на основе диалогов. Однако можно обойтись и без его помощи."
      ForeColor       =   &H00400000&
      Height          =   612
      Left            =   1800
      TabIndex        =   6
      Top             =   720
      Width           =   6192
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
   Begin VB.Label RunBut 
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      ForeColor       =   &H80000008&
      Height          =   492
      Left            =   1860
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
      Caption         =   "Выполнить"
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
      Left            =   1920
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
      Caption         =   "Назад"
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
   Begin VB.Image RunImg 
      Height          =   480
      Left            =   1860
      MousePointer    =   10  'Стрелка вверх
      Stretch         =   -1  'True
      Top             =   6060
      Width           =   1572
   End
   Begin VB.Image Image3 
      Height          =   660
      Left            =   1800
      Picture         =   "QueryMasterForm.frx":962B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6432
   End
   Begin VB.Image Image1 
      Height          =   660
      Left            =   0
      Picture         =   "QueryMasterForm.frx":9749
      Stretch         =   -1  'True
      Top             =   0
      Width           =   192
   End
   Begin VB.Image TopImg 
      Height          =   660
      Left            =   180
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1620
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   1740
      X2              =   1740
      Y1              =   660
      Y2              =   6600
   End
End
Attribute VB_Name = "QueryMasterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public QMFDBIndex%

Sub AddStr(str$)
    If (str <> "") Then
        QueryList.AddItem str
    Else
        Call MsgForm.ErrorMsg("Запрос отменен!")
    End If
End Sub

Private Sub AddImage_Click()
Call SoundClick
With QueryList
    Select Case QueryTypeCombo.ListIndex
        '******************* Добавление ***********************
        Case 0
            Select Case QuerySubtypeCombo.ListIndex
                Case 0  '   добавление столбца
                    Call AddStr(Generate_Add(sCol))
                Case 1  '   добавление записи
                     Call AddStr(Generate_Add(sRow))
            End Select
        '*******************  Удаление  ***********************
        Case 1
            Select Case QuerySubtypeCombo.ListIndex
                Case 0  '   удаление столбца
                    Call AddStr(Generate_Del(sCol))
                Case 1  '   удаление записи
                    Call AddStr(Generate_Del(sRow))
            End Select
            
        '******************* Сортировка ***********************
        Case 2
            Select Case QuerySubtypeCombo.ListIndex
                Case 0  '   сортировка по алфавиту
                    Call AddStr(Generate_Sort(sAZ))
                Case 1  '   сортировка против алфавита
                    Call AddStr(Generate_Sort(sZA))
            End Select
            
        '*******************    Вывод   ***********************
        Case 3
            Select Case QuerySubtypeCombo.ListIndex
                Case 0  '   вывод на равенство записи
                    Call AddStr(Generate_Out(sEqual))
                Case 1  '   вывод больше записи
                    Call AddStr(Generate_Out(sAbove))
                Case 2  '   вывод меньше записи
                    Call AddStr(Generate_Out(sBelow))
                Case 3  '   вывод на равенство кол-ву
                    Call AddStr(Generate_Out(sCountEqual))
                Case 4  '   вывод больше кол-ва
                    Call AddStr(Generate_Out(sCountAbove))
                Case 5  '   вывод меньше кол-ва
                    Call AddStr(Generate_Out(sCountBelow))
            End Select
            
        '*******************    Обмен   ***********************
        Case 4
            Select Case QuerySubtypeCombo.ListIndex
                Case 0  '   обмен столбцов
                    Call AddStr(Generate_Swap(sCol))
                Case 1  '   обмен строк
                    Call AddStr(Generate_Swap(sRow))
            End Select
            
        '*******************    Смена   ***********************
        Case 5
            Select Case QuerySubtypeCombo.ListIndex
                Case 0  '   смена типа поля
                    Call AddStr(Generate_Change(sType))
                Case 1  '   смена названия поля
                    Call AddStr(Generate_Change(sName))
            End Select
    End Select
    
End With
End Sub

Private Sub CancelBut_Click()
    Call SoundClick
    If (QueryList.ListCount > 0) Then
        If (MsgForm.QuestMsg("Список запросов не пуст. Выйти?") = resOk) Then Unload Me
    Else
        Unload Me
    End If
End Sub

' замена запроса
Private Sub ChangeImage_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (Trim(Text1) <> "") Then
        Call SoundClick
        With QueryList
            If (.ListIndex = -1) Or (Shift And vbShiftMask <> 0) Then
                .AddItem Text1
            Else
                .List(.ListIndex) = Text1
            End If
        End With
    End If
    Text1 = ""
    Text1.SetFocus
End Sub

' очистка запросов
Private Sub ClearImage_Click()
    If (QueryList.ListCount > 0) Then
        Call SoundClick
        If (MsgForm.QuestMsg("Очистить список запросов?") = resOk) Then
            QueryList.Clear
            Text1 = ""
            Text1.SetFocus
        End If
    End If
End Sub

' удаление запроса
Private Sub DelImage_Click()
    If (QueryList.ListIndex >= 0) Then
        Call SoundClick
        If (MsgForm.QuestMsg("Удалить выбранный запрос из списка?") = resOk) Then
            QueryList.RemoveItem QueryList.ListIndex
            Text1 = ""
            Text1.SetFocus
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
    Case 27
        CancelBut_Click
    Case 13
        RunBut_Click
 End Select
End Sub

Private Sub Form_Load()
    QueryTypeCombo.ListIndex = 0
    Call ButEnabled(RunImg, RunBut, True)
    Call ButEnabled(CancelImg, CancelBut, True)
    TopImg.Picture = MainForm.TopImageList.ListImages(1).Picture
End Sub

Private Sub QueryList_DblClick()
 With QueryList
  If (.ListIndex <> -1) Then
    Text1 = .List(.ListIndex)
    Text1.SetFocus
  End If
 End With
End Sub

Private Sub QueryTypeCombo_Click()
 With QuerySubtypeCombo
  .Clear
  Select Case QueryTypeCombo.ListIndex
   Case 0
    .AddItem "Поля"
    .AddItem "Записи"
   Case 1
    .AddItem "Поля"
    .AddItem "Записи"
   Case 2
    .AddItem "По алфавиту"
    .AddItem "Против алфавита"
   Case 3
    .AddItem "Равно записи"
    .AddItem "Больше записи"
    .AddItem "Меньше записи"
    .AddItem "Равно кол-ву копий"
    .AddItem "Больше кол-ва копий"
    .AddItem "Меньше кол-ва копий"
   Case 4
    .AddItem "Полей"
    .AddItem "Записей"
   Case 5
    .AddItem "Типа поля"
    .AddItem "Названия поля"
  End Select
  .ListIndex = 0
 End With
End Sub

Private Sub RunBut_Click()
    If (QueryList.ListCount > 0) Then
        Call SoundClick
        For i% = 0 To QueryList.ListCount - 1
            Call RunQuery(QMFDBIndex, QueryList.List(i))
        Next i
        With MainForm
            .TabStrip.SelectedItem = .TabStrip.Tabs(QMFDBIndex + 1)
            Call ShowTable(QMFDBIndex)
        End With
        QueryList.Clear
        Call MsgForm.InfoMsg("Запросы выполнены.")
    End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 13) Then Call ChangeImage_MouseDown(vbLeftButton, Shift, 1, 1)
End Sub
