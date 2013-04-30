VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form DiagMasterForm 
   Appearance      =   0  'Плоска
   BackColor       =   &H00DCDCDC&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "DBX"
   ClientHeight    =   6600
   ClientLeft      =   6420
   ClientTop       =   3390
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
   Begin MSComctlLib.ImageList DiagTypeImgs 
      Left            =   4760
      Top             =   6048
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DiagMasterForm.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DiagMasterForm.frx":0C52
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DiagMasterForm.frx":18A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DiagMasterForm.frx":24F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DiagMasterForm.frx":3148
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DiagMasterForm.frx":3665
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Плоска
      BackColor       =   &H00DCDCDC&
      Caption         =   "Тип диаграммы"
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
      Height          =   2532
      Left            =   5100
      TabIndex        =   14
      Top             =   1500
      Width           =   2832
      Begin VB.Frame Frame2 
         Appearance      =   0  'Плоска
         BackColor       =   &H00DCDCDC&
         BorderStyle     =   0  'Нет
         ForeColor       =   &H80000008&
         Height          =   1078
         Left            =   112
         TabIndex        =   16
         Top             =   1400
         Width           =   2646
         Begin VB.CheckBox Enabled3DCheck 
            Appearance      =   0  'Плоска
            BackColor       =   &H00DCDCDC&
            Caption         =   "Строить диаграмму в объёме"
            ForeColor       =   &H80000008&
            Height          =   504
            Left            =   56
            MaskColor       =   &H00DCDCDC&
            TabIndex        =   17
            Top             =   56
            Width           =   2352
         End
         Begin VB.Image DimImg 
            Appearance      =   0  'Плоска
            Height          =   375
            Left            =   2160
            Top             =   600
            Width           =   375
         End
      End
      Begin VB.ComboBox DiagTypeCombo 
         Appearance      =   0  'Плоска
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         ItemData        =   "DiagMasterForm.frx":3B87
         Left            =   60
         List            =   "DiagMasterForm.frx":3B97
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   420
         Width           =   2712
      End
      Begin VB.Image DiagTypeImage 
         Height          =   384
         Left            =   2280
         Top             =   840
         Width           =   384
      End
   End
   Begin VB.ListBox SelectColList 
      Appearance      =   0  'Плоска
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1464
      IntegralHeight  =   0   'False
      ItemData        =   "DiagMasterForm.frx":3BBC
      Left            =   2400
      List            =   "DiagMasterForm.frx":3BC3
      TabIndex        =   13
      Top             =   4500
      Width           =   5532
   End
   Begin VB.ListBox TableColList 
      Appearance      =   0  'Плоска
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1932
      IntegralHeight  =   0   'False
      ItemData        =   "DiagMasterForm.frx":3BD6
      Left            =   2400
      List            =   "DiagMasterForm.frx":3BDD
      TabIndex        =   11
      Top             =   2160
      Width           =   2532
   End
   Begin VB.ComboBox TableIndexCombo 
      Appearance      =   0  'Плоска
      BackColor       =   &H00F8F8F8&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      ItemData        =   "DiagMasterForm.frx":3BEF
      Left            =   2400
      List            =   "DiagMasterForm.frx":3BF1
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1740
      Width           =   2532
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
         Caption         =   "Draw yourself"
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
         Caption         =   "Мастер диаграмм"
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
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1980
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image6 
         Height          =   1440
         Left            =   420
         Picture         =   "DiagMasterForm.frx":3BF3
         Top             =   2130
         Width           =   1440
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Центровка
      Appearance      =   0  'Плоска
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Выбранные записи и исходные таблицы"
      ForeColor       =   &H80000008&
      Height          =   264
      Left            =   2460
      TabIndex        =   12
      Top             =   4200
      Width           =   4140
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Центровка
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Таблица"
      ForeColor       =   &H80000008&
      Height          =   264
      Left            =   2400
      TabIndex        =   9
      Top             =   1440
      Width           =   2544
   End
   Begin VB.Label CaptionLabel 
      Alignment       =   2  'Центровка
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Создание диаграммы"
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
      Caption         =   "Выберите источники данных а также желаемый тип и параметры диаграммы."
      ForeColor       =   &H00400000&
      Height          =   552
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
   Begin VB.Label OkBut 
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
   Begin VB.Image OkImg 
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
      Picture         =   "DiagMasterForm.frx":ACBD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5772
   End
   Begin VB.Image Image1 
      Height          =   660
      Left            =   0
      Picture         =   "DiagMasterForm.frx":ADDB
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
      X1              =   2210
      X2              =   2210
      Y1              =   660
      Y2              =   6600
   End
End
Attribute VB_Name = "DiagMasterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DiagData()

Private Sub DiagTypeCombo_Click()
 DiagTypeImage.Picture = DiagTypeImgs.ListImages(DiagTypeCombo.ListIndex + 1).Picture
 Select Case DiagTypeCombo.ListIndex
    Case 0, 2: Frame2.Visible = False
    Case 1, 3: Frame2.Visible = True
 End Select
End Sub

Private Sub Enabled3DCheck_Click()
 DimImg.Picture = DiagTypeImgs.ListImages(5 + Enabled3DCheck.value).Picture
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
    Case 27
        CancelBut_Click
    Case 13
        OkBut_Click
 End Select
End Sub

Private Sub Form_Load()
    Call ButEnabled(OkImg, OkBut, False)
    Call ButEnabled(CancelImg, CancelBut, True)
    TopImg.Picture = MainForm.TopImageList.ListImages(1).Picture
    DiagTypeCombo.ListIndex = 0
    DimImg.Picture = DiagTypeImgs.ListImages(5).Picture
    
    TableIndexCombo.Clear
    SelectColList.Clear
    For i% = 1 To MainForm.TabStrip.Tabs.Count
        TableIndexCombo.AddItem MainForm.TabStrip.Tabs(i).Caption
    Next i
    TableIndexCombo.ListIndex = 0
End Sub

' по строке "{x, YYY} ZZZ" возвращает номер таблицы ( x )
Sub GetTableIndex(ByVal str As String, TI As Integer)
    s$ = Trim$(Mid$(str, 2, InStr(1, str, ",") - 2))
    TI = CInt(s)
End Sub

' по строке "{x, YYY} ZZZ" и номеру таблицы возвращает номер поля с заголовком ZZZ
Sub GetColIndex(ByVal str As String, ByVal TI As Integer, CI As Integer)
    s$ = Trim$(Mid$(str, InStr(1, str, "}") + 1))
    For i% = 0 To DB(TI).Header.ColCount - 1
        If (s = Trim(DB(TI).Cols(i).title)) Then
            CI = i
            Exit Sub
        End If
    Next i
    CI = -1 '   событие невозможное но вероятное
End Sub

Function GettingDiagData(OnlyOneCol As Boolean) As Boolean
 GettingDiagData = False

 Dim TI As Integer, CI As Integer
 
 Select Case OnlyOneCol
    Case True   '   ************************************************************************
        Call GetTableIndex(SelectColList.List(0), TI)
        Call GetColIndex(SelectColList.List(0), TI, CI)
        ' зная номер таблицы и номер поля данных нужно проверить тип поля
        If (DB(TI).Cols(CI).Class <> ccInteger) Then
            Call MsgForm.ErrorMsg("Нельзя строить диаграмму по нечисленным данным!")
            Exit Function
        End If
        ' заполнение массива данных
        ReDim DiagData(2 * DB(TI).Header.RowCount)
        For i% = 0 To DB(TI).Header.RowCount - 1
            DiagData(2 * i) = DB(TI).Rows(i).Fields(CI)
            DiagData(2 * i + 1) = DiagData(2 * i)
        Next i
        GettingDiagData = True
        
    Case False   '   ************************************************************************
        ReDim DiagData(2 * SelectColList.ListCount)
        For R% = 0 To SelectColList.ListCount - 1
            Call GetTableIndex(SelectColList.List(R), TI)
            Call GetColIndex(SelectColList.List(R), TI, CI)
            ' зная номер таблицы и номер поля данных нужно проверить тип поля
            If (DB(TI).Cols(CI).Class <> ccInteger) Then
                Call MsgForm.ErrorMsg("Нельзя строить диаграмму по нечисленным данным!")
                Exit Function
            End If
            Dim Summary As Long
            Summary = 0
            For i% = 0 To DB(TI).Header.RowCount - 1
                Summary = Summary + DB(TI).Rows(i).Fields(CI)
            Next i
            ' заполнение массива данных
            DiagData(2 * R) = Summary
            DiagData(2 * R + 1) = MainForm.TabStrip.Tabs(TI + 1).Caption + "." + DB(TI).Cols(CI).title
        Next R
        GettingDiagData = True
 End Select
 
End Function

Private Sub OkBut_Click()
 If (OkBut.Tag = 0) Then Exit Sub
 Call SoundClick
 
 If GettingDiagData(SelectColList.ListCount = 1) Then
    Load DiagResForm
    Call DiagResForm.InitDiagData(DiagData, DiagTypeCombo.ListIndex, (Enabled3DCheck.value = 1))
    DiagResForm.Show vbModal
    Unload DiagResForm
 End If
End Sub

Private Sub CancelBut_Click()
 Call SoundClick
 Unload Me
End Sub

Private Sub TableColList_DblClick()
 i% = TableColList.ListIndex
 s$ = "{ " + CStr(TableIndexCombo.ListIndex) + " , " + TableIndexCombo.Text + " } " + TableColList.List(i)
 For j% = 0 To SelectColList.ListCount - 1
    If (SelectColList.List(j) = s) Then Exit Sub
 Next j
 Call ButEnabled(OkImg, OkBut, True)
 SelectColList.AddItem s
End Sub

Private Sub SelectColList_DblClick()
 If (SelectColList.ListIndex > -1) Then SelectColList.RemoveItem SelectColList.ListIndex
 Call ButEnabled(OkImg, OkBut, (SelectColList.ListCount > 0))
End Sub

Private Sub TableIndexCombo_Click()
 DBI% = TableIndexCombo.ListIndex
 TableColList.Clear
 For i% = 0 To DB(DBI).Header.ColCount - 1
    TableColList.AddItem DB(DBI).Cols(i).title
 Next i
 If (TableColList.ListCount > 0) Then TableColList.ListIndex = 0
End Sub
