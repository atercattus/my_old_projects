VERSION 5.00
Begin VB.Form AddColForm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00DCDCDC&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Добавление поля"
   ClientHeight    =   3675
   ClientLeft      =   3960
   ClientTop       =   3285
   ClientWidth     =   7395
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1680
      Top             =   1740
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DCDCDC&
      Caption         =   "Настройки"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2892
      Left            =   3660
      TabIndex        =   9
      Top             =   60
      Width           =   3672
      Begin VB.TextBox ColTitle 
         Appearance      =   0  'Плоска
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         HideSelection   =   0   'False
         Left            =   360
         TabIndex        =   1
         Tag             =   "0"
         Text            =   "заголовок"
         ToolTipText     =   "Начальное значение для всех полей этого столбца"
         Top             =   600
         Width           =   3132
      End
      Begin VB.TextBox InitValue 
         Appearance      =   0  'Плоска
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         HideSelection   =   0   'False
         Left            =   360
         MaxLength       =   4
         TabIndex        =   3
         Tag             =   "0"
         Text            =   "0"
         ToolTipText     =   "Начальное значение для всех полей этого столбца"
         Top             =   2340
         Width           =   3132
      End
      Begin VB.ComboBox ColType 
         Appearance      =   0  'Плоска
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "TableForm.frx":0000
         Left            =   360
         List            =   "TableForm.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Тип значений в этом столбце"
         Top             =   1440
         Width           =   3132
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Плоска
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Прозрачно
         Caption         =   "Заголовок:"
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
         Height          =   240
         Left            =   180
         TabIndex        =   12
         Top             =   276
         Width           =   984
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Плоска
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Прозрачно
         Caption         =   "Начальное значение:"
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
         Height          =   240
         Left            =   180
         TabIndex        =   11
         Top             =   2040
         Width           =   1932
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Плоска
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Прозрачно
         Caption         =   "Тип данных в поле:"
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
         Height          =   240
         Left            =   180
         TabIndex        =   10
         Top             =   1140
         Width           =   1836
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DCDCDC&
      Caption         =   "Положение относительно поля"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3552
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3492
      Begin VB.CheckBox OnlyEndCheck 
         Appearance      =   0  'Плоска
         BackColor       =   &H00DCDCDC&
         Caption         =   "Только в конец"
         ForeColor       =   &H80000008&
         Height          =   264
         Left            =   780
         TabIndex        =   15
         ToolTipText     =   "Поле будет добавляться только в конец"
         Top             =   3180
         Value           =   1  'Отмечено
         Width           =   1932
      End
      Begin VB.ListBox StCol 
         Appearance      =   0  'Плоска
         BackColor       =   &H00FAFAFA&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2472
         IntegralHeight  =   0   'False
         ItemData        =   "TableForm.frx":0023
         Left            =   120
         List            =   "TableForm.frx":0025
         TabIndex        =   4
         Top             =   300
         Width           =   3252
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Плоска
         BackColor       =   &H00DCDCDC&
         Caption         =   "после"
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   2280
         TabIndex        =   6
         Top             =   2880
         Value           =   -1  'True
         Width           =   924
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Плоска
         BackColor       =   &H00DCDCDC&
         Caption         =   "перед"
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   300
         TabIndex        =   5
         Top             =   2880
         Width           =   936
      End
   End
   Begin VB.Label CancelBut 
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      ForeColor       =   &H80000008&
      Height          =   492
      Left            =   5700
      MousePointer    =   1  'Указатель
      TabIndex        =   8
      Top             =   3060
      Width           =   1572
   End
   Begin VB.Label Label0 
      Alignment       =   2  'Центровка
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Отмена"
      ForeColor       =   &H80000008&
      Height          =   348
      Left            =   5760
      MousePointer    =   1  'Указатель
      TabIndex        =   14
      Top             =   3144
      Width           =   1476
   End
   Begin VB.Label CreateBut 
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      ForeColor       =   &H80000008&
      Height          =   492
      Left            =   3720
      MousePointer    =   1  'Указатель
      TabIndex        =   7
      Top             =   3060
      Width           =   1572
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Центровка
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Создать"
      ForeColor       =   &H80000008&
      Height          =   348
      Left            =   3780
      MousePointer    =   1  'Указатель
      TabIndex        =   13
      Top             =   3144
      Width           =   1476
   End
   Begin VB.Image CreateImg 
      Height          =   480
      Left            =   3720
      MousePointer    =   10  'Стрелка вверх
      Stretch         =   -1  'True
      Top             =   3060
      Width           =   1572
   End
   Begin VB.Image CancelImg 
      Height          =   480
      Left            =   5700
      MousePointer    =   10  'Стрелка вверх
      Stretch         =   -1  'True
      Top             =   3060
      Width           =   1572
   End
End
Attribute VB_Name = "AddColForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmp As String

Public Function AddColDlg(DBIndex%) As String
 tmp = ""
 With StCol
  .Clear
  For i% = 1 To DB(DBIndex).Header.ColCount
   .AddItem DB(DBIndex).Cols(i - 1).title
  Next
  .ListIndex = .ListCount - 1
 End With
 ColType.ListIndex = 0
 Me.Show vbModal
 AddColDlg = tmp
 Unload Me
End Function

Private Sub ColType_Click()
 ' изменение допустимых длин
 If Visible Then
  Select Case ColType.ListIndex
   Case ccInteger: InitValue.MaxLength = 4
   Case ccString: InitValue.MaxLength = 255
  End Select
 End If

' контроль ввода
 If Visible And (ColType.ListIndex = ccInteger) Then
  If (Not IsInteger(InitValue.Text)) Then InitValue.Text = "0"
 End If
End Sub

Private Sub CreateBut_Click()
 Call SoundClick
 s1$ = Trim(ColTitle.Text)
 Do While (s1 = "")
  s1 = Trim(InputForm.InputVal("Вы не ввели заголовок столбца. Повторите ввод."))
 Loop
 tmp$ = s1 + " , "
 Dim ct
 Dim s2
 Select Case ColType.ListIndex
  Case ccInteger
   t$ = Trim(InitValue.Text)
   If (Not IsInteger(t)) Then
        Call MsgForm.InfoMsg("Введённое значение не является целым числом. Преобразовано к '0'.")
        t = "0"
   End If
   tmp = tmp + " " + sI + " , " + t
  Case ccString
   t$ = Trim(InitValue.Text)
   If (t = "") Then t = " "
   tmp = tmp + " " + sS + " , " + t
 End Select
 Dim pos%
 If (OnlyEndCheck.value = 1) Then
    pos = -1
 Else
    pos = StCol.ListIndex
    If (Option2.value = True) Then pos = pos + 1
 End If
 tmp = tmp + " , " + CStr(pos)
 Hide
End Sub

Private Sub CancelBut_Click()
 Call SoundClick
 Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
    Case 27
        CancelBut_Click
    Case 13
        CreateBut_Click
 End Select
End Sub

Private Sub Form_Load()
    Call ButEnabled(CreateImg, CreateBut, True)
    Call ButEnabled(CancelImg, CancelBut, True)
End Sub
