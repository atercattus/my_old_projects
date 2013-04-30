VERSION 5.00
Begin VB.Form PasswordForm 
   Appearance      =   0  'Плоска
   BackColor       =   &H00DCDCDC&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "DBX"
   ClientHeight    =   6600
   ClientLeft      =   5070
   ClientTop       =   3000
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
   Begin VB.Frame Frame1 
      Appearance      =   0  'Плоска
      BackColor       =   &H00DCDCDC&
      BorderStyle     =   0  'Нет
      ForeColor       =   &H00DCDCDC&
      Height          =   4272
      Left            =   2340
      TabIndex        =   8
      Top             =   1680
      Width           =   5592
      Begin VB.TextBox PassText 
         Appearance      =   0  'Плоска
         Height          =   336
         IMEMode         =   3  'DISABLE
         Left            =   180
         MaxLength       =   255
         PasswordChar    =   "#"
         TabIndex        =   21
         Top             =   780
         Width           =   5232
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Плоска
         BackColor       =   &H00DCDCDC&
         BorderStyle     =   0  'Нет
         ForeColor       =   &H80000008&
         Height          =   1332
         Left            =   60
         TabIndex        =   11
         Top             =   1500
         Width           =   5352
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Плоска
            BackColor       =   &H00DCDCDC&
            Caption         =   "Вход только для просмотра"
            ForeColor       =   &H80000008&
            Height          =   372
            Left            =   240
            TabIndex        =   13
            Top             =   600
            Width           =   3192
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Плоска
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Прозрачно
            Caption         =   "Без ввода пароля доступ к базе доступен в режиме Только чтение."
            ForeColor       =   &H80000008&
            Height          =   528
            Left            =   60
            TabIndex        =   12
            Top             =   0
            Width           =   5316
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Label NoFullLabel 
         Appearance      =   0  'Плоска
         BackColor       =   &H80000005&
         BackStyle       =   0  'Прозрачно
         Caption         =   "Внимание: данная БД не допускает просмотр с ограниченными правами"
         ForeColor       =   &H000000FF&
         Height          =   564
         Left            =   180
         TabIndex        =   10
         Top             =   3060
         Visible         =   0   'False
         Width           =   5028
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Плоска
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Прозрачно
         Caption         =   "Введите пароль для получения полных прав"
         ForeColor       =   &H80000008&
         Height          =   264
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   4512
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Плоска
      BackColor       =   &H00DCDCDC&
      BorderStyle     =   0  'Нет
      ForeColor       =   &H80000008&
      Height          =   4272
      Left            =   2340
      TabIndex        =   14
      Top             =   1680
      Visible         =   0   'False
      Width           =   5652
      Begin VB.TextBox SetPassText 
         Appearance      =   0  'Плоска
         BackColor       =   &H00FFFFFF&
         Height          =   336
         IMEMode         =   3  'DISABLE
         Left            =   180
         MaxLength       =   255
         PasswordChar    =   "#"
         TabIndex        =   20
         Top             =   420
         Width           =   5292
      End
      Begin VB.CheckBox CheckNoRO 
         Appearance      =   0  'Плоска
         BackColor       =   &H00DCDCDC&
         Caption         =   "Запретить доступ к БД без ввода пароля"
         ForeColor       =   &H80000008&
         Height          =   372
         Left            =   240
         TabIndex        =   19
         Top             =   2100
         Width           =   4992
      End
      Begin VB.CheckBox CheckCoded 
         Appearance      =   0  'Плоска
         BackColor       =   &H00DCDCDC&
         Caption         =   "Шифровать данные в БД по паролю"
         ForeColor       =   &H00000000&
         Height          =   372
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Width           =   4992
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Плоска
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Прозрачно
         Caption         =   "Вся информация будет зашифрована. Без пароля будет невозможно восстановить содердимое БД."
         ForeColor       =   &H00000000&
         Height          =   528
         Left            =   120
         TabIndex        =   22
         Top             =   1380
         Width           =   5568
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Плоска
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Прозрачно
         Caption         =   $"PasswordForm.frx":0000
         ForeColor       =   &H80000008&
         Height          =   1584
         Left            =   120
         TabIndex        =   16
         Top             =   2580
         Width           =   5412
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Плоска
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Прозрачно
         Caption         =   "Установка пароля для доступа к БД"
         ForeColor       =   &H80000008&
         Height          =   264
         Left            =   120
         TabIndex        =   15
         Top             =   60
         Width           =   3660
      End
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
         Caption         =   "Protect yourself"
         Enabled         =   0   'False
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
         Height          =   228
         Left            =   60
         TabIndex        =   18
         Top             =   5640
         Width           =   2088
         WordWrap        =   -1  'True
      End
      Begin VB.Label MainLabel 
         Alignment       =   2  'Центровка
         Appearance      =   0  'Плоска
         BackColor       =   &H80000005&
         BackStyle       =   0  'Прозрачно
         Caption         =   "Мастер защиты и безопасности базы данных"
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
         Height          =   1440
         Left            =   420
         Picture         =   "PasswordForm.frx":010E
         Top             =   2160
         Width           =   1440
      End
   End
   Begin VB.Label CaptionLabel 
      Alignment       =   2  'Центровка
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Заголовок сообщения"
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
      Caption         =   "Текст сообщения"
      ForeColor       =   &H00400000&
      Height          =   912
      Left            =   2340
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
      Picture         =   "PasswordForm.frx":71D8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5772
   End
   Begin VB.Image Image1 
      Height          =   660
      Left            =   0
      Picture         =   "PasswordForm.frx":72F6
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
End
Attribute VB_Name = "PasswordForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public res As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
    Case 27
        CancelBut_Click
    Case 13
        OkBut_Click
 End Select
End Sub

Private Sub Form_Activate()
 res = False
 If Frame1.Visible Then
  PassText.SetFocus
 Else
  SetPassText.SetFocus
 End If
End Sub

Private Sub Form_Load()
    Call ButEnabled(OkImg, OkBut, True)
    Call ButEnabled(CancelImg, CancelBut, True)
    TopImg.Picture = MainForm.TopImageList.ListImages(1).Picture
End Sub

Private Sub OkBut_Click()
 res = True
 Call SoundClick
 Hide
End Sub

Private Sub CancelBut_Click()
 Call SoundClick
 Hide
End Sub

Private Sub PassText_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 13) Then Call OkBut_Click
End Sub

Private Sub SetPassText_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 13) Then Call OkBut_Click
End Sub
