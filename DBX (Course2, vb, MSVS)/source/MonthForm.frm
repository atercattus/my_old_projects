VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form MonthForm 
   BackColor       =   &H00DCDCDC&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Выбор даты"
   ClientHeight    =   4335
   ClientLeft      =   5760
   ClientTop       =   4860
   ClientWidth     =   5805
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
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
   ScaleHeight     =   4335
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Плоска
      BackColor       =   &H00DCDCDC&
      Caption         =   "Настройка"
      ForeColor       =   &H80000008&
      Height          =   2235
      Left            =   2760
      TabIndex        =   2
      Top             =   1320
      Width           =   2895
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Плоска
         BackColor       =   &H00DCDCDC&
         Caption         =   "Всегда редактировать даты с помощью текстового редактора."
         ForeColor       =   &H00000000&
         Height          =   675
         Left            =   60
         TabIndex        =   3
         Top             =   300
         Width           =   2655
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Прозрачно
         Caption         =   "ВНИМАНИЕ: При выборе данного пункта данное окно больше не будет появляться на протяжении сеанса работы."
         ForeColor       =   &H00808080&
         Height          =   1005
         Left            =   120
         TabIndex        =   4
         Top             =   1140
         Width           =   2640
         WordWrap        =   -1  'True
      End
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2310
      Left            =   180
      TabIndex        =   0
      Top             =   1260
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   4075
      _Version        =   393216
      ForeColor       =   0
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowWeekNumbers =   -1  'True
      StartOfWeek     =   51970050
      TitleBackColor  =   16744448
      TitleForeColor  =   8454143
      TrailingForeColor=   12632256
      CurrentDate     =   38306
      MinDate         =   -36522
   End
   Begin VB.Label EditBut 
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1980
      MousePointer    =   1  'Указатель
      TabIndex        =   9
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Центровка
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Текст"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   2040
      MousePointer    =   1  'Указатель
      TabIndex        =   10
      Top             =   3810
      Width           =   1470
   End
   Begin VB.Label YesBut 
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   300
      MousePointer    =   1  'Указатель
      TabIndex        =   7
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Центровка
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Принять"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   360
      MousePointer    =   1  'Указатель
      TabIndex        =   8
      Top             =   3810
      Width           =   1470
   End
   Begin VB.Label CancelBut 
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3960
      MousePointer    =   1  'Указатель
      TabIndex        =   5
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Image YesImg 
      Height          =   480
      Left            =   300
      MousePointer    =   10  'Стрелка вверх
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Центровка
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Отмена"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   4020
      MousePointer    =   1  'Указатель
      TabIndex        =   6
      Top             =   3810
      Width           =   1470
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Центровка
      AutoSize        =   -1  'True
      BackStyle       =   0  'Прозрачно
      Caption         =   $"MonthForm.frx":0000
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   5685
      WordWrap        =   -1  'True
   End
   Begin VB.Image CancelImg 
      Height          =   480
      Left            =   3960
      MousePointer    =   10  'Стрелка вверх
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Image EditImg 
      Height          =   480
      Left            =   1980
      MousePointer    =   10  'Стрелка вверх
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   1575
   End
End
Attribute VB_Name = "MonthForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public res%

Private Sub CancelBut_Click()
    Hide
End Sub

Private Sub EditBut_Click()
    res = -1
    Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
    Case 27
        CancelBut_Click
    Case 13
        YesBut_Click
    Case 8
        EditBut_Click
 End Select
End Sub

Private Sub Form_Load()
    Call ButEnabled(YesImg, YesBut, True)
    Call ButEnabled(EditImg, EditBut, True)
    Call ButEnabled(CancelImg, CancelBut, True)
    res = 0
End Sub

Private Sub YesBut_Click()
    res = 1
    Hide
End Sub
