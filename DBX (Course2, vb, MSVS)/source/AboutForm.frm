VERSION 5.00
Begin VB.Form AboutForm 
   Appearance      =   0  'Плоска
   BackColor       =   &H00DCDCDC&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3195
   ClientLeft      =   5025
   ClientTop       =   4215
   ClientWidth     =   7680
   ClipControls    =   0   'False
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
   ScaleHeight     =   3195
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label6 
      Appearance      =   0  'Плоска
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      Caption         =   "v. 1.1"
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
      Height          =   300
      Left            =   165
      TabIndex        =   7
      Top             =   2805
      Width           =   630
   End
   Begin VB.Image Image2 
      Height          =   735
      Left            =   5370
      Picture         =   "AboutForm.frx":0000
      Top             =   1515
      Width           =   1830
   End
   Begin VB.Label NoViewLabel 
      BackStyle       =   0  'Прозрачно
      Height          =   2170
      Left            =   4200
      TabIndex        =   6
      Top             =   392
      Width           =   3430
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Плоска
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      Caption         =   "IX..XI MMIV"
      ForeColor       =   &H80000008&
      Height          =   266
      Left            =   6328
      TabIndex        =   5
      Top             =   2226
      Width           =   1288
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Плоска
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Автор"
      ForeColor       =   &H80000008&
      Height          =   266
      Left            =   4676
      TabIndex        =   4
      Top             =   1624
      Width           =   630
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Центровка
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Работа с однотабличной ненормализованной БД"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   204
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   700
      Left            =   4130
      TabIndex        =   2
      Top             =   420
      Width           =   3318
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Центровка
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Работа с однотабличной ненормализованной БД"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   204
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D0D0D0&
      Height          =   700
      Left            =   4074
      TabIndex        =   3
      Top             =   476
      Width           =   3318
      WordWrap        =   -1  'True
   End
   Begin VB.Label OkBut 
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   490
      Left            =   3108
      MousePointer    =   1  'Указатель
      TabIndex        =   0
      Top             =   2632
      Width           =   1568
   End
   Begin VB.Label Label1 
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
      Height          =   350
      Left            =   3164
      MousePointer    =   1  'Указатель
      TabIndex        =   1
      Top             =   2702
      Width           =   1470
   End
   Begin VB.Image OkImg 
      Height          =   476
      Left            =   3108
      MousePointer    =   10  'Стрелка вверх
      Stretch         =   -1  'True
      Top             =   2632
      Width           =   1568
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   2786
      Left            =   56
      Picture         =   "AboutForm.frx":1894
      Stretch         =   -1  'True
      Top             =   56
      Width           =   4270
   End
End
Attribute VB_Name = "AboutForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If (KeyCode = 13) Or (KeyCode = 27) Then OkBut_Click
End Sub

Private Sub Form_Load()
    Call MInit
    Call ButEnabled(OkImg, OkBut, True)
    Label6.Caption = "v. " + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MDown(x, y)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MMove(hwnd, x, y)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MUp
End Sub

Private Sub Image2_Click()
    Call ShellExecute(0, "", "mailto:xerx@nightmail.ru", "", "", 1)
End Sub

Private Sub NoViewLabel_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MDown(x, y)
End Sub

Private Sub NoViewLabel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MMove(hwnd, x, y)
End Sub

Private Sub NoViewLabel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MUp
End Sub

Private Sub OkBut_Click()
 Unload Me
End Sub
