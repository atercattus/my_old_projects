VERSION 5.00
Begin VB.Form InputForm 
   Appearance      =   0  'Плоска
   BackColor       =   &H00DCDCDC&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ввод значения"
   ClientHeight    =   1728
   ClientLeft      =   4344
   ClientTop       =   5088
   ClientWidth     =   6444
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   10.8
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
   ScaleHeight     =   1728
   ScaleWidth      =   6444
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Плоска
      Height          =   312
      Left            =   420
      TabIndex        =   1
      Top             =   660
      Width           =   5532
   End
   Begin VB.Label CancelBut 
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      ForeColor       =   &H80000008&
      Height          =   492
      Left            =   4140
      MousePointer    =   1  'Указатель
      TabIndex        =   4
      Top             =   1080
      Width           =   1572
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
      Height          =   348
      Left            =   4200
      MousePointer    =   1  'Указатель
      TabIndex        =   5
      Top             =   1176
      Width           =   1476
   End
   Begin VB.Label YesBut 
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      ForeColor       =   &H80000008&
      Height          =   492
      Left            =   660
      MousePointer    =   1  'Указатель
      TabIndex        =   2
      Top             =   1080
      Width           =   1572
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
      Height          =   348
      Left            =   720
      MousePointer    =   1  'Указатель
      TabIndex        =   3
      Top             =   1176
      Width           =   1476
   End
   Begin VB.Image CancelImg 
      Height          =   480
      Left            =   4140
      MousePointer    =   10  'Стрелка вверх
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   1572
   End
   Begin VB.Image YesImg 
      Height          =   480
      Left            =   660
      MousePointer    =   10  'Стрелка вверх
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   1572
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      ForeColor       =   &H80000008&
      Height          =   492
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6312
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "InputForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim res%

Private Sub CancelBut_Click()
    Call SoundClick
    Hide
End Sub

Private Sub Form_Activate()
    Text1.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13: Call YesBut_Click
        Case 27: Call CancelBut_Click
    End Select
End Sub

Private Sub Form_Load()
    Call ButEnabled(YesImg, YesBut, True)
    Call ButEnabled(CancelImg, CancelBut, True)
End Sub

Public Function InputVal(str$) As String
    Label1.Caption = str
    Text1.Text = ""
    res = 0
    Me.Show vbModal
    If (res = 1) Then InputVal = Text1.Text
    Unload Me
End Function

Private Sub YesBut_Click()
    Call SoundClick
    res = 1
    Hide
End Sub
