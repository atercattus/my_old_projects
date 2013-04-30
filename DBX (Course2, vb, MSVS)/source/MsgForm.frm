VERSION 5.00
Begin VB.Form MsgForm 
   Appearance      =   0  'Плоска
   BackColor       =   &H00DCDCDC&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "DBX"
   ClientHeight    =   2460
   ClientLeft      =   5940
   ClientTop       =   4050
   ClientWidth     =   6390
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   10.5
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
   ScaleHeight     =   2460
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame YesFrame 
      BackColor       =   &H00DCDCDC&
      BorderStyle     =   0  'Нет
      Height          =   732
      Left            =   360
      TabIndex        =   7
      Top             =   1620
      Width           =   1692
      Begin VB.Label YesBut 
         Appearance      =   0  'Плоска
         BackColor       =   &H80000005&
         BackStyle       =   0  'Прозрачно
         ForeColor       =   &H80000008&
         Height          =   492
         Left            =   60
         MousePointer    =   1  'Указатель
         TabIndex        =   8
         Top             =   180
         Width           =   1572
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Центровка
         Appearance      =   0  'Плоска
         BackColor       =   &H80000005&
         BackStyle       =   0  'Прозрачно
         Caption         =   "Да"
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
         Left            =   120
         MousePointer    =   1  'Указатель
         TabIndex        =   9
         Top             =   240
         Width           =   1476
      End
      Begin VB.Image YesImg 
         Height          =   480
         Left            =   60
         MousePointer    =   10  'Стрелка вверх
         Stretch         =   -1  'True
         Top             =   180
         Width           =   1572
      End
   End
   Begin VB.Frame CancelFrame 
      BackColor       =   &H00DCDCDC&
      BorderStyle     =   0  'Нет
      Height          =   732
      Left            =   2400
      TabIndex        =   4
      Top             =   1620
      Width           =   1692
      Begin VB.Label CancelBut 
         Appearance      =   0  'Плоска
         BackColor       =   &H80000005&
         BackStyle       =   0  'Прозрачно
         ForeColor       =   &H80000008&
         Height          =   492
         Left            =   60
         MousePointer    =   1  'Указатель
         TabIndex        =   5
         Top             =   180
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
            Size            =   13.5
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   348
         Left            =   120
         MousePointer    =   1  'Указатель
         TabIndex        =   6
         Top             =   240
         Width           =   1476
      End
      Begin VB.Image CancelImg 
         Height          =   480
         Left            =   60
         MousePointer    =   10  'Стрелка вверх
         Stretch         =   -1  'True
         Top             =   180
         Width           =   1572
      End
   End
   Begin VB.Frame NoFrame 
      BackColor       =   &H00DCDCDC&
      BorderStyle     =   0  'Нет
      Height          =   732
      Left            =   4380
      TabIndex        =   1
      Top             =   1620
      Width           =   1692
      Begin VB.Label NoBut 
         Appearance      =   0  'Плоска
         BackColor       =   &H80000005&
         BackStyle       =   0  'Прозрачно
         ForeColor       =   &H80000008&
         Height          =   492
         Left            =   60
         MousePointer    =   1  'Указатель
         TabIndex        =   2
         Top             =   180
         Width           =   1572
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Центровка
         Appearance      =   0  'Плоска
         BackColor       =   &H80000005&
         BackStyle       =   0  'Прозрачно
         Caption         =   "Нет"
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
         Left            =   120
         MousePointer    =   1  'Указатель
         TabIndex        =   3
         Top             =   240
         Width           =   1476
      End
      Begin VB.Image NoImg 
         Height          =   480
         Left            =   60
         MousePointer    =   10  'Стрелка вверх
         Stretch         =   -1  'True
         Top             =   180
         Width           =   1572
      End
   End
   Begin VB.Image QuestImage 
      Height          =   1515
      Left            =   60
      Picture         =   "MsgForm.frx":0000
      Stretch         =   -1  'True
      Top             =   60
      Width           =   1515
   End
   Begin VB.Image InfoImage 
      Height          =   1515
      Left            =   60
      Picture         =   "MsgForm.frx":11A2
      Stretch         =   -1  'True
      Top             =   60
      Width           =   1515
   End
   Begin VB.Image ErrImage 
      Height          =   1515
      Left            =   56
      Picture         =   "MsgForm.frx":2324
      Stretch         =   -1  'True
      Top             =   56
      Width           =   1515
   End
   Begin VB.Label Text 
      Alignment       =   2  'Центровка
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Text"
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
      Height          =   1330
      Left            =   1554
      TabIndex        =   0
      Top             =   112
      UseMnemonic     =   0   'False
      Width           =   4704
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "MsgForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim res As Byte

Public Function ErrorMsg(str$) As Integer
    Caption = "Ошибка"
    Text = str
    
    YesFrame.Visible = True
    NoFrame.Visible = False
    CancelFrame.Visible = False

    InfoImage.Visible = False
    ErrImage.Visible = True
    QuestImage.Visible = False

    YesFrame.Move 2400
    res = resBad
    Call sndPlaySound("Data\Error.wav", SND_ASYNC + SND_FILENAME + SND_LOOP + SND_APPLICATION)
    Show vbModal
    ErrorMsg = res
    Unload Me
End Function

Public Function InfoMsg(str$) As Integer
    Caption = "Информация"
    Text = str
    
    YesFrame.Visible = True
    NoFrame.Visible = False
    CancelFrame.Visible = False

    InfoImage.Visible = True
    ErrImage.Visible = False
    QuestImage.Visible = False
    
    YesFrame.Move 2400

    res = 0
    Call sndPlaySound("Data\Info.wav", SND_ASYNC + SND_FILENAME + SND_LOOP + SND_APPLICATION)
    Show vbModal
    InfoMsg = res
    Unload Me
End Function

Public Function QuestMsg(str$, Optional showcancel As Boolean = False) As Integer
    Caption = "Вопрос"
    Text = str
    
    If showcancel Then
        YesFrame.Visible = True
        NoFrame.Visible = True
        CancelFrame.Visible = True
        
        YesFrame.Move 360
        NoFrame.Move 4380
        CancelFrame.Move 2400
        
    Else
        YesFrame.Visible = True
        NoFrame.Visible = True
        CancelFrame.Visible = False
        
        YesFrame.Move 900
        NoFrame.Move 3840
    End If

    InfoImage.Visible = False
    ErrImage.Visible = False
    QuestImage.Visible = True

    res = 0
    Call sndPlaySound("Data\Quest.wav", SND_ASYNC + SND_FILENAME + SND_LOOP + SND_APPLICATION)
    Show vbModal
    QuestMsg = res
    Unload Me
End Function

Private Sub CancelBut_Click()
    res = resCancel
    Call SoundClick
    Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13
            Call YesBut_Click
        Case 27
            If (NoFrame.Visible = True) Then Call NoBut_Click
        Case 8
            If (CancelFrame.Visible = True) Then Call CancelBut_Click
    End Select
End Sub

Private Sub Form_Load()
    Call ButEnabled(YesImg, YesBut, True)
    Call ButEnabled(CancelImg, CancelBut, True)
    Call ButEnabled(NoImg, NoBut, True)
End Sub

Private Sub NoBut_Click()
    res = resNo
    Call SoundClick
    Hide
End Sub

Private Sub YesBut_Click()
    res = resOk
    Call SoundClick
    Hide
End Sub

