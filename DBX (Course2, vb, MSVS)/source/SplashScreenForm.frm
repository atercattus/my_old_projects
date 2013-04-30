VERSION 5.00
Begin VB.Form SplashScreenForm 
   Appearance      =   0  'Плоска
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5640
   ClientLeft      =   6195
   ClientTop       =   5850
   ClientWidth     =   8235
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   8235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   1  'Привязать вверх
      Appearance      =   0  'Плоска
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Нет
      ForeColor       =   &H80000008&
      Height          =   5640
      Left            =   0
      Picture         =   "SplashScreenForm.frx":0000
      ScaleHeight     =   5640
      ScaleWidth      =   8235
      TabIndex        =   0
      Top             =   0
      Width           =   8235
      Begin VB.Label Label2 
         Alignment       =   1  'Правая привязка
         Appearance      =   0  'Плоска
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Прозрачно
         Caption         =   "1.1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   7500
         TabIndex        =   1
         Top             =   2340
         Width           =   525
      End
   End
End
Attribute VB_Name = "SplashScreenForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 27) Or (KeyCode = 13) Then
        MainForm.Show
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Label2.Caption = "v. " + CStr(App.Major) + "." + CStr(App.Minor)
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MDown(x, y)
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MMove(hwnd, x, y)
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MUp
End Sub
