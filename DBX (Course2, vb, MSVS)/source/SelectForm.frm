VERSION 5.00
Begin VB.Form SelectForm 
   BackColor       =   &H00DCDCDC&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Выбор"
   ClientHeight    =   4410
   ClientLeft      =   4245
   ClientTop       =   3540
   ClientWidth     =   7515
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
   ScaleHeight     =   4410
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List2 
      Appearance      =   0  'Плоска
      BackColor       =   &H00F5F5F5&
      ForeColor       =   &H00800000&
      Height          =   3372
      IntegralHeight  =   0   'False
      ItemData        =   "SelectForm.frx":0000
      Left            =   60
      List            =   "SelectForm.frx":0002
      MultiSelect     =   1  'Просто
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   7392
   End
   Begin VB.CheckBox CheckConfirm 
      Appearance      =   0  'Плоска
      BackColor       =   &H00DCDCDC&
      Caption         =   "Требовать подтверждения"
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
      Height          =   312
      Left            =   120
      TabIndex        =   2
      Top             =   3920
      Value           =   1  'Отмечено
      Width           =   2592
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Плоска
      BackColor       =   &H00F5F5F5&
      ForeColor       =   &H00800000&
      Height          =   3372
      IntegralHeight  =   0   'False
      ItemData        =   "SelectForm.frx":0004
      Left            =   60
      List            =   "SelectForm.frx":0006
      TabIndex        =   1
      Top             =   360
      Width           =   7392
   End
   Begin VB.Label CancelBut 
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      ForeColor       =   &H80000008&
      Height          =   490
      Left            =   5768
      MousePointer    =   1  'Указатель
      TabIndex        =   5
      Top             =   3836
      Width           =   1568
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Центровка
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Отмена"
      ForeColor       =   &H80000008&
      Height          =   348
      Left            =   5820
      MousePointer    =   1  'Указатель
      TabIndex        =   6
      Top             =   3936
      Width           =   1476
   End
   Begin VB.Label SelectBut 
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      ForeColor       =   &H80000008&
      Height          =   490
      Left            =   3976
      MousePointer    =   1  'Указатель
      TabIndex        =   3
      Top             =   3808
      Width           =   1568
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Центровка
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Выбрать"
      ForeColor       =   &H80000008&
      Height          =   348
      Left            =   4020
      MousePointer    =   1  'Указатель
      TabIndex        =   4
      Top             =   3936
      Width           =   1476
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Центровка
      BackStyle       =   0  'Прозрачно
      Caption         =   "Label1"
      Height          =   324
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7464
   End
   Begin VB.Image SelectImg 
      Height          =   480
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   1572
   End
   Begin VB.Image CancelImg 
      Height          =   480
      Left            =   5760
      MousePointer    =   10  'Стрелка вверх
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   1572
   End
End
Attribute VB_Name = "SelectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmp%, tmps$

Public Function SelectDlg(DBIndex%, ByVal title$, ByVal what$) As Integer
 Dim s$
 List1.Visible = True
 List2.Visible = False
 List1.Clear
 Select Case what
    Case sRow   '   ******************* ...::: Select Row :::... ********************
        With MainForm.ListView.ListItems
            For i% = 1 To .Count
                s = CStr(i - 1) + ") " + .Item(i)
                For j% = 1 To DB(DBIndex).Header.ColCount - 1
                    s = s + " - " + .Item(i).SubItems(j)
                Next j
                List1.AddItem s
            Next i
        End With
        
    Case sCol   '   ******************* ...::: Select Col :::... ********************
        With MainForm.ListView.ColumnHeaders
            For i% = 1 To .Count
                List1.AddItem CStr(i - 1) + ") " + .Item(i)
            Next i
        End With
        
    Case sTable   '   ******************* ...::: Select Table :::... ********************
        For i% = 0 To (MainForm.TabStrip.Tabs.Count - 1)
            List1.AddItem CStr(i) + ") " + MainForm.TabStrip.Tabs.Item(i + 1)
        Next i
 End Select

 If (List1.ListCount > 0) Then
  List1.ListIndex = 0
  Call ButEnabled(SelectImg, SelectBut, True)
 Else
  Call ButEnabled(SelectImg, SelectBut, False)
 End If
 Label1.Caption = title
 tmp = -1
 Show vbModal
 SelectDlg = CStr(tmp)
End Function

Public Function MultiSelectDlg(DBIndex%, ByVal title$, ByVal what$) As String
 Dim s$
 List2.Visible = True
 List1.Visible = False
 List2.Clear
 CheckConfirm.Visible = False
 If (what = sRow) Then
  With MainForm.ListView.ListItems
   For i% = 1 To .Count
    s = CStr(i - 1) + ") " + .Item(i)
    For j% = 1 To DB(DBIndex).Header.ColCount - 1
     s = s + " - " + .Item(i).SubItems(j)
    Next j
    List2.AddItem s
   Next i
  End With
 Else
  With MainForm.ListView.ColumnHeaders
   For i% = 1 To .Count
    List2.AddItem CStr(i - 1) + ") " + .Item(i)
   Next i
  End With
 End If
 Call ButEnabled(SelectImg, SelectBut, False)
 Label1.Caption = title
 tmps = ""
 Show vbModal
 CheckConfirm.Visible = True
 MultiSelectDlg = tmps
End Function

Private Sub Form_Activate()
    Call ButEnabled(CancelImg, CancelBut, True)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If (KeyCode = 27) Then CancelBut_Click Else If (KeyCode = 13) Then SelectBut_Click
End Sub

Private Sub SelectBut_Click()
 If (SelectBut.Tag = 0) Then Exit Sub
 If (List1.Visible) Then
    tmp = List1.ListIndex
 Else
    For i = 0 To List2.ListCount - 1
        If List2.Selected(i) Then tmps = tmps + CStr(i) + ","
    Next i
    tmps = Strings.Left$(tmps, Len(tmps) - 1)
 End If
 Hide
End Sub

Private Sub CancelBut_Click()
 Hide
End Sub

Private Sub List1_Click()
    Call ButEnabled(SelectImg, SelectBut, (List1.ListIndex <> -1))
End Sub

Private Sub List2_Click()
    Call ButEnabled(SelectImg, SelectBut, (List2.SelCount = 2))
End Sub
