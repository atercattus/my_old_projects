VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form MainForm 
   Appearance      =   0  'Плоска
   AutoRedraw      =   -1  'True
   Caption         =   "DB Xtension"
   ClientHeight    =   5070
   ClientLeft      =   5985
   ClientTop       =   5205
   ClientWidth     =   8475
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "MainForm"
   ScaleHeight     =   5070
   ScaleWidth      =   8475
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView ListView 
      Height          =   3645
      Left            =   105
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   8310
      _ExtentX        =   14658
      _ExtentY        =   6429
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman CYR"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Timer CoolTimer 
      Interval        =   10
      Left            =   6660
      Top             =   2340
   End
   Begin MSComctlLib.ImageList CoolImgs 
      Left            =   1500
      Top             =   1980
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":19A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":227E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":2F58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":3C32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":450C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":51E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":5EC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":679A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":7474
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":814E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":8A28
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":9702
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":A3DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":ACB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":B990
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":C66A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":CF44
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1500
      Top             =   1320
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":DC1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":E1B8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog HTMLPath 
      Left            =   5040
      Top             =   3840
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      DefaultExt      =   "*.html"
      DialogTitle     =   "Путь к HTML"
      Filter          =   "HTML файл|*.html"
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   615
      Left            =   -30
      TabIndex        =   1
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   1085
      BandCount       =   1
      ForeColor       =   -2147483633
      EmbossHighlight =   -2147483633
      _CBWidth        =   8175
      _CBHeight       =   615
      _Version        =   "6.0.8169"
      BandBackColor1  =   15132390
      MinHeight1      =   555
      Width1          =   8115
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      BandStyle1      =   1
      Begin VB.Frame Frame1 
         Appearance      =   0  'Плоска
         BorderStyle     =   0  'Нет
         ForeColor       =   &H8000000F&
         Height          =   540
         Left            =   120
         TabIndex        =   2
         Top             =   40
         Width           =   3915
         Begin VB.Image CoolBut 
            Appearance      =   0  'Плоска
            Height          =   480
            Index           =   0
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   480
         End
         Begin VB.Image CoolBut 
            Appearance      =   0  'Плоска
            Height          =   480
            Index           =   1
            Left            =   600
            Stretch         =   -1  'True
            Top             =   0
            Width           =   480
         End
         Begin VB.Image CoolBut 
            Appearance      =   0  'Плоска
            Enabled         =   0   'False
            Height          =   480
            Index           =   2
            Left            =   1200
            Stretch         =   -1  'True
            Top             =   0
            Width           =   480
         End
         Begin VB.Image CoolBut 
            Appearance      =   0  'Плоска
            Enabled         =   0   'False
            Height          =   480
            Index           =   3
            Left            =   1800
            Stretch         =   -1  'True
            Top             =   0
            Width           =   480
         End
         Begin VB.Image CoolBut 
            Appearance      =   0  'Плоска
            Enabled         =   0   'False
            Height          =   480
            Index           =   4
            Left            =   2400
            Stretch         =   -1  'True
            Top             =   0
            Width           =   480
         End
         Begin VB.Image CoolBut 
            Appearance      =   0  'Плоска
            Height          =   480
            Index           =   5
            Left            =   3000
            Stretch         =   -1  'True
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin MSComDlg.CommonDialog Dlgs 
      Left            =   5520
      Top             =   3300
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      DefaultExt      =   "dbx"
      DialogTitle     =   "Выберите БД"
      Filter          =   "Файл базы данных (DBX)|*.dbx"
   End
   Begin MSComctlLib.ImageList TopImageList 
      Left            =   540
      Top             =   1320
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   135
      ImageHeight     =   55
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":E552
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ButtonImageList 
      Left            =   540
      Top             =   1980
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   255
      ImageWidth      =   131
      ImageHeight     =   40
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":13D4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":1436F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Привязать вниз
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   4770
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   529
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   476
            MinWidth        =   476
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   952
            MinWidth        =   952
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   11933
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Bevel           =   0
            Object.Width           =   970
            MinWidth        =   970
            TextSave        =   "11:37"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TabStrip TabStrip 
      Height          =   4095
      Left            =   60
      TabIndex        =   3
      Top             =   675
      Visible         =   0   'False
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   7223
      TabWidthStyle   =   1
      Style           =   2
      HotTracking     =   -1  'True
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Главная таблица"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VB.Menu File 
      Caption         =   "Файл"
      Begin VB.Menu CreateDB 
         Caption         =   "Создать"
         Shortcut        =   ^C
      End
      Begin VB.Menu OpenDB 
         Caption         =   "Открыть"
         Shortcut        =   ^O
      End
      Begin VB.Menu SaveDB 
         Caption         =   "Сохранить"
         Shortcut        =   ^S
      End
      Begin VB.Menu CloseDB 
         Caption         =   "Закрыть"
         Shortcut        =   ^F
      End
      Begin VB.Menu ResCopyDB 
         Caption         =   "Резервная копия"
         Shortcut        =   ^R
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu ExitPr 
         Caption         =   "Выход"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu QueryDB 
      Caption         =   "Запросы"
      Begin VB.Menu QueryM 
         Caption         =   "Мастер запросов"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu ResDB 
      Caption         =   "Результаты"
      Begin VB.Menu DiagDraw 
         Caption         =   "Мастер диаграмм"
         Shortcut        =   ^D
      End
      Begin VB.Menu HTMLCreator 
         Caption         =   "Формирование HTML"
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu OptDB 
      Caption         =   "Настройки"
      Begin VB.Menu Security 
         Caption         =   "Защита"
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu Help 
      Caption         =   "?"
      Begin VB.Menu AboutProg 
         Caption         =   "О программе"
         Shortcut        =   ^A
      End
      Begin VB.Menu HelpProg 
         Caption         =   "Помощь"
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu TSMenu 
      Caption         =   "TabStripMenu"
      Visible         =   0   'False
      Begin VB.Menu TSClose 
         Caption         =   "Закрыть"
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' разница ширины и высоты формы и TabStrip'а
Dim dW1%, dH1%
' разница ширины и высоты TabStrip'а и ListView'а
Dim dW2%, dH2%
' последний выбранный элемент
Dim saveItemIndex%
' текущая таблица
Public DBCurIndex%
' последний Image, над которым был курсор
Dim OldImageIndex%

Private Sub AboutProg_Click()
 CoolTimer.Enabled = False
 AboutForm.Show vbModal
 CoolTimer.Enabled = True
End Sub

Private Sub CloseDB_Click()
    CoolTimer.Enabled = False
    
    If DBChanged Then
        If (MsgForm.QuestMsg("В БД внесены не сохранённые изменения. Закрыть не сохраняя?") <> resOk) Then GoTo exit_
    End If
    
    SB.Panels(3).Text = ""
    Call ClearAll
    Call ShowTable(-1)
    Call DisEnImage(2, 1)
    Call DisEnImage(3, 1)
    Call DisEnImage(4, 1)
    
exit_:

 CoolTimer.Enabled = True
End Sub

' index,mode / сегмент,смещение
Sub DisEnImage(Index%, Mode%)
    CoolBut(Index).Picture = CoolImgs.ListImages(1 + (Index * 3 + Mode)).Picture
    CoolBut(Index).Enabled = (Mode <> 1)
End Sub

Sub RetImage()
    If (OldImageIndex > -1) Then
        If CoolBut(OldImageIndex).Enabled Then
            Call DisEnImage(OldImageIndex, 0)
        Else
            Call DisEnImage(OldImageIndex, 1)
        End If
    End If
    OldImageIndex = -1
End Sub

Sub CoolMouseMove(Index%)
 If (Index = OldImageIndex) Then Exit Sub
 Call DisEnImage(Index, 2)
 Call RetImage
 OldImageIndex = Index
End Sub

Private Sub CoolBut_Click(Index As Integer)
    Call DisEnImage(Index, 0)
    Select Case Index
        Case 0: Call CreateDB_Click
        Case 1: Call OpenDB_Click
        Case 2: Call SaveDB_Click
        Case 3: Call CloseDB_Click
        Case 4: Call ResCopyDB_Click
        Case 5: Call ExitPr_Click
    End Select
End Sub

Private Sub CoolTimer_Timer()
 Dim Point As POINTAPI
 Dim R As RECT, R2 As RECT
 Call GetCursorPos(Point)
 Call GetWindowRect(Frame1.hwnd, R)
 For i% = 0 To 5
    If (Not CoolBut(i).Enabled) Then GoTo loop_
    x% = R.Left + CoolBut(i).Left / Screen.TwipsPerPixelX
    y% = R.Top + CoolBut(i).Top / Screen.TwipsPerPixelY
    X2% = x + CoolBut(i).Width / Screen.TwipsPerPixelX
    Y2% = y + CoolBut(i).Height / Screen.TwipsPerPixelY
    R2.Left = x
    R2.Top = y
    R2.Right = X2
    R2.Bottom = Y2
    If ((Point.x >= R2.Left) And (Point.x <= R2.Right) And (Point.y >= R2.Top) And (Point.y <= R2.Bottom)) Then
        Call CoolMouseMove(i)
        Exit Sub
    End If
loop_:
 Next i
 Call RetImage
End Sub

Private Sub CreateDB_Click()
 CoolTimer.Enabled = False
 Dlgs.FileName = ""
 Dlgs.ShowSave
 If (Dlgs.FileName <> "") Then
  ' создаю новую БД
  Call NewDB(Dlgs.FileName)
  ' вывожу путь к БД
  SB.Panels(3).Text = DBPath
  ' разрешения
  Call DisEnImage(2, 0)
  Call DisEnImage(3, 0)
  Call DisEnImage(4, 0)
  Call ShowTable(DBCurIndex)
 End If
 CoolTimer.Enabled = True
End Sub

Private Sub DiagDraw_Click()
 CoolTimer.Enabled = False
 DiagMasterForm.Show vbModal
 CoolTimer.Enabled = True
End Sub

Private Sub ExitBut_Click()
 Call ExitPr_Click
End Sub

Private Sub ExitPr_Click()
    CoolTimer.Enabled = False
    If Not DBChanged Then
        End
    Else
        If (MsgForm.QuestMsg("В БД внесены не сохранённые изменения. Выйти не сохраняя?") = resOk) Then End
    End If
    CoolTimer.Enabled = True
End Sub

Private Sub File_Click()
 SaveDB.Enabled = DBPath <> ""
 CloseDB.Enabled = SaveDB.Enabled
 ResCopyDB.Enabled = SaveDB.Enabled
End Sub

Private Sub HelpProg_Click()
 CoolTimer.Enabled = False
 Call ShellExecute(hwnd, "open", "Help\index.html", "", "", 0)
 CoolTimer.Enabled = True
End Sub

Sub CreateHTML(Path$)
 Call DeleteFile(Path)
 DBI% = FreeFile
 Open Path For Binary As DBI
 
 Capt$ = InputForm.InputVal("Введите заголовок для таблицы")
 
 HTMLHeader$ = Replace("<html><head><meta http-equiv=~Content-Language~ content=~ru~>" + _
               "<meta http-equiv=~Content-Type~ content=~text/html; charset=windows-1251~>", "~", Chr(34))

 HTMLInfo$ = "<title>" + Capt + "</title>"
 
 HTMLStart$ = Replace("</head><body bgcolor=~#D6E0E0~><div align=~center~><table border=~1~ cellspacing=~2~ style=~border-collapse: collapse~>", "~", Chr(34))

 HTMLEnd$ = "</table></div><br><hr><i>Файл сгенерирован программой DB Xtension по содержимому БД </i><b>&#39;" + DBPath + "&#39;</b></body></html>"
 
 HTMLCaption$ = Replace("<tr><td colspan=~" + CStr(DB(DBCurIndex).Header.ColCount) + "~ align=~center~ bgcolor=~3399FF~><font color=~#FFFF00~ size=~5~>" + Capt + "</font></td></tr>", "~", Chr(34))

 HTMLRowS$ = "<tr>"
 HTMLRowE$ = "</tr>"
 
 If (DB(DBCurIndex).Header.ColCount > 0) Then ColWidth% = 100 \ DB(DBCurIndex).Header.ColCount
 
 HTMLCols$ = Replace("<td bgcolor=~FFFFCC~ width=~" + CStr(ColWidth) + "%~ align=~center~><b><font face=~Arial~>^</font></b></td>", "~", Chr(34))
 
 HTMLCells$ = Replace("<td bgcolor=~#FFFFFF~ width=~" + CStr(ColWidth) + "%~ align=~center~>^</td>", "~", Chr(34))

 Put DBI, , HTMLHeader
 Put DBI, , HTMLInfo
 
 If (DB(DBCurIndex).Header.ColCount > 0) Then
    Put DBI, , HTMLStart
    Put DBI, , HTMLCaption
 
    Put DBI, , HTMLRowS
    For c% = 0 To DB(DBCurIndex).Header.ColCount - 1
        Put DBI, , Replace(HTMLCols, "^", CStr(DB(DBCurIndex).Cols(c).title))
    Next c
    Put DBI, , HTMLRowE
 
    For R% = 0 To DB(DBCurIndex).Header.RowCount - 1
        Put DBI, , HTMLRowS
        For c% = 0 To DB(DBCurIndex).Header.ColCount - 1
            tmp$ = CStr(DB(DBCurIndex).Rows(R).Fields(c))
            If (Trim(tmp) = "") Then tmp = "&nbsp;"
            Put DBI, , Replace(HTMLCells, "^", tmp)
        Next c
        Put DBI, , HTMLRowE
    Next R
 
    Put DBI, , HTMLEnd
 Else
    Put DBI, , "</head><body>База не содержит данных</body></html>"
 End If
 
 Close DBI
 
 If (MsgForm.QuestMsg("Файл '" + Path + "' создан. Открыть?") = resOk) Then
    Call ShellExecute(hwnd, "open", Path, "", "", 0)
 End If
End Sub

Private Sub HTMLCreator_Click()
 CoolTimer.Enabled = False
 HTMLPath.FileName = ""
 HTMLPath.ShowSave
 If (HTMLPath.FileName <> "") Then
  Call CreateHTML(HTMLPath.FileName)
 Else
  Call MsgForm.ErrorMsg("Формирование HTML-документа отменено!")
 End If
 CoolTimer.Enabled = True
End Sub

Private Sub ListView_DblClick()
 If (saveItemIndex > 0) Then
  Load EditRecordForm
  With EditRecordForm
   .CellList.Clear
   .ERFDBIndex = DBCurIndex
   Call .LoadData(saveItemIndex - 1)
   Call .OverloadList
   .Show vbModal
  End With
 End If
End Sub

Private Sub ListView_ItemClick(ByVal Item As MSComctlLib.ListItem)
 saveItemIndex = Item.Index
End Sub

Private Sub ListView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 saveItemIndex = 0
End Sub

Private Sub OptDB_Click()
 Security.Enabled = DBPath <> ""
End Sub

Private Sub Form_Load()
' регистрации расширения
 Call ShellExecute(0, "", "assoc.exe", App.Path + "\" + App.EXEName + ".exe", "", 0)
 DBCurIndex = 0
 UserIsAdmin = True
 saveItemIndex = 0
 OldImageIndex = -1
 Call ClearAll
 dW1 = Width - TabStrip.Width
 dH1 = Height - TabStrip.Height
 dW2 = Width - ListView.Width
 dH2 = Height - ListView.Height
 Call DisEnImage(0, 0)
 Call DisEnImage(1, 0)
 Call DisEnImage(2, 1)
 Call DisEnImage(3, 1)
 Call DisEnImage(4, 1)
 Call DisEnImage(5, 0)
End Sub

Private Sub Form_Resize()
 CoolBar1.Width = 2 * Width

 Min% = MainForm.Width - dW2
 If (Min < 0) Then: Min = 0
 ListView.Width = Min
 
 Min = MainForm.Height - dH2
 If (Min < 0) Then: Min = 0
 ListView.Height = Min
 
 Min = MainForm.Width - dW1
 If (Min < 0) Then: Min = 0
 TabStrip.Width = Min
 
 Min = MainForm.Height - dH1
 If (Min < 0) Then: Min = 0
 TabStrip.Height = Min
End Sub

Private Sub Form_Unload(Cancel%)
 If DBChanged Then
  If (MsgForm.QuestMsg("Выйти?") = resNo) Then Cancel = 1
 End If
 Close  ' пожалуй, это лишнее, но да мало ли :)
End Sub

Private Sub OpenDB_Click()
 CoolTimer.Enabled = False
 Dlgs.FileName = ""
 Dlgs.ShowOpen
 If (Dlgs.FileName <> "") Then
  ' открываю БД
  If LoadDB(DBCurIndex, Dlgs.FileName) Then
    ' вывожу путь к БД
    SB.Panels(3).Text = DBPath
    Call DisEnImage(2, 0)
    Call DisEnImage(3, 0)
    Call DisEnImage(4, 0)
    Call ShowTable(DBCurIndex)
  End If
 End If
 CoolTimer.Enabled = True
End Sub

Private Sub QueryDB_Click()
    QueryM.Enabled = DBPath <> ""
End Sub

Private Sub ResDB_Click()
    DiagDraw.Enabled = DBPath <> ""
    HTMLCreator.Enabled = DBPath <> ""
End Sub

Private Sub QueryM_Click()
 CoolTimer.Enabled = False
 With QueryMasterForm
    .QMFDBIndex = DBCurIndex
    .Show vbModal
 End With
 CoolTimer.Enabled = True
End Sub

Private Sub ResCopyDB_Click()
 CoolTimer.Enabled = False
 Dlgs.FileName = ""
 Dlgs.ShowSave
 If (Dlgs.FileName <> "") Then
  If (Dlgs.FileName = DBPath) Then
    Call MsgForm.ErrorMsg("Нельзя копировать файл сам в себя!")
  Else
   Call CopyFile(DBPath, Dlgs.FileName, False)
   Call MsgForm.InfoMsg("Архивная копия БД создана.")
  End If
 Else
  Call MsgForm.ErrorMsg("Резервное копирование БД отменено!")
 End If
 CoolTimer.Enabled = True
End Sub

Private Sub SaveDB_Click()
    CoolTimer.Enabled = False
    Dlgs.FileName = ""
    Dlgs.ShowSave
    If (Dlgs.FileName <> "") Then
        DBPath = Dlgs.FileName
        Call FlushDB(DBCurIndex)
    End If
    CoolTimer.Enabled = True
End Sub

Private Sub Security_Click()
 CoolTimer.Enabled = False
 If UserIsAdmin Then
    With PasswordForm
        .SetPassText = DB(DBCurIndex).Password
        
        If (DB(DBCurIndex).Header.Flags And flCoded) Then
            .CheckCoded = 1
        Else
            .CheckCoded = 0
        End If
        If (DB(DBCurIndex).Header.Flags And flReadOnlyEnable) Then
            .CheckNoRO = 1
        Else
            .CheckNoRO = 0
        End If
        .CaptionLabel = "Настройка защиты"
        .TextLabel = "Вы можете изменить пароль и права доступа к данной БД. Наличие пароля предполагает ограниченный доступ."
        .Frame1.Visible = False
        .Frame2.Visible = True
        .Show vbModal
        If (.res) Then
            If (Trim(.SetPassText) <> "") Then
                DB(DBCurIndex).Password = Trim(.SetPassText)
                DB(DBCurIndex).Header.Flags = flPasswordNeed + DB(DBCurIndex).Header.Flags + (flCoded * .CheckCoded) + (flReadOnlyEnable * .CheckNoRO)
                Call MsgForm.InfoMsg("Был задан пароль!")
            Else
                DB(DBCurIndex).Header.Flags = 0
            End If
        End If
        Unload PasswordForm
    End With
 Else
    Call ProtectedMsg
 End If
 CoolTimer.Enabled = True
End Sub

Private Sub TabStrip_Click()
 If (TabStrip.Tabs.Count = 0) Then Exit Sub
 If (DBCurIndex <> TabStrip.SelectedItem.Index - 1) Then
    DBCurIndex = TabStrip.SelectedItem.Index - 1
    Call ShowTable(DBCurIndex)
End If
End Sub

Private Sub TabStrip_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (Shift = vbCtrlMask) Then PopupMenu TSMenu
End Sub

Private Sub TSClose_Click()
    If (MsgForm.QuestMsg("Закрыть закладку?") = resOk) Then
        TabIndex% = TabStrip.SelectedItem.Index
        TabStrip.Tabs.Remove (TabIndex)
        Call DelTable(TabIndex - 1)
        
        If (TabStrip.Tabs.Count = 0) Then
            DBChanged = False
            Call DisEnImage(2, 1)
            Call DisEnImage(3, 1)
            Call DisEnImage(4, 1)
            Call ShowTable(-1)
        Else
            TabStrip.SelectedItem = TabStrip.Tabs.Item(1)
        End If
    End If
End Sub
