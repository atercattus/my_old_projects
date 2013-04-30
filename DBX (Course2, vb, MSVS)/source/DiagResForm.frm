VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form DiagResForm 
   Appearance      =   0  'Плоска
   AutoRedraw      =   -1  'True
   BackColor       =   &H00DCDCDC&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Результирующая диаграмма"
   ClientHeight    =   6735
   ClientLeft      =   4905
   ClientTop       =   3450
   ClientWidth     =   10185
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   10185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.VScrollBar VScroll 
      Height          =   6135
      Left            =   9840
      Max             =   100
      Min             =   4
      TabIndex        =   5
      Top             =   540
      Value           =   40
      Width           =   315
   End
   Begin VB.PictureBox Chart 
      Appearance      =   0  'Плоска
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   6090
      Left            =   0
      ScaleHeight     =   6060
      ScaleWidth      =   9735
      TabIndex        =   2
      Top             =   585
      Width           =   9765
      Begin MSComDlg.CommonDialog CD 
         Left            =   780
         Top             =   1260
         _ExtentX        =   820
         _ExtentY        =   820
         _Version        =   393216
         DefaultExt      =   "*.bmp"
         DialogTitle     =   "Укажите имя создаваемого bmp-файла"
         Filter          =   "BMP|*.bmp"
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Привязать вверх
      Appearance      =   0  'Плоска
      BackColor       =   &H00DCDCDC&
      BorderStyle     =   0  'Нет
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   10185
      TabIndex        =   0
      Top             =   0
      Width           =   10185
      Begin VB.Frame Frame1 
         Appearance      =   0  'Плоска
         BackColor       =   &H00DCDCDC&
         BorderStyle     =   0  'Нет
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   56
         TabIndex        =   1
         Top             =   56
         Width           =   6832
         Begin VB.Image Image2 
            Height          =   240
            Left            =   420
            MousePointer    =   1  'Указатель
            Picture         =   "DiagResForm.frx":0000
            ToolTipText     =   "Настройки"
            Top             =   75
            Width           =   240
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Плоска
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Прозрачно
            Caption         =   "0"
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
            Height          =   225
            Left            =   4620
            MousePointer    =   1  'Указатель
            TabIndex        =   4
            Top             =   90
            Width           =   90
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Правая привязка
            Appearance      =   0  'Плоска
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Прозрачно
            Caption         =   "Значение функции под курсором:"
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
            Height          =   225
            Left            =   1680
            MousePointer    =   1  'Указатель
            TabIndex        =   3
            Top             =   90
            Width           =   2865
         End
         Begin VB.Image Image3 
            Height          =   240
            Left            =   840
            MousePointer    =   1  'Указатель
            Picture         =   "DiagResForm.frx":058A
            ToolTipText     =   "Вернуться назад"
            Top             =   75
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Left            =   60
            MousePointer    =   1  'Указатель
            Picture         =   "DiagResForm.frx":0B14
            ToolTipText     =   "Сохранить диаграмму"
            Top             =   70
            Width           =   240
         End
      End
   End
End
Attribute VB_Name = "DiagResForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dW%, dH%, dX%, dH2%
Dim DiagData() As TDiagElem
Dim DrawingMode As Byte, Use3D As Boolean

' константы для вывода куска более 270 градусов (выводимая часть)
Const mode270begin As Byte = 1
Const mode270end As Byte = 2

' данные для процедур рисования
    Const Pi_180 As Double = 1.74532925199433E-02
    Const Pi_2 As Double = 1.5707963267949
    Const NearZero As Double = 1E-45

    Dim Xc%, Yc%            '   центр диаграммы
    Dim Radius#             '   радиус кусков
    Dim InRad#              '   радиус разноса кусков
    Dim OneGradus#          '   единиц в одном градусе
    Dim ChartHeight%        '   высота графика
    Dim ChartWidth%         '   ширина графика
    Dim ChartTop%           '   верх графика
    Dim ChartDown%          '   низ графика
    Dim ItemCount%          '   кол-во элементов
    Dim Max As Long
    Dim Sum As Long         '   максимальное значение и сумма всех значений
    Dim OldGrad#            '   предыдущий угол
    Dim LineCount As Long   '   количество полос заливки
    Dim d3D%                '   смещение в 3D, в пикселях
    Dim dWidth As Single    '   ширина одного столбца
    Dim dHeight As Single   '   высота 'единицы высоты'
    Dim StartFillColor As Long
    Dim EndFillColor As Long
    Dim LineColor As Long
    Dim LineWidth As Byte
    Dim PointRadius%
    Dim Ellipce#
    Dim UseColorFill As Boolean
    Dim UseCircleLegend As Boolean
    Dim UseLineLeftValues As Boolean

Public Sub InitDiagData(Data(), ByVal Mode As Byte, ByVal May3D As Boolean)
    ReDim DiagData(UBound(Data) \ 2 - 1)
    d# = 255 / (UBound(Data) \ 2 - 1)
    For i% = 0 To (UBound(Data) \ 2 - 1)
        DiagData(i).Val = Abs(Data(2 * i))
        DiagData(i).Text = Data(2 * i + 1)
        DiagData(i).Color = RGB(i * d, i * d, i * d)
    Next i
    DrawingMode = Mode
    Use3D = May3D
    
    Label2.Visible = (DrawingMode <> 3)
    Label3.Visible = Label2.Visible
    VScroll.Enabled = Not Label2.Visible
End Sub

Public Sub ColorFill(PB As PictureBox, ByVal StColor As Long, ByVal EnColor As Long)
    Dim dR#, dG#, DB#, dC1 As Long, dC2 As Long
    Dim R#, G#, B#
    Dim intLoop As Long
    
    PB.Line (0, 0)-(PB.Width, PB.Height), EnColor, BF

    ' get Red
    dC1 = StColor - (StColor \ &H100) * &H100
    R = dC1
    dC2 = EnColor - (EnColor \ &H100) * &H100
    dR = (dC1 - dC2) / LineCount
    
    ' get Green
    dC1 = (StColor - (StColor \ &H10000) * &H10000 - dC1) \ &H100
    G = dC1
    dC2 = (EnColor - (EnColor \ &H10000) * &H10000 - dC2) \ &H100
    dG = (dC1 - dC2) / LineCount
    
    ' get Blue
    dC1 = StColor \ &H10000
    B = dC1
    dC2 = EnColor \ &H10000
    DB = (dC1 - dC2) / LineCount

    With PB
        .DrawStyle = 1
        .DrawMode = vbCopyPen
        .ScaleMode = vbPixels
        .DrawWidth = 2
        .ScaleHeight = LineCount
        For intLoop = 0 To LineCount - 1
            PB.Line (0, intLoop)-(PB.Width, intLoop - 1), RGB(R, G, B), BF
            R = R - dR: If (R < 0) Then R = 255: If (R > 255) Then R = 0
            G = G - dG: If (G < 0) Then G = 255: If (G > 255) Then G = 0
            B = B - DB: If (B < 0) Then B = 255: If (B > 255) Then B = 0
        Next intLoop
        .ScaleMode = vbTwips
        .DrawWidth = 1
    End With
End Sub

Sub OutOneElem(ElemIndex As Integer, StAn#, EnAn#, Optional Mode270Mode As Byte = 0)
    ' центральный угол
    angle# = (StAn + (EnAn - StAn) / 2) * Pi_180
    
    ' динамическая глубина
    d3D_% = Round(d3D / 100 * (100 - Round(100 * Ellipce)))
    If (d3D_ = 0) Then d3D_ = 1
    ' динамическое смещение центров кусков
    r_# = Ellipce * d3D / 100
    
    X1# = Xc + Radius * Cos(angle)
    Y1# = Yc - Radius * Sin(angle)
    
    x# = Xc + InRad / Radius * (X1 - Xc)
    y# = Yc + InRad / Radius * (Y1 - Yc) * r_
    
    If (Not Use3D) Then
        Chart.FillStyle = 0
        Chart.FillColor = DiagData(ElemIndex).Color
        If (StAn <> 0) Then
            Chart.Circle (x, y), Radius, LineColor, -StAn * Pi_180, -EnAn * Pi_180, Ellipce
        Else
            Chart.Circle (x, y), Radius, LineColor, -1E-45, -EnAn * Pi_180, Ellipce
        End If
        Chart.FillStyle = 1
        
        ' вывод значений
        R# = 1.3 * Radius
        X2# = x + R * Cos(angle)
        Y2# = y - Ellipce * R * Sin(angle)
        
        x0# = x + Radius * Cos(angle)
        y0# = y - Ellipce * Radius * Sin(angle)
        
        str_1$ = CStr(DiagData(ElemIndex).Text)
        d1# = Chart.TextWidth(str_1)
        str_2$ = CStr(DiagData(ElemIndex).Val)
        d2# = Chart.TextWidth(str_2)
        
        If UseCircleLegend Then
            Chart.DrawStyle = 4
            Chart.Line (x0, y0)-(X2, Y2), LineColor
            Chart.DrawStyle = 0
            
            If Not ((angle > Pi_2) And (angle <= 3 * Pi_2)) Then
                Chart.Line (X2, Y2)-(X2 + d1, Y2), LineColor
                Chart.CurrentX = X2
                Chart.CurrentY = Y2
                Chart.Print CStr(str_1)
                
                Chart.CurrentX = X2
                Chart.CurrentY = Y2 - Chart.TextHeight(str_2)
                Chart.Print CStr(str_2)
            Else
                Chart.Line (X2, Y2)-(X2 - d1, Y2), LineColor
                Chart.CurrentX = X2 - d1
                Chart.CurrentY = Y2
                Chart.Print CStr(str_1)
                
                Chart.CurrentX = X2 - d1
                Chart.CurrentY = Y2 - Chart.TextHeight(str_2)
                Chart.Print CStr(str_2)
            End If
        End If
        
    Else
        Chart.FillStyle = 0
        Chart.FillColor = DiagData(ElemIndex).Color
        
        Select Case Mode270Mode
            Case 0
                sa# = StAn
                If (sa = 0) Then sa = 1E-45 Else sa = sa * Pi_180
                For i% = d3D_ To 1 Step -1
                    If (i = d3D_) Then
                        Chart.DrawStyle = vbSolid
                        Chart.Circle (x, Screen.TwipsPerPixelY * (i - 1) + y), Radius, LineColor, -sa, -EnAn * Pi_180, Ellipce
                        Chart.DrawStyle = vbInvisible
                    ElseIf (i = 1) Then
                        Chart.DrawStyle = vbSolid
                        Chart.Circle (x, y), Radius, LineColor, -sa, -EnAn * Pi_180, Ellipce
                        Chart.DrawStyle = vbInvisible
                    Else
                        Chart.Circle (x, Screen.TwipsPerPixelY * (i - 1) + y), Radius, LineColor, -sa, -EnAn * Pi_180, Ellipce
                    End If
                Next i
                
            Case mode270begin
                For i% = d3D_ To 1 Step -1
                    If (i = d3D_) Then
                        Chart.DrawStyle = vbSolid
                        Chart.Circle (x, Screen.TwipsPerPixelY * (i - 1) + y), Radius, LineColor, -StAn * Pi_180, -EnAn * Pi_180, Ellipce
                        Chart.DrawStyle = vbInvisible
                    Else
                        Chart.Circle (x, Screen.TwipsPerPixelY * (i - 1) + y), Radius, LineColor, -StAn * Pi_180, -angle, Ellipce
                    End If
                Next i
                
            Case mode270end
                For i% = d3D_ To 1 Step -1
                    If (i = 1) Then
                        Chart.DrawStyle = vbSolid
                        Chart.Circle (x, y), Radius, LineColor, -StAn * Pi_180, -EnAn * Pi_180, Ellipce
                    Else
                        Chart.DrawStyle = vbInvisible
                        Chart.Circle (x, Screen.TwipsPerPixelY * (i - 1) + y), Radius, LineColor, -angle, -EnAn * Pi_180, Ellipce
                    End If
                Next i
        End Select
        
        Chart.FillStyle = 1
        Chart.DrawStyle = vbSolid
        
        ' вывод значений
        R# = 1.3 * Radius
        X2# = x + R * Cos(angle)
        Y2# = y - Ellipce * R * Sin(angle)
        
        x0# = x + Radius * Cos(angle)
        y0# = y - Ellipce * Radius * Sin(angle)
        
        str_1$ = CStr(DiagData(ElemIndex).Text)
        d1# = Chart.TextWidth(str_1)
        str_2$ = CStr(DiagData(ElemIndex).Val)
        d2# = Chart.TextWidth(str_2)
        
        If UseCircleLegend Then
            Chart.DrawStyle = 4
            Chart.Line (x0, y0)-(X2, Y2), LineColor
            Chart.DrawStyle = 0
        
            If Not ((angle > Pi_2) And (angle <= 3 * Pi_2)) Then
                Chart.Line (X2, Y2)-(X2 + d1, Y2), LineColor
                Chart.CurrentX = X2
                Chart.CurrentY = Y2
                Chart.Print CStr(str_1)
                
                Chart.CurrentX = X2
                Chart.CurrentY = Y2 - Chart.TextHeight(str_2)
                Chart.Print CStr(str_2)
            Else
                Chart.Line (X2, Y2)-(X2 - d1, Y2), LineColor
                Chart.CurrentX = X2 - d1
                Chart.CurrentY = Y2
                Chart.Print CStr(str_1)
                
                Chart.CurrentX = X2 - d1
                Chart.CurrentY = Y2 - Chart.TextHeight(str_2)
                Chart.Print CStr(str_2)
            End If
        End If
        
        ' а теперь вывод боковых линий
        Chart.DrawStyle = 0

            ' начальный угол
            If Not ((StAn > 90) And (StAn < 180)) Then
                sa# = StAn * Pi_180
                x0 = x + Radius * Cos(sa)
                y0 = y - Radius * Ellipce * Sin(sa)

                If (Mode270Mode <> mode270end) Then
                    Chart.Line (x0, y0)-(x0, y0 + d3D_ * Screen.TwipsPerPixelY), LineColor
                End If
            End If

            ' конечный угол
            If Not ((EnAn > 0) And (EnAn < 90)) Then
                x0 = x + Radius * Cos(EnAn * Pi_180)
                y0 = y - Radius * Ellipce * Sin(EnAn * Pi_180)

                Chart.Line (x0, y0)-(x0, y0 + d3D_ * Screen.TwipsPerPixelY), LineColor
            End If
            
            ' центр
            If Not ((EnAn >= 270) And (StAn <= 270)) Then
                Chart.Line (x, y)-(x, y + d3D_ * Screen.TwipsPerPixelY), LineColor
            End If
            
            ' левый край
            If ((StAn <= 180) And (EnAn >= 180)) Then
                Chart.Line (x - Radius, y)-(x - Radius, y + d3D_ * Screen.TwipsPerPixelY), LineColor
            End If
            
    End If
    
    OldGrad = Grad
End Sub


' рисование круговой диаграммы
Sub DrawCircle()
    Dim Mode270 As Boolean
    Dim Item270%

    ItemCount = UBound(DiagData) + 1
    
    With Chart
        Max = -1
        Sum = 0
        For i% = 1 To ItemCount
            If (DiagData(i - 1).Val > Max) Then Max = DiagData(i - 1).Val
            Sum = Sum + DiagData(i - 1).Val
        Next i
        
        Mode270 = (Max > 3 / 4 * Sum)
        
        OneGradus = 360 / Sum
        OldGrad = 0.00001
        
        Xc = Chart.Width \ 2
        Yc = Chart.Height \ 2
        
        Dim pos90%, pos270% '   индексы ключевых элементов
        pos90 = -1
        pos270 = -1
        OldGrad = 0
        
        Dim Angles() As Double
        ReDim Angles(ItemCount - 1, 1)
        
        For i% = 1 To ItemCount
            If Mode270 Then If (DiagData(i - 1).Val = Max) Then Item270 = i - 1
            Grad# = DiagData(i - 1).Val * OneGradus + OldGrad
            If (OldGrad <= 90) And (Grad >= 90) Then pos90 = i - 1
            If (OldGrad <= 270) And (Grad >= 270) Then pos270 = i - 1
            Angles(i - 1, 0) = OldGrad
            Angles(i - 1, 1) = Grad
            OldGrad = Grad
        Next i
        
        Chart.DrawStyle = 0
        
        If Not Mode270 Then
            
            For i% = pos90 To 0 Step -1
                Call OutOneElem(i, Angles(i, 0), Angles(i, 1))
            Next i
            
            For i% = pos90 + 1 To pos270 - 1
                Call OutOneElem(i, Angles(i, 0), Angles(i, 1))
            Next i
            
            For i% = ItemCount - 1 To pos270 Step -1
                Call OutOneElem(i, Angles(i, 0), Angles(i, 1))
            Next i
        Else
            
            i% = pos90 - 1
            If (i < 0) Then i = ItemCount - 1
            
            Call OutOneElem(Item270, Angles(Item270, 0), Angles(Item270, 1), mode270begin)
            
            Do While (i <> Item270)
                Call OutOneElem(i, Angles(i, 0), Angles(i, 1))
                
                i = i - 1
                If (i < 0) Then i = ItemCount - 1
            Loop
            
            Call OutOneElem(Item270, Angles(Item270, 0), Angles(Item270, 1), mode270end)
            
        End If
    End With
End Sub

' рисование линейной, точечной и столбчатой диаграмм
Sub DrawPoint()
    Dim d3DX%
    Dim d3DY%
    Dim OldX%, OldY%        '   координаты предыдущей точки
    
    ItemCount = UBound(DiagData) + 1
    ChartHeight = Chart.Height * 0.8
    ChartTop = Chart.Height * 0.1
    ChartDown = Chart.Height * 0.9
    
    With Chart
        dWidth = Chart.Width / (2 * ItemCount + 1)
        
        Max = -1
        Sum = 0
        For i% = 1 To ItemCount
            If (DiagData(i - 1).Val > Max) Then Max = DiagData(i - 1).Val
            Sum = Sum + DiagData(i - 1).Val
        Next i
        
        dHeight = ChartHeight / Max
        
        d3DX = Screen.TwipsPerPixelX
        d3DY = Screen.TwipsPerPixelY
        
        With Chart
            .DrawWidth = 1
            .DrawStyle = 3
            Chart.Line (dWidth * 0.9, ChartTop \ 2)-(dWidth * 0.9, ChartDown), LineColor
            Chart.Line (dWidth * 0.9, ChartDown)-((2 * ItemCount + 0.5) * dWidth, ChartDown), LineColor
            .DrawStyle = 0

            .FontSize = .FontSize + 3
            .FontUnderline = True

            .CurrentX = 2 * d3DX
            .CurrentY = 2 * d3DY
            Chart.Print "Значения"
            
            str_$ = "Подписи"
            .CurrentX = .Width - .TextWidth(str_) - 10 * d3DX
            .CurrentY = ChartDown + .TextHeight(str_)
            Chart.Print str_

            .FontSize = .FontSize - 3
            .FontUnderline = False
        End With


        For i% = 1 To ItemCount
            j% = 2 * i - 1
            Dim y#, x#
            y = ChartTop + dHeight * (Max - DiagData(i - 1).Val)
            
           Select Case DrawingMode
                Case 0  '   ///////////////////////////////// ЛИНИИ /////////////////////////////////////////
                        x# = (j + 0.5) * dWidth
                        
                        If (i > 1) Then
                            Chart.DrawWidth = LineWidth
                            Chart.Line (OldX, OldY)-(x, y), DiagData(i - 1).Color
                            Chart.DrawWidth = 1
                        End If
                        Chart.DrawStyle = 1
                        Chart.Line (x, y)-(x, ChartDown), DiagData(i - 1).Color
                        Chart.DrawStyle = 0
                        OldX = x
                        OldY = y
                        
                        str_$ = CStr(DiagData(i - 1).Text)
                        Chart.CurrentX = j * dWidth + (dWidth - Chart.TextWidth(str_)) \ 2
                        Chart.CurrentY = ChartDown + Chart.TextHeight(str_) \ 10
                        Chart.Print str_
                        
                        str_ = CStr(Round(DiagData(i - 1).Val / Sum * 100)) + "%"
                        Chart.CurrentX = j * dWidth + (dWidth - Chart.TextWidth(str_)) \ 2
                        Chart.CurrentY = y - Chart.TextHeight(str_) * 1.2
                        Chart.Print str_
                        
                        ' значение слева с засечкой и линией
                        str_ = CStr(DiagData(i - 1).Val)
                        If UseLineLeftValues Then
                            Chart.CurrentX = dWidth * 0.8 - Chart.TextWidth(str_)
                            Chart.DrawStyle = 2
                            Chart.Line (dWidth * 0.9, y)-(x, y), LineColor
                            Chart.DrawStyle = 0
                        End If

                        Chart.DrawWidth = 2
                        Chart.Line (dWidth * 0.85, y)-(dWidth * 0.95, y), LineColor
                        Chart.DrawWidth = 1
                        x# = dWidth * 0.8 - Chart.TextWidth(str_)
                        Chart.CurrentX = x
                        Chart.CurrentY = y - Chart.TextHeight(str_) \ 2
                        Chart.Print str_
                        
                Case 1  '   ///////////////////////////////// КОЛОНКИ ///////////////////////////////////////
                        If (Not Use3D) Then
                            Chart.Line (j * dWidth, y)-((j + 1) * dWidth, ChartDown), DiagData(i - 1).Color, BF
                            Chart.Line (j * dWidth, y)-((j + 1) * dWidth, ChartDown), LineColor, B
                            
                            str_ = CStr(DiagData(i - 1).Text)
                            Chart.CurrentX = j * dWidth + (dWidth - Chart.TextWidth(str_)) \ 2
                            Chart.CurrentY = ChartDown + Chart.TextHeight(str_) \ 10
                            Chart.Print str_
                            
                            str_ = CStr(Round(DiagData(i - 1).Val / Sum * 100)) + "%"
                            Chart.CurrentX = j * dWidth + (dWidth - Chart.TextWidth(str_)) \ 2
                            Chart.CurrentY = y - Chart.TextHeight(str_) * 1.2
                            Chart.Print str_
                            
                            ' значение слева с засечкой и линией
                            str_ = CStr(DiagData(i - 1).Val)
                            If UseLineLeftValues Then
                                Chart.CurrentX = dWidth * 0.8 - Chart.TextWidth(str_)
                                Chart.DrawStyle = 2
                                Chart.Line (dWidth * 0.9, y)-(j * dWidth, y), LineColor
                                Chart.DrawStyle = 0
                            End If
                            
                            x# = dWidth * 0.8 - Chart.TextWidth(str_)
                            Chart.CurrentX = x
                            Chart.CurrentY = y - Chart.TextHeight(str_) \ 2
                            Chart.Print str_
                            Chart.CurrentX = x
                            Chart.CurrentY = y
                            Chart.DrawWidth = 2
                            Chart.Line (dWidth * 0.85, y)-(dWidth * 0.95, y), LineColor
                            Chart.DrawWidth = 1
                        Else
                            For k% = 0 To d3D - 1
                                Chart.Line (j * dWidth + k * d3DX, y - k * d3DY)-((j + 1) * dWidth + k * d3DX, ChartDown - k * d3DY), DiagData(i - 1).Color, B
                            Next k
                            Chart.Line (j * dWidth, y)-((j + 1) * dWidth, ChartDown), DiagData(i - 1).Color, BF
                            ' верхняя левая в глубине
                            ltdx% = j * dWidth + (d3D - 1) * d3DX
                            ltdy% = y - (d3D - 1) * d3DY
                            ' верхняя правая в глубине
                            rtdx% = (j + 1) * dWidth + (d3D - 1) * d3DX
                            rtdy% = y - (d3D - 1) * d3DY
                            ' нижняя правая в глубине
                            rddx% = (j + 1) * dWidth + (d3D - 1) * d3DX
                            rddy% = ChartDown - (d3D - 1) * d3DY
                            ' верхняя в глубине
                            Chart.Line (rtdx, rtdy)-(rddx, rddy), LineColor
                            ' правая в глубине
                            Chart.Line (ltdx, ltdy)-(rtdx, rtdy), LineColor
                            
                            ' левая переходная
                            Chart.Line (ltdx, ltdy)-(ltdx - d3D * d3DX, ltdy + d3D * d3DY), LineColor
                            ' правая верхняя переходная
                            Chart.Line (rtdx, rtdy)-(rtdx - d3D * d3DX, rtdy + d3D * d3DY), LineColor
                            ' правая нижняя переходная
                            Chart.Line (rddx, rddy)-(rddx - d3D * d3DX, rddy + d3D * d3DY), LineColor
                            Chart.Line (j * dWidth, y)-((j + 1) * dWidth, ChartDown), LineColor, B
                            
                            ' надпись внизу
                            str_ = CStr(DiagData(i - 1).Text)
                            Chart.CurrentX = j * dWidth + (dWidth - Chart.TextWidth(str_)) \ 2
                            Chart.CurrentY = ChartDown + Chart.TextHeight(str_) \ 10
                            Chart.Print str_
                            ' процент вверху
                            str_ = CStr(Round(DiagData(i - 1).Val / Sum * 100)) + "%"
                            Chart.CurrentX = d3D * d3DX + j * dWidth + (dWidth - Chart.TextWidth(str_)) \ 2
                            Chart.CurrentY = y - d3D * d3DY - Chart.TextHeight(str_) * 1.2
                            Chart.Print str_
                            ' значение слева с засечкой и линией
                            str_ = CStr(DiagData(i - 1).Val)
                            If UseLineLeftValues Then
                                Chart.CurrentX = dWidth * 0.8 - Chart.TextWidth(str_)
                                Chart.DrawStyle = 2
                                Chart.Line (dWidth * 0.9, y)-(j * dWidth, y), LineColor
                                Chart.DrawStyle = 0
                            End If
                            
                            x# = dWidth * 0.8 - Chart.TextWidth(str_)
                            Chart.CurrentX = x
                            Chart.CurrentY = y - Chart.TextHeight(str_) \ 2
                            Chart.Print str_
                            Chart.CurrentX = x
                            Chart.CurrentY = y
                            Chart.DrawWidth = 2
                            Chart.Line (dWidth * 0.85, y)-(dWidth * 0.95, y), LineColor
                            Chart.DrawWidth = 1
                        End If
                        
                Case 2  '   ///////////////////////////////// ТОЧКИ /////////////////////////////////////////
                        Chart.FillStyle = 0
                        Chart.FillColor = DiagData(i - 1).Color
                        x# = (j + 0.5) * dWidth
                        Chart.Circle (x, y), PointRadius * d3DX, LineColor
                        Chart.FillStyle = 1
                        Chart.DrawStyle = 1
                        Chart.Line (x, y)-(x, ChartDown), DiagData(i - 1).Color
                        Chart.DrawStyle = 0
                        
                        str_ = CStr(DiagData(i - 1).Text)
                        Chart.CurrentX = j * dWidth + (dWidth - Chart.TextWidth(str_)) \ 2
                        Chart.CurrentY = ChartDown + Chart.TextHeight(str_) \ 10
                        Chart.Print str_
                        
                        str_ = CStr(Round(DiagData(i - 1).Val / Sum * 100)) + "%"
                        Chart.CurrentX = j * dWidth + (dWidth - Chart.TextWidth(str_)) \ 2
                        Chart.CurrentY = y - PointRadius * d3D - Chart.TextHeight(str_) * 1.2
                        Chart.Print str_
                        
                        ' значение слева с засечкой и линией
                        str_ = CStr(DiagData(i - 1).Val)
                        Chart.CurrentX = dWidth * 0.8 - Chart.TextWidth(str_)
                        Chart.DrawStyle = 2
                        Chart.Line (dWidth * 0.9, y)-(x, y), LineColor
                        Chart.DrawStyle = 0
                        
                        Chart.DrawWidth = 2
                        Chart.Line (dWidth * 0.85, y)-(dWidth * 0.95, y), LineColor
                        Chart.DrawWidth = 1
                        x# = dWidth * 0.8 - Chart.TextWidth(str_)
                        Chart.CurrentX = x
                        Chart.CurrentY = y - Chart.TextHeight(str_) \ 2
                        Chart.Print str_
            End Select
        Next i
        
    End With
End Sub

Sub DrawDiagram()
    If (Chart.Height > Screen.TwipsPerPixelX * 5) And (UseColorFill) Then
        Call ColorFill(Chart, StartFillColor, EndFillColor)
    Else
        Chart.Line (0, 0)-(Chart.Width, Chart.Height), StartFillColor, BF
    End If

    Select Case DrawingMode
        Case 3:         Call DrawCircle
        Case Else:      Call DrawPoint
    End Select
End Sub

Private Sub Chart_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (DrawingMode <> 3) Then
        y = Round((ChartDown - y) * Max / (ChartDown - ChartTop))
        Label3.Caption = CStr(y)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyF5) Then Call DrawDiagram
    If (KeyCode = 27) Then Image3_Click
End Sub

Private Sub Form_Load()
 dW = Width - Chart.Width
 dH = Height - Chart.Height
 dX = Width - VScroll.Left
 dH2 = Height - VScroll.Height
 DrawingMode = 0
 Use3D = False
 LineCount = 100
 d3D = 15
 StartFillColor = RGB(255, 255, 128)
 EndFillColor = RGB(0, 128, 255)
 LineColor = 0
 LineWidth = 1
 Ellipce = 2 / 5
 PointRadius = 15
 
 UseColorFill = True
 UseCircleLegend = True
 UseLineLeftValues = True
 
 ChartHeight = Chart.Height * 0.85
 ChartWidth = Chart.Width * 0.85
 ChartTop = Chart.Height * 0.075
 ChartDown = Chart.Height * 0.925
 If (ChartWidth < ChartHeight) Then Radius = ChartWidth Else Radius = ChartHeight
 Radius = Radius * 0.5
 InRad = 0.1 * Radius
End Sub

Private Sub Form_Resize()
 Min% = Width - dW + 5 * Screen.TwipsPerPixelX
 If (Min < 0) Then Min = 0
 Chart.Width = Min
 
 Min% = Height - dH + Screen.TwipsPerPixelY
 If (Min < 0) Then Min = 0
 Chart.Height = Min
 
 VScroll.Left = Width - dX
 
 Min% = Height - dH2 + Screen.TwipsPerPixelY
 If (Min < 0) Then Min = 0
 VScroll.Height = Min
 
 Call DrawDiagram
End Sub

Private Sub Image1_Click()
    CD.FileName = ""
    CD.ShowSave
    If (CD.FileName <> "") Then
        Call SavePicture(Chart.Image, CD.FileName)
    End If
End Sub

Private Sub Image2_Click()
    With DiagOptForm
        ' цвета
        .Frame2(0).BackColor = StartFillColor
        .Frame2(1).BackColor = EndFillColor
        .Frame2(2).BackColor = Chart.ForeColor
        .Frame2(3).BackColor = LineColor
        ' размеры
        .UpDown1.value = LineWidth
        .UpDown2.value = d3D
        .UpDown3.value = PointRadius
        .UpDown4.value = LineCount
        .UpDown5.value = Round(Ellipce * 100)
        
        .UpDown6.Max = Chart.Width
        If (Chart.Height < Chart.Width) Then .UpDown6.Max = Chart.Width
        .UpDown6.Max = Round(.UpDown6.Max / Screen.TwipsPerPixelX)
        .UpDown6.value = Round(Radius / Screen.TwipsPerPixelX)

        .UpDown7.Max = .UpDown6.Max * 0.9
        .UpDown7.value = Round(InRad / Screen.TwipsPerPixelX)
        
        ' цвета и надписи
        .List1.Clear
        For i% = 1 To ItemCount
            .List1.AddItem (DiagData(i - 1).Text)
            .List1.ItemData(i - 1) = DiagData(i - 1).Color
        Next i
        If (.List1.ListCount > 0) Then .List1.ListIndex = 0
        
        ' флаги
        .Check1.value = -CInt(UseColorFill)
        .Check3.value = -CInt(UseCircleLegend)
        .Check2.value = -CInt(UseLineLeftValues)
        
        .Show vbModal
        If (.res = 1) Then
            ' цвета
            StartFillColor = .Frame2(0).BackColor
            EndFillColor = .Frame2(1).BackColor
            Chart.ForeColor = .Frame2(2).BackColor
            LineColor = .Frame2(3).BackColor
            ' размеры
            LineWidth = .UpDown1.value
            d3D = .UpDown2.value
            PointRadius = .UpDown3.value
            LineCount = .UpDown4.value
            Ellipce = .UpDown5.value / 100
            Radius = .UpDown6.value * Screen.TwipsPerPixelX
            InRad = .UpDown7.value * Screen.TwipsPerPixelX
            ' цвета и надписи
            For i% = 1 To ItemCount
                DiagData(i - 1).Text = .List1.List(i - 1)
                DiagData(i - 1).Color = .List1.ItemData(i - 1)
            Next i
            ' флаги
            UseColorFill = (.Check1.value = 1)
            UseCircleLegend = (.Check3.value = 1)
            UseLineLeftValues = (.Check2.value = 1)
            Call DrawDiagram
        End If
    End With
End Sub

Private Sub Image3_Click()
 Hide
End Sub

Private Sub VScroll_Change()
    Ellipce = VScroll.value / 100
    Call DrawDiagram
End Sub
