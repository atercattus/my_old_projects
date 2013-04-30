VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form DiagOptForm 
   BackColor       =   &H00DCDCDC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Настройки диаграмм"
   ClientHeight    =   4950
   ClientLeft      =   6285
   ClientTop       =   4410
   ClientWidth     =   7275
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
   ScaleHeight     =   4950
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4155
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   7329
      _Version        =   393216
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BackColor       =   14474460
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Цвета и текст"
      TabPicture(0)   =   "DiagOpt.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label7"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame2(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "ColorDlg"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Размеры"
      TabPicture(1)   =   "DiagOpt.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "UpDown5"
      Tab(1).Control(1)=   "UpDown1"
      Tab(1).Control(2)=   "UpDown2"
      Tab(1).Control(3)=   "UpDown3"
      Tab(1).Control(4)=   "UpDown4"
      Tab(1).Control(5)=   "UpDown6"
      Tab(1).Control(6)=   "UpDown7"
      Tab(1).Control(7)=   "Label13"
      Tab(1).Control(8)=   "Label26"
      Tab(1).Control(9)=   "Label11"
      Tab(1).Control(10)=   "Label25"
      Tab(1).Control(11)=   "Label24"
      Tab(1).Control(12)=   "Label17"
      Tab(1).Control(13)=   "Label16"
      Tab(1).Control(14)=   "Label23"
      Tab(1).Control(15)=   "Label14"
      Tab(1).Control(16)=   "Label22"
      Tab(1).Control(17)=   "Label12"
      Tab(1).Control(18)=   "Label21"
      Tab(1).Control(19)=   "Label20"
      Tab(1).Control(20)=   "Label9"
      Tab(1).ControlCount=   21
      TabCaption(2)   =   "Параметры"
      TabPicture(2)   =   "DiagOpt.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Check3"
      Tab(2).Control(1)=   "Check2"
      Tab(2).Control(2)=   "Check1"
      Tab(2).Control(3)=   "Label3"
      Tab(2).Control(4)=   "Label2"
      Tab(2).Control(5)=   "Label1"
      Tab(2).ControlCount=   6
      Begin ComCtl2.UpDown UpDown5 
         Height          =   225
         Left            =   -68520
         TabIndex        =   34
         Top             =   2460
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   397
         _Version        =   327681
         Value           =   4
         BuddyControl    =   "Label24"
         BuddyDispid     =   196627
         OrigLeft        =   5940
         OrigTop         =   2460
         OrigRight       =   6495
         OrigBottom      =   2685
         Max             =   100
         Min             =   4
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   225
         Left            =   -68520
         TabIndex        =   21
         Top             =   840
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   397
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "Label20"
         BuddyDispid     =   196635
         OrigLeft        =   5940
         OrigTop         =   960
         OrigRight       =   6495
         OrigBottom      =   1185
         Min             =   1
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComDlg.CommonDialog ColorDlg 
         Left            =   480
         Top             =   3360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00DCDCDC&
         Caption         =   "Цвета и надписи элементов"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   3660
         TabIndex        =   16
         Top             =   420
         Width           =   3495
         Begin VB.TextBox Text1 
            Appearance      =   0  'Плоска
            Height          =   345
            Left            =   780
            TabIndex        =   42
            Top             =   3120
            Width           =   2535
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  'Плоска
            BackColor       =   &H00000000&
            BorderStyle     =   0  'Нет
            Caption         =   "Frame2"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   4
            Left            =   180
            TabIndex        =   18
            Top             =   3120
            Width           =   375
         End
         Begin VB.ListBox List1 
            Appearance      =   0  'Плоска
            ForeColor       =   &H00000000&
            Height          =   2670
            IntegralHeight  =   0   'False
            ItemData        =   "DiagOpt.frx":0054
            Left            =   120
            List            =   "DiagOpt.frx":0056
            TabIndex        =   17
            Top             =   300
            Width           =   3255
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Плоска
         BackColor       =   &H00000000&
         BorderStyle     =   0  'Нет
         Caption         =   "Frame2"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   3
         Left            =   2760
         TabIndex        =   14
         Top             =   3600
         Width           =   375
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Плоска
         BackColor       =   &H00000000&
         BorderStyle     =   0  'Нет
         Caption         =   "Frame2"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   2760
         TabIndex        =   12
         Top             =   3060
         Width           =   375
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00DCDCDC&
         Caption         =   "Цвета заливки"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2475
         Left            =   120
         TabIndex        =   7
         Top             =   420
         Width           =   3435
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Плоска
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   2055
            Left            =   120
            ScaleHeight     =   2025
            ScaleWidth      =   405
            TabIndex        =   20
            Top             =   300
            Width           =   435
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  'Плоска
            BackColor       =   &H00000000&
            BorderStyle     =   0  'Нет
            Caption         =   "Frame2"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   1
            Left            =   2640
            TabIndex        =   11
            Top             =   1995
            Width           =   375
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  'Плоска
            BackColor       =   &H00000000&
            BorderStyle     =   0  'Нет
            Caption         =   "Frame2"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   0
            Left            =   2640
            TabIndex        =   10
            Top             =   375
            Width           =   375
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Правая привязка
            Appearance      =   0  'Плоска
            BackColor       =   &H80000005&
            BackStyle       =   0  'Прозрачно
            Caption         =   "Начальный цвет также является цветом заливки фона, если градиентная заливка не используется."
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00606060&
            Height          =   1155
            Left            =   360
            TabIndex        =   19
            Top             =   780
            Width           =   2715
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Правая привязка
            BackStyle       =   0  'Прозрачно
            Caption         =   "Конечный"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   180
            TabIndex        =   9
            Top             =   2040
            Width           =   2325
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Правая привязка
            BackStyle       =   0  'Прозрачно
            Caption         =   "Начальный"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   180
            TabIndex        =   8
            Top             =   420
            Width           =   2325
         End
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Делать выноски в круговой диаграмме"
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
         Height          =   315
         Left            =   -74760
         TabIndex        =   3
         Top             =   1740
         Value           =   1  'Отмечено
         Width           =   6915
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Рисовать сноски с точек на оси"
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
         Height          =   315
         Left            =   -74760
         TabIndex        =   2
         Top             =   2760
         Value           =   1  'Отмечено
         Width           =   6915
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Использовать цветовую заливку"
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
         Height          =   315
         Left            =   -74760
         TabIndex        =   1
         Top             =   480
         Value           =   1  'Отмечено
         Width           =   6915
      End
      Begin ComCtl2.UpDown UpDown2 
         Height          =   225
         Left            =   -68520
         TabIndex        =   24
         Top             =   1200
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   397
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "Label21"
         BuddyDispid     =   196634
         OrigLeft        =   5940
         OrigTop         =   1320
         OrigRight       =   6495
         OrigBottom      =   1545
         Max             =   100
         Min             =   1
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown3 
         Height          =   225
         Left            =   -68520
         TabIndex        =   27
         Top             =   1560
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   397
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "Label22"
         BuddyDispid     =   196632
         OrigLeft        =   5940
         OrigTop         =   1680
         OrigRight       =   6495
         OrigBottom      =   1905
         Max             =   50
         Min             =   1
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown4 
         Height          =   225
         Left            =   -68520
         TabIndex        =   30
         Top             =   1920
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   397
         _Version        =   327681
         Value           =   2
         BuddyControl    =   "Label23"
         BuddyDispid     =   196630
         OrigLeft        =   5940
         OrigTop         =   2040
         OrigRight       =   6495
         OrigBottom      =   2265
         Max             =   10000
         Min             =   1
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown6 
         Height          =   225
         Left            =   -68520
         TabIndex        =   36
         Top             =   2880
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   397
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "Label25"
         BuddyDispid     =   196626
         OrigLeft        =   5940
         OrigTop         =   3180
         OrigRight       =   6495
         OrigBottom      =   3405
         Increment       =   4
         Max             =   999999
         Min             =   1
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown7 
         Height          =   225
         Left            =   -68520
         TabIndex        =   39
         Top             =   3300
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   397
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "Label26"
         BuddyDispid     =   196624
         OrigLeft        =   5940
         OrigTop         =   3180
         OrigRight       =   6495
         OrigBottom      =   3405
         Increment       =   2
         Max             =   100
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Правая привязка
         BackStyle       =   0  'Прозрачно
         Caption         =   "Внутренний радиус круговой диаграммы"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74580
         TabIndex        =   41
         Top             =   3300
         Width           =   5300
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Центровка
         BackStyle       =   0  'Прозрачно
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -69300
         TabIndex        =   40
         Top             =   3300
         Width           =   800
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Правая привязка
         BackStyle       =   0  'Прозрачно
         Caption         =   "Внешний радиус круговой диаграммы"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74580
         TabIndex        =   38
         Top             =   2880
         Width           =   5300
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Центровка
         BackStyle       =   0  'Прозрачно
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -69300
         TabIndex        =   37
         Top             =   2880
         Width           =   800
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Центровка
         BackStyle       =   0  'Прозрачно
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -69300
         TabIndex        =   35
         Top             =   2460
         Width           =   800
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Правая привязка
         BackStyle       =   0  'Прозрачно
         Caption         =   "Коэффициент сжатия круговой диаграммы (% от круга)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74580
         TabIndex        =   33
         Top             =   2460
         Width           =   5300
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Правая привязка
         AutoSize        =   -1  'True
         BackStyle       =   0  'Прозрачно
         Caption         =   "Количество переходов градиента"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74580
         TabIndex        =   32
         Top             =   1920
         Width           =   5300
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Центровка
         BackStyle       =   0  'Прозрачно
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -69300
         TabIndex        =   31
         Top             =   1920
         Width           =   800
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Правая привязка
         AutoSize        =   -1  'True
         BackStyle       =   0  'Прозрачно
         Caption         =   "Размер точек"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74580
         TabIndex        =   29
         Top             =   1560
         Width           =   5300
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Центровка
         BackStyle       =   0  'Прозрачно
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -69300
         TabIndex        =   28
         Top             =   1560
         Width           =   800
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Правая привязка
         AutoSize        =   -1  'True
         BackStyle       =   0  'Прозрачно
         Caption         =   "Глубина объёмных диаграмм"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74580
         TabIndex        =   26
         Top             =   1200
         Width           =   5300
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Центровка
         BackStyle       =   0  'Прозрачно
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -69300
         TabIndex        =   25
         Top             =   1200
         Width           =   800
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Центровка
         BackStyle       =   0  'Прозрачно
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -69300
         TabIndex        =   23
         Top             =   840
         Width           =   800
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Правая привязка
         AutoSize        =   -1  'True
         BackStyle       =   0  'Прозрачно
         Caption         =   "Толщина линий линейной диаграммы"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74580
         TabIndex        =   22
         Top             =   840
         Width           =   5300
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Правая привязка
         BackStyle       =   0  'Прозрачно
         Caption         =   "Цвет линий"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   3630
         Width           =   2535
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Правая привязка
         BackStyle       =   0  'Прозрачно
         Caption         =   "Цвет надписей"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   3105
         Width           =   2535
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Плоска
         BackColor       =   &H80000005&
         BackStyle       =   0  'Прозрачно
         Caption         =   "В диаграммах столбчатого типа (все, кроме круговых) рисуются оси с отметками и линиями сносок."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   795
         Left            =   -72720
         TabIndex        =   6
         Top             =   3120
         Width           =   4815
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Плоска
         BackColor       =   &H80000005&
         BackStyle       =   0  'Прозрачно
         Caption         =   "В круговой строятся выноски, заменяющие легенду."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   555
         Left            =   -72660
         TabIndex        =   5
         Top             =   2100
         Width           =   4815
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Плоска
         BackColor       =   &H80000005&
         BackStyle       =   0  'Прозрачно
         Caption         =   "Фон заливается с плавным переходом цветов. Если режим отключён, фон заливается начальным цветом градиентной заливки."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   795
         Left            =   -72660
         TabIndex        =   4
         Top             =   840
         Width           =   4815
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Label Label15 
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
      Height          =   345
      Left            =   5460
      TabIndex        =   46
      Top             =   4380
      Width           =   1470
   End
   Begin VB.Label Label10 
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
      Height          =   345
      Left            =   360
      TabIndex        =   45
      Top             =   4380
      Width           =   1470
   End
   Begin VB.Image SelectImg 
      Height          =   480
      Left            =   300
      MousePointer    =   10  'Стрелка вверх
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Image CancelImg 
      Height          =   480
      Left            =   5400
      MousePointer    =   10  'Стрелка вверх
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label SelectBut 
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   300
      MousePointer    =   1  'Указатель
      TabIndex        =   44
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label CancelBut 
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      BackStyle       =   0  'Прозрачно
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5400
      MousePointer    =   1  'Указатель
      TabIndex        =   43
      Top             =   4320
      Width           =   1575
   End
End
Attribute VB_Name = "DiagOptForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public res%

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
    Case 27
        Label15_Click
    Case 13
        Label10_Click
 End Select
End Sub

Private Sub Form_Load()
    res = 0
    Call ButEnabled(SelectImg, SelectBut, True)
    Call ButEnabled(CancelImg, CancelBut, True)
End Sub

Private Sub Form_Paint()
    Call DiagResForm.ColorFill(Picture1, Frame2(0).BackColor, Frame2(1).BackColor)
End Sub

Private Sub Frame2_Click(Index As Integer)
 ColorDlg.Color = Frame2(Index).BackColor
 ColorDlg.ShowColor
 Frame2(Index).BackColor = ColorDlg.Color
 If (Index < 2) Then Call DiagResForm.ColorFill(Picture1, Frame2(0).BackColor, Frame2(1).BackColor)
 If (Index = 4) Then List1.ItemData(List1.ListIndex) = Frame2(4).BackColor
End Sub

Private Sub Label10_Click()
    res = 1
    Hide
End Sub

Private Sub Label15_Click()
    Hide
End Sub

Private Sub List1_Click()
    If (List1.ListIndex > -1) Then
        Text1.Text = List1.List(List1.ListIndex)
        Frame2(4).BackColor = List1.ItemData(List1.ListIndex)
    End If
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
    Call List1_Click
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 13) Then
        List1.List(List1.ListIndex) = Text1.Text
        List1.ItemData(List1.ListIndex) = Frame2(4).BackColor
    End If
End Sub
