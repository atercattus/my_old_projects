VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form TextEditForm 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Редактор текстовых полей"
   ClientHeight    =   4455
   ClientLeft      =   5265
   ClientTop       =   3420
   ClientWidth     =   5295
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog FontDlg 
      Left            =   2820
      Top             =   1500
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      CancelError     =   -1  'True
      FontSize        =   10
   End
   Begin RichTextLib.RichTextBox TextEdit 
      Height          =   4032
      Left            =   60
      TabIndex        =   1
      Top             =   360
      Width           =   5172
      _ExtentX        =   9128
      _ExtentY        =   7117
      _Version        =   393217
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"TextEditForm.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Привязать вверх
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   1
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ClearText"
            Object.ToolTipText     =   "Очистить"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SaveText"
            Object.ToolTipText     =   "Сохранить в БД"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CopyText"
            Object.ToolTipText     =   "Скопировать фрагмент"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PasteText"
            Object.ToolTipText     =   "Вставить фрагмент"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CutText"
            Object.ToolTipText     =   "Вырезать фрагмент"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DeleteText"
            Object.ToolTipText     =   "Удалить фрагмент"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Properties"
            Object.ToolTipText     =   "Настройки шрифта"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   3960
      Top             =   660
      _ExtentX        =   820
      _ExtentY        =   820
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TextEditForm.frx":0084
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TextEditForm.frx":0196
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TextEditForm.frx":02A8
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TextEditForm.frx":03BA
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TextEditForm.frx":04CC
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TextEditForm.frx":05DE
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TextEditForm.frx":06F0
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TextEditForm.frx":0802
            Key             =   "Redo"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TextEditForm.frx":0914
            Key             =   "Properties"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "TextEditForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public res%
Dim dW%, dH%

Private Sub Form_Activate()
 With TextEdit
  .SelStart = Len(.Text)
 End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If (KeyCode = 27) Then Me.Hide
End Sub

Private Sub Form_Load()
 res = 0
 dW = Width - TextEdit.Width
 dH = Height - TextEdit.Height
End Sub

Private Sub Form_Resize()
 Min% = Height - dH
 If (Min <= 1000) Then: Min = 1000: Height = dH + Min
 TextEdit.Height = Min
 
 Min = Width - dW
 If (Min <= 1000) Then: Min = 1000: Width = dW + Min
 TextEdit.Width = Min
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
 On Error Resume Next
 Select Case Button.Key
  Case "ClearText"
    TextEdit.TextRTF = ""
  Case "SaveText"
    res = 1
    Hide
  Case "CopyText"
    Clipboard.SetText (TextEdit.SelText)
  Case "PasteText"
    TextEdit.SelText = VB.Clipboard.GetText
  Case "CutText"
    Clipboard.SetText (TextEdit.SelText)
    TextEdit.SelText = ""
  Case "DeleteText"
    TextEdit.SelText = ""
  Case "Properties"
    On Error GoTo checkerror
    FontDlg.ShowFont
    TextEdit.Font.Name = FontDlg.FontName
    TextEdit.Font.Bold = FontDlg.FontBold
    TextEdit.Font.Italic = FontDlg.FontItalic
    TextEdit.Font.Size = FontDlg.FontSize
    TextEdit.Font.Strikethrough = FontDlg.FontStrikethru
    TextEdit.Font.Underline = FontDlg.FontUnderline
    Exit Sub
checkerror:
    MsgBox "error"
 End Select
End Sub

