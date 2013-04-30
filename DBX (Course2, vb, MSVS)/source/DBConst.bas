Attribute VB_Name = "DBConst"
' результаты работы диалогов из MsgBox
Public Const resBad = 0     '   выход, закрытием окна
Public Const resOk = 1      '   Да
Public Const resNo = 2      '   Нет
Public Const resCancel = 3  '   Отмена

' константы типов данных
Public Const ccInteger As Byte = 0
Public Const ccString As Byte = 1

' флаги доступа доступа к БД
    ' требовать пароль для входа
Public Const flPasswordNeed As Byte = 1
    ' запрещать доступ на чтение без пароля
Public Const flReadOnlyEnable As Byte = 2
    ' зашифрованность данных
Public Const flCoded As Byte = 4

' для диаграмм
Type TDiagElem
    Text As String
    Val As Integer
    Color As Long
End Type

' права Только чтение
Public Sub ProtectedMsg()
 Call MsgForm.ErrorMsg("Недостаточно прав для выполнения действия!")
End Sub

' звук нажатия кнопки
Public Sub SoundClick()
    Call sndPlaySound("Data\Click.wav", SND_ASYNC + SND_FILENAME + SND_LOOP + SND_APPLICATION)
End Sub

Public Function IsInteger(ByVal str$) As Boolean
 Dim Arr(1 To 4) As String * 1
 Arr(1) = "e": Arr(2) = "E": Arr(3) = ",": Arr(4) = "."
 IsInteger = True
 If IsNumeric(str) Then
  For i% = LBound(Arr) To UBound(Arr)
   If (InStr(1, str, Arr(i)) > 0) Then
    IsInteger = False
    Exit For
   End If
  Next i
 Else
  IsInteger = False
 End If
End Function

Public Sub ButEnabled(Pict As Image, Lbl As Label, enbl As Boolean)
    If enbl Then
        Pict.Picture = MainForm.ButtonImageList.ListImages(1).Picture
        Lbl.MousePointer = 1
    Else
        Pict.Picture = MainForm.ButtonImageList.ListImages(2).Picture
        Lbl.MousePointer = 12
    End If
    Lbl.Tag = CInt(enbl)
End Sub
