Attribute VB_Name = "PictureModule"
Sub MakeGostFigure()
    ' HW
    Dim PicStyleName As String
    Dim PicNameStyleName As String
    Dim PicCaption As String
    
    PicStyleName = "РисКонтейн"
    PicNameStyleName = "Название объекта"
    PicCaption = "Рисунок"

    If Selection.Type <> wdSelectionInlineShape Then
        MsgBox "Выделен не рисунок"
    End If
    
    Selection.ParagraphFormat.FirstLineIndent = InchesToPoints(0)
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    On Error Resume Next
    Selection.Style = PicStyleName
    On Error GoTo 0
    
    'Переход на след строку
    Selection.EndKey Unit:=wdLine
    Selection.TypeParagraph
    
    ' Применяем стиль "Название" (Caption)
    Selection.ParagraphFormat.FirstLineIndent = InchesToPoints(0)
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    On Error Resume Next ' На случай если стиль называется иначе
    Selection.Style = PicNameStyleName
    On Error GoTo 0

    'Вставка названия
    Selection.InsertCaption Label:=PicCaption, TitleAutoText:="InsertCaption", _
            Title:="", Position:=wdCaptionPositionBelow, ExcludeLabel:=0


End Sub

Sub InsertGostFigure()
    
    Dim figPath As String
    Dim PicStyleName As String
    Dim PicNameStyleName As String
    Dim PicCaption As String
    
    PicStyleName = "РисКонтейн"
    PicNameStyleName = "Название объекта"
    PicCaption = "Рисунок"
    
    ' 1. Открываем диалог выбора файла
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Выберите изображение"
        .Filters.Add "Images", "*.jpg; *.jpeg; *.png; *.bmp; *.gif", 1
        If .Show = -1 Then
            figPath = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    ' 2. Вставляем рисунок
    Selection.TypeParagraph ' Создаем новый абзац для рисунка
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter ' Центрируем абзац рисунка
    Selection.ParagraphFormat.FirstLineIndent = InchesToPoints(0)
    On Error Resume Next
    Selection.Style = PicStyleName
    On Error GoTo 0
    Selection.InlineShapes.AddPicture figPath
    
    ' 3. Переходим в конец строки с рисунком и делаем отступ
    Selection.EndKey Unit:=wdLine
    Selection.TypeParagraph
    
    ' 4. Вставляем подпись
    ' Центрируем абзац подписи перед применением стиля (надежная страховка)
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.ParagraphFormat.FirstLineIndent = InchesToPoints(0)
    
    ' Применяем стиль "Название" (Caption)
    On Error Resume Next ' На случай если стиль называется иначе
    Selection.Style = PicNameStyleName
    On Error GoTo 0
    
    ' Если стиль не применился программно, вызываем стандартное окно,
    ' но абзац уже отцентрован, так что проблем не будет
    Selection.InsertCaption Label:=PicCaption, TitleAutoText:="InsertCaption", _
            Title:="", Position:=wdCaptionPositionBelow, ExcludeLabel:=0

End Sub

