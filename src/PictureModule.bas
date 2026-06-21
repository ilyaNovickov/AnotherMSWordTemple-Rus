Attribute VB_Name = "PictureModule"
Sub InsertGostFigure()
    Dim figPath As String

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
    Selection.InlineShapes.AddPicture figPath
    
    MakeFigure
    
End Sub

Sub MakeGostFigure()

    If Selection.Type <> wdSelectionInlineShape Then
        MsgBox "Выделен не рисунок"
        Exit Sub
    End If
    
    MakeFigure

End Sub

Private Sub MakeFigure()
    Dim PicStyleName As String
    Dim PicNameStyleName As String
    Dim PicCaption As String
    
    PicStyleName = PICTURE_STYLE
    PicNameStyleName = PUCTURE_NAME_STYLE
    PicCaption = PICTURE_CAPTION
    
    ' 3. Вставляем рисунок
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter ' Центрируем абзац рисунка
    Selection.ParagraphFormat.FirstLineIndent = InchesToPoints(0)
    Selection.ParagraphFormat.KeepWithNext = True
    On Error Resume Next
    Selection.Style = PicStyleName
    On Error GoTo 0
    
    ' 4. Переходим в конец строки с рисунком и делаем отступ
    Selection.EndKey Unit:=wdLine
    Selection.TypeParagraph
    
    ' 5. Вставляем подпись
    ' Центрируем абзац подписи перед применением стиля (надежная страховка)
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.ParagraphFormat.FirstLineIndent = InchesToPoints(0)
        
    ' Если стиль не применился программно, вызываем стандартное окно,
    ' но абзац уже отцентрован, так что проблем не будет
    Selection.InsertCaption Label:=PicCaption, TitleAutoText:="InsertCaption", _
            Title:="", Position:=wdCaptionPositionBelow, ExcludeLabel:=0
            
    ' Применяем стиль "Название" (Caption)
    On Error Resume Next ' На случай если стиль называется иначе
    Selection.Style = PicNameStyleName
    On Error GoTo 0

End Sub
'кракакатау
