Attribute VB_Name = "TableModule"
' 1. Макрос для вставки новой таблицы 3x3
Sub InsertGostTable3x3()
    Dim doc As Document
    Dim tbl As Table
    Dim rng As Range
    
    Set doc = ActiveDocument
    Set rng = Selection.Range
    
    ' Вставляем пустую строку перед таблицей для отступа (необязательно, зависит от стиля)
    ' rng.InsertParagraphAfter
    ' Set rng = doc.Paragraphs(doc.Paragraphs.Count).Range
    
    ' Создаем таблицу 3x3
    Set tbl = doc.Tables.Add(Range:=rng, NumRows:=3, NumColumns:=3)
    
    ' Применяем оформление и вставляем подпись
    FormatTableAndAddCaption tbl
End Sub

' 2. Макрос для форматирования выделенной таблицы
Sub FormatSelectedTable()
    If Selection.Information(wdWithInTable) Then
        Dim tbl As Table
        Set tbl = Selection.Tables(1)
        FormatTableAndAddCaption tbl
    Else
        MsgBox "Курсор должен находиться внутри таблицы!", vbExclamation, "Ошибка"
    End If
End Sub

' Вспомогательная процедура для форматирования и вставки подписи
Private Sub FormatTableAndAddCaption(ByVal tbl As Table)
    Dim captionRange As Range
    Dim doc As Document
    Set doc = ActiveDocument
    
    ' --- 0. Применение стиля текста внутри таблицы ---
    ' Если стиля нет, применяем базовое форматирование вручную
    With tbl.Range.ParagraphFormat
        .CharacterUnitFirstLineIndent = 0
        .FirstLineIndent = CentimetersToPoints(0)
        .Alignment = wdAlignParagraphCenter
    End With
    ' Применяем созданный тобой стиль ко всей таблице
    On Error Resume Next
    tbl.Range.Style = TABLE_TEXT_STYLE
    On Error GoTo 0
    
    ' --- 1. Форматирование самой таблицы по ГОСТ ---
    With tbl
        ' Установка границ (тонкие черные линии)
        With .Borders
            .InsideLineStyle = wdLineStyleSingle
            .OutsideLineStyle = wdLineStyleSingle
            .InsideLineWidth = wdLineWidth050pt
            .OutsideLineWidth = wdLineWidth050pt
            .InsideColor = wdColorAutomatic
            .OutsideColor = wdColorAutomatic
        End With
        
        .Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter ' ВЕРТИКАЛЬНО ПО ЦЕНТРУ
        .Range.ParagraphFormat.Alignment = wdAlignParagraphCenter  ' ГОРИЗОНТАЛЬНО ПО ЦЕНТРУ
        
        .TopPadding = CentimetersToPoints(0)
        .BottomPadding = CentimetersToPoints(0)
        .LeftPadding = CentimetersToPoints(0.1)
        .RightPadding = CentimetersToPoints(0.1)
        .Spacing = 0
        .AllowPageBreaks = True
        .AllowAutoFit = True
        
        ' Автоподбор по ширине окна (обычно для ГОСТ)
        .AutoFitBehavior (wdAutoFitWindow)
        
        ' Повторять заголовок на каждой странице (ГОСТ требование)
        .Rows(1).HeadingFormat = True
    End With
    
    ' --- 2. Вставка подписи ---
    ' Устанавливаем диапазон ПЕРЕД таблицей
    Set captionRange = tbl.Range
    captionRange.Collapse Direction:=wdCollapseStart
    
    ' Вставляем название (Word автоматически добавит номер)
    ' Используем стандартный метод InsertCaption
    captionRange.InsertCaption Label:=TABLE_CAPTION, _
                               Title:="", _
                               Position:=wdCaptionPositionAbove


                               
    ' --- 3. Применение стиля к подписи ---
    ' После InsertCaption курсор или диапазон может сместиться.
    ' Находим абзац непосредственно перед таблицей.
    Dim captionParagraph As Paragraph
    Set captionParagraph = tbl.Range.Paragraphs(1).Previous
    
    On Error Resume Next ' На случай, если стиль не создан
    captionParagraph.Range.Style = TABLE_NAME_STYLE
    On Error GoTo 0
    
    ' Устанавливаем курсор в конец подписи, чтобы пользователь мог сразу писать название
    captionParagraph.Range.Select
    Selection.Collapse Direction:=wdCollapseEnd
    Selection.MoveLeft Unit:=wdCharacter, Count:=1 ' Встаем перед знаком абзаца
End Sub
' 3. МАКРОС ДЛЯ РАЗРЫВА ТАБЛИЦЫ (ГОСТ)
Sub SplitTableForGost()
    Dim tbl As Table
    Dim rowIdx As Long
    Dim doc As Document: Set doc = ActiveDocument
    Dim headerRow As Range
    Dim tblNew As Table
    
    ' Проверка: находится ли курсор в таблице
    If Not Selection.Information(wdWithInTable) Then
        MsgBox "Поставьте курсор в строку, с которой начнется новая страница.", vbExclamation
        Exit Sub
    End If

    Set tbl = Selection.Tables(1)
    rowIdx = Selection.Cells(1).RowIndex
    
    ' Нельзя разбить по первой строке (шапке)
    If rowIdx = 1 Then
        MsgBox "Нельзя разбить таблицу перед первой строкой (шапкой).", vbExclamation
        Exit Sub
    End If
    
    ' 0. Запоминаем шапку первой таблицы (первую строку)
    ' Мы копируем её содержимое и форматирование
    Set headerRow = tbl.Rows(1).Range
    headerRow.End = headerRow.End - 1 ' Исключаем маркер конца строки, чтобы избежать ошибок вставки
    headerRow.Copy

    ' 1. Разделяем таблицу
    ' Вставляем разрыв страницы прямо перед выбранной строкой
    Selection.InsertBreak Type:=wdPageBreak
    
    ' 2. Вставляем текст между таблицами
    ' После разрыва страницы курсор обычно оказывается в новом абзаце перед второй частью таблицы
    Selection.TypeParagraph ' Создаем дополнительный абзац, если нужно
    Selection.MoveUp Unit:=wdLine, Count:=1
    
    ' Пишем текст
    Selection.TypeText "Продолжение таблицы "
    
    ' 3. Вставляем ДИНАМИЧЕСКОЕ ПОЛЕ (SEQ Таблица \c)
    ' Это поле берет текущий номер последовательности без увеличения
    doc.Fields.Add Range:=Selection.Range, _
                   Type:=wdFieldEmpty, _
                   Text:="SEQ " & TABLE_CAPTION & " \c", _
                   PreserveFormatting:=True
                
    ' Применяем стиль названия
    On Error Resume Next
    Selection.Style = TABLE_NAME_STYLE
    On Error GoTo 0
    
    ' 4. Финальное оформление
    ' Убеждаемся, что заголовок (шапка) первой таблицы дублируется во второй
    tbl.Rows(1).HeadingFormat = True
    
    '5. Работа со второй (новой) таблицей
    ' Спускаемся к новой таблице
    Selection.MoveDown Unit:=wdLine, Count:=1
    
    ' Теперь Selection находится в первой ячейке второй таблицы.
    ' Нам нужно вставить новую строку сверху и вставить туда скопированную шапку.
    Set tblNew = Selection.Tables(1)
    tblNew.Rows.Add BeforeRow:=tblNew.Rows(1)
    
    ' Вставляем данные шапки в новую созданную строку
    tblNew.Rows(1).Range.Paste
    
    ' Чтобы не было пустой строки внутри ячеек после вставки (баг Word),
    ' иногда требуется небольшая очистка, но Paste обычно работает хорошо.
    
    ' 6. Убеждаемся, что обе таблицы имеют правильные свойства
    ' (Вертикальное выравнивание и отсутствие обтекания)
    FormatTableAfterSplit tbl
    FormatTableAfterSplit tblNew
    
    ' Ставим курсор в конец текста "Продолжение таблицы..."
    ' чтобы пользователь мог сразу видеть результат
    'doc.Paragraphs(tblNew.Range.Paragraphs(1).Index - 1).Range.Select
    'Selection.Collapse Direction:=wdCollapseEnd
End Sub
' Дополнительная мини-процедура для фиксации свойств после разрыва
Private Sub FormatTableAfterSplit(ByVal tbl As Table)
    With tbl
        .Rows.WrapAroundText = False ' Отключаем обтекание (обязательно!)
        .Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter
        .Rows(1).HeadingFormat = True ' На случай дальнейших разрывов
    End With
End Sub

