Attribute VB_Name = "SyncModule"
Option Explicit
'кракакатау1
' Процедура для экспорта всех компонентов VBA в папку /src
Public Sub ExportVbaSourceFiles()
    Dim vbComp As VBIDE.VBComponent
    Dim destFolder As String
    Dim extension As String
    Dim fileName As String
    
    ' 1. Определяем путь к папке src относительно документа
    ' Если документ еще не сохранен, выходим
    If Len(ThisDocument.Path) = 0 Then
        MsgBox "Сначала сохраните документ на диске!", vbCritical
        Exit Sub
    End If
    
    destFolder = ThisDocument.Path & "\src\"
    
    ' 2. Создаем папку src, если она не существует
    If Dir(destFolder, vbDirectory) = "" Then
        MkDir destFolder
    End If
    
    ' 3. Перебираем все компоненты проекта
    For Each vbComp In ThisDocument.VBProject.VBComponents
        
        ' Определяем расширение в зависимости от типа компонента
        Select Case vbComp.Type
            Case vbext_ct_StdModule
                extension = ".bas"
            Case vbext_ct_ClassModule
                extension = ".cls"
            Case vbext_ct_MSForm
                extension = ".frm"
            Case vbext_ct_Document
                ' ThisDocument и модули страниц
                extension = ".cls"
            Case Else
                extension = ".txt"
        End Select
        
        ' Формируем полное имя файла
        fileName = destFolder & vbComp.Name & extension
        
        ' 4. Экспортируем компонент
        On Error Resume Next
        vbComp.Export fileName
        If Err.Number <> 0 Then
            Debug.Print "Ошибка при экспорте " & vbComp.Name & ": " & Err.Description
            Err.Clear
        Else
            Debug.Print "Экспортировано: " & vbComp.Name & extension
        End If
        On Error GoTo 0
    Next vbComp
    
    MsgBox "Экспорт завершен! Файлы находятся в: " & destFolder, vbInformation
End Sub

' Процедура для импорта всех файлов из папки /src обратно в документ
Public Sub ImportVbaSourceFiles()
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim fso As Object
    Dim srcFolder As String
    Dim file As Object
    Dim fileName As String
    Dim compName As String
    Dim ext As String

    Set vbProj = ThisDocument.VBProject
    
    ' 1. Путь к папке src
    srcFolder = ThisDocument.Path & "\src\"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(srcFolder) Then
        MsgBox "Папка src не найдена!", vbCritical
        Exit Sub
    End If

    ' Отключаем предупреждения, чтобы процесс шел быстрее
    Application.ScreenUpdating = False

    ' 2. Перебираем файлы в папке src
    For Each file In fso.GetFolder(srcFolder).Files
        fileName = file.Path
        ext = LCase(fso.GetExtensionName(fileName))
        
        ' Нас интересуют только исходники (.bas, .cls, .frm)
        If ext = "bas" Or ext = "cls" Or ext = "frm" Then
            compName = fso.GetBaseName(fileName)
            
            ' Игнорируем ThisDocument, так как его нельзя удалить/импортировать просто так
            ' Также игнорируем текущий модуль, чтобы не прервать выполнение кода
            If compName <> "ThisDocument" And compName <> "SyncModule" Then
                
                ' 3. Если компонент с таким именем уже есть, удаляем его
                On Error Resume Next
                Set vbComp = vbProj.VBComponents(compName)
                If Not vbComp Is Nothing Then
                    vbProj.VBComponents.Remove vbComp
                End If
                On Error GoTo 0
                
                ' 4. Импортируем файл
                vbProj.VBComponents.Import fileName
                Debug.Print "Импортирован: " & compName
            End If
        End If
    Next file

    Application.ScreenUpdating = True
    MsgBox "Импорт завершен!", vbInformation
End Sub
