PRODUCT_COLUMNS = {
    1: "Код товара",
    2: "Артикул",
    3: "Категория",
    4: "Название товара",
    5: "Название товара (укр)",
    6: "Единица измерения",
    7: "Инструкция",
    8: "Сертификат",
    9: "Кратность количества в заказе",
    10: "Бренд",
    11: "Описание",
    12: "Описание (укр)",
    13: "Видео с ЮТУБА",
    14: "Обзоры статьи",
    15: "Сопутствующие товары",

    # Фото № 1-15; cells: 16 - 30
    **{i: f"Фото № {i - 15}" for i in range(16, 30 + 1)},

    # С этим товаром покупают № 1-12; cells: 31 - 42
    **{i: f"С этим товаром покупают № {i - 30}" for i in range(31, 42 + 1)},
}

PULT_MACROS = """
Sub Показать()
    With Sheets("Data")
        arr = .UsedRange.Value
        Application.ScreenUpdating = False
        For i = 41 To UBound(arr, 2)
            .Columns(i).Hidden = False
        Next
    End With
    Application.ScreenUpdating = True
End Sub
Sub Скрыть()
    With Sheets("Data")
        arr = .UsedRange.Value
        Application.ScreenUpdating = False
        For i = 41 To UBound(arr, 2)
            arr(1, i) = ""
        Next
        For i = 2 To UBound(arr)
            If .Rows(i).Hidden = False Then
                For j = 41 To UBound(arr, 2)
                    If arr(i, j) <> "" Then arr(1, j) = 1
                Next
            End If
        Next
        For i = 41 To UBound(arr, 2)
            If arr(1, i) = "" Then .Columns(i).Hidden = True
        Next
    End With
    Application.ScreenUpdating = True
End Sub
Sub Открыть_файл_с_путем()
    Application.ScreenUpdating = False
    FilesToOpen = Application.GetOpenFilename _
        ("Excel files(*.xls*),*.xls*", 1, "Выбрать файл", , False)
    If TypeName(FilesToOpen) = "Boolean" Then
        MsgBox "Файл не выбран!"
        Exit Sub
    End If
    With Sheets("Data")
        .Cells.Clear
        .Cells.Columns.Hidden = True
        On Error Resume Next
            .ShowAllData
        On Error GoTo 0
    End With
    
    Set importWb = Workbooks.Open(FilesToOpen)
        ActiveSheet.Cells.Copy ThisWorkbook.Sheets("Data").Cells
    importWb.Close False
    Sheets("Data").Select
    [A1].AutoFilter
    Application.ScreenUpdating = True
    MsgBox "Данные загружены успешно!", vbInformation, "Информация"
End Sub
Sub Убрать_переводы_строки()
    Application.ScreenUpdating = False

    With Sheets("Data")
        arr = .UsedRange.Value
        Application.ScreenUpdating = False
        For i = 2 To UBound(arr)
            arr(i, 10) = Replace(arr(i, 10), vbCrLf, "")
        Next
        .UsedRange = arr
    End With
    Application.ScreenUpdating = True
End Sub
Sub DuplicateSearch()
Dim ps As Long, myRange As Range, i1 As Long, i2 As Long, flag As Boolean
flag = False

'Определяем номер последней строки таблицы
Sheets("Data").Activate
ps = Cells(1, 1).CurrentRegion.Columns.Count
    'Нет смысла искать дубликаты в таблице, состоящей из одной строки
    If ps > 1 Then
    'Присваиваем объектной переменной ссылку на исследуемый столбец
    Set myRange = Range(Cells(1, 1), Cells(1, ps))
        With myRange
        'Очищаем ячейки столбца от предыдущих закрашиваний
        .Interior.Color = xlNone
            For i1 = 1 To ps - 1
                For i2 = i1 + 1 To ps
                    If .Cells(i1) = .Cells(i2) Then
                        'Если значения сравниваемых ячеек совпадают,
                        'обеим присваиваем новый цвет заливки
                        .Cells(i1).Interior.Color = 6740479
                        .Cells(i2).Interior.Color = 6740479
                        flag = True
                    End If
                Next
            Next
        End With
    End If
    If flag = True Then MsgBox ("Знайдені однакові характеристики!")
End Sub
Sub DeleteDoubleSpaces()
    Sheets("Data").Activate
    Selection.Cells.Replace What:="  ", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="  ", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
End Sub
Sub colorsRUtoUA()

    Selection.Cells.Replace What:="серый", Replacement:="сірий", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="красный", Replacement:="червоний", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="синий", Replacement:="синій", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="черный", Replacement:="чорний", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="красный", Replacement:="червоний", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="белый", Replacement:="білий", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="оранжевый", Replacement:="помаранчевий", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="зеленый", Replacement:="зелений", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="желтый", Replacement:="жовтий", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="коричневый", Replacement:="коричневий", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="фиолетовый", Replacement:="фіолетовий", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="прозрачный", Replacement:="прозорий", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="желто", Replacement:="жовто", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="серая", Replacement:="сіра", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="красная", Replacement:="червона", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="синяя", Replacement:="синя", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="черная", Replacement:="чорна", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="красная", Replacement:="червона", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="белая", Replacement:="біла", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="оранжевая", Replacement:="помаранчева", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="зеленая", Replacement:="зелена", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="желтая", Replacement:="жовта", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="коричневая", Replacement:="коричнева", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="фиолетовая", Replacement:="фіолетова", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="прозрачная", Replacement:="прозора", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
End Sub
Sub colorsUAtoRU()

    Selection.Cells.Replace What:="сірий", Replacement:="серый", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="червоний", Replacement:="красный", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="синій", Replacement:="синий", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="червоний", Replacement:="красный", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="білий", Replacement:="белый", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="чорний", Replacement:="черный", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="помаранчевий", Replacement:="оранжевый", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="зелений", Replacement:="зеленый", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="жовтий", Replacement:="желтый", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="коричневий", Replacement:="коричневый", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="фіолетовий", Replacement:="фиолетовый", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="прозорий", Replacement:="прозрачный", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="жовто", Replacement:="желто", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="сіра", Replacement:="серая", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="червона", Replacement:="красная", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="синя", Replacement:="синяя", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="червона", Replacement:="красная", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="біла", Replacement:="белая", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="чорна", Replacement:="черная", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="помаранчева", Replacement:="оранжевая", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="зелена", Replacement:="зеленая", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="жовта", Replacement:="желтая", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="коричнева", Replacement:="коричневая", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="фіолетова", Replacement:="фиолетовая", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Cells.Replace What:="прозориа", Replacement:="прозрачная", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
End Sub
Function zagl(ByVal x As String) As String
    zagl = UCase(Left(x, 1)) & LCase(Mid(x, 2))
End Function
Sub UpFirstSimbol()
    For Each c In Selection.Cells
        c.Value = zagl(c)
    Next
    
End Sub
Sub LowerSimbol()
    For Each c In Selection.Cells
        c.Value = LCase(c)
    Next
    
End Sub
Sub AppendTextRight()
    AppendText = InputBox("Що приклеїти?")
    For Each c In Selection.Cells
        c.Value = c & AppendText
    Next
End Sub
Sub HideColumnsNoColor()
    Dim ws As Worksheet
    Dim lastColumn As Long
    Dim i As Long
    
    ' Устанавливаем ссылку на лист "Data"
    Set ws = ThisWorkbook.Sheets("Data")
    
    ' Определяем последнюю используемую колонку
    lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Проверяем каждую колонку, начиная с 43
    For i = 43 To lastColumn
        ' Проверяем цвет заливки первой ячейки в колонке
        If ws.Cells(1, i).Interior.ColorIndex = xlNone Then
            ' Скрываем колонку
            ws.Columns(i).Hidden = True
        End If
    Next i
End Sub
"""