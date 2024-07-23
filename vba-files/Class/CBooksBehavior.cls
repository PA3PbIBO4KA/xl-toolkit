VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBooksBehavior"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' Name: CBooksBehavior
' Author: Михаил Красюк
' Date: 13.06.2024

Option Explicit

Private m_wbMacros As Workbook
Private m_wsMacrosMainMenu As Worksheet
Private m_wsMacrosReferences As Worksheet

Private m_wbCommon As Workbook
Private m_wsCommonPrimary As Worksheet

Private m_wbPreviousCommon As Workbook
Private m_wsPreviousCommonPrimary As Worksheet

Private m_wbForm As Workbook
Private m_wsFormPrimary As Worksheet

Private m_wbASR As Workbook
Private m_wsASRPrimary As Worksheet

Private m_wbSplitting As Workbook
Private m_wsSplittingPrimary As Worksheet

Private m_sCommonBookName As String

' Инициализация
Private Sub class_initialize()
    
    On Error GoTo ErrorHandler_OpenFile
    
    Lib.AddLogNote "Macros Launched"
    
    Set m_wbMacros = Workbooks(ThisWorkbook.Name)
    
    With m_wbMacros
        Set m_wsMacrosMainMenu = .Worksheets("Главное меню")
        Set m_wsMacrosReferences = .Worksheets("Референсы")
    End With
    
    Exit Sub
    
ErrorHandler_OpenFile:

    Call Lib.FatalError("Не удалось открыть файл")
    Lib.DisableOptimization
    End
    
End Sub

' Переименнование книг в зависимости от даты
Private Sub RenameWorkbook(ByRef Book As Workbook, ByRef Reference As String)
    
    On Error GoTo ErrorHandler_Rename
    
    Dim sYearToday As String: sYearToday = CStr(Year(Now))
    Dim sOldName   As String: sOldName = Book.Name
    Dim sOldRef    As String: sOldRef = Reference
    Dim iYear      As Long
    
    For iYear = Year(Now) To 2024 Step -1
        If InStr(sOldName, CStr(iYear)) Then
            
            sOldName = Replace(sOldName, CStr(iYear), sYearToday, 1, 1)
            Reference = Replace(sOldRef, CStr(iYear), sYearToday, 1, 1)
            
            Exit For
            
        End If
    Next iYear
    
    Exit Sub
    
ErrorHandler_Rename:

    Call Lib.FatalError("Не удалось переименовать файл")
    Lib.DisableOptimization
    End
    
End Sub

' Открытие последнего свода
Private Sub OpenPreviousCommon(Optional ByVal ReadOnly As Boolean = False)
        
    On Error GoTo ErrorHandler_OpenFile
        
    Dim sReference  As String
    Dim i           As Long
        
    ' Получение ссылки на книгу для сравнения сумм
    For i = 13 To 2 Step -1
        If m_wsCommonPrimary.Cells(2, 35).Value Like m_wsMacrosReferences.Cells(i, 6).Value & "*" Then
                
            ' NOTE -- Мы берем на месяц назад, потому что отчет составляется за прошлый месяц
            ' P.S. -- тут должен быть прописан путь
            If Month(Now) = 1 Then
                sReference = Year(Now) & "-" & "12" & "\1\" & "*" & ".xlsx"
            ElseIf Month(Now) > 10 Then
                sReference = Year(Now) & "-" & Month(Now) - 1 & "\1\" & "*" & ".xlsx"
            Else
                sReference = Year(Now) & "-" & "0" & Month(Now) - 1 & "\1\" & "*" & ".xlsx"
            End If
            
            If Dir(sReference) <> "" Then
                Exit For
            Else
                Exit Sub
            End If
                
        End If
    Next i
    
    Set m_wbPreviousCommon = Workbooks.Open(Filename:=sReference, ReadOnly:=ReadOnly)
    Set m_wsPreviousCommonPrimary = m_wbPreviousCommon.Worksheets(1)
    
    Exit Sub
    
ErrorHandler_OpenFile:

    Call Lib.FatalError("Не удалось открыть файл")
    Lib.DisableOptimization
    End
    
End Sub

Public Sub CreateXLSXFile(sPath As String)
    
    On Error GoTo ErrorHandler_CreateFile
    
    ' Loading 5%
    With frmLoading
        .SetInfo "Создание .xlsx файла..."
        .SetProgress 5
    End With
    
    Call Lib.Delay(2000)
    
    ' Формируем имя для итогового свода
    If Month(Now) > 9 Then
        m_sCommonBookName = "Свод " & Month(Now) & "." & Year(Now)
    Else
        m_sCommonBookName = "Свод 0" & Month(Now) & "." & Year(Now)
    End If
    
    ' Удаляем существующий итоговый свод если есть
    If Lib.IsFileExists(sPath & "\" & m_sCommonBookName & ".xlsx") Then
        Call Kill(sPath & "\" & m_sCommonBookName & ".xlsx")
    End If
    
    ' Создаем файл итогового свода
    Set m_wbCommon = Workbooks.Add: Set m_wsCommonPrimary = m_wbCommon.Worksheets(1)
        
    With m_wbCommon
        .Worksheets(1).Name = "свод"
        .SaveAs sPath & "\" & m_sCommonBookName & ".xlsx"
    End With
    
    Exit Sub
    
ErrorHandler_CreateFile:

    Call Lib.FatalError("Не удалось создать .xlsx файл")
    Lib.DisableOptimization
    End
    
End Sub

' Создание таблицы
Public Sub CreateCommonTable()
    
    On Error GoTo ErrorHandler_CreateCommonTable
    
    ' Loading 15%
    With frmLoading
        .SetInfo "Формирование табличной шапки..."
        .SetProgress 15
    End With
    
    Call Lib.Delay(2000)
    
    ' Кол-во ссылок на сводные файлы по регионам и ссылка на проформу таблицы
    Dim nRegionalReferences As Long: nRegionalReferences = m_wsMacrosReferences.Cells(2, 1).End(xlDown).Row
    Dim sHeadersReference   As String: sHeadersReference = m_wsMacrosReferences.Cells(2, 3).Value
    
    ' Для первого месяца
    Dim rngPeriod As Range
    
    ' Книга с шаблоном
    Dim wbHeaders As Workbook:  Set wbHeaders = Workbooks.Open(Filename:=sHeadersReference, ReadOnly:=True)
    Dim wsHeaders As Worksheet: Set wsHeaders = wbHeaders.Worksheets("свод")
    
    ' Снимает фильтры, раскрывает столбцы
    With wsHeaders
        .AutoFilterMode = False
        .Columns.Hidden = False
        .Rows.Hidden = False ' NOTE -- Этого в ТЗ нет но я добавил
    End With
    
    ' Ширина шапки
    Dim nHeaderColumns As Long: nHeaderColumns = Lib.GetColumnsCount(Row:=3, Sheet:=wsHeaders)
    
    ' NOTE -- iPasteRow - строка вставки диапозона
    Dim i, j, k, iPasteRow As Long: iPasteRow = 5
    
    ' Года, месяца, шапка шаблона и их инициализация
    
    Dim arrYears(), arrMonth() As String
    Dim arrHeaders()           As String
    Dim arrRefs()              As String
    
    ReDim arrYears(2 To 5)
    ReDim arrMonth(2 To 13)
    ReDim arrHeaders(1 To nHeaderColumns)
    ReDim arrRefs(2 To nRegionalReferences)
    
    For i = 2 To 5
        arrYears(i) = m_wsMacrosReferences.Cells(i, 5).Value
    Next i
    
    For i = 2 To 13
        arrMonth(i) = m_wsMacrosReferences.Cells(i, 6).Value
    Next i
    
    For i = 1 To nHeaderColumns
        arrHeaders(i) = LCase(wsHeaders.Cells(3, i).Value)
    Next i
    
    ' Добавляет шапку в итоговый свод
    wsHeaders.Range("A1:" & "AQ4").Copy Destination:=m_wsCommonPrimary.Range("A1")
    
    ' Дублирование периодов шапки
    ' NOTE -- минус 2 потому что, начинаем с нуля - раз; два - потому что формируем за прошлый месяц
    For i = 0 To Month(Now) - 2
        With m_wsCommonPrimary
            .Range(.Cells(1, 35 + (i * 9)), .Cells(4, 43 + (i * 9))).Copy Destination:=.Range(.Cells(1, 35 + ((i + 1) * 9)), .Cells(4, 43 + ((i + 1) * 9)))
        End With
    Next i
    
    ' Форматирует названия шапок
    i = 0
    For j = Month(Now) + 1 To 2 Step -1
    
        With m_wsCommonPrimary
            If Not arrMonth(j) = "Декабрь" Then
                .Cells(2, 35 + (i * 9)).Value = m_wsMacrosReferences.Cells(j, 6).Value & " " & Year(Now)
            Else
                .Cells(2, 35 + (i * 9)).Value = m_wsMacrosReferences.Cells(j, 6).Value & " " & Year(Now) - 1
            End If
        End With
        
        i = i + 1
        
    Next j
    
    ' Ширина итогового свода
    Dim nCommonColumns As Long: nCommonColumns = Lib.GetColumnsCount(Row:=3, Method:=mcByFiltering, Sheet:=m_wsCommonPrimary)
    
    ' Новый Итоговый СВОД
    Dim wbRegional As Workbook
    Dim wsRegionalPrimary As Worksheet
    
    With m_wsCommonPrimary
        .Range(.Cells(5, 1), .Cells(1048576, nCommonColumns)).NumberFormat = "@"
    End With
    
    For i = 2 To nRegionalReferences
        
        ' Присваеваем ссылку .xlsx файла
        Dim sRef As String: sRef = m_wsMacrosReferences.Cells(i, 1).Value
        
        On Error Resume Next
        Set wbRegional = Workbooks(sRef)
        On Error GoTo ErrorHandler_CreateCommonTable
        
        If Not wbRegional Is Nothing Then
            If Not InStr(wbRegional.Name, CStr(Year(Now))) > 0 Then
                Call RenameWorkbook(wbRegional, sRef)
            End If
        End If
        
        arrRefs(i) = CStr(sRef)
        
    Next i
    
    Lib.AddLogNote "Form of common table"
    
    ' Формируем итоговый свод
    For i = 2 To nRegionalReferences
        
        ' Открываем .xlsx файл по ссылке
        Application.DisplayAlerts = False
        Set wbRegional = Workbooks.Open(Filename:=arrRefs(i), ReadOnly:=True)
        Application.DisplayAlerts = True
        
        ' Loading +5% за каждый региональный свод
        With frmLoading
            .SetInfo "Заполнение из " & wbRegional.Name & "..."
            .IncreaseProgress 5
        End With
        
        Set wsRegionalPrimary = wbRegional.Worksheets("свод")
        
        ' Снимает фильтры, раскрывает столбцы
        With wsRegionalPrimary
            .AutoFilterMode = False
            .Columns.Hidden = False
        End With
        
        ' Кол-во столбцов и строк конкретно открытого файла
        Dim nRows    As Long: nRows = Lib.GetRowsCount(Column:=1, Method:=mcByEnd, Sheet:=wsRegionalPrimary)
        Dim nColumns As Long: nColumns = Lib.GetColumnsCount(Row:=3, Method:=mcByFiltering, Sheet:=wsRegionalPrimary)
        
        ' Буферная переменная
        Dim rngBuffer As Range
        
        Dim iYear, iMonth, iMonthOld As Long
        
        ' Вставка строк в первый диапозон
        For j = 1 To nHeaderColumns - 9
            For k = 1 To nColumns
                If arrHeaders(j) = LCase(wsRegionalPrimary.Cells(3, k).Value) Then
                    
                    ' Копипаст ячеек
                    
                    With wsRegionalPrimary
                        Set rngBuffer = .Range(.Cells(5, k), .Cells(nRows, k))
                    End With
                    
                    With m_wsCommonPrimary
                        
                        ' NOTE -- при первом заходе вставляются все данные ячеек
                        If i = 2 Then
                            rngBuffer.Copy
                            .Cells(iPasteRow, j).PasteSpecial xlPasteAll
                            Application.CutCopyMode = False
                        Else
                            .Range(.Cells(iPasteRow, j), .Cells(rngBuffer.Cells.Count + (iPasteRow - 1), j)).Value = rngBuffer.Value
                        End If
                        
                    End With
                    
                End If
            Next k
        Next j
        
        ' Выбирается первый месяц, с которой начинается Итоговый СВОД
        
        With wsRegionalPrimary
            Set rngPeriod = .Range(.Cells(2, 1), .Cells(2, nColumns))
        End With
        
        For Each rngBuffer In rngPeriod.Columns
            For iYear = Year(Now) To 2023 Step -1
                For iMonth = 13 To 2 Step -1
                    
                    Dim sDate   As String:  sDate = arrMonth(iMonth) & " " & iYear
                    Dim dt      As Date:    dt = CDate(sDate)
                    
                    If rngBuffer.Value = dt Then
                    
                        iMonth = Month(Now) - (iMonth - 1)
                        iMonthOld = iMonth
                        
                        GoTo BuildPeriods
                        
                    End If
                    
                Next iMonth
            Next iYear
        Next rngBuffer
        
BuildPeriods:
        
        ' Номер 35-ой строки в открытой таблице
        Set rngBuffer = wsRegionalPrimary.Cells(4, 1).EntireRow.Find("35", LookIn:=xlValues, LookAt:=xlPart)
        Dim n As Long: n = rngBuffer.Column
        
        ' Вставка строк в периодные диапозоны
        For iMonth = iMonth To Month(Now)
            For j = 35 To 43
                For k = n + ((iMonth - iMonthOld) * 9) To (n + 8) + ((iMonth - iMonthOld) * 9)
                    If arrHeaders(j) = LCase(wsRegionalPrimary.Cells(3, k).Value) Then
                        
                        ' Копипаст ячеек
                        
                        With wsRegionalPrimary
                            Set rngBuffer = .Range(.Cells(5, k), .Cells(nRows, k))
                        End With
                        
                        With m_wsCommonPrimary
                            
                            ' NOTE -- при первом заходе вставляются все данные ячеек
                            If i = 2 Then
                                rngBuffer.Copy
                                .Cells(iPasteRow, j + (iMonth * 9)).PasteSpecial xlPasteAll
                                Application.CutCopyMode = False
                            Else
                                With m_wsCommonPrimary
                                    .Range(.Cells(iPasteRow, j + (iMonth * 9)), .Cells(rngBuffer.Cells.Count + (iPasteRow - 1), j + (iMonth * 9))).Value = rngBuffer.Value
                                End With
                            End If
                            
                        End With
                        
                        Exit For
                        
                    End If
                Next k
            Next j
        Next iMonth
        
        ' Сдвигаем на сточку вниз место нашего последующего копипаста
        iPasteRow = Lib.GetRowsCount(Column:=1, Sheet:=m_wsCommonPrimary) + 1
        
        wbRegional.Close SaveChanges:=False
        
        Call Lib.Delay(3000)
        
    Next i
    
    Lib.AddLogNote "Form of additional table"
    
    ' Кол-во строк итогового свода
    Dim nCommonRows As Long: nCommonRows = Lib.GetRowsCount(Column:=1, Sheet:=m_wsCommonPrimary)
    
    nCommonColumns = Lib.GetColumnsCount(Row:=3, Method:=mcByFiltering, Sheet:=m_wsCommonPrimary)
    
    ' Форматирует таблицу
    With m_wsCommonPrimary
    
        Call Lib.AcceptCellsBorders(TableRange:=.Range(.Cells(5, 1), .Cells(5, nCommonColumns)))
        
        .Range(.Cells(5, 1), .Cells(5, nCommonColumns)).Font.Size = 8
        .Range(.Cells(5, 1), .Cells(5, nCommonColumns)).Font.Name = "Arial"
        
        .Range(.Cells(5, 1), .Cells(5, nCommonColumns)).Copy
        .Range(.Cells(6, 1), .Cells(nCommonRows, nCommonColumns)).PasteSpecial xlPasteFormats
        
        .Range(.Cells(5, 1), .Cells(5, nCommonColumns)).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range(.Cells(5, nCommonColumns), .Cells(nCommonRows, nCommonColumns)).Borders(xlEdgeRight).LineStyle = xlContinuous
        
    End With
    
    ' Loading 99%
    With frmLoading
        .SetInfo "Формирование сравнительной таблицы..."
        .SetProgress 99
    End With
    
    Call Lib.Delay(2000)
    
    ' Лист последнего свода
    Call OpenPreviousCommon(ReadOnly:=True)
    
    ' NOTE -- Итоговый СВОД за прошлый месяц может быть и не сформирован ко времени запуска макроса
    If Not m_wbPreviousCommon Is Nothing Then
        
        With m_wbCommon
            .Worksheets.Add After:=m_wbCommon.Worksheets("свод")
            .Worksheets(2).Name = "для сравнения"
        End With
        
        ' Лист для сравнения
        Dim wsAdditional As Worksheet: Set wsAdditional = m_wbCommon.Worksheets(2)
        
        ' Формируем таблицу для сравнения
        With wsAdditional
        
            Call Lib.AcceptCellsBorders(TableRange:=.Range("A1:C3"))
            
            .Cells(1, 2).Value = m_sCommonBookName
            
            If Month(Now) > 9 Then
                .Cells(1, 3).Value = Month(Now) & "." & Year(Now) & " ИТОГ  СВОД"
            Else
                .Cells(1, 3).Value = "0" & Month(Now) & "." & Year(Now) & " ИТОГ  СВОД"
            End If
            
            .Cells(2, 1).Value = "Сумма ДЗ"
            .Cells(3, 1).Value = "Сумм резерва ДЗ"
            
            .Range("A1:C1").Interior.Color = mcGreenColor
            .Range("A2:A3").Interior.Color = mcLightGreyColor
            
            .Range("A1:C1").Font.Bold = True
            .Range("A1:A3").Font.Bold = True
            
            .Range("B2:C3").NumberFormat = "0.00"
            
            ' Вставляет значения
            
            .Cells(2, 2).Value = m_wsPreviousCommonPrimary.Cells(1, 36).Value
            .Cells(3, 2).Value = m_wsPreviousCommonPrimary.Cells(1, 41).Value
            
            .Cells(2, 3).Value = m_wsCommonPrimary.Cells(1, 36).Value
            .Cells(3, 3).Value = m_wsCommonPrimary.Cells(1, 41).Value
            
            .Range("A:C").Columns.AutoFit
            
        End With
        
        With m_wsCommonPrimary
            .Activate
            .Range("A1").Select
            .Rows(4).AutoFilter
        End With
        
        wsAdditional.Activate: wsAdditional.Range("A1").Select
        
        m_wbPreviousCommon.Close SaveChanges:=False
        
    Else
        With m_wsCommonPrimary
            .Activate
            .Range("A1").Select
            .Rows(4).AutoFilter
        End With
    End If
    
    Application.CutCopyMode = False
    
    wbHeaders.Close SaveChanges:=False
    
    Lib.AddLogNote "Forming has finished"
    
    Exit Sub
    
ErrorHandler_CreateCommonTable:

    Call Lib.FatalError("Не удалось сформировать итоговой свод")
    Lib.DisableOptimization
    End
    
End Sub

Public Sub UpdateRegionalTable()
    
    On Error GoTo ErrorHandler_UpdateRegionalTable
    
    Lib.AddLogNote "Updating a regional table"
    
    ' Книга со сводом
    Dim wbRegional        As Workbook: Set wbRegional = Workbooks.Open(Filename:=frmMain.txtRegionalPath.Value, ReadOnly:=False)
    Dim wsRegionalPrimary As Worksheet: Set wsRegionalPrimary = wbRegional.Worksheets("свод")
    
    Dim nRegionalRows As Long: nRegionalRows = wsRegionalPrimary.Cells(wsRegionalPrimary.Rows.Count, 1).End(xlUp).Row
    Dim i, j, k       As Long
    
    Dim colFilteredIndices As New Collection
    Dim arrRegional()      As Variant: ReDim arrRegional(5 To nRegionalRows)
    
    Dim dictForm      As Object: Set dictForm = CreateObject("Scripting.Dictionary")
    Dim dictASR       As Object: Set dictASR = CreateObject("Scripting.Dictionary")
    Dim dictSplitting As Object: Set dictSplitting = CreateObject("Scripting.Dictionary")
    
    ' Проверка по Форме 7.4
    If Not frmMain.txtFormPath.Value = "" Then
        
        ' Loading 15%
        With frmLoading
            .SetInfo "Проверка по Форме 7.4..."
            .SetProgress 15
        End With
        
        Call Lib.Delay(2000)
        
        With wsRegionalPrimary
            .AutoFilterMode = False
            .Range("A:AA").AutoFilter Field:=6, Criteria1:="OEBS"
        End With
        
        ' Инициализирует массив транзакций
        For i = 5 To nRegionalRows
            
            arrRegional(i) = CStr(wsRegionalPrimary.Cells(i, 1).Value)
            
            If wsRegionalPrimary.Range("A" & i).EntireRow.Hidden = False Then
                
                colFilteredIndices.Add i
                
                If wsRegionalPrimary.Cells(i, 9).Value = "ок" Then
                    With wsRegionalPrimary
                        .Range(.Cells(i, 9), .Cells(i, 9)).Interior.Color = mcYellowColor
                    End With
                End If
                
            End If
            
        Next i
        
        Set m_wbForm = Workbooks.Open(Filename:=CStr(frmMain.txtFormPath.Value), ReadOnly:=True)
        Set m_wsFormPrimary = m_wbForm.Worksheets(1)
        
        Dim nFormRows As Long:    nFormRows = Lib.GetRowsCount(Column:=1, Sheet:=m_wsFormPrimary)
        Dim arrForm() As Variant: ReDim arrForm(2 To nFormRows)
        
        For i = 2 To nFormRows
            arrForm(i) = CStr(m_wsFormPrimary.Cells(i, 7).Value)
        Next i
        
        For i = 2 To nFormRows
            If Not dictForm.Exists(arrForm(i)) Then
                dictForm.Add arrForm(i), True
            End If
        Next i
        
        For Each i In colFilteredIndices
            If Not dictForm.Exists(arrRegional(i)) Then
                With wsRegionalPrimary
                    .Range(.Cells(i, 9), .Cells(i, 9)).Interior.Color = mcWhiteColor
                    .Cells(i, 9).Value = "ок"
                End With
            End If
        Next
        
        Set colFilteredIndices = Nothing
        
        m_wbForm.Close SaveChanges:=False
        
    End If
    
    ' Проверка по АСР Расчетам
    If Not frmMain.txtASRPath.Value = "" Then
        
        ' Loading 45%
        With frmLoading
            .SetInfo "Проверка по АСР Расчетам..."
            .SetProgress 45
        End With
        
        Call Lib.Delay(2000)
        
        With wsRegionalPrimary
            .AutoFilterMode = False
            .Range("A:AA").AutoFilter Field:=6, Criteria1:="АСР Расчеты"
        End With
        
        ' Инициализирует массив транзакций
        For i = 5 To nRegionalRows
            
            arrRegional(i) = CStr(wsRegionalPrimary.Cells(i, 1).Value)
            
            If wsRegionalPrimary.Range("A" & i).EntireRow.Hidden = False Then
                
                colFilteredIndices.Add i
            
                If wsRegionalPrimary.Cells(i, 9).Value = "ок" Then
                    With wsRegionalPrimary
                        .Range(.Cells(i, 9), .Cells(i, 9)).Interior.Color = mcYellowColor
                    End With
                End If
                
            End If
            
        Next i
        
        Set m_wbASR = Workbooks.Open(Filename:=CStr(frmMain.txtASRPath.Value), ReadOnly:=True)
        Set m_wsASRPrimary = m_wbASR.Worksheets(1)
        
        Dim nASRRows   As Long:    nASRRows = Lib.GetRowsCount(Column:=1, Sheet:=m_wsASRPrimary)
        Dim arrASR()   As Variant: ReDim arrASR(2 To nASRRows)
        
        For i = 2 To nASRRows
            arrASR(i) = CStr(m_wsASRPrimary.Cells(i, 4).Value)
        Next i
        
        For i = 2 To nASRRows
            If Not dictASR.Exists(arrASR(i)) Then
                dictASR.Add arrASR(i), True
            End If
        Next i
        
        For Each i In colFilteredIndices
            If Not dictASR.Exists(arrRegional(i)) Then
                With wsRegionalPrimary
                    .Range(.Cells(i, 9), .Cells(i, 9)).Interior.Color = mcWhiteColor
                    .Cells(i, 9).Value = "ок"
                End With
            End If
        Next
        
        Set colFilteredIndices = Nothing
        
        m_wbASR.Close SaveChanges:=False
        
    End If
    
    ' Проверка по раздельной предоплате
    If Not frmMain.txtSplittingPath.Value = "" Then
        
        ' Loading 75%
        With frmLoading
            .SetInfo "Проверка по раздельной предоплате..."
            .SetProgress 75
        End With
        
        Call Lib.Delay(2000)
        
        With wsRegionalPrimary
            .AutoFilterMode = False
            .Range("A:AA").AutoFilter Field:=6, Criteria1:="Кредиторы"
        End With
        
        ' Инициализирует массив транзакций
        For i = 5 To nRegionalRows
        
            arrRegional(i) = CStr(wsRegionalPrimary.Cells(i, 1).Value)
            
            If wsRegionalPrimary.Range("A" & i).EntireRow.Hidden = False Then
                
                colFilteredIndices.Add i
            
                If wsRegionalPrimary.Cells(i, 9).Value = "ок" Then
                    With wsRegionalPrimary
                        .Range(.Cells(i, 9), .Cells(i, 9)).Interior.Color = mcYellowColor
                    End With
                End If
            
            End If
            
        Next i
        
        Set m_wbSplitting = Workbooks.Open(Filename:=CStr(frmMain.txtSplittingPath.Value), ReadOnly:=True)
        Set m_wsSplittingPrimary = m_wbSplitting.Worksheets(1)
        
        Dim nSplittingRows   As Long:    nSplittingRows = Lib.GetRowsCount(Column:=1, Sheet:=m_wsSplittingPrimary)
        Dim arrSplitting()   As Variant: ReDim arrSplitting(14 To nSplittingRows)
        
        For i = 14 To nSplittingRows
            arrSplitting(i) = CStr(m_wsSplittingPrimary.Cells(i, 7).Value)
        Next i
        
        For i = 14 To nSplittingRows
            If Not dictSplitting.Exists(arrSplitting(i)) Then
                dictSplitting.Add arrSplitting(i), True
            End If
        Next i
        
        For Each i In colFilteredIndices
            If Not dictSplitting.Exists(arrRegional(i)) Then
                With wsRegionalPrimary
                    .Range(.Cells(i, 9), .Cells(i, 9)).Interior.Color = mcWhiteColor
                    .Cells(i, 9).Value = "ок"
                End With
            End If
        Next
        
        Set colFilteredIndices = Nothing
        
        m_wbSplitting.Close SaveChanges:=False
        
    End If
    
    ' Loading 99%
        With frmLoading
            .SetInfo ""
            .SetProgress 99
        End With
        
        Call Lib.Delay(2000)
    
    Application.CutCopyMode = False
    wsRegionalPrimary.AutoFilterMode = False
    
    ' Сбрасывает выделенный диапозон после работы кода
    With wsRegionalPrimary
        .Activate
        .Range("A1").Select
        .Rows(4).AutoFilter
    End With
    
    Lib.AddLogNote "Updating has finished"
    
    Exit Sub
    
ErrorHandler_UpdateRegionalTable:

    Call Lib.FatalError("Не удалось обновить сводный файл за предыдущий период")
    Lib.DisableOptimization
    End
    
End Sub
