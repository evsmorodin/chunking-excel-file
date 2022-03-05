Attribute VB_Name = "ChunkingExcelFile"
Option Explicit

' Предполагается, что данные начинаются со второй строки, в первой - заголовок
' Предполагается, что в колонке "A" хранится код объекта, который является строкой
' Предполагается, что в колонке "B" хранится дата
' Для своих потребностей вы можете изменить значения переменных, либо убрать ненужные

Sub Chunk_File()
  Dim Limit As Long, Count As Long, CodeLength As String, SaveDir As String, LastDayOfMonth As Long, Period As Date, LastRow As Long, Iterator As Long, CellValue As String

  Period = ActiveWorkbook.ActiveSheet.Range("B2")
  'MsgBox TypeName(Period)
  'MsgBox "Период: " & Period

  LastDayOfMonth = Day(DateAdd("d", -1, DateAdd("m", 1, Period)))
  'MsgBox "Количество дней в месяце: " & LastDayOfMonth

  Count = 1: Limit = 300 * LastDayOfMonth + 1 ' Количество строк
  'MsgBox "Строк для каждого файла: " & Limit

  LastRow = Cells(1, 1).CurrentRegion.Rows.Count
  'MsgBox "Количество строк в файле: " & LastRow

  Range("A2:A" & LastRow).NumberFormat = "@"
  
  CodeLength = 10 ' Длина строки кода объекта
  Iterator = 2 ' Проходим по кодам объектов и добавляем ведущий ноль, если он не указан
  Do While Iterator <= LastRow
    CellValue = ActiveWorkbook.ActiveSheet.Range("A" & Iterator).Value
    If (Len(CellValue) < CodeLength) Then ActiveWorkbook.ActiveSheet.Range("A" & Iterator).Value = "0" & CellValue
    Iterator = Iterator + 1
  Loop

  SaveDir = ThisWorkbook.Path ' Или вписать полный путь для сохранения "C:\"
  Application.DisplayAlerts = False
  Do While Not IsEmpty(Cells(1, 1)) ' Предполагается, что в колонке A нет пустых ячеек
    ActiveSheet.Range("A2").Select ' Если в следующем блоке пустые ячейки, то заканчиваем цикл
    If IsEmpty(ActiveCell.Value) Then Exit Do
    Rows("2:" & Limit).Cut ' Если есть заголовок, заменить 1 на 2
    Workbooks.Add xlWBATWorksheet
    ActiveSheet.Paste: Cells(1, 1).Select
    Selection.EntireRow.Insert
    ActiveWorkbook.SaveAs Filename:=SaveDir & "\" & Format$(Period, "yyyy\-mm\-dd") & "_" & Count & ".xlsx", _
      FileFormat:=xlOpenXMLWorkbook
    ActiveWindow.Close
    Rows("2:" & Limit).Delete Shift:=xlUp ' Если есть заголовок, заменить 1 на 2
    Count = Count + 1
  Loop: MsgBox "Файл разбит на " & Count - 1 & " файл(ов). "
End Sub
