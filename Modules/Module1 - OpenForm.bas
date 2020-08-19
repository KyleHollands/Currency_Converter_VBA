Attribute VB_Name = "Module1"
Option Explicit

Sub OpenForm()

Dim i As Integer
Dim DateArray As Variant
Dim currencies As Worksheet, importedCurrencies As Worksheet
Dim DateBox As Variant
Dim DateDay As Integer, DateMonth As Integer, DateYear As Integer
Dim url As String
Dim currImp As Variant, currImpTwo As Variant, currImpLabel As Variant, currCopy As Variant
Dim currImpCode As String, currImpName As String

Set currencies = Workbooks("Currency Converter_Hollands.xlsm").Worksheets("Currencies")
Set importedCurrencies = Workbooks("Currency Converter_Hollands.xlsm").Worksheets("Sheet1")

currencies.Select: Range("A1").Select

'EXTRACT DATE INFORMATION FROM DATEBOX VARIABLE

DateArray = Split(Now())
UserForm1.DateBox = DateArray(0)

DateDay = Day(UserForm1.DateBox)
DateMonth = Month(UserForm1.DateBox)
DateYear = Year(UserForm1.DateBox)

'IMPORT CURRENCY INFORMATION

    url = "URL;https://www.xe.com/currencytables/?from=USD&date=" & DateYear & "-" & DateMonth & "-" & DateDay

        With Worksheets("Sheet1").QueryTables.Add(Connection:=url, Destination:=Worksheets("Sheet1").Range("A1"))
            .Name = "My Query"
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .BackgroundQuery = False
            .RefreshStyle = xlOverwriteCells
            .SavePassword = False
            .SaveData = True
            .AdjustColumnWidth = True
            .RefreshPeriod = 0
            .WebSelectionType = xlEntirePage
            .WebFormatting = xlWebFormattingNone
            .WebPreFormattedTextToColumns = True
            .WebConsecutiveDelimitersAsOne = True
            .WebSingleBlockTextImport = False
            .WebDisableDateRecognition = False
            .WebDisableRedirections = False
            .Refresh BackgroundQuery:=False
        End With

'IMPORT/UPDATE CURRENCIES LIST INTO CURRENCIES SHEET

For Each currImp In importedCurrencies.Range("A:A")
    If Len(currImp) = 3 Then
        currImpCode = currImp.Value
        currImpName = currImp.Offset(0, 1).Value
            For Each currCopy In currencies.Range("A:A")
                If currCopy = currImpCode Then
                    Exit For
                ElseIf currCopy = "" Then
                    currCopy.Value = currImpCode
                    currCopy.Offset(0, 1).Value = currImpName
                    Exit For
                End If
            Next
    ElseIf currImp.Offset(1, 0).Value = "" And currImp.Offset(2, 0).Value = "" Then
        Exit For
    End If
Next

'POPULATE COMBOBOXES WITH CURRENCIES

For i = 1 To WorksheetFunction.CountA(Columns("A:A"))
    UserForm1.convFromBox.AddItem ActiveCell.Offset(i - 1, 0) & " - " & ActiveCell.Offset(i - 1, 1)
Next i

For i = 1 To WorksheetFunction.CountA(Columns("A:A"))
    UserForm1.convToBox.AddItem ActiveCell.Offset(i - 1, 0) & " - " & ActiveCell.Offset(i - 1, 1)
Next i

UserForm1.convFromBox.Text = Range("A1") & " - " & Range("B1")
UserForm1.convToBox.Text = Range("A1") & " - " & Range("B1")
UserForm1.Show

End Sub
