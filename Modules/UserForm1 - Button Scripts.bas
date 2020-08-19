VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   4005
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7785
   OleObjectBlob   =   "UserForm1 - Button Scripts.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TextBox1_Change()

End Sub

Private Sub CommandButton1_Click()

'IMPORT CURRENT CURRENCY DATA INTO SHEET 1
    
Dim DateDay As String, DateMonth As String, DateYear As String
Dim url As String
Dim Exists As Boolean
Dim currImp As Variant, currImpTwo As Variant, currImpLabel As Variant, currCopy As Variant
Dim initAmt As Double, convAmt As Double
Dim convFrom As String, convTo As String
Dim initAmtUnit As String, convToUnit As String
Dim testVar As String
Dim currentDate As String
Dim currYear As Integer

Set currencies = Workbooks("Currency Converter_Hollands.xlsm").Worksheets("Currencies")
Set importedCurrencies = Workbooks("Currency Converter_Hollands.xlsm").Worksheets("Sheet1")
    
Application.ScreenUpdating = False
importedCurrencies.Visible = True

'SAVE CURRENT YEAR

DateArray = Split(Now())
currentDate = DateArray(0)
currYear = Year(currentDate)

'ERROR CHECKING AND DATE ACQUISITION

If UserForm1.initAmtBox = "" Then MsgBox ("Input cannot be blank. "): GoTo Reset
If Not IsNumeric(UserForm1.initAmtBox) Then MsgBox ("Value must be a number."): GoTo Reset
If UserForm1.initAmtBox <= 0 Then MsgBox ("Please enter a positive number."): GoTo Reset
If DateBox = "" Then MsgBox ("Date cannot be empty."): GoTo Reset
If Not IsDate(DateBox) Then MsgBox ("Incorrect date format."): GoTo Reset

DateDay = Day(DateBox)
DateMonth = Month(DateBox)
DateYear = Year(DateBox)

If Len(DateDay) = 1 Then DateDay = "0" & DateDay
If Len(DateMonth) = 1 Then DateMonth = "0" & DateMonth
If Len(DateYear) < 4 Then MsgBox ("Incorrect year length."): GoTo Reset
If DateYear > currYear Then MsgBox ("Year cannot exceed current year."): GoTo Reset

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

'CONVERT SELECTED CURRENCY

    initAmt = UserForm1.initAmtBox.Value
    convFrom = Left(UserForm1.convFromBox, 3)
    convTo = Left(UserForm1.convToBox, 3)
    For Each currImp In importedCurrencies.Range("A:A")
        If currImp = convFrom Then
            initAmtUnit = currImp.Offset(0, 2).Value
            For Each currImpTwo In importedCurrencies.Range("A:A")
                If currImpTwo = convTo Then
                    convToUnit = currImpTwo.Offset(0, 2).Value
                    UserForm1.convAmtBox.Text = Round(initAmt * (convToUnit / initAmtUnit))
                    GoTo Reset:
                End If
            Next
        End If
    Next

Reset:
    
Sheets("Sheet1").Visible = False
Application.ScreenUpdating = True
        
End Sub

Private Sub CommandButton2_Click()

Unload UserForm1

End Sub

Private Sub CommandButton3_Click()

End Sub

Private Sub CommandButton4_Click()

Dim TodaysDate As String
Dim SelectedCurrency As String
Dim i As Integer
Dim url As String
Dim plotVar As Variant
Dim plotData As Worksheet, importedCurrencies As Worksheet
Dim currentDate As String

Application.ScreenUpdating = False

Set importedCurrencies = Workbooks("Currency Converter_Hollands.xlsm").Worksheets("Sheet1")
Set plotData = Workbooks("Currency Converter_Hollands.xlsm").Worksheets("PlotData")

plotData.Visible = True
importedCurrencies.Visible = True

If UserForm1.initAmtBox = "" Then MsgBox ("Input cannot be blank. "): GoTo ErrorReset
If Not IsNumeric(UserForm1.initAmtBox) Then MsgBox ("Value must be a number."): GoTo ErrorReset
If UserForm1.initAmtBox <= 0 Then MsgBox ("Please enter a positive number."): GoTo ErrorReset

plotData.Select: Range("A30").Select

TodaysDate = UserForm1.DateBox.Value
SelectedCurrency = Left(UserForm1.convFromBox, 3)

'SAVE CURRENT YEAR

DateArray = Split(Now())
currentDate = DateArray(0)
currYear = Year(currentDate)

'ERROR CHECKING AND DATE ACQUISITION

If UserForm1.initAmtBox = "" Then MsgBox ("Input cannot be blank. "): GoTo ErrorReset
If Not IsNumeric(UserForm1.initAmtBox) Then MsgBox ("Value must be a number."): GoTo ErrorReset
If UserForm1.initAmtBox <= 0 Then MsgBox ("Please enter a positive number."): GoTo ErrorReset
If DateBox = "" Then MsgBox ("Date cannot be empty."): GoTo ErrorReset
If Not IsDate(DateBox) Then MsgBox ("Incorrect date format."): GoTo ErrorReset

DateDay = Day(DateBox)
DateMonth = Month(DateBox)
DateYear = Year(DateBox)

If Len(DateDay) = 1 Then DateDay = "0" & DateDay
If Len(DateMonth) = 1 Then DateMonth = "0" & DateMonth
If Len(DateYear) < 4 Then MsgBox ("Incorrect year length."): GoTo ErrorReset
If DateYear > currYear Then MsgBox ("Year cannot exceed current year."): GoTo ErrorReset

'CALCULATE 30 DAYS PRIOR TO SELECTED DATE
For i = 30 To 1 Step -1
    Range("A" & 30 - i + 1) = DateAdd("d", -i + 1, TodaysDate)
Next i

'IMPORT CURRENCY INFORMATION BASED OFF DATE INFORMATION ABOVE

For Each plotVar In Sheets("PlotData").Range("A:A")
    
    DateDay = Day(plotVar)
    DateMonth = Month(plotVar)
    DateYear = Year(plotVar)
    
    If Len(DateDay) = 1 Then DateDay = "0" & DateDay
    If Len(DateMonth) = 1 Then DateMonth = "0" & DateMonth
    
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
    
'CONVERT SELECTED CURRENCY

    initAmt = UserForm1.initAmtBox.Value
    convFrom = Left(UserForm1.convFromBox, 3)
    convTo = Left(UserForm1.convToBox, 3)
    For Each currImp In importedCurrencies.Range("A:A")
        If currImp = convFrom Then
            initAmtUnit = currImp.Offset(0, 2).Value
            For Each currImpTwo In importedCurrencies.Range("A:A")
                If currImpTwo = convTo Then
                    convToUnit = currImpTwo.Offset(0, 2).Value
                    plotVar.Offset(0, 1) = initAmt * (convToUnit / initAmtUnit)
                    GoTo Reset:
                End If
            Next
        End If
    Next

Reset:

If plotVar.Offset(1, 0).Value = "" Then
    Exit For
End If

Next

'PLOT DATA ON A GRAPH

plotData.Range("A1:B30").Select
Chart6.SetSourceData Source:=Range("PlotData!$A$1:$B$30")
Chart6.ChartTitle.Select
ActiveChart.ChartTitle.Text = "Last 30 Days: " & UserForm1.convFromBox & " to " & UserForm1.convToBox

ErrorReset:

importedCurrencies.Visible = False
plotData.Visible = False
Worksheets("Currencies").Select
Application.ScreenUpdating = True

End Sub

Private Sub convAmtBox_Change()

End Sub

Private Sub convFromBox_Change()

End Sub

Private Sub dateBox_Change()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub UserForm_Click()

End Sub
