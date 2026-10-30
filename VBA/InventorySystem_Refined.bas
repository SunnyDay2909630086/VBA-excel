Option Explicit

'============================================================
' ?????????????????妫?
' ??????????????????????????????????????
'      ?????????????????????????????????????
'============================================================

Public Const BASE_SHEET As String = "???????"
Public Const SHEET_RK As String = "??????"
Public Const SHEET_XS As String = "?????"
Public Const SHEET_XSTH As String = "?????????"
Public Const SHEET_JHTH As String = "?????????"
Public Const SHEET_BS As String = "????"
Public Const SHEET_BY As String = "???绲?"
Public Const SHEET_KC As String = "??????"
Public Const SHEET_LR As String = "???????"

Public Sub SetupWorkbook()
    Dim ws As Worksheet

    Set ws = GetOrCreateSheet(BASE_SHEET)
    ws.Cells.Clear

    EnsureBaseDataStructure
    EnsureTransactionSheets
    EnsureInventorySummary
    EnsureProfitSummary
    EnsureNavigationSheet
    EnsureDateValidationList
    BuildMainNavigationButtons
    On Error Resume Next
    RefreshAllMasterDropdowns
    On Error GoTo 0
    MsgBox "??????????????????????????????????", vbInformation, "???????????"
End Sub

Private Sub EnsureNavigationSheet()
    Dim ws As Worksheet

    Set ws = GetOrCreateSheet("????")
    ws.Cells.Clear
    ws.Activate

    ws.Range("A1").Value = "???????????"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16
    ws.Range("A3").Value = "????????"
    ws.Range("A3").Font.Bold = True
    ws.Range("A5").Value = "???????"
    ws.Range("A5").Font.Bold = True
    ws.Range("A7").Value = "????????"
    ws.Range("A7").Font.Bold = True
    ws.Range("A9").Value = "??小???????"
    ws.Range("A9").Font.Bold = True
    ws.Range("A10").Value = "1. ??????/?????/???  2. ??????  3. ???????  4. ??????  5. ????????????"
    ws.Columns("A").ColumnWidth = 42
    ws.Range("A1:A10").VerticalAlignment = -4160
    ws.Range("A1:A10").HorizontalAlignment = -4131
    ws.Range("A1:H1").Interior.Color = RGB(31, 78, 121)
    ws.Range("A1:H1").Font.Color = RGB(255, 255, 255)
    ws.Range("A1:H1").RowHeight = 28
    ws.Range("A3:H3").Interior.Color = RGB(221, 235, 247)
    ws.Range("A5:H5").Interior.Color = RGB(252, 228, 214)
    ws.Range("A7:H7").Interior.Color = RGB(226, 239, 218)
    ws.Range("A9:H10").WrapText = True
End Sub

Private Sub BuildMainNavigationButtons()
    Dim ws As Worksheet

    Set ws = GetOrCreateSheet("????")

    AddNavigationButton ws, "btn_???????", "GoToSheetFromButton", 2, 3
    AddNavigationButton ws, "btn_???????", "GoToSheetFromButton", 2, 4
    AddNavigationButton ws, "btn_?????????", "GoToSheetFromButton", 2, 5
    AddNavigationButton ws, "btn_???????", "GoToSheetFromButton", 2, 6

    AddNavigationButton ws, "btn_???", "GoToSheetFromButton", 5, 3
    AddNavigationButton ws, "btn_?????", "GoToSheetFromButton", 5, 4
    AddNavigationButton ws, "btn_?????????", "GoToSheetFromButton", 5, 5
    AddNavigationButton ws, "btn_?????????", "GoToSheetFromButton", 5, 6
    AddNavigationButton ws, "btn_????", "GoToSheetFromButton", 5, 7
    AddNavigationButton ws, "btn_???绲?", "GoToSheetFromButton", 5, 8

    AddNavigationButton ws, "btn_??????", "GoToSheetFromButton", 8, 3
    AddNavigationButton ws, "btn_???????", "GoToSheetFromButton", 8, 4
    AddNavigationButton ws, "btn_?????", "GoToSheetFromButton", 8, 5
    AddNavigationButton ws, "btn_???", "GoToSheetFromButton", 8, 6
    AddNavigationButton ws, "btn_???????????", "LoadDemoData", 8, 7
    AddNavigationButton ws, "btn_???????", "RefreshFormDropdowns", 8, 8
End Sub

Private Sub AddNavigationButton(ByVal ws As Worksheet, ByVal buttonName As String, ByVal macroName As String, ByVal colIndex As Long, ByVal rowIndex As Long)
    Dim shp As Object
    Dim leftPos As Double
    Dim topPos As Double

    On Error Resume Next
    For Each shp In ws.Shapes
        If shp.Name = buttonName Then
            shp.Delete
        End If
    Next shp
    On Error GoTo 0

    leftPos = ws.Cells(rowIndex, colIndex).Left + 4
    topPos = ws.Cells(rowIndex, colIndex).Top + 4

    Set shp = ws.Shapes.AddShape(5, leftPos, topPos, 120, 26)
    shp.Name = buttonName
    shp.TextFrame.Characters.Text = GetButtonText(buttonName)
    shp.TextFrame.Characters.Font.Bold = True
    shp.Fill.ForeColor.RGB = RGB(31, 78, 121)
    shp.Line.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    shp.OnAction = macroName
End Sub

Private Sub AddFormButton(ByVal ws As Worksheet, ByVal buttonName As String, ByVal macroName As String, ByVal colIndex As Long, ByVal rowIndex As Long)
    Dim shp As Object
    Dim leftPos As Double
    Dim topPos As Double

    On Error Resume Next
    ws.Shapes(buttonName).Delete
    On Error GoTo 0

    leftPos = ws.Cells(rowIndex, colIndex).Left + 4
    topPos = ws.Cells(rowIndex, colIndex).Top + 4

    Set shp = ws.Shapes.AddShape(5, leftPos, topPos, 90, 26)
    shp.Name = buttonName
    shp.TextFrame.Characters.Text = GetFormButtonText(buttonName)
    shp.TextFrame.Characters.Font.Bold = True
    shp.Fill.ForeColor.RGB = RGB(237, 125, 49)
    shp.Line.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    shp.OnAction = macroName
End Sub

Private Function GetFormButtonText(ByVal buttonName As String) As String
    If InStr(1, buttonName, "???", 1) > 0 Then
        GetFormButtonText = "???"
    Else
        GetFormButtonText = "????"
    End If
End Function

Public Sub ClearReceiptEntry()
    Dim ws As Worksheet
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_RK)
    ws.Range("B2:C2").ClearContents
    ws.Range("E2:F2").ClearContents
    ws.Range("B5:G50").ClearContents
    For i = 5 To 50
        ws.Cells(i, 1).Value = i - 4
    Next i
End Sub

Public Sub ClearSalesEntry()
    ClearEntryRows ThisWorkbook.Worksheets(SHEET_XS), "I"
End Sub

Public Sub ClearSalesReturnEntry()
    ClearEntryRows ThisWorkbook.Worksheets(SHEET_XSTH), "I"
End Sub

Public Sub ClearPurchaseReturnEntry()
    ClearEntryRows ThisWorkbook.Worksheets(SHEET_JHTH), "I"
End Sub

Public Sub ClearLossEntry()
    ClearEntryRows ThisWorkbook.Worksheets(SHEET_BS), "F"
End Sub

Public Sub ClearOverflowEntry()
    ClearEntryRows ThisWorkbook.Worksheets(SHEET_BY), "F"
End Sub

Public Sub AddEntryRow()
    Dim ws As Worksheet
    Dim sourceRow As Long
    Dim targetRow As Long
    Dim lastColumn As String

    Set ws = ActiveSheet
    If ws.Name = SHEET_BS Or ws.Name = SHEET_BY Then
        lastColumn = "F"
    Else
        lastColumn = "I"
    End If

    sourceRow = FindLastEntryRow(ws, lastColumn)
    targetRow = sourceRow + 1
    If targetRow > 50 Then
        MsgBox "????????????? 49 ?小?", vbExclamation, "????????"
        Exit Sub
    End If

    ws.Range("A" & sourceRow & ":" & lastColumn & sourceRow).Copy
    ws.Range("A" & targetRow & ":" & lastColumn & targetRow).PasteSpecial -4122
    ws.Range("A" & targetRow & ":" & lastColumn & targetRow).ClearContents
    ws.Range("A" & targetRow).Value = targetRow - 1
    If lastColumn = "I" Then
        ws.Range("B" & targetRow).Value = ws.Range("B2").Value
        ws.Range("C" & targetRow).Value = ws.Range("C2").Value
    Else
        ws.Range("B" & targetRow).Value = ws.Range("B2").Value
    End If
    Application.CutCopyMode = False
End Sub

Public Sub RemoveEntryRow()
    Dim ws As Worksheet
    Dim lastColumn As String
    Dim lastRow As Long

    Set ws = ActiveSheet
    If ws.Name = SHEET_BS Or ws.Name = SHEET_BY Then
        lastColumn = "F"
    Else
        lastColumn = "I"
    End If
    lastRow = FindLastEntryRow(ws, lastColumn)
    If lastRow <= 2 Then
        MsgBox "??????????????小?", vbInformation, "????????"
        Exit Sub
    End If
    ws.Range("A" & lastRow & ":" & lastColumn & lastRow).ClearContents
End Sub

Public Sub PrintCurrentDocument()
    Dim ws As Worksheet
    Dim printLastRow As Long

    Set ws = ActiveSheet
    If ws.Name = SHEET_RK Then
        printLastRow = 13
        ws.PageSetup.PrintArea = "$A$1:$G$" & printLastRow
    Else
        printLastRow = FindLastEntryRow(ws, "I") + 2
        ws.PageSetup.PrintArea = "$A$1:$R$" & printLastRow
    End If
    ws.PrintPreview
End Sub

Private Function FindLastEntryRow(ByVal ws As Worksheet, ByVal lastColumn As String) As Long
    Dim lastRow As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    If lastRow < 2 Then lastRow = 2
    If lastRow > 50 Then lastRow = 50
    FindLastEntryRow = lastRow
End Function

Private Function GetButtonText(ByVal buttonName As String) As String
    Select Case buttonName
        Case "btn_???????": GetButtonText = "???????"
        Case "btn_???????": GetButtonText = "???????"
        Case "btn_?????????": GetButtonText = "?????????"
        Case "btn_???????": GetButtonText = "???????"
        Case "btn_???": GetButtonText = "??????"
        Case "btn_?????": GetButtonText = "?????"
        Case "btn_?????????": GetButtonText = "?????????"
        Case "btn_?????????": GetButtonText = "?????????"
        Case "btn_????": GetButtonText = "????"
        Case "btn_???绲?": GetButtonText = "???绲?"
        Case "btn_??????": GetButtonText = "??????"
        Case "btn_???????": GetButtonText = "???????"
        Case "btn_?????": GetButtonText = "?????"
        Case "btn_???": GetButtonText = "???"
        Case "btn_???????????": GetButtonText = "???????????"
        Case "btn_???????": GetButtonText = "???????"
        Case Else: GetButtonText = buttonName
    End Select
End Function

Public Sub LoadDemoData()
    Dim ws As Worksheet

    If Not WorksheetExists(BASE_SHEET) Or Not TableExistsOnSheet(BASE_SHEET, "tbl_??????") Then
        MsgBox "???????? SetupWorkbook??", vbExclamation, "???????"
        Exit Sub
    End If

    Set ws = ThisWorkbook.Worksheets(BASE_SHEET)
    EnsureDemoMasterData

    If Not DemoDocumentExists(SHEET_RK, "tbl_??????", "RK-DEMO-001") Then
        PrepareReceiptDemo
        SaveReceipt
    End If
    If Not DemoDocumentExists(SHEET_XS, "tbl_???????", "XS-DEMO-001") Then
        PrepareSalesDemo
        SaveSales
    End If
    If Not DemoDocumentExists(SHEET_XSTH, "tbl_??????????", "XSTH-DEMO-001") Then
        PrepareSalesReturnDemo
        SaveSalesReturn
    End If
    If Not DemoDocumentExists(SHEET_JHTH, "tbl_??????????", "JHTH-DEMO-001") Then
        PreparePurchaseReturnDemo
        SavePurchaseReturn
    End If
    If Not DemoDocumentExists(SHEET_BS, "tbl_???????", "BS-DEMO-001") Then
        PrepareLossDemo
        SaveLoss
    End If
    If Not DemoDocumentExists(SHEET_BY, "tbl_???????", "BY-DEMO-001") Then
        PrepareOverflowDemo
        SaveOverflow
    End If
    LoadAdditionalDemoTransactions

    RefreshInventorySummary
    RefreshProfitSummary
    MsgBox "???????????????????????????????", vbInformation, "???????"
End Sub

Private Sub EnsureDemoMasterData()
    EnsureDemoProduct "??????A", "??", "???????"
    EnsureDemoProduct "??????B", "??", "???????"
    EnsureDemoProduct "??????C", "?", "???????"
    EnsureDemoProduct "??????D", "??", "???????"
    EnsureDemoSupplier "????????", "??????", "???????"
    EnsureDemoSupplier "???霉????", "??????", "???????"
    EnsureDemoCustomer "??????", "????", "???????"
    EnsureDemoCustomer "???????", "??????", "???????"
End Sub

Private Sub EnsureDemoProduct(ByVal productName As String, ByVal unit As String, ByVal note As String)
    Dim tbl As ListObject
    Dim i As Long

    Set tbl = ThisWorkbook.Worksheets(BASE_SHEET).ListObjects("tbl_??????")
    For i = 1 To tbl.ListRows.Count
        If StrComp(CStr(tbl.DataBodyRange(i, 1).Value), productName, 1) = 0 Then Exit Sub
    Next i
    With tbl.ListRows.Add
        .Range(1, 1).Value = productName
        .Range(1, 2).Value = unit
        .Range(1, 3).Value = note
    End With
End Sub

Private Sub EnsureDemoSupplier(ByVal supplierName As String, ByVal contact As String, ByVal note As String)
    Dim tbl As ListObject
    Dim i As Long

    Set tbl = ThisWorkbook.Worksheets(BASE_SHEET).ListObjects("tbl_????????")
    For i = 1 To tbl.ListRows.Count
        If StrComp(CStr(tbl.DataBodyRange(i, 1).Value), supplierName, 1) = 0 Then Exit Sub
    Next i
    With tbl.ListRows.Add
        .Range(1, 1).Value = supplierName
        .Range(1, 2).Value = contact
        .Range(1, 3).Value = note
    End With
End Sub

Private Sub EnsureDemoCustomer(ByVal customerName As String, ByVal contact As String, ByVal note As String)
    Dim tbl As ListObject
    Dim i As Long

    Set tbl = ThisWorkbook.Worksheets(BASE_SHEET).ListObjects("tbl_??????")
    For i = 1 To tbl.ListRows.Count
        If StrComp(CStr(tbl.DataBodyRange(i, 1).Value), customerName, 1) = 0 Then Exit Sub
    Next i
    With tbl.ListRows.Add
        .Range(1, 1).Value = customerName
        .Range(1, 2).Value = contact
        .Range(1, 3).Value = note
    End With
End Sub

Private Function DemoDocumentExists(ByVal sheetName As String, ByVal tableName As String, ByVal docNo As String) As Boolean
    Dim tbl As ListObject
    Dim i As Long

    Set tbl = GetTableIfExists(sheetName, tableName)
    If tbl Is Nothing Or tbl.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To tbl.DataBodyRange.Rows.Count
        If StrComp(CStr(tbl.DataBodyRange(i, 2).Value), docNo, 1) = 0 Then
            DemoDocumentExists = True
            Exit Function
        End If
    Next i
End Function

Private Sub LoadAdditionalDemoTransactions()
    If Not DemoDocumentExists(SHEET_RK, "tbl_??????", "RK-DEMO-002") Then
        PrepareAdditionalReceiptDemo "RK-DEMO-002", "???霉????", "??????B", 30, "??", 80
        SaveReceipt
    End If
    If Not DemoDocumentExists(SHEET_RK, "tbl_??????", "RK-DEMO-003") Then
        PrepareAdditionalReceiptDemo "RK-DEMO-003", "????????", "??????C", 60, "?", 6
        SaveReceipt
    End If
    If Not DemoDocumentExists(SHEET_XS, "tbl_???????", "XS-DEMO-002") Then
        PrepareAdditionalSalesDemo "XS-DEMO-002", "???????", "??????B", 8, "??", 120
        SaveSales
    End If
    If Not DemoDocumentExists(SHEET_XS, "tbl_???????", "XS-DEMO-003") Then
        PrepareAdditionalSalesDemo "XS-DEMO-003", "??????", "??????C", 15, "?", 10
        SaveSales
    End If
End Sub

Private Sub PrepareAdditionalReceiptDemo(ByVal docNo As String, ByVal supplierName As String, ByVal productName As String, ByVal qty As Double, ByVal unit As String, ByVal price As Double)
    With ThisWorkbook.Worksheets(SHEET_RK)
        .Range("B2:C2").ClearContents
        .Range("E2:F2").ClearContents
        .Range("B5:G50").ClearContents
        .Range("B2").Value = supplierName
        .Range("E2").Value = Date
        .Range("I2").Value = docNo
        .Range("B5").Value = productName
        .Range("C5").Value = qty
        .Range("D5").Value = unit
        .Range("E5").Value = price
    End With
End Sub

Private Sub PrepareAdditionalSalesDemo(ByVal docNo As String, ByVal customerName As String, ByVal productName As String, ByVal qty As Double, ByVal unit As String, ByVal price As Double)
    With ThisWorkbook.Worksheets(SHEET_XS)
        .Range("A2:I2").ClearContents
        .Range("A2").Value = docNo
        .Range("B2").Value = Date
        .Range("C2").Value = customerName
        .Range("D2").Value = productName
        .Range("E2").Value = qty
        .Range("F2").Value = unit
        .Range("G2").Value = price
    End With
End Sub

Private Sub PrepareReceiptDemo()
    With ThisWorkbook.Worksheets(SHEET_RK)
        .Range("B2:C2").ClearContents
        .Range("E2:F2").ClearContents
        .Range("B5:G50").ClearContents
        .Range("B2").Value = "????????"
        .Range("E2").Value = Date
        .Range("I2").Value = "RK-DEMO-001"
        .Range("B5").Value = "??????A"
        .Range("C5").Value = 100
        .Range("D5").Value = "??"
        .Range("E5").Value = 10
    End With
End Sub

Private Sub PrepareSalesDemo()
    With ThisWorkbook.Worksheets(SHEET_XS)
        .Range("A2:I2").ClearContents
        .Range("A2").Value = "XS-DEMO-001"
        .Range("B2").Value = Date
        .Range("C2").Value = "??????"
        .Range("D2").Value = "??????A"
        .Range("E2").Value = 20
        .Range("F2").Value = "??"
        .Range("G2").Value = 16
    End With
End Sub

Private Sub PrepareSalesReturnDemo()
    With ThisWorkbook.Worksheets(SHEET_XSTH)
        .Range("A2:I2").ClearContents
        .Range("A2").Value = "XSTH-DEMO-001"
        .Range("B2").Value = Date
        .Range("C2").Value = "??????"
        .Range("D2").Value = "??????A"
        .Range("E2").Value = 2
        .Range("F2").Value = "??"
        .Range("G2").Value = 16
    End With
End Sub

Private Sub PreparePurchaseReturnDemo()
    With ThisWorkbook.Worksheets(SHEET_JHTH)
        .Range("A2:I2").ClearContents
        .Range("A2").Value = "JHTH-DEMO-001"
        .Range("B2").Value = Date
        .Range("C2").Value = "????????"
        .Range("D2").Value = "??????A"
        .Range("E2").Value = 5
        .Range("F2").Value = "??"
        .Range("G2").Value = 10
    End With
End Sub

Private Sub PrepareLossDemo()
    With ThisWorkbook.Worksheets(SHEET_BS)
        .Range("A2:F2").ClearContents
        .Range("A2").Value = "BS-DEMO-001"
        .Range("B2").Value = Date
        .Range("C2").Value = "??????A"
        .Range("D2").Value = 3
        .Range("E2").Value = "??"
    End With
End Sub

Private Sub PrepareOverflowDemo()
    With ThisWorkbook.Worksheets(SHEET_BY)
        .Range("A2:F2").ClearContents
        .Range("A2").Value = "BY-DEMO-001"
        .Range("B2").Value = Date
        .Range("C2").Value = "??????A"
        .Range("D2").Value = 1
        .Range("E2").Value = "??"
    End With
End Sub

Public Sub GoToSheetFromButton()
    Dim callerName As String
    Dim targetSheet As String

    callerName = Application.Caller
    targetSheet = GetButtonTargetSheet(callerName)

    If targetSheet = "" Then Exit Sub
    If WorksheetExists(targetSheet) Then
        ThisWorkbook.Worksheets(targetSheet).Activate
    Else
        Select Case targetSheet
            Case "???????": EnsureBaseDataStructure
            Case SHEET_RK: EnsureReceiptSheet
            Case SHEET_XS: EnsureSalesSheet
            Case SHEET_XSTH: EnsureSalesReturnSheet
            Case SHEET_JHTH: EnsurePurchaseReturnSheet
            Case SHEET_BS: EnsureLossSheet
            Case SHEET_BY: EnsureOverflowSheet
            Case SHEET_KC: EnsureInventorySummary
            Case SHEET_LR: EnsureProfitSummary
        End Select
        ThisWorkbook.Worksheets(targetSheet).Activate
    End If

    Select Case callerName
        Case "btn_???": RunSystemSelfCheck
        Case "btn_?????": RefreshInventorySummary
        Case "btn_???????": RefreshProfitSummary
    End Select
End Sub

Private Function GetButtonTargetSheet(ByVal callerName As String) As String
    Select Case callerName
        Case "btn_???????": GetButtonTargetSheet = BASE_SHEET
        Case "btn_???????", "btn_?????????", "btn_???????": GetButtonTargetSheet = BASE_SHEET
        Case "btn_???": GetButtonTargetSheet = SHEET_RK
        Case "btn_?????": GetButtonTargetSheet = SHEET_XS
        Case "btn_?????????": GetButtonTargetSheet = SHEET_XSTH
        Case "btn_?????????": GetButtonTargetSheet = SHEET_JHTH
        Case "btn_????": GetButtonTargetSheet = SHEET_BS
        Case "btn_???绲?": GetButtonTargetSheet = SHEET_BY
        Case "btn_??????": GetButtonTargetSheet = SHEET_KC
        Case "btn_???????": GetButtonTargetSheet = SHEET_LR
        Case "btn_???": GetButtonTargetSheet = "????"
        Case "btn_?????": GetButtonTargetSheet = SHEET_KC
        Case Else: GetButtonTargetSheet = ""
    End Select
End Function

Public Sub RunSystemSelfCheck()
    Dim msg As String

    If Not WorksheetExists(BASE_SHEET) Then msg = msg & "?????????????????" & vbCrLf
    If Not WorksheetExists(SHEET_RK) Then msg = msg & "?????????????" & vbCrLf
    If Not WorksheetExists(SHEET_XS) Then msg = msg & "???????????????" & vbCrLf
    If Not WorksheetExists(SHEET_XSTH) Then msg = msg & "???????????????????" & vbCrLf
    If Not WorksheetExists(SHEET_JHTH) Then msg = msg & "???????????????????" & vbCrLf
    If Not WorksheetExists(SHEET_BS) Then msg = msg & "??????????????" & vbCrLf
    If Not WorksheetExists(SHEET_BY) Then msg = msg & "?????????????绲?" & vbCrLf
    If Not WorksheetExists(SHEET_KC) Then msg = msg & "????????????????" & vbCrLf
    If Not WorksheetExists(SHEET_LR) Then msg = msg & "?????????????????" & vbCrLf

    If Not TableExistsOnSheet(BASE_SHEET, "tbl_??????") Then msg = msg & "??????????tbl_??????" & vbCrLf
    If Not TableExistsOnSheet(BASE_SHEET, "tbl_????????") Then msg = msg & "??????????tbl_????????" & vbCrLf
    If Not TableExistsOnSheet(BASE_SHEET, "tbl_??????") Then msg = msg & "??????????tbl_??????" & vbCrLf
    If Not TableExistsOnSheet(SHEET_RK, "tbl_??????") Then msg = msg & "??????????tbl_??????" & vbCrLf
    If Not TableExistsOnSheet(SHEET_XS, "tbl_???????") Then msg = msg & "??????????tbl_???????" & vbCrLf
    If Not TableExistsOnSheet(SHEET_XSTH, "tbl_??????????") Then msg = msg & "??????????tbl_??????????" & vbCrLf
    If Not TableExistsOnSheet(SHEET_JHTH, "tbl_??????????") Then msg = msg & "??????????tbl_??????????" & vbCrLf
    If Not TableExistsOnSheet(SHEET_BS, "tbl_???????") Then msg = msg & "??????????tbl_???????" & vbCrLf
    If Not TableExistsOnSheet(SHEET_BY, "tbl_???????") Then msg = msg & "??????????tbl_???????" & vbCrLf
    If Not TableExistsOnSheet(SHEET_KC, "tbl_??????") Then msg = msg & "??????????tbl_??????" & vbCrLf
    If Not TableExistsOnSheet(SHEET_LR, "tbl_???????") Then msg = msg & "??????????tbl_???????" & vbCrLf

    If msg = "" Then
        MsgBox "?????????????????????????????????????", vbInformation, "???????????"
    Else
        MsgBox "?????????????" & vbCrLf & msg, vbExclamation, "???????????"
    End If
End Sub

Private Sub EnsureBaseDataStructure()
    Dim ws As Worksheet
    Dim tbl As ListObject

    Set ws = GetOrCreateSheet(BASE_SHEET)
    ws.Cells.Clear

    ws.Range("A1:C1").Value = Array("???????", "??位", "???")
    ws.Range("E1:G1").Value = Array("?????????", "?????", "???")
    ws.Range("J1:L1").Value = Array("???????", "?????", "???")

    ws.Range("A1:C1").Font.Bold = True
    ws.Range("E1:G1").Font.Bold = True
    ws.Range("J1:L1").Font.Bold = True

    Set tbl = EnsureTable(ws, "tbl_??????", "A1:C1")
    tbl.Name = "tbl_??????"

    Set tbl = EnsureTable(ws, "tbl_????????", "E1:G1")
    tbl.Name = "tbl_????????"

    Set tbl = EnsureTable(ws, "tbl_??????", "J1:L1")
    tbl.Name = "tbl_??????"

    ws.Columns("A:L").AutoFit
        ApplyMasterSheetStyle ws
End Sub

Private Sub EnsureTransactionSheets()
    EnsureReceiptSheet
    EnsureSalesSheet
    EnsureSalesReturnSheet
    EnsurePurchaseReturnSheet
    EnsureLossSheet
    EnsureOverflowSheet
End Sub

Private Sub EnsureReceiptSheet()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long

    Set ws = GetOrCreateSheet(SHEET_RK)
    ws.Cells.Clear
    ws.Range("A1:G1").Merge
    ws.Range("A1").Value = "??????"
    ws.Range("A2").Value = "??????位??"
    ws.Range("D2").Value = "?????"
    ws.Range("E2:F2").Merge
    ws.Range("B2").Value = ""
    ws.Range("E2").Value = Date
    ws.Range("E2").NumberFormat = "yyyy/m/d"
    ws.Range("I2").Value = "RK-0001"
    ws.Range("I2").Font.Color = RGB(255, 255, 255)
    ws.Range("A4:G4").Value = Array("???", "???????", "????", "??位", "????????", "??????", "???")
    For i = 5 To 13
        ws.Cells(i, 1).Value = i - 4
    Next i
    ws.Range("K1:R1").Value = Array("????", "?????", "??????位", "???????", "????", "??位", "????", "???")
    ws.Range("A1:G1").Font.Bold = True
    ws.Range("A4:G4").Font.Bold = True
    ws.Range("K1:R1").Font.Bold = True

    Set tbl = EnsureTable(ws, "tbl_??????", "K1:R1")
    tbl.Name = "tbl_??????"

    ws.Columns("A:G").ColumnWidth = 14
    ws.Columns("A").ColumnWidth = 8
    ws.Columns("B").ColumnWidth = 24
    ws.Columns("G").ColumnWidth = 20
    ws.Columns("I:R").ColumnWidth = 12
    ApplySafeValidation ws.Range("B2"), "tbl_????????", "?????????"
    ApplySafeValidation ws.Range("B5:B50"), "tbl_??????", "???????"
    ws.Range("D5:D50").Interior.Color = RGB(242, 242, 242)
    ApplyDateValidation ws.Range("E2")
    AddFormButton ws, "btn_?????", "SaveReceipt", 11, 3
    AddFormButton ws, "btn_??????", "ClearReceiptEntry", 13, 3
    AddFormButton ws, "btn_?????????", "AddEntryRow", 20, 2
    AddFormButton ws, "btn_????????", "RemoveEntryRow", 20, 3
    AddFormButton ws, "btn_?????", "PrintCurrentDocument", 20, 4
    ApplyReceiptFormStyle ws
End Sub

Public Sub RefreshFormDropdowns()
    RefreshDropdownSources
    ApplySafeValidation ThisWorkbook.Worksheets(SHEET_RK).Range("B2"), "tbl_????????", "?????????"
    ApplySafeValidation ThisWorkbook.Worksheets(SHEET_RK).Range("B5:B50"), "tbl_??????", "???????"
    ApplyDateValidation ThisWorkbook.Worksheets(SHEET_RK).Range("E2")
    AutoFillReceiptUnits
    MsgBox "?????斜???????????????????????位????????????????", vbInformation, "?????斜?"
End Sub

Public Sub RefreshDropdownSources()
    RefreshAllMasterDropdowns
End Sub

Public Sub AutoFillReceiptUnits()
    Dim ws As Worksheet
    Dim i As Long
    Dim productName As String

    Set ws = ThisWorkbook.Worksheets(SHEET_RK)
    For i = 5 To 50
        productName = Trim$(CStr(ws.Cells(i, 2).Value))
        If productName <> "" Then
            ws.Cells(i, 4).Value = GetProductUnit(productName, "")
        Else
            ws.Cells(i, 4).ClearContents
        End If
    Next i
End Sub

Private Sub EnsureSalesSheet()
    Dim ws As Worksheet
    Dim tbl As ListObject

    Set ws = GetOrCreateSheet(SHEET_XS)
    ws.Cells.Clear
    ws.Range("A1:I1").Value = Array("?????", "????", "???", "???????", "????", "??位", "???????", "???", "???")
    ws.Range("K1:R1").Value = Array("????", "?????", "???", "???????", "????", "??位", "???????", "???")
    ws.Range("A1:I1").Font.Bold = True
    ws.Range("K1:R1").Font.Bold = True

    Set tbl = EnsureTable(ws, "tbl_???????", "K1:R1")
    tbl.Name = "tbl_???????"

    ws.Range("A2").Value = "XS-0001"
    ws.Range("B2").Value = Date
    ws.Range("C2").Value = ""
    ws.Range("D2").Value = ""
    ws.Range("E2").Value = 0
    ws.Range("F2").Value = ""
    ws.Range("G2").Value = 0
    ws.Range("H2").Value = 0

    ws.Columns("A:R").AutoFit
    ApplySafeValidation ws.Range("C2:C50"), "tbl_??????", "???????"
    ApplySafeValidation ws.Range("D2:D50"), "tbl_??????", "???????"
    ApplySafeValidation ws.Range("F2:F50"), "tbl_??????", "??位"
    AddFormButton ws, "btn_???????", "SaveSales", 11, 3
    AddFormButton ws, "btn_???????", "ClearSalesEntry", 13, 3
    AddFormButton ws, "btn_??????????", "AddEntryRow", 20, 2
    AddFormButton ws, "btn_?????????", "RemoveEntryRow", 20, 3
    AddFormButton ws, "btn_??????", "PrintCurrentDocument", 20, 4
        ApplyEntrySheetStyle ws, "A1:I2", "K1:R1", "R"
End Sub

Private Sub EnsureSalesReturnSheet()
    Dim ws As Worksheet
    Dim tbl As ListObject

    Set ws = GetOrCreateSheet(SHEET_XSTH)
    ws.Cells.Clear
    ws.Range("A1:I1").Value = Array("?????", "????", "???", "???????", "????", "??位", "???????", "???", "???")
    ws.Range("K1:R1").Value = Array("????", "?????", "???", "???????", "????", "??位", "???????", "???")
    ws.Range("A1:I1").Font.Bold = True
    ws.Range("K1:R1").Font.Bold = True

    Set tbl = EnsureTable(ws, "tbl_??????????", "K1:R1")
    tbl.Name = "tbl_??????????"

    ws.Range("A2").Value = "XSTH-0001"
    ws.Range("B2").Value = Date
    ws.Range("C2").Value = ""
    ws.Range("D2").Value = ""
    ws.Range("E2").Value = 0
    ws.Range("F2").Value = ""
    ws.Range("G2").Value = 0
    ws.Range("H2").Value = 0

    ws.Columns("A:R").AutoFit
    ApplySafeValidation ws.Range("C2:C50"), "tbl_??????", "???????"
    ApplySafeValidation ws.Range("D2:D50"), "tbl_??????", "???????"
    ApplySafeValidation ws.Range("F2:F50"), "tbl_??????", "??位"
    AddFormButton ws, "btn_???????????", "SaveSalesReturn", 11, 3
    AddFormButton ws, "btn_??????????", "ClearSalesReturnEntry", 13, 3
    AddFormButton ws, "btn_?????????????", "AddEntryRow", 20, 2
    AddFormButton ws, "btn_?????????????", "RemoveEntryRow", 20, 3
    AddFormButton ws, "btn_??????????", "PrintCurrentDocument", 20, 4
        ApplyEntrySheetStyle ws, "A1:I2", "K1:R1", "R"
End Sub

Private Sub EnsurePurchaseReturnSheet()
    Dim ws As Worksheet
    Dim tbl As ListObject

    Set ws = GetOrCreateSheet(SHEET_JHTH)
    ws.Cells.Clear
    ws.Range("A1:I1").Value = Array("?????", "????", "?????", "???????", "????", "??位", "???????", "???", "???")
    ws.Range("K1:R1").Value = Array("????", "?????", "?????", "???????", "????", "??位", "???????", "???")
    ws.Range("A1:I1").Font.Bold = True
    ws.Range("K1:R1").Font.Bold = True

    Set tbl = EnsureTable(ws, "tbl_??????????", "K1:R1")
    tbl.Name = "tbl_??????????"

    ws.Range("A2").Value = "JHTH-0001"
    ws.Range("B2").Value = Date
    ws.Range("C2").Value = ""
    ws.Range("D2").Value = ""
    ws.Range("E2").Value = 0
    ws.Range("F2").Value = ""
    ws.Range("G2").Value = 0
    ws.Range("H2").Value = 0

    ws.Columns("A:R").AutoFit
    ApplySafeValidation ws.Range("C2:C50"), "tbl_????????", "?????????"
    ApplySafeValidation ws.Range("D2:D50"), "tbl_??????", "???????"
    ApplySafeValidation ws.Range("F2:F50"), "tbl_??????", "??位"
    AddFormButton ws, "btn_???????????", "SavePurchaseReturn", 11, 3
    AddFormButton ws, "btn_??????????", "ClearPurchaseReturnEntry", 13, 3
    AddFormButton ws, "btn_?????????????", "AddEntryRow", 20, 2
    AddFormButton ws, "btn_?????????????", "RemoveEntryRow", 20, 3
    AddFormButton ws, "btn_??????????", "PrintCurrentDocument", 20, 4
    ApplyEntrySheetStyle ws, "A1:I2", "K1:R1", "R"
End Sub

Private Sub EnsureLossSheet()
    Dim ws As Worksheet
    Dim tbl As ListObject

    Set ws = GetOrCreateSheet(SHEET_BS)
    ws.Cells.Clear
    ws.Range("A1:F1").Value = Array("?????", "????", "???????", "????", "??位", "???")
    ws.Range("H1:M1").Value = Array("????", "?????", "???????", "????", "??位", "???")
    ws.Range("A1:F1").Font.Bold = True
    ws.Range("H1:M1").Font.Bold = True

    Set tbl = EnsureTable(ws, "tbl_???????", "H1:M1")
    tbl.Name = "tbl_???????"

    ws.Range("A2").Value = "BS-0001"
    ws.Range("B2").Value = Date
    ws.Range("C2").Value = ""
    ws.Range("D2").Value = 0
    ws.Range("E2").Value = ""

    ws.Columns("A:M").AutoFit
    ApplySafeValidation ws.Range("C2:C50"), "tbl_??????", "???????"
    ApplySafeValidation ws.Range("E2:E50"), "tbl_??????", "??位"
    AddFormButton ws, "btn_?????", "SaveLoss", 8, 3
    AddFormButton ws, "btn_???????", "ClearLossEntry", 10, 3
    AddFormButton ws, "btn_??????????", "AddEntryRow", 15, 2
    AddFormButton ws, "btn_?????????", "RemoveEntryRow", 15, 3
    AddFormButton ws, "btn_??????", "PrintCurrentDocument", 15, 4
    ApplyEntrySheetStyle ws, "A1:F2", "H1:M1", "M"
End Sub

Private Sub EnsureOverflowSheet()
    Dim ws As Worksheet
    Dim tbl As ListObject

    Set ws = GetOrCreateSheet(SHEET_BY)
    ws.Cells.Clear
    ws.Range("A1:F1").Value = Array("?????", "????", "???????", "????", "??位", "???")
    ws.Range("H1:M1").Value = Array("????", "?????", "???????", "????", "??位", "???")
    ws.Range("A1:F1").Font.Bold = True
    ws.Range("H1:M1").Font.Bold = True

    Set tbl = EnsureTable(ws, "tbl_???????", "H1:M1")
    tbl.Name = "tbl_???????"

    ws.Range("A2").Value = "BY-0001"
    ws.Range("B2").Value = Date
    ws.Range("C2").Value = ""
    ws.Range("D2").Value = 0
    ws.Range("E2").Value = ""

    ws.Columns("A:M").AutoFit
    ApplySafeValidation ws.Range("C2:C50"), "tbl_??????", "???????"
    ApplySafeValidation ws.Range("E2:E50"), "tbl_??????", "??位"
    AddFormButton ws, "btn_???绫???", "SaveOverflow", 8, 3
    AddFormButton ws, "btn_???????", "ClearOverflowEntry", 10, 3
    AddFormButton ws, "btn_??????????", "AddEntryRow", 15, 2
    AddFormButton ws, "btn_?????????", "RemoveEntryRow", 15, 3
    AddFormButton ws, "btn_??????", "PrintCurrentDocument", 15, 4
    ApplyEntrySheetStyle ws, "A1:F2", "H1:M1", "M"
End Sub

Private Sub EnsureInventorySummary()
    Dim ws As Worksheet
    Dim tbl As ListObject

    Set ws = GetOrCreateSheet(SHEET_KC)
    ws.Cells.Clear
    ws.Range("A1:E1").Value = Array("???????", "??位", "???????", "???????", "??????")
    ws.Range("A1:E1").Font.Bold = True

    Set tbl = EnsureTable(ws, "tbl_??????", "A1:E1")
    tbl.Name = "tbl_??????"
    ws.Columns("A:E").AutoFit
    ApplySummarySheetStyle ws, "A1:E1", "E"
End Sub

Private Sub EnsureProfitSummary()
    Dim ws As Worksheet
    Dim tbl As ListObject

    Set ws = GetOrCreateSheet(SHEET_LR)
    ws.Cells.Clear
    ws.Range("A1:H1").Value = Array("????", "?????", "???????", "????", "???????", "???????", "????", "????")
    ws.Range("A1:H1").Font.Bold = True

    Set tbl = EnsureTable(ws, "tbl_???????", "A1:H1")
    tbl.Name = "tbl_???????"
    ws.Columns("A:H").AutoFit
    ApplySummarySheetStyle ws, "A1:H1", "H"
End Sub

Private Sub ApplyMasterSheetStyle(ByVal ws As Worksheet)
    Dim tbl As ListObject

    ws.Tab.Color = RGB(46, 125, 115)
    ws.Range("A1:L1").Interior.Color = RGB(31, 78, 121)
    ws.Range("A1:L1").Font.Color = RGB(255, 255, 255)
    ws.Range("A1:L1").Font.Bold = True
    ws.Range("A1:L1").RowHeight = 24
    For Each tbl In ws.ListObjects
        tbl.TableStyle = "TableStyleMedium2"
    Next tbl
    ws.Range("A2:L200").Borders.LineStyle = 1
    ws.Range("A:L").ColumnWidth = 14
End Sub

Private Sub ApplyReceiptFormStyle(ByVal ws As Worksheet)
    ws.Tab.Color = RGB(237, 125, 49)
    ws.Range("A1:G1").Interior.Color = RGB(31, 78, 121)
    ws.Range("A1:G1").Font.Color = RGB(255, 255, 255)
    ws.Range("A1:G1").Font.Size = 18
    ws.Range("A1:G1").HorizontalAlignment = -4108
    ws.Range("A1:G1").RowHeight = 32
    ws.Range("A2:G2").Font.Bold = True
    ws.Range("B2:C2").Interior.Color = RGB(255, 242, 204)
    ws.Range("E2:F2").Interior.Color = RGB(255, 242, 204)
    ws.Range("A4:G4").Interior.Color = RGB(221, 235, 247)
    ws.Range("A4:G4").Font.Bold = True
    ws.Range("A4:G50").Borders.LineStyle = 1
    ws.Range("A4:G50").VerticalAlignment = -4160
    ws.Range("A4:G50").HorizontalAlignment = -4108
    ws.Range("C5:C50").NumberFormat = "0.00"
    ws.Range("E5:F50").NumberFormat = "0.00"
    ws.Range("B2:C2").Borders.LineStyle = 1
    ws.Range("E2:F2").Borders.LineStyle = 1
End Sub

Private Sub ApplyEntrySheetStyle(ByVal ws As Worksheet, ByVal entryAddress As String, ByVal tableAddress As String, ByVal lastColumn As String)
    Dim tbl As ListObject

    ws.Tab.Color = RGB(237, 125, 49)
    ws.Range(entryAddress).Borders.LineStyle = 1
    ws.Range(entryAddress).Rows(1).Interior.Color = RGB(31, 78, 121)
    ws.Range(entryAddress).Rows(1).Font.Color = RGB(255, 255, 255)
    ws.Range(entryAddress).Rows(1).Font.Bold = True
    ws.Range(entryAddress).Rows(2).Interior.Color = RGB(255, 242, 204)
    ws.Range(entryAddress).Rows(2).RowHeight = 24
    ws.Range(tableAddress).Interior.Color = RGB(31, 78, 121)
    ws.Range(tableAddress).Font.Color = RGB(255, 255, 255)
    ws.Range(tableAddress).Font.Bold = True
    ws.Range("A1:" & lastColumn & "200").Borders.LineStyle = 1
    For Each tbl In ws.ListObjects
        tbl.TableStyle = "TableStyleMedium2"
    Next tbl
    ws.Columns(lastColumn).ColumnWidth = 12
    ws.Range("E2:H200").NumberFormat = "0.00"
End Sub

Private Sub ApplySummarySheetStyle(ByVal ws As Worksheet, ByVal headerAddress As String, ByVal lastColumn As String)
    Dim tbl As ListObject

    ws.Tab.Color = RGB(112, 173, 71)
    ws.Range(headerAddress).Interior.Color = RGB(31, 78, 121)
    ws.Range(headerAddress).Font.Color = RGB(255, 255, 255)
    ws.Range(headerAddress).Font.Bold = True
    ws.Range(headerAddress).RowHeight = 24
    ws.Range("A1:" & lastColumn & "200").Borders.LineStyle = 1
    For Each tbl In ws.ListObjects
        tbl.TableStyle = "TableStyleMedium2"
    Next tbl
    ws.Range("C2:E200").NumberFormat = "0.00"
    ws.Range("D2:F200").NumberFormat = "0.00"
End Sub

Public Sub ValidateAndAddProduct(ByVal productName As String, ByVal unit As String, Optional ByVal note As String = "")
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long

    If Trim$(productName) = "" Then
        MsgBox "??????????????", vbExclamation, "校?????"
        Exit Sub
    End If

    Set ws = GetOrCreateSheet(BASE_SHEET)
    Set tbl = ws.ListObjects("tbl_??????")
    If tbl Is Nothing Then
        EnsureBaseDataStructure
        Set tbl = ws.ListObjects("tbl_??????")
    End If

    For i = 1 To tbl.ListRows.Count
        If StrComp(Trim$(CStr(tbl.DataBodyRange(i, 1).Value)), Trim$(productName), 1) = 0 Then
            MsgBox "???????????????????????", vbExclamation, "???校??"
            Exit Sub
        End If
    Next i

    With tbl.ListRows.Add
        .Range(1, 1).Value = productName
        .Range(1, 2).Value = unit
        .Range(1, 3).Value = note
    End With
    RefreshDropdownSources
    MsgBox "???????????", vbInformation, "????????"
End Sub

Public Sub ValidateAndAddSupplier(ByVal supplierName As String, Optional ByVal contact As String = "", Optional ByVal note As String = "")
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long

    If Trim$(supplierName) = "" Then
        MsgBox "????????????????", vbExclamation, "校?????"
        Exit Sub
    End If

    Set ws = GetOrCreateSheet(BASE_SHEET)
    Set tbl = ws.ListObjects("tbl_????????")
    If tbl Is Nothing Then
        EnsureBaseDataStructure
        Set tbl = ws.ListObjects("tbl_????????")
    End If

    For i = 1 To tbl.ListRows.Count
        If StrComp(Trim$(CStr(tbl.DataBodyRange(i, 1).Value)), Trim$(supplierName), 1) = 0 Then
            MsgBox "?????????????????????????", vbExclamation, "???校??"
            Exit Sub
        End If
    Next i

    With tbl.ListRows.Add
        .Range(1, 1).Value = supplierName
        .Range(1, 2).Value = contact
        .Range(1, 3).Value = note
    End With
    RefreshDropdownSources
End Sub

Public Sub ValidateAndAddCustomer(ByVal customerName As String, Optional ByVal contact As String = "", Optional ByVal note As String = "")
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long

    If Trim$(customerName) = "" Then
        MsgBox "??????????????", vbExclamation, "校?????"
        Exit Sub
    End If

    Set ws = GetOrCreateSheet(BASE_SHEET)
    Set tbl = ws.ListObjects("tbl_??????")
    If tbl Is Nothing Then
        EnsureBaseDataStructure
        Set tbl = ws.ListObjects("tbl_??????")
    End If

    For i = 1 To tbl.ListRows.Count
        If StrComp(Trim$(CStr(tbl.DataBodyRange(i, 1).Value)), Trim$(customerName), 1) = 0 Then
            MsgBox "???????????????????????", vbExclamation, "???校??"
            Exit Sub
        End If
    Next i

    With tbl.ListRows.Add
        .Range(1, 1).Value = customerName
        .Range(1, 2).Value = contact
        .Range(1, 3).Value = note
    End With
    RefreshDropdownSources
End Sub

Public Sub SaveReceipt()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim productName As String, supplierName As String, qty As Double, price As Double, unit As String
    Dim amount As Double, docNo As String, note As String
    Dim i As Long

    Set ws = GetOrCreateSheet(SHEET_RK)
    Set tbl = ws.ListObjects("tbl_??????")
    If tbl Is Nothing Then EnsureReceiptSheet: Set tbl = ws.ListObjects("tbl_??????")

    docNo = Trim$(CStr(ws.Range("I2").Value))
    supplierName = Trim$(CStr(ws.Range("B2").Value))
    If docNo = "" Then docNo = "RK-" & Format(Date, "yyyymmdd") & "-" & GetNextDocNo(SHEET_RK)
    If supplierName = "" Then
        MsgBox "??????位?????????????????????????斜??????", vbExclamation, "??????"
        Exit Sub
    End If
    If Not MasterValueExists("tbl_????????", "?????????", supplierName) Then
        MsgBox "??????位???????????校??????????????斜????", vbExclamation, "??????"
        Exit Sub
    End If

    For i = 5 To 50
        productName = Trim$(CStr(ws.Cells(i, 2).Value))
        qty = ToDouble(ws.Cells(i, 3).Value)
        unit = Trim$(CStr(ws.Cells(i, 4).Value))
        price = ToDouble(ws.Cells(i, 5).Value)
        note = Trim$(CStr(ws.Cells(i, 7).Value))

        If productName = "" And qty = 0 Then GoTo NextReceiptRow
        If productName = "" Then
            MsgBox "?? " & i - 4 & " ?械???????????????", vbExclamation, "??????"
            Exit Sub
        End If
        If Not MasterValueExists("tbl_??????", "???????", productName) Then
            MsgBox "?? " & i - 4 & " ?械??????????????????校?????????????斜????", vbExclamation, "??????"
            Exit Sub
        End If
        If qty <= 0 Then
            MsgBox "?? " & i - 4 & " ?械??????????????? 0??", vbExclamation, "??????"
            Exit Sub
        End If

        unit = GetProductUnit(productName, unit)
        ws.Cells(i, 4).Value = unit
        amount = qty * price

        With tbl.ListRows.Add
            .Range(1, 1).Value = Date
            .Range(1, 2).Value = docNo
            .Range(1, 3).Value = supplierName
            .Range(1, 4).Value = productName
            .Range(1, 5).Value = qty
            .Range(1, 6).Value = unit
            .Range(1, 7).Value = price
            .Range(1, 8).Value = amount
        End With
NextReceiptRow:
    Next i

    RefreshInventorySummary
    ClearReceiptEntry
    ws.Range("I2").Value = "RK-" & Format(Date, "yyyymmdd") & "-" & GetNextDocNo(SHEET_RK)
    MsgBox "??????????妗?", vbInformation, "??????"
End Sub

Public Sub SaveSales()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim customerName As String, productName As String, qty As Double, price As Double, unit As String
    Dim amount As Double, docNo As String

    Set ws = GetOrCreateSheet(SHEET_XS)
    docNo = Trim$(CStr(ws.Range("A2").Value))
    customerName = Trim$(CStr(ws.Range("C2").Value))
    productName = Trim$(CStr(ws.Range("D2").Value))
    qty = ToDouble(ws.Range("E2").Value)
    unit = Trim$(CStr(ws.Range("F2").Value))
    price = ToDouble(ws.Range("G2").Value)

    If docNo = "" Or customerName = "" Or productName = "" Then
        MsgBox "?????????????????????????", vbExclamation, "?????"
        Exit Sub
    End If
    If qty <= 0 Then
        MsgBox "??????????????? 0??", vbExclamation, "?????"
        Exit Sub
    End If

    unit = GetProductUnit(productName, unit)
    amount = qty * price

    Set tbl = ws.ListObjects("tbl_???????")
    If tbl Is Nothing Then EnsureSalesSheet: Set tbl = ws.ListObjects("tbl_???????")

    With tbl.ListRows.Add
        .Range(1, 1).Value = Date
        .Range(1, 2).Value = docNo
        .Range(1, 3).Value = customerName
        .Range(1, 4).Value = productName
        .Range(1, 5).Value = qty
        .Range(1, 6).Value = unit
        .Range(1, 7).Value = price
        .Range(1, 8).Value = amount
    End With

    SaveProfitSnapshot productName, qty, price, docNo, Date
    RefreshInventorySummary
    ClearEntryRange ws, "A2:I2"
    ws.Range("A2").Value = "XS-" & Format(Date, "yyyymmdd") & "-" & GetNextDocNo(SHEET_XS)
    MsgBox "?????????妗?", vbInformation, "?????"
End Sub

Public Sub SaveSalesReturn()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim customerName As String, productName As String, qty As Double, returnPrice As Double, unit As String
    Dim amount As Double, docNo As String

    Set ws = GetOrCreateSheet(SHEET_XSTH)
    docNo = Trim$(CStr(ws.Range("A2").Value))
    customerName = Trim$(CStr(ws.Range("C2").Value))
    productName = Trim$(CStr(ws.Range("D2").Value))
    qty = ToDouble(ws.Range("E2").Value)
    unit = Trim$(CStr(ws.Range("F2").Value))
    returnPrice = ToDouble(ws.Range("G2").Value)

    If docNo = "" Or customerName = "" Or productName = "" Then
        MsgBox "?????????????????????????", vbExclamation, "?????????"
        Exit Sub
    End If
    If qty <= 0 Then
        MsgBox "?????????????? 0??", vbExclamation, "?????????"
        Exit Sub
    End If

    unit = GetProductUnit(productName, unit)
    amount = qty * returnPrice

    Set tbl = ws.ListObjects("tbl_??????????")
    If tbl Is Nothing Then EnsureSalesReturnSheet: Set tbl = ws.ListObjects("tbl_??????????")

    With tbl.ListRows.Add
        .Range(1, 1).Value = Date
        .Range(1, 2).Value = docNo
        .Range(1, 3).Value = customerName
        .Range(1, 4).Value = productName
        .Range(1, 5).Value = qty
        .Range(1, 6).Value = unit
        .Range(1, 7).Value = returnPrice
        .Range(1, 8).Value = amount
    End With

    SaveProfitSnapshot productName, -qty, returnPrice, docNo, Date
    RefreshInventorySummary
    ClearEntryRange ws, "A2:I2"
    ws.Range("A2").Value = "XSTH-" & Format(Date, "yyyymmdd") & "-" & GetNextDocNo(SHEET_XSTH)
    MsgBox "?????????????妗?", vbInformation, "?????????"
End Sub

Public Sub SavePurchaseReturn()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim supplierName As String, productName As String, qty As Double, returnPrice As Double, unit As String
    Dim amount As Double, docNo As String

    Set ws = GetOrCreateSheet(SHEET_JHTH)
    docNo = Trim$(CStr(ws.Range("A2").Value))
    supplierName = Trim$(CStr(ws.Range("C2").Value))
    productName = Trim$(CStr(ws.Range("D2").Value))
    qty = ToDouble(ws.Range("E2").Value)
    unit = Trim$(CStr(ws.Range("F2").Value))
    returnPrice = ToDouble(ws.Range("G2").Value)

    If docNo = "" Or supplierName = "" Or productName = "" Then
        MsgBox "??????????????????????????", vbExclamation, "?????????"
        Exit Sub
    End If
    If qty <= 0 Then
        MsgBox "?????????????? 0??", vbExclamation, "?????????"
        Exit Sub
    End If

    unit = GetProductUnit(productName, unit)
    amount = qty * returnPrice

    Set tbl = ws.ListObjects("tbl_??????????")
    If tbl Is Nothing Then EnsurePurchaseReturnSheet: Set tbl = ws.ListObjects("tbl_??????????")

    With tbl.ListRows.Add
        .Range(1, 1).Value = Date
        .Range(1, 2).Value = docNo
        .Range(1, 3).Value = supplierName
        .Range(1, 4).Value = productName
        .Range(1, 5).Value = qty
        .Range(1, 6).Value = unit
        .Range(1, 7).Value = returnPrice
        .Range(1, 8).Value = amount
    End With

    RefreshInventorySummary
    ClearEntryRange ws, "A2:I2"
    ws.Range("A2").Value = "JHTH-" & Format(Date, "yyyymmdd") & "-" & GetNextDocNo(SHEET_JHTH)
    MsgBox "?????????????妗?", vbInformation, "?????????"
End Sub

Public Sub SaveLoss()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim productName As String, qty As Double, unit As String, docNo As String

    Set ws = GetOrCreateSheet(SHEET_BS)
    docNo = Trim$(CStr(ws.Range("A2").Value))
    productName = Trim$(CStr(ws.Range("C2").Value))
    qty = ToDouble(ws.Range("D2").Value)
    unit = Trim$(CStr(ws.Range("E2").Value))

    If docNo = "" Or productName = "" Then
        MsgBox "????????????????????", vbExclamation, "????"
        Exit Sub
    End If
    If qty <= 0 Then
        MsgBox "??????????????? 0??", vbExclamation, "????"
        Exit Sub
    End If

    unit = GetProductUnit(productName, unit)

    Set tbl = ws.ListObjects("tbl_???????")
    If tbl Is Nothing Then EnsureLossSheet: Set tbl = ws.ListObjects("tbl_???????")

    With tbl.ListRows.Add
        .Range(1, 1).Value = Date
        .Range(1, 2).Value = docNo
        .Range(1, 3).Value = productName
        .Range(1, 4).Value = qty
        .Range(1, 5).Value = unit
    End With

    RefreshInventorySummary
    ClearEntryRange ws, "A2:F2"
    ws.Range("A2").Value = "BS-" & Format(Date, "yyyymmdd") & "-" & GetNextDocNo(SHEET_BS)
    MsgBox "????????妗?", vbInformation, "????"
End Sub

Public Sub SaveOverflow()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim productName As String, qty As Double, unit As String, docNo As String

    Set ws = GetOrCreateSheet(SHEET_BY)
    docNo = Trim$(CStr(ws.Range("A2").Value))
    productName = Trim$(CStr(ws.Range("C2").Value))
    qty = ToDouble(ws.Range("D2").Value)
    unit = Trim$(CStr(ws.Range("E2").Value))

    If docNo = "" Or productName = "" Then
        MsgBox "????????????????????", vbExclamation, "???绲?"
        Exit Sub
    End If
    If qty <= 0 Then
        MsgBox "??????????????? 0??", vbExclamation, "???绲?"
        Exit Sub
    End If

    unit = GetProductUnit(productName, unit)

    Set tbl = ws.ListObjects("tbl_???????")
    If tbl Is Nothing Then EnsureOverflowSheet: Set tbl = ws.ListObjects("tbl_???????")

    With tbl.ListRows.Add
        .Range(1, 1).Value = Date
        .Range(1, 2).Value = docNo
        .Range(1, 3).Value = productName
        .Range(1, 4).Value = qty
        .Range(1, 5).Value = unit
    End With

    RefreshInventorySummary
    ClearEntryRange ws, "A2:F2"
    ws.Range("A2").Value = "BY-" & Format(Date, "yyyymmdd") & "-" & GetNextDocNo(SHEET_BY)
    MsgBox "???绲?????妗?", vbInformation, "???绲?"
End Sub

Public Sub RefreshInventorySummary()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dict As Object
    Dim prodName As Variant
    Dim productList As Variant
    Dim i As Long, j As Long
    Dim qtyIn As Double, qtyOut As Double, qtySalesRet As Double, qtyPurRet As Double, qtyLoss As Double, qtyOverflow As Double
    Dim unitName As String, avgCost As Double, totalCost As Double
    Dim rowCount As Long

    Set ws = GetOrCreateSheet(BASE_SHEET)
    Set tbl = ws.ListObjects("tbl_??????")
    Set dict = CreateObject("Scripting.Dictionary")

    If Not tbl Is Nothing And Not tbl.DataBodyRange Is Nothing Then
        productList = tbl.DataBodyRange.Value
        For i = 1 To UBound(productList, 1)
            prodName = Trim$(CStr(productList(i, 1)))
            unitName = Trim$(CStr(productList(i, 2)))
            If prodName <> "" Then
                dict(prodName) = Array(unitName, 0, 0, 0)
            End If
        Next i
    End If

    For Each prodName In dict.Keys
        qtyIn = 0: qtyOut = 0: qtySalesRet = 0: qtyPurRet = 0: qtyLoss = 0: qtyOverflow = 0

        qtyIn = SumQtyFromTable(SHEET_RK, "tbl_??????", 4, prodName)
        qtyOut = SumQtyFromTable(SHEET_XS, "tbl_???????", 4, prodName)
        qtySalesRet = SumQtyFromTable(SHEET_XSTH, "tbl_??????????", 4, prodName)
        qtyPurRet = SumQtyFromTable(SHEET_JHTH, "tbl_??????????", 4, prodName)
        qtyLoss = SumQtyFromTable(SHEET_BS, "tbl_???????", 3, prodName)
        qtyOverflow = SumQtyFromTable(SHEET_BY, "tbl_???????", 3, prodName)

        dict(prodName) = Array(dict(prodName)(0), qtyIn - qtyOut + qtySalesRet - qtyPurRet + qtyOverflow - qtyLoss, 0, 0)
    Next prodName

    Set ws = GetOrCreateSheet(SHEET_KC)
    If ws.ListObjects("tbl_??????") Is Nothing Then EnsureInventorySummary
    Set tbl = ws.ListObjects("tbl_??????")
    DeleteAllTableRows tbl

    rowCount = 0
    For Each prodName In dict.Keys
        unitName = CStr(dict(prodName)(0))
        qtyIn = CDbl(dict(prodName)(1))
        avgCost = GetAverageCost(prodName)
        totalCost = qtyIn * avgCost

        If qtyIn <> 0 Or prodName <> "" Then
            rowCount = rowCount + 1
            With tbl.ListRows.Add
                .Range(1, 1).Value = prodName
                .Range(1, 2).Value = unitName
                .Range(1, 3).Value = qtyIn
                .Range(1, 4).Value = avgCost
                .Range(1, 5).Value = totalCost
            End With
        End If
    Next prodName
End Sub

Public Sub RefreshProfitSummary()
    Dim ws As Worksheet
    Dim tbl As ListObject

    Set ws = GetOrCreateSheet(SHEET_LR)
    Set tbl = ws.ListObjects("tbl_???????")
    If tbl Is Nothing Then EnsureProfitSummary
    MsgBox "????????????????????????????????????", vbInformation, "???????"
End Sub

Private Sub SaveProfitSnapshot(ByVal productName As String, ByVal qty As Double, ByVal salePrice As Double, ByVal docNo As String, ByVal docDate As Date)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim avgCost As Double, adjustment As Double, profit As Double

    Set ws = GetOrCreateSheet(SHEET_LR)
    Set tbl = ws.ListObjects("tbl_???????")
    If tbl Is Nothing Then EnsureProfitSummary: Set tbl = ws.ListObjects("tbl_???????")

    avgCost = GetAverageCost(productName)
    adjustment = 0
    profit = qty * (salePrice - avgCost) + adjustment

    With tbl.ListRows.Add
        .Range(1, 1).Value = docDate
        .Range(1, 2).Value = docNo
        .Range(1, 3).Value = productName
        .Range(1, 4).Value = qty
        .Range(1, 5).Value = salePrice
        .Range(1, 6).Value = avgCost
        .Range(1, 7).Value = adjustment
        .Range(1, 8).Value = profit
    End With
End Sub

Private Function GetAverageCost(ByVal productName As String) As Double
    Dim qtyIn As Double
    Dim qtyReturn As Double
    Dim totalIn As Double
    Dim totalReturn As Double
    Dim netQty As Double

    qtyIn = SumQtyFromTable(SHEET_RK, "tbl_??????", 4, productName)
    qtyReturn = SumQtyFromTable(SHEET_JHTH, "tbl_??????????", 4, productName)
    totalIn = SumAmountFromTable(SHEET_RK, "tbl_??????", 4, productName)
    totalReturn = SumAmountFromTable(SHEET_JHTH, "tbl_??????????", 4, productName)
    netQty = qtyIn - qtyReturn

    If netQty > 0 Then
        GetAverageCost = (totalIn - totalReturn) / netQty
    Else
        GetAverageCost = 0
    End If
End Function

Private Function GetProductUnit(ByVal productName As String, Optional ByVal fallback As String = "") As String
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long

    Set ws = GetOrCreateSheet(BASE_SHEET)
    Set tbl = ws.ListObjects("tbl_??????")
    If tbl Is Nothing Or tbl.DataBodyRange Is Nothing Then
        GetProductUnit = fallback
        Exit Function
    End If

    For i = 1 To tbl.DataBodyRange.Rows.Count
        If StrComp(Trim$(CStr(tbl.DataBodyRange(i, 1).Value)), Trim$(productName), 1) = 0 Then
            GetProductUnit = Trim$(CStr(tbl.DataBodyRange(i, 2).Value))
            If GetProductUnit = "" Then GetProductUnit = fallback
            Exit Function
        End If
    Next i

    GetProductUnit = fallback
End Function

Private Function MasterValueExists(ByVal tableName As String, ByVal columnName As String, ByVal valueText As String) As Boolean
    Dim tbl As ListObject
    Dim i As Long
    Dim columnIndex As Long
    Dim dataRange As Range

    Set tbl = GetTableIfExists(BASE_SHEET, tableName)
    If tbl Is Nothing Or tbl.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To tbl.HeaderRowRange.Columns.Count
        If StrComp(Trim$(CStr(tbl.HeaderRowRange.Cells(1, i).Value)), Trim$(columnName), 1) = 0 Then
            columnIndex = i
            Exit For
        End If
    Next i
    If columnIndex = 0 Then Exit Function

    Set dataRange = tbl.DataBodyRange.Columns(columnIndex)
    If dataRange Is Nothing Then Exit Function
    For i = 1 To dataRange.Rows.Count
        If StrComp(Trim$(CStr(dataRange.Cells(i, 1).Value)), Trim$(valueText), 1) = 0 Then
            MasterValueExists = True
            Exit Function
        End If
    Next i
End Function

Private Function SumQtyFromTable(ByVal sheetName As String, ByVal tableName As String, ByVal productColumn As Long, ByVal productName As String) As Double
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    Dim qty As Double

    Set ws = GetOrCreateSheet(sheetName)
    Set tbl = ws.ListObjects(tableName)
    If tbl Is Nothing Or tbl.DataBodyRange Is Nothing Then
        SumQtyFromTable = 0
        Exit Function
    End If

    For i = 1 To tbl.DataBodyRange.Rows.Count
        If StrComp(Trim$(CStr(tbl.DataBodyRange(i, productColumn).Value)), Trim$(productName), 1) = 0 Then
            qty = qty + ToDouble(tbl.DataBodyRange(i, 5).Value)
        End If
    Next i

    SumQtyFromTable = qty
End Function

Private Function SumAmountFromTable(ByVal sheetName As String, ByVal tableName As String, ByVal productColumn As Long, ByVal productName As String) As Double
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    Dim amount As Double

    Set ws = GetOrCreateSheet(sheetName)
    Set tbl = ws.ListObjects(tableName)
    If tbl Is Nothing Or tbl.DataBodyRange Is Nothing Then
        SumAmountFromTable = 0
        Exit Function
    End If

    For i = 1 To tbl.DataBodyRange.Rows.Count
        If StrComp(Trim$(CStr(tbl.DataBodyRange(i, productColumn).Value)), Trim$(productName), 1) = 0 Then
            amount = amount + ToDouble(tbl.DataBodyRange(i, 8).Value)
        End If
    Next i

    SumAmountFromTable = amount
End Function

Private Function GetTableIfExists(ByVal sheetName As String, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    If ws Is Nothing Then
        Set GetTableIfExists = Nothing
        Exit Function
    End If
    Set GetTableIfExists = ws.ListObjects(tableName)
End Function

Private Sub RefreshAllMasterDropdowns()
    Dim ws As Worksheet

    On Error Resume Next

    Set ws = GetOrCreateSheet(SHEET_RK)
    ApplySafeValidation ws.Range("B2"), "tbl_????????", "?????????"
    ApplySafeValidation ws.Range("B5:B50"), "tbl_??????", "???????"
    ApplyDateValidation ws.Range("E2")

    Set ws = GetOrCreateSheet(SHEET_XS)
    ApplySafeValidation ws.Range("C2:C50"), "tbl_??????", "???????"
    ApplySafeValidation ws.Range("D2:D50"), "tbl_??????", "???????"
    ApplySafeValidation ws.Range("F2:F50"), "tbl_??????", "??位"

    Set ws = GetOrCreateSheet(SHEET_XSTH)
    ApplySafeValidation ws.Range("C2:C50"), "tbl_??????", "???????"
    ApplySafeValidation ws.Range("D2:D50"), "tbl_??????", "???????"
    ApplySafeValidation ws.Range("F2:F50"), "tbl_??????", "??位"

    Set ws = GetOrCreateSheet(SHEET_JHTH)
    ApplySafeValidation ws.Range("C2:C50"), "tbl_????????", "?????????"
    ApplySafeValidation ws.Range("D2:D50"), "tbl_??????", "???????"
    ApplySafeValidation ws.Range("F2:F50"), "tbl_??????", "??位"

    Set ws = GetOrCreateSheet(SHEET_BS)
    ApplySafeValidation ws.Range("C2:C50"), "tbl_??????", "???????"
    ApplySafeValidation ws.Range("E2:E50"), "tbl_??????", "??位"

    Set ws = GetOrCreateSheet(SHEET_BY)
    ApplySafeValidation ws.Range("C2:C50"), "tbl_??????", "???????"
    ApplySafeValidation ws.Range("E2:E50"), "tbl_??????", "??位"

    On Error GoTo 0
End Sub

Private Sub ApplySafeValidation(ByVal rng As Range, ByVal tableName As String, ByVal columnName As String)
    Dim srcRange As Range
    Dim listName As String
    Dim formulaText As String

    Set srcRange = GetMasterValidationRange(tableName)
    If srcRange Is Nothing Then Exit Sub
    listName = GetMasterValidationName(tableName)
    EnsureNamedListRange listName, srcRange
    formulaText = "=" & listName

    On Error Resume Next
    With rng.Validation
        .Delete
        .Add Type:=3, AlertStyle:=1, Operator:=1, Formula1:=formulaText
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    On Error GoTo 0
End Sub

Private Function GetMasterValidationRange(ByVal tableName As String) As Range
    Dim ws As Worksheet
    Dim columnNumber As Long
    Dim lastRow As Long
    Dim rowNumber As Long
    Dim cellValue As String

    Set ws = GetOrCreateSheet(BASE_SHEET)
    Select Case tableName
        Case "tbl_??????": columnNumber = 1
        Case "tbl_????????": columnNumber = 5
        Case "tbl_??????": columnNumber = 10
    End Select

    If columnNumber = 0 Then Exit Function
    For rowNumber = 2 To 200
        cellValue = Trim$(CStr(ws.Cells(rowNumber, columnNumber).Value))
        If cellValue <> "" Then lastRow = rowNumber
    Next rowNumber
    If lastRow < 2 Then lastRow = 2
    Set GetMasterValidationRange = ws.Range(ws.Cells(2, columnNumber), ws.Cells(lastRow, columnNumber))
End Function

Private Function GetMasterValidationName(ByVal tableName As String) As String
    Select Case tableName
        Case "tbl_??????": GetMasterValidationName = "ListProducts"
        Case "tbl_????????": GetMasterValidationName = "ListSuppliers"
        Case "tbl_??????": GetMasterValidationName = "ListCustomers"
        Case Else: GetMasterValidationName = "ListMasterValues"
    End Select
End Function

Private Sub EnsureDateValidationList()
    Dim ws As Worksheet
    Dim i As Long

    Set ws = GetOrCreateSheet("????????")
    ws.Range("A1").Value = "???????"
    For i = 2 To 732
        ws.Cells(i, 1).Value = Date - 365 + (i - 2)
        ws.Cells(i, 1).NumberFormat = "yyyy/m/d"
    Next i
    ws.Visible = False
    EnsureNamedListRange "ListReceiptDates", ws.Range("A2:A732")
End Sub

Private Sub ApplyDateValidation(ByVal rng As Range)
    On Error Resume Next
    With rng.Validation
        .Delete
        .Add Type:=3, AlertStyle:=1, Operator:=1, Formula1:="=ListReceiptDates"
        .IgnoreBlank = False
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
        .InputTitle = "????"
        .InputMessage = "????????????斜??????"
        .ErrorTitle = "??????效"
        .ErrorMessage = "?????????????斜??????"
    End With
    On Error GoTo 0
End Sub

Private Sub EnsureNamedListRange(ByVal listName As String, ByVal srcRange As Range)
    Dim refText As String

    refText = "=" & srcRange.Address(True, True, 1, True)

    On Error Resume Next
    ThisWorkbook.Names(listName).Delete
    ThisWorkbook.Names.Add Name:=listName, RefersTo:=refText
    On Error GoTo 0
End Sub

Private Function EnsureTable(ByVal ws As Worksheet, ByVal tableName As String, ByVal startCell As String) As ListObject
    Dim tbl As ListObject
    Dim rng As Range

    On Error Resume Next
    Set tbl = ws.ListObjects(tableName)
    On Error GoTo 0

    If tbl Is Nothing Then
        Set rng = ws.Range(startCell)
        Set tbl = ws.ListObjects.Add(Source:=rng, XlListObjectHasHeaders:=True, Destination:=rng)
        tbl.Name = tableName
    End If

    Set EnsureTable = tbl
End Function

Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If

    Set GetOrCreateSheet = ws
End Function

Private Function WorksheetExists(ByVal sheetName As String) As Boolean
    On Error Resume Next
    WorksheetExists = Not ThisWorkbook.Worksheets(sheetName) Is Nothing
    On Error GoTo 0
End Function

Private Function TableExistsOnSheet(ByVal sheetName As String, ByVal tableName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    If ws Is Nothing Then
        TableExistsOnSheet = False
        Exit Function
    End If
    TableExistsOnSheet = Not ws.ListObjects(tableName) Is Nothing
    On Error GoTo 0
End Function

Private Function ToDouble(ByVal value As Variant) As Double
    If IsNumeric(value) Then
        ToDouble = CDbl(value)
    Else
        ToDouble = 0
    End If
End Function

Private Sub ClearEntryRange(ByVal ws As Worksheet, ByVal rangeAddress As String)
    ws.Range(rangeAddress).ClearContents
End Sub

Private Sub ClearEntryRows(ByVal ws As Worksheet, ByVal lastColumn As String)
    ws.Range("A2:" & lastColumn & FindLastEntryRow(ws, lastColumn)).ClearContents
End Sub

Private Sub DeleteAllTableRows(ByVal tbl As ListObject)
    Dim i As Long

    If tbl Is Nothing Then Exit Sub
    For i = tbl.ListRows.Count To 1 Step -1
        tbl.ListRows(i).Delete
    Next i
End Sub

Private Function GetNextDocNo(ByVal sheetName As String) As String
    Dim ws As Worksheet
    Dim i As Long
    Dim valueText As String
    Dim maxNo As Long

    Set ws = GetOrCreateSheet(sheetName)
    maxNo = 0

    For i = 2 To 2000
        valueText = Trim$(CStr(ws.Cells(i, 1).Value))
        If valueText <> "" Then
            If Len(valueText) >= 2 Then
                On Error Resume Next
                If IsNumeric(Right$(valueText, 3)) Then maxNo = Application.Max(maxNo, CLng(Right$(valueText, 3)))
                On Error GoTo 0
            End If
        End If
    Next i

    GetNextDocNo = Format(maxNo + 1, "000")
End Function
