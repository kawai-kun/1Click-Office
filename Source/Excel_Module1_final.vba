Attribute VB_Name = "Module1"
Private mHT_Co As Boolean
Private mHT_Ws As Object
Private mHT_R1 As Long, mHT_R2 As Long, mHT_C1 As Long, mHT_C2 As Long
Private mHT_ColW() As Double
Private mHT_RowH() As Double
Private mHT_FontAll As Variant, mHT_VAll As Variant, mHT_ShrinkAll As Variant, mHT_WrapAll As Variant
Private mHT_Font() As String
Private mHT_VAl() As Long
Private mHT_Shrink() As Boolean
Private mHT_Wrap() As Boolean
Private mHT_PerCell As Boolean
Private mHT_HBreaks As String
Private mHT_VBreaks As String
Private mHT_PS(1 To 20) As Variant
Private Const LIC_CAN_NHAP_MA As Boolean = False

#If VBA7 Then
    Private Declare PtrSafe Function MsgBoxW Lib "user32" Alias "MessageBoxW" ( _
        ByVal hWnd As LongPtr, _
        ByVal lpText As LongPtr, _
        ByVal lpCaption As LongPtr, _
        ByVal uType As Long) As Long
#Else
    Private Declare Function MsgBoxW Lib "user32" Alias "MessageBoxW" ( _
        ByVal hWnd As Long, _
        ByVal lpText As Long, _
        ByVal lpCaption As Long, _
        ByVal uType As Long) As Long
#End If

Public Sub RibbonChuanHoaExcel(control As Object)
Call ChuanHoa_TrangIn_Excel(False)
End Sub
Public Sub ChuanHoa_TrangIn_Excel_Silent()
Call ChuanHoa_TrangIn_Excel(True)
End Sub
Private Sub HT_InAn(ByVal bat As Boolean)
On Error Resume Next
Application.PrintCommunication = bat
On Error GoTo 0
End Sub
Private Sub HT_ChupTrangThai(ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, ByVal c1 As Long, ByVal c2 As Long)
mHT_Co = False
On Error Resume Next
Set mHT_Ws = ws
mHT_R1 = r1: mHT_R2 = r2: mHT_C1 = c1: mHT_C2 = c2
Dim r As Long, c As Long
ReDim mHT_ColW(c1 To c2)
For c = c1 To c2
mHT_ColW(c) = ws.Columns(c).ColumnWidth
Next c
ReDim mHT_RowH(r1 To r2)
For r = r1 To r2
mHT_RowH(r) = ws.Rows(r).RowHeight
Next r
Dim rg As Range
Set rg = ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2))
mHT_FontAll = rg.Font.Name
mHT_VAll = rg.VerticalAlignment
mHT_ShrinkAll = rg.ShrinkToFit
mHT_WrapAll = rg.WrapText
mHT_PerCell = (IsNull(mHT_FontAll) Or IsNull(mHT_VAll) Or IsNull(mHT_ShrinkAll) Or IsNull(mHT_WrapAll)) And _
(CDbl(r2 - r1 + 1) * CDbl(c2 - c1 + 1) <= 200000#)
If mHT_PerCell Then
ReDim mHT_Font(r1 To r2, c1 To c2)
ReDim mHT_VAl(r1 To r2, c1 To c2)
ReDim mHT_Shrink(r1 To r2, c1 To c2)
ReDim mHT_Wrap(r1 To r2, c1 To c2)
For r = r1 To r2
For c = c1 To c2
With ws.Cells(r, c)
mHT_Font(r, c) = .Font.Name
mHT_VAl(r, c) = .VerticalAlignment
mHT_Shrink(r, c) = .ShrinkToFit
mHT_Wrap(r, c) = .WrapText
End With
Next c
Next r
End If
mHT_HBreaks = "": mHT_VBreaks = ""
Dim hb As HPageBreak, vb As VPageBreak
For Each hb In ws.HPageBreaks
If hb.Type = xlPageBreakManual Then mHT_HBreaks = mHT_HBreaks & hb.Location.Row & ","
Next hb
For Each vb In ws.VPageBreaks
If vb.Type = xlPageBreakManual Then mHT_VBreaks = mHT_VBreaks & vb.Location.Column & ","
Next vb
With ws.PageSetup
mHT_PS(1) = .PaperSize
mHT_PS(2) = .Orientation
mHT_PS(3) = .LeftMargin
mHT_PS(4) = .RightMargin
mHT_PS(5) = .TopMargin
mHT_PS(6) = .BottomMargin
mHT_PS(7) = .HeaderMargin
mHT_PS(8) = .FooterMargin
mHT_PS(9) = .CenterHorizontally
mHT_PS(10) = .CenterVertically
mHT_PS(11) = .PrintGridlines
mHT_PS(12) = .PrintHeadings
mHT_PS(13) = .PrintArea
mHT_PS(14) = .PrintTitleRows
mHT_PS(15) = .Zoom
mHT_PS(16) = .FitToPagesWide
mHT_PS(17) = .FitToPagesTall
mHT_PS(18) = .CenterFooter
End With
On Error GoTo 0
mHT_Co = True
End Sub
Private Function HT_KhoiPhuc() As Boolean
HT_KhoiPhuc = False
If Not mHT_Co Then Exit Function
Dim ws As Worksheet
Set ws = Nothing
On Error Resume Next
Set ws = mHT_Ws
Dim chk As String
chk = ws.Name
If Err.Number <> 0 Or ws Is Nothing Then
mHT_Co = False
Exit Function
End If
Application.ScreenUpdating = False
Dim r As Long, c As Long
Dim rg As Range
Set rg = ws.Range(ws.Cells(mHT_R1, mHT_C1), ws.Cells(mHT_R2, mHT_C2))
If Not IsNull(mHT_FontAll) Then rg.Font.Name = mHT_FontAll
If Not IsNull(mHT_VAll) Then rg.VerticalAlignment = mHT_VAll
If Not IsNull(mHT_ShrinkAll) Then rg.ShrinkToFit = mHT_ShrinkAll
If Not IsNull(mHT_WrapAll) Then rg.WrapText = mHT_WrapAll
If mHT_PerCell Then
For r = mHT_R1 To mHT_R2
For c = mHT_C1 To mHT_C2
With ws.Cells(r, c)
If IsNull(mHT_FontAll) Then .Font.Name = mHT_Font(r, c)
If IsNull(mHT_VAll) Then .VerticalAlignment = mHT_VAl(r, c)
If IsNull(mHT_ShrinkAll) Then .ShrinkToFit = mHT_Shrink(r, c)
If IsNull(mHT_WrapAll) Then .WrapText = mHT_Wrap(r, c)
End With
Next c
Next r
End If
ws.ResetAllPageBreaks
Dim parts() As String, k As Long
parts = Split(mHT_HBreaks, ",")
For k = LBound(parts) To UBound(parts)
If Len(parts(k)) > 0 Then ws.HPageBreaks.Add Before:=ws.Rows(CLng(parts(k)))
Next k
parts = Split(mHT_VBreaks, ",")
For k = LBound(parts) To UBound(parts)
If Len(parts(k)) > 0 Then ws.VPageBreaks.Add Before:=ws.Columns(CLng(parts(k)))
Next k
Call HT_InAn(False)
With ws.PageSetup
.PaperSize = mHT_PS(1)
.Orientation = mHT_PS(2)
.LeftMargin = mHT_PS(3)
.RightMargin = mHT_PS(4)
.TopMargin = mHT_PS(5)
.BottomMargin = mHT_PS(6)
.HeaderMargin = mHT_PS(7)
.FooterMargin = mHT_PS(8)
.CenterHorizontally = mHT_PS(9)
.CenterVertically = mHT_PS(10)
.PrintGridlines = mHT_PS(11)
.PrintHeadings = mHT_PS(12)
.PrintArea = mHT_PS(13)
.PrintTitleRows = mHT_PS(14)
If VarType(mHT_PS(15)) = vbBoolean Then
.Zoom = False
.FitToPagesWide = mHT_PS(16)
.FitToPagesTall = mHT_PS(17)
Else
.Zoom = mHT_PS(15)
End If
.CenterFooter = mHT_PS(18)
End With
Call HT_InAn(True)
For c = mHT_C1 To mHT_C2
ws.Columns(c).ColumnWidth = mHT_ColW(c)
Next c
For r = mHT_R1 To mHT_R2
ws.Rows(r).RowHeight = mHT_RowH(r)
Next r
Application.ScreenUpdating = True
On Error GoTo 0
mHT_Co = False
HT_KhoiPhuc = True
End Function
Public Sub HoanTac_ChuanHoa_Excel()
If Not HT_KhoiPhuc() Then
MsgBoxUni "Kh" & ChrW(244) & "ng c" & ChrW(243) & " l" & ChrW(7847) & "n chu" & ChrW(7849) & "n h" & ChrW(243) & "a n" & ChrW(224) & "o " & ChrW(273) & ChrW(7875) & " ho" & ChrW(224) & "n t" & ChrW(225) & "c (ho" & ChrW(7863) & "c Excel " & ChrW(273) & ChrW(227) & " " & ChrW(273) & ChrW(243) & "ng/m" & ChrW(7903) & " l" & ChrW(7841) & "i sau l" & ChrW(7847) & "n chu" & ChrW(7849) & "n h" & ChrW(243) & "a).", vbInformation, "Ho" & ChrW(224) & "n T" & ChrW(225) & "c"
End If
End Sub
Public Sub RibbonHoanTacExcel(control As Object)
Call HoanTac_ChuanHoa_Excel
End Sub
Sub ChuanHoa_TrangIn_Excel(Optional ByVal silentMode As Boolean = False)
If Not silentMode Then
Call CN_TuKiemTraHangTuan(True)
If Not CheckLicenseActivation() Then Exit Sub
End If
Dim ws As Worksheet
Set ws = Nothing
On Error Resume Next
Set ws = ActiveSheet
MsgBoxUni "Kh" & ChrW(244) & "ng c" & ChrW(243) & " l" & ChrW(7847) & "n chu" & ChrW(7849) & "n h" & ChrW(243) & "a n" & ChrW(224) & "o " & ChrW(273) & ChrW(7875) & " ho" & ChrW(224) & "n t" & ChrW(225) & "c (ho" & ChrW(7863) & "c Excel " & ChrW(273) & ChrW(227) & " " & ChrW(273) & ChrW(243) & "ng/m" & ChrW(7903) & " l" & ChrW(7841) & "i sau l" & ChrW(7847) & "n chu" & ChrW(7849) & "n h" & ChrW(243) & "a).", vbInformation, "Ho" & ChrW(224) & "n T" & ChrW(225) & "c"
If ws Is Nothing Then Exit Sub
If ws.ProtectContents Then
If Not silentMode Then MsgBoxUni "B" & ChrW(7843) & "ng t" & ChrW(237) & "nh (Sheet) " & ChrW(273) & "ang b" & ChrW(7883) & " kh" & ChrW(243) & "a b" & ChrW(7843) & "o v" & ChrW(7879) & ". Vui l" & ChrW(242) & "ng b" & ChrW(7887) & " kh" & ChrW(243) & "a Sheet tr" & ChrW(432) & ChrW(7899) & "c khi chu" & ChrW(7849) & "n h" & ChrW(243) & "a trang in.", vbExclamation, "Chu" & ChrW(7849) & "n H" & ChrW(243) & "a Excel"
Exit Sub
End If
On Error GoTo LoiChuanHoa
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim firstRow As Long, lastRow As Long, firstCol As Long, lastCol As Long
If Not GetTrueDataBounds(ws, firstRow, lastRow, firstCol, lastCol) Then
Application.ScreenUpdating = True
Application.DisplayAlerts = True
If Not silentMode Then
MsgBoxUni "B" & ChrW(7843) & "ng t" & ChrW(237) & "nh hi" & ChrW(7879) & "n t" & ChrW(7841) & "i kh" & ChrW(244) & "ng c" & ChrW(243) & " d" & ChrW(7919) & " li" & ChrW(7879) & "u!", vbExclamation, "Chu" & ChrW(7849) & "n H" & ChrW(243) & "a Excel"
End If
Exit Sub
End If
If Not silentMode Then MsgBoxUni "B" & ChrW(7843) & "ng t" & ChrW(237) & "nh (Sheet) " & ChrW(273) & "ang b" & ChrW(7883) & " kh" & ChrW(243) & "a b" & ChrW(7843) & "o v" & ChrW(7879) & ". Vui l" & ChrW(242) & "ng b" & ChrW(7887) & " kh" & ChrW(243) & "a Sheet tr" & ChrW(432) & ChrW(7899) & "c khi chu" & ChrW(7849) & "n h" & ChrW(243) & "a trang in.", vbExclamation, "Chu" & ChrW(7849) & "n H" & ChrW(243) & "a Excel"
Set rng = ws.Range(ws.Cells(firstRow, firstCol), ws.Cells(lastRow, lastCol))
Call HT_ChupTrangThai(ws, firstRow, lastRow, firstCol, lastCol)
Dim origOrientation As Long
origOrientation = xlPortrait
On Error Resume Next
origOrientation = ws.PageSetup.Orientation
On Error GoTo LoiChuanHoa
ws.ResetAllPageBreaks
rng.Font.Name = "Times New Roman"
rng.VerticalAlignment = xlCenter
rng.ShrinkToFit = False
Dim rIdx As Long, cIdx As Long
Dim hasHeader As Boolean
Dim hStart As Long, hEnd As Long
hasHeader = FindTableHeaderRows(ws, firstRow, lastRow, firstCol, lastCol, hStart, hEnd)
If hasHeader Then
Call ClampLastColToTableHeader(ws, firstRow, lastRow, firstCol, lastCol, hStart, hEnd)
Set rng = ws.Range(ws.Cells(firstRow, firstCol), ws.Cells(lastRow, lastCol))
End If
Dim tableEndRow As Long
tableEndRow = FindTableDataEndRow(ws, firstRow, lastRow, firstCol, lastCol, hEnd, hasHeader)
Call ApplySmartWrapText(ws, firstRow, lastRow, firstCol, lastCol, hStart, tableEndRow, hasHeader)
Call FitColumnsSmart(ws, firstRow, lastRow, firstCol, lastCol, hStart, tableEndRow, hasHeader)
Dim totalWidthPt As Double
Dim visibleCols As Long
totalWidthPt = GetVisibleColumnsWidthPt(ws, firstCol, lastCol, visibleCols)
Dim logicalCols As Long
logicalCols = visibleCols
If hasHeader Then
Dim lCnt As Long: lCnt = 0
For cIdx = firstCol To lastCol
If Not ws.Columns(cIdx).Hidden Then
If ws.Cells(hStart, cIdx).Address = ws.Cells(hStart, cIdx).MergeArea.Cells(1, 1).Address Then
lCnt = lCnt + 1
End If
End If
Next cIdx
If lCnt >= 2 Then logicalCols = lCnt
End If
Dim chosenOrientation As Long
Dim targetPageWidthPt As Double
Dim targetPageHeightPt As Double
If origOrientation = xlLandscape Then
chosenOrientation = xlLandscape
targetPageWidthPt = 762#
targetPageHeightPt = 500#
ElseIf visibleCols <= 6 And logicalCols <= 6 And totalWidthPt <= 545 Then
chosenOrientation = xlPortrait
targetPageWidthPt = 518#
targetPageHeightPt = 745#
Else
chosenOrientation = xlLandscape
targetPageWidthPt = 762#
targetPageHeightPt = 500#
End If
If totalWidthPt > 40 And totalWidthPt < targetPageWidthPt * 0.78 Then
Dim expandRatio As Double
expandRatio = (targetPageWidthPt * 0.92) / totalWidthPt
If expandRatio > 1.35 Then expandRatio = 1.35
For cIdx = firstCol To lastCol
If Not ws.Columns(cIdx).Hidden Then
Dim newColW As Double
newColW = ws.Columns(cIdx).ColumnWidth * expandRatio
If newColW <= 75 Then ws.Columns(cIdx).ColumnWidth = newColW
End If
Next cIdx
totalWidthPt = GetVisibleColumnsWidthPt(ws, firstCol, lastCol, visibleCols)
End If
Dim printScale As Double
printScale = 1#
If totalWidthPt > targetPageWidthPt And totalWidthPt > 0 Then
printScale = targetPageWidthPt / totalWidthPt
If printScale < 0.45 Then printScale = 0.45
End If
Call AutoFitRowsAccurate(ws, firstRow, lastRow, firstCol, lastCol, hStart, hEnd, tableEndRow, hasHeader, printScale)
Dim totalHeightPt As Double
totalHeightPt = GetVisibleRowsHeightPt(ws, firstRow, lastRow)
Dim tableHeightPt As Double
tableHeightPt = GetVisibleRowsHeightPt(ws, firstRow, tableEndRow)
Dim scaledHeightPt As Double
scaledHeightPt = totalHeightPt * printScale
Dim scaledTableHeightPt As Double
scaledTableHeightPt = tableHeightPt * printScale
Dim targetTallPages As Long
targetTallPages = 0
If scaledHeightPt <= targetPageHeightPt * 1.50 Then
targetTallPages = 1
ElseIf tableEndRow < lastRow And scaledTableHeightPt <= targetPageHeightPt * 1.35 And scaledHeightPt <= targetPageHeightPt * 1.75 Then
targetTallPages = 1
Else
Dim fullPages As Long
fullPages = Int(scaledHeightPt / targetPageHeightPt)
Dim remPages As Double
remPages = (scaledHeightPt / targetPageHeightPt) - fullPages
If fullPages >= 1 And fullPages <= 12 Then
If remPages > 0 And remPages <= 0.45 Then
targetTallPages = fullPages
ElseIf tableEndRow < lastRow And Int(scaledTableHeightPt / targetPageHeightPt) <= fullPages And remPages <= 0.65 Then
targetTallPages = fullPages
End If
End If
End If
Dim zoomW As Double, zoomH As Double
zoomW = 100#
If totalWidthPt > targetPageWidthPt And totalWidthPt > 0 Then
zoomW = (targetPageWidthPt / totalWidthPt) * 98#
End If
zoomH = 100#
If targetTallPages > 0 And totalHeightPt > 0 Then
Dim availTallPt As Double
availTallPt = targetPageHeightPt * targetTallPages
If totalHeightPt > availTallPt Then
zoomH = (availTallPt / totalHeightPt) * 97#
End If
End If
Dim calcZoom As Long
If zoomH < zoomW Then
calcZoom = Int(zoomH)
Else
calcZoom = Int(zoomW)
End If
If calcZoom > 100 Then calcZoom = 100
If calcZoom < 45 Then calcZoom = 45
Call HT_InAn(False)
On Error Resume Next
With ws.PageSetup
.PaperSize = xlPaperA4
.Orientation = chosenOrientation
.LeftMargin = Application.CentimetersToPoints(1#)
.RightMargin = Application.CentimetersToPoints(1#)
.TopMargin = Application.CentimetersToPoints(1.2)
.BottomMargin = Application.CentimetersToPoints(1.2)
.HeaderMargin = Application.CentimetersToPoints(0.5)
.FooterMargin = Application.CentimetersToPoints(0.5)
.CenterHorizontally = True
.CenterVertically = False
.PrintGridlines = False
.PrintHeadings = False
.PrintArea = rng.Address
If hasHeader And targetTallPages <> 1 And hStart <= 10 Then
.PrintTitleRows = "$" & hStart & ":$" & hEnd
Else
.PrintTitleRows = ""
End If
.FitToPagesWide = False
.FitToPagesTall = False
.Zoom = calcZoom
End With
On Error GoTo LoiChuanHoa
Call HT_InAn(True)
Dim stepGuard As Long
stepGuard = 0
Do While stepGuard < 5 And calcZoom > 45
Dim needShrink As Boolean
needShrink = False
If ws.VPageBreaks.Count > 0 Then needShrink = True
If targetTallPages > 0 Then
If ws.HPageBreaks.Count >= targetTallPages Then needShrink = True
End If
If Not needShrink Then Exit Do
calcZoom = calcZoom - 3
If calcZoom < 45 Then calcZoom = 45
ws.PageSetup.Zoom = calcZoom
stepGuard = stepGuard + 1
Loop
Dim soTrang As Long
soTrang = 1
On Error Resume Next
soTrang = (ws.HPageBreaks.Count + 1) * (ws.VPageBreaks.Count + 1)
Dim uocTinh As Long
uocTinh = Int((GetVisibleRowsHeightPt(ws, firstRow, lastRow) * calcZoom / 100) / targetPageHeightPt) + 1
If uocTinh > soTrang Then soTrang = uocTinh
If soTrang > 1 And Len(CStr(mHT_PS(18))) = 0 Then
ws.PageSetup.CenterFooter = "&""Times New Roman""&11Trang &P/&N"
End If
On Error GoTo LoiChuanHoa
Dim minFont As Double
minFont = 0
Dim fc As Range
Dim scanned As Long
On Error Resume Next
For Each fc In rng.Cells
If Len(CStr(fc.Value)) > 0 Then
Dim fsz As Double
fsz = Val(CStr(fc.Font.Size))
If fsz > 0 And (minFont = 0 Or fsz < minFont) Then minFont = fsz
End If
scanned = scanned + 1
If scanned > 20000 Then Exit For
Next fc
On Error GoTo LoiChuanHoa
Dim chuNho As String
chuNho = ""
If minFont > 0 Then
If minFont * calcZoom / 100 < 8.5 Then
chuNho = vbCrLf & "- CANH BAO: chu in ra chi con khoang " & Format(minFont * calcZoom / 100, "0.0") & " pt (thu nho " & calcZoom & "%), co the kho doc." & _
vbCrLf & "  Nen chuyen kho ngang, bo bot cot hoac tach bang."
End If
End If
On Error Resume Next
Application.OnUndo "Hoan tac Chuan hoa trang in", "HoanTac_ChuanHoa_Excel"
On Error GoTo LoiChuanHoa
Application.ScreenUpdating = True
Application.DisplayAlerts = True
If Not silentMode Then
Dim orientName As String
If chosenOrientation = xlPortrait Then
orientName = "A4 Doc (Portrait)"
Else
orientName = "A4 Ngang (Landscape)"
End If
MsgBoxUni ChrW(272) & ChrW(195) & " CHU" & ChrW(7848) & "N H" & ChrW(211) & "A B" & ChrW(7842) & "NG BI" & ChrW(7874) & "U & TRANG IN HO" & ChrW(192) & "N H" & ChrW(7842) & "O!" & vbCrLf & _
"- Kh" & ChrW(7893) & " gi" & ChrW(7845) & "y t" & ChrW(7921) & " " & ChrW(273) & ChrW(7897) & "ng: " & orientName & vbCrLf & _
"- " & ChrW(272) & ChrW(227) & " c" & ChrW(259) & "n ch" & ChrW(7881) & "nh Fit trang in A4 & ch" & ChrW(7889) & "ng che khu" & ChrW(7845) & "t ch" & ChrW(7919) & " 100%." & chuNho & vbCrLf & vbCrLf & _
"B" & ChrW(7845) & "m Ctrl+Z (ho" & ChrW(7863) & "c n" & ChrW(250) & "t Ho" & ChrW(224) & "n t" & ChrW(225) & "c) " & ChrW(273) & ChrW(7875) & " tr" & ChrW(7843) & " l" & ChrW(7841) & "i nh" & ChrW(432) & " tr" & ChrW(432) & ChrW(7899) & "c khi chu" & ChrW(7849) & "n h" & ChrW(243) & "a." & vbCrLf & vbCrLf & _
"1Click Office - T" & ChrW(225) & "c gi" & ChrW(7843) & ": Tr" & ChrW(7847) & "n Qu" & ChrW(7889) & "c " & ChrW(272) & ChrW(432) & ChrW(7907) & "c (TTDVTH Ph" & ChrW(432) & ChrW(7901) & "ng T" & ChrW(226) & "n Khai) - " & ChrW(272) & "T: 0984.724.131", vbInformation, "Chu" & ChrW(7849) & "n H" & ChrW(243) & "a Trang In Excel - 1Click"
End If
On Error GoTo 0
Exit Sub
LoiChuanHoa:
Dim errMsg As String
errMsg = Err.Number & " - " & Err.Description
On Error Resume Next
Call HT_InAn(True)
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Dim daTra As Boolean
daTra = HT_KhoiPhuc()
On Error GoTo 0
If Not silentMode Then
If daTra Then
MsgBoxUni "C" & ChrW(243) & " l" & ChrW(7895) & "i khi chu" & ChrW(7849) & "n h" & ChrW(243) & "a: " & errMsg & vbCrLf & ChrW(272) & ChrW(227) & " tr" & ChrW(7843) & " b" & ChrW(7843) & "ng t" & ChrW(237) & "nh v" & ChrW(7873) & " nh" & ChrW(432) & " tr" & ChrW(432) & ChrW(7899) & "c khi chu" & ChrW(7849) & "n h" & ChrW(243) & "a.", vbExclamation, "Chu" & ChrW(7849) & "n H" & ChrW(243) & "a Excel"
Else
MsgBoxUni "C" & ChrW(243) & " l" & ChrW(7895) & "i khi chu" & ChrW(7849) & "n h" & ChrW(243) & "a: " & errMsg, vbExclamation, "Chu" & ChrW(7849) & "n H" & ChrW(243) & "a Excel"
End If
End If
End Sub
Private Function GetTrueDataBounds(ByVal ws As Worksheet, ByRef firstRow As Long, ByRef lastRow As Long, ByRef firstCol As Long, ByRef lastCol As Long) As Boolean
On Error Resume Next
GetTrueDataBounds = False
Dim lastRCell As Range, lastCCell As Range
Dim firstRCell As Range, firstCCell As Range
Set lastRCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
If lastRCell Is Nothing Then Exit Function
Set lastCCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
Set firstRCell = ws.Cells.Find(What:="*", After:=ws.Cells(lastRCell.Row, lastCCell.Column), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
Set firstCCell = ws.Cells.Find(What:="*", After:=ws.Cells(lastRCell.Row, lastCCell.Column), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
firstRow = 1
firstCol = 1
If Not firstRCell Is Nothing Then firstRow = firstRCell.Row
MsgBoxUni "C" & ChrW(243) & " l" & ChrW(7895) & "i khi chu" & ChrW(7849) & "n h" & ChrW(243) & "a: " & errMsg & vbCrLf & ChrW(272) & ChrW(227) & " tr" & ChrW(7843) & " b" & ChrW(7843) & "ng t" & ChrW(237) & "nh v" & ChrW(7873) & " nh" & ChrW(432) & " tr" & ChrW(432) & ChrW(7899) & "c khi chu" & ChrW(7849) & "n h" & ChrW(243) & "a.", vbExclamation, "Chu" & ChrW(7849) & "n H" & ChrW(243) & "a Excel"
If firstRow > 3 Then firstRow = 1
If firstCol > 2 Then firstCol = 1
lastRow = lastRCell.Row
lastCol = lastCCell.Column
Dim r As Long, c As Long
For r = firstRow To lastRow
Dim chkCell As Range
Set chkCell = ws.Cells(r, lastCol)
If chkCell.MergeCells Then
Dim mEndCol As Long
mEndCol = chkCell.MergeArea.Column + chkCell.MergeArea.Columns.Count - 1
If mEndCol > lastCol And mEndCol <= lastCol + 15 Then
lastCol = mEndCol
End If
End If
Next r
For c = firstCol To lastCol
Dim chkBotCell As Range
Set chkBotCell = ws.Cells(lastRow, c)
If chkBotCell.MergeCells Then
Dim mEndRow As Long
mEndRow = chkBotCell.MergeArea.Row + chkBotCell.MergeArea.Rows.Count - 1
If mEndRow > lastRow And mEndRow <= lastRow + 15 Then
lastRow = mEndRow
End If
End If
Next c
If lastRow > firstRow + 6 Then
Dim botNonEmpty As Long: botNonEmpty = 0
Dim botIsNum As Boolean: botIsNum = False
For c = firstCol To lastCol
If Not IsEmpty(ws.Cells(lastRow, c).Value) Then
botNonEmpty = botNonEmpty + 1
If IsNumeric(ws.Cells(lastRow, c).Value) And ws.Cells(lastRow, c).Borders(xlEdgeTop).LineStyle = xlNone Then
botIsNum = True
End If
End If
Next c
If botNonEmpty = 1 And botIsNum Then
Dim prevHasData As Boolean: prevHasData = False
For r = lastRow - 2 To lastRow - 1
For c = firstCol To lastCol
If Not IsEmpty(ws.Cells(r, c).Value) Then prevHasData = True
Next c
Next r
If Not prevHasData Then
For r = lastRow - 3 To firstRow Step -1
For c = firstCol To lastCol
If Not IsEmpty(ws.Cells(r, c).Value) Then
lastRow = r
GoTo DoneTrimBot
End If
Next c
Next r
End If
End If
End If
DoneTrimBot:
GetTrueDataBounds = (lastRow >= firstRow And lastCol >= firstCol)
On Error GoTo 0
End Function
Private Sub ApplySmartWrapText(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal firstCol As Long, ByVal lastCol As Long, ByVal hStart As Long, ByVal tableEndRow As Long, ByVal hasHeader As Boolean)
On Error Resume Next
Dim fullRng As Range
Set fullRng = ws.Range(ws.Cells(firstRow, firstCol), ws.Cells(lastRow, lastCol))
fullRng.WrapText = True
If lastCol <= firstCol Then Exit Sub
Dim r As Long, c As Long
For r = firstRow To lastRow
Dim isOutsideTable As Boolean
isOutsideTable = False
If hasHeader And r < hStart Then isOutsideTable = True
If r > tableEndRow Then isOutsideTable = True
If isOutsideTable And Not ws.Rows(r).Hidden Then
For c = firstCol To lastCol - 1
Dim cell As Range
Set cell = ws.Cells(r, c)
If Not cell.MergeCells Then
Dim txt As String
txt = Trim(CStr(cell.Value))
If Len(txt) > 0 And Len(txt) > ws.Columns(c).ColumnWidth Then
Dim nextCell As Range
Set nextCell = ws.Cells(r, c + 1)
If Not nextCell.MergeCells And Len(Trim(CStr(nextCell.Value))) = 0 Then
cell.WrapText = False
End If
End If
End If
Next c
End If
Next r
On Error GoTo 0
End Sub
Private Sub FitColumnsSmart(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal firstCol As Long, ByVal lastCol As Long, ByVal hStart As Long, ByVal tableEndRow As Long, ByVal hasHeader As Boolean)
On Error Resume Next
Dim c As Long, r As Long
Dim scanStart As Long
scanStart = IIf(hasHeader, hStart, firstRow)
For c = firstCol To lastCol
Dim colObj As Range
Set colObj = ws.Columns(c)
If Not colObj.Hidden Then
Dim curW As Double
curW = colObj.ColumnWidth
Dim maxHashNumLen As Long: maxHashNumLen = 0
Dim hasAnyText As Boolean: hasAnyText = False
For r = scanStart To tableEndRow
If Not ws.Rows(r).Hidden Then
Dim cell As Range
Set cell = ws.Cells(r, c)
If cell.MergeArea.Columns.Count = 1 Then
If Not IsEmpty(cell.Value) Then
hasAnyText = True
Dim dispTxt As String
dispTxt = Trim(CStr(cell.Text))
If Left(dispTxt, 1) = "#" Then
Dim cleanNumStr As String
If IsNumeric(cell.Value) Then
cleanNumStr = CStr(Round(CDbl(cell.Value), 2))
Else
cleanNumStr = Trim(CStr(cell.Value))
End If
If Len(cleanNumStr) > maxHashNumLen Then maxHashNumLen = Len(cleanNumStr)
End If
End If
End If
End If
Next r
If maxHashNumLen > 0 Then
Dim reqNumW As Double
reqNumW = (maxHashNumLen * 1.15) + 2.2
If reqNumW > 26 Then reqNumW = 26
If curW < reqNumW Then colObj.ColumnWidth = reqNumW
ElseIf hasAnyText And curW < 4.2 Then
colObj.ColumnWidth = 4.5
End If
End If
Next c
On Error GoTo 0
End Sub
Private Sub AutoFitRowsAccurate(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal firstCol As Long, ByVal lastCol As Long, ByVal hStart As Long, ByVal hEnd As Long, ByVal tableEndRow As Long, ByVal hasHeader As Boolean, ByVal printScale As Double)
On Error Resume Next
Dim fullRng As Range
Set fullRng = ws.Range(ws.Cells(firstRow, firstCol), ws.Cells(lastRow, lastCol))
fullRng.Rows.AutoFit
Dim scaleSafety As Double
scaleSafety = 1.04
If printScale < 0.9 Then
scaleSafety = 1.04 + ((0.9 - printScale) * 0.1)
If scaleSafety > 1.08 Then scaleSafety = 1.08
End If
Dim r As Long, c As Long
For r = firstRow To lastRow
Dim rowObj As Range
Set rowObj = ws.Rows(r)
If Not rowObj.Hidden Then
Dim rowHasText As Boolean
rowHasText = False
Dim maxRowFontSize As Double
maxRowFontSize = 0#
For c = firstCol To lastCol
Dim scanCell As Range
Set scanCell = ws.Cells(r, c)
If Not IsEmpty(scanCell.Value) Then
If Len(Trim(CStr(scanCell.Value))) > 0 Then
rowHasText = True
Dim cfs As Double
cfs = Val(CStr(scanCell.Font.Size))
If cfs >= 6 And cfs <= 36 And cfs > maxRowFontSize Then
maxRowFontSize = cfs
End If
End If
End If
Next c
If Not rowHasText Then
If r > tableEndRow And rowObj.RowHeight > 15.5 Then
rowObj.RowHeight = 14.5
ElseIf rowObj.RowHeight > 20# Then
rowObj.RowHeight = 16#
End If
Else
If maxRowFontSize < 6 Then maxRowFontSize = 11#
Dim minRowH As Double
minRowH = (maxRowFontSize * 1.35) + 4#
If hasHeader And r >= hStart And r <= hEnd Then
If minRowH < (maxRowFontSize * 1.45) + 5# Then
minRowH = (maxRowFontSize * 1.45) + 5#
End If
End If
Dim maxReqHeight As Double
maxReqHeight = minRowH
Dim hasVertMergeStart As Boolean
hasVertMergeStart = False
For c = firstCol To lastCol
Dim chkV As Range
Set chkV = ws.Cells(r, c)
If chkV.MergeCells Then
If chkV.MergeArea.Row = r And chkV.MergeArea.Rows.Count > 1 Then
If Len(Trim(CStr(chkV.MergeArea.Cells(1, 1).Value))) > 12 Then
hasVertMergeStart = True
Exit For
End If
End If
End If
Next c
If Not hasVertMergeStart Then
If rowObj.RowHeight > maxReqHeight And rowObj.RowHeight <= 400 Then
maxReqHeight = rowObj.RowHeight + 3#
End If
End If
For c = firstCol To lastCol
Dim cell As Range
Set cell = ws.Cells(r, c)
If cell.Address = cell.MergeArea.Cells(1, 1).Address Then
If cell.MergeArea.Rows.Count = 1 Then
Dim cellReqH As Double
cellReqH = CalcCellRequiredHeightPt(cell, scaleSafety)
If cellReqH > maxReqHeight Then
maxReqHeight = cellReqH
End If
End If
End If
Next c
If maxReqHeight > 409# Then maxReqHeight = 409#
rowObj.RowHeight = maxReqHeight
End If
End If
Next r
Dim targetSpan As Long
For targetSpan = 2 To 12
For r = firstRow To lastRow
If Not ws.Rows(r).Hidden Then
For c = firstCol To lastCol
Dim mCell As Range
Set mCell = ws.Cells(r, c)
If mCell.MergeCells Then
Dim mRowSpan As Long
mRowSpan = mCell.MergeArea.Rows.Count
If mCell.Address = mCell.MergeArea.Cells(1, 1).Address And _
((targetSpan < 12 And mRowSpan = targetSpan) Or (targetSpan = 12 And mRowSpan >= 12)) Then
Dim totalReqH As Double
totalReqH = CalcCellRequiredHeightPt(mCell, scaleSafety)
If totalReqH > 0 Then
Dim mStartR As Long, mEndR As Long, mr As Long
mStartR = mCell.MergeArea.Row
mEndR = mStartR + mRowSpan - 1
Dim curTotalH As Double: curTotalH = 0
Dim visRowsInMerge As Long: visRowsInMerge = 0
For mr = mStartR To mEndR
If Not ws.Rows(mr).Hidden Then
curTotalH = curTotalH + ws.Rows(mr).RowHeight
visRowsInMerge = visRowsInMerge + 1
End If
Next mr
If visRowsInMerge > 0 And curTotalH < totalReqH Then
Dim addPerRow As Double
addPerRow = (totalReqH - curTotalH) / visRowsInMerge
For mr = mStartR To mEndR
If Not ws.Rows(mr).Hidden Then
Dim newH As Double
newH = ws.Rows(mr).RowHeight + addPerRow
If newH > 409# Then newH = 409#
ws.Rows(mr).RowHeight = newH
End If
Next mr
End If
End If
End If
End If
Next c
End If
Next r
Next targetSpan
On Error GoTo 0
End Sub
Private Function CalcCellRequiredHeightPt(ByVal cell As Range, ByVal scaleSafety As Double) As Double
On Error Resume Next
CalcCellRequiredHeightPt = 0
If IsEmpty(cell.Value) Then Exit Function
Dim valStr As String
If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
Dim dispNum As String
dispNum = Trim(CStr(cell.Text))
If Len(dispNum) > 0 And Left(dispNum, 1) <> "#" Then
valStr = dispNum
Else
valStr = CStr(Round(CDbl(cell.Value), 2))
End If
Else
valStr = Trim(CStr(cell.Value))
End If
If Len(valStr) = 0 Then Exit Function
valStr = Replace(valStr, ChrW(8230), "...")
Dim usableWidthPt As Double
usableWidthPt = GetVisibleMergeAreaWidthPt(cell.MergeArea) - 10#
If usableWidthPt < 24# Then usableWidthPt = 24#
Dim fSize As Double
fSize = Val(CStr(cell.Font.Size))
If fSize < 6 Or fSize > 36 Then fSize = 12#
Dim isBold As Boolean
isBold = False
If IsNull(cell.Font.Bold) Then
isBold = True
ElseIf CBool(cell.Font.Bold) = True Then
isBold = True
End If
Dim charWidthPt As Double
charWidthPt = fSize * 0.47
If isBold Then charWidthPt = charWidthPt * 1.1
If valStr = UCase(valStr) And Len(valStr) > 6 Then charWidthPt = charWidthPt * 1.12
charWidthPt = charWidthPt * scaleSafety
Dim maxCharsPerLine As Long
maxCharsPerLine = Int(usableWidthPt / charWidthPt)
If maxCharsPerLine < 4 Then maxCharsPerLine = 4
Dim normalizedTxt As String
normalizedTxt = Replace(valStr, vbCrLf, vbLf)
normalizedTxt = Replace(normalizedTxt, vbCr, vbLf)
Dim hardLines() As String
hardLines = Split(normalizedTxt, vbLf)
Dim totalLines As Long
totalLines = 0
If Not cell.MergeCells And CBool(cell.WrapText) = False Then
totalLines = UBound(hardLines) - LBound(hardLines) + 1
Else
Dim k As Long
For k = LBound(hardLines) To UBound(hardLines)
Dim pLine As String
pLine = Trim(hardLines(k))
If Len(pLine) = 0 Then
totalLines = totalLines + 1
Else
totalLines = totalLines + CountWordWrappedLines(pLine, maxCharsPerLine)
End If
Next k
End If
If totalLines < 1 Then totalLines = 1
Dim lineHeightPt As Double
lineHeightPt = fSize * 1.35
CalcCellRequiredHeightPt = (totalLines * lineHeightPt) + 5#
On Error GoTo 0
End Function
Private Function CountWordWrappedLines(ByVal paragraphText As String, ByVal maxCharsPerLine As Long) As Long
Dim words() As String
words = Split(paragraphText, " ")
Dim linesCount As Long: linesCount = 1
Dim curLineLen As Long: curLineLen = 0
Dim i As Long
For i = LBound(words) To UBound(words)
Dim w As String
w = words(i)
Dim wLen As Long
wLen = Len(w)
If wLen > 0 Then
If curLineLen = 0 Then
If wLen <= maxCharsPerLine Then
curLineLen = wLen
Else
Dim extraLines As Long
extraLines = Int((wLen - 1) / maxCharsPerLine)
linesCount = linesCount + extraLines
curLineLen = wLen - (extraLines * maxCharsPerLine)
End If
ElseIf curLineLen + 1 + wLen <= maxCharsPerLine Then
curLineLen = curLineLen + 1 + wLen
Else
linesCount = linesCount + 1
If wLen <= maxCharsPerLine Then
curLineLen = wLen
Else
Dim extraWLines As Long
extraWLines = Int((wLen - 1) / maxCharsPerLine)
linesCount = linesCount + extraWLines
curLineLen = wLen - (extraWLines * maxCharsPerLine)
End If
End If
End If
Next i
CountWordWrappedLines = linesCount
End Function
Private Function GetVisibleMergeAreaWidthPt(ByVal mergeRng As Range) As Double
On Error Resume Next
Dim totalPt As Double: totalPt = 0
Dim mCol As Range
For Each mCol In mergeRng.Columns
If Not mCol.Hidden Then
totalPt = totalPt + mCol.Width
End If
Next mCol
GetVisibleMergeAreaWidthPt = totalPt
On Error GoTo 0
End Function
Private Function GetVisibleColumnsWidthPt(ByVal ws As Worksheet, ByVal firstCol As Long, ByVal lastCol As Long, ByRef visibleColsCount As Long) As Double
On Error Resume Next
Dim totalPt As Double: totalPt = 0
visibleColsCount = 0
Dim c As Long
For c = firstCol To lastCol
If Not ws.Columns(c).Hidden Then
totalPt = totalPt + ws.Columns(c).Width
visibleColsCount = visibleColsCount + 1
End If
Next c
GetVisibleColumnsWidthPt = totalPt
On Error GoTo 0
End Function
Private Function GetVisibleRowsHeightPt(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long) As Double
On Error Resume Next
Dim totalPt As Double: totalPt = 0
Dim r As Long
For r = firstRow To lastRow
If Not ws.Rows(r).Hidden Then
totalPt = totalPt + ws.Rows(r).RowHeight
End If
Next r
GetVisibleRowsHeightPt = totalPt
On Error GoTo 0
End Function
Private Function FindTableDataEndRow(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal firstCol As Long, ByVal lastCol As Long, ByVal hEnd As Long, ByVal hasHeader As Boolean) As Long
FindTableDataEndRow = lastRow
Dim minCheckRow As Long
minCheckRow = IIf(hasHeader, hEnd + 1, firstRow + 1)
If lastRow <= minCheckRow Then Exit Function
Dim startScanFooter As Long
startScanFooter = lastRow - 18
If startScanFooter < minCheckRow Then startScanFooter = minCheckRow
Dim r As Long, c As Long
For r = startScanFooter To lastRow
Dim rowTextAsc As String: rowTextAsc = ""
Dim nonEmptyCells As Long: nonEmptyCells = 0
For c = firstCol To lastCol
If ws.Cells(r, c).Address = ws.Cells(r, c).MergeArea.Cells(1, 1).Address Then
Dim cellVal As String
cellVal = Trim(CStr(ws.Cells(r, c).Value))
If Len(cellVal) > 0 Then
nonEmptyCells = nonEmptyCells + 1
rowTextAsc = rowTextAsc & " " & ToAscii(cellVal)
End If
End If
Next c
If nonEmptyCells > 0 And nonEmptyCells <= 4 Then
If InStr(rowTextAsc, "T/M ") > 0 Or InStr(rowTextAsc, "TM.") > 0 Or _
InStr(rowTextAsc, "K/T ") > 0 Or InStr(rowTextAsc, "KT.") > 0 Or _
InStr(rowTextAsc, "T/L ") > 0 Or InStr(rowTextAsc, "TL.") > 0 Or _
InStr(rowTextAsc, "NGUOI LAP") > 0 Or InStr(rowTextAsc, "LAP BIEU") > 0 Or _
InStr(rowTextAsc, "BI THU") > 0 Or InStr(rowTextAsc, "PHO BI THU") > 0 Or _
InStr(rowTextAsc, "CHU TICH") > 0 Or InStr(rowTextAsc, "TRUONG BAN") > 0 Or _
InStr(rowTextAsc, "CHANH VAN PHONG") > 0 Or InStr(rowTextAsc, "VAN PHONG") > 0 Or _
InStr(rowTextAsc, "THU TRUONG") > 0 Or InStr(rowTextAsc, "GIAM DOC") > 0 Or _
InStr(rowTextAsc, "XAC NHAN") > 0 Or InStr(rowTextAsc, "DUYET") > 0 Or _
InStr(rowTextAsc, "NOI NHAN") > 0 Or _
(InStr(rowTextAsc, "NGAY") > 0 And InStr(rowTextAsc, "THANG") > 0 And InStr(rowTextAsc, "NAM") > 0) Then
If r - 1 >= minCheckRow Then
FindTableDataEndRow = r - 1
Exit Function
End If
End If
End If
Next r
End Function
Private Function FindTableHeaderRows(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal startCol As Long, ByVal endCol As Long, ByRef hStart As Long, ByRef hEnd As Long) As Boolean
Dim r As Long, c As Long
Dim maxScanRow As Long
maxScanRow = startRow + 25
If maxScanRow > endRow Then maxScanRow = endRow
Dim bestRow As Long: bestRow = 0
Dim maxScore As Long: maxScore = 0
For r = startRow To maxScanRow
Dim colCountWithText As Long: colCountWithText = 0
Dim coveredCols As Long: coveredCols = 0
Dim rowAscText As String: rowAscText = ""
For c = startCol To endCol
If ws.Cells(r, c).Address = ws.Cells(r, c).MergeArea.Cells(1, 1).Address Then
Dim cellVal As String
cellVal = Trim(CStr(ws.Cells(r, c).Value))
If Len(cellVal) > 0 Then
colCountWithText = colCountWithText + 1
coveredCols = coveredCols + ws.Cells(r, c).MergeArea.Columns.Count
rowAscText = rowAscText & " " & ToAscii(cellVal)
End If
End If
Next c
Dim isHeaderCandidate As Boolean: isHeaderCandidate = False
If colCountWithText >= 3 Then
isHeaderCandidate = True
ElseIf colCountWithText >= 2 And coveredCols >= 3 Then
isHeaderCandidate = True
End If
If isHeaderCandidate Then
Dim isTieuNgu As Boolean: isTieuNgu = False
If InStr(rowAscText, "DANG CONG SAN") > 0 Or _
InStr(rowAscText, "CONG HOA XA HOI") > 0 Or _
InStr(rowAscText, "DOC LAP") > 0 Then
isTieuNgu = True
ElseIf InStr(rowAscText, "DANG BO") > 0 Or InStr(rowAscText, "DANG UY") > 0 Or InStr(rowAscText, "UY BAN") > 0 Then
isTieuNgu = True
ElseIf InStr(rowAscText, "NGAY") > 0 And InStr(rowAscText, "THANG") > 0 And InStr(rowAscText, "NAM") > 0 Then
isTieuNgu = True
End If
If Not isTieuNgu Then
Dim isDocTitle As Boolean: isDocTitle = False
If colCountWithText <= 2 Then
If InStr(rowAscText, "DANH MUC") > 0 Or _
InStr(rowAscText, "PHU LUC") > 0 Or _
InStr(rowAscText, "KE HOACH") > 0 Or _
InStr(rowAscText, "BAO CAO") > 0 Or _
InStr(rowAscText, "DANH SACH") > 0 Or _
InStr(rowAscText, "BANG TONG HOP") > 0 Or _
InStr(rowAscText, "BANG PHAN CONG") > 0 Or _
InStr(rowAscText, "QUYET DINH") > 0 Or _
InStr(rowAscText, "THONG BAO") > 0 Then
isDocTitle = True
End If
End If
If Not isDocTitle Then
Dim currentScore As Long
currentScore = colCountWithText * 10
Dim kws As Variant
kws = Array("STT", "TT", "TRUC", "NHOM", "NOI DUNG", "NHIEM VU", "SAN PHAM", _
"DON VI", "CHI TIEU", "THOI GIAN", "PHAN CONG", "HO VA TEN", "HO TEN", _
"CHUC VU", "SO LUONG", "GHI CHU", "TIEN DO", "KET QUA", _
"MO TA", "CHI TIET", "DIEN GIAI", "DON GIA", "THANH TIEN", "NGUOI THUC HIEN", _
"DU TOAN", "DIEU CHINH", "TANG", "GIAM", "KINH PHI")
Dim kwIdx As Long, kwHits As Long
kwHits = 0
For kwIdx = LBound(kws) To UBound(kws)
If InStr(rowAscText, CStr(kws(kwIdx))) > 0 Then
kwHits = kwHits + 1
End If
Next kwIdx
currentScore = currentScore + (kwHits * 20)
If (kwHits >= 1 Or colCountWithText >= 4) And currentScore >= 50 Then
maxScore = currentScore
bestRow = r
Exit For
End If
End If
End If
End If
Next r
If bestRow = 0 Or maxScore < 40 Then
FindTableHeaderRows = False
Exit Function
End If
hStart = bestRow
hEnd = bestRow
Dim maxMergeRow As Long
maxMergeRow = bestRow
For c = startCol To endCol
If ws.Cells(bestRow, c).MergeCells Then
Dim mRows As Long
mRows = ws.Cells(bestRow, c).MergeArea.Row + ws.Cells(bestRow, c).MergeArea.Rows.Count - 1
If mRows > maxMergeRow Then maxMergeRow = mRows
End If
Next c
If maxMergeRow > hEnd Then hEnd = maxMergeRow
If hEnd < endRow Then
Dim nextRow As Long
nextRow = hEnd + 1
Dim numColsWithNumber As Long: numColsWithNumber = 0
For c = startCol To endCol
Dim nVal As String
nVal = Trim(CStr(ws.Cells(nextRow, c).Value))
If Len(nVal) > 0 Then
If nVal Like "(*)" Or IsNumeric(nVal) Or nVal = "a" Or nVal = "b" Or InStr(nVal, "=") > 0 Then
numColsWithNumber = numColsWithNumber + 1
End If
End If
Next c
If numColsWithNumber >= 3 Then
hEnd = nextRow
End If
End If
FindTableHeaderRows = True
End Function
Private Sub ClampLastColToTableHeader(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal firstCol As Long, ByRef lastCol As Long, ByVal hStart As Long, ByVal hEnd As Long)
On Error Resume Next
If lastCol <= firstCol + 1 Then Exit Sub
Dim trueTableLastCol As Long
trueTableLastCol = 0
Dim c As Long
For c = firstCol To lastCol
Dim topHdrCell As Range
Set topHdrCell = ws.Cells(hStart, c).MergeArea.Cells(1, 1)
Dim hdrTxt As String
hdrTxt = Trim(CStr(topHdrCell.Value))
If Len(hdrTxt) = 0 Then
Exit For
End If
Dim hdrAsc As String
hdrAsc = ToAscii(hdrTxt)
If InStr(hdrAsc, "KEM THEO") > 0 Then
Exit For
End If
trueTableLastCol = c
Next c
If trueTableLastCol >= firstCol + 1 And trueTableLastCol < lastCol Then
lastCol = trueTableLastCol
End If
On Error GoTo 0
End Sub
Private Function LicApiUrl() As String
LicApiUrl = ""
End Function
Private Function CheckLicenseActivation() As Boolean
    CheckLicenseActivation = True
End Function
Private Function CallHttpApi(ByVal url As String) As String
Dim res As String
res = ""
#If Mac Then
On Error Resume Next
Dim sc As String
sc = "do shell script ""curl -s -L --max-time 8 '" & url & "'"""
res = MacScript(sc)
On Error GoTo 0
#Else
On Error Resume Next
Dim http As Object
Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
If http Is Nothing Then
Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
End If
http.Open "GET", url, False
http.setRequestHeader "User-Agent", "Mozilla/5.0"
http.setTimeouts 5000, 8000, 8000, 8000
http.Send
If Err.Number = 0 Then
res = http.responseText
End If
On Error GoTo 0
#End If
CallHttpApi = Trim(res)
End Function
Private Function GetMachineName() As String
Dim mName As String
#If Mac Then
On Error Resume Next
mName = MacScript("do shell script ""scutil --get ComputerName || hostname -s || whoami""")
mName = Trim(Replace(mName, vbCrLf, ""))
mName = Trim(Replace(mName, vbLf, ""))
If Len(mName) = 0 Then mName = Environ("USER")
On Error GoTo 0
#Else
mName = Environ("COMPUTERNAME")
If Len(mName) = 0 Then mName = Environ("USERNAME")
#End If
If Len(mName) = 0 Then mName = "UNKNOWN_PC"
GetMachineName = mName
End Function
Private Function ParseJsonField(ByVal json As String, ByVal fieldName As String) As String
Dim pos As Long, startPos As Long, endPos As Long
pos = InStr(json, """" & fieldName & """")
If pos > 0 Then
startPos = InStr(pos + Len(fieldName) + 2, json, """")
If startPos > 0 Then
endPos = InStr(startPos + 1, json, """")
If endPos > startPos Then
ParseJsonField = Mid(json, startPos + 1, endPos - startPos - 1)
End If
End If
End If
End Function
Private Function EncodeUrl(ByVal s As String) As String
Dim i As Long, ch As String, code As Long
Dim res As String
res = ""
For i = 1 To Len(s)
ch = Mid(s, i, 1)
code = AscW(ch) And &HFFFF&
If (code >= 48 And code <= 57) Or (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Or ch = "-" Or ch = "_" Or ch = "." Then
res = res & ch
ElseIf code < 128 Then
res = res & "%" & Right("0" & Hex(code), 2)
Else
res = res & "_"
End If
Next i
EncodeUrl = res
End Function
Private Function ToAscii(ByVal s As String) As String
Dim n As Long
n = Len(s)
If n = 0 Then
ToAscii = ""
Exit Function
End If
Dim res As String
res = Space$(n)
Dim k As Long, w As Long
For k = 1 To n
w = AscW(Mid$(s, k, 1))
If w < 0 Then w = w + 65536
Select Case w
Case 272, 273: Mid$(res, k, 1) = "D"
Case 192 To 195, 224 To 227, 258, 259, 7840 To 7863: Mid$(res, k, 1) = "A"
Case 200 To 202, 232 To 234, 278, 279, 7864 To 7879: Mid$(res, k, 1) = "E"
Case 204, 205, 236, 237, 296, 297, 7880 To 7883: Mid$(res, k, 1) = "I"
Case 210 To 214, 242 To 246, 416, 417, 7884 To 7907: Mid$(res, k, 1) = "O"
Case 217, 218, 249, 250, 360, 361, 431, 432, 7908 To 7921: Mid$(res, k, 1) = "U"
Case 221, 253, 7922 To 7929: Mid$(res, k, 1) = "Y"
Case Else
If w >= 97 And w <= 122 Then
Mid$(res, k, 1) = Chr$(w - 32)
ElseIf (w >= 65 And w <= 90) Or (w >= 48 And w <= 57) Or w = 32 Or w = 42 Or w = 45 Or w = 47 Then
Mid$(res, k, 1) = Chr$(w)
Else
Mid$(res, k, 1) = " "
End If
End Select
Next k
ToAscii = res
End Function
Public Sub Auto_Open()
' 1Click Excel Auto initialized
End Sub
Public Sub RibbonOnLoad(ribbon As Object)
' 1Click Excel Auto initialized
End Sub
Public Function CN_PhienBan() As String
CN_PhienBan = "2026.11.17"
End Function
Private Function CN_ManifestUrl() As String
CN_ManifestUrl = "https://raw.githubusercontent.com/kawai-kun/1Click-Office/main/version.json"
End Function
Private Function CN_SoSanh(ByVal a As String, ByVal b As String) As Integer
Dim pa() As String, pb() As String
pa = Split(a, ".")
pb = Split(b, ".")
Dim i As Integer, va As Long, vb As Long
For i = 0 To 5
va = 0: vb = 0
If i <= UBound(pa) Then va = Val(pa(i))
If i <= UBound(pb) Then vb = Val(pb(i))
If va > vb Then
CN_SoSanh = 1
Exit Function
ElseIf va < vb Then
CN_SoSanh = -1
Exit Function
End If
Next i
CN_SoSanh = 0
End Function
Private Function CN_FileTuChay() As String
On Error Resume Next
CN_FileTuChay = ThisWorkbook.FullName
On Error GoTo 0
End Function
Private Function CN_DungFileAddIn() As Boolean
Dim n As String
n = LCase(Mid(CN_FileTuChay(), InStrRev(Replace(CN_FileTuChay(), "\", "/"), "/") + 1))
CN_DungFileAddIn = (n = "chuanhoa_excel.xlam")
End Function
Private Function CN_ShQ(ByVal s As String) As String
CN_ShQ = "'" & Replace(s, "'", "'\''") & "'"
End Function
Private Function CN_TaiManifest(Optional ByVal nhanh As Boolean = False) As String
Dim i As Integer, res As String, soLan As Integer, cho As Integer
soLan = IIf(nhanh, 1, 2)
cho = IIf(nhanh, 3, 15)
For i = 1 To soLan
res = ""
#If Mac Then
On Error Resume Next
res = MacScript("do shell script ""curl -s -L --max-time " & cho & " " & CN_ShQ(CN_ManifestUrl()) & """")
On Error GoTo 0
#Else
On Error Resume Next
Dim http As Object
Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
http.Open "GET", CN_ManifestUrl(), False
http.setRequestHeader "User-Agent", "Mozilla/5.0"
http.setTimeouts 5000, 5000, cho * 1000, cho * 1000
http.Send
If Err.Number = 0 Then
If http.Status = 200 Then res = http.responseText
End If
On Error GoTo 0
#End If
res = Trim(res)
If Len(ParseJsonField(res, "version")) > 0 Then Exit For
Next i
CN_TaiManifest = res
End Function
Public Sub RibbonCapNhat(control As Object)
    Dim s As String
    s = "Ph" & ChrW(7847) & "n m" & ChrW(7873) & "m 1Click Excel - Chu" & ChrW(7849) & "n H" & ChrW(243) & "a Trang In & B" & ChrW(7843) & "ng Bi" & ChrW(7875) & "u (v3.0 Enterprise)." & vbCrLf & _
        "- T" & ChrW(225) & "c gi" & ChrW(7843) & ": Tr" & ChrW(7847) & "n Qu" & ChrW(7889) & "c " & ChrW(272) & ChrW(432) & ChrW(7907) & "c (TTDVTH Ph" & ChrW(432) & ChrW(7901) & "ng T" & ChrW(226) & "n Khai)" & vbCrLf & _
        "- " & ChrW(272) & "i" & ChrW(7879) & "n tho" & ChrW(7841) & "i / Zalo: 0984.724.131" & vbCrLf & _
        "- Tr" & ChrW(7841) & "ng th" & ChrW(225) & "i: " & ChrW(272) & ChrW(227) & " k" & ChrW(237) & "ch ho" & ChrW(7841) & "t V" & ChrW(296) & "NH VI" & ChrW(7876) & "N tr" & ChrW(234) & "n m" & ChrW(225) & "y t" & ChrW(237) & "nh n" & ChrW(224) & "y."
    MsgBoxUni s, vbInformation, "1Click Excel - Th" & ChrW(244) & "ng Tin T" & ChrW(225) & "c Gi" & ChrW(7843)
End Sub
Public Sub CN_HenKiemTraKhiMo()
Static daHen As Boolean
If daHen Then Exit Sub
daHen = True
#If Mac Then
Call CN_TuKiemTraHangTuan(True)
#Else
On Error Resume Next
Application.OnTime Now + TimeSerial(0, 0, 20), "CN_KiemTraNhanhKhiMo"
#End If
End Sub
Public Sub CN_KiemTraNhanhKhiMo()
Call CN_TuKiemTraHangTuan(True)
End Sub
Public Sub CN_TuKiemTraHangTuan(Optional ByVal nhanh As Boolean = False)
On Error Resume Next
Dim lastChk As String, daDocDuoc As Boolean
lastChk = GetSetting("OfficePartySuite", "CapNhat", "LanKiemTra_EXCEL", "")
If Len(lastChk) > 0 Then
If IsDate(lastChk) Then
If DateDiff("d", CDate(lastChk), Date) < 1 Then Exit Sub
End If
End If
Dim lanThu As String
lanThu = GetSetting("OfficePartySuite", "CapNhat", "LanThu_EXCEL", "")
If Len(lanThu) > 0 Then
If IsDate(lanThu) Then
If DateDiff("n", CDate(lanThu), Now) < 60 Then Exit Sub
End If
End If
SaveSetting "OfficePartySuite", "CapNhat", "LanThu_EXCEL", Format(Now, "yyyy-mm-dd hh:nn:ss")
On Error GoTo 0
Call CN_KiemTraCapNhat(False, nhanh, daDocDuoc)
If daDocDuoc Then SaveSetting "OfficePartySuite", "CapNhat", "LanKiemTra_EXCEL", Format(Date, "yyyy-mm-dd")
End Sub
Public Sub CN_KiemTraCapNhat(ByVal hoiKhiKhongCo As Boolean, Optional ByVal nhanh As Boolean = False, Optional ByRef daDocDuoc As Boolean = False)
Dim js As String
js = CN_TaiManifest(nhanh)
Dim ver As String, url As String, sha As String, notes As String, kt As Long
ver = ParseJsonField(js, "version")
url = ParseJsonField(js, "excel_url")
sha = LCase(ParseJsonField(js, "excel_sha256"))
notes = ParseJsonField(js, "notes_ascii")
kt = Val(ParseJsonField(js, "excel_size"))
#If Mac Then
#Else
If Len(ParseJsonField(js, "excel_url_win")) > 0 Then
url = ParseJsonField(js, "excel_url_win")
sha = LCase(ParseJsonField(js, "excel_sha256_win"))
kt = Val(ParseJsonField(js, "excel_size_win"))
End If
#End If
If Len(ver) = 0 Or Len(url) = 0 Or Len(sha) <> 64 Then
If hoiKhiKhongCo Then MsgBoxUni "Kh" & ChrW(244) & "ng k" & ChrW(7871) & "t n" & ChrW(7889) & "i " & ChrW(273) & ChrW(432) & ChrW(7907) & "c m" & ChrW(225) & "y ch" & ChrW(7911) & " c" & ChrW(7853) & "p nh" & ChrW(7853) & "t (GitHub). Vui l" & ChrW(242) & "ng ki" & ChrW(7875) & "m tra m" & ChrW(7841) & "ng Internet r" & ChrW(7891) & "i th" & ChrW(7917) & " l" & ChrW(7841) & "i." & vbCrLf & "Phi" & ChrW(234) & "n b" & ChrW(7843) & "n " & ChrW(273) & "ang d" & ChrW(249) & "ng: " & CN_PhienBan(), vbExclamation, "C" & ChrW(7853) & "p Nh" & ChrW(7853) & "t Add-in"
Exit Sub
End If
daDocDuoc = True
If CN_SoSanh(ver, CN_PhienBan()) <= 0 Then
If hoiKhiKhongCo Then MsgBoxUni "B" & ChrW(7841) & "n " & ChrW(273) & "ang d" & ChrW(249) & "ng phi" & ChrW(234) & "n b" & ChrW(7843) & "n m" & ChrW(7899) & "i nh" & ChrW(7845) & "t: " & CN_PhienBan() & " Enterprise.", vbInformation, "Ki" & ChrW(7875) & "m Tra C" & ChrW(7853) & "p Nh" & ChrW(7853) & "t"
Exit Sub
End If
If Not CN_DungFileAddIn() Then
MsgBoxUni "C" & ChrW(243) & " phi" & ChrW(234) & "n b" & ChrW(7843) & "n m" & ChrW(7899) & "i " & ver & ". Vui l" & ChrW(242) & "ng ch" & ChrW(7841) & "y b" & ChrW(7897) & " c" & ChrW(224) & "i " & ChrW(273) & ChrW(7863) & "t " & ChrW(273) & ChrW(7875) & " c" & ChrW(7853) & "p nh" & ChrW(7853) & "t.", vbInformation, "C" & ChrW(7853) & "p Nh" & ChrW(7853) & "t Add-in"
Exit Sub
End If
Dim ans As VbMsgBoxResult
ans = MsgBoxUni("C" & ChrW(243) & " phi" & ChrW(234) & "n b" & ChrW(7843) & "n m" & ChrW(7899) & "i: " & ver & " (" & ChrW(273) & "ang d" & ChrW(249) & "ng " & CN_PhienBan() & ")." & vbCrLf & vbCrLf & notes & vbCrLf & vbCrLf & "B" & ChrW(7841) & "n c" & ChrW(243) & " mu" & ChrW(7889) & "n c" & ChrW(7853) & "p nh" & ChrW(7853) & "t ngay kh" & ChrW(244) & "ng?", vbYesNo + vbQuestion, "C" & ChrW(7853) & "p Nh" & ChrW(7853) & "t 1Click Office")
If ans <> vbYes Then Exit Sub
Dim msg As String
msg = CN_TaiVaCai(url, sha, kt)
If Len(msg) > 0 Then MsgBoxUni msg, vbInformation, "C" & ChrW(7853) & "p Nh" & ChrW(7853) & "t Add-in"
End Sub
Private Function CN_TaiVaCai(ByVal url As String, ByVal sha As String, Optional ByVal kichThuoc As Long = 0) As String
Dim target As String
target = CN_FileTuChay()
#If Mac Then
Dim tmp As String
tmp = Environ("HOME") & "/cn_capnhat_EXCEL.tmp"
Dim got As String
On Error Resume Next
got = MacScript("do shell script """ & "curl -s -L --fail --max-time 180 -o " & CN_ShQ(tmp) & " " & CN_ShQ(url) & _
" && shasum -a 256 " & CN_ShQ(tmp) & " | cut -d' ' -f1" & """")
On Error GoTo 0
If LCase(Trim(got)) <> sha Then
On Error Resume Next
MacScript "do shell script """ & "rm -f " & CN_ShQ(tmp) & """"
On Error GoTo 0
CN_TaiVaCai = "Tai ban cap nhat that bai hoac file khong khop ma kiem tra (SHA-256). Chua thay doi gi."
Exit Function
End If
Dim okCp As String
On Error Resume Next
okCp = MacScript("do shell script """ & "cp -f " & CN_ShQ(tmp) & " " & CN_ShQ(target) & " && xattr -c " & CN_ShQ(target) & _
" ; rm -f " & CN_ShQ(tmp) & " ; echo OK" & """")
On Error GoTo 0
If InStr(1, okCp, "OK") = 0 Then
CN_TaiVaCai = "Khong ghi duoc file add-in moi. Hay cai lai bang bo cai dat."
Exit Function
End If
CN_TaiVaCai = "Da cap nhat xong. Hay THOAT han Excel (Cmd+Q) roi mo lai de dung phien ban moi."
#Else
On Error GoTo WinFail
Dim http As Object
Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
http.Open "GET", url, False
http.setRequestHeader "User-Agent", "Mozilla/5.0"
http.setTimeouts 10000, 15000, 60000, 180000
http.Send
If http.Status <> 200 Then GoTo WinFail
Dim b() As Byte, n As Long
b = http.responseBody
n = UBound(b) - LBound(b) + 1
If n < 10240 Then GoTo WinBad
If b(LBound(b)) <> 80 Or b(LBound(b) + 1) <> 75 Then GoTo WinBad
If kichThuoc > 0 And n <> kichThuoc Then GoTo WinBad
Dim thuMuc As String
thuMuc = Environ("APPDATA") & "\OfficePartySuite"
If Len(Dir(thuMuc, vbDirectory)) = 0 Then MkDir thuMuc
Dim moiW As String
moiW = thuMuc & "\capnhat_EXCEL.moi"
If Len(Dir(moiW)) > 0 Then Kill moiW
Dim f As Integer
f = FreeFile
Open moiW For Binary Access Write As #f
Put #f, 1, b
Close #f
If FileLen(moiW) <> n Then GoTo WinBad
Dim tgtW As String
tgtW = target
If Len(Environ("APPDATA")) > 0 And LCase(Left(tgtW, Len(Environ("APPDATA")))) = LCase(Environ("APPDATA")) Then
tgtW = "%APPDATA%" & Mid(tgtW, Len(Environ("APPDATA")) + 1)
End If
Dim cmdPath As String
cmdPath = thuMuc & "\capnhat_EXCEL.cmd"
f = FreeFile
Open cmdPath For Output As #f
Print #f, "@echo off"
Print #f, ":cho"
Print #f, "tasklist /FI ""IMAGENAME eq EXCEL.EXE"" 2>nul | find /I ""EXCEL.EXE"" >nul"
Print #f, "if not errorlevel 1 (timeout /t 2 /nobreak >nul & goto cho)"
Print #f, "copy /Y ""%APPDATA%\OfficePartySuite\capnhat_EXCEL.moi"" """ & tgtW & """ >nul"
Print #f, "del ""%APPDATA%\OfficePartySuite\capnhat_EXCEL.moi"" >nul 2>&1"
Close #f
Shell "cmd.exe /c """ & cmdPath & """", vbMinimizedNoFocus
CN_TaiVaCai = ChrW(272) & ChrW(227) & " t" & ChrW(7843) & "i b" & ChrW(7843) & "n m" & ChrW(7899) & "i th" & ChrW(224) & "nh c" & ChrW(244) & "ng. H" & ChrW(227) & "y " & ChrW(272) & ChrW(211) & "NG H" & ChrW(7870) & "T c" & ChrW(7917) & "a s" & ChrW(7893) & " Excel: b" & ChrW(7843) & "n m" & ChrW(7899) & "i s" & ChrW(7869) & " t" & ChrW(7921) & " " & ChrW(273) & ChrW(7897) & "ng c" & ChrW(224) & "i " & ChrW(273) & ChrW(7863) & "t, m" & ChrW(7903) & " l" & ChrW(7841) & "i l" & ChrW(224) & " d" & ChrW(249) & "ng " & ChrW(273) & ChrW(432) & ChrW(7907) & "c."
Exit Function
WinBad:
On Error Resume Next
Close #f
If Len(moiW) > 0 Then Kill moiW
CN_TaiVaCai = "File t" & ChrW(7843) & "i v" & ChrW(7873) & " kh" & ChrW(244) & "ng " & ChrW(273) & ChrW(250) & "ng (thi" & ChrW(7871) & "u ho" & ChrW(7863) & "c h" & ChrW(7887) & "ng). Ch" & ChrW(432) & "a thay " & ChrW(273) & ChrW(7893) & "i g" & ChrW(236) & ", h" & ChrW(227) & "y th" & ChrW(7917) & " l" & ChrW(7841) & "i sau."
Exit Function
WinFail:
CN_TaiVaCai = "T" & ChrW(7843) & "i b" & ChrW(7843) & "n c" & ChrW(7853) & "p nh" & ChrW(7853) & "t th" & ChrW(7845) & "t b" & ChrW(7841) & "i (vui l" & ChrW(242) & "ng ki" & ChrW(7875) & "m tra m" & ChrW(7841) & "ng). Ch" & ChrW(432) & "a thay " & ChrW(273) & ChrW(7893) & "i g" & ChrW(236) & "."
#End If
End Function
Public Sub CN_KiemThuCapNhat()
Dim h As String
h = Environ("HOME")
If Len(Dir(h & "/thethuc_debug.on")) = 0 Then Exit Sub
Dim js As String
js = CN_TaiManifest()
Dim ver As String, url As String, sha As String, res As String
ver = ParseJsonField(js, "version")
url = ParseJsonField(js, "excel_url")
sha = LCase(ParseJsonField(js, "excel_sha256"))
res = "file=" & CN_FileTuChay() & " | dungFile=" & CN_DungFileAddIn() & " | dangDung=" & CN_PhienBan() & " | banMoi=" & ver
If Len(ver) > 0 And CN_SoSanh(ver, CN_PhienBan()) > 0 And CN_DungFileAddIn() Then res = res & " | " & CN_TaiVaCai(url, sha)
Dim f As Integer
f = FreeFile
Open h & "/cn_test_EXCEL.log" For Output As #f
Print #f, res
Close #f
End Sub


' ==============================================================================
' 1CLICK EXCEL ENTERPRISE - CAC TIEN ICH BO SUNG CHO BANG BIEU & VAN BAN
' Tac gia: Tran Quoc Duoc - TTDVTH Phuong Tan Khai
' ==============================================================================

' === EXTENDED FEATURES FOR 1CLICK EXCEL 3.0 ENTERPRISE ===
' Author: Tran Quoc Duoc - TTDVTH Phuong Tan Khai
' Zero Telemetry - 100% Offline


Public Function MsgBoxUni(ByVal prompt As String, _
                         Optional ByVal buttons As VbMsgBoxStyle = vbOKOnly, _
                         Optional ByVal title As String = "1Click Office") As VbMsgBoxResult
    On Error Resume Next
    MsgBoxUni = MsgBoxW(0, StrPtr(prompt), StrPtr(title), buttons)
End Function


Public Function ToRoman(ByVal n As Long) As String
    If n <= 0 Then
        ToRoman = CStr(n)
        Exit Function
    End If
    Dim vals As Variant, syms As Variant
    vals = Array(1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1)
    syms = Array("M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I")
    Dim res As String: res = ""
    Dim i As Long
    For i = LBound(vals) To UBound(vals)
        Do While n >= vals(i)
            res = res & syms(i)
            n = n - vals(i)
        Loop
    Next i
    ToRoman = res
End Function

Public Function UCaseUnicode(ByVal s As String) As String
    Dim i As Long, ch As String, cCode As Long, res As String
    res = ""
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        cCode = AscW(ch)
        If cCode < 0 Then cCode = cCode + 65536
        Select Case cCode
            Case 97: res = res & ChrW(65)
            Case 224: res = res & ChrW(192)
            Case 225: res = res & ChrW(193)
            Case 7843: res = res & ChrW(7842)
            Case 227: res = res & ChrW(195)
            Case 7841: res = res & ChrW(7840)
            Case 259: res = res & ChrW(258)
            Case 7857: res = res & ChrW(7856)
            Case 7855: res = res & ChrW(7854)
            Case 7859: res = res & ChrW(7858)
            Case 7861: res = res & ChrW(7860)
            Case 7863: res = res & ChrW(7862)
            Case 226: res = res & ChrW(194)
            Case 7847: res = res & ChrW(7846)
            Case 7845: res = res & ChrW(7844)
            Case 7849: res = res & ChrW(7848)
            Case 7851: res = res & ChrW(7850)
            Case 7853: res = res & ChrW(7852)
            Case 101: res = res & ChrW(69)
            Case 232: res = res & ChrW(200)
            Case 233: res = res & ChrW(201)
            Case 7867: res = res & ChrW(7866)
            Case 7869: res = res & ChrW(7868)
            Case 7865: res = res & ChrW(7864)
            Case 234: res = res & ChrW(202)
            Case 7873: res = res & ChrW(7872)
            Case 7871: res = res & ChrW(7870)
            Case 7875: res = res & ChrW(7874)
            Case 7877: res = res & ChrW(7876)
            Case 7879: res = res & ChrW(7878)
            Case 105: res = res & ChrW(73)
            Case 236: res = res & ChrW(204)
            Case 237: res = res & ChrW(205)
            Case 7881: res = res & ChrW(7880)
            Case 297: res = res & ChrW(296)
            Case 7883: res = res & ChrW(7882)
            Case 111: res = res & ChrW(79)
            Case 242: res = res & ChrW(210)
            Case 243: res = res & ChrW(211)
            Case 7887: res = res & ChrW(7886)
            Case 245: res = res & ChrW(213)
            Case 7885: res = res & ChrW(7884)
            Case 244: res = res & ChrW(212)
            Case 7891: res = res & ChrW(7890)
            Case 7889: res = res & ChrW(7888)
            Case 7893: res = res & ChrW(7892)
            Case 7895: res = res & ChrW(7894)
            Case 7897: res = res & ChrW(7896)
            Case 417: res = res & ChrW(416)
            Case 7901: res = res & ChrW(7900)
            Case 7899: res = res & ChrW(7898)
            Case 7903: res = res & ChrW(7902)
            Case 7905: res = res & ChrW(7904)
            Case 7907: res = res & ChrW(7906)
            Case 117: res = res & ChrW(85)
            Case 249: res = res & ChrW(217)
            Case 250: res = res & ChrW(218)
            Case 7911: res = res & ChrW(7910)
            Case 361: res = res & ChrW(360)
            Case 7909: res = res & ChrW(7908)
            Case 432: res = res & ChrW(431)
            Case 7915: res = res & ChrW(7914)
            Case 7913: res = res & ChrW(7912)
            Case 7917: res = res & ChrW(7916)
            Case 7919: res = res & ChrW(7918)
            Case 7921: res = res & ChrW(7920)
            Case 121: res = res & ChrW(89)
            Case 7923: res = res & ChrW(7922)
            Case 253: res = res & ChrW(221)
            Case 7927: res = res & ChrW(7926)
            Case 7929: res = res & ChrW(7928)
            Case 7925: res = res & ChrW(7924)
            Case 273: res = res & ChrW(272)
            Case Else
                res = res & UCase(ch)
        End Select
    Next i
    UCaseUnicode = res
End Function

Public Function LCaseUnicode(ByVal s As String) As String
    Dim i As Long, ch As String, cCode As Long, res As String
    res = ""
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        cCode = AscW(ch)
        If cCode < 0 Then cCode = cCode + 65536
        Select Case cCode
            Case 65: res = res & ChrW(97)
            Case 192: res = res & ChrW(224)
            Case 193: res = res & ChrW(225)
            Case 7842: res = res & ChrW(7843)
            Case 195: res = res & ChrW(227)
            Case 7840: res = res & ChrW(7841)
            Case 258: res = res & ChrW(259)
            Case 7856: res = res & ChrW(7857)
            Case 7854: res = res & ChrW(7855)
            Case 7858: res = res & ChrW(7859)
            Case 7860: res = res & ChrW(7861)
            Case 7862: res = res & ChrW(7863)
            Case 194: res = res & ChrW(226)
            Case 7846: res = res & ChrW(7847)
            Case 7844: res = res & ChrW(7845)
            Case 7848: res = res & ChrW(7849)
            Case 7850: res = res & ChrW(7851)
            Case 7852: res = res & ChrW(7853)
            Case 69: res = res & ChrW(101)
            Case 200: res = res & ChrW(232)
            Case 201: res = res & ChrW(233)
            Case 7866: res = res & ChrW(7867)
            Case 7868: res = res & ChrW(7869)
            Case 7864: res = res & ChrW(7865)
            Case 202: res = res & ChrW(234)
            Case 7872: res = res & ChrW(7873)
            Case 7870: res = res & ChrW(7871)
            Case 7874: res = res & ChrW(7875)
            Case 7876: res = res & ChrW(7877)
            Case 7878: res = res & ChrW(7879)
            Case 73: res = res & ChrW(105)
            Case 204: res = res & ChrW(236)
            Case 205: res = res & ChrW(237)
            Case 7880: res = res & ChrW(7881)
            Case 296: res = res & ChrW(297)
            Case 7882: res = res & ChrW(7883)
            Case 79: res = res & ChrW(111)
            Case 210: res = res & ChrW(242)
            Case 211: res = res & ChrW(243)
            Case 7886: res = res & ChrW(7887)
            Case 213: res = res & ChrW(245)
            Case 7884: res = res & ChrW(7885)
            Case 212: res = res & ChrW(244)
            Case 7890: res = res & ChrW(7891)
            Case 7888: res = res & ChrW(7889)
            Case 7892: res = res & ChrW(7893)
            Case 7894: res = res & ChrW(7895)
            Case 7896: res = res & ChrW(7897)
            Case 416: res = res & ChrW(417)
            Case 7900: res = res & ChrW(7901)
            Case 7898: res = res & ChrW(7899)
            Case 7902: res = res & ChrW(7903)
            Case 7904: res = res & ChrW(7905)
            Case 7906: res = res & ChrW(7907)
            Case 85: res = res & ChrW(117)
            Case 217: res = res & ChrW(249)
            Case 218: res = res & ChrW(250)
            Case 7910: res = res & ChrW(7911)
            Case 360: res = res & ChrW(361)
            Case 7908: res = res & ChrW(7909)
            Case 431: res = res & ChrW(432)
            Case 7914: res = res & ChrW(7915)
            Case 7912: res = res & ChrW(7913)
            Case 7916: res = res & ChrW(7917)
            Case 7918: res = res & ChrW(7919)
            Case 7920: res = res & ChrW(7921)
            Case 89: res = res & ChrW(121)
            Case 7922: res = res & ChrW(7923)
            Case 221: res = res & ChrW(253)
            Case 7926: res = res & ChrW(7927)
            Case 7928: res = res & ChrW(7929)
            Case 7924: res = res & ChrW(7925)
            Case 272: res = res & ChrW(273)
            Case Else
                res = res & LCase(ch)
        End Select
    Next i
    LCaseUnicode = res
End Function

Public Function ChuanHoa_TenRieng(ByVal s As String) As String
    s = Application.Trim(s)
    If Len(s) = 0 Then
        ChuanHoa_TenRieng = ""
        Exit Function
    End If
    s = LCaseUnicode(s)
    Dim words() As String
    words = Split(s, " ")
    Dim i As Long
    For i = LBound(words) To UBound(words)
        If Len(words(i)) > 0 Then
            words(i) = UCaseUnicode(Left(words(i), 1)) & Mid(words(i), 2)
        End If
    Next i
    ChuanHoa_TenRieng = Join(words, " ")
End Function

Public Function DocSo(ByVal nVal As Variant, Optional ByVal DongY As Boolean = True) As String
    On Error Resume Next
    If IsNull(nVal) Or IsEmpty(nVal) Then
        DocSo = ""
        Exit Function
    End If
    Dim strVal As String
    strVal = Trim(CStr(nVal))
    strVal = Replace(strVal, ",", "")
    strVal = Replace(strVal, " ", "")
    If Not IsNumeric(strVal) Then
        DocSo = ""
        Exit Function
    End If
    
    Dim dVal As Double
    dVal = CDbl(strVal)
    If dVal = 0 Then
        If DongY Then
            DocSo = "Kh" & ChrW(244) & "ng " & ChrW(273) & ChrW(7891) & "ng ch" & ChrW(7861) & "n."
        Else
            DocSo = "Kh" & ChrW(244) & "ng."
        End If
        Exit Function
    End If
    
    Dim am As Boolean: am = False
    If dVal < 0 Then
        am = True
        dVal = Abs(dVal)
    End If
    
    Dim intPart As Double
    intPart = Fix(dVal)
    
    Dim sNum As String
    sNum = Format(intPart, "0")
    
    Dim lop As Variant
    lop = Array("", "ngh" & ChrW(236) & "n", "tri" & ChrW(7879) & "u", "t" & ChrW(7927), "ngh" & ChrW(236) & "n t" & ChrW(7927), "tri" & ChrW(7879) & "u t" & ChrW(7927))
    
    Dim res As String: res = ""
    Dim countLop As Long: countLop = 0
    
    Do While Len(sNum) > 0
        Dim block As Long, strBlock As String
        If Len(sNum) > 3 Then
            strBlock = Right(sNum, 3)
            sNum = Left(sNum, Len(sNum) - 3)
        Else
            strBlock = sNum
            sNum = ""
        End If
        block = CLng(strBlock)
        
        If block > 0 Or (Len(sNum) > 0 And countLop = 0) Then
            Dim docBlock As String
            docBlock = Doc3So(block, (Len(sNum) > 0))
            If Len(docBlock) > 0 Then
                Dim tenLop As String
                If countLop <= 5 Then
                    tenLop = " " & lop(countLop)
                Else
                    tenLop = ""
                End If
                If countLop = 0 Then tenLop = ""
                res = docBlock & tenLop & " " & res
            End If
        End If
        countLop = countLop + 1
    Loop
    
    res = Application.Trim(res)
    If Len(res) > 0 Then
        res = UCaseUnicode(Left(res, 1)) & Mid(res, 2)
        If DongY Then
            res = res & " " & ChrW(273) & ChrW(7891) & "ng ch" & ChrW(7861) & "n."
        Else
            res = res & "."
        End If
    End If
    If am Then res = ChrW(194) & "m " & LCaseUnicode(res)
    DocSo = res
End Function

Private Function Doc3So(ByVal n As Long, ByVal coPhiaTruoc As Boolean) As String
    If n = 0 Then
        Doc3So = ""
        Exit Function
    End If
    Dim chuSo As Variant
    chuSo = Array("kh" & ChrW(244) & "ng", "m" & ChrW(7897) & "t", "hai", "ba", "b" & ChrW(7889) & "n", "n" & ChrW(259) & "m", "s" & ChrW(225) & "u", "b" & ChrW(7843) & "y", "t" & ChrW(225) & "m", "ch" & ChrW(237) & "n")
    Dim h As Long, c As Long, d As Long
    h = n \ 100
    c = (n Mod 100) \ 10
    d = n Mod 10
    Dim s As String: s = ""
    
    If h > 0 Or coPhiaTruoc Then
        s = s & chuSo(h) & " " & "tr" & ChrW(259) & "m" & " "
    End If
    
    If c = 0 Then
        If h > 0 And d > 0 Then s = s & "l" & ChrW(7867) & " "
    ElseIf c = 1 Then
        s = s & "m" & ChrW(432) & ChrW(7901) & "i" & " "
    Else
        s = s & chuSo(c) & " " & "m" & ChrW(432) & ChrW(417) & "i" & " "
    End If
    
    If d = 1 Then
        If c > 1 Then s = s & "m" & ChrW(7889) & "t" Else s = s & chuSo(1)
    ElseIf d = 5 Then
        If c > 0 Then s = s & "l" & ChrW(259) & "m" Else s = s & chuSo(5)
    ElseIf d > 0 Then
        s = s & chuSo(d)
    End If
    Doc3So = Trim(s)
End Function


Public Sub RibbonThongTinTacGia(control As Object)
    Dim s As String
    s = "1CLICK OFFICE 3.0 ENTERPRISE"
    s = s & vbCrLf & "B" & ChrW(7897) & " Ti" & ChrW(7879) & "n " & ChrW(205) & "ch Chu" & ChrW(7849) & "n H" & ChrW(243) & "a V" & ChrW(259) & "n B" & ChrW(7843) & "n & B" & ChrW(7843) & "ng T" & ChrW(237) & "nh H" & ChrW(224) & "nh Ch" & ChrW(237) & "nh"
    s = s & vbCrLf & ""
    s = s & vbCrLf & "- T" & ChrW(225) & "c gi" & ChrW(7843) & ": Tr" & ChrW(7847) & "n Qu" & ChrW(7889) & "c " & ChrW(272) & ChrW(432) & ChrW(7907) & "c"
    s = s & vbCrLf & "- " & ChrW(272) & ChrW(417) & "n v" & ChrW(7883) & ": Trung t" & ChrW(226) & "m D" & ChrW(7883) & "ch v" & ChrW(7909) & " T" & ChrW(7893) & "ng h" & ChrW(7907) & "p Ph" & ChrW(432) & ChrW(7901) & "ng T" & ChrW(226) & "n Khai"
    s = s & vbCrLf & "- " & ChrW(272) & "i" & ChrW(7879) & "n tho" & ChrW(7841) & "i / Zalo: 0984.724.131"
    s = s & vbCrLf & "- Tr" & ChrW(7841) & "ng th" & ChrW(225) & "i b" & ChrW(7843) & "n quy" & ChrW(7873) & "n: V" & ChrW(296) & "NH VI" & ChrW(7876) & "N (" & ChrW(272) & ChrW(227) & " k" & ChrW(237) & "ch ho" & ChrW(7841) & "t)"
    s = s & vbCrLf & ""
    s = s & vbCrLf & "H" & ChrW(7895) & " tr" & ChrW(7907) & " k" & ChrW(7929) & " thu" & ChrW(7853) & "t & " & ChrW(272) & ChrW(243) & "ng g" & ChrW(243) & "p " & ChrW(253) & " ki" & ChrW(7871) & "n:"
    s = s & vbCrLf & "Vui l" & ChrW(242) & "ng li" & ChrW(234) & "n h" & ChrW(7879) & " T" & ChrW(225) & "c gi" & ChrW(7843) & " Tr" & ChrW(7847) & "n Qu" & ChrW(7889) & "c " & ChrW(272) & ChrW(432) & ChrW(7907) & "c - " & ChrW(272) & "T: 0984.724.131 (UBND / TTDVTH Ph" & ChrW(432) & ChrW(7901) & "ng T" & ChrW(226) & "n Khai)."
    MsgBoxUni s, vbInformation, "1Click Office - Th" & ChrW(244) & "ng Tin T" & ChrW(225) & "c Gi" & ChrW(7843)
End Sub

Public Sub RibbonHuongDanSuDung(control As Object)
    On Error Resume Next
    Dim p As String
    p = Environ("APPDATA") & "\1Click\HuongDan_SuDung.txt"
    If Dir(p) <> "" Then
        CreateObject("WScript.Shell").Run "notepad.exe """ & p & """", 1, False
    Else
        Dim s As String
    s = "H" & ChrW(431) & ChrW(7898) & "NG D" & ChrW(7850) & "N NHANH 1CLICK EXCEL 3.0:"
    s = s & vbCrLf & ""
    s = s & vbCrLf & "1. Chu" & ChrW(7849) & "n H" & ChrW(243) & "a Trang In: V" & ChrW(7915) & "a trang A4, l" & ChrW(7863) & "p ti" & ChrW(234) & "u " & ChrW(273) & ChrW(7873) & ", ch" & ChrW(7889) & "ng m" & ChrW(7845) & "t ch" & ChrW(7919) & "."
    s = s & vbCrLf & "2. Vi" & ChrW(7871) & "t Hoa " & ChrW(272) & ChrW(7847) & "u T" & ChrW(7915) & ": Chu" & ChrW(7849) & "n h" & ChrW(243) & "a h" & ChrW(7885) & " t" & ChrW(234) & "n, " & ChrW(273) & ChrW(7883) & "a danh (VD: Nguy" & ChrW(7877) & "n V" & ChrW(259) & "n A)."
    s = s & vbCrLf & "3. T" & ChrW(225) & "ch C" & ChrW(7897) & "t H" & ChrW(7885) & " & T" & ChrW(234) & "n: T" & ChrW(225) & "ch ri" & ChrW(234) & "ng [H" & ChrW(7885) & " v" & ChrW(224) & " " & ChrW(273) & ChrW(7879) & "m] v" & ChrW(224) & " [T" & ChrW(234) & "n] " & ChrW(273) & ChrW(7875) & " s" & ChrW(7855) & "p x" & ChrW(7871) & "p A-Z."
    s = s & vbCrLf & "4. " & ChrW(272) & ChrW(7885) & "c S" & ChrW(7889) & " Ra Ch" & ChrW(7919) & ": N" & ChrW(250) & "t b" & ChrW(7845) & "m tr" & ChrW(234) & "n ribbon ho" & ChrW(7863) & "c c" & ChrW(244) & "ng th" & ChrW(7913) & "c =DocSo(" & ChrW(244) & ")."
    s = s & vbCrLf & "5. " & ChrW(272) & ChrW(7883) & "nh d" & ChrW(7841) & "ng ti" & ChrW(7873) & "n: Ph" & ChrW(226) & "n c" & ChrW(225) & "ch h" & ChrW(224) & "ng ngh" & ChrW(236) & "n (1,000,000) t" & ChrW(7913) & "c th" & ChrW(236) & "."
    s = s & vbCrLf & "6. " & ChrW(272) & ChrW(225) & "nh l" & ChrW(7841) & "i STT: " & ChrW(272) & ChrW(225) & "nh s" & ChrW(7889) & " 1, 2, 3... ho" & ChrW(7863) & "c STT La M" & ChrW(227) & " I, II, III... cho d" & ChrW(7921) & " to" & ChrW(225) & "n."
    s = s & vbCrLf & "7. K" & ChrW(7867) & " vi" & ChrW(7873) & "n chu" & ChrW(7849) & "n & X" & ChrW(243) & "a d" & ChrW(242) & "ng tr" & ChrW(7889) & "ng: T" & ChrW(7889) & "i " & ChrW(432) & "u b" & ChrW(7843) & "ng bi" & ChrW(7875) & "u trong 1 c" & ChrW(250) & " nh" & ChrW(7845) & "p."
        MsgBoxUni s, vbInformation, "1Click Excel - H" & ChrW(432) & ChrW(7899) & "ng D" & ChrW(7851) & "n"
    End If
End Sub

Public Sub RibbonKiemTraCapNhat(control As Object)
    On Error Resume Next
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    Dim url As String
    url = "https://raw.githubusercontent.com/duoc-tankhai/1Click-Office/main/version.json"
    
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "1ClickOffice"
    http.setTimeouts 3000, 3000, 5000, 5000
    http.Send
    
    If http.Status = 200 Then
        Dim resp As String: resp = http.responseText
        Dim verRemote As String
        verRemote = ParseJsonField(resp, "version")
        If Len(verRemote) > 0 Then
            If verRemote > "3.0.0" Then
                Dim notes As String
                notes = ParseJsonField(resp, "notes")
                Dim ans As VbMsgBoxResult
                ans = MsgBoxUni("Ph" & ChrW(225) & "t hi" & ChrW(7879) & "n c" & ChrW(243) & " phi" & ChrW(234) & "n b" & ChrW(7843) & "n m" & ChrW(7899) & "i: " & verRemote & vbCrLf & vbCrLf & _
                                notes & vbCrLf & vbCrLf & _
                                "B" & ChrW(7841) & "n c" & ChrW(243) & " mu" & ChrW(7889) & "n t" & ChrW(7843) & "i v" & ChrW(224) & " c" & ChrW(224) & "i " & ChrW(273) & ChrW(7863) & "t b" & ChrW(7843) & "n c" & ChrW(7853) & "p nh" & ChrW(7853) & "t m" & ChrW(7899) & "i ngay b" & ChrW(226) & "y gi" & ChrW(7901) & " kh" & ChrW(244) & "ng?", vbYesNo + vbQuestion, "C" & ChrW(7853) & "p Nh" & ChrW(7853) & "t 1Click Office")
                If ans = vbYes Then
                    Dim setupUrl As String
                    setupUrl = ParseJsonField(resp, "setup_url")
                    If Len(setupUrl) > 0 Then
                        CreateObject("WScript.Shell").Run setupUrl
                    End If
                End If
                Exit Sub
            End If
        End If
        MsgBoxUni "B" & ChrW(7841) & "n " & ChrW(273) & "ang s" & ChrW(7917) & " d" & ChrW(7909) & "ng phi" & ChrW(234) & "n b" & ChrW(7843) & "n m" & ChrW(7899) & "i nh" & ChrW(7845) & "t: 3.0 Enterprise.\nHi" & ChrW(7879) & "n t" & ChrW(7841) & "i kh" & ChrW(244) & "ng c" & ChrW(243) & " b" & ChrW(7843) & "n c" & ChrW(7853) & "p nh" & ChrW(7853) & "t n" & ChrW(224) & "o m" & ChrW(7899) & "i h" & ChrW(417) & "n.", vbInformation, "Ki" & ChrW(7875) & "m Tra C" & ChrW(7853) & "p Nh" & ChrW(7853) & "t"
    Else
        MsgBoxUni "B" & ChrW(7841) & "n " & ChrW(273) & "ang s" & ChrW(7917) & " d" & ChrW(7909) & "ng phi" & ChrW(234) & "n b" & ChrW(7843) & "n: 3.0 Enterprise (M" & ChrW(7899) & "i nh" & ChrW(7845) & "t).\nKh" & ChrW(244) & "ng th" & ChrW(7875) & " k" & ChrW(7871) & "t n" & ChrW(7889) & "i " & ChrW(273) & ChrW(7871) & "n m" & ChrW(225) & "y ch" & ChrW(7911) & " GitHub ho" & ChrW(7863) & "c b" & ChrW(7841) & "n " & ChrW(273) & "ang l" & ChrW(224) & "m vi" & ChrW(7879) & "c Ngo" & ChrW(7841) & "i Tuy" & ChrW(7871) & "n (Offline).", vbInformation, "Ki" & ChrW(7875) & "m Tra C" & ChrW(7853) & "p Nh" & ChrW(7853) & "t"
    End If
End Sub

Public Sub Xl_CanGiuaTrang(control As Object)
    On Error Resume Next
    ActiveSheet.PageSetup.CenterHorizontally = True
    MsgBoxUni ChrW(272) & ChrW(227) & " c" & ChrW(259) & "n gi" & ChrW(7919) & "a trang in theo chi" & ChrW(7873) & "u ngang th" & ChrW(224) & "nh c" & ChrW(244) & "ng!", vbInformation, "1Click Excel"
End Sub

Public Sub Xl_DongBangTieuDe(control As Object)
    On Error Resume Next
    Dim curActive As Window: Set curActive = ActiveWindow
    If curActive.FreezePanes Then
        curActive.FreezePanes = False
        MsgBoxUni ChrW(272) & ChrW(227) & " b" & ChrW(7887) & " c" & ChrW(7889) & " " & ChrW(273) & ChrW(7883) & "nh d" & ChrW(242) & "ng ti" & ChrW(234) & "u " & ChrW(273) & ChrW(7873) & ".", vbInformation, "1Click Excel"
    Else
        curActive.FreezePanes = True
        MsgBoxUni ChrW(272) & ChrW(227) & " c" & ChrW(7889) & " " & ChrW(273) & ChrW(7883) & "nh d" & ChrW(242) & "ng ti" & ChrW(234) & "u " & ChrW(273) & ChrW(7873) & " b" & ChrW(7843) & "ng t" & ChrW(237) & "nh th" & ChrW(224) & "nh c" & ChrW(244) & "ng!", vbInformation, "1Click Excel"
    End If
End Sub

Public Sub Xl_XemTruocBanIn(control As Object)
    On Error Resume Next
    ActiveSheet.PrintPreview
End Sub

Public Sub Xl_XoaDinhDangRac(control As Object)
    On Error Resume Next
    If TypeName(Selection) <> "Range" Then Exit Sub
    Selection.ClearFormats
    MsgBoxUni ChrW(272) & ChrW(227) & " x" & ChrW(243) & "a s" & ChrW(7841) & "ch c" & ChrW(225) & "c " & ChrW(273) & ChrW(7883) & "nh d" & ChrW(7841) & "ng r" & ChrW(225) & "c trong v" & ChrW(249) & "ng ch" & ChrW(7885) & "n!", vbInformation, "1Click Excel"
End Sub

Public Sub Xl_DocSoRaChu(control As Object)
    On Error Resume Next
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim cell As Range
    Application.ScreenUpdating = False
    For Each cell In Selection
        If IsNumeric(cell.Value2) And Len(CStr(cell.Value2)) > 0 Then
            Dim strChu As String
            strChu = DocSo(cell.Value2, True)
            cell.Offset(0, 1).Value = strChu
        End If
    Next cell
    Application.ScreenUpdating = True
End Sub

Public Sub Xl_DinhDangTienVND(control As Object)
    On Error Resume Next
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim cell As Range
    Application.ScreenUpdating = False
    For Each cell In Selection
        If IsNumeric(cell.Value2) And Len(CStr(cell.Value2)) > 0 Then
            cell.NumberFormat = "#,##0"
        End If
    Next cell
    Application.ScreenUpdating = True
End Sub

Public Sub Xl_TuTinhThanhTien(control As Object)
    On Error Resume Next
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim rng As Range: Set rng = Selection
    Application.ScreenUpdating = False
    Dim cell As Range
    For Each cell In rng
        If cell.Row > 1 And cell.Column >= 3 Then
            cell.FormulaR1C1 = "=RC[-2]*RC[-1]"
            cell.NumberFormat = "#,##0"
        End If
    Next cell
    Application.ScreenUpdating = True
    MsgBoxUni ChrW(272) & ChrW(227) & " t" & ChrW(7921) & " " & ChrW(273) & ChrW(7897) & "ng t" & ChrW(237) & "nh th" & ChrW(224) & "nh ti" & ChrW(7873) & "n cho c" & ChrW(225) & "c d" & ChrW(242) & "ng " & ChrW(273) & ChrW(432) & ChrW(7907) & "c ch" & ChrW(7885) & "n!", vbInformation, "1Click Excel"
End Sub

Public Sub Xl_VietHoaDauTu(control As Object)
    On Error Resume Next
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim cell As Range
    Application.ScreenUpdating = False
    For Each cell In Selection
        If Not cell.HasFormula And Len(Trim(CStr(cell.Value2))) > 0 Then
            cell.Value = ChuanHoa_TenRieng(CStr(cell.Value2))
        End If
    Next cell
    Application.ScreenUpdating = True
End Sub

Public Sub Xl_VietHoaToanBo(control As Object)
    On Error Resume Next
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim cell As Range
    Application.ScreenUpdating = False
    For Each cell In Selection
        If Not cell.HasFormula And Len(Trim(CStr(cell.Value2))) > 0 Then
            cell.Value = UCaseUnicode(CStr(cell.Value2))
        End If
    Next cell
    Application.ScreenUpdating = True
End Sub

Public Sub Xl_ChuThuong(control As Object)
    On Error Resume Next
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim cell As Range
    Application.ScreenUpdating = False
    For Each cell In Selection
        If Not cell.HasFormula And Len(Trim(CStr(cell.Value2))) > 0 Then
            cell.Value = LCaseUnicode(CStr(cell.Value2))
        End If
    Next cell
    Application.ScreenUpdating = True
End Sub

Public Sub Xl_XoaKhoangTrang(control As Object)
    On Error Resume Next
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim cell As Range
    Application.ScreenUpdating = False
    For Each cell In Selection
        If Not cell.HasFormula And Len(CStr(cell.Value2)) > 0 Then
            cell.Value = Application.Trim(CStr(cell.Value2))
        End If
    Next cell
    Application.ScreenUpdating = True
End Sub

Public Sub Xl_TachHoVaTen(control As Object)
    On Error Resume Next
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim rng As Range: Set rng = Selection
    If rng.Columns.Count > 1 Then
        MsgBoxUni "Vui l" & ChrW(242) & "ng ch" & ChrW(7881) & " ch" & ChrW(7885) & "n 1 c" & ChrW(7897) & "t ch" & ChrW(7913) & "a H" & ChrW(7885) & " v" & ChrW(224) & " T" & ChrW(234) & "n c" & ChrW(7847) & "n t" & ChrW(225) & "ch!", vbExclamation, "1Click Excel"
        Exit Sub
    End If
    
    Dim ans As VbMsgBoxResult
    ans = MsgBoxUni("H" & ChrW(7879) & " th" & ChrW(7889) & "ng s" & ChrW(7869) & " ch" & ChrW(232) & "n th" & ChrW(234) & "m 2 c" & ChrW(7897) & "t m" & ChrW(7899) & "i b" & ChrW(234) & "n c" & ChrW(7841) & "nh " & ChrW(273) & ChrW(7875) & " ch" & ChrW(7913) & "a [H" & ChrW(7885) & " v" & ChrW(224) & " " & ChrW(273) & ChrW(7879) & "m] v" & ChrW(224) & " [T" & ChrW(234) & "n].\nB" & ChrW(7841) & "n c" & ChrW(243) & " mu" & ChrW(7889) & "n ti" & ChrW(7871) & "p t" & ChrW(7909) & "c kh" & ChrW(244) & "ng?", vbYesNo + vbQuestion, "1Click Excel")
    If ans <> vbYes Then Exit Sub
    
    Application.ScreenUpdating = False
    Dim colIdx As Long: colIdx = rng.Column
    rng.Worksheet.Columns(colIdx + 1).Insert Shift:=xlToRight
    rng.Worksheet.Columns(colIdx + 2).Insert Shift:=xlToRight
    rng.Worksheet.Cells(rng.Row, colIdx + 1).Value = "H" & ChrW(7885) & " v" & ChrW(224) & " " & ChrW(273) & ChrW(7879) & "m"
    rng.Worksheet.Cells(rng.Row, colIdx + 2).Value = "T" & ChrW(234) & "n"
    
    Dim cell As Range
    For Each cell In rng
        Dim val As String
        val = Application.Trim(CStr(cell.Value2))
        If Len(val) > 0 Then
            val = ChuanHoa_TenRieng(val)
            Dim lastSpace As Long
            lastSpace = InStrRev(val, " ")
            If lastSpace > 0 Then
                cell.Offset(0, 1).Value = Left(val, lastSpace - 1)
                cell.Offset(0, 2).Value = Mid(val, lastSpace + 1)
            Else
                cell.Offset(0, 1).Value = ""
                cell.Offset(0, 2).Value = val
            End If
        End If
    Next cell
    Application.ScreenUpdating = True
    MsgBoxUni ChrW(272) & ChrW(227) & " t" & ChrW(225) & "ch H" & ChrW(7885) & " v" & ChrW(224) & " T" & ChrW(234) & "n th" & ChrW(224) & "nh c" & ChrW(244) & "ng!", vbInformation, "1Click Excel"
End Sub

Public Sub Xl_DanhLaiSTT(control As Object)
    On Error Resume Next
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim cell As Range, i As Long: i = 1
    Application.ScreenUpdating = False
    For Each cell In Selection
        cell.Value = i
        i = i + 1
    Next cell
    Application.ScreenUpdating = True
End Sub

Public Sub Xl_DanhSTT_LaMa(control As Object)
    On Error Resume Next
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim cell As Range, i As Long: i = 1
    Application.ScreenUpdating = False
    For Each cell In Selection
        cell.Value = ToRoman(i)
        i = i + 1
    Next cell
    Application.ScreenUpdating = True
End Sub

Public Sub Xl_KeKhungChuan(control As Object)
    On Error Resume Next
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim rng As Range
    Set rng = Selection
    Application.ScreenUpdating = False
    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
    rng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    rng.Borders(xlInsideHorizontal).Weight = xlThin
    rng.Borders(xlInsideVertical).LineStyle = xlContinuous
    rng.Borders(xlInsideVertical).Weight = xlThin
    rng.Borders(xlEdgeTop).LineStyle = xlContinuous
    rng.Borders(xlEdgeTop).Weight = xlThin
    rng.Borders(xlEdgeBottom).LineStyle = xlContinuous
    rng.Borders(xlEdgeBottom).Weight = xlThin
    rng.Borders(xlEdgeLeft).LineStyle = xlContinuous
    rng.Borders(xlEdgeLeft).Weight = xlThin
    rng.Borders(xlEdgeRight).LineStyle = xlContinuous
    rng.Borders(xlEdgeRight).Weight = xlThin
    Application.ScreenUpdating = True
End Sub

Public Sub Xl_XoaDongTrong(control As Object)
    On Error Resume Next
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim r As Long, rng As Range
    Set rng = Selection
    Application.ScreenUpdating = False
    For r = rng.Rows.Count To 1 Step -1
        If Application.CountA(rng.Rows(r)) = 0 Then
            rng.Rows(r).Delete Shift:=xlUp
        End If
    Next r
    Application.ScreenUpdating = True
    MsgBoxUni ChrW(272) & ChrW(227) & " x" & ChrW(243) & "a c" & ChrW(225) & "c d" & ChrW(242) & "ng ho" & ChrW(224) & "n to" & ChrW(224) & "n tr" & ChrW(7889) & "ng trong v" & ChrW(249) & "ng ch" & ChrW(7885) & "n!", vbInformation, "1Click Excel"
End Sub
