Attribute VB_Name = "ExcelLabLevel6Macros"
'=====================================================================
' Excel Lab - Level 6: Advanced Excel
' Ready-to-import VBA module with the macros taught in the macro
' lessons: "Recording Macros", "VBA Fundamentals" and
' "Automating Reports with VBA".
'
' HOW TO USE THIS FILE
'   1. Open the macro practice workbook and re-save it as .xlsm
'      (File > Save As > Excel Macro-Enabled Workbook) - a plain
'      .xlsx file cannot store macros.
'   2. Press Alt+F11 to open the VBA editor.
'   3. Either: File > Import File... and choose this .bas file,
'      or:     Insert > Module, then copy-paste this code in.
'   4. Put the cursor inside a Sub and press F5 to run it, or F8 to
'      step through it line by line. From Excel, use Alt+F8.
'
' Contents:
'   FormatReport       recorded-style macro (uses .Select)
'   RelativeInitials   recorded with Use Relative References
'   FillReportTitle    first hand-written Sub (Range + Format)
'   CellTest           Range vs Cells
'   OrderSummary       Dim variables + MsgBox
'   WhoIsIt            InputBox with Cancel handling
'   FormatReportClean  the With-block rewrite (no .Select)
'   CountPens          For loop + If counter
'   FilterLaptops      bonus: AutoFilter the Sales sheet
'   MonthlyReport      full loop/summary/save macro with On Error
'=====================================================================
Option Explicit

Sub FormatReport()
' Formats the daily sales report header - exactly as the macro
' recorder writes it (record on the Exercise sheet of the practice
' workbook, with A1:E1 holding the report header).

  Range("A1:E1").Select
  Selection.Font.Bold = True
  With Selection.Interior
    .Color = RGB(255, 217, 102)
  End With
  Selection.HorizontalAlignment = xlCenter
  Range("A1").Select

End Sub

Sub RelativeInitials()
' Recorded with Use Relative References switched ON.
' From wherever the cursor is: go two rows down and type initials
' (replace "AB" with your own initials). Run it from different
' starting cells - the initials always land two rows below the
' ACTIVE cell, never at a fixed address.

  ActiveCell.Offset(2, 0).Range("A1").Select
  ActiveCell.FormulaR1C1 = "AB"

End Sub

Sub FillReportTitle()
' Puts the title in A1 and today's date in A2

  Range("A1").Value = "Sunrise Traders - Daily Sales"
  Range("A2").Value = "Report date: " & Format(Date, "dd-mmm-yyyy")

End Sub

Sub CellTest()
' Writes the same value to C2 two ways: Range (address) and
' Cells (row, column). The second form is the one loops use.

  Range("C2").Value = "Staplers"     ' the cell C2
  Cells(2, 3).Value = "Staplers"     ' row 2, column 3 = C2

End Sub

Sub OrderSummary()
' Shows a one-line summary of a sample order.
' Step through it with F8 and hover over the variables to watch
' product, qty, price and total fill in.

  Dim product As String
  Dim qty As Long
  Dim price As Double
  Dim total As Double

  product = "Staplers"
  qty = 12
  price = 210
  total = qty * price

  MsgBox "Sold " & qty & " " & product & " for " & total, vbInformation

End Sub

Sub WhoIsIt()
' Asks for a name and greets the user

  Dim who As String

  who = InputBox("Which salesperson is this report for?")
  If who = "" Then Exit Sub          ' user pressed Cancel - stop politely

  MsgBox "Preparing the report for " & who, vbInformation

End Sub

Sub FormatReportClean()
' Hand-written replacement for the recorded FormatReport: the same
' result with a With block and no .Select anywhere, so it works no
' matter which cell the cursor is in.

  With Range("A1:E1")
    .Font.Bold = True
    .Interior.Color = RGB(255, 217, 102)
    .HorizontalAlignment = xlCenter
  End With

End Sub

Sub CountPens()
' Loops the Sales sheet rows and counts every Pens sale.
' Reports ONCE at the end, never inside the loop.

  Dim ws As Worksheet
  Dim lastRow As Long, r As Long
  Dim penCount As Long

  Set ws = ThisWorkbook.Worksheets("Sales")
  lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

  For r = 2 To lastRow
    If ws.Cells(r, "C").Value = "Pens" Then
      penCount = penCount + 1
    End If
  Next r

  MsgBox "Pens sales found: " & penCount, vbInformation

End Sub

Sub FilterLaptops()
' Bonus: applies an AutoFilter to the Sales sheet and shows only
' the laptop rows. Reset with Data > Clear when done.

  With ThisWorkbook.Worksheets("Sales")
    .Range("A1:D1").AutoFilter Field:=3, Criteria1:="Laptops"
  End With

End Sub

Sub MonthlyReport()
' Scans the Sales sheet, highlights laptop sales,
' writes a summary sheet and saves a dated copy.

  Dim ws As Worksheet, rep As Worksheet
  Dim lastRow As Long, r As Long
  Dim total As Double, laptopTotal As Double

  On Error GoTo ErrHandler

  Set ws = ThisWorkbook.Worksheets("Sales")
  lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

  total = 0
  laptopTotal = 0
  For r = 2 To lastRow
    total = total + ws.Cells(r, "D").Value
    If ws.Cells(r, "C").Value = "Laptops" Then
      ws.Rows(r).Interior.Color = RGB(255, 235, 205)
      laptopTotal = laptopTotal + ws.Cells(r, "D").Value
    End If
  Next r

  Set rep = ThisWorkbook.Worksheets.Add(After:=Sheets(Sheets.Count))
  On Error Resume Next           ' name may already exist this month
  rep.Name = "Report " & Format(Date, "yyyy-mm")
  On Error GoTo ErrHandler

  rep.Range("A1").Value = "Sunrise Traders - Monthly Report"
  rep.Range("A2").Value = "Total sales"
  rep.Range("B2").Value = total
  rep.Range("A3").Value = "Laptop sales"
  rep.Range("B3").Value = laptopTotal
  rep.Range("B2:B3").NumberFormat = "#,##0"
  rep.Range("A1").Font.Bold = True

  ThisWorkbook.SaveAs Filename:= _
    ThisWorkbook.Path & Application.PathSeparator & _
    "Sunrise Traders " & Format(Date, "yyyy-mm-dd") & ".xlsm", _
    FileFormat:=xlOpenXMLWorkbookMacroEnabled

  MsgBox "Report ready. Total " & Format(total, "#,##0"), vbInformation
  Exit Sub

ErrHandler:
  MsgBox "The macro stopped: " & Err.Description, vbExclamation
End Sub
