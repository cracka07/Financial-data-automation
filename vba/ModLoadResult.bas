Attribute VB_Name = "ModLoadResult"
Option Explicit

Sub LoadResult(Ticker)


     Dim row As Long
     Dim lastRow As Long
     Dim wsRatio As Worksheet
     Dim ct As Long
     Dim sym As String
     
     Set wsRatio = Worksheets("FACT_FINANCIALS_MAX")
     
    With frmFinancial.lstResult
         .Clear
         .ColumnHeads = False
         .ColumnCount = 6
        .ColumnWidths = "100 pt;70 pt;80 pt; 80 pt;80 pt;100 pt"
         
         lastRow = wsRatio.Cells(Rows.Count, "A").End(xlUp).row
         ct = 0
         For row = 2 To lastRow
                sym = wsRatio.Cells(row, "A").Value
                If sym Like Ticker Then
                .AddItem
                .List(ct, 0) = wsRatio.Cells(row, 2).Value
                .List(ct, 1) = wsRatio.Cells(row, 1).Value
                .List(ct, 2) = wsRatio.Cells(row, 16).Value
                .List(ct, 3) = wsRatio.Cells(row, 18).Value
                .List(ct, 4) = wsRatio.Cells(row, 17).Value
                If wsRatio.Cells(row, 10) = "Strong" Then
                .List(ct, 5) = wsRatio.Cells(row, 10).Value
                frmFinancial.TextBox1.BackColor = vbGreen
                ElseIf wsRatio.Cells(row, 10) = "Moderate" Then
                .List(ct, 5) = wsRatio.Cells(row, 10).Value
                frmFinancial.TextBox1.BackColor = vbYellow
                Else
                .List(ct, 5) = wsRatio.Cells(row, 10).Value
                frmFinancial.TextBox1.BackColor = vbRed
                End If
          ct = ct + 1
          End If
         Next row
         End With
End Sub

Sub LoadHistoricalData(Ticker)

     Dim row As Long
     Dim lastRow As Long
     Dim wsRatio As Worksheet
     Dim ct As Long
     Dim dYear As Long
     Dim Y As Long
     Dim sym As String
     
     
     Y = 2004
     
     Set wsRatio = Worksheets("FACT_FINANCIALS_CORRECT")
     
    With frmFinancial.lstHistorical
         .Clear
         .ColumnHeads = False
         .ColumnCount = 8
        .ColumnWidths = "60 pt;60 pt;70 pt; 80 pt;50 pt;70 pt; 60 pt;60 pt"
        
         lastRow = wsRatio.Cells(Rows.Count, "A").End(xlUp).row
         ct = 0
         For row = 2 To lastRow
         
                dYear = wsRatio.Cells(row, 12).Value
             
                sym = wsRatio.Cells(row, "J").Value
                If sym Like Ticker Then
                .AddItem
                .List(ct, 0) = wsRatio.Cells(row, 12).Value
                .List(ct, 1) = wsRatio.Cells(row, 16).Value
                .List(ct, 2) = wsRatio.Cells(row, 18).Value
                .List(ct, 3) = wsRatio.Cells(row, 17).Value
                .List(ct, 4) = wsRatio.Cells(row, 5).Value
                .List(ct, 5) = wsRatio.Cells(row, 3).Value
                .List(ct, 6) = wsRatio.Cells(row, 2).Value
                .List(ct, 7) = wsRatio.Cells(row, 4).Value
          ct = ct + 1
          End If
         Next row
         End With
    
End Sub
'========================================================
' Project: Financial Dashboard
' Author: Mariano Ferrer
' Role: Excel VBA Developer
' Date: 2026
'
' Description:
' Interactive financial dashboard built in Excel using VBA.
' Automates financial analysis, dynamic charts, KPI tracking
' and data visualization for business decision making.
'
' GitHub: https://github.com/cracka07
'========================================================
