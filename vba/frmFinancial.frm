VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFinancial 
   Caption         =   "FINANCIAL DATA EXPLORER"
   ClientHeight    =   8085
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11250
   OleObjectBlob   =   "frmFinancial.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFinancial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Application.Quit
End Sub

Private Sub cmdDashboard_Click()
 
 Application.Visible = True
 Me.Hide

End Sub

Private Sub cmdSearch_Click()
        Dim Ticker As String
        Dim cell As Range
        Dim wsTicker As Worksheet
        Dim tblTIcker As ListObject
        Dim exists As Boolean
        Dim newRow As ListRow
        
        
        Ticker = UCase(Trim(Me.txtTicker))
        
        
        
        If Ticker = "" Then
            MsgBox "Please, enter a ticker", vbExclamation, "Financial Data"
            
            Exit Sub
        End If
        
        Set wsTicker = Worksheets("TICKERS")
        Set tblTIcker = wsTicker.ListObjects("tblTickers")
        
        For Each cell In tblTIcker.ListColumns(1).DataBodyRange
                  If cell.Value = Ticker Then
                         exists = True
                        Exit For
                End If
        Next cell
      If exists = False Then
             Set newRow = tblTIcker.ListRows.Add
             newRow.Range(1, 1).Value = Ticker
       End If
       
'     ThisWorkbook.RefreshAll
       
        LoadResult (Ticker)
        Me.TextBox1.Visible = True
        LoadHistoricalData (Ticker)
        
        
End Sub

Private Sub Frame1_Layout()
Frame1.BackColor = RGB(15, 28, 46)
End Sub

Private Sub Frame2_Layout()
Frame2.BackColor = RGB(15, 28, 46)
End Sub

Private Sub Frame3_Layout()
Frame3.BackColor = RGB(15, 28, 46)
End Sub




Private Sub lstResult_Change()
    Dim Ticker As String
        If Me.lstResult <> "" Then
                With Me.lstResult
                 Ticker = .List(.ListIndex, 1)
                 Worksheets("PVTABLES").Range("B1").Value = Ticker
                End With
                Me.cmdDashboard.Visible = True
        End If
End Sub


Private Sub txtTicker_Change()

          If Me.txtTicker = "" Then
                Me.cmdDashboard.Visible = False
        Exit Sub
        End If
    
End Sub

Private Sub UserForm_Activate()
            Me.TextBox1.Visible = False
           
End Sub
Private Sub UserForm_Initialize()

  Me.cmdSearch.BackColor = RGB(33, 150, 243)
  Me.cmdClose.BackColor = RGB(33, 150, 243)
  Me.cmdDashboard.BackColor = RGB(33, 150, 243)
  Me.lstResult.BackColor = RGB(15, 28, 46)
  Me.lstHistorical.BackColor = RGB(15, 28, 46)
  Me.BackColor = RGB(14, 40, 65)
  Me.lblHist.BackColor = RGB(14, 40, 65)
  Me.lblResult.BackColor = RGB(14, 40, 65)
  Me.txtTicker.BackColor = RGB(15, 28, 46)
          If Me.txtTicker = "" Then
                Me.cmdDashboard.Visible = False
        Exit Sub
        End If
       
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
