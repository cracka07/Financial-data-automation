Attribute VB_Name = "ModBack"
Option Explicit

Sub ModBackToSearch()
        
        Application.Visible = False
        frmFinancial.txtTicker = ""
        frmFinancial.lstResult.Clear
        frmFinancial.lstHistorical.Clear
        frmFinancial.Show
        
End Sub
'========================================================
' Project: Financial Data Automation
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

