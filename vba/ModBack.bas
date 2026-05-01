Attribute VB_Name = "ModBack"
Option Explicit

Sub ModBackToSearch()
        
        Application.Visible = False
        frmFinancial.txtTicker = ""
        frmFinancial.lstResult.Clear
        frmFinancial.lstHistorical.Clear
        frmFinancial.Show
        
End Sub
