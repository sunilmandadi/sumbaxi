Dim coBalancesLimits_Page : Set coBalancesLimits_Page = New clsBalancesLimits

Class clsBalancesLimits

	Private Sub Class_Initialize()
	End Sub
	
	Private Sub Class_Terminate()
	End Sub
	
	'******************************************************************** Object Initialization **********************************************************************************
	Public Function lblBalancesLimits()
		Set lblBalancesLimits = gObjIServePage.WebElement("xpath:=//isrv-balance-and-limits/div")
	End Function
	
	Public Function lblCardLimitsCCCL()
		Set lblCardLimitsCCCL = gObjIServePage.WebElement("xpath:=(//isrv-balance-and-limits/div/div)[1]")
	End Function
	
	Public Function lblBalancesLimitsCCCL()
		Set lblBalancesLimitsCCCL = gObjIServePage.WebElement("xpath:=(//isrv-balance-and-limits/div/div)[2]")
	End Function
	
	Public Function lblLimitsDCATM()
		Set lblLimitsDCATM = gObjIServePage.WebElement("xpath:=//isrv-limits-usage/div")
	End Function
	
	'******************************************************************** End of Object Initialization *************************************************************************************
End Class
