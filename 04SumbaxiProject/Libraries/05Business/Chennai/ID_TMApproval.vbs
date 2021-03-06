'[Verify fields displayed under Pending Approval tab]
Public Function VerifyFieldsPendingApproval_TM(arrLblValPairs,strProductType)
If Not IsNull (arrLblValPairs) Then	
   bverifyfieldvalues = VerifyIDLabelValuePairsRandom(coTMApproval_Page.lblPendingApprovals,arrLblValPairs,strProductType,"Pending Approval Fields")
   VerifyFieldsPendingApproval_TM = bverifyfieldvalues
Else
   VerifyFieldsPendingApproval_TM = True
End If
End Function

'[Verify Additional Info table details displayed under Pending Approval tab]
Public Function VerifytblAdditionalInfo_TM(lstAdditioanalInfo)
VerifytblAdditionalInfo_TM = VerifyTableSingleRowData(coTMApproval_Page.tblAdditionalSRInfoHeader,coTMApproval_Page.tblAdditionalSRInfoContent,lstAdditioanalInfo,"Additional Info")
End Function

'[Click on Approve OR Reject Button in TM Pending Approval Page]
Public Function clickButtonApproveReject_TM(sButtonFlag)
	If sButtonFlag = "Approve" Then
		clickButtonApproveReject_TM = clickButtonApprove_TM
	ElseIf  sButtonFlag = "Reject" Then
		clickButtonApproveReject_TM = clickButtonReject_TM
	End If
End Function

'[Click on Approve Button in Pending Approval Page]
Public Function clickButtonApprove_TM()
	coTMApproval_Page.btnApprove.click 
	If Err.Number <> 0 Then
	  clickButtonApprove_TM = False
	  LogMessage "WARN","Verification","Failed to Click Button: Approve", False
	  Exit Function
	End If
	WaitForIServeLoading
	clickButtonApprove_TM = True
End Function

'[Click on Reject Button in Pending Approval Page]
Public Function clickButtonReject_TM()
	coTMApproval_Page.btnReject.click 
	If Err.Number <> 0 Then
	  clickButtonReject_TM = False
	  LogMessage "WARN","Verification","Failed to Click Button: Reject", False
	  Exit Function
	End If
	WaitForIServeLoading
	clickButtonReject_TM = True
End Function

'[Verify display of Approve Button in Pending Approval Page]
Public Function VerifyButtondisplayApprove_TM(strCheckFlag)
	VerifyButtondisplayApprove_TM = VerifyObjectDisabled(coTMApproval_Page.btnApprove,strCheckFlag,"Approve Button")
End Function

'[Verify display of Reject Button in Pending Approval Page]
Public Function VerifyButtondisplayReject_TM(strCheckFlag)
	VerifyButtondisplayReject_TM = VerifyObjectDisabled(coTMApproval_Page.btnReject,strCheckFlag,"Reject Button")
End Function
