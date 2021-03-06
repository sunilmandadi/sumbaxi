'**************Created for Past Verification by Kalai-24/11/16********************

'[Click on Past verification Tab from Customer Overview page]
Public Function clickPastVerificationTab()   
   clickPastVerificationTab=PastVerificationNew.tabPastVerification1.Click
   WaitForIcallLoading
End Function


'[Click Verification ID hyperlink on Past verification tab]
Public Function viewPastVerifDetails(strVerificationID, strStatus)
Dim bviewPastVerifDetails: viewPastVerifDetails = True
lstPastVerifDetails = checknull("VerificationID:"&strVerificationID&"|Status:"&strStatus&"")
bviewPastVerifDetails=selectTableLink(PastVerificationNew.tblPastVerificationHeader,PastVerificationNew.tblPastVerificationContent,lstPastVerifDetails,"PastVerification" ,"Action",true,PastVerificationNew.lnkNext ,PastVerificationNew.lnkNext1 ,PastVerificationNew.lnkPrevious)
WaitForICallLoading
viewPastVerifDetails=bviewPastVerifDetails
End Function

'[Verify Tab Past verification exists]
Public Function verifytabVerificationRexist(strTabName)
   bDevPending=false
   verifytabVerificationexist=verifyTabExist(strTabName)
End Function

'[Select Tab Past verification]
Public Function setectTabPastVerification(strTabName)
   bDevPending=false
   setectTabPastVerification=selectTab(strTabName)
End Function

'[Close Tab Past verification]
Public Function closeTabPastVerification(strTabName)
   bDevPending=false
   closeTabPastVerification=closeTab(strTabName)
End Function

'[Verify Field CTI Reference in Verification tab displayed as]
Public Function verifyCTIReference(strCTIReferenceTxt)
   bDevPending=false
   bverifyCTIReference=true
   If Not IsNull(strCTIReferenceTxt) Then
     If Not VerifyInnerText (PastVerificationNew.lblCTIReferance(), strCTIReferenceTxt, "CTI Reference")Then
	   bverifyCTIReference=false
	End If
   End If
   verifyCTIReference=bverifyCTIReference
End Function

'[Verify Field Created By in Verification tab displayed as]
Public Function verifyCreatedBy(strCreatedByTxt)
   bDevPending=false
   bverifyCreatedBy=true
   If Not IsNull(strCreatedByTxt) Then
     If Not VerifyInnerText (PastVerificationNew.lblCreatedby(), strCreatedByTxt, "Created By")Then
	   bverifyCreatedBy=false
	End If
   End If
   verifyCreatedBy=bverifyCreatedBy
End Function

'[Verify Field Call Direction(Inbound/Outbound) in Verification tab displayed as]
Public Function verifyCallDirection(strCallDirectionTxt)
   bDevPending=false
   bverifyCallDirection=true
   If Not IsNull(strCallDirectionTxt) Then
     If Not VerifyInnerText (PastVerificationNew.lblCreatedby(), strCallDirectionTxt, "Call Direction")Then
	   bverifyCallDirection=false
	End If
   End If
   verifyCallDirection=bverifyCallDirection
End Function

'[Verify Field Verification Status in Verification tab displayed as]
Public Function verifyVerificationStatus(strVerificationStatusTxt)
   bDevPending=false
   bverifyVerificationStatus=true
   If Not IsNull(strVerificationStatusTxt) Then
     If Not VerifyInnerText (PastVerificationNew.lblVerificationStatus(), strVerificationStatusTxt, "Verification Status")Then
	   bverifyVerificationStatus=false
	End If
   End If
   verifyVerificationStatus=bverifyVerificationStatus
End Function

'[Verify Field Verification StartDateTime in Verification tab displayed as]
Public Function verifyVerifStartDateTime(strVerifStartDateTimeTxt)
   bDevPending=false
   bverifyVerifStartDateTime=true
   If Not IsNull(strVerifStartDateTimeTxt) Then
     If Not VerifyInnerText (PastVerificationNew.lblVerifStartDateTime(), strVerifStartDateTimeTxt, "Verification StartDateTime")Then
	   bverifyVerifStartDateTime=false
	End If
   End If
   verifyVerifStartDateTime=bverifyVerifStartDateTime
End Function

'[Verify Field Modified dateTime in Verification tab displayed as]
Public Function verifyModifieddateTime(strModifieddateTimeTxt)
   bDevPending=false
   bverifyModifieddateTime=true
   If Not IsNull(strModifieddateTimeTxt) Then
     If Not VerifyInnerText (PastVerificationNew.lblModifieddateTime(), strModifieddateTimeTxt, "Modified dateTime")Then
	   bverifyModifieddateTime=false
	End If
   End If
   verifyModifieddateTime=bverifyModifieddateTime
End Function

'[Verify Field Completion DateTime in Verification tab displayed as]
Public Function verifyCompletionDateTime(strCompletionDateTimeTxt)
   bDevPending=false
   bverifyCompletionDateTime=true
   If Not IsNull(strCompletionDateTimeTxt) Then
     If Not VerifyInnerText (PastVerificationNew.lblCompletionDate(), strCompletionDateTimeTxt, "Completion DateTime")Then
	   bverifyCompletionDateTime=false
	End If
   End If
   verifyCompletionDateTime=bverifyCompletionDateTime
End Function

'[Verify Field Owner in Verification tab displayed as]
Public Function verifyOwner(strOwnerTxt)
   bDevPending=false
   bverifyOwner=true
   If Not IsNull(strOwnerTxt) Then
     If Not VerifyInnerText (PastVerificationNew.lblOwner(), strOwnerTxt, "Owner")Then
	   bverifyOwner=false
	End If
   End If
   verifyOwner=bverifyOwner
End Function

'[Verify Field Last ModifiedBy in Verification tab displayed as]
Public Function verifyLastModifiedBy(strLastModifiedByTxt)
   bDevPending=false
   bverifyLastModifiedBy=true
   If Not IsNull(strLastModifiedByTxt) Then
     If Not VerifyInnerText (PastVerificationNew.lblLastModifiedBy(), strLastModifiedByTxt, "Last ModifiedBy")Then
	   bverifyLastModifiedBy=false
	End If
   End If
   verifyLastModifiedBy=bverifyLastModifiedBy
End Function

'[Verify Tab Verification displayed]
Public Function verifyTabVerificationexist(strTabName)
   bDevPending=false
   verifyVerificationexist=verifyTabExist(strTabName)
End Function

'[Select Tab Verification]
Public Function setectTabVerification(strTabName)
   bDevPending=false
   setectTabVerification=selectTab(strTabName)
End Function

'[Close Tab Verification]
Public Function closeTabVerification(strTabName)
   bDevPending=false
   closeTabVerification=closeTab(strTabName)
End Function

'[Execute DB Query to verify Identifcation ID question answers]
Public Function verifyIdentQuesAns_DBQuery(strDBServer, arrExpectedData)
bverifyIdentQuesAns_DBQuery=true
strQuery="select answers from `iservesgdev`.`cca_cust_ma_rec` where `VERIFICATION_REC_ID` = '"&strVerificationID&"'"
bverifyIdentQuesAns_DBQuery=CompareDBValue_icall(strDBServer,strQuery, arrExpectedData)
WaitForICallLoading
verifyIdentQuesAns_DBQuery=bverifyIdentQuesAns_DBQuery
End Function

'[Execute DB Query to verify Authentication ID question answers]
Public Function verifyAuthQuesAns_DBQuery(strDBServer, arrExpectedData)
bverifyAuthQuesAns_DBQuery=true
strQuery="select answers from `iservesgdev`.`cca_cust_ma_rec` where `VERIFICATION_REC_ID` = '"&strVerificationID&"'"
bverifyAuthQuesAns_DBQuery=CompareDBValue_icall(strDBServer,strQuery, arrExpectedData)
WaitForICallLoading
verifyAuthQuesAns_DBQuery=bverifyAuthQuesAns_DBQuery
End Function

'[Execute DB Query to verify Custom question answers]
Public Function verifyCustQuesAns_DBQuery(strDBServer, arrExpectedData)
bverifyCustQuesAns_DBQuery=true
strQuery="select answers from `iservesgdev`.`cca_cust_ma_rec` where `VERIFICATION_REC_ID` = '"&strVerificationID&"'"
bverifyCustQuesAns_DBQuery=CompareDBValue_icall(strDBServer,strQuery, arrExpectedData)
WaitForICallLoading
verifyCustQuesAns_DBQuery=bverifyCustQuesAns_DBQuery
End Function
	
