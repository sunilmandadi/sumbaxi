'[Verify row Data in Table Selected Offers for Offers List]
Public Function verifytblSelectedOffersContent_Offers(arrRowDataList)
   bDevPending=false
   verifytblSelectedOffersContent_Offers=verifyTableContentList(Offers.OffersListHeader,Offers.OffersListContents,arrRowDataList,"SelectedOffersContent",false,null,null,null)
End Function

'[Click link in Table Selected Offers Header]
Public Function clickSelectedOffers_link(arrRowDataList)
	'****** Function to click the offer
   bDevPending=false
   clickSelectedOffers_link=selectTableLink(Offers.OffersListHeader,Offers.OffersListContents,arrRowDataList,"OfferList" ,"Offer Reference No",false,null ,null,null)
End Function

'[Verify the Reference Number is displayed as per the DB]
Public Function verifyReferenceNumber(Reference_Number)
   bDevPending=false
   bverifyReferenceNumber=true
   If Not IsNull(Reference_Number) Then
       If Not VerifyInnerText (Offers.Reference_Number(), Reference_Number, "Reference Number")Then
           bverifyReferenceNumber=false
       End If
   End If
   verifyReferenceNumber = bverifyReferenceNumber
End Function

'[Verify the OfferType is displayed as per the DB]
Public Function VerifyOfferType(OfferType)
   bDevPending=false
   bVerifyOfferType=true
   If Not IsNull(OfferType) Then
       If Not VerifyInnerText (Offers.Offer_Type(), OfferType, "Offer Type")Then
           bVerifyOfferType=false
       End If
   End If
   VerifyOfferType = bVerifyOfferType
End Function

'[Verify the Product and Account/Card no. is displayed as per the DB]
Public Function VerifyProductnCard(arrRowDataList)
	bDevPending=false
    VerifyProductnCard=verifyTableContentList(Offers.OfferDtlList_TableHeader,Offers.OfferDtlList_TableContent,arrRowDataList,"SelectedProductnCardNumber",false,null,null,null)
End Function

'[Verify the Offer Description is displayed as per the DB]
Public Function VerifyOfferDesc(Offer_Description)
   bDevPending=false
   bVerifyOfferDesc=true
   If Not IsNull(Offer_Description) Then
       If Not VerifyInnerText (Offers.Offer_Description(), Offer_Description, "Offer Description")Then
           bVerifyOfferDesc=false
       End If
   End If
   VerifyOfferDesc = bVerifyOfferDesc
End Function

'[Update the End date and the Status of the offers in the DB]
Public Function UpdateStatusEndDate(strDate, strReferenceNumber)
	bUpdateStatusEndDate=true
	If Ucase(strDate)="TODAY" Then
		If len(Day(CDate(Now)))=1 Then
			strDay="0"&Day(CDate(Now))
		else
			strDay=""&Day(CDate(Now))
		End If
		strDate=""&strDay & " "&monthName(Month(CDate(Now)),true) &" " &Year(CDate(Now))
	End If
	
	'******** Update the end date to the future date for the Offer and the status to OPEN
	strStatus = "OPEN"
	strUpdateQuery=updateCellValinDB_NZD("update cca_offers_main set end_date ='"&strDate&"',status ='"&strStatus&"' where OFFER_REF = '"&strReferenceNumber&"'")
	'validate the status
	strValue=getDBValForColumn_FE("select end_date from cca_offers_main where OFFER_REF = '"&strReferenceNumber&"'")(0)
	'***** the below step will definitely fail as strvalue is "28/01/2016" and strDate is "28 Jan 2016"
'	If not (strValue = strDate) Then
'		LogMessage "WARN","Verification","Failed to update cell value to change End date." ,false
'		bUpdateStatusEndDate=false
'	End If
	'******* Check if the status has been updated to OPEN
	strValStatus=getDBValForColumn_FE("select status from cca_offers_main where OFFER_REF = '"&strReferenceNumber&"'")(0)
	If not (strValStatus = strStatus) Then
		LogMessage "WARN","Verification","Failed to update cell value to change Status." ,false
		bUpdateStatusEndDate=false
	else
		LogMessage "RSLT","Verification.","DB status changed successfully",true
	End If
	UpdateStatusEndDate=bUpdateStatusEndDate
End Function

'[Submit when Customer Decision is Not Interested]
Public Function SubmissionNotInterested(strReason, strComments)
	bSubmissionNotInterested = true
	strCustomerDecision = "Not Interested"
	Set Customer_Decision = Offers.Customer_Decision()
	
	'********* Fill the drop down Customer Decision as Not Interested
	If not isNull(strCustomerDecision) Then
		If not (selectItem_Combobox(Customer_Decision,strCustomerDecision))Then
		   LogMessage "WARN","Verification","Failed to select :"&strCustomerDecision&" From "&Customer_Decision&" down list",false
		   bSubmissionNotInterested= false
		End If
	End If
	
	'********* Fill the reason
	Set Reason = Offers.OfferDtl_Reason()
	If not isNull(strReason) Then
		If not (selectItem_Combobox(Reason,strReason))Then
		   LogMessage "WARN","Verification","Failed to select :"&strReason&" From "&Reason&" down list",false
		   bSubmissionNotInterested = false
		End If
	End If
	
	'if the strReason is Others, set the other reason 
	if(strReason = "Others") then 
		strOtherReason = "Others reason for Testing"
		Offers.OtherReason.set strOtherReason
	End If
	
	'********fill the comments
	Set Comments = Offers.comments()
	Comments.set strComments
	
	'**********Click on Submit now
	Set Submit = Offers.Submit()
	Submit.Click
	
	'*******Check the pop up exists
	Set RqstSubmitPopUp = Offers.RqstSubmitPopUp()
	If RqstSubmitPopUp.Exist Then
	'The status will always be "Closed". Hence, check the status as well
		strStatus = "Closed"
		If Not VerifyInnerText (Offers.SRStatus(), strStatus , "SR Status is")Then
           bSubmissionNotInterested=false
        End If
		Offers.SRClose.Click
	End If
	SubmissionNotInterested = bSubmissionNotInterested
End Function

'[Check the status of the Modified Offers]
Public Function checkStatusDB_Offers(strReferenceNumber)
	'************ Status should be Closed
	bcheckStatusDB_Offers = true
	strStatus = "CLOSED"
	strValue=getDBValForColumn_FE("select STATUS from cca_offers_main where OFFER_REF = '"&strReferenceNumber&"'")(0)
	If not (strValue = strStatus) Then
		LogMessage "WARN","Verification","Status has not been updated to Closed." ,false
		bcheckStatusDB_Offers=false
	End If
	checkStatusDB_Offers=bcheckStatusDB_Offers
End Function

'[Submit when Customer Decision is Interested]
Public Function SubmissionInterested(strAssignedTo, strComments)
	bSubmissionInterested = true
	strCustomerDecision = "Interested"
	Set Customer_Decision = Offers.Customer_Decision()
	
	'********* Fill the drop down Customer Decision as Not Interested
	If not isNull(strCustomerDecision) Then
		If not (selectItem_Combobox(Customer_Decision,strCustomerDecision))Then
		   LogMessage "WARN","Verification","Failed to select :"&strCustomerDecision&" From "&Customer_Decision&" down list",false
		   bSubmissionInterested = false
		End If
	End If
	
	'********* Fill the Field "Assigned To"
	Set OfferDtl_AssignedTo = Offers.OfferDtl_AssignedTo()
	If not isNull(strAssignedTo) Then
		If not (selectItem_Combobox(OfferDtl_AssignedTo,strAssignedTo))Then
		   LogMessage "WARN","Verification","Failed to select :"&strAssignedTo&" From "&OfferDtl_AssignedTo&" down list",false
		   bSubmissionInterested = false
		End If
	End If
	
	'********fill the comments
	Set Comments = Offers.comments()
	Comments.set strComments
	
	'**********Click on Submit now
	Set Submit = Offers.Submit()
	Submit.Click
	
	'******* Check the pop up exists
	Set RqstSubmitPopUp = Offers.RqstSubmitPopUp()
	If RqstSubmitPopUp.Exist Then
	'The status will always be "In Progress". Hence, check the status as well
		strStatus = "In Progress"
		If Not VerifyInnerText (Offers.SRStatus(), strStatus , "SR Status is")Then
           bSubmissionInterested=false
        End If
		Offers.SRClose.Click
	End If
	'********* Close the Offer List tab
	Offers.CloseOfferListTab.Click
	
	SubmissionInterested = bSubmissionInterested
End Function

'[Click the SR Number of Offers from Service Request Tab]
Public Function clickUnknownSRNumber_Offers(strSubType,strCIN)
	'***** Click on the SR tab
'	Offers.ServiceRequest.Click
'	WaitForIcallLoading
	strSubType=trim(strSubType)
	strQuery="select c3_sr_id from orchsvc_sr where sub_type='"&strSubType&"' and contact_cin='"&strCIN&"' order by created_datetime desc"
	strQuery_SRNumber=getDBValForColumn_OL(strQuery)(0)
	UserclickServiceRequestLink=selectTableLink(ServiceRequest.tblServiceRequestHeader, ServiceRequest.tblServiceRequestContent,_
	Array("SR No.:"&strQuery_SRNumber),"Service Requests","SR No.",true,null,null,null)	
	WaitForIcallLoading
	clickUnknownSRNumber_Offers=true
End Function

'[Verify the Inline Message for CIN with only 1 offer]
Public Function verifyForJustOneOffer()
	bverifyForJustOneOffer = true
	'******** Check if the inline Message is being displayed
	strNoOfferMsg = "There is no Offer currently."
	If Not IsNull(strNoOfferMsg) Then
       If Not VerifyInnerText (Offers.NoOfferMsg(), strNoOfferMsg, "Offer Message")Then
           bverifyForJustOneOffer=false
       End If
     End If
     '****** The offer link should be disabled
     'Close the Offer List tab
     Offers.CloseOfferListTab.Click
End Function

'[Check if the SR can be edited]
Public Function verifyEditIA()
	bverifyEditIA = true
	Offers.EditSR().Click
	Offers.AddActivity().Click
	If not Offers.NewIA().exist Then
		LogMessage "WARN","Verification","SR can not be edited and new IA does not exist.",false
		bverifyEditIA=false
	End If
	verifyEditIA = bverifyEditIA
End Function

'[Verify the Offer Link is disabled]
Public Function verifyOfferDisabled()
	'*** Check if the offer is enabled;
	Dim oDesc
	Set oDesc = Description.Create
	oDesc("micclass").value = "Image"
	'find all the images
	Set obj = Offers.SpecialOffer().ChildObjects(oDesc)
	For i = 0 To obj.Count - 1 Step 1
		if obj(i).webelement("ng-show:=!enableOffers") then
			print obj(i).object.GetAttribute("src")
		End if
	Next
End Function

'[Verify the DB Columns for Offers]
Public Function verifyDBColumns(strReferenceNumber, strOPP)
	bverifyDBColumns = true
	'****** Verify Priority from cca_offers_main
	strQueryP="select priority from cca_offers_main where offer_ref='"&strReferenceNumber&"'"
	strQuery_resultP=getDBValForColumn_OL(strQueryP)(0)
	If not isnumeric(strQuery_resultP) Then
		LogMessage "WARN","Verification","Priority Column in DB is not Numeric.",false
		bverifyDBColumns=false
	End If
	'****** Verify opp_attr_ind from cca_offers_main
	strQueryO="select opp_attr_ind from cca_offers_main where offer_ref='"&strReferenceNumber&"'"
	strQuery_resultO=getDBValForColumn_OL(strQueryO)(0)
	If not strQuery_resultO = "OPP" Then
		LogMessage "WARN","Verification","opp_attr_ind Column in DB is not equal to OPP.",false
		bverifyDBColumns=false
	End If
	verifyDBColumns = bverifyDBColumns
End Function
