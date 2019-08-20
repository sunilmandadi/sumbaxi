Feature:SendSMS

Scenario:TC02_Verify_SendSMS_SpecialOffers_PreValidation_EP2
Given User logged in as CSO "$CSO" "Password1"
Then Verify IServe Dashboard page is Displayed
Search Customer from Dashboard page "$strSearchType" "$strSearchVal"
Navigate to Customer Overview page "$arrResult" "$strSearchType"

Click Menu List Button from Customer Overview Page
Click link View Profile from MenuList
Verify tab exists "CUSTOMER PROFILE"
Verify default value displayed in the Customer Profile dropdown "$PhoneType" "Customer Phone Type"
Verify fields displayed in Customer Profile Contact Info "$ContactInfo" "Contact Info"
Verify all tab Close except Dashboard or Overview
Change the contact CIF status in Database "$strSearchVal"
Click link displayed below Action section "SPECIAL OFFER" "Banking-Facility"
Verify tab exists "OFFER LIST"
Select Row from Special Offers List table displayed "$lstOffersList"
Verify tab exists "NEW OFFER"
Click on scroll down button in page displayed "4"
Set Customer Decision dropdown as "$setCustomerDecision"
Click on scroll down button in page displayed "4"
Set Reason Customer Decision dropdown as "$setReasonCustomerDecision"
Click on scroll down button in page displayed "2"
Enter Comments Optional field displayed in New Offers "$Comments"
Click on scroll down button in page displayed "2"
Click on Submit button in New Offers "$strSubmit"
Verify Submission popup Message displayed for send SMS option in New Offers Page "$SubMsg" "$SubMsgOption"
Verify tab exists "SEND SMS"
Verify Validation message displayed "$strValidationMsg"

