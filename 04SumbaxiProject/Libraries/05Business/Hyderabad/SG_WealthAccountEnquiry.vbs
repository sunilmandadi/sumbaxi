'*******************Added by Kalyan for Wealth Account Enquiry 1702 08052017 ****************************

'[Verify Wealth Account Balance Limit BL Section Details]
Public Function verifyWealthBLDtls(strAvlBal,strBLLdgrBal,strBLShrEarmrkAmt,strBLGldEarmrkAmt,strBLRFEarmrkAmt,strBLOthrEarmrkAmt,strBLAmtAppldFrmCPF,strBLRgtCashTopUpLmt,strBLCrntYrPlAdjstDt)
	bverifyWealthBLDtls=true
   If Not IsNull(strAvlBal) Then
       If Not verifyInnerText(WealthAccountEnq.lblWealthBLAvlBal(), strAvlBal, "strAvlBal")Then
           bverifyWealthBLDtls=false
       End If
   End If
   
   If Not IsNull(strBLLdgrBal) Then
       If Not verifyInnerText(WealthAccountEnq.lblWealthBLLedgerBal(), strBLLdgrBal, "strBLLdgrBal")Then
           bverifyWealthBLDtls=false
       End If
   End If
   
   If Not IsNull(strBLShrEarmrkAmt) Then
       If Not verifyInnerText(WealthAccountEnq.lblWealthBLShrsEarmrkAmt(), strBLShrEarmrkAmt, "strBLShrEarmrkAmt")Then
           bverifyWealthBLDtls=false
       End If
   End If
   
   If Not IsNull(strBLGldEarmrkAmt) Then
       If Not verifyInnerText(WealthAccountEnq.lblWealthBLGldEarmrkAmt(), strBLGldEarmrkAmt, "strBLGldEarmrkAmt")Then
           bverifyWealthBLDtls=false
       End If
   End If
   
   If Not IsNull(strBLRFEarmrkAmt) Then
       If Not verifyInnerText(WealthAccountEnq.lblWealthBLRFEarmrkAmt(), strBLRFEarmrkAmt, "strBLRFEarmrkAmt")Then
           bverifyWealthBLDtls=false
       End If
   End If
   
   If Not IsNull(strBLOthrEarmrkAmt) Then
       If Not verifyInnerText(WealthAccountEnq.lblWealthBLOthrEarmrkAmt(), strBLOthrEarmrkAmt, "strBLOthrEarmrkAmt")Then
           bverifyWealthBLDtls=false
       End If
   End If
   
   If Not IsNull(strBLAmtAppldFrmCPF) Then
       If Not verifyInnerText(WealthAccountEnq.lblWealthBLAmtAppldFrmCPF(), strBLAmtAppldFrmCPF, "strBLAmtAppldFrmCPF")Then
           bverifyWealthBLDtls=false
       End If
   End If
   
   If Not IsNull(strBLRgtCashTopUpLmt) Then
       If Not verifyInnerText(WealthAccountEnq.lblWealthBLRgtCashTopUpLmt(), strBLRgtCashTopUpLmt, "strBLRgtCashTopUpLmt")Then
           bverifyWealthBLDtls=false
       End If
   End If
   
   If Not IsNull(strBLCrntYrPlAdjstDt) Then
       If Not verifyInnerText(WealthAccountEnq.lblWealthBLCrntYrPlAdjstDt(), strBLCrntYrPlAdjstDt, "strBLCrntYrPlAdjstDt")Then
           bverifyWealthBLDtls=false
       End If
   End If
   
   verifyWealthBLDtls=bverifyWealthBLDtls
End Function

'[Verify Wealth Account Balance Limit Stock Section Details]
Public Function verifyWealthBLStockDtls(strOrgStockLmt,strAvlStockLmt)
	bverifyWealthBLStockDtls=true
   If Not IsNull(strOrgStockLmt) Then
       If Not verifyInnerText(WealthAccountEnq.lblBLOrgStockLmt(), strOrgStockLmt, "strOrgStockLmt")Then
           bverifyWealthBLStockDtls=false
       End If
   End If
   
   If Not IsNull(strAvlStockLmt) Then
       If Not verifyInnerText(WealthAccountEnq.lblBLAvlStockLmt(), strAvlStockLmt, "strAvlStockLmt")Then
           bverifyWealthBLStockDtls=false
       End If
   End If
   
verifyWealthBLStockDtls=bverifyWealthBLStockDtls

End Function

'[Verify Wealth Account Balance Limit Gold Section Details]
Public Function verifyWealthBLGldDtls(strOrgGldLmt,strAvlGldLmt)
	bverifyWealthBLGldDtls=true
   If Not IsNull(strOrgGldLmt) Then
       If Not verifyInnerText(WealthAccountEnq.lblBLOrgGldLmt(), strOrgGldLmt, "strOrgGldLmt")Then
           bverifyWealthBLGldDtls=false
       End If
   End If
   
   If Not IsNull(strAvlGldLmt) Then
       If Not verifyInnerText(WealthAccountEnq.lblBLAvlGldLmt(), strAvlGldLmt, "strAvlGldLmt")Then
           bverifyWealthBLGldDtls=false
       End If
   End If
   
verifyWealthBLGldDtls=bverifyWealthBLGldDtls

End Function

'[Verify Wealth Account Balance Limit Hold Section Details]
Public Function verifyWealthBLHoldDtls(strHalfDay,strOneDay,strTwoDays)
	bverifyWealthBLHoldDtls=true
   If Not IsNull(strHalfDay) Then
       If Not verifyInnerText(WealthAccountEnq.lblBLHalfDay(), strHalfDay, "strHalfDay")Then
           bverifyWealthBLHoldDtls=false
       End If
   End If
   
   If Not IsNull(strOneDay) Then
       If Not verifyInnerText(WealthAccountEnq.lblBLOneDay(), strOneDay, "strOneDay")Then
           bverifyWealthBLHoldDtls=false
       End If
   End If
   
   If Not IsNull(strTwoDays) Then
       If Not verifyInnerText(WealthAccountEnq.lblBLTwoDays(), strTwoDays, "strTwoDays")Then
           bverifyWealthBLHoldDtls=false
       End If
   End If
   
verifyWealthBLHoldDtls=bverifyWealthBLHoldDtls

End Function

'[Select Radio Button Show on Holdings Page]
Public Function selectShowHoldingsRadio(strShow)
	bselectShowHoldingsRadio=true
	bselectShowHoldingsRadio=SelectRadioButtonGrp(strShow, WealthAccountEnq.rbtnShowHoldings, Array("Shares/Unit Trust","Enhanced","Settled Gold"))
   	WaitForICallLoading
	If Err.Number<>0 Then
		bselectShowHoldingsRadio=false
		LogMessage "WARN","Verification","Failed to Click Button : Show on Holdings" ,false
		Exit Function
   End If
   selectShowHoldingsRadio=bselectShowHoldingsRadio
End Function

'[Verify row Data in Table CPF Holdings]
Public Function verifytblCPFHoldings(arrRowDataList)
   bDevPending=false
   verifytblCPFHoldings=verifyTableContentList(WealthAccountEnq.tblShowHoldingsHeader,WealthAccountEnq.tblShowHoldingsContent,arrRowDataList,"CPFHoldings_Records" ,True,WealthAccountEnq.lnkNext ,WealthAccountEnq.lnkNext1,WealthAccountEnq.lnkPrevious)
End Function

'[Verify Wealth Account Holdings DBS Section Details]
Public Function verifyWealthHoldDtlsDBS(strCPFHldngGldSvgAccnt,strCPFHldngGldCertfcts,strCPFHldngPhyGldDBS1,strCPFHldngPhyGldDBS2pnt5,strCPFHldngPhyGldDBS5,strCPFHldngPhyGldDBS10,strCPFHldngPhyGldDBS20,strCPFHldngPhyGldDBS50)
	bverifyWealthHoldDtlsDBS=true
	
	If Not IsNull(strCPFHldngGldSvgAccnt) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFHldngGldSvgAccnt(), strCPFHldngGldSvgAccnt, "strCPFHldngGldSvgAccnt")Then
           bverifyWealthHoldDtlsDBS=false
       End If
   End If
   
   If Not IsNull(strCPFHldngGldCertfcts) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFHldngGldCertfcts(), strCPFHldngGldCertfcts, "strCPFHldngGldCertfcts")Then
           bverifyWealthHoldDtlsDBS=false
       End If
   End If
   
   If Not IsNull(strCPFHldngPhyGldDBS1) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFHldngPhyGldDBS1(), strCPFHldngPhyGldDBS1, "strCPFHldngPhyGldDBS1")Then
           bverifyWealthHoldDtlsDBS=false
       End If
   End If
   
   If Not IsNull(strCPFHldngPhyGldDBS2pnt5) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFHldngPhyGldDBS2pnt5(), strCPFHldngPhyGldDBS2pnt5, "strCPFHldngPhyGldDBS2pnt5")Then
           bverifyWealthHoldDtlsDBS=false
       End If
   End If
   
   If Not IsNull(strCPFHldngPhyGldDBS5) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFHldngPhyGldDBS5(), strCPFHldngPhyGldDBS5, "strCPFHldngPhyGldDBS5")Then
           bverifyWealthHoldDtlsDBS=false
       End If
   End If
   
   If Not IsNull(strCPFHldngPhyGldDBS10) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFHldngPhyGldDBS10(), strCPFHldngPhyGldDBS10, "strCPFHldngPhyGldDBS10")Then
           bverifyWealthHoldDtlsDBS=false
       End If
   End If
   
   If Not IsNull(strCPFHldngPhyGldDBS20) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFHldngPhyGldDBS20(), strCPFHldngPhyGldDBS20, "strCPFHldngPhyGldDBS20")Then
           bverifyWealthHoldDtlsDBS=false
       End If
   End If
   
   If Not IsNull(strCPFHldngPhyGldDBS50) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFHldngPhyGldDBS50(), strCPFHldngPhyGldDBS50, "strCPFHldngPhyGldDBS50")Then
           bverifyWealthHoldDtlsDBS=false
       End If
   End If

verifyWealthHoldDtlsDBS=bverifyWealthHoldDtlsDBS

End Function

'[Verify Wealth Account Holdings BCCS Section Details]
Public Function verifyWealthHoldDtlsBCCS(strCPFHldngPhyGldBCCS110,strCPFHldngPhyGldBCCS14,strCPFHldngPhyGldBCCS12,strCPFHldngPhyGldBCCS1)
	bverifyWealthHoldDtlsBCCS=true
   If Not IsNull(strCPFHldngPhyGldBCCS110) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFHldngPhyGldBCCS110(), strCPFHldngPhyGldBCCS110, "strCPFHldngPhyGldBCCS110")Then
           bverifyWealthHoldDtlsBCCS=false
       End If
   End If
   
   If Not IsNull(strCPFHldngPhyGldBCCS14) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFHldngPhyGldBCCS14(), strCPFHldngPhyGldBCCS14, "strCPFHldngPhyGldBCCS14")Then
           bverifyWealthHoldDtlsBCCS=false
       End If
   End If
   
   If Not IsNull(strCPFHldngPhyGldBCCS12) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFHldngPhyGldBCCS12(), strCPFHldngPhyGldBCCS12, "strCPFHldngPhyGldBCCS12")Then
           bverifyWealthHoldDtlsBCCS=false
       End If
   End If
   
   If Not IsNull(strCPFHldngPhyGldBCCS1) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFHldngPhyGldBCCS1(), strCPFHldngPhyGldBCCS1, "strCPFHldngPhyGldBCCS1")Then
           bverifyWealthHoldDtlsBCCS=false
       End If
   End If

verifyWealthHoldDtlsBCCS=bverifyWealthHoldDtlsBCCS

End Function

'[Verify Wealth Account Holdings SBC Section Details]
Public Function verifyWealthHoldDtlsSBC(strCPFHldngPhyGldSBC1,strCPFHldngPhyGldSBC1by2,strCPFHldngPhyGldSBC1OZ,strCPFHldngPhyGldSBC5OZ,strCPFHldngPhyGldSBC10OZ,strCPFHldngPhyGldSBC20OZ,strCPFHldngPhyGldSBC50OZ,strCPFHldngPhyGldSBC100OZ)
	bverifyWealthHoldDtlsSBC=true
   If Not IsNull(strCPFHldngPhyGldSBC1) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFHldngPhyGldSBC1(), strCPFHldngPhyGldSBC1, "strCPFHldngPhyGldSBC1")Then
           bverifyWealthHoldDtlsSBC=false
       End If
   End If
   
   If Not IsNull(strCPFHldngPhyGldSBC1by2) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFHldngPhyGldSBC1by2(), strCPFHldngPhyGldSBC1by2, "strCPFHldngPhyGldSBC1by2")Then
           bverifyWealthHoldDtlsSBC=false
       End If
   End If
   
   If Not IsNull(strCPFHldngPhyGldSBC1OZ) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFHldngPhyGldSBC1OZ(), strCPFHldngPhyGldSBC1OZ, "strCPFHldngPhyGldSBC1OZ")Then
           bverifyWealthHoldDtlsSBC=false
       End If
   End If
   
   If Not IsNull(strCPFHldngPhyGldSBC5OZ) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFHldngPhyGldSBC5OZ(), strCPFHldngPhyGldSBC5OZ, "strCPFHldngPhyGldSBC5OZ")Then
           bverifyWealthHoldDtlsSBC=false
       End If
   End If
   
   If Not IsNull(strCPFHldngPhyGldSBC10OZ) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFHldngPhyGldSBC10OZ(), strCPFHldngPhyGldSBC10OZ, "strCPFHldngPhyGldSBC10OZ")Then
           bverifyWealthHoldDtlsSBC=false
       End If
   End If
   
   If Not IsNull(strCPFHldngPhyGldSBC20OZ) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFHldngPhyGldSBC20OZ(), strCPFHldngPhyGldSBC20OZ, "strCPFHldngPhyGldSBC20OZ")Then
           bverifyWealthHoldDtlsSBC=false
       End If
   End If
   
   If Not IsNull(strCPFHldngPhyGldSBC50OZ) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFHldngPhyGldSBC50OZ(), strCPFHldngPhyGldSBC50OZ, "strCPFHldngPhyGldSBC50OZ")Then
           bverifyWealthHoldDtlsSBC=false
       End If
   End If
   
   If Not IsNull(strCPFHldngPhyGldSBC100OZ) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFHldngPhyGldSBC100OZ(), strCPFHldngPhyGldSBC100OZ, "strCPFHldngPhyGldSBC100OZ")Then
           bverifyWealthHoldDtlsSBC=false
       End If
   End If

verifyWealthHoldDtlsSBC=bverifyWealthHoldDtlsSBC

End Function

'[Verify Wealth Account Holdings MLeaf Section Details]
Public Function verifyWealthHoldDtlsMleaf(strCPFHldngPhyGldMLeaf110,strCPFHldngPhyGldMLeaf114,strCPFHldngPhyGldMLeaf112,strCPFHldngPhyGldMLeaf1)
	bverifyWealthHoldDtlsMleaf=true
   If Not IsNull(strCPFHldngPhyGldMLeaf110) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFHldngPhyGldMLeaf110(), strCPFHldngPhyGldMLeaf110, "strCPFHldngPhyGldMLeaf110")Then
           bverifyWealthHoldDtlsMleaf=false
       End If
   End If
   
   If Not IsNull(strCPFHldngPhyGldMLeaf114) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFHldngPhyGldMLeaf114(), strCPFHldngPhyGldMLeaf114, "strCPFHldngPhyGldMLeaf114")Then
           bverifyWealthHoldDtlsMleaf=false
       End If
   End If
   
   If Not IsNull(strCPFHldngPhyGldMLeaf112) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFHldngPhyGldMLeaf112(), strCPFHldngPhyGldMLeaf112, "strCPFHldngPhyGldMLeaf112")Then
           bverifyWealthHoldDtlsMleaf=false
       End If
   End If
   
   If Not IsNull(strCPFHldngPhyGldMLeaf1) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFHldngPhyGldMLeaf1(), strCPFHldngPhyGldMLeaf1, "strCPFHldngPhyGldMLeaf1")Then
           bverifyWealthHoldDtlsMleaf=false
       End If
   End If

verifyWealthHoldDtlsMleaf=bverifyWealthHoldDtlsMleaf

End Function

'[Verify Wealth Account Holdings Lion Section Details]
Public Function verifyWealthHoldDtlsLion(strCPFHldngPhyGldLion120,strCPFHldngPhyGldLion110,strCPFHldngPhyGldLion114,strCPFHldngPhyGldLion112,strCPFHldngPhyGldLion1)
	bverifyWealthHoldDtlsLion=true
   If Not IsNull(strCPFHldngPhyGldLion120) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFHldngPhyGldLion120(), strCPFHldngPhyGldLion120, "strCPFHldngPhyGldLion120")Then
           bverifyWealthHoldDtlsLion=false
       End If
   End If
   
   If Not IsNull(strCPFHldngPhyGldLion110) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFHldngPhyGldLion110(), strCPFHldngPhyGldLion110, "strCPFHldngPhyGldLion110")Then
           bverifyWealthHoldDtlsLion=false
       End If
   End If
   
   If Not IsNull(strCPFHldngPhyGldLion114) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFHldngPhyGldLion114(), strCPFHldngPhyGldLion114, "strCPFHldngPhyGldLion114")Then
           bverifyWealthHoldDtlsLion=false
       End If
   End If
   
   If Not IsNull(strCPFHldngPhyGldLion112) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFHldngPhyGldLion112(), strCPFHldngPhyGldLion112, "strCPFHldngPhyGldLion112")Then
           bverifyWealthHoldDtlsLion=false
       End If
   End If
   
   If Not IsNull(strCPFHldngPhyGldLion1) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFHldngPhyGldLion1(), strCPFHldngPhyGldLion1, "strCPFHldngPhyGldLion1")Then
           bverifyWealthHoldDtlsLion=false
       End If
   End If

verifyWealthHoldDtlsLion=bverifyWealthHoldDtlsLion

End Function

'[Verify Wealth Account KeyInfo Section Details]
Public Function verifyWealthKeyInfoDtls(strCPFKeyInfoAcctInfoInvstmntSchm,strCPFKeyInfoAcctInfoSts,strCPFKeyInfoTrdStlmntInstAuthGvn,strCPFKeyInfoTrdStlmntInstDtUpdate,strCPFKeyInfoBnkChrgStlmntAccntMode,strCPFKeyInfoBnkChrgStlmntAccntNo,strCPFKeyInfoBnkChrgDtUpdate)
	bverifyWealthKeyInfoDtls=true
   If Not IsNull(strCPFKeyInfoAcctInfoInvstmntSchm) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFKeyInfoAcctInfoInvstmntSchm(), strCPFKeyInfoAcctInfoInvstmntSchm, "strCPFKeyInfoAcctInfoInvstmntSchm")Then
           bverifyWealthKeyInfoDtls=false
       End If
   End If
   
   If Not IsNull(strCPFKeyInfoAcctInfoSts) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFKeyInfoAcctInfoSts(), strCPFKeyInfoAcctInfoSts, "strCPFKeyInfoAcctInfoSts")Then
           bverifyWealthKeyInfoDtls=false
       End If
   End If
   
   If Not IsNull(strCPFKeyInfoTrdStlmntInstAuthGvn) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFKeyInfoTrdStlmntInstAuthGvn(), strCPFKeyInfoTrdStlmntInstAuthGvn, "strCPFKeyInfoTrdStlmntInstAuthGvn")Then
           bverifyWealthKeyInfoDtls=false
       End If
   End If
   
   If Not IsNull(strCPFKeyInfoTrdStlmntInstDtUpdate) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFKeyInfoTrdStlmntInstDtUpdate(), strCPFKeyInfoTrdStlmntInstDtUpdate, "strCPFKeyInfoTrdStlmntInstDtUpdate")Then
           bverifyWealthKeyInfoDtls=false
       End If
   End If
   
   If Not IsNull(strCPFKeyInfoBnkChrgStlmntAccntMode) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFKeyInfoBnkChrgStlmntAccntMode(), strCPFKeyInfoBnkChrgStlmntAccntMode, "strCPFKeyInfoBnkChrgStlmntAccntMode")Then
           bverifyWealthKeyInfoDtls=false
       End If
   End If
 If Not IsNull(strCPFKeyInfoBnkChrgStlmntAccntNo) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFKeyInfoBnkChrgStlmntAccntNo(), strCPFKeyInfoBnkChrgStlmntAccntNo, "strCPFKeyInfoBnkChrgStlmntAccntNo")Then
           bverifyWealthKeyInfoDtls=false
       End If
   End If
   
   If Not IsNull(strCPFKeyInfoBnkChrgDtUpdate) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFKeyInfoBnkChrgDtUpdate(), strCPFKeyInfoBnkChrgDtUpdate, "strCPFKeyInfoBnkChrgDtUpdate")Then
           bverifyWealthKeyInfoDtls=false
       End If
   End If
verifyWealthKeyInfoDtls=bverifyWealthKeyInfoDtls

End Function

'[Verify Wealth Account SRS Balance Limits Section Details]
Public Function verifySRSWealthBalLmtDtls(strSRSBLBalLmtAvlBal,strSRSBLBalLmtLdgrBal,strSRSBLBalLmtProdEarMrkAmt,strSRSBLBalLmtTempEarMrkAmt,strSRSBLBalLmtIOUBnkChrgs,strSRSBLBalLmtIOUOthrChrgs,strSRSBLBalLmtAccrdCRIntrst)
	bverifySRSWealthBalLmtDtls=true
   If Not IsNull(strSRSBLBalLmtAvlBal) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSBLBalLmtAvlBal(), strSRSBLBalLmtAvlBal, "strSRSBLBalLmtAvlBal")Then
           bverifySRSWealthBalLmtDtls=false
       End If
   End If
   
   If Not IsNull(strSRSBLBalLmtLdgrBal) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSBLBalLmtLdgrBal(), strSRSBLBalLmtLdgrBal, "strSRSBLBalLmtLdgrBal")Then
           bverifySRSWealthBalLmtDtls=false
       End If
   End If
   
   If Not IsNull(strSRSBLBalLmtProdEarMrkAmt) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSBLBalLmtProdEarMrkAmt(), strSRSBLBalLmtProdEarMrkAmt, "strSRSBLBalLmtProdEarMrkAmt")Then
           bverifySRSWealthBalLmtDtls=false
       End If
   End If
   
   If Not IsNull(strSRSBLBalLmtTempEarMrkAmt) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSBLBalLmtTempEarMrkAmt(), strSRSBLBalLmtTempEarMrkAmt, "strSRSBLBalLmtTempEarMrkAmt")Then
           bverifySRSWealthBalLmtDtls=false
       End If
   End If
   
   If Not IsNull(strSRSBLBalLmtIOUBnkChrgs) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSBLBalLmtIOUBnkChrgs(), strSRSBLBalLmtIOUBnkChrgs, "strSRSBLBalLmtIOUBnkChrgs")Then
           bverifySRSWealthBalLmtDtls=false
       End If
   End If
 If Not IsNull(strSRSBLBalLmtIOUOthrChrgs) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSBLBalLmtIOUOthrChrgs(), strSRSBLBalLmtIOUOthrChrgs, "strSRSBLBalLmtIOUOthrChrgs")Then
           bverifySRSWealthBalLmtDtls=false
       End If
   End If
   
   If Not IsNull(strSRSBLBalLmtAccrdCRIntrst) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSBLBalLmtAccrdCRIntrst(), strSRSBLBalLmtAccrdCRIntrst, "strSRSBLBalLmtAccrdCRIntrst")Then
           bverifySRSWealthBalLmtDtls=false
       End If
   End If
verifySRSWealthBalLmtDtls=bverifySRSWealthBalLmtDtls

End Function


''###############################################################################################
''# Name: verifyLabelValuePairs()
''# Description: Function to verify label and value combinations
''# Author:
''# Date: 12-April-2017
''# Input Parameters: strLabelValuePair
''# Output Parameters: None
''###############################################################################################
'
'
'Public Function verifyLabelValuePairs(strLabelValuePair)
'    Dim objDiv
'    Dim blnIndicator:blnIndicator=true
'    set gObjIServePage=Browser("micclass:=Browser").Page("miccclass:=Page")
'    For Iterator2 = 0 To Ubound(strLabelValuePair) Step 1
'        strValues1 = strLabelValuePair(Iterator2)
'        
'        Set objDiv=Description.Create
'        'objDiv("class").value="layout-padding layout-column flex"
'        objDiv("class").value="flex"
'        objDiv("html tag").value="MD-CONTENT"
'        
'        
'        Set objDivChild = gObjIServePage.ChildObjects(objDiv)
'        
'        For k = 0 To objDivChild.Count-1 Step 1
'            
'            Set objLabel=Description.Create
'            objLabel("class").value="ng-scope flex"
'            objLabel("html tag").value="LABEL"
'           objLabel("visible").value="True"
'            'Set obj = gObjIServePage.WebElement("xpath:=//div[@class='layout-padding layout-column flex']["&k+1&"]").ChildObjects(objLabel)
'            Set obj = objDivChild(k).ChildObjects(objLabel)
'            
'            Set objValue=Description.Create
'            objValue("class").value="md-body-2 ng-binding"
'            objValue("html tag").value="SPAN"
'           ' objValue("visible").value="True"
'            'Set obj1 = gObjIServePage.WebElement("xpath:=(//div[@class='md-padding.*'])["&k+1&"]").ChildObjects(objValue)
'            Set obj1 = objDivChild(k).ChildObjects(objValue)
'            intCount = obj.Count
'            
'            For i = 0 To intCount-1 Step 1
'                strInnertextLable = Trim(obj(i).GetROProperty("innertext"))
'                strInnertextValue = Trim(obj1(i).GetROProperty("innertext"))
'                 
'                strValue1 = strInnertextLable &":"&strInnertextValue
'                strValues = strValues &"|"& strValue1
'            Next
'        Next
'        
'        If instr(1,strValues,strValues1,1)>0 Then
'            LogMessage "RSLT","Verification","Text is displayed as expected." & strValues1,True
'            blnIndicator=true
'        Else
'            LogMessage "WARN","Verifiation","Failed to display text as expected" & strValues1  ,false
'            blnIndicator=false
'            
'        End If
'    Next    
'    verifyLabelValuePairs=blnIndicator
'End Function
'
''[Verify labels and values in SRS Balance and Limits page]
'Public Function verifyKeyInfoLabelsAndValues(strlstKeyInfoLablesAndValues)
'
'    Dim blnverifyKeyInfoLabelsAndValues:blnverifyKeyInfoLabelsAndValues = True
'    
'    blnverifyKeyInfoLabelsAndValues = verifyLabelValuePairs(strlstKeyInfoLablesAndValues)
'
'    verifyKeyInfoLabelsAndValues=blnverifyKeyInfoLabelsAndValues
'End Function


'[Verify Wealth Account SRS Balance Limits Contribution Section Details]
Public Function verifySRSWealthBalLmtContribtnDtls(strSRSBLBalLmtMaxCntrbutionAmt,strSRSBLBalLmtBalCntrbutionAmt)
	bverifySRSWealthBalLmtContribtnDtls=true
   If Not IsNull(strSRSBLBalLmtMaxCntrbutionAmt) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSBLBalLmtMaxCntrbutionAmt(), strSRSBLBalLmtMaxCntrbutionAmt, "strSRSBLBalLmtMaxCntrbutionAmt")Then
           bverifySRSWealthBalLmtContribtnDtls=false
       End If
   End If
   
   If Not IsNull(strSRSBLBalLmtBalCntrbutionAmt) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSBLBalLmtBalCntrbutionAmt(), strSRSBLBalLmtBalCntrbutionAmt, "strSRSBLBalLmtBalCntrbutionAmt")Then
           bverifySRSWealthBalLmtContribtnDtls=false
       End If
   End If
   
verifySRSWealthBalLmtContribtnDtls=bverifySRSWealthBalLmtContribtnDtls

End Function

'[Verify Wealth Account SRS Balance Limits Consolidation Section Details]
Public Function verifySRSWealthBalLmtConsldtnDtls(strSRSBLBalLmtConsolidtdCntrbution,strSRSBLBalLmtConsolidtdCntWthdrw,strSRSBLBalLmtConsolidtdWthdrws,strSRSBLBalLmtConsolidtdRealzdPL,strSRSBLBalLmtConsolidtdIncmeErnd,strSRSBLBalLmtConsolidtdBnkChrg,strSRSBLBalLmtConsolidtdOthrChrg,strSRSBLBalLmtConsolidtdCstofSale,strSRSBLBalLmtConsolidtdPenalty,strSRSBLBalLmtConsolidtdWithHldngTax)
	bverifySRSWealthBalLmtConsldtnDtls=true
   If Not IsNull(strSRSBLBalLmtConsolidtdCntrbution) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSBLBalLmtConsolidtdCntrbution(), strSRSBLBalLmtConsolidtdCntrbution, "strSRSBLBalLmtConsolidtdCntrbution")Then
           bverifySRSWealthBalLmtConsldtnDtls=false
       End If
   End If
   
   If Not IsNull(strSRSBLBalLmtConsolidtdCntWthdrw) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSBLBalLmtConsolidtdCntWthdrw(), strSRSBLBalLmtConsolidtdCntWthdrw, "strSRSBLBalLmtConsolidtdCntWthdrw")Then
           bverifySRSWealthBalLmtConsldtnDtls=false
       End If
   End If
   
   If Not IsNull(strSRSBLBalLmtConsolidtdWthdrws) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSBLBalLmtConsolidtdWthdrws(), strSRSBLBalLmtConsolidtdWthdrws, "strSRSBLBalLmtConsolidtdWthdrws")Then
           bverifySRSWealthBalLmtConsldtnDtls=false
       End If
   End If
   
   If Not IsNull(strSRSBLBalLmtConsolidtdRealzdPL) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSBLBalLmtConsolidtdRealzdPL(), strSRSBLBalLmtConsolidtdRealzdPL, "strSRSBLBalLmtConsolidtdRealzdPL")Then
           bverifySRSWealthBalLmtConsldtnDtls=false
       End If
   End If
   
   If Not IsNull(strSRSBLBalLmtConsolidtdIncmeErnd) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSBLBalLmtConsolidtdIncmeErnd(), strSRSBLBalLmtConsolidtdIncmeErnd, "strSRSBLBalLmtConsolidtdIncmeErnd")Then
           bverifySRSWealthBalLmtConsldtnDtls=false
       End If
   End If
   
   If Not IsNull(strSRSBLBalLmtConsolidtdBnkChrg) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSBLBalLmtConsolidtdBnkChrg(), strSRSBLBalLmtConsolidtdBnkChrg, "strSRSBLBalLmtConsolidtdBnkChrg")Then
           bverifySRSWealthBalLmtConsldtnDtls=false
       End If
   End If
   
   If Not IsNull(strSRSBLBalLmtConsolidtdOthrChrg) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSBLBalLmtConsolidtdOthrChrg(), strSRSBLBalLmtConsolidtdOthrChrg, "strSRSBLBalLmtConsolidtdOthrChrg")Then
           bverifySRSWealthBalLmtConsldtnDtls=false
       End If
   End If
   
   If Not IsNull(strSRSBLBalLmtConsolidtdCstofSale) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSBLBalLmtConsolidtdCstofSale(), strSRSBLBalLmtConsolidtdCstofSale, "strSRSBLBalLmtConsolidtdCstofSale")Then
           bverifySRSWealthBalLmtConsldtnDtls=false
       End If
   End If
   
   If Not IsNull(strSRSBLBalLmtConsolidtdPenalty) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSBLBalLmtConsolidtdPenalty(), strSRSBLBalLmtConsolidtdPenalty, "strSRSBLBalLmtConsolidtdPenalty")Then
           bverifySRSWealthBalLmtConsldtnDtls=false
       End If
   End If
   
   If Not IsNull(strSRSBLBalLmtConsolidtdWithHldngTax) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSBLBalLmtConsolidtdWithHldngTax(), strSRSBLBalLmtConsolidtdWithHldngTax, "strSRSBLBalLmtConsolidtdWithHldngTax")Then
           bverifySRSWealthBalLmtConsldtnDtls=false
       End If
   End If
   
verifySRSWealthBalLmtConsldtnDtls=bverifySRSWealthBalLmtConsldtnDtls

End Function


'[Verify row Data in Table SRS Balance Limits Holdings]
Public Function verifytblSRSHoldings(arrRowDataList)
   verifytblSRSHoldings=verifyTableContentList(WealthAccountEnq.tblSRSHoldingsHeader,WealthAccountEnq.tblSRSHoldingsContent,arrRowDataList,"SRSHoldings_Records" ,True,WealthAccountEnq.lnkNext ,WealthAccountEnq.lnkNext1,WealthAccountEnq.lnkPrevious)
End Function

'[Click SRS Product Description Hypelink for Holdings]
Public Function clickLink_SRS_ProdDesc_Holdings(lstProdDescData)
   bDevPending=false
   clickLink_SRS_ProdDesc_Holdings=selectTableLink(WealthAccountEnq.tblSRSHoldingsHeader,WealthAccountEnq.tblSRSHoldingsContent,lstProdDescData,"SRSHoldings_Records","Product Description",True,WealthAccountEnq.lnkNext ,WealthAccountEnq.lnkNext1,WealthAccountEnq.lnkPrevious)
End Function

'[Verify row Data in Table SRS Balance Limits Holdings Protfolio Desc]
Public Function verifytblSRSHoldingsProtDesc(arrRowDataList)
   verifytblSRSHoldingsProtDesc=verifyTableContentList(WealthAccountEnq.tblSRSHoldingsHeaderProtDesc,WealthAccountEnq.tblSRSHoldingsContentProtDesc,arrRowDataList,"SRSHoldingsProtFolioDesc_Records" ,False,null,null,null)
End Function

'[Verify Wealth Account SRS Key Info Section Details]
Public Function verifySRSWealthKeyInfoDtls(strSRSKeyInfoAccntSts,strSRSKeyInfoIRASID,strSRSKeyInfoFrgnDeclSubmisn,strSRSKeyInfoDtFstContrbutn,strSRSKeyInfoDtFstPenltyFeeWithdrw,strSRSKeyInfoValDtDeemdWithdrw,strSRSKeyInfoTxnDtDeemdWithdrw,strSRSKeyInfoDtTaxClrnceReq,strSRSKeyInfoRetrmntAgeFstContributn,strSRSKeyInfoRetrmntAgeFstPenltyFreeWithdrw,strSRSKeyInfoAcctOpngDt,strSRSKeyInfoDtLstTxn)
	bverifySRSWealthKeyInfoDtls=true
   If Not IsNull(strSRSKeyInfoAccntSts) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSKeyInfoAccntSts(), strSRSKeyInfoAccntSts, "strSRSKeyInfoAccntSts")Then
           bverifySRSWealthKeyInfoDtls=false
       End If
   End If
   
   If Not IsNull(strSRSKeyInfoIRASID) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSKeyInfoIRASID(), strSRSKeyInfoIRASID, "strSRSKeyInfoIRASID")Then
           bverifySRSWealthKeyInfoDtls=false
       End If
   End If
   
   If Not IsNull(strSRSKeyInfoFrgnDeclSubmisn) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSKeyInfoFrgnDeclSubmisn(), strSRSKeyInfoFrgnDeclSubmisn, "strSRSKeyInfoFrgnDeclSubmisn")Then
           bverifySRSWealthKeyInfoDtls=false
       End If
   End If
   
   If Not IsNull(strSRSKeyInfoDtFstContrbutn) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSKeyInfoDtFstContrbutn(), strSRSKeyInfoDtFstContrbutn, "strSRSKeyInfoDtFstContrbutn")Then
           bverifySRSWealthKeyInfoDtls=false
       End If
   End If
   
   If Not IsNull(strSRSKeyInfoDtFstPenltyFeeWithdrw) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSKeyInfoDtFstPenltyFeeWithdrw(), strSRSKeyInfoDtFstPenltyFeeWithdrw, "strSRSKeyInfoDtFstPenltyFeeWithdrw")Then
           bverifySRSWealthKeyInfoDtls=false
       End If
   End If
   
   If Not IsNull(strSRSKeyInfoValDtDeemdWithdrw) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSKeyInfoValDtDeemdWithdrw(), strSRSKeyInfoValDtDeemdWithdrw, "strSRSKeyInfoValDtDeemdWithdrw")Then
           bverifySRSWealthKeyInfoDtls=false
       End If
   End If
   
   If Not IsNull(strSRSKeyInfoTxnDtDeemdWithdrw) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSKeyInfoTxnDtDeemdWithdrw(), strSRSKeyInfoTxnDtDeemdWithdrw, "strSRSKeyInfoTxnDtDeemdWithdrw")Then
           bverifySRSWealthKeyInfoDtls=false
       End If
   End If
   
   If Not IsNull(strSRSKeyInfoDtTaxClrnceReq) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSKeyInfoDtTaxClrnceReq(), strSRSKeyInfoDtTaxClrnceReq, "strSRSKeyInfoDtTaxClrnceReq")Then
           bverifySRSWealthKeyInfoDtls=false
       End If
   End If
   
   If Not IsNull(strSRSKeyInfoRetrmntAgeFstContributn) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSKeyInfoRetrmntAgeFstContributn(), strSRSKeyInfoRetrmntAgeFstContributn, "strSRSKeyInfoRetrmntAgeFstContributn")Then
           bverifySRSWealthKeyInfoDtls=false
       End If
   End If
   
   If Not IsNull(strSRSKeyInfoRetrmntAgeFstPenltyFreeWithdrw) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSKeyInfoRetrmntAgeFstPenltyFreeWithdrw(), strSRSKeyInfoRetrmntAgeFstPenltyFreeWithdrw, "strSRSKeyInfoRetrmntAgeFstPenltyFreeWithdrw")Then
           bverifySRSWealthKeyInfoDtls=false
       End If
   End If
   
   If Not IsNull(strSRSKeyInfoAcctOpngDt) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSKeyInfoAcctOpngDt(), strSRSKeyInfoAcctOpngDt, "strSRSKeyInfoAcctOpngDt")Then
           bverifySRSWealthKeyInfoDtls=false
       End If
   End If
   
   If Not IsNull(strSRSKeyInfoDtLstTxn) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSKeyInfoDtLstTxn(), strSRSKeyInfoDtLstTxn, "strSRSKeyInfoDtLstTxn")Then
           bverifySRSWealthKeyInfoDtls=false
       End If
   End If
   
   verifySRSWealthKeyInfoDtls=bverifySRSWealthKeyInfoDtls

End Function

'[Verify Wealth Account SRS Standing Instructions Section Details]
Public Function verifySRSWealthSIDtls(strSRSSISttlmntBnk,strSRSSISttlmntMde,strSRSSISttlmntAcct,strSRSSISttlmntFreqncy,strSRSSIContributnAmt,strSRSSIPrefrdDt,strSRSSIContributnInstSts,strSRSSIEffDtNewContributn)
	bverifySRSWealthSIDtls=true
   If Not IsNull(strSRSSISttlmntBnk) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSSISttlmntBnk(), strSRSSISttlmntBnk, "strSRSSISttlmntBnk")Then
           bverifySRSWealthSIDtls=false
       End If
   End If
   
   If Not IsNull(strSRSSISttlmntMde) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSSISttlmntMde(), strSRSSISttlmntMde, "strSRSSISttlmntMde")Then
           bverifySRSWealthSIDtls=false
       End If
   End If
   
   If Not IsNull(strSRSSISttlmntAcct) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSSISttlmntAcct(), strSRSSISttlmntAcct, "strSRSSISttlmntAcct")Then
           bverifySRSWealthSIDtls=false
       End If
   End If
   
   If Not IsNull(strSRSSISttlmntFreqncy) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSSISttlmntFreqncy(), strSRSSISttlmntFreqncy, "strSRSSISttlmntFreqncy")Then
           bverifySRSWealthSIDtls=false
       End If
   End If
   
   If Not IsNull(strSRSSIContributnAmt) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSSIContributnAmt(), strSRSSIContributnAmt, "strSRSSIContributnAmt")Then
           bverifySRSWealthSIDtls=false
       End If
   End If
   
   If Not IsNull(strSRSSIPrefrdDt) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSSIPrefrdDt(), strSRSSIPrefrdDt, "strSRSSIPrefrdDt")Then
           bverifySRSWealthSIDtls=false
       End If
   End If
   
   If Not IsNull(strSRSSIContributnInstSts) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSSIContributnInstSts(), strSRSSIContributnInstSts, "strSRSSIContributnInstSts")Then
           bverifySRSWealthSIDtls=false
       End If
   End If
   
   If Not IsNull(strSRSSIEffDtNewContributn) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSSIEffDtNewContributn(), strSRSSIEffDtNewContributn, "strSRSSIEffDtNewContributn")Then
           bverifySRSWealthSIDtls=false
       End If
   End If
  
   verifySRSWealthSIDtls=bverifySRSWealthSIDtls

End Function


'[Select Radio Button Transaction Type on Transaction History Page]
Public Function selectTxnTypTxnHisRadio(strTxnTyp)
	bDevPending=False
	bselectTxnTypTxnHisRadio=true
	bselectTxnTypTxnHisRadio=SelectRadioButtonGrp(strTxnTyp,WealthAccountEnq.rbtnSRStxnTypTxnHist, Array("Portfolio","Financial"))
   WaitForICallLoading
	If Err.Number<>0 Then
       bselectTxnTypTxnHisRadio=false
          LogMessage "WARN","Verification","Failed to Click Button : Tan Type on Txn Hist" ,false
       Exit Function
   End If
   selectTxnTypTxnHisRadio=bselectTxnTypTxnHisRadio
End Function


'[Select Combobox Transaction Period on Transaction History Page]
Public Function selectComboTxnPrdTxnHis(strTxnPrd)
	bselectComboTxnPrdTxnHis=true
	If not (selectItem_Combobox(WealthAccountEnq.comboboxSRStxnPrdTxnHist,strTxnPrd))Then
			LogMessage "WARN","Verification","Failed to select :"&strTxnPrd&" From Txn Prd drop down list",false
			bselectComboTxnPrdTxnHis=false
	End If
   WaitForICallLoading
   selectComboTxnPrdTxnHis=bselectComboTxnPrdTxnHis
End Function

'[Verify row Data in Table SRS Transaction History]
Public Function verifytblSRStxnHist(arrRowDataList)
   verifytblSRStxnHist=verifyTableContentList(WealthAccountEnq.tblSRSTxnHistHeader,WealthAccountEnq.tblSRSTxnHistContent,arrRowDataList,"SRSTxnHist_Records" ,True,WealthAccountEnq.lnkNext ,WealthAccountEnq.lnkNext1,WealthAccountEnq.lnkPrevious)
End Function

'[Set TextBox To Date for SRS Transaction History]
Public Function setSRSTxnHistToDt(strToDt)
	selectDateFromCalendar WealthAccountEnq.btnSRSTxnHistToDate,strToDt
	If Err.Number<>0 Then
			setSRSTxnHistToDt=false
			LogMessage "WARN","Verification","Failed to Set Text Box :To Date" ,false
			Exit Function
	End If
	setSRSTxnHistToDt=true
End Function


'[Select Radio Button CPF Product Type on Transaction History Page]
Public Function selectCPFProdTypTxnHisRadio(strProdTyp)
	bselectCPFProdTypTxnHisRadio=true
	bselectCPFProdTypTxnHisRadio=SelectRadioButtonGrp(strProdTyp, WealthAccountEnq.rbtnCPFProdTypTxnHist, Array("Shares","Enhanced"))
   WaitForICallLoading
	If Err.Number<>0 Then
       bselectCPFProdTypTxnHisRadio=false
          LogMessage "WARN","Verification","Failed to Click Button : Prod Type on Txn Hist" ,false
       Exit Function
   End If
   selectCPFProdTypTxnHisRadio=bselectCPFProdTypTxnHisRadio
End Function

'[Select Combobox CPF Transaction Period on Transaction History Page]
Public Function selectCPFComboTxnPrdTxnHis(strTxnPrd)
	bDevPending=False
	bselectCPFComboTxnPrdTxnHis=true
	If not (selectItem_Combobox(WealthAccountEnq.comboboxCPFtxnPrdTxnHist,strTxnPrd))Then
			LogMessage "WARN","Verification","Failed to select :"&strTxnPrd&" From Txn Prd drop down list",false
			bselectCPFComboTxnPrdTxnHis=false
	End If
   WaitForICallLoading
   selectCPFComboTxnPrdTxnHis=bselectCPFComboTxnPrdTxnHis
End Function

'[Verify row Data in Table CPF Transaction History]
Public Function verifytblCPFtxnHist(arrRowDataList)
   bDevPending=false
   verifytblCPFtxnHist=verifyTableContentList(WealthAccountEnq.tblCPFTxnHistHeader,WealthAccountEnq.tblCPFTxnHistContent,arrRowDataList,"CPFTxnHist_Records" ,True,WealthAccountEnq.lnkNext ,WealthAccountEnq.lnkNext1,WealthAccountEnq.lnkPrevious)
End Function

'[Set TextBox From Date for CPF Transaction History]
Public Function setCPFTxnHistFrmDt(strFrmDt)
	selectDateFromCalendar WealthAccountEnq.btnCPFTxnHistFromDate,strFrmDt
	If Err.Number<>0 Then
			setCPFTxnHistFrmDt=false
			LogMessage "WARN","Verification","Failed to Set Text Box :From Date" ,false
			Exit Function
	End If
	setCPFTxnHistFrmDt=true
End Function

'[Set TextBox To Date for CPF Transaction History]
Public Function setCPFTxnHistToDt(strToDt)
bDevPending=true
selectDateFromCalendar WealthAccountEnq.btnCPFTxnHistToDate,strToDt
If Err.Number<>0 Then
		setCPFTxnHistToDt=false
		LogMessage "WARN","Verification","Failed to Set Text Box :To Date" ,false
		Exit Function
End If
setCPFTxnHistToDt=true
End Function

'[Click Go button for CPF SRS Trans Hist]
Public Function clickGoSRSCPFTranHist()
	WealthAccountEnq.btnGOCPFSRSTxnHist().click()
	WaitForIcallLoading
	clickGoSRSCPFTranHist=true
End Function

'[Click CPF Transaction Type Hypelink for transaction History]
Public Function clickLink_CPF_TxnType_TxnHist(lstTxnHistData)
   bDevPending=false
   clickLink_CPF_TxnType_TxnHist=selectTableLink(WealthAccountEnq.tblCPFTxnHistHeader,WealthAccountEnq.tblCPFTxnHistContent,lstTxnHistData,"TxnHist_Data" ,"Transaction Type",True,WealthAccountEnq.lnkNext ,WealthAccountEnq.lnkNext1,WealthAccountEnq.lnkPrevious)
End Function

'[Click SRS Transaction Type Hypelink for transaction History]
Public Function clickLink_SRS_TxnType_TxnHist(lstTxnHistData)
   bDevPending=false
   clickLink_SRS_TxnType_TxnHist=selectTableLink(WealthAccountEnq.tblSRSTxnHistHeader,WealthAccountEnq.tblSRSTxnHistContent,lstTxnHistData,"TxnHist_Data" ,"Transaction Ref",True,WealthAccountEnq.lnkNext ,WealthAccountEnq.lnkNext1,WealthAccountEnq.lnkPrevious)
End Function


'[Verify Wealth Account CPF Transaction History Pop Up Details]
Public Function verifyCPFWealthtxnHistPopUpDtls(strCPFTxnHistPopUpCstofSale,strCPFTxnHistPopUpUntPrice,strCPFTxnHistPopUpQuantity,strCPFTxnHistPopUpTxnFee,strCPFTxnHistPopUpContrctAmt,strCPFTxnHistPopUpCDPChrg,strCPFTxnHistPopUpGSTAmnt,strCPFTxnHistPopUpCstNewProtfolio,strCPFTxnHistPopUpQntyofOldProtfolio,strCPFTxnHistPopUpActlOldQnty,strCPFTxnHistPopUpRsnCode,strCPFTxnHistPopUpDelInd,strCPFTxnHistPopUpPLUpdtInd,strCPFTxnHistPopUpAdjDt)
	bverifyCPFWealthtxnHistPopUpDtls=true
   If Not IsNull(strCPFTxnHistPopUpCstofSale) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFTxnHistPopUpCstofSale(), strCPFTxnHistPopUpCstofSale, "strCPFTxnHistPopUpCstofSale")Then
           bverifyCPFWealthtxnHistPopUpDtls=false
       End If
   End If
   
   If Not IsNull(strCPFTxnHistPopUpUntPrice) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFTxnHistPopUpUntPrice(), strCPFTxnHistPopUpUntPrice, "strCPFTxnHistPopUpUntPrice")Then
           bverifyCPFWealthtxnHistPopUpDtls=false
       End If
   End If
   
   If Not IsNull(strCPFTxnHistPopUpQuantity) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFTxnHistPopUpQuantity(), strCPFTxnHistPopUpQuantity, "strCPFTxnHistPopUpQuantity")Then
           bverifyCPFWealthtxnHistPopUpDtls=false
       End If
   End If
   
   If Not IsNull(strCPFTxnHistPopUpTxnFee) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFTxnHistPopUpTxnFee(), strCPFTxnHistPopUpTxnFee, "strCPFTxnHistPopUpTxnFee")Then
           bverifyCPFWealthtxnHistPopUpDtls=false
       End If
   End If
   
   If Not IsNull(strCPFTxnHistPopUpContrctAmt) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFTxnHistPopUpContrctAmt(), strCPFTxnHistPopUpContrctAmt, "strCPFTxnHistPopUpContrctAmt")Then
           bverifyCPFWealthtxnHistPopUpDtls=false
       End If
   End If
   
   If Not IsNull(strCPFTxnHistPopUpCDPChrg) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFTxnHistPopUpCDPChrg(), strCPFTxnHistPopUpCDPChrg, "strCPFTxnHistPopUpCDPChrg")Then
           bverifyCPFWealthtxnHistPopUpDtls=false
       End If
   End If
   
   If Not IsNull(strCPFTxnHistPopUpGSTAmnt) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFTxnHistPopUpGSTAmnt(), strCPFTxnHistPopUpGSTAmnt, "strCPFTxnHistPopUpGSTAmnt")Then
           bverifyCPFWealthtxnHistPopUpDtls=false
       End If
   End If
   
   If Not IsNull(strCPFTxnHistPopUpCstNewProtfolio) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFTxnHistPopUpCstNewProtfolio(), strCPFTxnHistPopUpCstNewProtfolio, "strCPFTxnHistPopUpCstNewProtfolio")Then
           bverifyCPFWealthtxnHistPopUpDtls=false
       End If
   End If
   
   If Not IsNull(strCPFTxnHistPopUpQntyofOldProtfolio) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFTxnHistPopUpQntyofOldProtfolio(), strCPFTxnHistPopUpQntyofOldProtfolio, "strCPFTxnHistPopUpQntyofOldProtfolio")Then
           bverifyCPFWealthtxnHistPopUpDtls=false
       End If
   End If
   
   If Not IsNull(strCPFTxnHistPopUpActlOldQnty) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFTxnHistPopUpActlOldQnty(), strCPFTxnHistPopUpActlOldQnty, "strCPFTxnHistPopUpActlOldQnty")Then
           bverifyCPFWealthtxnHistPopUpDtls=false
       End If
   End If
   
   If Not IsNull(strCPFTxnHistPopUpRsnCode) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFTxnHistPopUpRsnCode(), strCPFTxnHistPopUpRsnCode, "strCPFTxnHistPopUpRsnCode")Then
           bverifyCPFWealthtxnHistPopUpDtls=false
       End If
   End If
   
   If Not IsNull(strCPFTxnHistPopUpDelInd) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFTxnHistPopUpDelInd(), strCPFTxnHistPopUpDelInd, "strCPFTxnHistPopUpDelInd")Then
           bverifyCPFWealthtxnHistPopUpDtls=false
       End If
   End If
   
   If Not IsNull(strCPFTxnHistPopUpPLUpdtInd) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFTxnHistPopUpPLUpdtInd(), strCPFTxnHistPopUpPLUpdtInd, "strCPFTxnHistPopUpPLUpdtInd")Then
           bverifyCPFWealthtxnHistPopUpDtls=false
       End If
   End If
   
   If Not IsNull(strCPFTxnHistPopUpAdjDt) Then
       If Not verifyInnerText(WealthAccountEnq.lblCPFTxnHistPopUpAdjDt(), strCPFTxnHistPopUpAdjDt, "strCPFTxnHistPopUpAdjDt")Then
           bverifyCPFWealthtxnHistPopUpDtls=false
       End If
   End If
  
   verifyCPFWealthtxnHistPopUpDtls=bverifyCPFWealthtxnHistPopUpDtls

End Function

'[Verify Wealth Account CPF Transaction History Pop Up Details Enhance]
Public Function verifyCPFWealthtxnHistPopUpDtlsEnhance(strEnhanceCostAmt,strEnhanceQuantity,strEnhanceTranstnFee,strEnhanceGSTAmt,strEnhanceRsnCode,strEnhanceDelInd,strEnhancePLUpdtInd)
	bverifyCPFWealthtxnHistPopUpDtlsEnhance=true
	
   If Not IsNull(strEnhanceCostAmt) Then
       If Not verifyInnerText(WealthAccountEnq.lblEnhanceCostAmt(), strEnhanceCostAmt, "strEnhanceCostAmt")Then
           bverifyCPFWealthtxnHistPopUpDtlsEnhance=false
       End If
   End If
   
   If Not IsNull(strEnhanceQuantity) Then
       If Not verifyInnerText(WealthAccountEnq.lblEnhanceQuantity(), strEnhanceQuantity, "strEnhanceQuantity")Then
           bverifyCPFWealthtxnHistPopUpDtlsEnhance=false
       End If
   End If
   
   If Not IsNull(strEnhanceTranstnFee) Then
       If Not verifyInnerText(WealthAccountEnq.lblEnhanceTranstnFee(), strEnhanceTranstnFee, "strEnhanceTranstnFee")Then
           bverifyCPFWealthtxnHistPopUpDtlsEnhance=false
       End If
   End If
   
   If Not IsNull(strEnhanceGSTAmt) Then
       If Not verifyInnerText(WealthAccountEnq.lblEnhanceGSTAmt(), strEnhanceGSTAmt, "strEnhanceGSTAmt")Then
           bverifyCPFWealthtxnHistPopUpDtlsEnhance=false
       End If
   End If
   
   If Not IsNull(strEnhanceRsnCode) Then
       If Not verifyInnerText(WealthAccountEnq.lblEnhanceRsnCode(), strEnhanceRsnCode, "strEnhanceRsnCode")Then
           bverifyCPFWealthtxnHistPopUpDtlsEnhance=false
       End If
   End If
   
   If Not IsNull(strEnhanceDelInd) Then
       If Not verifyInnerText(WealthAccountEnq.lblEnhanceDelInd(), strEnhanceDelInd, "strEnhanceDelInd")Then
           bverifyCPFWealthtxnHistPopUpDtlsEnhance=false
       End If
   End If
   
   If Not IsNull(strEnhancePLUpdtInd) Then
       If Not verifyInnerText(WealthAccountEnq.lblEnhancePLUpdtInd(), strEnhancePLUpdtInd, "strEnhancePLUpdtInd")Then
           bverifyCPFWealthtxnHistPopUpDtlsEnhance=false
       End If
   End If
  
   verifyCPFWealthtxnHistPopUpDtlsEnhance=bverifyCPFWealthtxnHistPopUpDtlsEnhance

End Function

'[Set TextBox To Date for SRS Transaction History]
Public Function setSRSTxnHistToDt(strToDt)
bDevPending=true
	If Not IsNull(strToDt) Then
		selectDateFromCalendar WealthAccountEnq.btnSRSTxnHistToDate,strToDt
	End If	
If Err.Number<>0 Then
		setSRSTxnHistToDt=false
		LogMessage "WARN","Verification","Failed to Set Text Box :To Date" ,false
		Exit Function
End If
setSRSTxnHistToDt=true
End Function

'[Verify Wealth Account SRS Transaction History Pop Up Details]
Public Function verifySRSWealthtxnHistPopUpDtls(strSettlementAmount,strSettlementAccount,strSettlementMode,strReasonCode,strProductTyp,strTransactionCurrency,strStatementRemarks,strTransactionFee,strGSTAmount,strTransactionMode,strOfficerID,strTellerID,strAccountCurrency,strGSTIndicator,strWithdrawalReason,strWithdrawalType,strWithdrawalMode,strAmountPaidCustomer,strPenaltyFee,strWithholdingTax,strTaxableAmount,strWithdrawalFieldInd,strAmountRequestIRAS,strAmountSubjectPenalty)
	bverifySRSWealthtxnHistPopUpDtls=true
   If Not IsNull(strSettlementAmount) Then
       If Not verifyInnerText(WealthAccountEnq.lblSettlementAmount(), strSettlementAmount, "strSettlementAmount")Then
           bverifySRSWealthtxnHistPopUpDtls=false
       End If
   End If
   
   If Not IsNull(strSettlementAccount) Then
       If Not verifyInnerText(WealthAccountEnq.lblSettlementAccount(), strSettlementAccount, "strSettlementAccount")Then
           bverifySRSWealthtxnHistPopUpDtls=false
       End If
   End If
   
   If Not IsNull(strSettlementMode) Then
       If Not verifyInnerText(WealthAccountEnq.lblSettlementMode(), strSettlementMode, "strSettlementMode")Then
           bverifySRSWealthtxnHistPopUpDtls=false
       End If
   End If
   
   If Not IsNull(strReasonCode) Then
       If Not verifyInnerText(WealthAccountEnq.lblReasonCode(), strReasonCode, "strReasonCode")Then
           bverifySRSWealthtxnHistPopUpDtls=false
       End If
   End If
   
   If Not IsNull(strProductTyp) Then
       If Not verifyInnerText(WealthAccountEnq.lblProductType(), strProductTyp, "strProductTyp")Then
           bverifySRSWealthtxnHistPopUpDtls=false
       End If
   End If
   
   If Not IsNull(strTransactionCurrency) Then
       If Not verifyInnerText(WealthAccountEnq.lblTransactionCurrency(), strTransactionCurrency, "strTransactionCurrency")Then
           bverifySRSWealthtxnHistPopUpDtls=false
       End If
   End If
   
   If Not IsNull(strStatementRemarks) Then
       If Not verifyInnerText(WealthAccountEnq.lblStatementRemarks(), strStatementRemarks, "strStatementRemarks")Then
           bverifySRSWealthtxnHistPopUpDtls=false
       End If
   End If
   
   If Not IsNull(strTransactionFee) Then
       If Not verifyInnerText(WealthAccountEnq.lblTransactionFee(), strTransactionFee, "strTransactionFee")Then
           bverifySRSWealthtxnHistPopUpDtls=false
       End If
   End If
   
   If Not IsNull(strGSTAmount) Then
       If Not verifyInnerText(WealthAccountEnq.lblGSTAmount(), strGSTAmount, "strGSTAmount")Then
           bverifySRSWealthtxnHistPopUpDtls=false
       End If
   End If
   
   If Not IsNull(strTransactionMode) Then
       If Not verifyInnerText(WealthAccountEnq.lblTransactionMode(), strTransactionMode, "strTransactionMode")Then
           bverifySRSWealthtxnHistPopUpDtls=false
       End If
   End If
   
   If Not IsNull(strOfficerID) Then
       If Not verifyInnerText(WealthAccountEnq.lblOfficerID(), strOfficerID, "strOfficerID")Then
           bverifySRSWealthtxnHistPopUpDtls=false
       End If
   End If
   
   If Not IsNull(strTellerID) Then
       If Not verifyInnerText(WealthAccountEnq.lblTellerID(), strTellerID, "strTellerID")Then
           bverifySRSWealthtxnHistPopUpDtls=false
       End If
   End If
   
   If Not IsNull(strAccountCurrency) Then
       If Not verifyInnerText(WealthAccountEnq.lblAccountCurrency(), strAccountCurrency, "strAccountCurrency")Then
           bverifySRSWealthtxnHistPopUpDtls=false
       End If
   End If
   
   If Not IsNull(strGSTIndicator) Then
       If Not verifyInnerText(WealthAccountEnq.lblGSTIndicator(), strGSTIndicator, "strGSTIndicator")Then
           bverifySRSWealthtxnHistPopUpDtls=false
       End If
   End If
   
   If Not IsNull(strWithdrawalReason) Then
       If Not verifyInnerText(WealthAccountEnq.lblWithdrawalReason(), strWithdrawalReason, "strWithdrawalReason")Then
           bverifySRSWealthtxnHistPopUpDtls=false
       End If
   End If
   
   If Not IsNull(strWithdrawalType) Then
       If Not verifyInnerText(WealthAccountEnq.lblWithdrawalType(), strWithdrawalType, "strWithdrawalType")Then
           bverifySRSWealthtxnHistPopUpDtls=false
       End If
   End If
   
   If Not IsNull(strWithdrawalMode) Then
       If Not verifyInnerText(WealthAccountEnq.lblWithdrawalMode(), strWithdrawalMode, "strWithdrawalMode")Then
           bverifySRSWealthtxnHistPopUpDtls=false
       End If
   End If
   
   If Not IsNull(strAmountPaidCustomer) Then
       If Not verifyInnerText(WealthAccountEnq.lblAmountPaidCustomer(), strAmountPaidCustomer, "strAmountPaidCustomer")Then
           bverifySRSWealthtxnHistPopUpDtls=false
       End If
   End If
   
   If Not IsNull(strPenaltyFee) Then
       If Not verifyInnerText(WealthAccountEnq.lblPenaltyFee(), strPenaltyFee, "strPenaltyFee")Then
           bverifySRSWealthtxnHistPopUpDtls=false
       End If
   End If
   
   If Not IsNull(strWithholdingTax) Then
       If Not verifyInnerText(WealthAccountEnq.lblWithholdingTax(), strWithholdingTax, "strWithholdingTax")Then
           bverifySRSWealthtxnHistPopUpDtls=false
       End If
   End If
   
   If Not IsNull(strTaxableAmount) Then
       If Not verifyInnerText(WealthAccountEnq.lblTaxableAmount(), strTaxableAmount, "strTaxableAmount")Then
           bverifySRSWealthtxnHistPopUpDtls=false
       End If
   End If
   
   If Not IsNull(strWithdrawalFieldInd) Then
       If Not verifyInnerText(WealthAccountEnq.lblWithdrawalFieldInd(), strWithdrawalFieldInd, "strWithdrawalFieldInd")Then
           bverifySRSWealthtxnHistPopUpDtls=false
       End If
   End If
   
   If Not IsNull(strAmountRequestIRAS) Then
       If Not verifyInnerText(WealthAccountEnq.lblAmountRequestIRAS(), strAmountRequestIRAS, "strAmountRequestIRAS")Then
           bverifySRSWealthtxnHistPopUpDtls=false
       End If
   End If
   
   If Not IsNull(strAmountSubjectPenalty) Then
       If Not verifyInnerText(WealthAccountEnq.lblAmountSubjectPenalty(), strAmountSubjectPenalty, "strAmountSubjectPenalty")Then
           bverifySRSWealthtxnHistPopUpDtls=false
       End If
   End If
  
   verifySRSWealthtxnHistPopUpDtls=bverifySRSWealthtxnHistPopUpDtls

End Function

'[Verify Wealth Account SRS Balance Limits Cosolidated Current Year Details]
Public Function verifySRSWealthBLConCurrntYrDtls(strBLCont,strBLContPrevOprtr,strBLContWithdrawn,strBLContWithdrawnPrevOprtr,strBLWithdrawals,strBLWithdrawalsPrevOprtr,strBLRealisedPL,strBLIncomeEarned,strBLBankCharge,strBLOthrChargs,strBLCostofSale,strBLPenalty,strBLWithholdingTax)
	bverifySRSWealthBLConCurrntYrDtls=true
	
   If Not IsNull(strBLCont) Then
       If Not verifyInnerText(WealthAccountEnq.lblBLCont(), strBLCont, "strBLCont")Then
           bverifySRSWealthBLConCurrntYrDtls=false
       End If
   End If
   
   If Not IsNull(strBLContPrevOprtr) Then
       If Not verifyInnerText(WealthAccountEnq.lblBLContPrevOprtr(), strBLContPrevOprtr, "strBLContPrevOprtr")Then
           bverifySRSWealthBLConCurrntYrDtls=false
       End If
   End If
   
   If Not IsNull(strBLContWithdrawn) Then
       If Not verifyInnerText(WealthAccountEnq.lblBLContWithdrawn(), strBLContWithdrawn, "strBLContWithdrawn")Then
           bverifySRSWealthBLConCurrntYrDtls=false
       End If
   End If
   
   If Not IsNull(strBLContWithdrawnPrevOprtr) Then
       If Not verifyInnerText(WealthAccountEnq.lblBLContWithdrawnPrevOprtr(), strBLContWithdrawnPrevOprtr, "strBLContWithdrawnPrevOprtr")Then
           bverifySRSWealthBLConCurrntYrDtls=false
       End If
   End If
   
   If Not IsNull(strBLWithdrawals) Then
       If Not verifyInnerText(WealthAccountEnq.lblBLWithdrawals(), strBLWithdrawals, "strBLWithdrawals")Then
           bverifySRSWealthBLConCurrntYrDtls=false
       End If
   End If
   
   If Not IsNull(strBLWithdrawalsPrevOprtr) Then
       If Not verifyInnerText(WealthAccountEnq.lblBLWithdrawalsPrevOprtr(), strBLWithdrawalsPrevOprtr, "strBLWithdrawalsPrevOprtr")Then
           bverifySRSWealthBLConCurrntYrDtls=false
       End If
   End If
   
   If Not IsNull(strBLRealisedPL) Then
       If Not verifyInnerText(WealthAccountEnq.lblBLRealisedPL(), strBLRealisedPL, "strBLRealisedPL")Then
           bverifySRSWealthBLConCurrntYrDtls=false
       End If
   End If
   
   If Not IsNull(strBLIncomeEarned) Then
       If Not verifyInnerText(WealthAccountEnq.lblBLIncomeEarned(), strBLIncomeEarned, "strBLIncomeEarned")Then
           bverifySRSWealthBLConCurrntYrDtls=false
       End If
   End If
   
   If Not IsNull(strBLBankCharge) Then
       If Not verifyInnerText(WealthAccountEnq.lblBLBankCharge(), strBLBankCharge, "strBLBankCharge")Then
           bverifySRSWealthBLConCurrntYrDtls=false
       End If
   End If
   
   If Not IsNull(strBLOthrChargs) Then
       If Not verifyInnerText(WealthAccountEnq.lblBLOthrChargs(), strBLOthrChargs, "strBLOthrChargs")Then
           bverifySRSWealthBLConCurrntYrDtls=false
       End If
   End If
   
   If Not IsNull(strBLCostofSale) Then
       If Not verifyInnerText(WealthAccountEnq.lblBLCostofSale(), strBLCostofSale, "strBLCostofSale")Then
           bverifySRSWealthBLConCurrntYrDtls=false
       End If
   End If
   
   If Not IsNull(strBLPenalty) Then
       If Not verifyInnerText(WealthAccountEnq.lblBLPenalty(), strBLPenalty, "strBLPenalty")Then
           bverifySRSWealthBLConCurrntYrDtls=false
       End If
   End If
   
   If Not IsNull(strBLWithholdingTax) Then
       If Not verifyInnerText(WealthAccountEnq.lblBLWithholdingTax(), strBLWithholdingTax, "strBLWithholdingTax")Then
           bverifySRSWealthBLConCurrntYrDtls=false
       End If
   End If
  
   verifySRSWealthBLConCurrntYrDtls=bverifySRSWealthBLConCurrntYrDtls

End Function

'[Verify Wealth Account SRS Txn Hist Protfolio Pop UP Details]
Public Function verifySRSWealthTxnHistProtPopupDtls(strSRSTxnHistContAmount,strSRSTxnHistQuantity,strSRSTxnHistDueDt,strSRSTxnHistAcntName,strSRSTxnHistProdCode,strSRSTxnHistProdProvnCode,strSRSTxnHistContNumb,strSRSTxnHistBatchNumb,strSRSTxnHistRealisedPL,strSRSTxnHistCostofSale,strSRSTxnHistTranFeeWaiv,strSRSTxnHistWaivRsn,strSRSTxnHistStmntRmark,strSRSTxnHistTranCurr,strSRSTxnHistAcntCurr,strSRSTxnHistQuanTyp,strSRSTxnHistUnitPrice,strSRSTxnHistCDPCharge,strSRSTxnHistTranFee,strSRSTxnHistGSTAmt,strSRSTxnHistGSTInd,strSRSTxnHistOfficerID,strSRSTxnHistTellerID)
	bverifySRSWealthTxnHistProtPopupDtls=true
	
   If Not IsNull(strSRSTxnHistContAmount) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSTxnHistContAmount(), strSRSTxnHistContAmount, "strSRSTxnHistContAmount")Then
           bverifySRSWealthTxnHistProtPopupDtls=false
       End If
   End If
   
   If Not IsNull(strSRSTxnHistQuantity) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSTxnHistQuantity(), strSRSTxnHistQuantity, "strSRSTxnHistQuantity")Then
           bverifySRSWealthTxnHistProtPopupDtls=false
       End If
   End If
   
   If Not IsNull(strSRSTxnHistDueDt) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSTxnHistDueDt(), strSRSTxnHistDueDt, "strSRSTxnHistDueDt")Then
           bverifySRSWealthTxnHistProtPopupDtls=false
       End If
   End If
   
   If Not IsNull(strSRSTxnHistAcntName) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSTxnHistAcntName(), strSRSTxnHistAcntName, "strSRSTxnHistAcntName")Then
           bverifySRSWealthTxnHistProtPopupDtls=false
       End If
   End If
   
   If Not IsNull(strSRSTxnHistProdCode) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSTxnHistProdCode(), strSRSTxnHistProdCode, "strSRSTxnHistProdCode")Then
           bverifySRSWealthTxnHistProtPopupDtls=false
       End If
   End If
   
   If Not IsNull(strSRSTxnHistProdProvnCode) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSTxnHistProdProvnCode(), strSRSTxnHistProdProvnCode, "strSRSTxnHistProdProvnCode")Then
           bverifySRSWealthTxnHistProtPopupDtls=false
       End If
   End If
   
   If Not IsNull(strSRSTxnHistContNumb) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSTxnHistContNumb(), strSRSTxnHistContNumb, "strSRSTxnHistContNumb")Then
           bverifySRSWealthTxnHistProtPopupDtls=false
       End If
   End If
   
   If Not IsNull(strSRSTxnHistBatchNumb) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSTxnHistBatchNumb(), strSRSTxnHistBatchNumb, "strSRSTxnHistBatchNumb")Then
           bverifySRSWealthTxnHistProtPopupDtls=false
       End If
   End If
   
   If Not IsNull(strSRSTxnHistRealisedPL) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSTxnHistRealisedPL(), strSRSTxnHistRealisedPL, "strSRSTxnHistRealisedPL")Then
           bverifySRSWealthTxnHistProtPopupDtls=false
       End If
   End If
   
   If Not IsNull(strSRSTxnHistCostofSale) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSTxnHistCostofSale(), strSRSTxnHistCostofSale, "strSRSTxnHistCostofSale")Then
           bverifySRSWealthTxnHistProtPopupDtls=false
       End If
   End If
   
   If Not IsNull(strSRSTxnHistTranFeeWaiv) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSTxnHistTranFeeWaiv(), strSRSTxnHistTranFeeWaiv, "strSRSTxnHistTranFeeWaiv")Then
           bverifySRSWealthTxnHistProtPopupDtls=false
       End If
   End If
   
   If Not IsNull(strSRSTxnHistWaivRsn) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSTxnHistWaivRsn(), strSRSTxnHistWaivRsn, "strSRSTxnHistWaivRsn")Then
           bverifySRSWealthTxnHistProtPopupDtls=false
       End If
   End If
   
   If Not IsNull(strSRSTxnHistStmntRmark) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSTxnHistStmntRmark(), strSRSTxnHistStmntRmark, "strSRSTxnHistStmntRmark")Then
           bverifySRSWealthTxnHistProtPopupDtls=false
       End If
   End If
   
   If Not IsNull(strSRSTxnHistTranCurr) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSTxnHistTranCurr(), strSRSTxnHistTranCurr, "strSRSTxnHistTranCurr")Then
           bverifySRSWealthTxnHistProtPopupDtls=false
       End If
   End If
   
   If Not IsNull(strSRSTxnHistAcntCurr) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSTxnHistAcntCurr(), strSRSTxnHistAcntCurr, "strSRSTxnHistAcntCurr")Then
           bverifySRSWealthTxnHistProtPopupDtls=false
       End If
   End If
   
   If Not IsNull(strSRSTxnHistQuanTyp) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSTxnHistQuanTyp(), strSRSTxnHistQuanTyp, "strSRSTxnHistQuanTyp")Then
           bverifySRSWealthTxnHistProtPopupDtls=false
       End If
   End If
   
   If Not IsNull(strSRSTxnHistUnitPrice) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSTxnHistUnitPrice(), strSRSTxnHistUnitPrice, "strSRSTxnHistUnitPrice")Then
           bverifySRSWealthTxnHistProtPopupDtls=false
       End If
   End If
   
   If Not IsNull(strSRSTxnHistCDPCharge) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSTxnHistCDPCharge(), strSRSTxnHistCDPCharge, "strSRSTxnHistCDPCharge")Then
           bverifySRSWealthTxnHistProtPopupDtls=false
       End If
   End If
   
   If Not IsNull(strSRSTxnHistTranFee) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSTxnHistTranFee(), strSRSTxnHistTranFee, "strSRSTxnHistTranFee")Then
           bverifySRSWealthTxnHistProtPopupDtls=false
       End If
   End If
   
   If Not IsNull(strSRSTxnHistGSTAmt) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSTxnHistGSTAmt(), strSRSTxnHistGSTAmt, "strSRSTxnHistGSTAmt")Then
           bverifySRSWealthTxnHistProtPopupDtls=false
       End If
   End If
   
   If Not IsNull(strSRSTxnHistGSTInd) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSTxnHistGSTInd(), strSRSTxnHistGSTInd, "strSRSTxnHistGSTInd")Then
           bverifySRSWealthTxnHistProtPopupDtls=false
       End If
   End If
   
   If Not IsNull(strSRSTxnHistOfficerID) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSTxnHistOfficerID(), strSRSTxnHistOfficerID, "strSRSTxnHistOfficerID")Then
           bverifySRSWealthTxnHistProtPopupDtls=false
       End If
   End If
   
   If Not IsNull(strSRSTxnHistTellerID) Then
       If Not verifyInnerText(WealthAccountEnq.lblSRSTxnHistTellerID(), strSRSTxnHistTellerID, "strSRSTxnHistTellerID")Then
           bverifySRSWealthTxnHistProtPopupDtls=false
       End If
   End If
  
   verifySRSWealthTxnHistProtPopupDtls=bverifySRSWealthTxnHistProtPopupDtls

End Function
