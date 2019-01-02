<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 1건의 세금계산서 상세항목을 확인합니다.
	' - 응답항목에 대한 자세한 사항은 "[전자세금계산서 API 연동매뉴얼]
	'   > 4.1 (세금)계산서 구성" 을 참조하시기 바랍니다.
	'**************************************************************
	
	' 팝빌회원 사업자번호, "-" 제외 10자리
	testCorpNum = "1234567890"

	' 세금계산서 발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
	KeyType = "SELL"

	' 문서관리번호
	MgtKey = "20190103-001"

	On Error Resume Next

	Set taxInfo = m_TaxinvoiceService.GetDetailInfo(testCorpNum, KeyType, MgtKey)

	If Err.Number <> 0 Then
		code = Err.Number
		message = Err.Description
		Err.Clears
	End If

	On Error GoTo 0
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>세금계산서 상세정보 확인 </legend>
				<% 

					If code = 0 Then
				%>
				<ul>
					<li>ntsconfirmNum (국세청 승인번호) : <%=taxInfo.ntsconfirmNum%></li>
					<li>writeDate (작성일자) : <%=taxInfo.writeDate%></li>
					<li>chargeDirection (과금방향) : <%=taxInfo.chargeDirection%></li>
					<li>issueType (발행형태) : <%=taxInfo.issueType%></li>
					<li>issueTiming (발행시점) : <%=taxInfo.issueTiming%></li>
					<li>taxType (과세형태) : <%=taxInfo.taxType%></li>
					<li>supplyCostTotal (공급가액 합계) : <%=taxInfo.supplyCostTotal%></li>
					<li>taxTotal (세액 합계) : <%=taxInfo.taxTotal%></li>
					<li>totalAmount (합계금액) : <%=taxInfo.totalAmount%></li>
					<li>cash (현금) : <%=taxInfo.cash%></li>
					<li>chkBill (수표) : <%=taxInfo.chkBill%></li>
					<li>credit (외상) : <%=taxInfo.credit%></li>
					<li>note (어음) : <%=taxInfo.note%></li>
					<li>remark1 (비고1) : <%=taxInfo.remark1%></li>
					<li>remark2 (비고2) : <%=taxInfo.remark2%></li>
					<li>remark3 (비고3) : <%=taxInfo.remark3%></li>

					<li>invoicerCorpNum (공급자 사업자번호) : <%=taxInfo.invoicerCorpNum%> </li>
					<li>invoicerMgtKey (공급자 문서관리번호) : <%=taxInfo.invoicerMgtKey%></li>
					<li>invoicerTaxRegID (공급자 종사업장 식별번호) : <%=taxInfo.invoicerTaxRegID%></li>
					<li>invoicerCorpName (공급자 상호) : <%=taxInfo.invoicerCorpName%></li>
					<li>invoicerCEOName (공급자 대표자명) : <%=taxInfo.invoicerCEOName%></li>
					<li>invoicerAddr (공급자 주소) : <%=taxInfo.invoicerAddr%></li>
					<li>invoicerBizClass (공급자 종목) : <%=taxInfo.invoicerBizClass%></li>
					<li>invoicerBizType (공급자 업태) : <%=taxInfo.invoicerBizType%></li>
					<li>invoicerContactName (공급자 담당자명) : <%=taxInfo.invoicerContactName%></li>
					<li>invoicerTEL (공급자 연락처) : <%=taxInfo.invoicerTEL%></li>
					<li>invoicerHP (공급자 휴대폰번호) : <%=taxInfo.invoicerHP%></li>
					<li>invoicerEmail (공급자 메일) : <%=taxInfo.invoicerEmail%></li>
					<li>invoicerSMSSendYN (알림문자 전송여부) : <%=taxInfo.invoicerSMSSendYN%></li>

					<li>invoiceeType (공급받는자 구분) : <%=taxInfo.invoiceeType%></li>
					<li>invoiceeCorpNum (공급받는자 사업자번호) : <%=taxInfo.invoiceeCorpNum%></li>
					<li>invoiceeMgtKey (공급받는자 문서관리번호) : <%=taxInfo.invoiceeMgtKey%></li>
					<li>invoiceeTaxRegID (공급받는자 종사업장 식별번호) : <%=taxInfo.invoiceeTaxRegID%></li>
					<li>invoiceeCorpName (공급받는자 상호) : <%=taxInfo.invoiceeCorpName%></li>
					<li>invoiceeCEOName (공급받는자 대표자명) : <%=taxInfo.invoiceeCEOName%></li>
					<li>invoiceeAddr (공급받는자 주소) : <%=taxInfo.invoiceeAddr%></li>
					<li>invoiceeBizClass (공급받는자 종목) : <%=taxInfo.invoiceeBizClass%></li>
					<li>invoiceeBizType (공급받는자 업태) : <%=taxInfo.invoiceeBizType%></li>
					<li>invoiceeContactName1 (공급받는자 담당자명) : <%=taxInfo.invoiceeContactName1%></li>
					<li>closeDownState (공급받는자 휴폐업상태) : <%=taxInfo.closeDownState%></li>
					<li>closeDownStateDate (공급받는자 휴폐업일자) : <%=taxInfo.closeDownStateDate%></li>

					<%
						For i=0 To UBound(taxInfo.detailList)-1
					%>
						<fieldset class="fieldset2">
						<legend>상세항목(품목) 정보 <%=i+1%> </legend>
						<ul>
							<li>serialNum (일련번호) : <%=taxInfo.detailList(i).serialNum%></li>
							<li>purchaseDT (거래일자) : <%=taxInfo.detailList(i).purchaseDT%></li>
							<li>itemName (품명) : <%=taxInfo.detailList(i).itemName%></li>
							<li>spec (규격) : <%=taxInfo.detailList(i).spec%></li>
							<li>qty (수량) : <%=taxInfo.detailList(i).qty%></li>
							<li>unitCost (단가) : <%=taxInfo.detailList(i).unitCost%></li>
							<li>supplyCost (공급가액) : <%=taxInfo.detailList(i).supplyCost%></li>
							<li>tax (세액) : <%=taxInfo.detailList(i).tax%></li>
							<li>remark (비고) : <%=taxInfo.detailList(i).remark%></li>
						</ul>
						</fieldset>
					<%
						Next
					%>
					<%
						For i=0 To UBound(taxInfo.addContactList)-1
					%>
						<fieldset class="fieldset2">
							<legend>추가담당자 정보 <%=i+1%> </legend>
								<ul>
									<li>serialNum (일련번호) : <%=taxInfo.addContactList(i).serialNum%></li>
									<li>email (담당자 메일) : <%=taxInfo.addContactList(i).email%></li>
									<li>contactName (담당자명) : <%=taxInfo.addContactList(i).contactName%></li>
								</ul>
							</fieldset>
					<%
						Next
					%>
				</ul>

				<% 
					Else
				%>
					<ul>
						<li>Response.dcode : <%=code%> </li>
						<li>Response.message : <%=message%> </li>
					</ul>
				<%
					End If
				%>
			</fieldset>
		 </div>
	</body>
</html>