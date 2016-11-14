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
	MgtKey = "20161114-02"

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
					<li>writeDate : <%=taxInfo.writeDate%></li>
					<li>chargeDirection : <%=taxInfo.chargeDirection%></li>
					<li>issueType : <%=taxInfo.issueType%></li>
					<li>issueTiming : <%=taxInfo.issueTiming%></li>
					<li>taxType : <%=taxInfo.taxType%></li>

					<li>invoicerCorpNum : <%=taxInfo.invoicerCorpNum%> </li>
					<li>invoicerMgtKey : <%=taxInfo.invoicerMgtKey%></li>
					<li>invoicerTaxRegID : <%=taxInfo.invoicerTaxRegID%></li>
					<li>invoicerCorpName : <%=taxInfo.invoicerCorpName%></li>
					<li>invoicerCEOName : <%=taxInfo.invoicerCEOName%></li>
					<li>invoicerAddr : <%=taxInfo.invoicerAddr%></li>
					<li>invoicerBizClass : <%=taxInfo.invoicerBizClass%></li>
					<li>invoicerBizType : <%=taxInfo.invoicerBizType%></li>
					<li>invoicerContactName : <%=taxInfo.invoicerContactName%></li>
					<li>invoicerDeptName : <%=taxInfo.invoicerDeptName%></li>
					<li>invoicerTEL : <%=taxInfo.invoicerTEL%></li>
					<li>invoicerHP : <%=taxInfo.invoicerHP%></li>
					<li>invoicerEmail : <%=taxInfo.invoicerEmail%></li>
					<li>invoicerSMSSendYN : <%=taxInfo.invoicerSMSSendYN%></li>

					<li>invoiceeType : <%=taxInfo.invoiceeType%></li>
					<li>invoiceeCorpNum : <%=taxInfo.invoiceeCorpNum%></li>
					<li>invoiceeMgtKey : <%=taxInfo.invoiceeMgtKey%></li>
					<li>invoiceeTaxRegID : <%=taxInfo.invoiceeTaxRegID%></li>
					<li>invoiceeCorpName : <%=taxInfo.invoiceeCorpName%></li>
					<li>invoiceeCEOName : <%=taxInfo.invoiceeCEOName%></li>
					<li>invoiceeAddr : <%=taxInfo.invoiceeAddr%></li>
					<li>invoiceeBizClass : <%=taxInfo.invoiceeBizClass%></li>
					<li>invoiceeBizType : <%=taxInfo.invoiceeBizType%></li>
					<li>invoiceeContactName1 : <%=taxInfo.invoiceeContactName1%></li>
					<li>invoiceeDeptName1 : <%=taxInfo.invoiceeDeptName1%></li>
					<li>invoiceeCorpTEL1 : <%=taxInfo.invoiceeTEL1%></li>
					<li>invoiceeCorpHP1 : <%=taxInfo.invoiceeHP1%></li>
					<li>invoiceeCorpEmail : <%=taxInfo.invoiceeEmail1%></li>
					
					<%
						For i=0 To UBound(taxInfo.detailList)-1
					%>
						<fieldset class="fieldset2">
						<legend>detailList <%=i+1%> </legend>
						<ul>
							<li>serialNum : <%=taxInfo.detailList(i).serialNum%></li>
							<li>purchaseDT : <%=taxInfo.detailList(i).purchaseDT%></li>
							<li>itemName : <%=taxInfo.detailList(i).itemName%></li>
							<li>spec : <%=taxInfo.detailList(i).spec%></li>
							<li>qty : <%=taxInfo.detailList(i).qty%></li>
							<li>unitCost : <%=taxInfo.detailList(i).unitCost%></li>
							<li>supplyCost : <%=taxInfo.detailList(i).supplyCost%></li>
							<li>tax : <%=taxInfo.detailList(i).tax%></li>
							<li>remark : <%=taxInfo.detailList(i).remark%></li>
						</ul>
						</fieldset>
					<%
						Next
					%>
					<%
						For i=0 To UBound(taxInfo.addContactList)-1
					%>
						<fieldset class="fieldset2">
							<legend>addContactList <%=i+1%> </legend>
								<ul>
									<li>serialNum : <%=taxInfo.addContactList(i).serialNum%></li>
									<li>email : <%=taxInfo.addContactList(i).email%></li>
									<li>contactName : <%=taxInfo.addContactList(i).contactName%></li>
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