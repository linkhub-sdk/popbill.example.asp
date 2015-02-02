<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"	'회원 사업자번호, "-" 제외
	KeyType= "SELL"             '발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
	MgtKey = "20150121-18"      '연동관리번호 
	UserID = "testkorea"		'회원아이디

	Set result = m_TaxinvoiceService.GetInfo(testCorpNum, KeyType, MgtKey, UserID)

	If Err.Number <> 0 Then
		code = Err.Number
		message = Err.Description
		Err.Clears
	End If

	
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>세금계산서 상태/요약 정보 확인 </legend>
				<% 
					If code = 0 Then 
				%>
					<ul>
						<li> itemKey :  <%=result.itemKey%> </li>
						<li> stateCode :  <%=result.stateCode%> </li>
						<li> taxType :  <%=result.taxType%> </li>
						<li> purposeType :  <%=result.purposeType%> </li>
						<li> modifyCode : <%=result.modifyCode%></li> 
						<li> issueType :  <%=result.issueType %> </li>
						<li> writeDate :  <%=result.writeDate%> </li>
						<li> invoicerCorpName :  <%=result.invoicerCorpName%> </li>
						<li> invoicerCorpNum :  <%=result.invoicerCorpNum%> </li>
						<li> invoicerMgtKey :  <%=result.invoicerMgtKey%> </li>
						<li> invoiceeCorpName :  <%=result.invoiceeCorpName%> </li>
						<li> invoiceeCorpNum :  <%=result.invoiceeCorpNum%> </li>
						<li> invoiceeMgtKey :  <%=result.invoiceeMgtKey%> </li>
						<li> trusteeCorpName :  <%=result.trusteeCorpName%> </li>
						<li> trusteeCorpNum :  <%=result.trusteeCorpName%> </li>
						<li> trusteeMgtKey :  <%=result.trusteeMgtKey%> </li> 
						<li> supplyCostTotal :  <%=result.supplyCostTotal%> </li>
						<li> taxTotal :  <%=result.taxTotal%> </li>
						<li> issueDT :  <%=result.issueDT%> </li>
						<li> preIssueDT :  <%=result.preIssueDT%> </li>
						<li> stateDT :  <%=result.stateDT%> </li>
						<li> openYN :  <%=result.openYN%> </li>
						<li> openDT :  <%=result.openDT%> </li>
						<li> ntsresult :  <%=result.ntsresult%> </li>
						<li> ntsconfirmNum :  <%=result.ntsconfirmNum %> </li>
						<li> ntssendDT :  <%=result.ntssendDT%> </li>
						<li> ntsresultDT :  <%=result.ntsresultDT%> </li>
						<li> ntssendErrCode :  <%=result.ntssendErrCode%> </li>
						<li> stateMemo :  <%=result.stateMemo%> </li>
						<li> regDT :  <%=result.regDT%> </li>
					</ul>	
					<%	
						Else
					%>
						<ul>
							<li>Response.code: <%=code%> </li>
							<li>Response.message: <%=message%> </li>
						</ul>	
					<%	
						End If
					%>
			</fieldset>
		 </div>
	</body>
</html>