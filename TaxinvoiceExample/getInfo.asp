<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	'1건의 세금계산서 상태/요약 정보를 확인합니다.
	' - 세금계산서 상태정보(GetInfo API) 응답항목에 대한 자세한 정보는
	'  "[전자세금계산서 API 연동매뉴얼] > 4.2. (세금)계산서 상태정보 구성"
	'   을 참조하시기 바랍니다.
	'**************************************************************

	' 팝빌회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"	

	' 발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
	KeyType= "SELL"             

	' 문서관리번호 
	MgtKey = "20150121-18"      

	' 팝빌회원아이디
	UserID = "testkorea"		

	On Error Resume Next

	Set result = m_TaxinvoiceService.GetInfo(testCorpNum, KeyType, MgtKey, UserID)

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
						<li> invoicerPrintYN :  <%=result.invoicerPrintYN%> </li>
						<li> invoiceeCorpName :  <%=result.invoiceeCorpName%> </li>
						<li> invoiceeCorpNum :  <%=result.invoiceeCorpNum%> </li>
						<li> invoiceeMgtKey :  <%=result.invoiceeMgtKey%> </li>
						<li> invoiceePrintYN :  <%=result.invoiceePrintYN%> </li>
						<li> trusteeCorpName :  <%=result.trusteeCorpName%> </li>
						<li> trusteeCorpNum :  <%=result.trusteeCorpName%> </li>
						<li> trusteeMgtKey :  <%=result.trusteeMgtKey%> </li> 
						<li> trusteePrintYN :  <%=result.trusteePrintYN%> </li> 
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
						<li> lateIssueYN :  <%=result.lateIssueYN%> </li>
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