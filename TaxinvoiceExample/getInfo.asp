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
	' - https://docs.popbill.com/taxinvoice/asp/api#GetInfo
	'**************************************************************

	' 팝빌회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"

	' 발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
	KeyType= "SELL"

	' 문서번호 
	MgtKey = "20190103-001"

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
						<li> itemKey (세금계산서 아이템키) :  <%=result.itemKey%> </li>
						<li> stateCode (상태코드) :  <%=result.stateCode%> </li>
						<li> taxType (과세형태) :  <%=result.taxType%> </li>
						<li> purposeType (영수/청구) :  <%=result.purposeType%> </li>
						<li> modifyCode (수정사유코드) : <%=result.modifyCode%></li> 
						<li> issueType (발행형태) :  <%=result.issueType %> </li>
						<li> writeDate (작성일자) :  <%=result.writeDate%> </li>

						<li> invoicerCorpName (공급자 상호) :  <%=result.invoicerCorpName%> </li>
						<li> invoicerCorpNum (공급자 사업자번호) :  <%=result.invoicerCorpNum%> </li>
						<li> invoicerMgtKey (공급자 문서번호) :  <%=result.invoicerMgtKey%> </li>
						<li> invoicerPrintYN (공급자 인쇄여부) :  <%=result.invoicerPrintYN%> </li>

						<li> invoiceeCorpName (공급받는자 상호) :  <%=result.invoiceeCorpName%> </li>
						<li> invoiceeCorpNum (공급받는자 사업자번호) :  <%=result.invoiceeCorpNum%> </li>
						<li> invoiceeMgtKey (공급받는자 문서번호) :  <%=result.invoiceeMgtKey%> </li>
						<li> invoiceePrintYN (공급받는자 인쇄여부) :  <%=result.invoiceePrintYN%> </li>
						<li> closeDownState (공급받는자 휴폐업상태) :  <%=result.closeDownState%> </li>
						<li> closeDownStateDate (공급받는자 휴폐업일자) :  <%=result.closeDownStateDate%> </li>
						<li> interOPYN (연동문서여부) :  <%=result.interOPYN%> </li>
						
						<li> supplyCostTotal (공급가액 합계) :  <%=result.supplyCostTotal%> </li>
						<li> taxTotal (세액 합계) :  <%=result.taxTotal%> </li>
						<li> issueDT (발행일시) :  <%=result.issueDT%> </li>

						<li> stateDT (상태 변경일시) :  <%=result.stateDT%> </li>
						<li> openYN (개봉 여부) :  <%=result.openYN%> </li>
						<li> openDT (개봉 일시) :  <%=result.openDT%> </li>
						<li> ntsresult (국세청 전송결과) :  <%=result.ntsresult%> </li>
						<li> ntsconfirmNum (국세청 승인번호) :  <%=result.ntsconfirmNum %> </li>
						<li> ntssendDT (국세청 전송일시) :  <%=result.ntssendDT%> </li>
						<li> ntsresultDT  (국세청 결과 수신일시) :  <%=result.ntsresultDT%> </li>
						<li> ntssendErrCode (전송실패 사유코드) :  <%=result.ntssendErrCode%> </li>
						<li> stateMemo (상태메모) :  <%=result.stateMemo%> </li>
						<li> regDT (임시저장 일자) :  <%=result.regDT%> </li>
						<li> lateIssueYN (지연발행 여부) :  <%=result.lateIssueYN%> </li>
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