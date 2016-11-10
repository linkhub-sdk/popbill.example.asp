<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 대량의 세금계산서 상태/요약 정보를 확인합니다. (최대 1000건)
	' - 세금계산서 상태정보(GetInfos API) 응답항목에 대한 자세한 정보는
	'  "[전자세금계산서 API 연동매뉴얼] > 4.2. (세금)계산서 상태정보 구성"
	'  을 참조하시기 바랍니다.
	'**************************************************************

	' 팝빌회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"	

	' 발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
	KeyType= "SELL"        

	' 팝빌회원 아이디
	UserID = "testkorea"	  

	' 세금계산서 문서관리번호 배열, 최대 1000건
	Dim MgtKeyList(3) 
	MgtKeyList(0) = "20150121-01"
	MgtKeyList(1) = "20150121-02"
	MgtKeyList(2) = "20150121-03"
	
	Set result = m_TaxinvoiceService.GetInfos(testCorpNum, KeyType, MgtKeyList, UserID)

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
				<legend>세금계산서 상태/요약 정보 확인 - 대량</legend>
				<%
					If code = 0 Then
						For i=0 To result.Count-1
				%>
							<fieldset class="fieldset2">					
								<legend> TaxinvoiceResult : <%=i+1%> </legend>
									<ul>
										<li> itemKey :  <%=result.Item(i).itemKey%> </li>
										<li> stateCode :  <%=result.Item(i).stateCode%> </li>
										<li> taxType :  <%=result.Item(i).taxType%> </li>
										<li> purposeType :  <%=result.Item(i).purposeType%> </li>
										<li> issueType :  <%=result.Item(i).issueType %> </li>
										<li> writeDate :  <%=result.Item(i).writeDate%> </li>
										<li> invoicerCorpName :  <%=result.Item(i).invoicerCorpName%> </li>
										<li> invoicerCorpNum :  <%=result.Item(i).invoicerCorpNum%> </li>
										<li> invoicerMgtKey :  <%=result.Item(i).invoicerMgtKey%> </li>
										<li> invoicerPrintYN :  <%=result.Item(i).invoicerPrintYN%> </li>
										<li> invoiceeCorpName :  <%=result.Item(i).invoiceeCorpName%> </li>
										<li> invoiceeCorpNum :  <%=result.Item(i).invoiceeCorpNum%> </li>
										<li> invoiceeMgtKey :  <%=result.Item(i).invoiceeMgtKey%> </li>
										<li> invoiceePrintYN :  <%=result.Item(i).invoiceePrintYN%> </li>
										<li> trusteeCorpName :  <%=result.Item(i).trusteeCorpName%> </li>
										<li> trusteeCorpNum :  <%=result.Item(i).trusteeCorpName%> </li>
										<li> trusteeMgtKey :  <%=result.Item(i).trusteeMgtKey%> </li> 
										<li> trusteePrintYN :  <%=result.Item(i).trusteePrintYN%> </li> 
										<li> supplyCostTotal :  <%=result.Item(i).supplyCostTotal%> </li>
										<li> taxTotal :  <%=result.Item(i).taxTotal%> </li>
										<li> issueDT :  <%=result.Item(i).issueDT%> </li>
										<li> preIssueDT :  <%=result.Item(i).preIssueDT%> </li>
										<li> stateDT :  <%=result.Item(i).stateDT%> </li>
										<li> openYN :  <%=result.Item(i).openYN%> </li>
										<li> openDT :  <%=result.Item(i).openDT%> </li>
										<li> ntsresult :  <%=result.Item(i).ntsresult%> </li>
										<li> ntsconfirmNum :  <%=result.Item(i).ntsconfirmNum %> </li>
										<li> ntssendDT :  <%=result.Item(i).ntssendDT%> </li>
										<li> ntsresultDT :  <%=result.Item(i).ntsresultDT%> </li>
										<li> ntssendErrCode :  <%=result.Item(i).ntssendErrCode%> </li>
										<li> stateMemo :  <%=result.Item(i).stateMemo%> </li>
										<li> regDT :  <%=result.Item(i).regDT%> </li>
										<li> lateIssueYN :  <%=result.Item(i).lateIssueYN%> </li>
									</ul>
								</fieldset>
				<%
						Next
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
